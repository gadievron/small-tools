/*
 * Email to Name Matcher for Gmail
 * Author: Gadi Evron
 * Version: 1.2.0
 * License: MIT
 *
 * Given emails (possibly with display names attached like
 * "John Doe <john@example.com>") in column B, finds the most likely real-world
 * display name via Gmail headers and Calendar guest lists.
 *
 * SHEET LAYOUT (auto-created in tab "Email to Name" if needed)
 *   A - Name (output)
 *   B - Email (input; may contain "Name <email>" pasted form, will be cleaned)
 *   C - Status (live during run; final on completion)
 *   D - Other name variants seen
 *   E - Confidence, e.g. "High confidence (from: 12 hits)"
 *
 * CASCADE (first non-empty result wins)
 *   INPUT -> FROM -> TO -> CC -> BCC -> CALENDAR (last 5y) -> LOCAL-PART
 *   If column B already contains "Name <email>", the embedded name is used
 *   directly (Confirmed tier). If nothing else matches, the local part is
 *   parsed as first.last/first_last/first-last (Low tier, can be overridden
 *   by re-runs on richer mailboxes).
 *
 * CONFIDENCE TIERS (ranked for re-run merge)
 *   1. Confirmed (input)   - name pasted in by user
 *   2. High               - >= 10 hits in winning phase
 *   3. Medium             - 3-9 hits
 *   4. Low                - 1-2 hits
 *   5. Not found
 *
 * RE-RUN BEHAVIOR (run again on another mailbox/browser)
 *   For each row, compare new result tier to existing tier:
 *     - New > existing  -> replace name+email; demote old name into "Other names"
 *     - New == existing -> keep existing; add new name to "Other names" if different
 *     - New < existing  -> no change
 *     - "Not found" never overwrites anything
 *   This makes results monotonically improve across multiple accounts.
 *
 * LOGGING
 *   Execution events (sheet creation, input parsing, skips, merges, downgrades)
 *   are written to console.log() and visible in Apps Script's Execution log.
 *
 * LIMITATIONS
 *   - Calendar phase reads getDefaultCalendar() only.
 *   - Long lists may exceed Apps Script's 6-min runtime; run in batches if so.
 */

// ============================================================================
// Sheet setup
// ============================================================================

function getOrCreateResultsSheet(ss, headers) {
  var name = "Email to Name";

  // Highest priority: if the results tab already exists anywhere in this
  // workbook, use it. Makes cross-mailbox enrichment work regardless of
  // which tab happens to be active when the user clicks Run.
  var existing = ss.getSheetByName(name);
  if (existing) {
    existing.activate();
    console.log("Reusing existing tab '" + name + "' (found in workbook)");
    return existing;
  }

  // No results tab yet. Check if the active sheet looks set up correctly.
  var active = ss.getActiveSheet();
  var firstRow = active.getRange(1, 1, 1, headers.length).getValues()[0]
    .map(function(v) { return String(v || "").trim().toLowerCase(); });

  var headersMatch = true;
  for (var i = 0; i < headers.length; i++) {
    var expected = headers[i].toLowerCase();
    var actual = firstRow[i];
    if (actual && actual !== expected) { headersMatch = false; break; }
  }
  var activeHasContent = !!active.getRange("A1").getValue() || !!active.getRange("A2").getValue() ||
                         !!active.getRange("B1").getValue() || !!active.getRange("B2").getValue();
  if (headersMatch && activeHasContent) {
    console.log("Using active sheet '" + active.getName() + "' (headers already match)");
    return active;
  }

  // Create the results tab and copy emails over from active sheet's column A
  var sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.activate();
  console.log("Created new tab '" + name + "' with headers");

  var sourceColA = active.getRange("A1:A").getValues().map(function(r) {
    return String(r[0] || "").trim();
  });
  var lastIdx = -1;
  for (var k = sourceColA.length - 1; k >= 0; k--) {
    if (sourceColA[k] !== "") { lastIdx = k; break; }
  }
  if (lastIdx >= 0) {
    var startIdx = extractEmailAddress(sourceColA[0]) ? 0 : 1;
    var rows = [];
    var foundOne = false;
    for (var n = startIdx; n <= lastIdx; n++) {
      var val = sourceColA[n];
      rows.push([val]);
      if (!foundOne && extractEmailAddress(val)) foundOne = true;
    }
    if (foundOne && rows.length) {
      sheet.getRange(2, 2, rows.length, 1).setValues(rows);
      console.log("Copied " + rows.length + " row(s) from '" + active.getName() + "' column A into column B");
    }
  }
  return sheet;
}

// ============================================================================
// Input parsing
// ============================================================================

// Last-resort heuristic: derive a name from the email's local part if it
// contains a separator (. _ -). Splits on those and title-cases each token.
// Single-letter first token becomes an initial ("J."). No rejection rules:
// 'info.team@x.com' will yield "Info Team", which is acceptable noise at the
// Low confidence tier.
function nameFromLocalPart(localPart) {
  if (!localPart) return null;
  var tokens = localPart.toLowerCase().split(/[._\-]+/).filter(Boolean);
  if (tokens.length < 2) return null;

  var formatted = tokens.map(function(tok, idx) {
    if (tok.length === 1 && idx === 0) return tok.toUpperCase() + ".";
    return tok.charAt(0).toUpperCase() + tok.slice(1);
  });
  return formatted.join(" ");
}

// Returns { name, email } from "Cory Scott <cory@x.com>" or { email } from a
// bare address. name is null when not present.
function parseInputCell(raw) {
  if (!raw) return { name: null, email: "" };
  var s = String(raw).trim();

  // "Name <email>" form
  var match = s.match(/^(.*?)<([^>]+)>\s*[,;]?\s*$/);
  if (match) {
    var rawName = match[1].trim().replace(/^['"\u2018\u2019\u201C\u201D]+|['"\u2018\u2019\u201C\u201D]+$/g, "").trim();
    var rawAddr = match[2].trim();
    var email = extractEmailAddress(rawAddr);
    if (email && rawName) {
      // Reject if "name" is actually an email or matches the local-part
      if (/@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/.test(rawName)) return { name: null, email: email };
      var localPart = email.split('@')[0].toLowerCase();
      var nameCompact = rawName.toLowerCase().replace(/[._\-\s]/g, "");
      var localCompact = localPart.replace(/[._\-]/g, "");
      if (!/\s/.test(rawName) && nameCompact === localCompact) return { name: null, email: email };
      return { name: cleanDisplayName(rawName, localPart) || null, email: email };
    }
    if (email) return { name: null, email: email };
  }

  return { name: null, email: extractEmailAddress(s) };
}

// Split a Gmail header value (From/To/Cc/Bcc) on commas, but ignore commas
// inside angle brackets or quoted strings. Handles "Doe, Jane" <jane@x.com>
// correctly, which a single regex with a lookahead cannot.
function splitHeaderParts(header) {
  var parts = [];
  var cur = "";
  var inAngle = false;
  var inQuote = false;
  for (var i = 0; i < header.length; i++) {
    var c = header.charAt(i);
    if (c === '"' && !inAngle) inQuote = !inQuote;
    else if (c === '<' && !inQuote) inAngle = true;
    else if (c === '>' && !inQuote) inAngle = false;
    if (c === ',' && !inAngle && !inQuote) {
      parts.push(cur);
      cur = "";
      continue;
    }
    cur += c;
  }
  if (cur.length) parts.push(cur);
  return parts;
}

function extractEmailAddress(raw) {
  if (!raw) return "";
  var s = String(raw).trim();
  var angle = s.match(/<([^>]+)>/);
  if (angle) s = angle[1].trim();
  s = s.replace(/^mailto:/i, "");
  s = s.replace(/^['"\u2018\u2019\u201C\u201D]+|['"\u2018\u2019\u201C\u201D]+$/g, "");
  s = s.replace(/^[\s,;:<>()\[\]]+|[\s,;:.<>()\[\]]+$/g, "");
  var m = s.match(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/);
  return m ? m[0].toLowerCase() : "";
}

// ============================================================================
// Confidence ranking for re-run merge
// ============================================================================

// Tiers: 4 Confirmed > 3 High > 2 Medium > 1 Low > 0 Not found / empty
function confidenceTier(confStr) {
  if (!confStr) return 0;
  var s = String(confStr).toLowerCase();
  if (s.indexOf("confirmed") !== -1) return 4;
  if (s.indexOf("high") !== -1) return 3;
  if (s.indexOf("medium") !== -1) return 2;
  if (s.indexOf("low") !== -1) return 1;
  return 0;
}

// ============================================================================
// Rate limiting
// ============================================================================

var lastAPICall_e2n = 0;
function rateLimitedGmailSearch_e2n(query, start, max) {
  if (!query) throw new Error("rateLimitedGmailSearch_e2n is a helper. Run matchEmailsToNames instead.");
  var elapsed = Date.now() - lastAPICall_e2n;
  if (elapsed < 300) Utilities.sleep(300 - elapsed);
  lastAPICall_e2n = Date.now();
  return GmailApp.search(query, start, max);
}

// ============================================================================
// Main entry point
// ============================================================================

function matchEmailsToNames() {
  console.log("=== matchEmailsToNames started at " + new Date().toISOString() + " ===");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var headers = ["Name", "Email", "Status", "Other name variants", "Confidence"];
  var sheet = getOrCreateResultsSheet(ss, headers);

  for (var c = 0; c < headers.length; c++) {
    if (!sheet.getRange(1, c + 1).getValue()) sheet.getRange(1, c + 1).setValue(headers[c]);
  }

  // Read input column (B) starting at row 2
  var colB = sheet.getRange("B2:B").getValues().map(function(r) { return String(r[0] || "").trim(); });
  var lastIdx = -1;
  for (var i = colB.length - 1; i >= 0; i--) { if (colB[i] !== "") { lastIdx = i; break; } }
  if (lastIdx === -1) throw new Error("No emails found. Add emails under column B before running.");

  console.log("Processing " + (lastIdx + 1) + " row(s)");
  var inputs = colB.slice(0, lastIdx + 1);

  for (var j = 0; j < inputs.length; j++) {
    var raw = inputs[j];
    var row = j + 2;

    try {
      if (!raw) {
        writeRow(sheet, row, { name: "", status: "Empty row", others: "", confidence: "" });
        continue;
      }

      var parsed = parseInputCell(raw);
      if (!parsed.email) {
        writeRow(sheet, row, { name: "", status: "Invalid email format", others: "", confidence: "" });
        console.log("Row " + row + ": invalid input '" + raw + "'");
        continue;
      }

      // Clean column B: write back canonical email regardless of how it was pasted
      if (raw !== parsed.email) {
        sheet.getRange(row, 2).setValue(parsed.email);
      }

      // Build new result for this run
      var newResult;
      if (parsed.name) {
        newResult = {
          name: parsed.name,
          status: "Found in INPUT",
          others: "",
          confidence: "Confirmed confidence (input: pasted name)"
        };
        console.log("Row " + row + " (" + parsed.email + "): used name '" + parsed.name + "' from input, skipping search");
      } else {
        var res = resolveNameCascade(parsed.email, row, sheet);
        var finalStatus = "Not found";
        if (res.status) {
          var srcMatch = res.status.match(/\((from|to|cc|bcc|calendar|local-part):/i);
          if (srcMatch) {
            var src = srcMatch[1].toUpperCase();
            finalStatus = "Found in " + (
              src === "FROM"       ? "FROM headers"     :
              src === "TO"         ? "TO headers"       :
              src === "CC"         ? "CC headers"       :
              src === "BCC"        ? "BCC headers"      :
              src === "CALENDAR"   ? "CALENDAR"         :
              "EMAIL LOCAL PART"
            );
          }
        }
        newResult = {
          name: res.best || "",
          status: finalStatus,
          others: res.others || "",
          confidence: res.status || ""
        };
        console.log("Row " + row + " (" + parsed.email + "): " + finalStatus +
                    (res.best ? " -> '" + res.best + "'" : ""));
      }

      // Merge with existing result on the row (re-run logic)
      var existing = {
        name:       String(sheet.getRange(row, 1).getValue() || "").trim(),
        status:     String(sheet.getRange(row, 3).getValue() || "").trim(),
        others:     String(sheet.getRange(row, 4).getValue() || "").trim(),
        confidence: String(sheet.getRange(row, 5).getValue() || "").trim()
      };
      var merged = mergeResults(existing, newResult, row, parsed.email);
      writeRow(sheet, row, merged);

    } catch (error) {
      console.log("Row " + row + ": ERROR " + ((error && error.message) || error));
      writeRow(sheet, row, {
        name: "Error",
        status: String((error && error.message) || error),
        others: "",
        confidence: ""
      });
    }
  }
  console.log("=== matchEmailsToNames finished at " + new Date().toISOString() + " ===");
}

function writeRow(sheet, row, r) {
  sheet.getRange(row, 1, 1, 5).setValues([[
    r.name || "",
    sheet.getRange(row, 2).getValue(),  // preserve column B (email)
    r.status || "",
    r.others || "",
    r.confidence || ""
  ]]);
}

// ============================================================================
// Re-run merge logic
// ============================================================================

function mergeResults(existing, fresh, row, email) {
  var oldTier = confidenceTier(existing.confidence);
  var newTier = confidenceTier(fresh.confidence);

  // No prior result: just take new
  if (oldTier === 0 && !existing.name) {
    return fresh;
  }

  // New found nothing: keep existing (never downgrade to "Not found")
  if (newTier === 0) {
    console.log("Row " + row + " (" + email + "): kept existing (new run found nothing)");
    return existing;
  }

  if (newTier > oldTier) {
    // Upgrade: replace, demote old name into others
    var demoted = existing.name ? appendToOthers(fresh.others, existing.name) : fresh.others;
    console.log("Row " + row + " (" + email + "): UPGRADED tier " + oldTier + " -> " + newTier +
                ", '" + existing.name + "' -> '" + fresh.name + "'");
    return {
      name: fresh.name,
      status: fresh.status,
      others: demoted,
      confidence: fresh.confidence
    };
  }

  if (newTier === oldTier) {
    // Same tier: keep existing name, add new to others if different
    if (fresh.name && fresh.name !== existing.name) {
      var withAlt = appendToOthers(existing.others, fresh.name);
      console.log("Row " + row + " (" + email + "): equal tier, added '" + fresh.name + "' to alternates");
      return {
        name: existing.name,
        status: existing.status,
        others: withAlt,
        confidence: existing.confidence
      };
    }
    console.log("Row " + row + " (" + email + "): equal tier, no change");
    return existing;
  }

  // newTier < oldTier: keep existing
  console.log("Row " + row + " (" + email + "): kept existing (new tier " + newTier + " < old " + oldTier + ")");
  return existing;
}

function appendToOthers(existingOthers, newName) {
  if (!newName) return existingOthers || "";
  if (!existingOthers) return newName;
  // Avoid duplicates (case-insensitive, ignoring bracketed counts)
  var existingNames = existingOthers.split(/,\s*/).map(function(s) {
    return s.replace(/\s*\[[^\]]*\]\s*$/, "").trim().toLowerCase();
  });
  if (existingNames.indexOf(newName.toLowerCase()) !== -1) return existingOthers;
  return existingOthers + ", " + newName;
}

// ============================================================================
// Cascade
// ============================================================================

function resolveNameCascade(email, statusRow, sheet) {
  var phases = [
    { label: "Searching FROM headers...",    fn: function() { return searchHeadersForEmail(email, 'from:', 50); }, src: 'from' },
    { label: "Searching TO headers...",      fn: function() { return searchHeadersForEmail(email, 'to:',   50); }, src: 'to' },
    { label: "Searching CC headers...",      fn: function() { return searchHeadersForEmail(email, 'cc:',   50); }, src: 'cc' },
    { label: "Searching BCC headers...",     fn: function() { return searchHeadersForEmail(email, 'bcc:',  50); }, src: 'bcc' },
    { label: "Searching Calendar guests...", fn: function() { return searchCalendarForEmail(email, 200); },       src: 'calendar' }
  ];

  for (var i = 0; i < phases.length; i++) {
    sheet.getRange(statusRow, 3).setValue(phases[i].label);
    SpreadsheetApp.flush();
    var result = phases[i].fn();
    if (result) return labelWithSource(result, phases[i].src);
  }

  // Last-resort: derive from local part (Low tier so re-runs can override)
  sheet.getRange(statusRow, 3).setValue("Deriving from email local part...");
  SpreadsheetApp.flush();
  var localPart = email.split('@')[0];
  var derived = nameFromLocalPart(localPart);
  if (derived) {
    return {
      best: derived,
      status: "Low confidence (local-part: derived from " + localPart + ")",
      others: ""
    };
  }

  return { best: null, status: "Not found", others: "" };
}

function labelWithSource(result, src) {
  var hits = result._hits;
  var level = hits >= 10 ? "High" : (hits >= 3 ? "Medium" : "Low");
  var out = {
    best: result.name,
    status: level + " confidence (" + src + ": " + hits + " hit" + (hits === 1 ? "" : "s") + ")"
  };
  if (result._alts && result._alts.length) {
    out.others = result._alts.slice(0, 5)
      .map(function(a) { return a.name + " [" + a.count + "]"; })
      .join(", ");
  }
  return out;
}

// ============================================================================
// Header search
// ============================================================================

function searchHeadersForEmail(email, field, maxThreads) {
  var query = field + '"' + email + '" newer_than:5y';
  var threads = rateLimitedGmailSearch_e2n(query, 0, maxThreads);
  if (!threads.length) return null;

  var groups = {};
  var emailLower = email.toLowerCase();
  var localPart = email.split('@')[0].toLowerCase();

  for (var t = 0; t < threads.length; t++) {
    var msgs = threads[t].getMessages();
    for (var m = 0; m < msgs.length; m++) {
      var header = "";
      if (field === 'from:')      header = msgs[m].getFrom();
      else if (field === 'to:')   header = msgs[m].getTo();
      else if (field === 'cc:')   header = msgs[m].getCc();
      else if (field === 'bcc:')  header = msgs[m].getBcc();
      if (!header) continue;

      var date = msgs[m].getDate().getTime();
      var parts = splitHeaderParts(header);

      for (var p = 0; p < parts.length; p++) {
        var raw = parts[p].trim();
        var match = raw.match(/^(.*)<(.+?)>$/);
        var disp = "", addr = "";
        if (match) { disp = (match[1] || "").trim(); addr = match[2].trim(); }
        else { addr = raw; disp = ""; }

        if (!addr || addr.toLowerCase() !== emailLower) continue;
        if (!disp) continue;

        var cleaned = cleanDisplayName(disp, localPart);
        if (!cleaned) continue;

        var key = canonicalKey(cleaned);
        if (!key) continue;

        if (!groups[key]) groups[key] = { count: 0, lastSeen: 0, variants: {} };
        groups[key].count++;
        if (date > groups[key].lastSeen) groups[key].lastSeen = date;
        groups[key].variants[cleaned] = (groups[key].variants[cleaned] || 0) + 1;
      }
    }
  }

  return pickBestGroup(groups);
}

// ============================================================================
// Calendar search
// ============================================================================

function searchCalendarForEmail(email, maxEvents) {
  var cal = CalendarApp.getDefaultCalendar();
  var endDate = new Date();
  var startDate = new Date();
  startDate.setFullYear(endDate.getFullYear() - 5);

  var events = cal.getEvents(startDate, endDate);
  if (!events || !events.length) return null;

  var emailLower = email.toLowerCase();
  var localPart = email.split('@')[0].toLowerCase();
  var groups = {};

  var processed = 0;
  for (var i = 0; i < events.length && processed < maxEvents; i++) {
    processed++;
    var ev = events[i];
    var when = ev.getStartTime() ? ev.getStartTime().getTime() : ev.getLastUpdated().getTime();
    var guests = (ev.getGuestList && ev.getGuestList(true)) || [];
    if (!guests.length) continue;

    for (var g = 0; g < guests.length; g++) {
      var ge = guests[g];
      var gEmail = ((ge.getEmail && ge.getEmail()) || "").toLowerCase();
      if (gEmail !== emailLower) continue;

      var gName = (ge.getName && ge.getName()) || "";
      if (!gName) continue;

      var cleaned = cleanDisplayName(gName, localPart);
      if (!cleaned) continue;

      var key = canonicalKey(cleaned);
      if (!key) continue;

      if (!groups[key]) groups[key] = { count: 0, lastSeen: 0, variants: {} };
      groups[key].count++;
      if (when > groups[key].lastSeen) groups[key].lastSeen = when;
      groups[key].variants[cleaned] = (groups[key].variants[cleaned] || 0) + 1;
    }
  }

  return pickBestGroup(groups);
}

// ============================================================================
// Pick best variant: most frequent, ties by longer name, then recency
// ============================================================================

function pickBestGroup(groups) {
  var keys = Object.keys(groups);
  if (!keys.length) return null;

  var list = keys.map(function(k) {
    var g = groups[k];
    var variantList = Object.keys(g.variants).map(function(v) { return { v: v, c: g.variants[v] }; });
    variantList.sort(function(a, b) { return (b.c - a.c) || (b.v.length - a.v.length); });
    return { name: variantList[0].v, count: g.count, lastSeen: g.lastSeen };
  });

  list.sort(function(a, b) {
    return (b.count - a.count) || (b.name.length - a.name.length) || (b.lastSeen - a.lastSeen);
  });

  var best = list[0];
  best._hits = best.count;
  best._alts = list.slice(1, 6).map(function(c) { return { name: c.name, count: c.count }; });
  return best;
}

// ============================================================================
// Display name cleanup & canonicalization
// ============================================================================

function cleanDisplayName(raw, localPart) {
  if (!raw) return null;
  var s = String(raw).trim();

  s = s.replace(/^['"\u2018\u2019\u201C\u201D]+|['"\u2018\u2019\u201C\u201D]+$/g, "").trim();
  s = s.replace(/\s*\([^)]*@[^)]*\)\s*$/, "").trim();
  s = s.replace(/\s+via\s+\S.*$/i, "").trim();
  s = s.replace(/\s+\((?:external|guest|via [^)]+)\)\s*$/i, "").trim();

  if (!s) return null;
  if (/@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/.test(s)) return null;

  if (!/\s/.test(s)) {
    var sCompact = s.toLowerCase().replace(/[._\-]/g, "");
    var localCompact = localPart.replace(/[._\-]/g, "");
    if (sCompact === localCompact) return null;
  }

  if (/^\d+$/.test(s)) return null;
  if (s.length < 2) return null;

  if (s === s.toLowerCase() && /[a-z]/.test(s)) {
    s = s.replace(/\b([a-z])/g, function(_, ch) { return ch.toUpperCase(); });
  }

  return s;
}

function canonicalKey(s) {
  if (!s) return "";
  return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "")
          .toLowerCase()
          .replace(/[^a-z0-9\s]/g, " ")
          .replace(/\s+/g, " ")
          .trim();
}
