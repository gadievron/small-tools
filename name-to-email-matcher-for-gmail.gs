/*
 * Name: Name to Email Matcher for Gmail
 * Author: Gadi Evron
 * Version: 1.9.1
 * License: MIT
 *
 * OVERVIEW
 *   This is a complete overkill Apps Script to match names on a sheet to emails
 *   through Gmail.
 *   
 *   Given names in Column A, this script searches Gmail and Calendar to find likely
 *   email addresses, writing results to:
 *     - B: Email
 *     - C: Status (live during search; final on completion)
 *     - D: Other addresses (top alternates from the winning phase)
 *     - E: Confidence (e.g., "High confidence (from: 24.5)")
 *
 * SHEET LAYOUT (Row 1 headers auto-created if missing)
 *   A - Names
 *   B - Emails
 *   C - Status
 *   D - Other addresses
 *   E - Confidence
 *
 * CASCADE LOGIC (FIRST HIT WINS; NO ENRICHMENT)
 *   1) Gmail FROM headers
 *   2) Gmail TO headers
 *   3) Gmail CC headers
 *   4) Gmail BCC headers (visible on your sent mail)
 *   5) Calendar guests (last 5 years)
 *   6) Email body harvest
 *   The search stops at the first phase yielding a candidate. Later phases do not
 *   enrich earlier results.
 *
 * BEHAVIOR / UX
 *   - Per-row processing with try/catch. Each row writes B..E immediately.
 *   - Live status in Column C during processing ("Searching FROM headers...", etc.),
 *     then overwritten with the final status ("Found in FROM headers", ..., "Not found").
 *   - "Skip" rule for idempotence/resume:
 *       If B already has an email AND E contains a numeric score >= 10 (Medium/High),
 *       the row is skipped on re-runs. If E is missing/unparsable (or score < 10),
 *       the row is processed again.
 *   - Empty name rows write: ["", "Empty row", "", ""] to B..E.
 *
 * RATE LIMITING
 *   - All Gmail searches are spaced by at least 300 ms via a wrapper
 *     to reduce transient rate-limit errors.
 *
 * SCORING (SUMMARY)
 *   Local-part patterns (highest weight):
 *     +20:  firstlast / lastfirst / first.last / last.first (normalized)
 *     +12:  local contains both first and last anywhere
 *     +10:  local == first (if domain includes last) OR local == last (if domain includes first)
 *      +8:  local startsWith(first) OR startsWith(last)
 *      +8:  initial+last OR last+initial (e.g., flastname, lastnamef)
 *      +4:  local contains first; +4 if local contains last (partials)
 *   Recency bonus (headers/body/calendar):
 *      +6:  <= 1 year
 *      +3:  <= 3 years
 *      +0:  otherwise
 *   Display-name overlap (headers/calendar):
 *      +4:  two or more tokens hit
 *      +2:  one token hit
 *   Opaque local-part with exact display-name match (no local overlap):
 *      +8:  outbound (your sent TO/CC or BCC)
 *      +4:  inbound (headers, calendar)
 *   Outbound channel bonus:
 *      +4:  if address is in TO/CC of an outbound message
 *   Header body-corroboration (same message only):
 *      +4:  header candidate matches only first OR only last, and the message body
 *           contains the missing token
 *   Body phase:
 *      -2:  noise penalty
 *      +2:  if local starts with first initial
 *   Calendar participant bonus (applied in calendar phase only, scaled by recency):
 *     +10:  if recency bonus is +6 (<= 1 year)
 *      +5:  if recency bonus is +3 (<= 3 years)
 *   Junk filtering:
 *     Excludes noreply/donotreply/mailer-daemon/bounces/notifications and Google system senders.
 *   Confidence labels:
 *     High:   score >= 20
 *     Medium: 10 <= score < 20
 *     Low:    score < 10
 *
 * ACCEPTANCE GATES (HEADERS/CALENDAR)
 *   - Prefer display-name overlap with tokens from the query name.
 *   - If no display overlap, allow strong local-part patterns:
 *       contains normalized first & last, or initial+last / last+initial.
 *
 * CONTRACT (DO NOT CHANGE WITHOUT INTENT)
 *   - Cascade order: FROM -> TO -> CC -> BCC -> CALENDAR -> BODY.
 *   - Column semantics: B=email, C=final status, D=alternates (winning phase only), E=confidence string.
 *   - Skip rule: only skip if B present and E's numeric score >= 10; otherwise process.
 *   - Rate limiting: Gmail searches must use the wrapper (>= 300 ms spacing).
 *   - No enrichment across phases; first hit wins.
 */

/* === Rate limit wrapper for Gmail searches (>=300ms) === */
let lastAPICall = 0;
function rateLimitedGmailSearch(query, start, max) {
  const elapsed = Date.now() - lastAPICall;
  if (elapsed < 300) Utilities.sleep(300 - elapsed);
  lastAPICall = Date.now();
  return GmailApp.search(query, start, max);
}

function matchNamesToEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // Ensure headers
  if (!sheet.getRange(1, 1).getValue()) sheet.getRange(1, 1).setValue("Names");
  if (!sheet.getRange(1, 2).getValue()) sheet.getRange(1, 2).setValue("Emails");
  if (!sheet.getRange(1, 3).getValue()) sheet.getRange(1, 3).setValue("Status");
  if (!sheet.getRange(1, 4).getValue()) sheet.getRange(1, 4).setValue("Other addresses");
  if (!sheet.getRange(1, 5).getValue()) sheet.getRange(1, 5).setValue("Confidence");

  // Find last non-empty row in column A (below header)
  const colA = sheet.getRange("A2:A").getValues().map(r => String(r[0] || "").trim());
  let lastIdx = -1;
  for (let i = colA.length - 1; i >= 0; i--) { if (colA[i] !== "") { lastIdx = i; break; } }
  if (lastIdx === -1) throw new Error("No names found. Add names under column A before running.");

  const names = colA.slice(0, lastIdx + 1);

  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    const row = i + 2;

    try {
      if (!name) {
        sheet.getRange(row, 2, 1, 4).setValues([[ "", "Empty row", "", "" ]]);
        continue;
      }

      // Skip if Email (B) present AND Confidence (E) parses to score >= 10
      const existingEmail = String(sheet.getRange(row, 2).getValue() || "").trim();
      const existingConf  = String(sheet.getRange(row, 5).getValue() || "").trim();
      let parsedScore = null;
      if (existingConf) {
        const m = existingConf.match(/\((?:from|to|cc|bcc|calendar|body):\s*([0-9]+(?:\.[0-9]+)?)\)/i);
        if (m) parsedScore = parseFloat(m[1]);
      }
      if (existingEmail && parsedScore !== null && parsedScore >= 10) {
        // Leave row intact; do not re-query
        continue;
      }

      const res = resolveEmailCascadeWithStatus(name, row);

      // Final plain status in Column C derived from res.status
      let finalStatus = "Not found";
      if (res.status && /\((from|to|cc|bcc|calendar|body):/i.test(res.status)) {
        const src = (res.status.match(/\((from|to|cc|bcc|calendar|body):/i) || [,""])[1].toUpperCase();
        finalStatus = src ? ("Found in " + (src === "FROM" ? "FROM headers" :
                                            src === "TO" ? "TO headers" :
                                            src === "CC" ? "CC headers" :
                                            src === "BCC" ? "BCC headers" :
                                            src === "CALENDAR" ? "CALENDAR" : "BODY")) : "Found";
      }

      // Write B..E for this row
      sheet.getRange(row, 2, 1, 4).setValues([[ 
        res.best || "Not found",
        finalStatus,
        res.others || "",
        res.status || ""
      ]]);

    } catch (error) {
      sheet.getRange(row, 2, 1, 4).setValues([[ "Error", String((error && error.message) || error), "", "" ]]);
    }
  }
}

function resolveEmailCascade(name) {
  const tokens = name.split(/\s+/).filter(Boolean).map(s => s.toLowerCase());

  // Phase A: FROM
  let result = searchHeadersForName(name, tokens, 'from:', 30);
  if (result) return labelWithSource(result, 'from');

  // Phase B: TO / CC
  result = searchHeadersForName(name, tokens, 'to:', 30);
  if (!result) result = searchHeadersForName(name, tokens, 'cc:', 30);
  if (result) return labelWithSource(result, result._src);

  // Phase B.5: BCC
  result = searchHeadersForName(name, tokens, 'bcc:', 30);
  if (result) return labelWithSource(result, result._src);

  // Phase C: CALENDAR guests (before body)
  result = searchCalendarForName(name, tokens, 200);
  if (result) return labelWithSource(result, result._src);

  // Phase D: BODY harvest
  result = searchBodiesForEmail(name, tokens, 30);
  if (result) return labelWithSource(result, 'body');

  return { best: null, status: "Not found", others: "" };
}

function resolveEmailCascadeWithStatus(name, statusRow) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const tokens = name.split(/\s+/).filter(Boolean).map(s => s.toLowerCase());

  // Phase A: FROM
  sheet.getRange(statusRow, 3).setValue("Searching FROM headers...");
  SpreadsheetApp.flush();
  let result = searchHeadersForName(name, tokens, 'from:', 30);
  if (result) return labelWithSource(result, 'from');

  // Phase B: TO
  sheet.getRange(statusRow, 3).setValue("Searching TO headers...");
  SpreadsheetApp.flush();
  result = searchHeadersForName(name, tokens, 'to:', 30);
  if (result) return labelWithSource(result, result._src);

  // Phase B: CC
  sheet.getRange(statusRow, 3).setValue("Searching CC headers...");
  SpreadsheetApp.flush();
  result = searchHeadersForName(name, tokens, 'cc:', 30);
  if (result) return labelWithSource(result, result._src);

  // Phase B.5: BCC
  sheet.getRange(statusRow, 3).setValue("Searching BCC headers...");
  SpreadsheetApp.flush();
  result = searchHeadersForName(name, tokens, 'bcc:', 30);
  if (result) return labelWithSource(result, result._src);

  // Phase C: CALENDAR
  sheet.getRange(statusRow, 3).setValue("Searching Calendar guests...");
  SpreadsheetApp.flush();
  result = searchCalendarForName(name, tokens, 200);
  if (result) return labelWithSource(result, result._src);

  // Phase D: BODY
  sheet.getRange(statusRow, 3).setValue("Scanning message bodies...");
  SpreadsheetApp.flush();
  result = searchBodiesForEmail(name, tokens, 30);
  if (result) return labelWithSource(result, 'body');

  return { best: null, status: "Not found", others: "" };
}

function labelWithSource(result, src) {
  // Map score to confidence, annotate with source; include alternates
  const s = result._score;
  let level;
  if (s >= 20) level = "High";
  else if (s >= 10) level = "Medium";
  else level = "Low";

  const out = { best: result.email, status: `${level} confidence (${src}: ${s.toFixed(1)})` };

  if (result._alts && result._alts.length) {
    const formatted = result._alts
      .filter(a => /@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/.test(a.email))
      .slice(0, 5)
      .map(a => `${a.email} [${a.score.toFixed(1)}]`)
      .join(", ");
    out.others = formatted;
  }
  return out;
}

// -------- Phase A/B: header-based search (from:/to:/cc:/bcc:) --------
function searchHeadersForName(name, tokens, field, maxThreads) {
  const qName = `"${name}"`;
  const last = deriveLast(tokens);
  const first = tokens[0];

  // Search both "First Last" and "Last, First" in the header field
  const query = `${field}(${qName} OR "${last}, ${first}") newer_than:5y -noreply -docs.google.com -calendar.google.com`;
  const threads = rateLimitedGmailSearch(query, 0, maxThreads);
  if (!threads.length) return null;

  // Collect self emails (primary + aliases) to detect outbound
  let selfEmails = [];
  try {
    const primary = (Session.getEffectiveUser && Session.getEffectiveUser().getEmail()) || "";
    const aliases = (GmailApp.getAliases && GmailApp.getAliases()) || [];
    selfEmails = [primary].concat(aliases).filter(Boolean).map(s => s.toLowerCase());
  } catch (e) { selfEmails = []; }

  const candidates = {}; // email -> {email, score, date, bump?}
  for (let t = 0; t < threads.length; t++) {
    const msgs = threads[t].getMessages();
    for (let m = 0; m < msgs.length; m++) {
      let header = "";
      if (field === 'from:')      header = msgs[m].getFrom();
      else if (field === 'to:')   header = msgs[m].getTo();
      else if (field === 'cc:')   header = msgs[m].getCc();
      else if (field === 'bcc:')  header = msgs[m].getBcc();
      if (!header) continue;

      // Split by commas not inside angle brackets
      const parts = header.split(/,(?![^<]*>)/);
      for (let p = 0; p < parts.length; p++) {
        const raw = parts[p].trim();
        const match = raw.match(/^(.*)<(.+?)>$/);
        let disp = "", email = "";
        if (match) { disp = (match[1] || "").trim().toLowerCase(); email = match[2]; }
        else { email = raw; disp = raw.toLowerCase(); }

        if (!email || isJunkEmail(email)) continue;
        // Require actual email format
        if (!/@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/.test(email)) continue;

        // Acceptance gates
        const overlap = tokens.filter(tok => disp.includes(tok));
        let accept = overlap.length > 0;
        if (!accept) {
          const localRaw = (email.split('@')[0] || "").toLowerCase();
          const nlocal = localRaw.replace(/[.\-_' ]/g, "");
          const lastSimple = tokens[tokens.length - 1];
          const lastCompound = deriveLast(tokens);
          const strong = (lastTok) =>
            ((nlocal.includes(first) && nlocal.includes(lastTok)) ||
             nlocal.startsWith(first[0] + lastTok) ||
             nlocal.startsWith(lastTok + first[0]));
          if (strong(lastSimple) || strong(lastCompound)) accept = true;
        }
        if (!accept) continue;

        // Base score (consider compound surname variant too)
        const tokensCompound = tokens.slice(0, -1).concat([deriveLast(tokens)]);
        const base = Math.max(
          scoreEmailAgainstName(email, tokens),
          scoreEmailAgainstName(email, tokensCompound)
        );

        const date = msgs[m].getDate().getTime();
        const rec = recencyBonus(date);
        const recWeighted = rec * 1.5; // heavier recency in headers

        // Display-name overlap bonus
        const dispHits = tokens.filter(tok => disp.includes(tok)).length;
        const dispBonus = dispHits >= 2 ? 4 : (dispHits === 1 ? 2 : 0);

        const nlocal = (email.split('@')[0] || "").toLowerCase().replace(/[.\-_' ]/g, "");
        const lastSimple = tokens[tokens.length - 1];
        const lastCompound = deriveLast(tokens);
        const hasFirstLocal = nlocal.includes(first);
        const hasLastLocal = nlocal.includes(lastSimple) || nlocal.includes(lastCompound);
        const hasLocalOverlap = hasFirstLocal || hasLastLocal;

        // Outbound detection (FROM matches self OR BCC phase)
        let fromEmail = "";
        try {
          const fraw = msgs[m].getFrom() || "";
          const mfrom = fraw.match(/<(.+?)>/);
          fromEmail = (mfrom ? mfrom[1] : fraw).trim().toLowerCase();
        } catch (e) {}
        const isOutbound = (selfEmails.indexOf(fromEmail) !== -1) || (field === 'bcc:');

        // Opaque-local exact-display boost (no local overlap)
        const dispExactNoLocalBoost = (dispHits >= 2 && !hasLocalOverlap)
          ? (isOutbound ? 8 : 4)
          : 0;

        // Outbound TO/CC channel bonus
        const channelOutboundBonus = (isOutbound && (field === 'to:' || field === 'cc:')) ? 4 : 0;

        // Header body-corroboration for partial matches (same message)
        let corroborationBonus = 0;
        const needFirst = (!hasFirstLocal && !disp.includes(first));
        const needLast  = (!hasLastLocal && !(disp.includes(lastSimple) || disp.includes(lastCompound)));
        if (needFirst || needLast) {
          let text = "";
          try { text = msgs[m].getPlainBody() || ""; } catch (e) { text = ""; }
          if (!text) {
            try { text = (msgs[m].getBody() || "").replace(/<[^>]*>/g, " "); } catch (e) {}
          }
          const low = (text || "").toLowerCase();
          if ((needFirst && low.indexOf(first) !== -1) ||
              (needLast && (low.indexOf(lastSimple) !== -1 || low.indexOf(lastCompound) !== -1))) {
            corroborationBonus = 4;
          }
        }

        // Multi-hit recency bump (headers only): +2 per <=1y, +1 per <=3y, cap +6 per email
        const inc = (rec >= 6 ? 2 : (rec >= 3 ? 1 : 0));
        const prev = candidates[email];
        const prevB = prev ? (prev.bump || 0) : 0;
        const newBump = Math.min(6, prevB + inc);
        const scored = base + recWeighted + dispBonus + dispExactNoLocalBoost + channelOutboundBonus + corroborationBonus + newBump;

        if (!candidates[email] || scored > candidates[email].score) {
          candidates[email] = { email, score: scored, date, bump: newBump };
        } else {
          const delta = newBump - prevB;
          if (delta > 0) candidates[email].score += delta;
          candidates[email].bump = newBump;
        }
      }
    }
  }

  const list = Object.values(candidates);
  if (!list.length) return null;

  list.sort((a, b) => (b.score === a.score) ? (b.date - a.date) : (b.score - a.score));
  const best = list[0];
  best._score = best.score;
  best._src = field.replace(':','');
  best._alts = list.slice(1, 6).map(c => ({ email: c.email, score: c.score }));

  return best;
}

// -------- Phase D (after calendar): body-based search --------
function searchBodiesForEmail(name, tokens, maxThreads) {
  const q = `"${name}" newer_than:5y -noreply -docs.google.com -calendar.google.com`;
  const threads = rateLimitedGmailSearch(q, 0, maxThreads);
  if (!threads.length) return null;

  const lastSimple = tokens[tokens.length - 1];
  const lastCompound = deriveLast(tokens);
  const first = tokens[0];
  const reEmail = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g;

  const candidates = {};
  for (let t = 0; t < threads.length; t++) {
    const msgs = threads[t].getMessages();
    for (let m = 0; m < msgs.length; m++) {
      const date = msgs[m].getDate().getTime();

      // prefer plainText, fallback to HTML-stripped
      let text = "";
      try { text = msgs[m].getPlainBody() || ""; } catch(e) { text = ""; }
      if (!text) {
        try {
          const html = msgs[m].getBody() || "";
          text = html.replace(/<[^>]*>/g, " ");
        } catch(e) {}
      }
      if (!text) continue;

      const found = text.match(reEmail);
      if (!found) continue;

      for (let k = 0; k < found.length; k++) {
        const email = found[k];
        if (isJunkEmail(email)) continue;

        const local = email.split('@')[0].toLowerCase();
        const domain = (email.split('@')[1] || "").toLowerCase();

        // keep only if last name appears in local or domain (supports compound)
        const lastHit = local.includes(lastSimple) || domain.includes(lastSimple) ||
                        local.includes(lastCompound) || domain.includes(lastCompound);
        if (!lastHit) continue;

        const tokensCompound = tokens.slice(0, -1).concat([deriveLast(tokens)]);
        const base = Math.max(
          scoreEmailAgainstName(email, tokens),
          scoreEmailAgainstName(email, tokensCompound)
        );
        const firstInitHit = local.startsWith(first[0]);
        const initBump = firstInitHit ? 2 : 0;
        const rec = recencyBonus(date);
        const total = base + initBump + rec - 2;

        if (!candidates[email] || total > candidates[email].score) {
          candidates[email] = { email, score: total, date };
        }
      }
    }
  }

  const list = Object.values(candidates);
  if (!list.length) return null;

  list.sort((a, b) => (b.score === a.score) ? (b.date - a.date) : (b.score - a.score));
  const best = list[0];
  best._score = best.score;
  best._src = 'body';
  best._alts = list.slice(1, 6).map(c => ({ email: c.email, score: c.score }));
  return best;
}

// -------- Phase C: calendar-based search (before body) --------
function searchCalendarForName(name, tokens, maxEvents) {
  const cal = CalendarApp.getDefaultCalendar();
  const end = new Date();
  const start = new Date();
  start.setFullYear(end.getFullYear() - 5);

  const events = cal.getEvents(start, end);
  if (!events || !events.length) return null;

  const candidates = {};
  const first = tokens[0];
  const lastSimple = tokens[tokens.length - 1];
  const lastCompound = deriveLast(tokens);

  let processed = 0;
  for (let i = 0; i < events.length; i++) {
    if (processed >= maxEvents) break;
    processed++;

    const ev = events[i];
    const when = ev.getStartTime() ? ev.getStartTime().getTime() : ev.getLastUpdated().getTime();
    const guests = (ev.getGuestList && ev.getGuestList(true)) || [];
    if (!guests || !guests.length) continue;

    for (let g = 0; g < guests.length; g++) {
      const ge = guests[g];
      const gEmail = (ge.getEmail && ge.getEmail()) || "";
      const gName  = ((ge.getName && ge.getName()) || "").toLowerCase();
      if (!gEmail || isJunkEmail(gEmail)) continue;

      // Acceptance gates (word-level for display name)
      const nameWords = gName.split(/[^a-z]+/).filter(Boolean);
      const overlap = tokens.filter(tok => nameWords.indexOf(tok) !== -1);
      let accept = overlap.length > 0;
      if (!accept) {
        const nlocal = (gEmail.split('@')[0] || "").toLowerCase().replace(/[.\-_' ]/g, "");
        const strongLocal =
          ((nlocal.includes(first) && nlocal.includes(lastSimple)) ||
           (nlocal.includes(first) && nlocal.includes(lastCompound)) ||
           nlocal.startsWith(first[0] + lastSimple) ||
           nlocal.startsWith(lastSimple + first[0]) ||
           nlocal.startsWith(first[0] + lastCompound) ||
           nlocal.startsWith(lastCompound + first[0]));
        if (strongLocal) accept = true;
      }
      if (!accept) continue;

      const tokensCompound = tokens.slice(0, -1).concat([deriveLast(tokens)]);
      const base = Math.max(
        scoreEmailAgainstName(gEmail, tokens),
        scoreEmailAgainstName(gEmail, tokensCompound)
      );
      const rec = recencyBonus(when);
      const dispHits = tokens.filter(tok => nameWords.indexOf(tok) !== -1).length;
      const dispBonus = dispHits >= 2 ? 4 : (dispHits === 1 ? 2 : 0);

      const nlocal2 = (gEmail.split('@')[0] || "").toLowerCase().replace(/[.\-_' ]/g, "");
      const hasLocalOverlap = nlocal2.includes(first) || nlocal2.includes(lastSimple) || nlocal2.includes(lastCompound);
      const opaqueExactDisplayBoost = (dispHits >= 2 && !hasLocalOverlap) ? 4 : 0;

      // Participant bonus scaled by recency (based on recencyBonus result)
      let participantBonus = 0;
      if (rec >= 6) participantBonus = 10;      // <=1y
      else if (rec >= 3) participantBonus = 5;  // <=3y

      // Severe penalty when local part lacks first/last evidence; also dampen participant bonus
      let localNamePenalty = 0;
      if (!hasLocalOverlap) {
        localNamePenalty = -8;
        if (participantBonus > 0) participantBonus = Math.max(0, participantBonus - 5);
      }

      const total = base + rec + dispBonus + opaqueExactDisplayBoost + participantBonus + localNamePenalty;

      if (!candidates[gEmail] || total > candidates[gEmail].score) {
        candidates[gEmail] = { email: gEmail, score: total, date: when };
      }
    }
  }

  const list = Object.values(candidates);
  if (!list.length) return null;

  list.sort((a, b) => (b.score === a.score) ? (b.date - a.date) : (b.score - a.score));
  const best = list[0];
  best._score = best.score;
  best._src = 'calendar';
  best._alts = list.slice(1, 6).map(c => ({ email: c.email, score: c.score }));
  return best;
}

// -------- Scoring & helpers --------
function recencyBonus(dateMs) {
  const years = (Date.now() - dateMs) / (1000 * 60 * 60 * 24 * 365);
  if (years <= 1) return 6;
  if (years <= 3) return 3;
  return 0;
}

function scoreEmailAgainstName(email, tokens) {
  const parts = email.split('@');
  const local = norm(parts[0]);
  const domain = (parts[1] || "").toLowerCase();
  const first = norm(tokens[0]);
  const last  = norm(tokens[tokens.length - 1]);

  let score = 0;

  // Perfect patterns
  if (local === first + last || local === last + first) score += 20;
  if (local === first + "." + last || local === last + "." + first) score += 20;

  // Both names present
  if (local.includes(first) && local.includes(last)) score += 12;

  // First-only / last-only strong signals
  if (local === first) { score += domain.includes(last) ? 10 : 6; }
  else if (local.startsWith(first)) score += 8;

  if (local === last) { score += domain.includes(first) ? 10 : 6; }
  else if (local.startsWith(last)) score += 8;

  // Initial+last / last+initial
  if (local.startsWith(first[0] + last)) score += 8;
  if (local.startsWith(last + first[0])) score += 8;

  // Partial weak bumps
  if (local.includes(first)) score += 4;
  if (local.includes(last)) score += 4;

  // Penalties for noise
  if (/\d{3,}/.test(local)) score -= 3;
  if (/[_.-]{2,}/.test(local)) score -= 2;

  return score;
}

function norm(s) {
  return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "")
          .toLowerCase().replace(/['’\-\._]/g,"").trim();
}

function deriveLast(tokens) {
  if (tokens.length >= 2) {
    const t2 = tokens[tokens.length - 2];
    if (["de","da","del","la","van","von"].includes(t2)) {
      return t2 + tokens[tokens.length - 1];
    }
  }
  return tokens[tokens.length - 1];
}

function isJunkEmail(email) {
  const lower = email.toLowerCase();
  const parts = lower.split("@");
  const local = parts[0];
  const domain = parts[1] || "";

  if (/(?:^|[^a-z])(no[\-_\.]?reply|do[\-_\.]?not[\-_\.]?reply|mailer[-_]?daemon|bounces?)(?:[^a-z]|$)/.test(local)) {
    return true;
  }
  if (/^notifications?$/.test(local)) {
    return true;
  }
  if (domain.includes("docs.google.com") || domain.includes("calendar.google.com")) {
    return true;
  }
  return false;
}
