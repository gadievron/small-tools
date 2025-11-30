/**
 * Gmail bounce extractor to Google Sheets.
 *
 * Notes:
 * - Uses GmailApp (search, from, subject, plain body/html body).
 * - Detection is based on FROM, SUBJECT, and BODY text (including 5xx codes).
 * - Extraction prefers the original sent message (To/Cc/Bcc); falls back to body anchors.
 * - Does NOT parse raw MIME or DSN headers (no X-Failed-Recipients, Final-Recipient, etc.).
 * 
 * By: Gadi Evron (with ChatGPT)
 * License: MIT
 * 
 */

/**** CONFIG ****/
const SHEET_NAME  = 'Bounced';  // Sheet tab to dump results into
const SENDER_EMAIL = '';        // Optional: hard-code your sender. If empty, uses Session.getActiveUser()

/**** PATTERNS & REGEXES ****/

// DSN-like senders
const FROM_INDICATORS = [
  'mailer-daemon',
  'mail delivery subsystem',
  'postmaster'
];

// Bounce-ish subjects (hard-bounce oriented, real-world)
const SUBJECT_INDICATORS = [
  'delivery status notification (failure)',
  'undeliverable',                       // covers "Undeliverable:" subject too
  'delivery failure',
  'delivery has failed',
  'message not delivered',
  'returned mail',
  'failure notice',
  'undelivered mail returned to sender',
  'mail delivery failed',
  'mail delivery problem'
];

// Body phrases that indicate a hard bounce and are useful anchors for extraction
const BODY_INDICATOR_PHRASES = [
  // permanent / generic
  'permanent fatal errors',
  'delivery has failed',
  'message not delivered',
  'address not found',

  // recipient does not exist
  'no such user',
  'user unknown',
  'recipient not found',
  'recipient address rejected',
  'invalid recipient',
  'mailbox not found',
  'mailbox unavailable',
  'mailbox disabled',
  'mailbox does not exist',
  'unrouteable address',
  'unknown recipient',
  'the email account that you tried to reach does not exist',

  // domain issues (specific dns failures only)
  'domain not found',
  'domain does not exist',
  'unrouteable domain',
  'no mx record',
  'dns domain does not exist',
  'dns error: dns domain',
  'dns error: dns type \'mx\' lookup of',

  // outlook / ms specific wordings
  'your message wasn\'t delivered to',
  'your message couldn\'t be delivered',
  'the recipient\'s email system rejected your message',

  // resolver family (Exchange / Cisco ESA / O365)
  'resolver.adr.',
  'recipient not found by smtp address lookup'
];

// Numeric SMTP-style hard-bounce codes (5xx only; no 4xx)
const CODE_REGEXES = [
  // generic “5xx 5.x.x” enhanced codes (e.g. 550 5.1.1, 554-5.7.1)
  /5[0-9]{2}(\s*[\-#:])?\s*5\.[0-9]\.[0-9]+/i,

  // bare 550 / 554 in body (permanent failures)
  /550(?![0-9])/i,
  /554(?![0-9])/i,

  // specific enhanced codes we care about
  /5\.1\.1\b/i,
  /5\.4\.310\b/i
];

// Email pattern (shared)
const EMAIL_REGEX = /[a-zA-Z0-9._%+\-']+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;

/**** MENU ****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Bounce Cleaner')
    .addItem('Scan bounces…', 'scanBouncesMenu')
    .addToUi();
}

/**
 * Menu entry: ask "how many days back", default 1 (last 24 hours).
 */
function scanBouncesMenu() {
  const ui = SpreadsheetApp.getUi();
  const defaultDays = 1;

  const response = ui.prompt(
    'Scan Gmail bounces',
    'Search how many days back? (1 = last 24 hours)\nLeave blank for 1 day.',
    ui.ButtonSet.OK_CANCEL
  );

  const button = response.getSelectedButton();
  if (button !== ui.Button.OK) {
    return; // user cancelled
  }

  const txt = (response.getResponseText() || '').trim();
  let days = parseInt(txt, 10);

  if (isNaN(days) || days <= 0) {
    days = defaultDays;
  }

  scanBouncesWithDays_(days);
}

/**
 * Direct helper: always last 24h (1 day).
 */
function scanBouncesLast24h() {
  scanBouncesWithDays_(1);
}

/**** MAIN SCAN IMPLEMENTATION ****/
function scanBouncesWithDays_(daysBack) {
  const sheet = getOrCreateSheet_(SHEET_NAME);
  initHeaderRow_(sheet);

  // Build Gmail time filter from daysBack
  const timeFilter = 'newer_than:' + daysBack + 'd';
  // Include Inbox + Spam + Trash
  const query = timeFilter + ' in:anywhere';

  const threads = GmailApp.search(query);
  Logger.log(
    'scanBouncesWithDays_: daysBack=' + daysBack +
    ', threads=' + threads.length +
    ', query="' + query + '"'
  );

  const senderEmail =
    SENDER_EMAIL ||
    (Session.getActiveUser() && Session.getActiveUser().getEmail()) ||
    '';

  const index = buildExistingIndex_(sheet);
  const existingEmails = index.emails;
  const rowByEmail = index.rowByEmail;

  let bounceMessagesCount = 0;
  let rowsCreated = 0;
  let rowsUpdated = 0;

  for (var t = 0; t < threads.length; t++) {
    const thread = threads[t];
    const messages = thread.getMessages();
    const threadUrl = 'https://mail.google.com/mail/u/0/#all/' + thread.getId();

    for (var m = 0; m < messages.length; m++) {
      const msg = messages[m];

      const from = msg.getFrom() || '';
      const subject = msg.getSubject() || '';
      const body = safePlainBody_(msg);

      if (!isBounceMessage_(from, subject, body)) continue;

      bounceMessagesCount++;

      // 1) Prefer original message we sent in this thread
      const originalMessage = findOriginalSentMessage_(messages, senderEmail);
      let bouncedEmails = [];

      if (originalMessage) {
        bouncedEmails = extractEmailsFromOriginal_(originalMessage, senderEmail);
      }

      // 2) Fallback ONLY if we can't locate original recipients
      if (bouncedEmails.length === 0) {
        bouncedEmails = extractEmailsFromBounceBody_(body);
      }

      if (bouncedEmails.length === 0) {
        Logger.log('Bounce detected but no emails extracted. Subject="' + subject + '", From="' + from + '"');
        continue;
      }

      // 3) Resolve onmicrosoft aliases only if needed
      bouncedEmails = resolveOnMicrosoftAliases_(bouncedEmails, senderEmail);
      if (bouncedEmails.length === 0) {
        Logger.log('Bounce after resolution but no usable emails. Subject="' + subject + '", From="' + from + '"');
        continue;
      }

      const uniqueEmails = Array.from(new Set(bouncedEmails));
      const now = new Date();

      uniqueEmails.forEach(function(rawEmail) {
        const email = (rawEmail || '').trim();
        if (!email) return;

        const key = email.toLowerCase();

        if (existingEmails.has(key)) {
          // UPDATE existing row: Last Seen, Bounce Count, latest context
          const info = rowByEmail[key];
          if (!info || !info.row) return;

          const rowNum = info.row;
          const currentCount = info.count || 1;
          const newCount = currentCount + 1;

          sheet.getRange(rowNum, 6).setValue(now);       // Last Seen
          sheet.getRange(rowNum, 7).setValue(newCount);  // Bounce Count
          sheet.getRange(rowNum, 3).setValue(subject);   // Bounce Subject (latest)
          sheet.getRange(rowNum, 4).setValue(from);      // Bounce From (latest)
          sheet.getRange(rowNum, 5).setValue(threadUrl); // Gmail Thread URL (latest)

          rowByEmail[key].count = newCount;
          rowsUpdated++;
        } else {
          // CREATE new row for this email
          const row = [
            now,        // Timestamp (first seen)
            email,      // Bounced Email
            subject,    // Bounce Subject
            from,       // Bounce From
            threadUrl,  // Gmail Thread URL
            now,        // Last Seen
            1           // Bounce Count
          ];

          sheet.appendRow(row);
          const newRowNum = sheet.getLastRow();

          existingEmails.add(key);
          rowByEmail[key] = { row: newRowNum, count: 1 };
          rowsCreated++;
        }
      });
    }
  }

  Logger.log(
    'scanBouncesWithDays_ summary: daysBack=' + daysBack +
    ', bounceMessagesCount=' + bounceMessagesCount +
    ', rowsCreated=' + rowsCreated +
    ', rowsUpdated=' + rowsUpdated
  );

  SpreadsheetApp.getActive().toast(
    'Scan done (last ' + daysBack + ' day' + (daysBack === 1 ? '' : 's') + '): ' +
    threads.length + ' threads, ' +
    bounceMessagesCount + ' bounce messages, ' +
    rowsCreated + ' new emails, ' +
    rowsUpdated + ' updated.'
  );
}

/**** HELPERS: SHEET ****/

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function initHeaderRow_(sheet) {
  const header = [
    'Timestamp',         // First seen
    'Bounced Email',
    'Bounce Subject',
    'Bounce From',
    'Gmail Thread URL',
    'Last Seen',
    'Bounce Count'
  ];

  sheet.getRange(1, 1, 1, header.length).setValues([header]);
}

/**
 * Build an index of existing emails:
 *   emails: Set<emailLower>
 *   rowByEmail: { emailLower -> { row, count } }
 */
function buildExistingIndex_(sheet) {
  const emails = new Set();
  const rowByEmail = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { emails: emails, rowByEmail: rowByEmail };
  }

  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  for (var i = 0; i < data.length; i++) {
    const r = data[i];
    const rowNum = i + 2;

    const email = (r[1] || '').toString().trim().toLowerCase(); // column B = Bounced Email
    if (!email) continue;

    let countVal = r[6]; // column G = Bounce Count (0-based index 6)
    let count = parseInt(countVal, 10);
    if (isNaN(count) || count <= 0) count = 1; // default

    emails.add(email);
    rowByEmail[email] = { row: rowNum, count: count };
  }

  return { emails: emails, rowByEmail: rowByEmail };
}

/**** TEXT NORMALIZATION ****/

/**
 * Normalize text for matching:
 * - lowercase
 * - convert curly apostrophes to straight (’ / ‘ -> ')
 */
function normalizeText_(text) {
  if (!text) return '';
  return text
    .toLowerCase()
    .replace(/\u2019/g, "'")
    .replace(/\u2018/g, "'");
}

/**** HELPERS: GMAIL / BOUNCE DETECTION ****/

function safePlainBody_(msg) {
  try {
    const plain = msg.getPlainBody();
    if (plain) return plain;
  } catch (e) {}
  try {
    return msg.getBody();
  } catch (e2) {}
  return '';
}

/**
 * Short-circuit bounce detection:
 *   FROM -> SUBJECT -> BODY -> CODES
 * Stop at first match.
 */
function isBounceMessage_(from, subject, body) {
  const normFrom = normalizeText_(from);
  const normSubject = normalizeText_(subject);
  const normBody = normalizeText_(body);

  // 1. FROM
  if (FROM_INDICATORS.some(function(p) { return normFrom.indexOf(p) !== -1; })) {
    return true;
  }

  // 2. SUBJECT
  if (SUBJECT_INDICATORS.some(function(p) { return normSubject.indexOf(p) !== -1; })) {
    return true;
  }

  // 3. BODY phrases
  if (BODY_INDICATOR_PHRASES.some(function(p) { return normBody.indexOf(p) !== -1; })) {
    return true;
  }

  // 4. Numeric error codes (use raw body, numeric only)
  for (var i = 0; i < CODE_REGEXES.length; i++) {
    if (CODE_REGEXES[i].test(body)) {
      return true;
    }
  }

  return false;
}

/**
 * Extract emails from bounce *body* around clear indicator phrases OR code lines.
 */
function extractEmailsFromBounceBody_(body) {
  const emails = new Set();
  const lines = (body || '').split(/\r?\n/);

  for (var i = 0; i < lines.length; i++) {
    const line = lines[i];
    const normLine = normalizeText_(line);

    const phraseHit = BODY_INDICATOR_PHRASES.some(function(p) {
      return normLine.indexOf(p) !== -1;
    });

    let codeHit = false;
    if (!phraseHit) {
      for (var k = 0; k < CODE_REGEXES.length; k++) {
        if (CODE_REGEXES[k].test(line)) {
          codeHit = true;
          break;
        }
      }
    }

    if (!phraseHit && !codeHit) continue;

    // Scan this line and next 3 lines for emails (use original lines for regex)
    for (var j = i; j < Math.min(i + 4, lines.length); j++) {
      const candidateLine = lines[j];
      const matches = candidateLine.match(EMAIL_REGEX);
      if (matches) {
        matches.forEach(function(e) {
          if (!isLikelySystemAddress_(e)) {
            emails.add(e);
          }
        });
      }
    }
  }

  return Array.from(emails);
}

/**
 * Filter out obvious infra/system addresses:
 * - postmaster / mailer-daemon
 * - mail.gmail.com, 1e100.net, mail-*.google.com
 * - internal routing domains like ant.amazon.com
 */
function isLikelySystemAddress_(email) {
  const lower = email.toLowerCase();
  const parts = lower.split('@');
  if (parts.length !== 2) return false;

  const local = parts[0];
  const domain = parts[1];

  if (local === 'postmaster' || local === 'mailer-daemon') return true;

  if (domain === 'mail.gmail.com') return true;
  if (domain === '1e100.net') return true;
  if (domain.indexOf('mail-') === 0 && domain.indexOf('.google.com') !== -1) return true;

  // Amazon internal routing example
  if (domain === 'ant.amazon.com') return true;

  return false;
}

/**
 * Try to find the original message we sent in a thread.
 */
function findOriginalSentMessage_(messages, senderEmail) {
  const lowerSender = (senderEmail || '').toLowerCase();
  let candidate = null;

  for (var i = 0; i < messages.length; i++) {
    const msg = messages[i];
    const from = (msg.getFrom() || '').toLowerCase();

    // Skip system senders
    if (
      from.indexOf('mailer-daemon') !== -1 ||
      from.indexOf('mail delivery subsystem') !== -1 ||
      from.indexOf('postmaster') !== -1
    ) {
      continue;
    }

    if (lowerSender && from.indexOf(lowerSender) !== -1) {
      return msg; // best case: exact match to us
    }

    // Fallback: remember first non-system message
    if (!candidate) {
      candidate = msg;
    }
  }
  return candidate;
}

/**
 * Extract candidate recipient emails from original message To/Cc/Bcc.
 */
function extractEmailsFromOriginal_(msg, senderEmail) {
  const fields = [
    msg.getTo() || '',
    msg.getCc() || '',
    msg.getBcc() || ''
  ].join(',');

  const matches = fields.match(EMAIL_REGEX) || [];
  const lowerSender = (senderEmail || '').toLowerCase();

  const filtered = matches.filter(function(e) {
    const lower = e.toLowerCase();
    if (lower === lowerSender) return false;
    if (isLikelySystemAddress_(e)) return false;
    return true;
  });

  return Array.from(new Set(filtered));
}

/**** onmicrosoft alias resolution ****/

/**
 * For addresses like local@tenant.onmicrosoft.com,
 * search last 24h sent mail for real addresses:
 *   local@*tenant*
 */
function resolveOnMicrosoftAliases_(emails, senderEmail) {
  const resolved = new Set();

  emails.forEach(function(address) {
    if (!address) return;
    const email = address.trim();

    const m = email.match(/^([a-zA-Z0-9._%+\-']+)@([a-zA-Z0-9.\-]+)\.onmicrosoft\.com$/i);
    if (!m) {
      resolved.add(email); // not an alias
      return;
    }

    const localPart = m[1];     // e.g. "vadalad"
    const tenantToken = m[2];   // e.g. "moodys"

    const originals = findOriginalRecipientsForAlias_(localPart, tenantToken, senderEmail);

    if (originals.length > 0) {
      originals.forEach(function(o) { resolved.add(o); });
    } else {
      resolved.add(email); // better to keep alias than drop it
    }
  });

  return Array.from(resolved);
}

/**
 * Search last 24h sent mail for recipients whose:
 *   - local part == localPart
 *   - domain contains tenantToken (e.g. "moodys")
 */
function findOriginalRecipientsForAlias_(localPart, tenantToken, senderEmail) {
  localPart = (localPart || '').toLowerCase();
  tenantToken = (tenantToken || '').toLowerCase();
  if (!localPart && !tenantToken) return [];

  const queryParts = ['newer_than:1d', 'from:me', 'in:anywhere'];
  if (localPart) queryParts.push(localPart);
  if (tenantToken) queryParts.push(tenantToken);

  const query = queryParts.join(' ');
  Logger.log('findOriginalRecipientsForAlias_: query="' + query + '"');

  let threads;
  try {
    threads = GmailApp.search(query);
  } catch (e) {
    Logger.log('findOriginalRecipientsForAlias_ search failed: ' + e);
    return [];
  }

  const originals = new Set();
  const lowerSender = (senderEmail || '').toLowerCase();

  for (var t = 0; t < threads.length; t++) {
    const messages = threads[t].getMessages();
    for (var m = 0; m < messages.length; m++) {
      const msg = messages[m];
      const from = (msg.getFrom() || '').toLowerCase();

      if (lowerSender && from.indexOf(lowerSender) === -1) {
        continue;
      }

      const fields = [
        msg.getTo() || '',
        msg.getCc() || '',
        msg.getBcc() || ''
      ].join(',');

      const matches = fields.match(EMAIL_REGEX) || [];
      matches.forEach(function(addr) {
        const lowerAddr = addr.toLowerCase();
        const parts = lowerAddr.split('@');
        if (parts.length !== 2) return;
        const lp = parts[0];
        const domain = parts[1];

        if (lp === localPart && domain.indexOf(tenantToken) !== -1) {
          if (!lowerSender || lowerAddr !== lowerSender) {
            originals.add(addr);
          }
        }
      });
    }
  }

  const result = Array.from(originals);
  Logger.log('findOriginalRecipientsForAlias_: found originals=' + JSON.stringify(result));
  return result;
}
