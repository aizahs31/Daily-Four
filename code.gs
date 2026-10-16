/**
 * Daily Four: Gmail -> Google Sheet
 *
 * Reads your coach's daily emails, finds the LeetCode problem links in them,
 * looks up each problem's title and difficulty, and adds one row per problem
 * to the "Problems" tab: date | problem | link | difficulty.
 *
 * Setup (once):
 *   1. Fill in CONFIG below (at least SENDER).
 *   2. Run backfill() to import every past email. Approve the permissions.
 *      If it stops at the time limit, run it again. It never adds duplicates.
 *   3. Run setup() to schedule the daily run.
 */

var CONFIG = {
  SENDER: 'squad83@kalvium.community',   // the coach's email address (required)
  SUBJECT_CONTAINS: 'Workout Exercise',          // optional, e.g. 'LeetCode'. Leave '' to accept any subject
  SHEET_NAME: 'daily four',
  TIMEZONE: 'Asia/Kolkata',      // used for the "date sent" column
  LOOKBACK_DAYS: 3,              // how far back the daily run looks
  RUN_HOUR: 7,                   // daily run time, 24h clock in TIMEZONE
  ALERT_EMAIL: ''                // where failure alerts go. '' means your own address
};

var HEADERS = ['date', 'problem', 'link', 'difficulty'];
var TIME_BUDGET_MS = 5 * 60 * 1000;   // Apps Script stops at 6 minutes

// ---------- Entry points ----------

/** Import every past email from the coach. Safe to run more than once. */
function backfill() {
  checkConfig_();
  var summary = run_(buildQuery_(''), 1000);
  Logger.log(summary.text);
}

/** Runs every day from the trigger. */
function dailyRun() {
  try {
    checkConfig_();
    var summary = run_(buildQuery_('newer_than:' + CONFIG.LOOKBACK_DAYS + 'd'), 100);
    fixMissingDifficulty_(40);
    Logger.log(summary.text);
    if (summary.emailsWithoutProblems > 0) {
      alert_('Daily Four: could not find problems in ' + summary.emailsWithoutProblems + ' email(s)',
        'The coach\'s email arrived but no leetcode.com/problems links were found in it. ' +
        'The email format may have changed. Open the email and check how the problems are written.');
    }
  } catch (e) {
    alert_('Daily Four: the daily run failed', String(e && e.stack ? e.stack : e));
    throw e;
  }
}

/** Creates the Problems tab and schedules dailyRun(). Run once. */
function setup() {
  checkConfig_();
  ensureSheet_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyRun') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyRun').timeBased().everyDays(1)
    .atHour(CONFIG.RUN_HOUR).inTimezone(CONFIG.TIMEZONE).create();
  Logger.log('Scheduled dailyRun() for about ' + CONFIG.RUN_HOUR + ':00 ' + CONFIG.TIMEZONE + ' every day.');
}

// ---------- Core ----------

function run_(query, maxThreads) {
  var deadline = Date.now() + TIME_BUDGET_MS;
  var sheet = ensureSheet_();
  var existing = readExisting_(sheet);
  var senderNeedle = CONFIG.SENDER.toLowerCase();
  var perMessage = [];
  var emailsSeen = 0, emailsWithoutProblems = 0;

  var start = 0, page = 100;
  while (start < maxThreads && Date.now() < deadline) {
    var threads = GmailApp.search(query, start, page);
    if (!threads.length) break;
    threads.forEach(function (thread) {
      thread.getMessages().forEach(function (msg) {
        if (msg.isInTrash()) return;
        if (msg.getFrom().toLowerCase().indexOf(senderNeedle) < 0) return;
        emailsSeen++;
        var items = extractProblems_(msg.getBody() || '', msg.getPlainBody() || '');
        if (!items.length) { emailsWithoutProblems++; return; }
        perMessage.push({
          ts: msg.getDate().getTime(),
          date: Utilities.formatDate(msg.getDate(), CONFIG.TIMEZONE, 'yyyy-MM-dd'),
          items: items
        });
      });
    });
    start += page;
  }

  // Oldest email first, keeping each email's own problem order.
  perMessage.sort(function (a, b) { return a.ts - b.ts; });

  var newRows = [], seenKeys = {}, slugsToLookUp = [], queued = {};
  perMessage.forEach(function (m) {
    m.items.forEach(function (it) {
      var key = m.date + '|' + it.slug;
      if (existing.keys[key] || seenKeys[key]) return;
      seenKeys[key] = true;
      newRows.push({ date: m.date, slug: it.slug, hint: it.title });
      if (!existing.known[it.slug] && !queued[it.slug]) { queued[it.slug] = true; slugsToLookUp.push(it.slug); }
    });
  });

  var looked = lookupMany_(slugsToLookUp, deadline);
  var out = newRows.map(function (r) {
    var info = looked[r.slug] || existing.known[r.slug] || {};
    return [
      r.date,
      info.title || cleanTitle_(r.hint) || slugToTitle_(r.slug),
      'https://leetcode.com/problems/' + r.slug + '/',
      info.difficulty || ''
    ];
  });
  appendRows_(sheet, out);

  var blanks = out.filter(function (r) { return !r[3]; }).length;
  return {
    added: out.length,
    emailsSeen: emailsSeen,
    emailsWithoutProblems: emailsWithoutProblems,
    text: 'Emails from coach: ' + emailsSeen + '. New rows added: ' + out.length +
      (blanks ? ' (' + blanks + ' without difficulty, they are retried by the daily run)' : '') + '.' +
      (Date.now() >= deadline ? ' Time limit reached, run backfill() again to continue.' : '')
  };
}

/** Finds LeetCode problem links in an email, in order, without repeats. */
function extractProblems_(html, text) {
  var out = [], seen = {}, m;
  function add(slug, title) {
    slug = String(slug || '').toLowerCase();
    if (!slug || seen[slug]) return;
    seen[slug] = true;
    out.push({ slug: slug, title: title || '' });
  }
  var reA = /<a\s[^>]*href\s*=\s*["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  while ((m = reA.exec(html))) {
    var slug = slugFromUrl_(decodeHref_(m[1]));
    if (slug) add(slug, stripTags_(m[2]));
  }
  var reU = /leetcode\.(?:com|cn)\/problems\/([a-z0-9-]+)/gi;
  [html, decodeHref_(html), text].forEach(function (src) {
    reU.lastIndex = 0;
    while ((m = reU.exec(src))) add(m[1], '');
  });
  return out;
}

function lookupMany_(slugs, deadline) {
  var res = {};
  var query = 'query q($titleSlug: String!){ question(titleSlug: $titleSlug){ title difficulty } }';
  for (var i = 0; i < slugs.length; i += 8) {
    if (Date.now() > deadline) break;
    var chunk = slugs.slice(i, i + 8);
    var reqs = chunk.map(function (s) {
      return {
        url: 'https://leetcode.com/graphql',
        method: 'post',
        contentType: 'application/json',
        headers: { Referer: 'https://leetcode.com/problems/' + s + '/' },
        payload: JSON.stringify({ query: query, variables: { titleSlug: s } }),
        muteHttpExceptions: true
      };
    });
    var resps = UrlFetchApp.fetchAll(reqs);
    resps.forEach(function (r, j) {
      try {
        if (r.getResponseCode() !== 200) return;
        var q = JSON.parse(r.getContentText()).data.question;
        if (q && q.difficulty) res[chunk[j]] = { title: q.title, difficulty: q.difficulty };
      } catch (e) {}
    });
    Utilities.sleep(300);
  }
  return res;
}

/** Retries rows where the difficulty lookup failed earlier. */
function fixMissingDifficulty_(limit) {
  var sheet = ensureSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return;
  var vals = sheet.getRange(2, 1, last - 1, 4).getValues();
  var todo = [];
  vals.forEach(function (r, i) {
    if (!r[3] && todo.length < limit) {
      var slug = slugFromUrl_(String(r[2]));
      if (slug) todo.push({ row: i + 2, slug: slug });
    }
  });
  if (!todo.length) return;
  var looked = lookupMany_(todo.map(function (t) { return t.slug; }), Date.now() + TIME_BUDGET_MS);
  todo.forEach(function (t) {
    var info = looked[t.slug];
    if (info) { sheet.getRange(t.row, 2).setValue(info.title); sheet.getRange(t.row, 4).setValue(info.difficulty); }
  });
}

// ---------- Sheet helpers ----------

function ensureSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.getRange('A:A').setNumberFormat('@');   // keep dates as plain yyyy-mm-dd text
  return sh;
}

function readExisting_(sheet) {
  var keys = {}, known = {};
  var last = sheet.getLastRow();
  if (last < 2) return { keys: keys, known: known };
  sheet.getRange(2, 1, last - 1, 4).getValues().forEach(function (r) {
    var slug = slugFromUrl_(String(r[2]));
    if (!slug) return;
    keys[normDate_(r[0]) + '|' + slug] = true;
    if (r[3]) known[slug] = { title: String(r[1]), difficulty: String(r[3]) };
  });
  return { keys: keys, known: known };
}

function appendRows_(sheet, rows) {
  if (!rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
}

// ---------- Small utilities ----------

function buildQuery_(extra) {
  var q = 'from:' + CONFIG.SENDER;
  if (CONFIG.SUBJECT_CONTAINS) q += ' subject:(' + CONFIG.SUBJECT_CONTAINS + ')';
  return (q + ' ' + extra).trim();
}

function checkConfig_() {
  if (!CONFIG.SENDER || CONFIG.SENDER.indexOf('example.com') >= 0) {
    throw new Error('Set CONFIG.SENDER to your coach\'s email address first.');
  }
}

function slugFromUrl_(s) {
  var m = String(s).match(/\/problems\/([a-z0-9-]+)/i);
  return m ? m[1].toLowerCase() : '';
}

function decodeHref_(h) {
  var s = String(h).replace(/&amp;/g, '&');
  for (var i = 0; i < 2; i++) { try { s = decodeURIComponent(s); } catch (e) { break; } }
  return s;
}

function stripTags_(s) {
  return String(s).replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/\s+/g, ' ').trim();
}

function cleanTitle_(t) {
  t = String(t || '').trim();
  return /^https?:\/\//i.test(t) || /leetcode\.(com|cn)/i.test(t) ? '' : t;
}

function slugToTitle_(slug) {
  return slug.split('-').map(function (w) { return w ? w.charAt(0).toUpperCase() + w.slice(1) : w; }).join(' ');
}

function normDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  return String(v).trim().slice(0, 10);
}

function alert_(subject, body) {
  var to = CONFIG.ALERT_EMAIL || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail(to, subject, body);
}
