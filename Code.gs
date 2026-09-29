/**
 * Meeting poll backend — Google Apps Script bound to one Google Sheet.
 *
 * One deployment serves every poll. Two tabs:
 *   Polls      one row per poll
 *   Responses  one row per person per poll
 *
 * Setup (once):
 *   1. Project Settings > Script properties > add CREATE_KEY = <a passphrase only you know>
 *   2. Run setup() once from the editor and approve the permissions.
 *   3. Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
 *   4. Paste the /exec URL into API_URL in index.html.
 */

const SHEET_POLLS = 'Polls';
const SHEET_RESP = 'Responses';
const POLL_HEADERS = ['id', 'title', 'details', 'location', 'durationMin', 'slots',
                      'required', 'adminHash', 'closed', 'finalSlot', 'created'];
const RESP_HEADERS = ['pollId', 'rid', 'tokenHash', 'name', 'answers', 'comment', 'updated'];
const MAX_SLOTS = 60;
const MAX_RESPONSES = 150;

function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SHEET_POLLS, POLL_HEADERS);
  ensureSheet_(ss, SHEET_RESP, RESP_HEADERS);
}

function doGet(e) {
  return route_((e && e.parameter) || {});
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: 'The request body was not valid JSON.' });
  }
  return route_(body);
}

function route_(p) {
  try {
    switch (p.action) {
      case 'get':     return out_(getPoll_(p.poll, p.admin));
      case 'create':  return out_(createPoll_(p));
      case 'respond': return out_(respond_(p));
      case 'admin':   return out_(adminAction_(p));
      default:        return out_({ ok: false, error: 'Unknown action.' });
    }
  } catch (err) {
    return out_({ ok: false, code: err.code || 'error', error: String(err.message || err) });
  }
}

/* ── Actions ─────────────────────────────────────────────────────────── */

function createPoll_(p) {
  const key = PropertiesService.getScriptProperties().getProperty('CREATE_KEY');
  if (!key) fail_('The organiser passphrase is not set. Add CREATE_KEY under Script properties.');
  if (String(p.createKey || '') !== key) fail_('That organiser passphrase is wrong.', 'bad_key');

  const title = clean_(p.title, 120);
  if (!title) fail_('Give the meeting a title.');
  const details = clean_(p.details, 1000);
  const location = clean_(p.location, 200);
  const durationMin = Math.max(5, Math.min(600, parseInt(p.durationMin, 10) || 60));

  if (!Array.isArray(p.slots) || p.slots.length === 0) fail_('Add at least one time option.');
  if (p.slots.length > MAX_SLOTS) fail_('Keep it to ' + MAX_SLOTS + ' time options or fewer.');
  const seen = {};
  const slots = p.slots.map(function (s, i) {
    const start = String(s.start || '');
    if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(start)) fail_('Time option ' + (i + 1) + ' is malformed.');
    if (seen[start]) fail_('Two time options share the same start: ' + start + '.');
    seen[start] = true;
    return { id: 's' + (i + 1), start: start };
  }).sort(function (a, b) { return a.start < b.start ? -1 : 1; });

  const required = (Array.isArray(p.required) ? p.required : [])
    .map(function (n) { return clean_(n, 60); })
    .filter(Boolean)
    .slice(0, 40);

  const id = shortId_(8);
  const adminKey = Utilities.getUuid();
  const sh = sheet_(SHEET_POLLS, POLL_HEADERS);
  sh.appendRow(safeRow_([id, title, details, location, durationMin, JSON.stringify(slots),
    JSON.stringify(required), hash_(adminKey), false, '', new Date()]));
  return { ok: true, id: id, adminKey: adminKey };
}

function getPoll_(pollId, adminKey) {
  const poll = findPoll_(pollId);
  const responses = readRows_(sheet_(SHEET_RESP, RESP_HEADERS))
    .filter(function (r) { return r.pollId === poll.id; })
    .map(function (r) {
      return { rid: r.rid, name: r.name, answers: parseJson_(r.answers, {}),
               comment: r.comment || '', updated: r.updated };
    });
  return {
    ok: true,
    poll: {
      id: poll.id, title: poll.title, details: poll.details, location: poll.location,
      durationMin: Number(poll.durationMin) || 60,
      slots: parseJson_(poll.slots, []), required: parseJson_(poll.required, []),
      closed: poll.closed === true || poll.closed === 'TRUE',
      finalSlot: poll.finalSlot || ''
    },
    responses: responses,
    isAdmin: !!adminKey && hash_(adminKey) === poll.adminHash
  };
}

function respond_(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const poll = findPoll_(p.poll);
    if (poll.closed === true || poll.closed === 'TRUE') fail_('This poll is closed.', 'closed');

    const name = clean_(p.name, 60);
    if (!name) fail_('Enter your name.');
    const slotIds = {};
    parseJson_(poll.slots, []).forEach(function (s) { slotIds[s.id] = true; });
    const answers = {};
    Object.keys(p.answers || {}).forEach(function (k) {
      const v = p.answers[k];
      if (slotIds[k] && (v === 'y' || v === 'm' || v === 'n')) answers[k] = v;
    });
    const comment = clean_(p.comment, 300);

    const sh = sheet_(SHEET_RESP, RESP_HEADERS);
    const rows = readRows_(sh).filter(function (r) { return r.pollId === poll.id; });

    // Update an existing answer when the browser holds its token.
    if (p.rid && p.token) {
      const mine = rows.filter(function (r) { return r.rid === p.rid; })[0];
      if (mine && mine.tokenHash === hash_(p.token)) {
        const clash = rows.some(function (r) { return r.rid !== p.rid && norm_(r.name) === norm_(name); });
        if (clash) fail_('Someone called ' + name + ' has already answered. Add your surname to tell you apart.', 'name_taken');
        sh.getRange(mine._row, 4, 1, 4).setValues([safeRow_([name, JSON.stringify(answers), comment, new Date()])]);
        return { ok: true, rid: p.rid, token: p.token };
      }
    }

    if (rows.length >= MAX_RESPONSES) fail_('This poll has reached its response limit.');
    if (rows.some(function (r) { return norm_(r.name) === norm_(name); })) {
      fail_('Someone called ' + name + ' has already answered. If that was you on another device, ask the organiser to remove the old entry, or add your surname.', 'name_taken');
    }
    const rid = shortId_(10);
    const token = Utilities.getUuid();
    sh.appendRow(safeRow_([poll.id, rid, hash_(token), name, JSON.stringify(answers), comment, new Date()]));
    return { ok: true, rid: rid, token: token };
  } finally {
    lock.releaseLock();
  }
}

function adminAction_(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const poll = findPoll_(p.poll);
    if (!p.admin || hash_(p.admin) !== poll.adminHash) fail_('That admin link is not valid.', 'not_admin');
    const psh = sheet_(SHEET_POLLS, POLL_HEADERS);
    const col = function (h) { return POLL_HEADERS.indexOf(h) + 1; };

    switch (p.op) {
      case 'close':
        psh.getRange(poll._row, col('closed')).setValue(!!p.value);
        break;
      case 'final': {
        const ids = parseJson_(poll.slots, []).map(function (s) { return s.id; });
        const v = String(p.value || '');
        if (v && ids.indexOf(v) === -1) fail_('That time option does not exist.');
        psh.getRange(poll._row, col('finalSlot')).setValue(v);
        break;
      }
      case 'required': {
        const req = (Array.isArray(p.value) ? p.value : [])
          .map(function (n) { return clean_(n, 60); }).filter(Boolean).slice(0, 40);
        psh.getRange(poll._row, col('required')).setValue(JSON.stringify(req));
        break;
      }
      case 'delete': {
        const rsh = sheet_(SHEET_RESP, RESP_HEADERS);
        const row = readRows_(rsh).filter(function (r) {
          return r.pollId === poll.id && r.rid === p.value;
        })[0];
        if (row) rsh.deleteRow(row._row);
        break;
      }
      default:
        fail_('Unknown admin action.');
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ── Helpers ─────────────────────────────────────────────────────────── */

function findPoll_(id) {
  const poll = readRows_(sheet_(SHEET_POLLS, POLL_HEADERS))
    .filter(function (r) { return r.id === String(id || ''); })[0];
  if (!poll) fail_('No poll matches this link. Check that you copied all of it.', 'not_found');
  return poll;
}

function sheet_(name, headers) {
  return ensureSheet_(SpreadsheetApp.getActive(), name, headers);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readRows_(sh) {
  const values = sh.getDataRange().getValues();
  const head = values.shift() || [];
  return values.map(function (v, i) {
    const o = { _row: i + 2 };
    head.forEach(function (h, j) { o[h] = v[j]; });
    return o;
  });
}

// Stop a name such as "=HYPERLINK(...)" from becoming a live formula.
function safeRow_(row) {
  return row.map(function (v) {
    return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
  });
}

function clean_(s, max) {
  return String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
}

function norm_(s) {
  return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function parseJson_(s, fallback) {
  try { return JSON.parse(s); } catch (e) { return fallback; }
}

function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s))
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); })
    .join('');
}

function shortId_(n) {
  return Utilities.getUuid().replace(/-/g, '').slice(0, n);
}

function fail_(msg, code) {
  const e = new Error(msg);
  e.code = code || 'error';
  throw e;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
