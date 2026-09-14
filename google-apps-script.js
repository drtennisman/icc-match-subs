/**
 * Match Subs — backend
 *
 * Deploy as a Web App (Execute as: Me, Access: Anyone) attached to the
 * "Match Subs Worksheet for App" Google Sheet.
 *
 * Bump VERSION here AND APP_VERSION in index.html whenever this changes,
 * then redeploy with Deploy → Manage deployments → edit → New version.
 */

var VERSION = 12;

/* Sheet tabs are created automatically on first run. */
var TABS = {
  Config:   ['Key', 'Value'],
  Teams:    ['TeamID', 'Team Name', 'Level', 'Captain', 'Captain Email', 'Active'],
  Subs:     ['SubID', 'Name', 'Email', 'Phone', 'Level', 'Verified',
             'Token', 'Active', 'Added By', 'Signed Up At', 'Sub Count', 'Last Sub',
             'Season', 'Team Subs', 'Pending Name', 'Pending Phone', 'Pending Level'],
  Requests: ['ID', 'TeamID', 'Level', 'Date', 'Time', 'Match Type', 'Location', 'Opponent',
             'Notes', 'Posted By', 'Posted At', 'Notified', 'Status', 'Claimed By',
             'Claimed At', 'Nudged'],
  History:  ['Request ID', 'Team', 'Match Date', 'Sub Name', 'Sub Email', 'Claimed At', 'Posted By']
};

var DEFAULT_CONFIG = [
  ['AppUrl', ''],
  ['CaptainPIN', '1234'],
  ['ManagerEmail', ''],
  ['NudgeHours', '24'],
  ['Season', '2026 Fall'],
  ['MaxSubsPerTeam', '3']
];

/* ═══════════════════════════════════════════════════════
   SHEET PLUMBING
   ═══════════════════════════════════════════════════════ */

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getSheet(name) {
  var book = ss();
  var sheet = book.getSheetByName(name);
  if (!sheet) {
    sheet = book.insertSheet(name);
    var headers = TABS[name];
    if (headers) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
           .setFontWeight('bold').setBackground('#021f3d').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    if (name === 'Config') {
      sheet.getRange(2, 1, DEFAULT_CONFIG.length, 2).setValues(DEFAULT_CONFIG);
    }
  } else if (TABS[name]) {
    ensureColumns(sheet, TABS[name]);
  }
  return sheet;
}

/**
 * Adds any header this version expects but the existing sheet doesn't have.
 *
 * Without this, a new column means the value is silently dropped on write
 * and someone has to be told to delete the tab. Existing columns are left
 * alone — including ones no longer used — so nothing is ever destroyed.
 */
function ensureColumns(sheet, wanted) {
  var lastCol = sheet.getLastColumn();
  var existing = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    : [];

  var missing = [];
  for (var i = 0; i < wanted.length; i++) {
    if (existing.indexOf(wanted[i]) === -1) missing.push(wanted[i]);
  }
  if (!missing.length) return;

  sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing])
       .setFontWeight('bold').setBackground('#021f3d').setFontColor('#ffffff');
}

function sheetToObjects(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row.join('') === '') continue;
    var obj = { _row: i + 1 };
    for (var c = 0; c < headers.length; c++) {
      if (headers[c] !== '') obj[headers[c]] = row[c];
    }
    out.push(obj);
  }
  return out;
}

/**
 * Appends a row by matching object keys to the sheet's actual header names.
 *
 * Positional appendRow breaks silently whenever a column is added or removed:
 * every value after the gap shifts one cell left and the data looks plausible
 * but is wrong. Matching on headers means the sheet and the code can drift
 * without corrupting anything.
 */
function appendByHeader(sheet, obj) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = [];
  for (var i = 0; i < headers.length; i++) {
    var key = headers[i];
    row.push(obj[key] === undefined ? '' : obj[key]);
  }
  sheet.appendRow(row);
}

function findRowById(sheet, id) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function colIndex(sheet, header) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === header) return i + 1;
  }
  return -1;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function htmlPage(title, message, tone) {
  var color = tone === 'error' ? '#dc3545' : '#28a745';
  var appUrl = appUrl_();
  var link = appUrl
    ? '<a href="' + appUrl + '" style="display:inline-block;margin-top:22px;padding:14px 28px;' +
      'background:#021f3d;color:#fff;text-decoration:none;border-radius:12px;font-weight:700;">Open Match Subs</a>'
    : '';
  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>' + title + '</title></head>' +
    '<body style="font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;' +
    'background:#f0f4f0;margin:0;padding:40px 20px;text-align:center;">' +
    '<div style="max-width:420px;margin:0 auto;background:#fff;border-radius:16px;padding:34px 26px;' +
    'box-shadow:0 2px 12px rgba(0,0,0,.08);">' +
    '<h1 style="color:' + color + ';font-size:21px;margin:0 0 12px;">' + title + '</h1>' +
    '<p style="color:#555;font-size:15px;line-height:1.6;margin:0;">' + message + '</p>' +
    link +
    '</div></body></html>'
  ).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ═══════════════════════════════════════════════════════
   CONFIG
   ═══════════════════════════════════════════════════════ */

function getConfig(key) {
  var rows = sheetToObjects(getSheet('Config'));
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].Key).trim() === key) return String(rows[i].Value).trim();
  }
  return '';
}

function webAppUrl() {
  return ScriptApp.getService().getUrl();
}

/*
 * Links in emails point at the app on Vercel, never at script.google.com.
 *
 * A browser signed into more than one Google account rewrites a
 * script.google.com link to /macros/u/1/s/..., which fails with Google
 * Drive's "unable to open the file" page. The app instead hands the request
 * to this script with fetch, which carries no Google sign-in at all, so the
 * rewrite never happens. It also means a link only acts when someone really
 * opens it: a mail scanner that pre-fetches URLs gets a page, not a claim.
 */
var DEFAULT_APP_URL = 'https://icc-match-subs.vercel.app';

function appUrl_() {
  return (getConfig('AppUrl') || DEFAULT_APP_URL).replace(/\/+$/, '');
}

function appLink_(params) {
  var parts = [];
  for (var k in params) {
    parts.push(encodeURIComponent(k) + '=' + encodeURIComponent(params[k]));
  }
  return appUrl_() + '/?' + parts.join('&');
}

/* ═══════════════════════════════════════════════════════
   SMALL HELPERS
   ═══════════════════════════════════════════════════════ */

function newId(prefix) {
  return prefix + Utilities.getUuid().replace(/-/g, '').substring(0, 10);
}

function makeToken() {
  return Utilities.getUuid().replace(/-/g, '');
}

function splitList(v) {
  if (!v) return [];
  return String(v).split(',').map(function (s) { return s.trim(); })
                  .filter(function (s) { return s !== ''; });
}

function isTrue(v) {
  var s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === 'y' || s === '1' || s === 'x';
}

/** At least 10 digits, so "n/a" or a stray word doesn't pass as a number. */
function hasUsablePhone(v) {
  return String(v || '').replace(/\D/g, '').length >= 10;
}

function normEmail(e) {
  return String(e || '').trim().toLowerCase();
}

/**
 * Levels always come back as "3.0", never "3".
 *
 * Google Sheets stores a typed "3.0" as the number 3, so the raw cell value
 * would never match the app's "3.0" option and those teams would silently
 * find no eligible subs. Non-numeric labels (e.g. "Open") pass through.
 */
function normLevel(v) {
  var s = String(v == null ? '' : v).trim();
  if (s === '') return '';
  var n = Number(s);
  if (isNaN(n)) return s;
  return n.toFixed(1);
}

/* ═══════════════════════════════════════════════════════
   THE 3-PER-TEAM RULE

   A sub may play at most MaxSubsPerTeam matches for any one team per
   season. Counts are stored per sub as "teamId:count, teamId:count"
   alongside the season they belong to, so when the season name in Config
   changes everyone's counts read as zero without editing a single row.

   Seasons run Spring and Fall only — set Config "Season" to e.g.
   "2026 Fall". The value is just a label; nothing parses it, so any
   consistent naming works as long as it changes between seasons.
   ═══════════════════════════════════════════════════════ */

function currentSeason() {
  return getConfig('Season') || '';
}

function maxSubsPerTeam() {
  var n = Number(getConfig('MaxSubsPerTeam'));
  return (isNaN(n) || n <= 0) ? 3 : n;
}

function parseTeamSubs(v) {
  var out = {};
  var parts = splitList(v);
  for (var i = 0; i < parts.length; i++) {
    var bits = parts[i].split(':');
    if (bits.length !== 2) continue;
    var id = bits[0].trim();
    var n = Number(bits[1]);
    if (id && !isNaN(n) && n > 0) out[id] = n;
  }
  return out;
}

function formatTeamSubs(obj) {
  var parts = [];
  for (var id in obj) {
    if (obj[id] > 0) parts.push(id + ':' + obj[id]);
  }
  return parts.join(', ');
}

/** Counts for a sub row, zeroed out if they belong to a past season. */
function seasonCountsFor(row) {
  if (String(row.Season || '').trim() !== currentSeason()) return {};
  return parseTeamSubs(row['Team Subs']);
}

/**
 * Adjusts one sub's count for one team and writes it back. Rolls the record
 * onto the current season first, so a stale season is never incremented.
 */
function bumpTeamSub(subId, teamId, delta) {
  if (!subId || !teamId) return;
  var sheet = getSheet('Subs');
  var row = findRowById(sheet, subId);
  if (row === -1) return;

  var seasonCol = colIndex(sheet, 'Season');
  var subsCol   = colIndex(sheet, 'Team Subs');
  var season    = String(sheet.getRange(row, seasonCol).getValue()).trim();

  var counts;
  if (season !== currentSeason()) {
    counts = {};
    sheet.getRange(row, seasonCol).setValue(currentSeason());
  } else {
    counts = parseTeamSubs(sheet.getRange(row, subsCol).getValue());
  }

  counts[teamId] = Math.max(0, (counts[teamId] || 0) + delta);
  sheet.getRange(row, subsCol).setValue(formatTeamSubs(counts));
}

/** Sheet cells come back as Date objects or strings — normalize to YYYY-MM-DD. */
function toDateString(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, ss().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  return String(v).trim();
}

/** Times may arrive as Date objects too — normalize to HH:mm. */
function toTimeString(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, ss().getSpreadsheetTimeZone(), 'HH:mm');
  }
  return String(v).trim();
}

/** "Saturday, September 12 at 6:00 PM - Doubles", dropping any part not set. */
function matchWhen_(req) {
  var out = prettyDate(req.date);
  if (req.time) out += ' at ' + prettyTime(req.time);
  if (req.matchType) out += ' - ' + req.matchType;
  return out;
}

function prettyDate(dateStr) {
  if (!dateStr) return '';
  var parts = String(dateStr).split('-');
  if (parts.length !== 3) return dateStr;
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  return Utilities.formatDate(d, ss().getSpreadsheetTimeZone(), 'EEEE, MMMM d');
}

function prettyTime(timeStr) {
  if (!timeStr) return '';
  var parts = String(timeStr).split(':');
  if (parts.length < 2) return timeStr;
  var h = Number(parts[0]);
  var m = parts[1];
  var ampm = h >= 12 ? 'PM' : 'AM';
  var hr = h % 12 === 0 ? 12 : h % 12;
  return hr + ':' + m + ' ' + ampm;
}

/* ═══════════════════════════════════════════════════════
   DATA LOADERS
   ═══════════════════════════════════════════════════════ */

/** Reads Teams, auto-filling any blank TeamID so hand-entered rows just work. */
function loadTeams() {
  var sheet = getSheet('Teams');
  var rows = sheetToObjects(sheet);
  var idCol = colIndex(sheet, 'TeamID');
  var out = [];

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!String(r['Team Name']).trim()) continue;

    var id = String(r.TeamID).trim();
    if (!id) {
      id = newId('t');
      sheet.getRange(r._row, idCol).setValue(id);
    }
    if (r.Active !== '' && r.Active !== undefined && !isTrue(r.Active)) continue;

    out.push({
      id: id,
      name: String(r['Team Name']).trim(),
      level: normLevel(r.Level),
      captain: String(r.Captain).trim(),
      captainEmail: normEmail(r['Captain Email'])
    });
  }
  return out;
}

function loadSubs() {
  var rows = sheetToObjects(getSheet('Subs'));
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!String(r.Name).trim()) continue;
    if (r.Active !== '' && r.Active !== undefined && !isTrue(r.Active)) continue;
    out.push({
      id: String(r.SubID).trim(),
      name: String(r.Name).trim(),
      email: normEmail(r.Email),
      phone: String(r.Phone || '').trim(),
      level: normLevel(r.Level),
      season: currentSeason(),
      teamSubs: seasonCountsFor(r),
      verified: isTrue(r.Verified),
      addedBy: String(r['Added By'] || '').trim(),
      subCount: Number(r['Sub Count']) || 0,
      lastSub: toDateString(r['Last Sub'])
    });
  }
  return out;
}

function loadRequests() {
  var rows = sheetToObjects(getSheet('Requests'));
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!String(r.ID).trim()) continue;
    out.push({
      id: String(r.ID).trim(),
      teamId: String(r.TeamID).trim(),
      level: normLevel(r.Level),
      date: toDateString(r.Date),
      time: toTimeString(r.Time),
      matchType: String(r['Match Type'] || '').trim(),
      location: String(r.Location || '').trim(),
      opponent: String(r.Opponent || '').trim(),
      notes: String(r.Notes || '').trim(),
      postedBy: String(r['Posted By'] || '').trim(),
      postedAt: r['Posted At'] ? new Date(r['Posted At']).toISOString() : '',
      notified: splitList(r.Notified),
      status: String(r.Status || 'open').trim(),
      claimedBy: String(r['Claimed By'] || '').trim(),
      claimedAt: r['Claimed At'] ? new Date(r['Claimed At']).toISOString() : ''
    });
  }
  return out;
}

function subById(id) {
  var subs = loadSubs();
  for (var i = 0; i < subs.length; i++) if (subs[i].id === id) return subs[i];
  return null;
}

/* ═══════════════════════════════════════════════════════
   ACCESS CONTROL

   Everything a sub needs (their matches, the teams list) is public.
   Personal contact details and the ability to post on a team's behalf
   require the shared captain PIN from the Config tab.
   ═══════════════════════════════════════════════════════ */

function checkPin(pin) {
  var expected = getConfig('CaptainPIN');
  if (!expected) return false;
  return String(pin || '').trim() === expected;
}

function requireCaptain(data) {
  if (!checkPin(data.pin)) {
    return { status: 'error', message: 'That captain PIN is not right.' };
  }
  return null;
}

function scrubSubs(subs, isCaptain) {
  if (isCaptain) return subs;
  var out = [];
  for (var i = 0; i < subs.length; i++) {
    var s = subs[i];
    out.push({
      id: s.id, name: s.name, level: s.level,
      season: s.season, teamSubs: s.teamSubs,
      verified: s.verified, subCount: s.subCount, lastSub: s.lastSub,
      email: '', phone: ''
    });
  }
  return out;
}

function scrubTeams(teams, isCaptain) {
  if (isCaptain) return teams;
  var out = [];
  for (var i = 0; i < teams.length; i++) {
    var t = teams[i];
    out.push({ id: t.id, name: t.name, level: t.level, captain: t.captain, captainEmail: '' });
  }
  return out;
}

function teamById(id) {
  var teams = loadTeams();
  for (var i = 0; i < teams.length; i++) if (teams[i].id === id) return teams[i];
  return { id: id, name: 'Unknown team', level: '', captain: '', captainEmail: '' };
}

/* ═══════════════════════════════════════════════════════
   doGet
   ═══════════════════════════════════════════════════════ */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'ping';

  try {
    switch (action) {

      case 'ping':
        return jsonResponse({
          status: 'ok',
          version: VERSION,
          /* Surfaced so a broken link in an email can be diagnosed without
             guessing at what getUrl() resolved to. */
          webAppUrl: webAppUrl()
        });

      /*
       * Sub contact details are only returned to a caller who knows the
       * captain PIN. The app URL is public, so without this gate anyone
       * with the link could scrape every member's email and phone number.
       */
      case 'getData':
        var isCaptain = checkPin(e.parameter.pin);
        return jsonResponse({
          status: 'ok',
          version: VERSION,
          captain: isCaptain,
          season: currentSeason(),
          maxSubsPerTeam: maxSubsPerTeam(),
          teams: scrubTeams(loadTeams(), isCaptain),
          subs: scrubSubs(loadSubs(), isCaptain),
          requests: loadRequests()
        });

      /* Email confirmation link → returns a friendly HTML page, not JSON. */
      case 'verify':
        return handleVerify(e.parameter.sub, e.parameter.token);

      /* One-tap claim straight from the notification email. */
      case 'claimFromEmail':
        return handleClaimFromEmail(e.parameter.r, e.parameter.sub, e.parameter.token);

      /* Unsubscribe link at the bottom of every email. */
      case 'optOut':
        return handleOptOut(e.parameter.sub, e.parameter.token);

      default:
        return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ status: 'error', message: String(err) });
  }
}

/* ═══════════════════════════════════════════════════════
   doPost
   ═══════════════════════════════════════════════════════ */

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Bad request body' });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return jsonResponse({ status: 'error', message: 'Server busy — try again in a moment.' });
  }

  /* Anything that posts on a team's behalf or touches the roster needs the PIN. */
  var captainOnly = ['addSub', 'postRequest', 'notifyMore', 'cancelRequest'];
  if (captainOnly.indexOf(data.action) !== -1) {
    var denied = requireCaptain(data);
    if (denied) { lock.releaseLock(); return jsonResponse(denied); }
  }

  try {
    switch (data.action) {
      case 'signup':        return jsonResponse(actionSignup(data));
      case 'addSub':        return jsonResponse(actionAddSub(data));
      case 'postRequest':   return jsonResponse(actionPostRequest(data));
      case 'notifyMore':    return jsonResponse(actionNotifyMore(data));
      case 'claim':         return jsonResponse(actionClaim(data));
      case 'withdraw':      return jsonResponse(actionWithdraw(data));
      case 'cancelRequest': return jsonResponse(actionCancelRequest(data));
      case 'emailClaim':    return jsonResponse(actionEmailClaim(data));
      case 'emailVerify':   return jsonResponse(actionEmailVerify(data));
      case 'emailOptOut':   return jsonResponse(actionEmailOptOut(data));
      case 'emailRelease':  return jsonResponse(actionEmailRelease(data));
      default:
        return jsonResponse({ status: 'error', message: 'Unknown action: ' + data.action });
    }
  } catch (err) {
    return jsonResponse({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ═══════════════════════════════════════════════════════
   SIGNUP & PROFILE
   ═══════════════════════════════════════════════════════ */

function actionSignup(data) {
  var name  = String(data.name || '').trim();
  var email = normEmail(data.email);
  if (!name)  return { status: 'error', message: 'Name is required.' };
  if (!email || email.indexOf('@') === -1) return { status: 'error', message: 'A valid email is required.' };
  /* Captains text their sub once the spot is taken, so a number is required. */
  if (!hasUsablePhone(data.phone)) {
    return { status: 'error', message: 'A cell phone number is required so the captain can text you.' };
  }

  var sheet = getSheet('Subs');
  var rows = sheetToObjects(sheet);

  /*
   * Already on the list? Nothing changes until whoever owns that inbox
   * confirms it. Updating in place would let anyone who knows a member's
   * email address replace their phone number or level just by signing up.
   */
  for (var i = 0; i < rows.length; i++) {
    if (normEmail(rows[i].Email) === email) {
      var existingRow = rows[i]._row;
      var existingId = String(rows[i].SubID).trim();
      var token = String(rows[i].Token).trim();
      if (!token) {
        token = makeToken();
        sheet.getRange(existingRow, colIndex(sheet, 'Token')).setValue(token);
      }
      var newPhone = String(data.phone || '').trim();
      var newLevel = normLevel(data.level);
      sheet.getRange(existingRow, colIndex(sheet, 'Pending Name')).setValue(name);
      sheet.getRange(existingRow, colIndex(sheet, 'Pending Phone')).setValue(newPhone);
      sheet.getRange(existingRow, colIndex(sheet, 'Pending Level')).setValue(newLevel);

      sendVerifyEmail(existingId, name, email, token, { level: newLevel, phone: newPhone });
      return {
        status: 'ok',
        subId: existingId,
        verified: isTrue(rows[i].Verified),
        pending: true,
        message: 'Check your email to confirm. Nothing changes until you tap the link.'
      };
    }
  }

  var id = newId('s');
  var tok = makeToken();
  appendByHeader(sheet, {
    'SubID': id, 'Name': name, 'Email': email, 'Phone': data.phone || '',
    'Level': normLevel(data.level), 'Verified': false, 'Token': tok, 'Active': true,
    'Added By': '', 'Signed Up At': new Date(), 'Sub Count': 0, 'Last Sub': '',
    'Season': currentSeason(), 'Team Subs': ''
  });

  sendVerifyEmail(id, name, email, tok);
  return { status: 'ok', subId: id, verified: false, message: 'Confirmation email sent.' };
}

/** Captain adds someone by hand. Active right away — the captain vouched for them. */
function actionAddSub(data) {
  var name  = String(data.name || '').trim();
  var email = normEmail(data.email);
  if (!name)  return { status: 'error', message: 'Name is required.' };
  if (!email || email.indexOf('@') === -1) return { status: 'error', message: 'A valid email is required.' };
  /* Captains text their sub once the spot is taken, so a number is required. */
  if (!hasUsablePhone(data.phone)) {
    return { status: 'error', message: 'A cell phone number is required so the captain can text you.' };
  }

  var sheet = getSheet('Subs');
  var rows = sheetToObjects(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (normEmail(rows[i].Email) === email) {
      return { status: 'error', message: name + ' is already on the sub list.' };
    }
  }

  var id = newId('s');
  var tok = makeToken();
  appendByHeader(sheet, {
    'SubID': id, 'Name': name, 'Email': email, 'Phone': data.phone || '',
    'Level': normLevel(data.level), 'Verified': true, 'Token': tok, 'Active': true,
    'Added By': String(data.addedBy || 'a captain'), 'Signed Up At': new Date(),
    'Sub Count': 0, 'Last Sub': '', 'Season': currentSeason(), 'Team Subs': ''
  });

  sendAddedByCaptainEmail(id, name, email, tok, String(data.addedBy || 'A captain'),
                          normLevel(data.level));
  return { status: 'ok', subId: id, message: name + ' added and notified.' };
}

/* ═══════════════════════════════════════════════════════
   REQUESTS
   ═══════════════════════════════════════════════════════ */

function actionPostRequest(data) {
  var notified = data.notified || [];
  if (!notified.length) return { status: 'error', message: 'Pick at least one sub to notify.' };

  /* The team determines the level — never trust a level sent by the client. */
  var team = teamById(String(data.teamId).trim());
  if (!team || !team.level) {
    return { status: 'error', message: 'That team is not set up with a level on the Teams tab.' };
  }

  var id = newId('r');
  appendByHeader(getSheet('Requests'), {
    'ID': id, 'TeamID': team.id, 'Level': team.level,
    'Date': data.date, 'Time': data.time || '', 'Match Type': data.matchType || '',
    'Location': data.location || '', 'Opponent': data.opponent || '',
    'Notes': data.notes || '', 'Posted By': data.postedBy || '',
    'Posted At': new Date(), 'Notified': notified.join(', '),
    'Status': 'open', 'Claimed By': '', 'Claimed At': '', 'Nudged': false
  });

  var sent = notifySubsOfRequest(id, notified);
  return { status: 'ok', id: id, emailed: sent, message: 'Posted — ' + sent + ' notified.' };
}

function actionNotifyMore(data) {
  var sheet = getSheet('Requests');
  var row = findRowById(sheet, data.id);
  if (row === -1) return { status: 'error', message: 'That match is no longer listed.' };

  var col = colIndex(sheet, 'Notified');
  var existing = splitList(sheet.getRange(row, col).getValue());
  var added = [];
  var incoming = data.notified || [];

  for (var i = 0; i < incoming.length; i++) {
    if (existing.indexOf(incoming[i]) === -1) {
      existing.push(incoming[i]);
      added.push(incoming[i]);
    }
  }
  if (!added.length) return { status: 'error', message: 'Those subs were already notified.' };

  sheet.getRange(row, col).setValue(existing.join(', '));
  /* Reset the nudge clock so the captain gets chased again if this round goes quiet. */
  sheet.getRange(row, colIndex(sheet, 'Nudged')).setValue(false);

  var sent = notifySubsOfRequest(data.id, added);
  return { status: 'ok', emailed: sent, message: sent + ' more notified.' };
}

function actionClaim(data) {
  var sheet = getSheet('Requests');
  var row = findRowById(sheet, data.id);
  if (row === -1) return { status: 'error', message: 'That match is no longer listed.' };

  var statusCol = colIndex(sheet, 'Status');
  if (String(sheet.getRange(row, statusCol).getValue()).trim() !== 'open') {
    return { status: 'error', message: 'Someone else just claimed this one.' };
  }

  var sub = subById(data.subId);
  if (!sub) return { status: 'error', message: 'We could not find you on the sub list.' };

  var teamId = String(sheet.getRange(row, colIndex(sheet, 'TeamID')).getValue()).trim();
  var team = teamById(teamId);
  var cap = maxSubsPerTeam();
  var used = (sub.teamSubs || {})[teamId] || 0;

  /*
   * Enforced here, not just in the picker. The one-tap link in a notification
   * email bypasses the UI entirely, so this is the only place that can
   * actually stop a sub exceeding the per-team limit.
   */
  if (used >= cap) {
    return {
      status: 'error',
      message: 'You have already subbed ' + cap + ' times for ' + team.name +
               ' this season, which is the limit.'
    };
  }

  sheet.getRange(row, statusCol).setValue('filled');
  sheet.getRange(row, colIndex(sheet, 'Claimed By')).setValue(sub.id);
  sheet.getRange(row, colIndex(sheet, 'Claimed At')).setValue(new Date());

  bumpSubCount(sub.id, 1);
  bumpTeamSub(sub.id, teamId, 1);
  logHistory(data.id, sub);
  notifyCaptainOfClaim(data.id, sub);
  sendClaimReceipt_(data.id, sub);

  var left = cap - (used + 1);
  return {
    status: 'ok',
    message: 'Confirmed — the captain will text you the details.',
    subsLeftForTeam: left
  };
}

/**
 * "I can no longer play" is confirmed by email.
 *
 * The app knows a sub only by an ID that anyone can see, so withdrawing
 * straight from it would let one person quietly pull another out of a match.
 * Instead the sub holding the spot is emailed a link, and the spot reopens
 * only when they tap it (actionEmailRelease). A spoofed request just sends
 * the real sub an email they can ignore.
 */
function actionWithdraw(data) {
  var sheet = getSheet('Requests');
  var row = findRowById(sheet, data.id);
  if (row === -1) return { status: 'error', message: 'That match is no longer listed.' };

  var status = String(sheet.getRange(row, colIndex(sheet, 'Status')).getValue()).trim();
  var claimedBy = String(sheet.getRange(row, colIndex(sheet, 'Claimed By')).getValue()).trim();
  if (status !== 'filled' || !claimedBy) {
    return { status: 'error', message: 'Nobody is signed up for that match right now.' };
  }
  if (String(data.subId || '').trim() !== claimedBy) {
    return { status: 'error', message: 'Only the sub who took this match can back out of it.' };
  }

  var sub = rawSub_(claimedBy);
  if (!sub || !sub.email) {
    return { status: 'error', message: 'We could not find an email address to confirm this with.' };
  }

  sendReleaseEmail_(data.id, sub);
  return {
    status: 'ok',
    pending: true,
    message: 'Check your email and tap the button to confirm. You stay on the match until you do.'
  };
}

/** Reopens a filled match and gives the sub's slot back. */
function doWithdraw_(requestId, subId) {
  var sheet = getSheet('Requests');
  var row = findRowById(sheet, requestId);
  if (row === -1) return;
  var teamId = String(sheet.getRange(row, colIndex(sheet, 'TeamID')).getValue()).trim();

  sheet.getRange(row, colIndex(sheet, 'Status')).setValue('open');
  sheet.getRange(row, colIndex(sheet, 'Claimed By')).setValue('');
  sheet.getRange(row, colIndex(sheet, 'Claimed At')).setValue('');
  sheet.getRange(row, colIndex(sheet, 'Nudged')).setValue(false);

  /* Backing out gives the slot back — it should not count against them. */
  bumpSubCount(subId, -1);
  bumpTeamSub(subId, teamId, -1);
  notifyCaptainOfWithdrawal(requestId, rawSub_(subId));
}

/** A sub by ID whether or not they're still active, for someone backing out. */
function rawSub_(subId) {
  var sheet = getSheet('Subs');
  var row = findRowById(sheet, subId);
  if (row === -1) return null;
  var get = function (header) { return sheet.getRange(row, colIndex(sheet, header)).getValue(); };
  return {
    id: String(get('SubID')).trim(),
    name: String(get('Name')).trim(),
    email: normEmail(get('Email')),
    phone: String(get('Phone') || '').trim()
  };
}

function actionCancelRequest(data) {
  var sheet = getSheet('Requests');
  var row = findRowById(sheet, data.id);
  if (row === -1) return { status: 'error', message: 'That match is no longer listed.' };

  notifyCancellation(data.id);
  sheet.getRange(row, colIndex(sheet, 'Status')).setValue('cancelled');
  return { status: 'ok', message: 'Cancelled — everyone notified has been told.' };
}

function bumpSubCount(subId, delta) {
  var sheet = getSheet('Subs');
  var row = findRowById(sheet, subId);
  if (row === -1) return;
  var col = colIndex(sheet, 'Sub Count');
  var current = Number(sheet.getRange(row, col).getValue()) || 0;
  sheet.getRange(row, col).setValue(Math.max(0, current + delta));
  if (delta > 0) sheet.getRange(row, colIndex(sheet, 'Last Sub')).setValue(new Date());
}

function logHistory(requestId, sub) {
  var reqs = loadRequests();
  for (var i = 0; i < reqs.length; i++) {
    if (reqs[i].id === requestId) {
      var team = teamById(reqs[i].teamId);
      appendByHeader(getSheet('History'), {
        'Request ID': requestId, 'Team': team.name, 'Match Date': reqs[i].date,
        'Sub Name': sub.name, 'Sub Email': sub.email,
        'Claimed At': new Date(), 'Posted By': reqs[i].postedBy
      });
      return;
    }
  }
}

/* ═══════════════════════════════════════════════════════
   EMAIL LINK HANDLERS

   The action* versions are what current emails use: the link opens the app,
   and the app posts here. The handle* versions below them answer links in
   emails sent before that change, so nobody's old email goes dead.
   ═══════════════════════════════════════════════════════ */

/** The sub's row if the token matches, otherwise null. */
function subRowForToken_(subId, token) {
  var sheet = getSheet('Subs');
  var row = findRowById(sheet, subId);
  if (row === -1) return null;
  var stored = String(sheet.getRange(row, colIndex(sheet, 'Token')).getValue()).trim();
  if (!token || String(token).trim() !== stored) return null;
  return { sheet: sheet, row: row };
}

function subSummary_(hit) {
  var sh = hit.sheet, r = hit.row;
  return {
    id: String(sh.getRange(r, colIndex(sh, 'SubID')).getValue()).trim(),
    name: String(sh.getRange(r, colIndex(sh, 'Name')).getValue()).trim(),
    level: normLevel(sh.getRange(r, colIndex(sh, 'Level')).getValue())
  };
}

function badLink_() {
  return {
    status: 'error',
    title: 'Link not recognized',
    message: "That link isn't valid any more. Open the app to check your matches or sign up again."
  };
}

function actionEmailVerify(data) {
  var hit = subRowForToken_(data.subId, data.token);
  if (!hit) return badLink_();
  var sh = hit.sheet, r = hit.row;

  /* A signup for an address already on the list parks its changes until now. */
  var changed = false;
  var fields = ['Name', 'Phone', 'Level'];
  for (var f = 0; f < fields.length; f++) {
    var pendingCol = colIndex(sh, 'Pending ' + fields[f]);
    var value = String(sh.getRange(r, pendingCol).getValue()).trim();
    if (!value) continue;
    sh.getRange(r, colIndex(sh, fields[f])).setValue(fields[f] === 'Level' ? normLevel(value) : value);
    sh.getRange(r, pendingCol).setValue('');
    changed = true;
  }

  sh.getRange(r, colIndex(sh, 'Verified')).setValue(true);
  sh.getRange(r, colIndex(sh, 'Active')).setValue(true);
  var sub = subSummary_(hit);

  if (changed) {
    return {
      status: 'ok',
      title: 'Your changes are saved',
      message: 'Thanks ' + sub.name + " — you're updated, and you'll hear about matches for " +
               'teams rated ' + sub.level + ' and above.',
      sub: sub
    };
  }

  return {
    status: 'ok',
    title: "You're on the sub list",
    message: 'Thanks ' + sub.name + " — you're all set. You'll get an email whenever a " +
             'team at your level or above needs someone.',
    sub: sub
  };
}

function actionEmailClaim(data) {
  var hit = subRowForToken_(data.subId, data.token);
  if (!hit) return badLink_();
  var sub = subSummary_(hit);

  var req = null;
  var reqs = loadRequests();
  for (var i = 0; i < reqs.length; i++) if (reqs[i].id === data.requestId) req = reqs[i];
  if (!req) {
    return { status: 'error', title: 'Match not found',
             message: 'That match is no longer listed. The captain may have cancelled it.', sub: sub };
  }

  var team = teamById(req.teamId);
  var described = team.name + ' on ' + matchWhen_(req) + ', ' + req.location;

  /* Tapping the button twice shouldn't read as "someone beat you to it". */
  if (req.status === 'filled' && req.claimedBy === sub.id) {
    return { status: 'ok', title: "You're already in",
             message: "You're subbing for " + described + '. ' + req.postedBy +
                      ' will text you the details.', sub: sub };
  }

  var result = actionClaim({ id: req.id, subId: sub.id });
  if (result.status !== 'ok') {
    return { status: 'error', title: "Couldn't take this match", message: result.message, sub: sub };
  }

  return {
    status: 'ok',
    title: "You're in",
    message: "You're subbing for " + described + '. ' + req.postedBy +
             ' has been notified and will text you the details.',
    sub: sub
  };
}

function actionEmailRelease(data) {
  var hit = subRowForToken_(data.subId, data.token);
  if (!hit) return badLink_();
  var sub = subSummary_(hit);

  var req = findRequest_(data.requestId);
  if (!req || req.status === 'cancelled') {
    return { status: 'error', title: 'Match not found',
             message: 'That match is no longer listed. The captain may have cancelled it.' };
  }

  /* Tapped twice, or the spot has already moved on. */
  if (req.status !== 'filled' || req.claimedBy !== sub.id) {
    return { status: 'ok', title: "You're not on this match",
             message: "Nothing to undo — you aren't signed up for this match any more.", sub: sub };
  }

  doWithdraw_(req.id, sub.id);
  return {
    status: 'ok',
    title: "You're off this match",
    message: 'Thanks for letting us know. ' + (req.postedBy || 'The captain') +
             ' has been told, and the spot is open again.',
    sub: sub
  };
}

function actionEmailOptOut(data) {
  var hit = subRowForToken_(data.subId, data.token);
  if (!hit) return badLink_();

  hit.sheet.getRange(hit.row, colIndex(hit.sheet, 'Active')).setValue(false);
  var sub = subSummary_(hit);
  return {
    status: 'ok',
    title: "You're off the list",
    message: 'No problem, ' + sub.name + " — you won't get any more sub requests. " +
             'You can sign up again in the app anytime.'
  };
}

function handleVerify(subId, token) {
  var sheet = getSheet('Subs');
  var row = findRowById(sheet, subId);
  if (row === -1) return htmlPage('Link not recognized', 'We could not find that signup. Try signing up again in the app.', 'error');

  var stored = String(sheet.getRange(row, colIndex(sheet, 'Token')).getValue()).trim();
  if (!token || token !== stored) {
    return htmlPage('Link not recognized', 'That confirmation link is not valid. Try signing up again in the app.', 'error');
  }

  sheet.getRange(row, colIndex(sheet, 'Verified')).setValue(true);
  sheet.getRange(row, colIndex(sheet, 'Active')).setValue(true);
  var name = String(sheet.getRange(row, colIndex(sheet, 'Name')).getValue()).trim();

  return htmlPage("You're on the sub list",
    'Thanks ' + name + " — you're all set. Captains can now call on you, and you'll get an email whenever a match fits what you play.");
}

function handleClaimFromEmail(requestId, subId, token) {
  var subSheet = getSheet('Subs');
  var subRow = findRowById(subSheet, subId);
  if (subRow === -1) return htmlPage('Link not recognized', 'We could not find you on the sub list.', 'error');

  var stored = String(subSheet.getRange(subRow, colIndex(subSheet, 'Token')).getValue()).trim();
  if (!token || token !== stored) {
    return htmlPage('Link not recognized', 'That link is not valid. Open the app to claim this match instead.', 'error');
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) {
    return htmlPage('Try again', 'The server was busy. Refresh this page to try once more.', 'error');
  }

  try {
    var result = actionClaim({ id: requestId, subId: subId });
    if (result.status !== 'ok') return htmlPage('Already taken', result.message, 'error');

    var reqs = loadRequests();
    for (var i = 0; i < reqs.length; i++) {
      if (reqs[i].id === requestId) {
        var team = teamById(reqs[i].teamId);
        return htmlPage("You're in",
          "You're subbing for " + team.name + ' on ' + matchWhen_(reqs[i]) +
          ', ' + reqs[i].location +
          '. ' + reqs[i].postedBy + ' has been notified and will text you the details.');
      }
    }
    return htmlPage("You're in", 'The captain has been notified and will text you the details.');
  } finally {
    lock.releaseLock();
  }
}

function handleOptOut(subId, token) {
  var sheet = getSheet('Subs');
  var row = findRowById(sheet, subId);
  if (row === -1) return htmlPage('Link not recognized', 'We could not find that account.', 'error');

  var stored = String(sheet.getRange(row, colIndex(sheet, 'Token')).getValue()).trim();
  if (!token || token !== stored) return htmlPage('Link not recognized', 'That link is not valid.', 'error');

  sheet.getRange(row, colIndex(sheet, 'Active')).setValue(false);
  var name = String(sheet.getRange(row, colIndex(sheet, 'Name')).getValue()).trim();
  return htmlPage("You're off the list",
    'No problem, ' + name + " — you won't get any more sub requests. Sign up again in the app anytime.");
}

/* ═══════════════════════════════════════════════════════
   EMAILS
   ═══════════════════════════════════════════════════════ */

/**
 * Every email goes out as HTML with real <a> buttons.
 *
 * These links run past 190 characters. Pasted bare into a plain-text email
 * they get wrapped by the receiving client, which then hyperlinks only the
 * first fragment — the recipient taps it and lands on a broken page. An
 * anchor tag cannot be split that way. A plain-text alternative still goes
 * along for clients that refuse HTML.
 */
function esc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function emailHtml(opts) {
  var rows = '';
  var details = opts.details || [];
  for (var i = 0; i < details.length; i++) {
    rows += '<tr>' +
      '<td style="padding:4px 14px 4px 0;color:#777;font-size:14px;white-space:nowrap;">' +
        esc_(details[i][0]) + '</td>' +
      '<td style="padding:4px 0;color:#1a1a1a;font-size:14px;font-weight:600;">' +
        esc_(details[i][1]) + '</td>' +
    '</tr>';
  }

  /* Declared explicitly so the em-dashes and curly quotes in this content
     can't arrive as mojibake in a client that guesses the encoding. */
  return '<!doctype html><html><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1"></head><body style="margin:0;">' +
    '<div style="font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,sans-serif;' +
      'background:#f0f4f0;padding:24px 12px;">' +
    '<div style="max-width:480px;margin:0 auto;background:#fff;border-radius:14px;' +
        'padding:26px 24px;box-shadow:0 1px 6px rgba(0,0,0,.08);">' +
      '<h1 style="margin:0 0 6px;font-size:19px;color:#052d54;">' + esc_(opts.heading) + '</h1>' +
      (opts.lead ? '<p style="margin:0 0 16px;font-size:15px;color:#555;line-height:1.55;">' +
          esc_(opts.lead) + '</p>' : '') +
      (rows ? '<table cellpadding="0" cellspacing="0" style="margin:0 0 18px;">' + rows + '</table>' : '') +
      (opts.note ? '<div style="background:#f8f9fa;border-radius:8px;padding:11px 13px;margin:0 0 18px;' +
          'font-size:14px;color:#555;line-height:1.5;">' + esc_(opts.note) + '</div>' : '') +
      (opts.buttonUrl ? '<a href="' + esc_(opts.buttonUrl) + '" ' +
          'style="display:inline-block;background:#052d54;color:#fff;text-decoration:none;' +
          'padding:14px 26px;border-radius:11px;font-weight:700;font-size:16px;">' +
          esc_(opts.buttonLabel) + '</a>' : '') +
      (opts.after ? '<p style="margin:18px 0 0;font-size:14px;color:#555;line-height:1.55;">' +
          opts.after + '</p>' : '') +
      (opts.footer ? '<p style="margin:22px 0 0;padding-top:14px;border-top:1px solid #eee;' +
          'font-size:12px;color:#999;line-height:1.5;">' + opts.footer + '</p>' : '') +
    '</div>' +
  '</div></body></html>';
}

function optOutUrl(subId, token) {
  return appLink_({ optout: subId, t: token || subToken(subId) });
}

/** Plain-text footer. The opt-out link is deliberately kept well clear of any
 *  action link so a mis-tap doesn't quietly remove someone from the list. */
function emailFooter(sub) {
  return '\n\n---\n' +
    'To stop receiving these, open this link:\n' + optOutUrl(sub.id);
}

function footerHtml(sub) {
  return 'Not interested in subbing any more? ' +
    '<a href="' + esc_(optOutUrl(sub.id)) + '" style="color:#999;">Take me off the list</a>.';
}

function subToken(subId) {
  var sheet = getSheet('Subs');
  var row = findRowById(sheet, subId);
  if (row === -1) return '';
  return String(sheet.getRange(row, colIndex(sheet, 'Token')).getValue()).trim();
}

/*
 * Who an email appears to come from, and where a reply goes.
 *
 * Apps Script can only send from the account that deployed it, and putting a
 * captain's Gmail address in From would be treated as spoofing and land in
 * spam. So the captain goes in the display name and in Reply-To instead: the
 * inbox shows "Jane Miller via ICC Match Subs", and Reply writes to Jane.
 */
var APP_SENDER_NAME = 'ICC Match Subs';

function senderName_(person) {
  var clean = String(person || '').replace(/["<>\r\n]/g, '').trim();
  return clean ? clean + ' via ' + APP_SENDER_NAME : APP_SENDER_NAME;
}

function isEmail_(v) {
  return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(v || '').trim());
}

/** MailApp.sendEmail with a display name, plus Reply-To when there's a valid address. */
function sendMail_(msg, fromPerson, replyTo) {
  msg.name = senderName_(fromPerson);
  if (isEmail_(replyTo)) msg.replyTo = String(replyTo).trim();
  if (!msg.cc) delete msg.cc;
  MailApp.sendEmail(msg);
}

/** A captain's address from the Teams tab, by the name they picked in the app. */
function captainEmailByName_(name) {
  var teams = loadTeams();
  for (var i = 0; i < teams.length; i++) {
    if (teams[i].captain === name && teams[i].captainEmail) return teams[i].captainEmail;
  }
  return '';
}

function sendVerifyEmail(subId, name, email, token, change) {
  var url = appLink_({ verify: subId, t: token });
  var subject, body, html;

  if (change) {
    /* Someone signed up again with an address that's already on the list. */
    subject = 'Confirm your sub list changes';
    body =
      'Hi ' + name + ',\n\n' +
      'We got a request to update your details on the tennis sub list:\n\n' +
      'Level: ' + change.level + '\n' +
      'Cell:  ' + change.phone + '\n\n' +
      'To save these changes, open this link:\n\n' + url + '\n\n' +
      "If you didn't ask for this, ignore this email and nothing changes.";
    html = emailHtml({
      heading: 'Confirm your changes',
      lead: 'Hi ' + name + ' — tap below to save your updated sub list details.',
      details: [['Level', change.level], ['Cell', change.phone]],
      buttonUrl: url,
      buttonLabel: 'Save my changes',
      after: 'Nothing changes until you do.',
      footer: "Didn't ask for this? Just ignore this email and nothing changes."
    });
  } else {
    subject = 'Confirm your spot on the sub list';
    body =
      'Hi ' + name + ',\n\n' +
      'Confirm your spot on the tennis sub list by opening this link:\n\n' + url + '\n\n' +
      "Until you do, you won't receive any match requests.\n\n" +
      "If you didn't sign up, just ignore this email.";
    html = emailHtml({
      heading: 'Confirm your spot',
      lead: 'Hi ' + name + ' — one tap and you\'re on the sub list.',
      buttonUrl: url,
      buttonLabel: 'Confirm my spot',
      after: "Until you do, you won't get any match requests.",
      footer: "Didn't sign up? Just ignore this email and nothing happens."
    });
  }

  try {
    sendMail_({ to: email, subject: subject, body: body, htmlBody: html }, '', '');
  } catch (err) {
    Logger.log('Verify email failed for ' + email + ': ' + err);
  }
}

function sendAddedByCaptainEmail(subId, name, email, token, addedBy, level) {
  var appUrl = appUrl_();

  var body =
    'Hi ' + name + ',\n\n' +
    addedBy + ' added you to the tennis sub list.\n\n' +
    (level ? 'Your level: ' + level + '\n' : '') +
    "\nThat means you'll get an email when a team at your level or above " +
    'needs a sub. You are never obligated to say yes.\n\n' +
    (appUrl ? 'Open the app to change your level:\n' + appUrl + '\n\n' : '') +
    "If you'd rather not be on the list at all:\n" + optOutUrl(subId, token);

  var html = emailHtml({
    heading: addedBy + ' added you to the sub list',
    lead: 'Hi ' + name + " — you'll get an email when a team at your level or above needs a sub. " +
          'You are never obligated to say yes.',
    details: level ? [['Your level', level]] : [],
    buttonUrl: appUrl || '',
    buttonLabel: 'Open the app',
    after: appUrl ? 'You can change your level there any time.' : '',
    footer: "Rather not be on the list? " +
      '<a href="' + esc_(optOutUrl(subId, token)) + '" style="color:#999;">Take me off</a>.'
  });

  try {
    var adder = addedBy === 'A captain' ? '' : addedBy;
    sendMail_({
      to: email, subject: "You've been added to the tennis sub list",
      body: body, htmlBody: html
    }, adder, captainEmailByName_(adder));
  } catch (err) {
    Logger.log('Added-by-captain email failed for ' + email + ': ' + err);
  }
}

function notifySubsOfRequest(requestId, subIds) {
  var reqs = loadRequests();
  var req = null;
  for (var i = 0; i < reqs.length; i++) if (reqs[i].id === requestId) req = reqs[i];
  if (!req) return 0;

  var team = teamById(req.teamId);
  var sent = 0;

  /* Whoever posted it, falling back to the team's listed captain. */
  var posterName  = req.postedBy || team.captain;
  var posterEmail = captainEmailByName_(posterName) || team.captainEmail;

  var cap = maxSubsPerTeam();

  for (var j = 0; j < subIds.length; j++) {
    var sub = subById(subIds[j]);
    if (!sub || !sub.email || !sub.verified) continue;
    /* Don't invite someone who would be turned away when they tap the link. */
    if (((sub.teamSubs || {})[req.teamId] || 0) >= cap) continue;

    var claimUrl = appLink_({ claim: req.id, sub: sub.id, t: subToken(sub.id) });

    var details = [
      ['When', prettyDate(req.date) + (req.time ? ' at ' + prettyTime(req.time) : '')],
      ['Playing', req.matchType],
      ['Where', req.location]
    ];
    if (req.opponent) details.push(['Vs', req.opponent]);
    details.push(['Level', req.level]);

    var body =
      'Hi ' + sub.name + ',\n\n' +
      team.name + ' needs a sub.\n\n' +
      'When:    ' + prettyDate(req.date) + (req.time ? ' at ' + prettyTime(req.time) : '') + '\n' +
      'Playing: ' + req.matchType + '\n' +
      'Where: ' + req.location + '\n' +
      (req.opponent ? 'Vs:    ' + req.opponent + '\n' : '') +
      'Level: ' + req.level + '\n' +
      (req.notes ? '\nNotes from ' + req.postedBy + ':\n' + req.notes + '\n' : '') +
      '\nFirst to accept gets the spot. Open this link to take it:\n' + claimUrl + '\n' +
      emailFooter(sub);

    var html = emailHtml({
      heading: team.name + ' needs a sub',
      lead: 'Hi ' + sub.name + ' — first to accept gets the spot.',
      details: details,
      note: req.notes ? 'From ' + req.postedBy + ': ' + req.notes : '',
      buttonUrl: claimUrl,
      buttonLabel: "Yes, I'll sub",
      after: 'Nothing is confirmed until you tap that button.',
      footer: footerHtml(sub)
    });

    try {
      sendMail_({
        to: sub.email,
        subject: 'Sub needed — ' + team.name + ', ' + prettyDate(req.date),
        body: body,
        htmlBody: html
      }, posterName, posterEmail);
      sent++;
    } catch (err) {
      Logger.log('Request email failed for ' + sub.email + ': ' + err);
    }
  }
  return sent;
}

/** The request with this ID, or null. */
function findRequest_(requestId) {
  var reqs = loadRequests();
  for (var i = 0; i < reqs.length; i++) if (reqs[i].id === requestId) return reqs[i];
  return null;
}

/** When / Playing / Where / Vs rows for an email. */
function matchDetails_(req) {
  var details = [
    ['When', prettyDate(req.date) + (req.time ? ' at ' + prettyTime(req.time) : '')],
    ['Playing', req.matchType],
    ['Where', req.location]
  ];
  if (req.opponent) details.push(['Vs', req.opponent]);
  return details;
}

function matchDetailsText_(req) {
  var rows = matchDetails_(req), out = '';
  for (var i = 0; i < rows.length; i++) out += rows[i][0] + ': ' + rows[i][1] + '\n';
  return out;
}

/**
 * Sent to whoever just took a match: what they signed up for, who will text
 * them, and a way out. Because it lands in the sub's own inbox, a claim made
 * in someone else's name can't go unnoticed.
 */
function sendClaimReceipt_(requestId, sub) {
  var req = findRequest_(requestId);
  if (!req || !sub || !sub.email) return;

  var team = teamById(req.teamId);
  var posterName  = req.postedBy || team.captain;
  var posterEmail = captainEmailByName_(posterName) || team.captainEmail;
  var who = posterName || 'The captain';
  var releaseUrl = appLink_({ release: req.id, sub: sub.id, t: subToken(sub.id) });

  var body =
    'Hi ' + sub.name + ',\n\n' +
    "You're subbing for " + team.name + '. ' + who + ' will text you the details.\n\n' +
    matchDetailsText_(req) +
    (req.notes ? '\nNotes from ' + who + ':\n' + req.notes + '\n' : '') +
    "\nCan't make it after all? Release your spot so the captain can find someone else:\n" +
    releaseUrl;

  var html = emailHtml({
    heading: "You're subbing for " + team.name,
    lead: 'Hi ' + sub.name + ' — ' + who + ' will text you the details.',
    details: matchDetails_(req),
    note: req.notes ? 'From ' + who + ': ' + req.notes : '',
    footer: "Can't make it after all? " +
      '<a href="' + esc_(releaseUrl) + '" style="color:#999;">Release your spot</a>' +
      ' so the captain can find someone else.'
  });

  try {
    sendMail_({
      to: sub.email,
      subject: "You're subbing — " + team.name + ', ' + prettyDate(req.date),
      body: body, htmlBody: html
    }, posterName, posterEmail);
  } catch (err) {
    Logger.log('Claim receipt failed for ' + sub.email + ': ' + err);
  }
}

/** Asks the sub holding a spot to confirm they're giving it up. */
function sendReleaseEmail_(requestId, sub) {
  var req = findRequest_(requestId);
  if (!req || !sub || !sub.email) return;

  var team = teamById(req.teamId);
  var posterName  = req.postedBy || team.captain;
  var posterEmail = captainEmailByName_(posterName) || team.captainEmail;
  /* go=1: this button is already the "are you sure", so the app won't ask again. */
  var url = appLink_({ release: req.id, sub: sub.id, t: subToken(sub.id), go: 1 });

  var body =
    'Hi ' + sub.name + ',\n\n' +
    'To give up your spot for ' + team.name + ', open this link:\n\n' + url + '\n\n' +
    matchDetailsText_(req) +
    '\nStill planning to play? Ignore this email and you stay on the match.';

  var html = emailHtml({
    heading: "Confirm you can't make it",
    lead: 'Hi ' + sub.name + ' — tap below to give up your spot for ' + team.name + '.',
    details: matchDetails_(req),
    buttonUrl: url,
    buttonLabel: 'Yes, release my spot',
    after: 'Still planning to play? Ignore this email and you stay on the match.'
  });

  try {
    sendMail_({
      to: sub.email,
      subject: "Confirm you're backing out — " + team.name + ', ' + prettyDate(req.date),
      body: body, htmlBody: html
    }, posterName, posterEmail);
  } catch (err) {
    Logger.log('Release email failed for ' + sub.email + ': ' + err);
  }
}

function notifyCaptainOfClaim(requestId, sub) {
  var reqs = loadRequests();
  var req = null;
  for (var i = 0; i < reqs.length; i++) if (reqs[i].id === requestId) req = reqs[i];
  if (!req) return;

  var team = teamById(req.teamId);
  var to = team.captainEmail;
  if (!to) return;

  var body =
    'Good news — ' + sub.name + ' is subbing.\n\n' +
    'Match: ' + team.name + '\n' +
    'When:    ' + prettyDate(req.date) + (req.time ? ' at ' + prettyTime(req.time) : '') + '\n' +
    'Playing: ' + req.matchType + '\n' +
    'Where: ' + req.location + '\n' +
    '\nReach ' + sub.name + ':\n' +
    '  ' + sub.email + '\n' +
    (sub.phone ? '  ' + sub.phone + '\n' : '') +
    '\nOpen the app: ' + appUrl_();

  var cc = getConfig('ManagerEmail');
  try {
    sendMail_({
      to: to, cc: cc || '',
      subject: sub.name + ' is subbing — ' + team.name + ', ' + prettyDate(req.date),
      body: body
    }, '', sub.email);
  } catch (err) {
    Logger.log('Captain claim email failed: ' + err);
  }
}

function notifyCaptainOfWithdrawal(requestId, sub) {
  var reqs = loadRequests();
  var req = null;
  for (var i = 0; i < reqs.length; i++) if (reqs[i].id === requestId) req = reqs[i];
  if (!req) return;

  var team = teamById(req.teamId);
  if (!team.captainEmail) return;

  var body =
    (sub ? sub.name : 'Your sub') + ' can no longer play ' + team.name +
    ' on ' + matchWhen_(req) + '.\n\n' +
    'The spot is open again. Open the app to notify more subs:\n' +
    appUrl_();

  try {
    sendMail_({
      to: team.captainEmail,
      subject: 'Sub dropped out — ' + team.name + ', ' + prettyDate(req.date),
      body: body
    }, '', sub ? sub.email : '');
  } catch (err) {
    Logger.log('Withdrawal email failed: ' + err);
  }
}

function notifyCancellation(requestId) {
  var reqs = loadRequests();
  var req = null;
  for (var i = 0; i < reqs.length; i++) if (reqs[i].id === requestId) req = reqs[i];
  if (!req) return;

  var team = teamById(req.teamId);
  var posterName  = req.postedBy || team.captain;
  var posterEmail = captainEmailByName_(posterName) || team.captainEmail;

  for (var j = 0; j < req.notified.length; j++) {
    var sub = subById(req.notified[j]);
    if (!sub || !sub.email) continue;
    try {
      sendMail_({
        to: sub.email,
        subject: 'No longer needed — ' + team.name + ', ' + prettyDate(req.date),
        body: 'Hi ' + sub.name + ',\n\n' + team.name + ' no longer needs a sub for ' +
              prettyDate(req.date) + '. Thanks anyway!' + emailFooter(sub)
      }, posterName, posterEmail);
    } catch (err) {
      Logger.log('Cancellation email failed for ' + sub.email + ': ' + err);
    }
  }
}

/* ═══════════════════════════════════════════════════════
   NO-RESPONSE NUDGE  (time-driven trigger)
   ═══════════════════════════════════════════════════════ */

/**
 * Emails a captain when nobody has claimed their match after NudgeHours.
 * Each request is nudged once — the Nudged flag resets if they widen the net.
 * Run setupNudgeTrigger() once to schedule this hourly.
 */
function sendNoResponseNudges() {
  var sheet = getSheet('Requests');
  var rows = sheetToObjects(sheet);
  var hours = Number(getConfig('NudgeHours')) || 24;
  var cutoff = new Date().getTime() - hours * 3600 * 1000;
  var nudgedCol = colIndex(sheet, 'Nudged');
  var appUrl = appUrl_();

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r.Status).trim() !== 'open') continue;
    if (isTrue(r.Nudged)) continue;
    if (!r['Posted At']) continue;
    if (new Date(r['Posted At']).getTime() > cutoff) continue;

    var matchDate = toDateString(r.Date);
    if (matchDate) {
      var parts = matchDate.split('-');
      var md = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
      var today = new Date(); today.setHours(0, 0, 0, 0);
      if (md < today) continue; /* match already happened */
    }

    var team = teamById(String(r.TeamID).trim());
    if (!team.captainEmail) continue;

    var notifiedCount = splitList(r.Notified).length;
    var body =
      'Nobody has claimed this match yet:\n\n' +
      team.name + '\n' +
      prettyDate(matchDate) + (toTimeString(r.Time) ? ' at ' + prettyTime(toTimeString(r.Time)) : '') +
      ' - ' + String(r['Match Type'] || '') + '\n' +
      String(r.Location || '') + '\n\n' +
      'You notified ' + notifiedCount + ' ' + (notifiedCount === 1 ? 'sub' : 'subs') +
      ' and none have accepted.\n\n' +
      'Open the app and tap "Notify more subs" to widen the net:\n' + appUrl;

    try {
      sendMail_({
        to: team.captainEmail,
        subject: 'Still no sub — ' + team.name + ', ' + prettyDate(matchDate),
        body: body
      }, '', '');
      sheet.getRange(r._row, nudgedCol).setValue(true);
    } catch (err) {
      Logger.log('Nudge email failed: ' + err);
    }
  }
}

/** Run once from the Apps Script editor to schedule the nudge check hourly. */
function setupNudgeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendNoResponseNudges') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('sendNoResponseNudges').timeBased().everyHours(1).create();
  return 'Nudge trigger scheduled — runs hourly.';
}

/** Run once from the editor to create all tabs without waiting for first use. */
function setupSheets() {
  for (var name in TABS) getSheet(name);
  return 'All tabs ready.';
}
