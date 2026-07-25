/**
 * Match Subs — backend
 *
 * Deploy as a Web App (Execute as: Me, Access: Anyone) attached to the
 * "Match Subs Worksheet for App" Google Sheet.
 *
 * Bump VERSION here AND APP_VERSION in index.html whenever this changes,
 * then redeploy with Deploy → Manage deployments → edit → New version.
 */

var VERSION = 1;

/* Sheet tabs are created automatically on first run. */
var TABS = {
  Config:   ['Key', 'Value'],
  Teams:    ['TeamID', 'Team Name', 'Level', 'Captain', 'Captain Email', 'Active'],
  Subs:     ['SubID', 'Name', 'Email', 'Phone', 'Levels', 'Teams', 'Verified',
             'Token', 'Active', 'Added By', 'Signed Up At', 'Sub Count', 'Last Sub'],
  Requests: ['ID', 'TeamID', 'Level', 'Date', 'Time', 'Location', 'Opponent', 'Line',
             'Notes', 'Posted By', 'Posted At', 'Notified', 'Status', 'Claimed By',
             'Claimed At', 'Nudged'],
  History:  ['Request ID', 'Team', 'Match Date', 'Sub Name', 'Sub Email', 'Claimed At', 'Posted By']
};

var DEFAULT_CONFIG = [
  ['AppUrl', ''],
  ['CaptainPIN', '1234'],
  ['ManagerEmail', ''],
  ['NudgeHours', '24']
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
  }
  return sheet;
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
  var appUrl = getConfig('AppUrl');
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

function normEmail(e) {
  return String(e || '').trim().toLowerCase();
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
      level: String(r.Level).trim(),
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
      levels: splitList(r.Levels),
      teams: splitList(r.Teams),
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
      level: String(r.Level).trim(),
      date: toDateString(r.Date),
      time: toTimeString(r.Time),
      location: String(r.Location || '').trim(),
      opponent: String(r.Opponent || '').trim(),
      line: String(r.Line || '').trim(),
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
      id: s.id, name: s.name, levels: s.levels, teams: s.teams,
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
        return jsonResponse({ status: 'ok', version: VERSION });

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
      case 'updateProfile': return jsonResponse(actionUpdateProfile(data));
      case 'addSub':        return jsonResponse(actionAddSub(data));
      case 'postRequest':   return jsonResponse(actionPostRequest(data));
      case 'notifyMore':    return jsonResponse(actionNotifyMore(data));
      case 'claim':         return jsonResponse(actionClaim(data));
      case 'withdraw':      return jsonResponse(actionWithdraw(data));
      case 'cancelRequest': return jsonResponse(actionCancelRequest(data));
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

  var sheet = getSheet('Subs');
  var rows = sheetToObjects(sheet);

  /* Already on the list? Update in place and re-send the confirmation. */
  for (var i = 0; i < rows.length; i++) {
    if (normEmail(rows[i].Email) === email) {
      var token = String(rows[i].Token).trim() || makeToken();
      sheet.getRange(rows[i]._row, colIndex(sheet, 'Name')).setValue(name);
      sheet.getRange(rows[i]._row, colIndex(sheet, 'Phone')).setValue(data.phone || '');
      sheet.getRange(rows[i]._row, colIndex(sheet, 'Levels')).setValue((data.levels || []).join(', '));
      sheet.getRange(rows[i]._row, colIndex(sheet, 'Teams')).setValue((data.teams || []).join(', '));
      sheet.getRange(rows[i]._row, colIndex(sheet, 'Token')).setValue(token);
      sheet.getRange(rows[i]._row, colIndex(sheet, 'Active')).setValue(true);

      var alreadyVerified = isTrue(rows[i].Verified);
      if (!alreadyVerified) sendVerifyEmail(String(rows[i].SubID).trim(), name, email, token);
      return {
        status: 'ok',
        subId: String(rows[i].SubID).trim(),
        verified: alreadyVerified,
        message: alreadyVerified ? 'Profile updated.' : 'Confirmation email sent.'
      };
    }
  }

  var id = newId('s');
  var tok = makeToken();
  sheet.appendRow([
    id, name, email, data.phone || '',
    (data.levels || []).join(', '), (data.teams || []).join(', '),
    false, tok, true, '', new Date(), 0, ''
  ]);

  sendVerifyEmail(id, name, email, tok);
  return { status: 'ok', subId: id, verified: false, message: 'Confirmation email sent.' };
}

function actionUpdateProfile(data) {
  var sheet = getSheet('Subs');
  var row = findRowById(sheet, data.subId);
  if (row === -1) return { status: 'error', message: 'We could not find you on the sub list.' };

  if (data.name)  sheet.getRange(row, colIndex(sheet, 'Name')).setValue(String(data.name).trim());
  if (data.phone !== undefined) sheet.getRange(row, colIndex(sheet, 'Phone')).setValue(data.phone);
  if (data.levels) sheet.getRange(row, colIndex(sheet, 'Levels')).setValue(data.levels.join(', '));
  if (data.teams)  sheet.getRange(row, colIndex(sheet, 'Teams')).setValue(data.teams.join(', '));

  return { status: 'ok', message: 'Profile updated.' };
}

/** Captain adds someone by hand. Active right away — the captain vouched for them. */
function actionAddSub(data) {
  var name  = String(data.name || '').trim();
  var email = normEmail(data.email);
  if (!name)  return { status: 'error', message: 'Name is required.' };
  if (!email || email.indexOf('@') === -1) return { status: 'error', message: 'A valid email is required.' };

  var sheet = getSheet('Subs');
  var rows = sheetToObjects(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (normEmail(rows[i].Email) === email) {
      return { status: 'error', message: name + ' is already on the sub list.' };
    }
  }

  var id = newId('s');
  var tok = makeToken();
  sheet.appendRow([
    id, name, email, data.phone || '',
    (data.levels || []).join(', '), (data.teams || []).join(', '),
    true, tok, true, String(data.addedBy || 'a captain'), new Date(), 0, ''
  ]);

  sendAddedByCaptainEmail(id, name, email, tok, String(data.addedBy || 'A captain'),
                          (data.levels || []).join(', '), (data.teams || []));
  return { status: 'ok', subId: id, message: name + ' added and notified.' };
}

/* ═══════════════════════════════════════════════════════
   REQUESTS
   ═══════════════════════════════════════════════════════ */

function actionPostRequest(data) {
  var notified = data.notified || [];
  if (!notified.length) return { status: 'error', message: 'Pick at least one sub to notify.' };

  var id = newId('r');
  getSheet('Requests').appendRow([
    id, data.teamId, data.level, data.date, data.time,
    data.location || '', data.opponent || '', data.line || '', data.notes || '',
    data.postedBy || '', new Date(), notified.join(', '),
    'open', '', '', false
  ]);

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

  sheet.getRange(row, statusCol).setValue('filled');
  sheet.getRange(row, colIndex(sheet, 'Claimed By')).setValue(sub.id);
  sheet.getRange(row, colIndex(sheet, 'Claimed At')).setValue(new Date());

  bumpSubCount(sub.id, 1);
  logHistory(data.id, sub);
  notifyCaptainOfClaim(data.id, sub);

  return { status: 'ok', message: 'Confirmed — the captain has been notified.' };
}

function actionWithdraw(data) {
  var sheet = getSheet('Requests');
  var row = findRowById(sheet, data.id);
  if (row === -1) return { status: 'error', message: 'That match is no longer listed.' };

  var sub = subById(data.subId);
  sheet.getRange(row, colIndex(sheet, 'Status')).setValue('open');
  sheet.getRange(row, colIndex(sheet, 'Claimed By')).setValue('');
  sheet.getRange(row, colIndex(sheet, 'Claimed At')).setValue('');
  sheet.getRange(row, colIndex(sheet, 'Nudged')).setValue(false);

  if (sub) bumpSubCount(sub.id, -1);
  notifyCaptainOfWithdrawal(data.id, sub);

  return { status: 'ok', message: 'The captain has been notified and the spot is open again.' };
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
      getSheet('History').appendRow([
        requestId, team.name, reqs[i].date, sub.name, sub.email, new Date(), reqs[i].postedBy
      ]);
      return;
    }
  }
}

/* ═══════════════════════════════════════════════════════
   EMAIL LINK HANDLERS
   ═══════════════════════════════════════════════════════ */

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
          "You're subbing for " + team.name + ' on ' + prettyDate(reqs[i].date) +
          ' at ' + prettyTime(reqs[i].time) + ', ' + reqs[i].location +
          '. ' + reqs[i].postedBy + ' has been notified.');
      }
    }
    return htmlPage("You're in", 'The captain has been notified.');
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

function emailFooter(sub) {
  var url = webAppUrl();
  return '\n\n─────────────\n' +
    'Stop getting these: ' + url + '?action=optOut&sub=' + sub.id + '&token=' + subToken(sub.id);
}

function subToken(subId) {
  var sheet = getSheet('Subs');
  var row = findRowById(sheet, subId);
  if (row === -1) return '';
  return String(sheet.getRange(row, colIndex(sheet, 'Token')).getValue()).trim();
}

function sendVerifyEmail(subId, name, email, token) {
  var url = webAppUrl() + '?action=verify&sub=' + subId + '&token=' + token;
  var body =
    'Hi ' + name + ',\n\n' +
    'Confirm your spot on the tennis sub list by opening this link:\n\n' + url + '\n\n' +
    "Until you do, you won't receive any match requests.\n\n" +
    "If you didn't sign up, just ignore this email.";
  try {
    MailApp.sendEmail({ to: email, subject: 'Confirm your spot on the sub list', body: body });
  } catch (err) {
    Logger.log('Verify email failed for ' + email + ': ' + err);
  }
}

function sendAddedByCaptainEmail(subId, name, email, token, addedBy, levels, teamIds) {
  var appUrl = getConfig('AppUrl');
  var teamNames = [];
  for (var i = 0; i < teamIds.length; i++) teamNames.push(teamById(teamIds[i]).name);

  var body =
    'Hi ' + name + ',\n\n' +
    addedBy + ' added you to the tennis sub list.\n\n' +
    (levels ? 'Levels: ' + levels + '\n' : '') +
    (teamNames.length ? 'Teams: ' + teamNames.join(', ') + '\n' : '') +
    "\nThat means you'll get an email when one of those teams needs a sub. " +
    'You are never obligated to say yes.\n\n' +
    (appUrl ? 'Open the app to change your levels or teams:\n' + appUrl + '\n\n' : '') +
    "If you'd rather not be on the list at all:\n" +
    webAppUrl() + '?action=optOut&sub=' + subId + '&token=' + token;

  try {
    MailApp.sendEmail({ to: email, subject: "You've been added to the tennis sub list", body: body });
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
  var url = webAppUrl();
  var sent = 0;

  for (var j = 0; j < subIds.length; j++) {
    var sub = subById(subIds[j]);
    if (!sub || !sub.email || !sub.verified) continue;

    var claimUrl = url + '?action=claimFromEmail&r=' + req.id +
                   '&sub=' + sub.id + '&token=' + subToken(sub.id);

    var body =
      'Hi ' + sub.name + ',\n\n' +
      team.name + ' needs a sub.\n\n' +
      'When:  ' + prettyDate(req.date) + ' at ' + prettyTime(req.time) + '\n' +
      'Where: ' + req.location + '\n' +
      (req.opponent ? 'Vs:    ' + req.opponent + '\n' : '') +
      (req.line ? 'Line:  ' + req.line + '\n' : '') +
      'Level: ' + req.level + '\n' +
      (req.notes ? '\nNotes from ' + req.postedBy + ':\n' + req.notes + '\n' : '') +
      '\nFirst to accept gets the spot:\n' + claimUrl + '\n\n' +
      (getConfig('AppUrl') ? 'Or open the app: ' + getConfig('AppUrl') + '\n' : '') +
      emailFooter(sub);

    try {
      MailApp.sendEmail({
        to: sub.email,
        subject: 'Sub needed — ' + team.name + ', ' + prettyDate(req.date),
        body: body
      });
      sent++;
    } catch (err) {
      Logger.log('Request email failed for ' + sub.email + ': ' + err);
    }
  }
  return sent;
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
    'When:  ' + prettyDate(req.date) + ' at ' + prettyTime(req.time) + '\n' +
    'Where: ' + req.location + '\n' +
    (req.line ? 'Line:  ' + req.line + '\n' : '') +
    '\nReach ' + sub.name + ':\n' +
    '  ' + sub.email + '\n' +
    (sub.phone ? '  ' + sub.phone + '\n' : '') +
    (getConfig('AppUrl') ? '\nOpen the app: ' + getConfig('AppUrl') : '');

  var cc = getConfig('ManagerEmail');
  try {
    MailApp.sendEmail({
      to: to, cc: cc || '',
      subject: sub.name + ' is subbing — ' + team.name + ', ' + prettyDate(req.date),
      body: body
    });
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
    ' on ' + prettyDate(req.date) + ' at ' + prettyTime(req.time) + '.\n\n' +
    'The spot is open again. Open the app to notify more subs:\n' +
    (getConfig('AppUrl') || webAppUrl());

  try {
    MailApp.sendEmail({
      to: team.captainEmail,
      subject: 'Sub dropped out — ' + team.name + ', ' + prettyDate(req.date),
      body: body
    });
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
  for (var j = 0; j < req.notified.length; j++) {
    var sub = subById(req.notified[j]);
    if (!sub || !sub.email) continue;
    try {
      MailApp.sendEmail({
        to: sub.email,
        subject: 'No longer needed — ' + team.name + ', ' + prettyDate(req.date),
        body: 'Hi ' + sub.name + ',\n\n' + team.name + ' no longer needs a sub for ' +
              prettyDate(req.date) + '. Thanks anyway!' + emailFooter(sub)
      });
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
  var appUrl = getConfig('AppUrl') || webAppUrl();

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
      prettyDate(matchDate) + ' at ' + prettyTime(toTimeString(r.Time)) + '\n' +
      String(r.Location || '') + '\n\n' +
      'You notified ' + notifiedCount + ' ' + (notifiedCount === 1 ? 'sub' : 'subs') +
      ' and none have accepted.\n\n' +
      'Open the app and tap "Notify more subs" to widen the net:\n' + appUrl;

    try {
      MailApp.sendEmail({
        to: team.captainEmail,
        subject: 'Still no sub — ' + team.name + ', ' + prettyDate(matchDate),
        body: body
      });
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
