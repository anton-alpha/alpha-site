/**
 * Alpha waitlist — Google Apps Script backend.
 *
 * Setup (once):
 *  1. Create a Google Sheet called "Alpha waitlist" in the Alpha Workspace.
 *  2. Extensions → Apps Script. Delete the sample code, paste this file, save.
 *  3. Run the function `setup` once (toolbar ▶). Approve the permissions
 *     (Sheets + send email as you). It writes the header row.
 *  4. Deploy → New deployment → type "Web app":
 *        Execute as: Me
 *        Who has access: Anyone
 *     Copy the Web app URL (ends in /exec) and send it to Claude.
 *  5. Any later code change: Deploy → Manage deployments → edit → New version.
 *
 * Rows: id | timestamp | email | name | country | ref_by | code | confirmed | confirmed_at | token | ua | source
 */

var SHEET = 'waitlist';
var SITE = 'https://alpha.com.ky';
var FROM_NAME = 'Alpha';
var REPLY_TO = 'hello@alpha.com.ky';
var MOVE_UP_PER_REFERRAL = 50;

function setup() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['id', 'timestamp', 'email', 'name', 'country', 'ref_by', 'code', 'confirmed', 'confirmed_at', 'token', 'ua', 'source']);
    sh.setFrozenRows(1);
  }
  PropertiesService.getScriptProperties().setProperty('ready', '1');
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet() {
  return SpreadsheetApp.getActive().getSheetByName(SHEET);
}

function rand(n) {
  var s = '', a = 'abcdefghjkmnpqrstuvwxyz23456789';
  for (var i = 0; i < n; i++) s += a.charAt(Math.floor(Math.random() * a.length));
  return s;
}

function findByEmail(sh, email) {
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) if (String(data[i][2]).toLowerCase() === email) return { row: i + 1, v: data[i] };
  return null;
}

/* ---------- signup ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var p = {};
    try { p = JSON.parse(e.postData.contents || '{}'); } catch (_) { p = e.parameter || {}; }

    // bot checks: honeypot filled, or form submitted inhumanly fast
    if (p.website) return out({ ok: true });                    // honeypot: pretend success, store nothing
    if (!(Number(p.t) >= 2500)) return out({ ok: false, error: 'Please try again.' });

    var email = String(p.email || '').trim().toLowerCase();
    var name = String(p.name || '').trim().slice(0, 80);
    var country = String(p.country || '').trim().slice(0, 64);
    var refBy = String(p.ref || '').trim().slice(0, 24).replace(/[^a-z0-9]/gi, '');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return out({ ok: false, error: 'That email address doesn\'t look right.' });
    if (!name) return out({ ok: false, error: 'Please tell us your first name.' });

    var sh = sheet();
    var existing = findByEmail(sh, email);
    if (existing) {
      if (existing.v[7] === true || existing.v[7] === 'TRUE') return out({ ok: true, state: 'confirmed', position: positionOf(sh, existing.v[0]), code: existing.v[6] });
      sendConfirm(email, name, existing.v[9]);                   // resend the confirmation
      return out({ ok: true, state: 'pending' });
    }

    var id = sh.getLastRow();                                     // header is row 1 → first signup is id 1
    var code = (name.toLowerCase().replace(/[^a-z]/g, '').slice(0, 6) || 'alpha') + rand(3);
    var token = Utilities.getUuid().replace(/-/g, '');
    sh.appendRow([id, new Date(), email, name, country, refBy, code, false, '', token, String(p.ua || '').slice(0, 160), String(p.source || 'site').slice(0, 40)]);
    sendConfirm(email, name, token);
    return out({ ok: true, state: 'pending' });
  } catch (err) {
    return out({ ok: false, error: 'Something went wrong. Please try again in a minute.' });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- confirmation link ---------- */
function doGet(e) {
  var token = (e.parameter.confirm || '').replace(/[^a-f0-9]/g, '');
  if (!token) return out({ ok: true, service: 'alpha-waitlist' });
  var sh = sheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][9] === token) {
      if (!(data[i][7] === true || data[i][7] === 'TRUE')) {
        sh.getRange(i + 1, 8).setValue(true);
        sh.getRange(i + 1, 9).setValue(new Date());
      }
      var pos = positionOf(sh, data[i][0]);
      var code = data[i][6];
      var url = SITE + '/?joined=1&pos=' + pos + '&code=' + encodeURIComponent(code) + '#join';
      return HtmlService.createHtmlOutput('<!doctype html><meta http-equiv="refresh" content="0;url=' + url + '"><p style="font-family:sans-serif">Confirmed. Taking you back to Alpha…</p>');
    }
  }
  return HtmlService.createHtmlOutput('<!doctype html><p style="font-family:sans-serif">This link has expired. <a href="' + SITE + '/#join">Join again</a>.</p>');
}

/* position among confirmed signups, moved up by confirmed referrals */
function positionOf(sh, id) {
  var data = sh.getDataRange().getValues();
  var code = null, ahead = 0, refs = 0;
  for (var i = 1; i < data.length; i++) if (data[i][0] === id) code = data[i][6];
  for (var j = 1; j < data.length; j++) {
    var confirmed = data[j][7] === true || data[j][7] === 'TRUE';
    if (!confirmed) continue;
    if (data[j][0] < id) ahead++;
    if (code && data[j][5] === code && data[j][0] !== id) refs++;
  }
  return Math.max(1, ahead + 1 - refs * MOVE_UP_PER_REFERRAL);
}

function sendConfirm(email, name, token) {
  var link = ScriptApp.getService().getUrl() + '?confirm=' + token;
  var first = name.split(' ')[0];
  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:520px;margin:0 auto;color:#0B1F18">' +
    '<p style="font-size:16px">Hi ' + first + ',</p>' +
    '<p style="font-size:16px">One tap to confirm your place on the Alpha waitlist:</p>' +
    '<p style="margin:26px 0"><a href="' + link + '" style="background:#1FC48C;color:#04140E;text-decoration:none;font-weight:bold;padding:14px 24px;border-radius:999px;display:inline-block">Confirm my place</a></p>' +
    '<p style="font-size:14px;color:#5F7A6F">You\'ll see your position in the queue and your referral link straight away. If you didn\'t ask to join, ignore this email and nothing happens.</p>' +
    '<p style="font-size:12px;color:#8FA59C;margin-top:34px">Alpha Corp · Cayman Islands · <a href="' + SITE + '" style="color:#1FC48C;text-decoration:none">alpha.com.ky</a><br>Alpha will never ask you for a password or a code by email.</p>' +
    '</div>';
  var opts = { htmlBody: html, name: FROM_NAME, replyTo: REPLY_TO };
  try {
    // send from the hello@ alias when it is configured on this account ("Send mail as" in Gmail)
    if (GmailApp.getAliases().indexOf(REPLY_TO) >= 0) opts.from = REPLY_TO;
    GmailApp.sendEmail(email, 'Confirm your place on the Alpha waitlist', 'Open this email in a client that shows HTML to confirm your place.', opts);
  } catch (e) {
    MailApp.sendEmail({ to: email, subject: 'Confirm your place on the Alpha waitlist', htmlBody: html, name: FROM_NAME, replyTo: REPLY_TO });
  }
}
