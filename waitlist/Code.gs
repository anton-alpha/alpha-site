/**
 * Alpha waitlist — Google Apps Script backend (v5).
 *
 * Setup / update:
 *  1. Extensions → Apps Script in the "Alpha waitlist" sheet. Replace all code with this file. Save.
 *  2. Run `setup` once (▶). It creates a tab called "Waitlist" with tidy columns and formatting.
 *     (An older tab called "waitlist" can be deleted by hand.)
 *  3. Deploy → Manage deployments → ✎ → Version: New version → Deploy.
 */

var SHEET = 'Waitlist';
var SITE = 'https://alpha.com.ky';
var FROM_NAME = 'Alpha';
var REPLY_TO = 'hello@alpha.com.ky';
var MOVE_UP_PER_REFERRAL = 50;
var POSITION_OFFSET = 2135;   // positions shown to customers start at 2136

// Column layout (1-based). Customer-facing first, housekeeping last.
var C = { ID: 1, JOINED: 2, FIRST: 3, SURNAME: 4, EMAIL: 5, MOBILE: 6, COUNTRY: 7, ACCOUNT: 8, BUSINESS: 9, USE: 10,
          CONFIRMED: 11, CONFIRMED_AT: 12, POSITION: 13, REFERRED_BY: 14, INVITE_CODE: 15, SOURCE: 16, TOKEN: 17, BROWSER: 18 };
var HEADERS = ['#', 'Joined', 'First name', 'Surname', 'Email', 'Mobile', 'Country', 'Account', 'Business', 'Would use Alpha for',
               'Confirmed', 'Confirmed at', 'Position', 'Referred by', 'Invite code', 'Source', 'Token', 'Browser'];

function setup() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET, 0); }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  var head = sh.getRange(1, 1, 1, HEADERS.length);
  head.setFontWeight('bold').setBackground('#0B1F18').setFontColor('#FFFFFF').setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 32);
  var widths = [50, 150, 110, 120, 240, 150, 150, 90, 160, 260, 90, 150, 80, 110, 110, 90, 60, 60];
  for (var i = 0; i < widths.length; i++) sh.setColumnWidth(i + 1, widths[i]);
  sh.getRange('B2:B').setNumberFormat('dd mmm yyyy hh:mm');
  sh.getRange('L2:L').setNumberFormat('dd mmm yyyy hh:mm');
  sh.getRange('F2:F').setNumberFormat('@');
  sh.getRange('A2:A').setHorizontalAlignment('center');
  sh.getRange('K2:K').setHorizontalAlignment('center');
  sh.getRange('M2:M').setHorizontalAlignment('center');
  sh.hideColumns(C.TOKEN, 2);   // token + browser: needed, not for reading
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, Math.max(2, sh.getLastRow()), HEADERS.length).createFilter();
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function sheet() { return SpreadsheetApp.getActive().getSheetByName(SHEET); }
function rand(n) { var s = '', a = 'abcdefghjkmnpqrstuvwxyz23456789'; for (var i = 0; i < n; i++) s += a.charAt(Math.floor(Math.random() * a.length)); return s; }
function isTrue(v) { return v === true || v === 'TRUE'; }

function findByEmail(sh, email) {
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) if (String(data[i][C.EMAIL - 1]).toLowerCase() === email) return { row: i + 1, v: data[i] };
  return null;
}

/* ---------- signup ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var p = {};
    try { p = JSON.parse(e.postData.contents || '{}'); } catch (_) { p = e.parameter || {}; }

    if (p.website) return out({ ok: true });                       // honeypot: pretend success, store nothing
    if (!(Number(p.t) >= 2500)) return out({ ok: false, error: 'Please try again.' });

    var email = String(p.email || '').trim().toLowerCase();
    var first = String(p.name || '').trim().slice(0, 60);
    var surname = String(p.surname || '').trim().slice(0, 60);
    var mobile = String(p.phone || '').replace(/[^0-9+() .-]/g, '').trim().slice(0, 24);
    var country = String(p.country || '').trim().slice(0, 64);
    var account = String(p.type || 'personal').toLowerCase() === 'business' ? 'Business' : 'Personal';
    var business = String(p.business || '').trim().slice(0, 100);
    var use = String(p.use || '').trim().slice(0, 200);
    var refBy = String(p.ref || '').trim().slice(0, 24).replace(/[^a-z0-9]/gi, '');
    var source = String(p.source || 'site').slice(0, 40);
    var browser = String(p.ua || '').slice(0, 160);

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return out({ ok: false, error: 'That email address doesn\'t look right.' });
    if (!first) return out({ ok: false, error: 'Please tell us your first name.' });
    if (!surname) return out({ ok: false, error: 'Please tell us your surname.' });
    if (mobile.replace(/\D/g, '').length < 7) return out({ ok: false, error: 'Please enter your mobile number with the country code.' });
    if (account === 'Business' && !business) return out({ ok: false, error: 'Please tell us your business name.' });

    var sh = sheet();
    var existing = findByEmail(sh, email);
    if (existing) {
      if (isTrue(existing.v[C.CONFIRMED - 1])) return out({ ok: true, state: 'confirmed', position: positionOf(sh, existing.v[C.ID - 1]), code: existing.v[C.INVITE_CODE - 1] });
      sendConfirm(email, first, existing.v[C.TOKEN - 1]);
      return out({ ok: true, state: 'pending' });
    }

    var id = sh.getLastRow();                                        // header is row 1 → first signup is #1
    var code = (first.toLowerCase().replace(/[^a-z]/g, '').slice(0, 6) || 'alpha') + rand(3);
    var token = Utilities.getUuid().replace(/-/g, '');
    var row = [];
    row[C.ID - 1] = id; row[C.JOINED - 1] = new Date(); row[C.FIRST - 1] = first; row[C.SURNAME - 1] = surname; row[C.EMAIL - 1] = email;
    row[C.MOBILE - 1] = mobile; row[C.COUNTRY - 1] = country; row[C.ACCOUNT - 1] = account; row[C.BUSINESS - 1] = business; row[C.USE - 1] = use;
    row[C.CONFIRMED - 1] = false; row[C.CONFIRMED_AT - 1] = ''; row[C.POSITION - 1] = ''; row[C.REFERRED_BY - 1] = refBy; row[C.INVITE_CODE - 1] = code;
    row[C.SOURCE - 1] = source; row[C.TOKEN - 1] = token; row[C.BROWSER - 1] = browser;
    sh.appendRow(row);
    sh.getRange(sh.getLastRow(), C.MOBILE).setNumberFormat('@').setValue(mobile);   // keep "+" and leading zeros
    sendConfirm(email, first, token);
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
    if (data[i][C.TOKEN - 1] === token) {
      if (!isTrue(data[i][C.CONFIRMED - 1])) {
        sh.getRange(i + 1, C.CONFIRMED).setValue(true);
        sh.getRange(i + 1, C.CONFIRMED_AT).setValue(new Date());
      }
      var pos = positionOf(sh, data[i][C.ID - 1]);
      sh.getRange(i + 1, C.POSITION).setValue(pos);
      var url = SITE + '/?joined=1&pos=' + pos + '&code=' + encodeURIComponent(data[i][C.INVITE_CODE - 1]) + '#join';
      return HtmlService.createHtmlOutput('<!doctype html><meta http-equiv="refresh" content="0;url=' + url + '"><p style="font-family:sans-serif">Confirmed. Taking you back to Alpha…</p>');
    }
  }
  return HtmlService.createHtmlOutput('<!doctype html><p style="font-family:sans-serif">This link has expired. <a href="' + SITE + '/#join">Join again</a>.</p>');
}

/* position among confirmed signups (offset), moved up by confirmed referrals */
function positionOf(sh, id) {
  var data = sh.getDataRange().getValues();
  var code = null, ahead = 0, refs = 0;
  for (var i = 1; i < data.length; i++) if (data[i][C.ID - 1] === id) code = data[i][C.INVITE_CODE - 1];
  for (var j = 1; j < data.length; j++) {
    if (!isTrue(data[j][C.CONFIRMED - 1])) continue;
    if (data[j][C.ID - 1] < id) ahead++;
    if (code && data[j][C.REFERRED_BY - 1] === code && data[j][C.ID - 1] !== id) refs++;
  }
  return POSITION_OFFSET + Math.max(1, ahead + 1 - refs * MOVE_UP_PER_REFERRAL);
}

function sendConfirm(email, first, token) {
  var link = ScriptApp.getService().getUrl() + '?confirm=' + token;
  var html =
    '<div style="background:#050C0A;padding:40px 20px;font-family:Arial,Helvetica,sans-serif">' +
    '<div style="max-width:520px;margin:0 auto;background:#0F1A17;border:1px solid rgba(255,255,255,0.08);border-radius:22px;padding:36px 32px;color:#F7FBF9">' +
    '<img src="' + SITE + '/wordmark.png" width="96" height="19" alt="ALPHA" style="display:block;border:0;width:96px;height:19px;margin-bottom:26px">' +
    '<h1 style="margin:0 0 10px;font-size:26px;letter-spacing:-0.5px;color:#F7FBF9">You\'re nearly in, ' + first + '.</h1>' +
    '<p style="margin:0 0 24px;font-size:15px;line-height:1.55;color:#CFE6DC">One tap confirms it\'s really you and reserves your place in the queue. Founding members get in first, country by country.</p>' +
    '<p style="margin:0 0 26px"><a href="' + link + '" style="background:#1FC48C;color:#04140E;text-decoration:none;font-weight:bold;font-size:15px;padding:15px 26px;border-radius:999px;display:inline-block">Confirm and see my place</a></p>' +
    '<p style="margin:0;font-size:13px;line-height:1.5;color:#8FA59C">If you didn\'t ask to join, ignore this and nothing happens.</p>' +
    '</div>' +
    '<p style="max-width:520px;margin:18px auto 0;font-size:11px;line-height:1.5;color:#5F7A6F;text-align:center">One account. Anywhere on earth. · Alpha Corp · Cayman Islands · <a href="' + SITE + '" style="color:#5FE3B3;text-decoration:none">alpha.com.ky</a><br>Alpha will never ask you for a password or a code by email.</p>' +
    '</div>';
  var opts = { htmlBody: html, name: FROM_NAME, replyTo: REPLY_TO };
  try {
    if (GmailApp.getAliases().indexOf(REPLY_TO) >= 0) opts.from = REPLY_TO;
    GmailApp.sendEmail(email, 'Confirm your place on the Alpha waitlist', 'Open this email in a client that shows HTML to confirm your place.', opts);
  } catch (e) {
    MailApp.sendEmail({ to: email, subject: 'Confirm your place on the Alpha waitlist', htmlBody: html, name: FROM_NAME, replyTo: REPLY_TO });
  }
}
