/**
 * Alpha waitlist — Google Apps Script backend (v7).
 *
 * Setup / update:
 *  1. Extensions → Apps Script in the "Alpha waitlist" sheet. Replace all code with this file. Save.
 *  2. Project Settings (gear icon) → Script Properties → Add property:
 *        TURNSTILE_SECRET = <the Cloudflare Turnstile secret key>
 *     (until this is set, submissions are accepted with the basic bot checks only)
 *  3. Run `setup` once (▶). It creates/refreshes the "Waitlist" tab.
 *  4. Deploy → Manage deployments → ✎ → Version: New version → Deploy.
 */

var SHEET = 'Waitlist';
var SITE = 'https://alpha.com.ky';
var FROM_NAME = 'Alpha';
var REPLY_TO = 'hello@alpha.com.ky';
var MOVE_UP_PER_REFERRAL = 50;
var POSITION_OFFSET = 2135;   // the first signup is shown as #2,136

var C = { NUMBER: 1, JOINED: 2, FIRST: 3, SURNAME: 4, EMAIL: 5, MOBILE: 6, COUNTRY: 7, ACCOUNT: 8, BUSINESS: 9, USE: 10,
          POSITION: 11, EMAIL_OK: 12, CONFIRMED_AT: 13, REFERRED_BY: 14, INVITE_CODE: 15, INVITE_LINK: 16, SOURCE: 17, TOKEN: 18, BROWSER: 19 };
var HEADERS = ['#', 'Joined', 'First name', 'Surname', 'Email', 'Mobile', 'Country', 'Account', 'Business', 'Would use Alpha for',
               'Position', 'Email confirmed', 'Confirmed at', 'Referred by', 'Invite code', 'Invite link', 'Source', 'Token', 'Browser'];

function setup() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) sh = ss.insertSheet(SHEET, 0);
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0B1F18').setFontColor('#FFFFFF').setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 32);
  var widths = [70, 150, 110, 120, 240, 150, 150, 90, 160, 260, 80, 120, 150, 110, 110, 230, 90, 60, 60];
  for (var i = 0; i < widths.length; i++) sh.setColumnWidth(i + 1, widths[i]);
  sh.getRange('B2:B').setNumberFormat('dd mmm yyyy hh:mm');
  sh.getRange('F2:F').setNumberFormat('@');
  sh.getRange('A2:A').setHorizontalAlignment('center').setNumberFormat('#,##0');
  sh.getRange('K2:K').setHorizontalAlignment('center').setNumberFormat('#,##0');
  sh.getRange('L2:L').setHorizontalAlignment('center');
  sh.getRange('M2:M').setNumberFormat('dd mmm yyyy hh:mm');
  sh.hideColumns(C.TOKEN, 2);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, Math.max(2, sh.getLastRow()), HEADERS.length).createFilter();
}

function out(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function sheet() { return SpreadsheetApp.getActive().getSheetByName(SHEET); }
function rand(n) { var s = '', a = 'abcdefghjkmnpqrstuvwxyz23456789'; for (var i = 0; i < n; i++) s += a.charAt(Math.floor(Math.random() * a.length)); return s; }

function findByEmail(data, email) {
  for (var i = 1; i < data.length; i++) if (String(data[i][C.EMAIL - 1]).toLowerCase() === email) return i;
  return -1;
}

/* Cloudflare Turnstile: verifies the token the form sent. Returns true if no secret is configured yet. */
function turnstileOk(token, ip) {
  var secret = PropertiesService.getScriptProperties().getProperty('TURNSTILE_SECRET');
  if (!secret) return true;
  if (!token) return false;
  try {
    var r = UrlFetchApp.fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', {
      method: 'post', payload: { secret: secret, response: token, remoteip: ip || '' }, muteHttpExceptions: true });
    var j = JSON.parse(r.getContentText());
    return j.success === true;
  } catch (e) { return false; }
}

/* current position: base number minus 50 per referral */
function positionFor(data, number, code) {
  var refs = 0;
  for (var j = 1; j < data.length; j++) if (code && data[j][C.REFERRED_BY - 1] === code) refs++;
  return Math.max(POSITION_OFFSET + 1, number - refs * MOVE_UP_PER_REFERRAL);
}

/* ---------- signup ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var p = {};
    try { p = JSON.parse(e.postData.contents || '{}'); } catch (_) { p = e.parameter || {}; }

    if (p.website) return out({ ok: true, state: 'ok' });                     // honeypot: pretend success, store nothing
    if (!(Number(p.t) >= 2500)) return out({ ok: false, error: 'Please try again.' });
    if (!turnstileOk(p.turnstile, p.ip)) return out({ ok: false, error: 'We couldn\'t verify you\'re human. Please try again.' });

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
    var data = sh.getDataRange().getValues();
    var i = findByEmail(data, email);
    if (i > 0) {                                                                // already on the list: show their place again
      var row = data[i];
      var pos = positionFor(data, row[C.NUMBER - 1], row[C.INVITE_CODE - 1]);
      sh.getRange(i + 1, C.POSITION).setValue(pos);
      return out({ ok: true, state: 'ok', position: pos, code: row[C.INVITE_CODE - 1], link: row[C.INVITE_LINK - 1], returning: true });
    }

    var number = POSITION_OFFSET + sh.getLastRow();                             // header is row 1 → first signup is 2136
    var code = (first.toLowerCase().replace(/[^a-z]/g, '').slice(0, 6) || 'alpha') + rand(3);
    var link = SITE + '/?r=' + code;
    var r = [];
    r[C.NUMBER - 1] = number; r[C.JOINED - 1] = new Date(); r[C.FIRST - 1] = first; r[C.SURNAME - 1] = surname; r[C.EMAIL - 1] = email;
    r[C.MOBILE - 1] = mobile; r[C.COUNTRY - 1] = country; r[C.ACCOUNT - 1] = account; r[C.BUSINESS - 1] = business; r[C.USE - 1] = use;
    var token = Utilities.getUuid().replace(/-/g, '');
    r[C.POSITION - 1] = number; r[C.EMAIL_OK - 1] = false; r[C.CONFIRMED_AT - 1] = ''; r[C.REFERRED_BY - 1] = refBy; r[C.INVITE_CODE - 1] = code; r[C.INVITE_LINK - 1] = link;
    r[C.SOURCE - 1] = source; r[C.TOKEN - 1] = token; r[C.BROWSER - 1] = browser;
    sh.appendRow(r);
    sh.getRange(sh.getLastRow(), C.MOBILE).setNumberFormat('@').setValue(mobile);

    // the referrer moves up 50: refresh their position
    if (refBy) for (var k = 1; k < data.length; k++) if (data[k][C.INVITE_CODE - 1] === refBy) {
      sh.getRange(k + 1, C.POSITION).setValue(positionFor(sh.getDataRange().getValues(), data[k][C.NUMBER - 1], refBy));
    }

    sendWelcome(email, first, number, link, token);
    return out({ ok: true, state: 'ok', position: number, code: code, link: link });
  } catch (err) {
    return out({ ok: false, error: 'Something went wrong. Please try again in a minute.' });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- email confirmation: a standalone page, no redirect ---------- */
function doGet(e) {
  var token = (e.parameter.confirm || '').replace(/[^a-f0-9]/g, '');
  if (!token) return out({ ok: true, service: 'alpha-waitlist' });
  var sh = sheet();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][C.TOKEN - 1] === token) {
      if (!(data[i][C.EMAIL_OK - 1] === true || data[i][C.EMAIL_OK - 1] === 'TRUE')) {
        sh.getRange(i + 1, C.EMAIL_OK).setValue(true);
        sh.getRange(i + 1, C.CONFIRMED_AT).setValue(new Date());
      }
      return confirmPage(data[i][C.FIRST - 1], Number(data[i][C.POSITION - 1]) || Number(data[i][C.NUMBER - 1]), data[i][C.INVITE_LINK - 1], true);
    }
  }
  return confirmPage('', 0, '', false);
}

function confirmPage(first, position, link, ok) {
  var body = ok
    ? '<div class="k">EMAIL CONFIRMED</div><h1>Thank you, ' + esc(first) + '.</h1><p>Your place on the Alpha waitlist is reserved.</p><div class="n">#' + position.toLocaleString('en-US') + '</div>' +
      '<p>Each friend who joins through your invite link moves you up 50 places and earns you 25 $AG when you both open an account.</p>' +
      '<code>' + esc(link) + '</code><p class="s">We\'ll email you once, when it\'s your turn. You can close this page.</p>'
    : '<div class="k">LINK EXPIRED</div><h1>That link no longer works.</h1><p>Your place is safe; this link has simply been used or has expired. If you need anything, write to <a href="mailto:hello@alpha.com.ky">hello@alpha.com.ky</a>.</p>';
  var html = '<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Alpha</title>' +
    '<style>body{margin:0;min-height:100vh;display:grid;place-items:center;background:#050C0A;color:#F7FBF9;font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;padding:24px;box-sizing:border-box}' +
    '.c{max-width:480px;width:100%;background:#0F1A17;border:1px solid rgba(255,255,255,.08);border-radius:24px;padding:36px 32px}.wm{display:block;width:96px;height:19px;margin-bottom:26px}' +
    '.k{font-family:Menlo,Consolas,monospace;font-size:11px;letter-spacing:.16em;color:#1FC48C;margin-bottom:12px}h1{margin:0 0 8px;font-size:26px;letter-spacing:-.5px}p{margin:0 0 14px;color:#CFE6DC;line-height:1.55;font-size:15px}' +
    '.n{font-size:56px;font-weight:700;letter-spacing:-2px;line-height:1;color:#5FE3B3;margin:6px 0 18px}code{display:block;background:rgba(0,0,0,.3);border-radius:12px;padding:12px 14px;font-size:13px;color:#fff;word-break:break-all;margin-bottom:14px}' +
    '.s{color:#8FA59C;font-size:13px}a{color:#5FE3B3;text-decoration:none}.f{margin-top:16px;font-size:11px;color:#5F7A6F;text-align:center;max-width:480px}</style></head><body><div>' +
    '<div class="c"><img class="wm" src="' + SITE + '/wordmark.png" alt="ALPHA">' + body + '</div>' +
    '<div class="f">One account. Anywhere on earth. · Alpha Corp · Cayman Islands<br>Alpha will never ask you for a password or a code by email.</div></div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('Alpha').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

function sendWelcome(email, first, number, link, token) {
  var confirmLink = ScriptApp.getService().getUrl() + '?confirm=' + token;
  var html =
    '<div style="background:#050C0A;padding:40px 20px;font-family:Arial,Helvetica,sans-serif">' +
    '<div style="max-width:520px;margin:0 auto;background:#0F1A17;border:1px solid rgba(255,255,255,0.08);border-radius:22px;padding:36px 32px;color:#F7FBF9">' +
    '<img src="' + SITE + '/wordmark.png" width="96" height="19" alt="ALPHA" style="display:block;border:0;width:96px;height:19px;margin-bottom:26px">' +
    '<h1 style="margin:0 0 6px;font-size:26px;letter-spacing:-0.5px;color:#F7FBF9">You\'re in, ' + first + '.</h1>' +
    '<p style="margin:0 0 22px;font-size:15px;line-height:1.55;color:#CFE6DC">Your place in the queue:</p>' +
    '<p style="margin:0 0 22px;font-size:54px;line-height:1;font-weight:bold;letter-spacing:-2px;color:#5FE3B3">#' + number.toLocaleString('en-US') + '</p>' +
    '<p style="margin:0 0 10px;font-size:15px;line-height:1.55;color:#CFE6DC">Founding members get in first, country by country. Every friend who joins through your link moves you up 50 places and earns you 25 $AG when you both open an account.</p>' +
    '<p style="margin:0 0 26px"><a href="' + confirmLink + '" style="background:#1FC48C;color:#04140E;text-decoration:none;font-weight:bold;font-size:15px;padding:15px 26px;border-radius:999px;display:inline-block">Confirm my email</a></p>' +
    '<p style="margin:0 0 6px;font-size:13px;color:#8FA59C">Your invite link</p><p style="margin:0 0 22px;font-size:13px;color:#F7FBF9;word-break:break-all">' + link + '</p>' +
    '<p style="margin:0;font-size:13px;line-height:1.5;color:#8FA59C">We\'ll email you once, when it\'s your turn. If you didn\'t join the Alpha waitlist, ignore this email.</p>' +
    '</div>' +
    '<p style="max-width:520px;margin:18px auto 0;font-size:11px;line-height:1.5;color:#5F7A6F;text-align:center">One account. Anywhere on earth. · Alpha Corp · Cayman Islands · <a href="' + SITE + '" style="color:#5FE3B3;text-decoration:none">alpha.com.ky</a><br>Alpha will never ask you for a password or a code by email.</p>' +
    '</div>';
  var opts = { htmlBody: html, name: FROM_NAME, replyTo: REPLY_TO };
  try {
    if (GmailApp.getAliases().indexOf(REPLY_TO) >= 0) opts.from = REPLY_TO;
    GmailApp.sendEmail(email, 'You\'re on the Alpha waitlist: #' + number.toLocaleString('en-US'), 'You are #' + number + ' on the Alpha waitlist. Your invite link: ' + link, opts);
  } catch (e) {
    MailApp.sendEmail({ to: email, subject: 'You\'re on the Alpha waitlist: #' + number.toLocaleString('en-US'), htmlBody: html, name: FROM_NAME, replyTo: REPLY_TO });
  }
}
