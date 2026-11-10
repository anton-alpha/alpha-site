/**
 * Alpha waitlist: one-off notice of the referral change (OI-206, Anton, 10 Nov 2026).
 *
 * Anton runs this himself; nothing here runs on its own and nothing is scheduled.
 *  1. Extensions → Apps Script in the "Alpha waitlist" sheet → ＋ → Script → name it ReferralNotice.
 *     Paste this file. Save. (Code.gs is not changed by this file.)
 *  2. Run `sendReferralNoticeTest` (▶). It sends one copy, to the account running it, and nothing else.
 *     Check that copy.
 *  3. Run `sendReferralNotice` (▶). It emails every row of the Waitlist tab that has an email address and
 *     no date in the "Referral notice sent" column, writes the time it sent into that column, and stops
 *     before the daily sending quota runs out. If it stops early, run it again the next day: rows already
 *     sent are skipped, so nobody gets it twice.
 *  4. When it reports 0 left, the site's stage-1 wording can be published.
 */

var NOTICE_HEADER = 'Referral notice sent';
var NOTICE_SUBJECT = 'A change to how referral $AG works';

function noticeBodies_(first) {
  var hi = first ? 'Hi ' + first + ',' : 'Hi,';
  var paras = [
    'Thank you for joining the Alpha waitlist.',
    'We are changing how the referral reward works, and we want you to hear it from us first.',
    'Until now, the site said you earn 25 $AG when someone you invite opens an account. From now on, when someone you refer becomes a verified Alpha customer, 25 $AG is recorded as owed to you. The limit is 40 people a year.',
    '$AG rewards are not on at launch. The $AG you are owed is added to your account when $AG rewards are switched on, under the programme terms.',
    'Your place on the waitlist does not change, and each friend who joins through your link still moves you up 50 places.',
    'The full terms are on our site. If you have a question, just reply to this email.',
    'The Alpha team'
  ];
  var text = hi + '\n\n' + paras.join('\n\n') + '\n\nAlpha will never ask you for a password or a code by email.';
  var html = '<div style="font-family:Helvetica,Arial,sans-serif;font-size:15px;line-height:1.5;color:#0B1F18">' +
    '<p>' + esc(hi) + '</p>' +
    paras.map(function (p) { return '<p>' + esc(p) + '</p>'; }).join('') +
    '<p style="font-size:12px;color:#5B6B64">Alpha will never ask you for a password or a code by email.</p></div>';
  return { text: text, html: html };
}

function sendNotice_(email, first) {
  var b = noticeBodies_(first);
  var opts = { htmlBody: b.html, name: FROM_NAME, replyTo: REPLY_TO };
  if (GmailApp.getAliases().indexOf(REPLY_TO) >= 0) opts.from = REPLY_TO;
  GmailApp.sendEmail(email, NOTICE_SUBJECT, b.text, opts);
}

function noticeColumn_(sh) {
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var i = headers.indexOf(NOTICE_HEADER);
  if (i >= 0) return i + 1;
  sh.getRange(1, lastCol + 1).setValue(NOTICE_HEADER).setFontWeight('bold');
  return lastCol + 1;
}

/** Sends one copy, to the account running the script. Sends to nobody else and writes nothing to the sheet. */
function sendReferralNoticeTest() {
  var me = Session.getActiveUser().getEmail();
  if (!me) throw new Error('Could not read the running account\'s email address.');
  sendNotice_(me, 'Anton');
  Logger.log('Test copy sent to ' + me + '.');
}

/** Sends the notice to every waitlist row not yet sent, within today's quota. Safe to run again. */
function sendReferralNotice() {
  var sh = sheet();
  var col = noticeColumn_(sh);
  var last = sh.getLastRow();
  if (last < 2) { Logger.log('No rows.'); return; }
  var emails = sh.getRange(2, C.EMAIL, last - 1, 1).getValues();
  var firsts = sh.getRange(2, C.FIRST, last - 1, 1).getValues();
  var sent = sh.getRange(2, col, last - 1, 1).getValues();
  var done = 0, left = 0, skipped = 0;
  for (var r = 0; r < emails.length; r++) {
    var email = String(emails[r][0] || '').trim();
    if (!email || email.indexOf('@') < 0) { skipped++; continue; }
    if (sent[r][0]) continue;
    if (MailApp.getRemainingDailyQuota() < 5) { left++; continue; }
    sendNotice_(email, String(firsts[r][0] || '').trim());
    sh.getRange(r + 2, col).setValue(new Date());
    done++;
  }
  Logger.log('Sent ' + done + ' today. ' + left + ' left for tomorrow (quota). ' + skipped + ' rows without an email.');
}
