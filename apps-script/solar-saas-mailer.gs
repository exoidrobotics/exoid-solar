/**
 * Solar app – proposal email sender for ALL companies (Google Apps Script)
 *
 * The app sends each proposal here with the user's login token. This script checks the token with
 * Supabase, checks the company's plan is active and under its daily limit, then emails the client
 * (PDF attached) from this Google account, with the installer's company name as the sender name
 * and replies going to the installer.
 *
 * This is a NEW script, separate from the Exoid team mailer, so the current Exoid app is not affected.
 *
 * SETUP (signed in as the Google account that should send the emails, e.g. mgb@exoidrobotics.com):
 * 1. Go to script.google.com → New project. Name it "Solar app mailer". Paste this file. Save.
 * 2. Project Settings (gear) → Script Properties → Add:
 *      SUPABASE_URL   = your Supabase Project URL   (https://xxxx.supabase.co)
 *      SUPABASE_KEY   = your Supabase anon / publishable key
 *    Optional:
 *      FROM_ALIAS     = a "Send mail as" alias set up in Gmail (e.g. proposals@exoidrobotics.com)
 *      DAILY_LIMIT    = emails per company per 24 hours (default 40)
 * 3. Deploy → New deployment → type "Web app". Execute as: Me. Who has access: Anyone. Deploy, approve.
 * 4. Copy the Web app URL (ends in /exec) into config.js → mailerUrl.
 *    UPDATING later: Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy.
 */

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function prop_(k, d) { const v = PropertiesService.getScriptProperties().getProperty(k); return v === null || v === '' ? d : v; }
const EMAIL_RE = /^[^@\s,;<>"]+@[^@\s,;<>"]+\.[^@\s,;<>"]+$/;
function emails_(v, max) {
  const list = String(v || '').split(/[,;\s]+/).filter(Boolean);
  const bad = list.filter(x => !EMAIL_RE.test(x));
  if (bad.length) throw new Error('Not a valid email address: ' + bad[0]);
  if (list.length > max) throw new Error('Too many addresses (max ' + max + ').');
  return list;
}

// Call Supabase as the signed-in user (row-level security applies)
function sb_(path, token, method, body, minimal) {
  const url = prop_('SUPABASE_URL', '').replace(/\/$/, '') + path;
  const res = UrlFetchApp.fetch(url, {
    method: method || 'get', muteHttpExceptions: true, contentType: 'application/json',
    headers: Object.assign({ apikey: prop_('SUPABASE_KEY', ''), Authorization: 'Bearer ' + token }, minimal ? { Prefer: 'return=minimal' } : {}),
    payload: body ? JSON.stringify(body) : undefined
  });
  const code = res.getResponseCode(), text = res.getContentText();
  if (code >= 300) throw new Error('auth:' + code + ':' + text.slice(0, 200));
  return text ? JSON.parse(text) : null;
}

function doGet() { return json_({ ok: true, service: 'solar-app-mailer' }); }

function doPost(e) {
  try {
    if (!prop_('SUPABASE_URL', '') || !prop_('SUPABASE_KEY', '')) return json_({ ok: false, error: 'The mailer is not set up yet (missing Supabase settings).' });
    const d = JSON.parse(e.postData.contents || '{}');
    if (!d.token) return json_({ ok: false, error: 'Please log in again.' });

    // 1. Who is this? (rejects expired or fake tokens)
    let user;
    try { user = sb_('/auth/v1/user', d.token); } catch (x) { return json_({ ok: false, error: 'Your login has expired. Log out and log in again.' }); }
    if (!user || !user.id) return json_({ ok: false, error: 'Please log in again.' });

    // 2. Their company, plan and today's count
    const ctx = sb_('/rest/v1/rpc/mail_context', d.token, 'post', {});
    if (!ctx || !ctx.company_id) return json_({ ok: false, error: 'Your account is not linked to a company.' });
    if (!ctx.active) return json_({ ok: false, error: 'Your plan is not active, so sending is turned off.' });
    const limit = Number(prop_('DAILY_LIMIT', 40));
    if (Number(ctx.sent_today) >= limit) return json_({ ok: false, error: 'Daily email limit reached (' + limit + ' per day). Try again tomorrow or download the PDF.' });
    if (MailApp.getRemainingDailyQuota() < 3) return json_({ ok: false, error: 'The email service is busy today. Download the PDF and send it yourself.' });

    // 3. Check the message
    const to = emails_(d.to, 3), cc = emails_(d.cc, 5), bcc = emails_(d.bcc, 5);
    if (!to.length) return json_({ ok: false, error: 'Client email is missing.' });
    const subject = String(d.subject || 'Your solar energy system quote').slice(0, 200);
    const body = String(d.body || '').slice(0, 20000);
    const replyTo = EMAIL_RE.test(String(d.replyTo || '')) ? d.replyTo : (ctx.user_email || ctx.company_email || user.email);
    const opts = { name: String(ctx.company_name || 'Solar proposal').slice(0, 80), replyTo: replyTo };
    if (cc.length) opts.cc = cc.join(',');
    if (bcc.length) opts.bcc = bcc.join(',');
    if (d.pdfBase64) {
      if (d.pdfBase64.length > 14 * 1024 * 1024) return json_({ ok: false, error: 'The PDF is too large to email. Download it instead.' });
      const name = String(d.pdfName || 'Solar Proposal.pdf').replace(/[\\\/:*?"<>|]/g, '').slice(0, 120);
      opts.attachments = [Utilities.newBlob(Utilities.base64Decode(d.pdfBase64), 'application/pdf', name)];
    }
    const alias = prop_('FROM_ALIAS', '');
    if (alias && GmailApp.getAliases().indexOf(alias) >= 0) opts.from = alias;

    // 4. Send, then move this account's Sent copy to the trash so installers' client emails
    //    don't pile up in your mailbox (the log row below still counts toward the daily limit)
    const sent = GmailApp.createDraft(to.join(','), subject, body, opts).send();
    try { sent.moveToTrash(); } catch (x) {}
    try { sb_('/rest/v1/email_log', d.token, 'post', { company_id: ctx.company_id, user_id: user.id, to_addr: to.join(','), subject: subject }, true); } catch (x) {}
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err).replace(/^Error:\s*/, '') });
  }
}
