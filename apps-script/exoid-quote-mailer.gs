/**
 * Exoid Robotics – Solar quote mailer (Google Apps Script)
 * Sends each quote from the calculator to the client, CCs the office,
 * and logs it to the "Quotes" tab of the Google Sheet this script is attached to.
 *
 * SETUP (about 5 minutes, signed in as the account that should send the emails):
 * 1. Create a Google Sheet, e.g. "Exoid Solar Quotes".
 * 2. Open Extensions → Apps Script, delete the sample code, paste this file, and save.
 * 3. Click Deploy → New deployment → type "Web app".
 *    Execute as: Me.   Who has access: Anyone.   Click Deploy and approve the permissions.
 * 4. Copy the Web app URL (ends in /exec). In the calculator, go to Assumptions → Quote email
 *    and paste it into "Google Apps Script URL".
 */
const CC_ALWAYS = 'mgb@exoidrobotics.com';
const SHEET_NAME = 'Quotes';

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (!d.to || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(d.to)) {
    return ContentService.createTextOutput('missing client email');
  }
  const cc = [d.cc, CC_ALWAYS].filter(String).filter((v, i, a) => a.indexOf(v) === i).join(',');
  const html = '<div style="font-family:Arial,sans-serif;font-size:14px;white-space:pre-wrap">' +
               String(d.body).replace(/[<>&]/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;'}[c])) + '</div>';
  MailApp.sendEmail({ to: d.to, cc: cc, subject: d.subject, body: d.body, htmlBody: html, name: 'Exoid Robotics' });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Date', 'Installer', 'Client', 'Email', 'Phone', 'Address', 'System', 'kWp',
                  'Panels', 'Inverter', 'Battery', 'Price (PHP)', 'Payback (yrs)']);
  }
  const c = d.client || {};
  sh.appendRow([new Date(), d.installer, c.name, d.to, c.phone, c.address, d.system, d.kWp,
                d.panels, d.inverter, d.battery, d.price, d.payback ? Math.round(d.payback * 10) / 10 : '']);
  return ContentService.createTextOutput('ok');
}
