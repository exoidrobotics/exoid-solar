/**
 * Exoid Robotics – Solar quote mailer (Google Apps Script)
 * Sends each quote from the calculator to the client (with the PDF proposal attached), CCs the office,
 * logs it to the "Quotes" tab of the Google Sheet this script is attached to, and stores the
 * shared equipment catalog (installer edits + product photos) as a file in your Google Drive.
 *
 * SETUP (about 5 minutes, signed in as the account that should send the emails):
 * 1. Create a Google Sheet, e.g. "Exoid Solar Quotes".
 * 2. Open Extensions → Apps Script, delete the sample code, paste this file, and save.
 * 3. Click Deploy → New deployment → type "Web app".
 *    Execute as: Me.   Who has access: Anyone.   Click Deploy and approve the permissions.
 *    UPDATING an existing script: paste the new code, Save, then Deploy → Manage deployments →
 *    pencil (Edit) → Version: New version → Deploy. The URL stays the same.
 * 4. Copy the Web app URL (ends in /exec). In the calculator, go to Assumptions → Quote email
 *    and paste it into "Google Apps Script URL".
 */
const CC_ALWAYS = 'mgb@exoidrobotics.com';
const SHEET_NAME = 'Quotes';

const CATALOG_FILE = 'Exoid Solar Catalog (do not delete).json';

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function catalogFile_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('CATALOG_FILE_ID');
  if (id) { try { return DriveApp.getFileById(id); } catch (err) {} }
  return null;
}

// GET ?action=catalog → shared equipment catalog edits (models added/edited by installers, photos included)
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'catalog') {
    const f = catalogFile_();
    return json_({ ok: true, edits: f ? JSON.parse(f.getBlob().getDataAsString() || 'null') : null });
  }
  return json_({ ok: true });
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.action === 'saveCatalog') {
    const text = JSON.stringify(d.edits || { over: {}, added: [] });
    let f = catalogFile_();
    if (f) f.setContent(text);
    else {
      f = DriveApp.createFile(CATALOG_FILE, text, MimeType.PLAIN_TEXT);
      PropertiesService.getScriptProperties().setProperty('CATALOG_FILE_ID', f.getId());
    }
    return json_({ ok: true });
  }
  if (!d.to || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(d.to)) {
    return ContentService.createTextOutput('missing client email');
  }
  const cc = [d.cc, CC_ALWAYS].filter(String).filter((v, i, a) => a.indexOf(v) === i).join(',');
  const html = '<div style="font-family:Arial,sans-serif;font-size:14px;white-space:pre-wrap">' +
               String(d.body).replace(/[<>&]/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;'}[c])) + '</div>';
  const mail = { to: d.to, cc: cc, subject: d.subject, body: d.body, htmlBody: html, name: 'Exoid Robotics' };
  if (d.pdfBase64) {
    mail.attachments = [Utilities.newBlob(Utilities.base64Decode(d.pdfBase64), 'application/pdf', d.pdfName || 'Exoid Solar Proposal.pdf')];
  }
  MailApp.sendEmail(mail);

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
