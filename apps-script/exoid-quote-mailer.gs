/**
 * Exoid Robotics – Solar quote mailer (Google Apps Script)
 * Sends each quote from the calculator to the client (with the PDF proposal attached), BCCs the team,
 * logs it to the "Quotes" tab of the Google Sheet this script is attached to, and stores the
 * shared equipment catalog (installer edits + product photos) as a file in your Google Drive,
 * and keeps saved proposals on the "Proposals" tab so any installer can reopen, duplicate or delete them.
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
const BCC_ALWAYS = 'mgb@exoidrobotics.com';   // hidden copy of every quote
const REPLY_TO = 'mgb@exoidrobotics.com';     // client replies go here
const SHEET_NAME = 'Quotes';

const CATALOG_FILE = 'Exoid Solar Catalog (do not delete).json';

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function catalogFile_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('CATALOG_FILE_ID');
  if (id) { try { return DriveApp.getFileById(id); } catch (err) {} }
  return null;
}

// ---- Saved proposals: one row per proposal on the "Proposals" tab ----
const PROPOSALS_SHEET = 'Proposals';
function propSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PROPOSALS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PROPOSALS_SHEET);
    sh.appendRow(['ID', 'Last saved', 'Prepared by', 'Client', 'Client address', 'Project address', 'System', 'kWp', 'Price (PHP)', 'Data (do not edit)']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function propRow_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function listProposals_() {
  const sh = propSheet_(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 9).getValues().filter(r => r[0]).map(r => ({
    id: String(r[0]), updated: r[1] instanceof Date ? r[1].toISOString() : String(r[1]), installer: r[2], client: r[3],
    clientAddress: r[4], projectAddress: r[5], system: r[6], kWp: r[7], price: r[8] }));
}

// GET ?action=catalog → shared equipment catalog edits (models added/edited by installers, photos included)
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'proposals') return json_({ ok: true, proposals: listProposals_() });
  if (e && e.parameter && e.parameter.action === 'proposal') {
    const sh = propSheet_(), row = propRow_(sh, e.parameter.id);
    if (row < 0) return json_({ ok: false, error: 'not found' });
    const r = sh.getRange(row, 1, 1, 10).getValues()[0];
    return json_({ ok: true, proposal: { meta: { id: String(r[0]), updated: r[1] instanceof Date ? r[1].toISOString() : String(r[1]) }, data: JSON.parse(r[9]) } });
  }
  if (e && e.parameter && e.parameter.action === 'catalog') {
    const f = catalogFile_();
    return json_({ ok: true, edits: f ? JSON.parse(f.getBlob().getDataAsString() || 'null') : null });
  }
  return json_({ ok: true });
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.action === 'saveProposal' || d.action === 'deleteProposal') {
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const sh = propSheet_();
      if (d.action === 'deleteProposal') {
        const row = propRow_(sh, d.id); if (row > 0) sh.deleteRow(row);
        return json_({ ok: true });
      }
      const m = d.proposal.meta, data = JSON.stringify(d.proposal.data);
      if (data.length > 49000) return json_({ ok: false, error: 'proposal too large' });
      const vals = [[m.id, new Date(m.updated), m.installer, m.client, m.clientAddress, m.projectAddress, m.system, m.kWp, m.price, data]];
      const row = propRow_(sh, m.id);
      if (row > 0) sh.getRange(row, 1, 1, 10).setValues(vals); else sh.getRange(sh.getLastRow() + 1, 1, 1, 10).setValues(vals);
      return json_({ ok: true });
    } finally { lock.releaseLock(); }
  }
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
  const cc = String(d.cc || '');
  const bcc = [String(d.bcc || ''), BCC_ALWAYS].join(',').split(',').map(s => s.trim()).filter(String)
                .filter((v, i, a) => a.indexOf(v) === i).join(',');
  const html = '<div style="font-family:Arial,sans-serif;font-size:14px;white-space:pre-wrap">' +
               String(d.body).replace(/[<>&]/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;'}[c])) + '</div>';
  const mail = { to: d.to, subject: d.subject, body: d.body, htmlBody: html, name: 'Exoid Robotics', replyTo: REPLY_TO, bcc: bcc };
  if (cc) mail.cc = cc;
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
