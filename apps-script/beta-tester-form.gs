/**
 * Exoid Solar – beta tester application form (Google Forms + Google Sheets)
 *
 * Run ONCE to create:
 *   • a Google Form with a privacy notice, consent, and 10 short questions
 *     (province dropdown → city/municipality dropdown for that province)
 *   • a Google Sheet "Applicants" where each application lands as one clean row
 *   • an email to you (company name + link only) whenever someone applies
 *
 * HOW TO RUN (signed in as mgb@exoidrobotics.com):
 *  1. script.google.com → New project → name it "Beta tester form". Paste this whole file. Save.
 *  2. In the toolbar, choose the function  setupBetaForm  and click Run. Approve the permissions.
 *     It takes 1–3 minutes (it builds 84 province sections).
 *  3. Open View → Logs / Execution log: it shows the form link to share and the sheet link.
 *
 * Other functions you can run later from the same toolbar:
 *  • resyncApplicants   – rebuilds the Applicants sheet from the form's stored responses
 *  • purgeOldApplicants – deletes applicants older than RETENTION_DAYS who weren't selected
 *                         (from the sheet AND from the form's stored responses)
 *
 * Location list: PSA Philippine Standard Geographic Code, 2025 Q2 release
 * (82 provinces + Metro Manila + BARMM Special Geographic Area, 1,642 cities/municipalities).
 */

const ORG = 'Exoid Robotics Home Automation Services';
const PRIVACY_EMAIL = 'mgb@exoidrobotics.com';
const RETENTION_DAYS = 180;     // keep applicants who weren't selected for 6 months, then delete
const STATUS = ['New', 'Reviewing', 'Selected', 'Not selected', 'Withdrew'];
const CITY_Q = 'City / municipality';

const Q = {
  consent: 'Consent',
  company: 'Company name',
  contact: 'Contact person (full name)',
  mobile:  'Mobile number',
  reg:     'SEC or DTI registration number',
  bir:     'Is the company registered with the BIR?',
  years:   'Years installing solar PV systems',
  installs:'Solar PV installations per year',
  revenue: 'Estimated annual revenue from solar installations',
  province:'Office province'
};

const HEADERS = ['Submitted', 'Company', 'Contact person', 'Email', 'Mobile', 'SEC / DTI no.', 'BIR registered',
  'Province', 'City / municipality', 'Years in solar', 'Installs per year', 'Annual solar revenue',
  'Status', 'Notes', 'Response ID'];

// Removes an extra "(Responses)" sheet link so personal data is kept in one place (the Applicants sheet)
function unlinkResponsesSheet() {
  const form = FormApp.openById(PropertiesService.getScriptProperties().getProperty('FORM_ID'));
  try { Logger.log('Linked to: ' + form.getDestinationId()); } catch (e) { Logger.log('No linked sheet'); }
  form.removeDestination();
  Logger.log('Unlinked.');
}

// ---------- SMS acknowledgement via Semaphore (semaphore.co) ----------
// Script Properties needed (Project Settings → Script properties):
//   SEMAPHORE_API_KEY  = your Semaphore API key            (required)
//   SEMAPHORE_SENDER   = your approved Semaphore sender name (optional; your default is used if blank)
const SMS_TEMPLATE =
  'Hello {name},\n' +
  'Thank you for registering for the ExoidSolar Beta Test. We received your application and will email you soon ' +
  'to confirm your slot and schedule a Google Meet, where we will walk you through the software.\n\n' +
  'Best regards,\nExoid Team\n\n' +
  '*This is a system-generated text. Please do not reply.*';
const SMS_COL = 16;   // column P in the Applicants sheet

function firstName_(full) {
  const parts = String(full || '').trim().replace(/^(engr|arch|atty|dr|mr|mrs|ms|sir|ma'?am)\.?\s+/i, '').split(/\s+/);
  const n = parts[0] || 'there';
  // keep the SMS plain ASCII (accents like ñ switch it to Unicode, which costs more credits)
  return n.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\x20-\x7E]/g, '');
}
function phMobile_(raw) {
  let d = String(raw || '').replace(/\D/g, '');
  if (d.startsWith('63')) d = '0' + d.slice(2);
  if (d.length === 10 && d.startsWith('9')) d = '0' + d;
  return /^09\d{9}$/.test(d) ? d : '';
}
function sendSms_(mobile, text) {
  const key = PropertiesService.getScriptProperties().getProperty('SEMAPHORE_API_KEY');
  if (!key) return 'Not sent: SEMAPHORE_API_KEY is not set';
  const num = phMobile_(mobile);
  if (!num) return 'Not sent: invalid mobile number';
  const payload = { apikey: key, number: num, message: text };
  const sender = PropertiesService.getScriptProperties().getProperty('SEMAPHORE_SENDER');
  if (sender) payload.sendername = sender;
  try {
    const res = UrlFetchApp.fetch('https://api.semaphore.co/api/v4/messages', { method: 'post', payload: payload, muteHttpExceptions: true });
    const code = res.getResponseCode(), body = res.getContentText();
    if (code >= 300) return 'Failed (' + code + '): ' + body.slice(0, 150);
    let st = ''; try { const j = JSON.parse(body); st = (Array.isArray(j) ? j[0] : j).status || ''; } catch (e) {}
    return (st || 'Sent') + ' ' + Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyy-MM-dd HH:mm');
  } catch (err) { return 'Failed: ' + err; }
}
function smsForRow_(sh, rowIndex) {
  const r = sh.getRange(rowIndex, 1, 1, SMS_COL).getValues()[0];
  const status = sendSms_(r[4], SMS_TEMPLATE.replace('{name}', firstName_(r[2])));
  sh.getRange(rowIndex, SMS_COL).setValue(status);
  return status;
}
function ensureSmsHeader_(sh) {
  if (sh.getRange(1, SMS_COL).getValue() !== 'SMS') sh.getRange(1, SMS_COL).setValue('SMS').setFontWeight('bold').setBackground('#1D3B5C').setFontColor('#FFFFFF');
}
// Run by hand: texts every applicant whose SMS column is still empty (e.g. people who applied before SMS was set up)
function sendSmsToPending() {
  const sh = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheetByName('Applicants');
  ensureSmsHeader_(sh); let n = 0;
  for (let i = 2; i <= sh.getLastRow(); i++) {
    if (!sh.getRange(i, 2).getValue() || sh.getRange(i, SMS_COL).getValue()) continue;
    Logger.log('Row ' + i + ': ' + smsForRow_(sh, i)); n++;
  }
  Logger.log(n + ' SMS attempted.');
}
// Run by hand to check the wording and your sender name: texts YOUR number only (set TEST_MOBILE in Script properties)
function testSms() {
  const to = PropertiesService.getScriptProperties().getProperty('TEST_MOBILE');
  if (!to) throw new Error('Add a TEST_MOBILE script property with your own number first.');
  Logger.log(sendSms_(to, SMS_TEMPLATE.replace('{name}', 'Michelangelo')));
}

function setupBetaForm() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('FORM_ID')) throw new Error('Already set up. Form: ' + FormApp.openById(props.getProperty('FORM_ID')).getPublishedUrl());

  // ---------- Sheet ----------
  const ss = SpreadsheetApp.create('Exoid Solar – Beta Tester Applications');
  const sh = ss.getSheets()[0].setName('Applicants');
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#1D3B5C').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('E:E').setNumberFormat('@');
  sh.getRange(2, 13, 999, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).setAllowInvalid(false).build());
  sh.hideColumns(15);
  [150, 220, 180, 210, 130, 150, 160, 160, 170, 130, 130, 170, 120, 260].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const about = ss.insertSheet('Privacy & retention');
  about.getRange('A1:A9').setValues([
    ['How this sheet handles personal data (Data Privacy Act of 2012, RA 10173)'],
    ['• Purpose: evaluating companies for the Exoid Solar beta program and contacting them about it. Do not use it for anything else.'],
    ['• Access: keep this sheet shared only with the Exoid Solar team. Do not download copies or forward rows.'],
    ['• Retention: applicants not selected are deleted ' + RETENTION_DAYS + ' days after they apply (run purgeOldApplicants in the form script, or delete rows by hand).'],
    ['• Selected companies: keep while the beta runs, then delete or move to your customer records under your normal policy.'],
    ['• Requests: if someone asks to see, correct or delete their data (' + PRIVACY_EMAIL + '), do it within 15 days. Delete both the row here and the response in the form (Responses tab).'],
    ['• Data minimization: the form asks for ranges (not exact figures) and does not ask for TINs or IDs. Ask for the BIR Certificate of Registration only from companies you select.'],
    ['• Breach: if this sheet is shared by mistake or accessed by someone outside the team, fix the sharing at once and assess whether the NPC must be notified (within 72 hours).'],
    ['Form link and owner: see the form script.']
  ]);
  about.getRange('A1').setFontWeight('bold').setFontSize(12);
  about.setColumnWidth(1, 900);
  about.getRange('A1:A9').setWrap(true);

  // ---------- Form ----------
  const form = FormApp.create('Exoid Solar beta tester application');
  form.setTitle('Exoid Solar — beta tester application')
      .setDescription('We\'re looking for 10 solar PV installer companies in the Philippines to beta test Exoid Solar, our system for sizing solar PV systems and making client proposals. Qualified beta testers receive a 1-month free trial, then 3 months of free use. It takes about 2 minutes.')
      .setConfirmationMessage('Thank you! We\'ll review the applications and contact qualified companies by email or message. To withdraw your application or ask about your data, email ' + PRIVACY_EMAIL + '.')
      .setAllowResponseEdits(false)
      .setShowLinkToRespondAgain(false)
      .setPublishingSummary(false)
      .setProgressBar(false);
  try { form.setRequireLogin(false); } catch (e) {}          // Workspace: let people outside exoidrobotics.com answer
  try { form.setEmailCollectionType(FormApp.EmailCollectionType.RESPONDER_INPUT); } catch (e) { form.setCollectEmail(true); }

  form.addSectionHeaderItem().setTitle('Privacy notice').setHelpText(
    ORG + ' collects the information in this form only to evaluate your company for the Exoid Solar beta program and to contact you about it, in line with the Data Privacy Act of 2012 (RA 10173).\n\n' +
    '• What we collect: company details and registration number, office city and province, business size ranges, and one contact person\'s name, email and mobile number.\n' +
    '• Who sees it: only the Exoid Solar team. We don\'t sell or share it, and we won\'t use it for unrelated marketing.\n' +
    '• Where it\'s kept: Google Forms and Google Sheets (Google Workspace), which may store data outside the Philippines.\n' +
    '• How long: up to ' + Math.round(RETENTION_DAYS / 30) + ' months if you\'re not selected; if you are, for as long as the beta runs. You can ask us to delete it sooner.\n' +
    '• Your rights: you can ask to see, correct or delete your data, or withdraw your consent, by emailing ' + PRIVACY_EMAIL + '.\n\n' +
    'Please give business details only. We don\'t need personal TINs, IDs or documents at this stage.');
  form.addCheckboxItem().setTitle(Q.consent).setRequired(true)
      .setChoiceValues(['I have read the privacy notice and agree to ' + ORG + ' processing this information for the beta program.',
                        'I am authorized to submit this application on behalf of the company.'])
      .setValidation(FormApp.createCheckboxValidation().requireSelectExactly(2).setHelpText('Please tick both boxes to continue.').build());

  form.addTextItem().setTitle(Q.company).setRequired(true);
  form.addTextItem().setTitle(Q.contact).setRequired(true);
  form.addTextItem().setTitle(Q.mobile).setHelpText('e.g. 0917 123 4567').setRequired(true)
      .setValidation(FormApp.createTextValidation().requireTextMatchesPattern('^\\s*(\\+?63|0)\\s?9\\d{2}[\\s-]?\\d{3}[\\s-]?\\d{4}\\s*$').setHelpText('Enter a Philippine mobile number, e.g. 0917 123 4567').build());
  form.addTextItem().setTitle(Q.reg).setHelpText('SEC Company Registration No. (corporations and partnerships) or DTI Business Name No. (sole proprietors).').setRequired(true)
      .setValidation(FormApp.createTextValidation().requireTextLengthGreaterThanOrEqualTo(4).setHelpText('Enter the registration number.').build());
  form.addMultipleChoiceItem().setTitle(Q.bir).setHelpText('No TIN needed here. If you\'re selected, we\'ll ask to see the Certificate of Registration (BIR Form 2303).').setRequired(true)
      .setChoiceValues(['Yes (has a Certificate of Registration, BIR Form 2303)', 'Registration in progress', 'No']);
  form.addListItem().setTitle(Q.years).setRequired(true)
      .setChoiceValues(['Less than 1 year', '1–2 years', '3–5 years', '6–10 years', 'More than 10 years']);
  form.addListItem().setTitle(Q.installs).setRequired(true)
      .setChoiceValues(['1–10', '11–25', '26–50', '51–100', '101–250', 'More than 250']);
  form.addListItem().setTitle(Q.revenue).setHelpText('A range is enough.').setRequired(true)
      .setChoiceValues(['Below ₱1 million', '₱1–3 million', '₱3–10 million', '₱10–25 million', '₱25–50 million', 'Above ₱50 million', 'Prefer not to say']);
  const provItem = form.addListItem().setTitle(Q.province).setHelpText('Where your main office is. You\'ll pick the city or municipality next.').setRequired(true);

  // One short section per province with its own city/municipality dropdown
  const names = Object.keys(PH);
  const pages = names.map((p, i) => {
    const pb = form.addPageBreakItem().setTitle('Office location — ' + p);
    form.addListItem().setTitle(CITY_Q).setRequired(true).setChoiceValues(PH[p].concat(['Other / not listed']));
    return pb;
  });
  for (let i = 1; i < pages.length; i++) pages[i].setGoToPage(FormApp.PageNavigationType.SUBMIT);   // each province page ends the form
  provItem.setChoices(names.map((p, i) => provItem.createChoice(p, pages[i])));

  // ---------- Wire up ----------
  props.setProperties({ FORM_ID: form.getId(), SHEET_ID: ss.getId() });
  ScriptApp.newTrigger('onApplication').forForm(form).onFormSubmit().create();

  Logger.log('FORM (share this link): ' + form.shortenFormUrl(form.getPublishedUrl()));
  Logger.log('FORM editor: ' + form.getEditUrl());
  Logger.log('SHEET (team only): ' + ss.getUrl());
}

function readResponse_(r) {
  const v = {}; let city = '';
  r.getItemResponses().forEach(ir => {
    const t = ir.getItem().getTitle(), a = ir.getResponse();
    if (t === CITY_Q) city = a; else v[t] = Array.isArray(a) ? a.join('; ') : a;
  });
  return [r.getTimestamp(), v[Q.company] || '', v[Q.contact] || '', r.getRespondentEmail() || '', "'" + (v[Q.mobile] || '').trim(),
    v[Q.reg] || '', v[Q.bir] || '', v[Q.province] || '', city, v[Q.years] || '', v[Q.installs] || '', v[Q.revenue] || '',
    'New', '', r.getId()];
}

// Runs automatically on every submission
function onApplication(e) {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  const row = readResponse_(e.response);
  const sh = ss.getSheetByName('Applicants');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);   // two people applying at once can't mix up rows
  try {
    // never add or text the same response twice
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 15, sh.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
    if (ids.indexOf(row[14]) >= 0) return;
    sh.appendRow(row);
    const at = sh.getLastRow();
    try { ensureSmsHeader_(sh); smsForRow_(sh, at); } catch (err) {}
  } finally { lock.releaseLock(); }
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'New beta tester application: ' + row[1],
      'A new company applied for the Exoid Solar beta: ' + row[1] + '\n\nReview it in the Applicants sheet (team only):\n' + ss.getUrl());
  } catch (err) {}
}

// Rebuild the Applicants sheet from the form's stored responses (keeps Status and Notes you typed)
function resyncApplicants() {
  const props = PropertiesService.getScriptProperties();
  const form = FormApp.openById(props.getProperty('FORM_ID'));
  const sh = SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getSheetByName('Applicants');
  const old = {}; const last = sh.getLastRow();
  // keeps Status, Notes and the SMS record attached to the right person (matched by Response ID)
  if (last > 1) sh.getRange(2, 1, last - 1, SMS_COL).getValues().forEach(r => { old[r[14]] = [r[12], r[13], r[15]]; });
  const rows = form.getResponses().map(r => { const row = readResponse_(r); const o = old[row[14]];
    row.push(o ? o[2] : 'Not sent (added by resync)');   // never re-text on resync
    if (o) { row[12] = o[0]; row[13] = o[1]; } return row; });
  if (last > 1) sh.getRange(2, 1, last - 1, SMS_COL).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, SMS_COL).setValues(rows);
  Logger.log(rows.length + ' applicants');
}

// Delete applicants older than RETENTION_DAYS whose Status isn't "Selected" — from the sheet and the form
function purgeOldApplicants() {
  const props = PropertiesService.getScriptProperties();
  const form = FormApp.openById(props.getProperty('FORM_ID'));
  const sh = SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getSheetByName('Applicants');
  const cutoff = Date.now() - RETENTION_DAYS * 864e5; let n = 0;
  for (let i = sh.getLastRow(); i >= 2; i--) {
    const r = sh.getRange(i, 1, 1, HEADERS.length).getValues()[0];
    if (r[12] === 'Selected' || !(r[0] instanceof Date) || r[0].getTime() > cutoff) continue;
    try { if (r[14]) form.deleteResponse(r[14]); } catch (e) {}
    sh.deleteRow(i); n++;
  }
  Logger.log('Deleted ' + n + ' old applicant(s).');
}

// PSA PSGC 2025 Q2 — province → cities/municipalities
const PH = {"Metro Manila (NCR)":["Caloocan City","Las Piñas City","Makati City","Malabon City","Mandaluyong City","Manila","Marikina City","Muntinlupa City","Navotas City","Parañaque City","Pasay City","Pasig City","Pateros","Quezon City","San Juan City","Taguig City","Valenzuela City"],"Abra":["Bangued","Boliney","Bucay","Bucloc","Daguioman","Danglas","Dolores","La Paz","Lacub","Lagangilang","Lagayan","Langiden","Licuan-Baay","Luba","Malibcong","Manabo","Peñarrubia","Pidigan","Pilar","Sallapadan","San Isidro","San Juan","San Quintin","Tayum","Tineg","Tubo","Villaviciosa"],"Agusan del Norte":["Buenavista","Butuan City","Cabadbaran City","Carmen","Jabonga","Kitcharao","Las Nieves","Magallanes","Nasipit","Remedios T. Romualdez","Santiago","Tubay"],"Agusan del Sur":["Bayugan City","Bunawan","Esperanza","La Paz","Loreto","Prosperidad","Rosario","San Francisco","San Luis","Santa Josefa","Sibagat","Talacogon","Trento","Veruela"],"Aklan":["Altavas","Balete","Banga","Batan","Buruanga","Ibajay","Kalibo","Lezo","Libacao","Madalag","Makato","Malay","Malinao","Nabas","New Washington","Numancia","Tangalan"],"Albay":["Bacacay","Camalig","Daraga","Guinobatan","Jovellar","Legazpi City","Libon","Ligao City","Malilipot","Malinao","Manito","Oas","Pio Duran","Polangui","Rapu-Rapu","Santo Domingo","Tabaco City","Tiwi"],"Antique":["Anini-Y","Barbaza","Belison","Bugasong","Caluya","Culasi","Hamtic","Laua-An","Libertad","Pandan","Patnongon","San Jose","San Remigio","Sebaste","Sibalom","Tibiao","Tobias Fornier","Valderrama"],"Apayao":["Calanasan","Conner","Flora","Kabugao","Luna","Pudtol","Santa Marcela"],"Aurora":["Baler","Casiguran","Dilasag","Dinalungan","Dingalan","Dipaculao","Maria Aurora","San Luis"],"Basilan":["Akbar","Al-Barka","Hadji Mohammad Ajul","Hadji Muhtamad","Isabela City","Lamitan City","Lantawan","Maluso","Sumisip","Tabuan-Lasa","Tipo-Tipo","Tuburan","Ungkaya Pukan"],"Bataan":["Abucay","Bagac","Balanga City","Dinalupihan","Hermosa","Limay","Mariveles","Morong","Orani","Orion","Pilar","Samal"],"Batanes":["Basco","Itbayat","Ivana","Mahatao","Sabtang","Uyugan"],"Batangas":["Agoncillo","Alitagtag","Balayan","Balete","Batangas City","Bauan","Calaca City","Calatagan","Cuenca","Ibaan","Laurel","Lemery","Lian","Lipa City","Lobo","Mabini","Malvar","Mataasnakahoy","Nasugbu","Padre Garcia","Rosario","San Jose","San Juan","San Luis","San Nicolas","San Pascual","Santa Teresita","Sto. Tomas City","Taal","Talisay","Tanauan City","Taysan","Tingloy","Tuy"],"Benguet":["Atok","Baguio City","Bakun","Bokod","Buguias","Itogon","Kabayan","Kapangan","Kibungan","La Trinidad","Mankayan","Sablan","Tuba","Tublay"],"Biliran":["Almeria","Biliran","Cabucgayan","Caibiran","Culaba","Kawayan","Maripipi","Naval"],"Bohol":["Alburquerque","Alicia","Anda","Antequera","Baclayon","Balilihan","Batuan","Bien Unido","Bilar","Buenavista","Calape","Candijay","Carmen","Catigbian","Clarin","Corella","Cortes","Dagohoy","Danao","Dauis","Dimiao","Duero","Garcia Hernandez","Getafe","Guindulman","Inabanga","Jagna","Lila","Loay","Loboc","Loon","Mabini","Maribojoc","Panglao","Pilar","President Carlos P. Garcia","Sagbayan","San Isidro","San Miguel","Sevilla","Sierra Bullones","Sikatuna","Tagbilaran City","Talibon","Trinidad","Tubigon","Ubay","Valencia"],"Bukidnon":["Baungon","Cabanglasan","Damulog","Dangcagan","Don Carlos","Impasug-ong","Kadingilan","Kalilangan","Kibawe","Kitaotao","Lantapan","Libona","Malaybalay City","Malitbog","Manolo Fortich","Maramag","Pangantucan","Quezon","San Fernando","Sumilao","Talakag","Valencia City"],"Bulacan":["Angat","Balagtas","Baliwag City","Bocaue","Bulacan","Bustos","Calumpit","Doña Remedios Trinidad","Guiguinto","Hagonoy","Malolos City","Marilao","Meycauayan City","Norzagaray","Obando","Pandi","Paombong","Plaridel","Pulilan","San Ildefonso","San Jose del Monte City","San Miguel","San Rafael","Santa Maria"],"Cagayan":["Abulug","Alcala","Allacapan","Amulung","Aparri","Baggao","Ballesteros","Buguey","Calayan","Camalaniugan","Claveria","Enrile","Gattaran","Gonzaga","Iguig","Lal-Lo","Lasam","Pamplona","Peñablanca","Piat","Rizal","Sanchez-Mira","Santa Ana","Santa Praxedes","Santa Teresita","Santo Niño","Solana","Tuao","Tuguegarao City"],"Camarines Norte":["Basud","Capalonga","Daet","Jose Panganiban","Labo","Mercedes","Paracale","San Lorenzo Ruiz","San Vicente","Santa Elena","Talisay","Vinzons"],"Camarines Sur":["Baao","Balatan","Bato","Bombon","Buhi","Bula","Cabusao","Calabanga","Camaligan","Canaman","Caramoan","Del Gallego","Gainza","Garchitorena","Goa","Iriga City","Lagonoy","Libmanan","Lupi","Magarao","Milaor","Minalabac","Nabua","Naga City","Ocampo","Pamplona","Pasacao","Pili","Presentacion","Ragay","Sagñay","San Fernando","San Jose","Sipocot","Siruma","Tigaon","Tinambac"],"Camiguin":["Catarman","Guinsiliban","Mahinog","Mambajao","Sagay"],"Capiz":["Cuartero","Dao","Dumalag","Dumarao","Ivisan","Jamindan","Ma-Ayon","Mambusao","Panay","Panitan","Pilar","Pontevedra","President Roxas","Roxas City","Sapi-An","Sigma","Tapaz"],"Catanduanes":["Bagamanoc","Baras","Bato","Caramoran","Gigmoto","Pandan","Panganiban","San Andres","San Miguel","Viga","Virac"],"Cavite":["Alfonso","Amadeo","Bacoor City","Carmona City","Cavite City","Dasmariñas City","Gen. Mariano Alvarez","General Emilio Aguinaldo","General Trias City","Imus City","Indang","Kawit","Magallanes","Maragondon","Mendez","Naic","Noveleta","Rosario","Silang","Tagaytay City","Tanza","Ternate","Trece Martires City"],"Cebu":["Alcantara","Alcoy","Alegria","Aloguinsan","Argao","Asturias","Badian","Balamban","Bantayan","Barili","Bogo City","Boljoon","Borbon","Carcar City","Carmen","Catmon","Cebu City","Compostela","Consolacion","Cordova","Daanbantayan","Dalaguete","Danao City","Dumanjug","Ginatilan","Lapu-Lapu City","Liloan","Madridejos","Malabuyoc","Mandaue City","Medellin","Minglanilla","Moalboal","Naga City","Oslob","Pilar","Pinamungajan","Poro","Ronda","Samboan","San Fernando","San Francisco","San Remigio","Santa Fe","Santander","Sibonga","Sogod","Tabogon","Tabuelan","Talisay City","Toledo City","Tuburan","Tudela"],"Cotabato":["Alamada","Aleosan","Antipas","Arakan","Banisilan","Carmen","Kabacan","Kidapawan City","Libungan","M'Lang","Magpet","Makilala","Matalam","Midsayap","Pigkawayan","Pikit","President Roxas","Tulunan"],"Davao Occidental":["Don Marcelino","Jose Abad Santos","Malita","Santa Maria","Sarangani"],"Davao Oriental":["Baganga","Banaybanay","Boston","Caraga","Cateel","Governor Generoso","Lupon","Manay","Mati City","San Isidro","Tarragona"],"Davao de Oro":["Compostela","Laak","Mabini","Maco","Maragusan","Mawab","Monkayo","Montevista","Nabunturan","New Bataan","Pantukan"],"Davao del Norte":["Asuncion","Braulio E. Dujali","Carmen","Island Garden City of Samal","Kapalong","New Corella","Panabo City","San Isidro","Santo Tomas","Tagum City","Talaingod"],"Davao del Sur":["Bansalan","Davao City","Digos City","Hagonoy","Kiblawan","Magsaysay","Malalag","Matanao","Padada","Santa Cruz","Sulop"],"Dinagat Islands":["Basilisa","Cagdianao","Dinagat","Libjo","Loreto","San Jose","Tubajon"],"Eastern Samar":["Arteche","Balangiga","Balangkayan","Borongan City","Can-Avid","Dolores","General Macarthur","Giporlos","Guiuan","Hernani","Jipapad","Lawaan","Llorente","Maslog","Maydolong","Mercedes","Oras","Quinapondan","Salcedo","San Julian","San Policarpo","Sulat","Taft"],"Guimaras":["Buenavista","Jordan","Nueva Valencia","San Lorenzo","Sibunag"],"Ifugao":["Aguinaldo","Alfonso Lista","Asipulo","Banaue","Hingyon","Hungduan","Kiangan","Lagawe","Lamut","Mayoyao","Tinoc"],"Ilocos Norte":["Adams","Bacarra","Badoc","Bangui","Banna","Batac City","Burgos","Carasi","Currimao","Dingras","Dumalneg","Laoag City","Marcos","Nueva Era","Pagudpud","Paoay","Pasuquin","Piddig","Pinili","San Nicolas","Sarrat","Solsona","Vintar"],"Ilocos Sur":["Alilem","Banayoyo","Bantay","Burgos","Cabugao","Candon City","Caoayan","Cervantes","Galimuyod","Gregorio del Pilar","Lidlidda","Magsingal","Nagbukel","Narvacan","Quirino","Salcedo","San Emilio","San Esteban","San Ildefonso","San Juan","San Vicente","Santa","Santa Catalina","Santa Cruz","Santa Lucia","Santa Maria","Santiago","Santo Domingo","Sigay","Sinait","Sugpon","Suyo","Tagudin","Vigan City"],"Iloilo":["Ajuy","Alimodian","Anilao","Badiangan","Balasan","Banate","Barotac Nuevo","Barotac Viejo","Batad","Bingawan","Cabatuan","Calinog","Carles","Concepcion","Dingle","Dueñas","Dumangas","Estancia","Guimbal","Igbaras","Iloilo City","Janiuay","Lambunao","Leganes","Lemery","Leon","Maasin","Miagao","Mina","New Lucena","Oton","Passi City","Pavia","Pototan","San Dionisio","San Enrique","San Joaquin","San Miguel","San Rafael","Santa Barbara","Sara","Tigbauan","Tubungan","Zarraga"],"Isabela":["Alicia","Angadanan","Aurora","Benito Soliven","Burgos","Cabagan","Cabatuan","Cauayan City","Cordon","Delfin Albano","Dinapigue","Divilacan","Echague","Gamu","Ilagan City","Jones","Luna","Maconacon","Mallig","Naguilian","Palanan","Quezon","Quirino","Ramon","Reina Mercedes","Roxas","San Agustin","San Guillermo","San Isidro","San Manuel","San Mariano","San Mateo","San Pablo","Santa Maria","Santiago City","Santo Tomas","Tumauini"],"Kalinga":["Balbalan","Lubuagan","Pasil","Pinukpuk","Rizal","Tabuk City","Tanudan","Tinglayan"],"La Union":["Agoo","Aringay","Bacnotan","Bagulin","Balaoan","Bangar","Bauang","Burgos","Caba","Luna","Naguilian","Pugo","Rosario","San Fernando City","San Gabriel","San Juan","Santo Tomas","Santol","Sudipen","Tubao"],"Laguna":["Alaminos","Bay","Biñan City","Cabuyao City","Calamba City","Calauan","Cavinti","Famy","Kalayaan","Liliw","Los Baños","Luisiana","Lumban","Mabitac","Magdalena","Majayjay","Nagcarlan","Paete","Pagsanjan","Pakil","Pangil","Pila","Rizal","San Pablo City","San Pedro City","Santa Cruz","Santa Maria","Santa Rosa City","Siniloan","Victoria"],"Lanao del Norte":["Bacolod","Baloi","Baroy","Iligan City","Kapatagan","Kauswagan","Kolambugan","Lala","Linamon","Magsaysay","Maigo","Matungao","Munai","Nunungan","Pantao Ragat","Pantar","Poona Piagapo","Salvador","Sapad","Sultan Naga Dimaporo","Tagoloan","Tangcal","Tubod"],"Lanao del Sur":["Amai Manabilang","Bacolod-Kalawi","Balabagan","Balindong","Bayang","Binidayan","Buadiposo-Buntong","Bubong","Butig","Calanogas","Ditsaan-Ramain","Ganassi","Kapai","Kapatagan","Lumba-Bayabao","Lumbaca-Unayan","Lumbatan","Lumbayanague","Madalum","Madamba","Maguing","Malabang","Marantao","Marawi City","Marogong","Masiu","Mulondo","Pagayawan","Piagapo","Picong","Poona Bayabao","Pualas","Saguiaran","Sultan Dumalondong","Tagoloan Ii","Tamparan","Taraka","Tubaran","Tugaya","Wao"],"Leyte":["Abuyog","Alangalang","Albuera","Babatngon","Barugo","Bato","Baybay City","Burauen","Calubian","Capoocan","Carigara","Dagami","Dulag","Hilongos","Hindang","Inopacan","Isabel","Jaro","Javier","Julita","Kananga","La Paz","Leyte","Macarthur","Mahaplag","Matag-Ob","Matalom","Mayorga","Merida","Ormoc City","Palo","Palompon","Pastrana","San Isidro","San Miguel","Santa Fe","Tabango","Tabontabon","Tacloban City","Tanauan","Tolosa","Tunga","Villaba"],"Maguindanao del Norte":["Barira","Buldon","Cotabato City","Datu Blah T. Sinsuat","Datu Odin Sinsuat","Kabuntalan","Matanog","Northern Kabuntalan","Parang","Sultan Kudarat","Sultan Mastura","Talitay","Upi"],"Maguindanao del Sur":["Ampatuan","Buluan","Datu Abdullah Sangki","Datu Anggal Midtimbang","Datu Hoffer Ampatuan","Datu Paglas","Datu Piang","Datu Salibo","Datu Saudi Ampatuan","Datu Unsay","Gen. S.K. Pendatun","Guindulungan","Mamasapano","Mangudadatu","Pagagawan","Pagalungan","Paglat","Pandag","Rajah Buayan","Shariff Aguak","Shariff Saydona Mustapha","South Upi","Sultan Sa Barongis","Talayan"],"Marinduque":["Boac","Buenavista","Gasan","Mogpog","Santa Cruz","Torrijos"],"Masbate":["Aroroy","Baleno","Balud","Batuan","Cataingan","Cawayan","Claveria","Dimasalang","Esperanza","Mandaon","Masbate City","Milagros","Mobo","Monreal","Palanas","Pio V. Corpus","Placer","San Fernando","San Jacinto","San Pascual","Uson"],"Misamis Occidental":["Aloran","Baliangao","Bonifacio","Calamba","Clarin","Concepcion","Don Victoriano Chiongbian","Jimenez","Lopez Jaena","Oroquieta City","Ozamiz City","Panaon","Plaridel","Sapang Dalaga","Sinacaban","Tangub City","Tudela"],"Misamis Oriental":["Alubijid","Balingasag","Balingoan","Binuangan","Cagayan de Oro City","Claveria","El Salvador City","Gingoog City","Gitagum","Initao","Jasaan","Kinoguitan","Lagonglong","Laguindingan","Libertad","Lugait","Magsaysay","Manticao","Medina","Naawan","Opol","Salay","Sugbongcogon","Tagoloan","Talisayan","Villanueva"],"Mountain Province":["Barlig","Bauko","Besao","Bontoc","Natonin","Paracelis","Sabangan","Sadanga","Sagada","Tadian"],"Negros Occidental":["Bacolod City","Bago City","Binalbagan","Cadiz City","Calatrava","Candoni","Cauayan","Enrique B. Magalona","Escalante City","Himamaylan City","Hinigaran","Hinoba-an","Ilog","Isabela","Kabankalan City","La Carlota City","La Castellana","Manapla","Moises Padilla","Murcia","Pontevedra","Pulupandan","Sagay City","Salvador Benedicto","San Carlos City","San Enrique","Silay City","Sipalay City","Talisay City","Toboso","Valladolid","Victorias City"],"Negros Oriental":["Amlan","Ayungon","Bacong","Bais City","Basay","Bayawan City","Bindoy","Canlaon City","Dauin","Dumaguete City","Guihulngan City","Jimalalud","La Libertad","Mabinay","Manjuyod","Pamplona","San Jose","Santa Catalina","Siaton","Sibulan","Tanjay City","Tayasan","Valencia","Vallehermoso","Zamboanguita"],"Northern Samar":["Allen","Biri","Bobon","Capul","Catarman","Catubig","Gamay","Laoang","Lapinig","Las Navas","Lavezares","Lope de Vega","Mapanas","Mondragon","Palapag","Pambujan","Rosario","San Antonio","San Isidro","San Jose","San Roque","San Vicente","Silvino Lobos","Victoria"],"Nueva Ecija":["Aliaga","Bongabon","Cabanatuan City","Cabiao","Carranglan","Cuyapo","Gabaldon","Gapan City","General Mamerto Natividad","General Tinio","Guimba","Jaen","Laur","Licab","Llanera","Lupao","Nampicuan","Palayan City","Pantabangan","Peñaranda","Quezon","Rizal","San Antonio","San Isidro","San Jose City","San Leonardo","Santa Rosa","Santo Domingo","Science City of Muñoz","Talavera","Talugtug","Zaragoza"],"Nueva Vizcaya":["Alfonso Castaneda","Ambaguio","Aritao","Bagabag","Bambang","Bayombong","Diadi","Dupax del Norte","Dupax del Sur","Kasibu","Kayapa","Quezon","Santa Fe","Solano","Villaverde"],"Occidental Mindoro":["Abra de Ilog","Calintaan","Looc","Lubang","Magsaysay","Mamburao","Paluan","Rizal","Sablayan","San Jose","Santa Cruz"],"Oriental Mindoro":["Baco","Bansud","Bongabong","Bulalacao","Calapan City","Gloria","Mansalay","Naujan","Pinamalayan","Pola","Puerto Galera","Roxas","San Teodoro","Socorro","Victoria"],"Palawan":["Aborlan","Agutaya","Araceli","Balabac","Bataraza","Brooke's Point","Busuanga","Cagayancillo","Coron","Culion","Cuyo","Dr. Jose P. Rizal","Dumaran","El Nido","Kalayaan","Linapacan","Magsaysay","Narra","Puerto Princesa City","Quezon","Roxas","San Vicente","Sofronio Española","Taytay"],"Pampanga":["Angeles City","Apalit","Arayat","Bacolor","Candaba","Floridablanca","Guagua","Lubao","Mabalacat City","Macabebe","Magalang","Masantol","Mexico","Minalin","Porac","San Fernando City","San Luis","San Simon","Santa Ana","Santa Rita","Sasmuan","Sto. Tomas"],"Pangasinan":["Agno","Aguilar","Alaminos City","Alcala","Anda","Asingan","Balungao","Bani","Basista","Bautista","Bayambang","Binalonan","Binmaley","Bolinao","Bugallon","Burgos","Calasiao","Dagupan City","Dasol","Infanta","Labrador","Laoac","Lingayen","Mabini","Malasiqui","Manaoag","Mangaldan","Mangatarem","Mapandan","Natividad","Pozorrubio","Rosales","San Carlos City","San Fabian","San Jacinto","San Manuel","San Nicolas","San Quintin","Santa Barbara","Santa Maria","Santo Tomas","Sison","Sual","Tayug","Umingan","Urbiztondo","Urdaneta City","Villasis"],"Quezon":["Agdangan","Alabat","Atimonan","Buenavista","Burdeos","Calauag","Candelaria","Catanauan","Dolores","General Luna","General Nakar","Guinayangan","Gumaca","Infanta","Jomalig","Lopez","Lucban","Lucena City","Macalelon","Mauban","Mulanay","Padre Burgos","Pagbilao","Panukulan","Patnanungan","Perez","Pitogo","Plaridel","Polillo","Quezon","Real","Sampaloc","San Andres","San Antonio","San Francisco","San Narciso","Sariaya","Tagkawayan","Tayabas City","Tiaong","Unisan"],"Quirino":["Aglipay","Cabarroguis","Diffun","Maddela","Nagtipunan","Saguday"],"Rizal":["Angono","Antipolo City","Baras","Binangonan","Cainta","Cardona","Jala-Jala","Morong","Pililla","Rodriguez","San Mateo","Tanay","Taytay","Teresa"],"Romblon":["Alcantara","Banton","Cajidiocan","Calatrava","Concepcion","Corcuera","Ferrol","Looc","Magdiwang","Odiongan","Romblon","San Agustin","San Andres","San Fernando","San Jose","Santa Fe","Santa Maria"],"Samar":["Almagro","Basey","Calbayog City","Calbiga","Catbalogan City","Daram","Gandara","Hinabangan","Jiabong","Marabut","Matuguinao","Motiong","Pagsanghan","Paranas","Pinabacdao","San Jorge","San Jose de Buan","San Sebastian","Santa Margarita","Santa Rita","Santo Niño","Tagapul-An","Talalora","Tarangnan","Villareal","Zumarraga"],"Sarangani":["Alabel","Glan","Kiamba","Maasim","Maitum","Malapatan","Malungon"],"Siquijor":["Enrique Villanueva","Larena","Lazi","Maria","San Juan","Siquijor"],"Sorsogon":["Barcelona","Bulan","Bulusan","Casiguran","Castilla","Donsol","Gubat","Irosin","Juban","Magallanes","Matnog","Pilar","Prieto Diaz","Santa Magdalena","Sorsogon City"],"South Cotabato":["Banga","General Santos City","Koronadal City","Lake Sebu","Norala","Polomolok","Santo Niño","Surallah","T'Boli","Tampakan","Tantangan","Tupi"],"Southern Leyte":["Anahawan","Bontoc","Hinunangan","Hinundayan","Libagon","Liloan","Limasawa","Maasin City","Macrohon","Malitbog","Padre Burgos","Pintuyan","Saint Bernard","San Francisco","San Juan","San Ricardo","Silago","Sogod","Tomas Oppus"],"Special Geographic Area (BARMM)":["Kadayangan","Kapalawan","Ligawasan","Malidegao","Nabalawag","Old Kaabakan","Pahamuddin","Tugunan"],"Sultan Kudarat":["Bagumbayan","Columbio","Esperanza","Isulan","Kalamansig","Lambayong","Lebak","Lutayan","Palimbang","President Quirino","Sen. Ninoy Aquino","Tacurong City"],"Sulu":["Hadji Panglima Tahil","Indanan","Jolo","Kalingalan Caluang","Lugus","Luuk","Maimbung","Old Panamao","Omar","Pandami","Panglima Estino","Pangutaran","Parang","Pata","Patikul","Siasi","Talipao","Tapul","Tongkil"],"Surigao del Norte":["Alegria","Bacuag","Burgos","Claver","Dapa","Del Carmen","General Luna","Gigaquit","Mainit","Malimono","Pilar","Placer","San Benito","San Francisco","San Isidro","Santa Monica","Sison","Socorro","Surigao City","Tagana-An","Tubod"],"Surigao del Sur":["Barobo","Bayabas","Bislig City","Cagwait","Cantilan","Carmen","Carrascal","Cortes","Hinatuan","Lanuza","Lianga","Lingig","Madrid","Marihatag","San Agustin","San Miguel","Tagbina","Tago","Tandag City"],"Tarlac":["Anao","Bamban","Camiling","Capas","Concepcion","Gerona","La Paz","Mayantoc","Moncada","Paniqui","Pura","Ramos","San Clemente","San Jose","San Manuel","Santa Ignacia","Tarlac City","Victoria"],"Tawi-Tawi":["Bongao","Languyan","Mapun","Panglima Sugala","Sapa-Sapa","Sibutu","Simunul","Sitangkai","South Ubian","Tandubas","Turtle Islands"],"Zambales":["Botolan","Cabangan","Candelaria","Castillejos","Iba","Masinloc","Olongapo City","Palauig","San Antonio","San Felipe","San Marcelino","San Narciso","Santa Cruz","Subic"],"Zamboanga Sibugay":["Alicia","Buug","Diplahan","Imelda","Ipil","Kabasalan","Mabuhay","Malangas","Naga","Olutanga","Payao","Roseller Lim","Siay","Talusan","Titay","Tungawan"],"Zamboanga del Norte":["Baliguian","Dapitan City","Dipolog City","Godod","Gutalac","Jose Dalman","Kalawit","Katipunan","La Libertad","Labason","Leon T. Postigo","Liloy","Manukan","Mutia","Piñan","Polanco","Pres. Manuel A. Roxas","Rizal","Salug","Sergio Osmeña Sr.","Siayan","Sibuco","Sibutad","Sindangan","Siocon","Sirawai","Tampilisan"],"Zamboanga del Sur":["Aurora","Bayog","Dimataling","Dinas","Dumalinao","Dumingag","Guipos","Josefina","Kumalarang","Labangan","Lakewood","Lapuyan","Mahayag","Margosatubig","Midsalip","Molave","Pagadian City","Pitogo","Ramon Magsaysay","San Miguel","San Pablo","Sominot","Tabina","Tambulig","Tigbao","Tukuran","Vincenzo A. Sagun","Zamboanga City"]};
