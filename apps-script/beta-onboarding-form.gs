/**
 * Exoid Solar – beta tester ONBOARDING form (Google Forms + Google Sheets)
 * For testers you have already selected. After they submit, they see the app link
 * on the confirmation page and also get it by email.
 *
 * Run ONCE to create:
 *   • a Google Form: privacy notice, name, company, email, contact number, agreement to share insights
 *   • a Google Sheet "Beta testers" where each submission lands as one row
 *   • a welcome email to the tester (with the app link) and a short note to you
 *
 * HOW TO RUN (signed in as mgb@exoidrobotics.com):
 *  1. script.google.com → New project → name it "Beta onboarding form". Paste this whole file. Save.
 *  2. Choose the function  setupOnboardingForm  in the toolbar and click Run. Approve the permissions.
 *  3. The Execution log shows the form link to send to your testers, and the sheet link (team only).
 *
 * To change the app link later: edit APP_URL below, save, and run  updateAppLink .
 */

const ORG = 'Exoid Robotics Home Automation Services';
const PRIVACY_EMAIL = 'mgb@exoidrobotics.com';
const APP_URL = 'https://solar-app.mgb-bd8.workers.dev';
const SUPPORT = 'mgb@exoidrobotics.com';

const Q = {
  consent: 'Privacy consent',
  name:    'Full name',
  company: 'Company name',
  email:   'Email address',
  phone:   'Contact number',
  insights:'Beta tester agreement'
};
const HEADERS = ['Submitted', 'Full name', 'Company', 'Email', 'Contact number', 'Agreed to share insights',
  'Welcome email', 'Status', 'Notes', 'Response ID'];
const STATUS = ['Onboarded', 'Active', 'Feedback received', 'Withdrew'];

function confirmation_() {
  return 'Welcome to the Exoid Solar beta! 🎉\n\n' +
    'Open the app here:\n' + APP_URL + '\n\n' +
    'Sign up with the same email you entered in this form and create your company account. ' +
    'We\'ve also emailed you this link so you can find it later.\n\n' +
    'Questions or problems? Email ' + SUPPORT + '.';
}

function setupOnboardingForm() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('FORM_ID')) throw new Error('Already set up. Form: ' + FormApp.openById(props.getProperty('FORM_ID')).getPublishedUrl());

  // ---------- Sheet ----------
  const ss = SpreadsheetApp.create('Exoid Solar – Beta Testers (onboarding)');
  const sh = ss.getSheets()[0].setName('Beta testers');
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#1D3B5C').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('E:E').setNumberFormat('@');
  sh.getRange(2, 8, 999, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).setAllowInvalid(false).build());
  sh.hideColumns(10);
  [150, 200, 220, 230, 140, 170, 140, 150, 280].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // ---------- Form ----------
  const form = FormApp.create('Exoid Solar beta tester onboarding');
  form.setTitle('Exoid Solar — beta tester onboarding')
      .setDescription('Congratulations on being selected as an Exoid Solar beta tester! Fill in this short form (about 1 minute). ' +
        'When you submit, you\'ll get the link to the app right away.')
      .setConfirmationMessage(confirmation_())
      .setAllowResponseEdits(false)
      .setShowLinkToRespondAgain(false)
      .setPublishingSummary(false)
      .setProgressBar(false);
  try { form.setRequireLogin(false); } catch (e) {}          // Workspace: let people outside exoidrobotics.com answer
  try { form.setCollectEmail(false); } catch (e) {}

  form.addSectionHeaderItem().setTitle('Privacy notice').setHelpText(
    ORG + ' collects the information in this form only to give you access to the Exoid Solar beta, contact you about it, ' +
    'and ask for your feedback, in line with the Data Privacy Act of 2012 (RA 10173).\n\n' +
    '• What we collect: your name, company name, email and contact number.\n' +
    '• Who sees it: only the Exoid Solar team. We don\'t sell or share it, and we won\'t use it for unrelated marketing.\n' +
    '• Where it\'s kept: Google Forms and Google Sheets (Google Workspace), which may store data outside the Philippines.\n' +
    '• How long: for as long as the beta runs, then deleted or kept under our normal customer-records policy if you become a customer.\n' +
    '• Your rights: you can ask to see, correct or delete your data, or withdraw your consent, by emailing ' + PRIVACY_EMAIL + '.');
  form.addCheckboxItem().setTitle(Q.consent).setRequired(true)
      .setChoiceValues(['I have read the privacy notice and agree to ' + ORG + ' processing this information for the beta program.']);

  form.addTextItem().setTitle(Q.name).setRequired(true)
      .setValidation(FormApp.createTextValidation().requireTextLengthGreaterThanOrEqualTo(2).setHelpText('Please enter your full name.').build());
  form.addTextItem().setTitle(Q.company).setRequired(true)
      .setValidation(FormApp.createTextValidation().requireTextLengthGreaterThanOrEqualTo(2).setHelpText('Please enter your company name.').build());
  form.addTextItem().setTitle(Q.email).setHelpText('Use this same email when you sign up in the app.').setRequired(true)
      .setValidation(FormApp.createTextValidation().requireTextIsEmail().setHelpText('Enter a valid email address.').build());
  form.addTextItem().setTitle(Q.phone).setHelpText('Mobile or landline, e.g. 0917 123 4567').setRequired(true)
      .setValidation(FormApp.createTextValidation().requireTextMatchesPattern('^\\s*\\+?[0-9][0-9\\s()\\-]{6,18}[0-9]\\s*$')
        .setHelpText('Enter a phone number using digits, e.g. 0917 123 4567').build());

  form.addCheckboxItem().setTitle(Q.insights).setRequired(true)
      .setHelpText('Your feedback is how we make Exoid Solar fit the way Philippine installers work.')
      .setChoiceValues(['I agree to share my insights and feedback about Exoid Solar with the Exoid team during and after the beta test (for example, a short survey or call).']);

  // ---------- Wire up ----------
  props.setProperties({ FORM_ID: form.getId(), SHEET_ID: ss.getId() });
  ScriptApp.newTrigger('onOnboard').forForm(form).onFormSubmit().create();

  Logger.log('FORM (send this link to your selected testers): ' + form.shortenFormUrl(form.getPublishedUrl()));
  Logger.log('FORM editor: ' + form.getEditUrl());
  Logger.log('SHEET (team only): ' + ss.getUrl());
}

function readResponse_(r) {
  const v = {};
  r.getItemResponses().forEach(ir => { const a = ir.getResponse(); v[ir.getItem().getTitle()] = Array.isArray(a) ? a.join('; ') : a; });
  return {
    when: r.getTimestamp(), name: (v[Q.name] || '').trim(), company: (v[Q.company] || '').trim(),
    email: (v[Q.email] || '').trim(), phone: (v[Q.phone] || '').trim(), insights: v[Q.insights] ? 'Yes' : 'No', id: r.getId()
  };
}

function firstName_(full) { return (full || '').trim().split(/\s+/)[0] || 'there'; }

function welcomeEmail_(d) {
  const subject = 'Your Exoid Solar beta access';
  const text = 'Hi ' + firstName_(d.name) + ',\n\n' +
    'Thank you for joining the Exoid Solar beta for ' + d.company + '.\n\n' +
    'Open the app here: ' + APP_URL + '\n\n' +
    'Getting started:\n' +
    '1. Sign up with this email address (' + d.email + ') and create your company account.\n' +
    '2. Add your logo and company details in Settings.\n' +
    '3. Check the equipment catalog and enter your own supplier prices.\n' +
    '4. Make your first proposal.\n\n' +
    'As agreed, we\'ll reach out during and after the beta to hear what works and what doesn\'t. ' +
    'You can also reply to this email any time with questions or ideas.\n\n' +
    'Exoid Solar team\n' + ORG;
  const html = '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.55;color:#1f2937;max-width:560px">' +
    '<p>Hi ' + esc_(firstName_(d.name)) + ',</p>' +
    '<p>Thank you for joining the Exoid Solar beta for <b>' + esc_(d.company) + '</b>.</p>' +
    '<p style="margin:24px 0"><a href="' + APP_URL + '" style="background:#F28C28;color:#fff;text-decoration:none;padding:12px 22px;border-radius:6px;font-weight:bold;display:inline-block">Open Exoid Solar</a></p>' +
    '<p style="font-size:13px;color:#6b7280">Or copy this link: ' + APP_URL + '</p>' +
    '<p><b>Getting started</b></p><ol style="padding-left:20px">' +
    '<li>Sign up with this email address (' + esc_(d.email) + ') and create your company account.</li>' +
    '<li>Add your logo and company details in Settings.</li>' +
    '<li>Check the equipment catalog and enter your own supplier prices.</li>' +
    '<li>Make your first proposal.</li></ol>' +
    '<p>As agreed, we\'ll reach out during and after the beta to hear what works and what doesn\'t. You can also reply to this email any time with questions or ideas.</p>' +
    '<p>Exoid Solar team<br><span style="color:#6b7280">' + ORG + '</span></p></div>';
  return { subject, text, html };
}

function esc_(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

// Runs automatically on every submission
function onOnboard(e) {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  const sh = ss.getSheetByName('Beta testers');
  const d = readResponse_(e.response);
  let mailNote = 'Not sent';
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 10, sh.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
    if (ids.indexOf(d.id) >= 0) return;                        // never add or email the same response twice
    try {
      const m = welcomeEmail_(d);
      MailApp.sendEmail({ to: d.email, subject: m.subject, body: m.text, htmlBody: m.html, name: 'Exoid Solar', replyTo: SUPPORT });
      mailNote = 'Sent ' + Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyy-MM-dd HH:mm');
    } catch (err) { mailNote = 'Failed: ' + err.message; }
    sh.appendRow([d.when, d.name, d.company, d.email, "'" + d.phone, d.insights, mailNote, 'Onboarded', '', d.id]);
  } finally { lock.releaseLock(); }
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Beta tester onboarded: ' + d.company,
      d.name + ' from ' + d.company + ' completed the onboarding form and received the app link.\n\nBeta testers sheet (team only):\n' + ss.getUrl());
  } catch (err) {}
}

// Resend the welcome email to every row whose "Welcome email" column doesn't start with "Sent"
function resendFailedWelcomeEmails() {
  const sh = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheetByName('Beta testers');
  const last = sh.getLastRow(); if (last < 2) return;
  const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  rows.forEach((r, i) => {
    if (String(r[6]).indexOf('Sent') === 0) return;
    const d = { name: r[1], company: r[2], email: r[3] };
    try { const m = welcomeEmail_(d);
      MailApp.sendEmail({ to: d.email, subject: m.subject, body: m.text, htmlBody: m.html, name: 'Exoid Solar', replyTo: SUPPORT });
      sh.getRange(i + 2, 7).setValue('Sent ' + Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyy-MM-dd HH:mm'));
    } catch (err) { sh.getRange(i + 2, 7).setValue('Failed: ' + err.message); }
  });
}

// After changing APP_URL above, run this so the form's confirmation page shows the new link
function updateAppLink() {
  FormApp.openById(PropertiesService.getScriptProperties().getProperty('FORM_ID')).setConfirmationMessage(confirmation_());
  Logger.log('Confirmation page now links to ' + APP_URL);
}

// Links the form's responses to the Beta testers spreadsheet (adds a "Form Responses 1" tab with every raw answer)
function linkToSheet() {
  const p = PropertiesService.getScriptProperties();
  const form = FormApp.openById(p.getProperty('FORM_ID'));
  form.setDestination(FormApp.DestinationType.SPREADSHEET, p.getProperty('SHEET_ID'));
  Logger.log('Linked. Responses sheet: ' + SpreadsheetApp.openById(p.getProperty('SHEET_ID')).getUrl());
}
