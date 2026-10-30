/**
 * Storm-Ready Leads — "Check your area" form → Google Sheet + email alert.
 *
 * Setup (about 5 minutes):
 * 1. Create a Google Sheet called "Storm-Ready Leads — Inbound".
 * 2. Extensions → Apps Script. Paste this whole file, replacing what's there.
 * 3. Set NOTIFY_EMAIL below to the inbox that should get alerts.
 * 4. Deploy → New deployment → type "Web app".
 *      Execute as: Me    Who has access: Anyone
 * 5. Authorize when asked, copy the Web app URL, and paste it into
 *    FORM_ENDPOINT near the bottom of index.html.
 */

const NOTIFY_EMAIL = 'you@example.com';
const FIELDS = ['submitted_at', 'name', 'company', 'email', 'phone', 'area',
                'jobs', 'advertising', 'website', 'page'];

function doPost(e) {
  const p = e.parameter || {};
  if (p.company_website) return ContentService.createTextOutput('ok'); // spam trap

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) sheet.appendRow(FIELDS);
  sheet.appendRow(FIELDS.map(function (f) { return p[f] || ''; }));

  if (NOTIFY_EMAIL && NOTIFY_EMAIL.indexOf('@') > 0) {
    const body = FIELDS.map(function (f) { return f + ': ' + (p[f] || ''); }).join('\n');
    MailApp.sendEmail(NOTIFY_EMAIL, 'New Storm-Ready lead: ' + (p.company || p.name || 'unknown'), body);
  }
  return ContentService.createTextOutput('ok');
}
