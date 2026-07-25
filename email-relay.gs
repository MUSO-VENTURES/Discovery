// MUSO VENTURES — Discovery Form Email Relay + Contact Sheet
// Deploy as: Extensions > Apps Script > Deploy > New deployment
//   Type: Web app | Execute as: Me | Who has access: Anyone
// Copy the web app URL into discovery.html as GAS_URL

const TO_EMAIL    = 'smusseau.ventures@gmail.com';
const SHEET_ID    = '1xEl6w-Bu7Anq85-mtLezaNJZurO9btG_Qr0cg5Wm7yo';
const SHEET_NAME  = 'Contacts';

const HEADERS = [
  'Date Submitted',
  'Business Name',
  'Industry / Business Type',
  'Years in Business',
  'Business Phone',
  'Business Email',
  'Address / Location',
  'Website',
  'How They Heard',
  'Primary Contact Name',
  'Primary Role / Title',
  'Primary Direct Phone',
  'Primary Direct Email',
  'Best Time to Reach',
  'Preferred Contact Method',
  'Secondary Contact Name',
  'Secondary Role / Title',
  'Secondary Direct Phone',
  'Secondary Direct Email',
  'Opted Into Contact List',
  'Notes'
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Send email
    MailApp.sendEmail({
      to: TO_EMAIL,
      subject: data._subject || 'MUSO VENTURES — New Discovery Submission',
      htmlBody: buildHTML(data),
      replyTo: data._replyto || ''
    });

    // Write to contact sheet
    appendToSheet(data);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function appendToSheet(d) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_NAME);

  // Create sheet + header row on first run
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#0D2B3E')
      .setFontColor('#3D9DB3');
    sheet.setFrozenRows(1);
  }

  // Parse secondary contact fields from the payload
  const secParts = (d.secondary_contact || '').split(' — ');

  sheet.appendRow([
    new Date().toLocaleDateString('en-US', { year:'numeric', month:'short', day:'numeric' }),
    d.business_name       || '',
    d.industry            || '',
    d.years               || '',
    d.phone               || '',
    d.email               || '',
    d.location            || '',
    d.website             || '',
    d.referral            || '',
    d.p_name              || '',
    d.p_role              || '',
    d.p_phone             || '',
    d.p_email             || '',
    d.p_time              || '',
    d.p_contact_method    || '',
    d.s_name              || '',
    d.s_role              || '',
    d.s_phone             || '',
    d.s_email             || '',
    d.opted_in            || '',
    ''  // Notes — filled manually
  ]);

  // Auto-resize columns for readability
  sheet.autoResizeColumns(1, HEADERS.length);
}

function buildHTML(d) {
  const accent = '#3D9DB3';
  const ink    = '#0D2B3E';

  const row = (label, value) => value
    ? `<tr><td style="padding:6px 12px 6px 0;font-size:11px;color:#666;font-family:monospace;text-transform:uppercase;letter-spacing:0.08em;white-space:nowrap;vertical-align:top;">${label}</td>
       <td style="padding:6px 0;font-size:13px;color:#1a1a1a;">${value}</td></tr>`
    : '';

  const section = (title, content) =>
    `<div style="margin-bottom:24px;">
      <div style="font-size:10px;letter-spacing:0.2em;text-transform:uppercase;color:${accent};font-family:monospace;padding-bottom:6px;border-bottom:2px solid ${accent};margin-bottom:12px;">${title}</div>
      ${content}
    </div>`;

  const responses = d.full_responses || '';
  const qaBlocks = responses
    .split(/\nQ\d+: /)
    .filter(b => b.trim() && b.includes('\nAnswer:'))
    .map(b => {
      const lines  = b.split('\n');
      const qText  = lines[0].trim();
      const answer = (lines.find(l => l.startsWith('Answer:')) || '').replace('Answer:', '').trim();
      const notes  = (lines.find(l => l.startsWith('Notes:'))  || '').replace('Notes:',  '').trim();
      return { qText, answer, notes };
    });

  const qaHTML = qaBlocks.map(({ qText, answer, notes }) => `
    <div style="padding:10px 0;border-bottom:1px solid #eee;">
      <div style="font-size:11px;color:#888;margin-bottom:4px;">${qText}</div>
      <div style="font-size:14px;color:${accent};font-weight:600;">${answer || '— Not answered'}</div>
      ${notes ? `<div style="font-size:12px;color:#555;margin-top:3px;">${notes}</div>` : ''}
    </div>`).join('');

  const contactInfo = section('Business Information', `
    <table style="border-collapse:collapse;width:100%;">
      ${row('Business',   d.business_name)}
      ${row('Industry',   d.industry)}
      ${row('Years',      d.years)}
      ${row('Phone',      d.phone)}
      ${row('Email',      d.email)}
      ${row('Location',   d.location)}
      ${row('Website',    d.website)}
      ${row('Referral',   d.referral)}
    </table>`);

  const primaryContact = section('Primary Contact', `
    <table style="border-collapse:collapse;width:100%;">
      ${row('Name',             d.p_name)}
      ${row('Role / Title',     d.p_role)}
      ${row('Direct Phone',     d.p_phone)}
      ${row('Direct Email',     d.p_email)}
      ${row('Best Time',        d.p_time)}
      ${row('Contact Method',   d.p_contact_method)}
    </table>`);

  const secondaryContact = d.s_name
    ? section('Secondary Contact', `
        <table style="border-collapse:collapse;width:100%;">
          ${row('Name',         d.s_name)}
          ${row('Role / Title', d.s_role)}
          ${row('Direct Phone', d.s_phone)}
          ${row('Direct Email', d.s_email)}
        </table>`)
    : '';

  let extraContacts = '';
  let i = 1;
  while (d[`extra_contact_${i}`]) {
    extraContacts += section(`Additional Contact ${i}`, `<table style="border-collapse:collapse;width:100%;">${row('Name / Role', d[`extra_contact_${i}`])}</table>`);
    i++;
  }

  const optinLine = `<div style="font-size:11px;color:${d.opted_in === 'Yes' ? '#22c55e' : '#999'};margin-top:4px;">${d.opted_in === 'Yes' ? '✓ Opted into contact list' : '— Did not opt in'}</div>`;

  const questionSection = qaBlocks.length ? section('Discovery Responses', qaHTML) : '';

  return `
    <!DOCTYPE html><html>
    <body style="margin:0;padding:0;background:#f4f4f4;font-family:'DM Sans',Arial,sans-serif;">
      <div style="max-width:640px;margin:32px auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.08);">
        <div style="background:${ink};padding:28px 36px;">
          <div style="font-size:10px;letter-spacing:0.24em;text-transform:uppercase;color:${accent};margin-bottom:6px;font-family:monospace;">MUSO VENTURES</div>
          <div style="font-size:24px;color:#fff;font-weight:700;">${d.business_name || 'New Client Discovery'}</div>
          <div style="font-size:11px;color:rgba(255,255,255,0.5);margin-top:4px;">${new Date().toLocaleDateString('en-US',{year:'numeric',month:'long',day:'numeric'})}</div>
        </div>
        <div style="padding:32px 36px;">
          ${contactInfo}${primaryContact}${secondaryContact}${extraContacts}${optinLine}
          <div style="margin-top:28px;">${questionSection}</div>
        </div>
        <div style="background:#f9f9f9;padding:16px 36px;border-top:1px solid #eee;font-size:11px;color:#aaa;font-family:monospace;">
          Sent via MUSO VENTURES Discovery Form — ${d._subject || ''}
        </div>
      </div>
    </body></html>`;
}
