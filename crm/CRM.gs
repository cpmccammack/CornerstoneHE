// ============================================================
// Cornerstone HE — CRM Backend
// 1. Create a new Google Sheet
// 2. Go to Extensions → Apps Script
// 3. Paste this entire file, replacing any existing code
// 4. Run setupCRM() once to create the sheet structure
// 5. Deploy as Web App: Deploy → New deployment → Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 6. Copy the Web App URL into the bid tool app
// ============================================================

const STAGES = ['New Lead', 'Contacted', 'Quote Sent', 'Follow-Up', 'Won', 'Lost'];
const SOURCES = ['Referral', 'Website', 'Job Site', 'Cold Call', 'Social Media', 'Bid Tool'];
const FOLLOWUP_DAYS = { 'New Lead': 1, 'Contacted': 2, 'Quote Sent': 4, 'Follow-Up': 3, 'Won': 7, 'Lost': 30 };

const COMPANY = {
  name: 'Cornerstone Hardscape & Excavation',
  address: '651 Reed Lane, Simpsonville, KY 40067',
  phone: '502-396-7887',
  email: 'isaacmosko@cornerstonehe.net',
  rep: 'Isaac Moskovich',
};

// ── Helpers ───────────────────────────────────────────────────
function corsResponse(data, callback) {
  const json = JSON.stringify(data);
  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${json})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function htmlPage(body) {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html><head><meta charset="UTF-8">
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <style>
      body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
        background:#f5f5f5;display:flex;align-items:center;justify-content:center;
        min-height:100vh;margin:0;padding:20px;box-sizing:border-box;}
      .card{background:white;border-radius:12px;padding:40px 36px;max-width:420px;
        width:100%;text-align:center;box-shadow:0 4px 24px rgba(0,0,0,0.08);}
      h2{margin:0 0 10px;font-size:22px;color:#1a1a1a;}
      p{color:#666;font-size:14px;line-height:1.6;margin:0;}
      .icon{font-size:48px;margin-bottom:16px;}
      .co{font-size:12px;color:#aaa;margin-top:20px;}
    </style></head><body><div class="card">${body}</div></body></html>
  `);
}

// ── GET handler ───────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  const cb = e.parameter.callback || null;
  try {
    // Data reads
    if (action === 'getLeads')  return corsResponse(getLeads(), cb);
    if (action === 'getLead')   return corsResponse(getLead(e.parameter.id), cb);
    if (action === 'ping')      return corsResponse({ ok: true }, cb);

    // Customer-facing approval links (return HTML pages — no JSONP needed)
    if (action === 'approveQuote') {
      const id = e.parameter.leadId;
      if (id) {
        updateLead({ id, stage: 'Won' });
        addNote({ id, note: 'Customer approved quote via email link' });
        markFinancing(id, e.parameter.financing === '1');
      }
      return htmlPage(`
        <div class="icon">✅</div>
        <h2>Quote Approved!</h2>
        <p>Thanks for choosing Cornerstone. We'll be in touch shortly to get your project scheduled.</p>
        <p class="co">Cornerstone Hardscape & Excavation · 502-396-7887</p>
      `);
    }

    if (action === 'declineQuote') {
      const id = e.parameter.leadId;
      if (id) {
        updateLead({ id, stage: 'Lost' });
        addNote({ id, note: 'Customer declined quote via email link' });
      }
      return htmlPage(`
        <div class="icon">👋</div>
        <h2>Got it, no worries.</h2>
        <p>Thanks for considering Cornerstone. If you change your mind or need anything in the future, we're always here.</p>
        <p class="co">Cornerstone Hardscape & Excavation · 502-396-7887</p>
      `);
    }

    if (action === 'requestFinancing') {
      const id = e.parameter.leadId;
      if (id) {
        addNote({ id, note: 'Customer requested financing info via email link' });
        updateLead({ id, stage: 'Follow-Up' });
        markFinancing(id, true);
      }
      return htmlPage(`
        <div class="icon">💰</div>
        <h2>Financing Request Received</h2>
        <p>We'll reach out with financing options shortly. Thanks for your interest!</p>
        <p class="co">Cornerstone Hardscape & Excavation · 502-396-7887</p>
      `);
    }

    // Write operations via GET + JSONP for CORS compatibility
    const data = e.parameter.data ? JSON.parse(e.parameter.data) : {};
    if (action === 'createLead')   return corsResponse(createLead(data), cb);
    if (action === 'updateLead')   return corsResponse(updateLead(data), cb);
    if (action === 'sendQuote')    return corsResponse(sendQuoteEmail(data), cb);
    if (action === 'sendFollowUp') return corsResponse(sendFollowUpEmail(data), cb);
    if (action === 'addNote')      return corsResponse(addNote(data), cb);
    if (action === 'scheduleJob')          return corsResponse(scheduleJob(data), cb);
    if (action === 'sendJobConfirmation')  return corsResponse(sendJobConfirmation(data), cb);

    return corsResponse({ error: 'Unknown action' }, cb);
  } catch (err) {
    return corsResponse({ error: err.message }, cb);
  }
}

// ── POST handler ──────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    if (action === 'createLead')   return corsResponse(createLead(data));
    if (action === 'updateLead')   return corsResponse(updateLead(data));
    if (action === 'sendQuote')    return corsResponse(sendQuoteEmail(data));
    if (action === 'sendFollowUp') return corsResponse(sendFollowUpEmail(data));
    if (action === 'addNote')      return corsResponse(addNote(data));
    return corsResponse({ error: 'Unknown action' });
  } catch (err) {
    return corsResponse({ error: err.message });
  }
}

// ── Sheet helpers ─────────────────────────────────────────────
function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leads');
}

const COLS = {
  id: 1, dateAdded: 2, name: 3, phone: 4, email: 5,
  address: 6, dealValue: 7, source: 8, stage: 9,
  lastTouch: 10, nextAction: 11, notes: 12,
  acreage: 13, density: 14, difficulty: 15, estimateTotal: 16,
  approved: 17, financing: 18
};

function rowToObj(row) {
  return {
    id:            row[COLS.id - 1],
    dateAdded:     row[COLS.dateAdded - 1],
    name:          row[COLS.name - 1],
    phone:         row[COLS.phone - 1],
    email:         row[COLS.email - 1],
    address:       row[COLS.address - 1],
    dealValue:     row[COLS.dealValue - 1],
    source:        row[COLS.source - 1],
    stage:         row[COLS.stage - 1],
    lastTouch:     row[COLS.lastTouch - 1],
    nextAction:    row[COLS.nextAction - 1],
    notes:         row[COLS.notes - 1],
    acreage:       row[COLS.acreage - 1],
    density:       row[COLS.density - 1],
    difficulty:    row[COLS.difficulty - 1],
    estimateTotal: row[COLS.estimateTotal - 1],
    approved:      row[COLS.approved - 1] || '',
    financing:     row[COLS.financing - 1] || false,
  };
}

// ── Read leads ────────────────────────────────────────────────
function getLeads() {
  const s = getSheet();
  const lastRow = s.getLastRow();
  if (lastRow < 2) return { leads: [] };
  const rows = s.getRange(2, 1, lastRow - 1, 18).getValues();
  const leads = rows
    .filter(r => r[0])
    .map(rowToObj)
    .map(l => ({
      ...l,
      daysSinceTouch: l.lastTouch ? Math.floor((new Date() - new Date(l.lastTouch)) / 86400000) : null,
      overdue: l.nextAction && new Date(l.nextAction) < new Date() && l.stage !== 'Won' && l.stage !== 'Lost',
      nextBestAction: getNextAction(l),
    }));
  return { leads };
}

function getLead(id) {
  const all = getLeads().leads;
  const lead = all.find(l => l.id === id);
  return lead ? { lead } : { error: 'Not found' };
}

function getNextAction(lead) {
  const d = lead.daysSinceTouch || 0;
  switch (lead.stage) {
    case 'New Lead':    return 'Call to introduce';
    case 'Contacted':   return d >= 2 ? 'Follow-up call' : 'Wait — reached out recently';
    case 'Quote Sent':  return d >= 4 ? 'Follow up on quote' : 'Wait — quote is recent';
    case 'Follow-Up':   return 'Check in — push for decision';
    case 'Won':         return 'Collect deposit & schedule work';
    case 'Lost':        return 'Check back in 30 days';
    default:            return '—';
  }
}

// ── Create lead ───────────────────────────────────────────────
function createLead(data) {
  const s = getSheet();
  const lastRow = s.getLastRow();
  const newRow = lastRow + 1;

  const lastId = lastRow >= 2 ? s.getRange(lastRow, 1).getValue() : 'L-000';
  const num = lastId ? parseInt(lastId.toString().replace('L-', '')) + 1 : 1;
  const id = 'L-' + String(num).padStart(3, '0');

  const today = new Date();
  const nextDate = new Date(today);
  nextDate.setDate(today.getDate() + (FOLLOWUP_DAYS['New Lead'] || 1));

  s.getRange(newRow, 1, 1, 18).setValues([[
    id, today,
    data.name || '', data.phone || '', data.email || '',
    data.address || '', data.dealValue || 0, data.source || 'Bid Tool',
    'New Lead', today, nextDate, data.notes || '',
    data.acreage || '', data.density || '', data.difficulty || 0, data.estimateTotal || 0,
    '', false
  ]]);

  return { success: true, id };
}

// ── Update lead ───────────────────────────────────────────────
function updateLead(data) {
  const s = getSheet();
  const lastRow = s.getLastRow();
  if (lastRow < 2) return { error: 'No leads found' };
  const rows = s.getRange(2, 1, lastRow - 1, 1).getValues();
  const rowIndex = rows.findIndex(r => r[0] === data.id);
  if (rowIndex === -1) return { error: 'Lead not found' };
  const sheetRow = rowIndex + 2;

  if (data.stage)     s.getRange(sheetRow, COLS.stage).setValue(data.stage);
  if (data.notes)     s.getRange(sheetRow, COLS.notes).setValue(data.notes);
  if (data.name)      s.getRange(sheetRow, COLS.name).setValue(data.name);
  if (data.phone)     s.getRange(sheetRow, COLS.phone).setValue(data.phone);
  if (data.email)     s.getRange(sheetRow, COLS.email).setValue(data.email);
  if (data.dealValue) s.getRange(sheetRow, COLS.dealValue).setValue(data.dealValue);

  if (data.stage) {
    s.getRange(sheetRow, COLS.lastTouch).setValue(new Date());
    const daysOut = FOLLOWUP_DAYS[data.stage] || 3;
    const nextDate = new Date();
    nextDate.setDate(nextDate.getDate() + daysOut);
    s.getRange(sheetRow, COLS.nextAction).setValue(nextDate);
  }

  return { success: true };
}

function markFinancing(id, wantsFinancing) {
  const s = getSheet();
  if (s.getLastRow() < 2) return;
  const rows = s.getRange(2, 1, s.getLastRow() - 1, 1).getValues();
  const rowIndex = rows.findIndex(r => r[0] === id);
  if (rowIndex === -1) return;
  s.getRange(rowIndex + 2, COLS.financing).setValue(wantsFinancing ? 'Yes' : 'No');
}

// ── Add note ──────────────────────────────────────────────────
function addNote(data) {
  const s = getSheet();
  if (s.getLastRow() < 2) return { error: 'No leads found' };
  const rows = s.getRange(2, 1, s.getLastRow() - 1, 1).getValues();
  const rowIndex = rows.findIndex(r => r[0] === data.id);
  if (rowIndex === -1) return { error: 'Lead not found' };
  const sheetRow = rowIndex + 2;
  const existing = s.getRange(sheetRow, COLS.notes).getValue();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yy');
  const newNote = `[${timestamp}] ${data.note}${existing ? '\n' + existing : ''}`;
  s.getRange(sheetRow, COLS.notes).setValue(newNote);
  s.getRange(sheetRow, COLS.lastTouch).setValue(new Date());
  return { success: true };
}

// ── Send Quote Email ──────────────────────────────────────────
function sendQuoteEmail(data) {
  if (!data.email) return { error: 'No email address provided' };

  // Create lead first so we have an ID for the approval links
  let leadId = data.leadId;
  if (!leadId) {
    const result = createLead({ ...data, source: 'Bid Tool' });
    leadId = result.id;
  }

  const scriptUrl = ScriptApp.getService().getUrl();
  const approveUrl  = `${scriptUrl}?action=approveQuote&leadId=${leadId}`;
  const declineUrl  = `${scriptUrl}?action=declineQuote&leadId=${leadId}`;
  const financeUrl  = `${scriptUrl}?action=requestFinancing&leadId=${leadId}`;

  const offerDate = new Date();
  offerDate.setDate(offerDate.getDate() + 7);
  const offerStr = Utilities.formatDate(offerDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');

  const densityLabel = { light: 'Light', medium: 'Medium', dense: 'Dense' }[data.density] || data.density || '';
  const RATES    = { light: 3500, medium: 4500, dense: 6500 };
  const PROD_DAY = { light: 2, medium: 1, dense: 0.5 };
  const days = data.timeline || (data.acreage && data.density ? Math.ceil(data.acreage / PROD_DAY[data.density]) : 1);
  const calcBase = data.acreage && data.density ? (days * RATES[data.density]) : 0;
  const calcTotal = Math.round(calcBase * (1 + ((data.difficulty || 0) / 100)));
  // Prefer the custom price from the bid tool; fall back to calculated
  const total = Math.round(data.estimateTotal || calcTotal || 0);

  const scopeLines = [
    `Cornerstone Hardscape & Excavation will perform forestry mulching services at ${data.address || 'the property'}.`,
    data.acreage ? `The area to be cleared is approximately ${data.acreage} acres of ${densityLabel.toLowerCase()} vegetation density.` : '',
    `Estimated project duration: ${days} ${days == 1 ? 'day' : 'days'}.`,
    `Our crew will use a professional forestry mulcher to clear, chip, and spread all vegetation on site — leaving a clean, mulched surface with no hauling required.`,
  ].filter(Boolean).join(' ');

  const html = `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  * { box-sizing: border-box; }
  body { margin: 0; padding: 20px; background: #f0f0f0; font-family: Arial, sans-serif; color: #222; }
  .page { max-width: 680px; margin: 0 auto; background: white; border-radius: 4px; overflow: hidden;
    box-shadow: 0 2px 16px rgba(0,0,0,0.1); }

  /* Header */
  .hdr { background: #000000; padding: 28px 36px; display: flex; justify-content: space-between; align-items: center; }
  .logo-block { color: white; }
  .logo-name { font-size: 22px; font-weight: 900; letter-spacing: 1px; }
  .logo-tag { font-size: 10px; color: rgba(255,255,255,0.5); letter-spacing: 2px; text-transform: uppercase; margin-top: 2px; }
  .quote-id { color: rgba(255,255,255,0.5); font-size: 12px; text-align: right; }
  .quote-id strong { display: block; color: white; font-size: 18px; margin-top: 2px; }

  /* Bill row */
  .bill { display: flex; justify-content: space-between; padding: 24px 36px; border-bottom: 1px solid #eee; gap: 20px; }
  .bill-col h4 { margin: 0 0 6px; font-size: 10px; text-transform: uppercase; letter-spacing: 1px; color: #aaa; }
  .bill-col p { margin: 2px 0; font-size: 13px; color: #444; }
  .bill-col strong { color: #1a1a1a; font-size: 14px; }
  .offer { font-size: 11px; color: #888; margin-top: 10px !important; }
  .offer strong { color: #c84444; }

  /* Body sections */
  .section { padding: 20px 36px; border-bottom: 1px solid #eee; }
  .section-title { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; color: #aaa; margin-bottom: 10px; }
  .scope-text { font-size: 13px; line-height: 1.7; color: #444; }

  /* Line items */
  .line { display: flex; justify-content: space-between; align-items: flex-start; padding: 12px 0; border-bottom: 1px solid #f5f5f5; gap: 16px; }
  .line:last-child { border-bottom: none; }
  .line-desc strong { font-size: 14px; display: block; margin-bottom: 3px; }
  .line-desc span { font-size: 12px; color: #888; }
  .line-price { font-size: 15px; font-weight: 700; color: #1a1a1a; white-space: nowrap; }

  /* Total */
  .total-row { display: flex; justify-content: space-between; padding: 6px 0; font-size: 13px; color: #888; }
  .total-row.grand { font-size: 18px; font-weight: 800; color: #1a1a1a; border-top: 2px solid #1a1a1a; margin-top: 8px; padding-top: 14px; }

  /* Actions */
  .actions { padding: 28px 36px; text-align: center; background: #fafafa; border-bottom: 1px solid #eee; }
  .actions p { font-size: 13px; color: #666; margin: 0 0 18px; }
  .btn-row { display: flex; gap: 12px; justify-content: center; flex-wrap: wrap; margin-bottom: 16px; }
  .btn { display: inline-block; padding: 13px 28px; border-radius: 6px; font-size: 14px; font-weight: 700;
    text-decoration: none; letter-spacing: 0.3px; }
  .btn-approve { background: #1a6b1a; color: white; }
  .btn-decline { background: white; color: #888; border: 1px solid #ddd; }
  .btn-finance { display: inline-block; font-size: 12px; color: #1a6b1a; text-decoration: underline; }

  /* Signature */
  .sig { padding: 20px 36px 24px; }
  .sig-name { font-size: 15px; font-weight: 700; color: #1a1a1a; margin-bottom: 2px; }
  .sig-title { font-size: 12px; color: #888; }
  .sig-contact { font-size: 12px; color: #555; margin-top: 8px; }

  /* Footer */
  .footer { background: #000000; padding: 14px 36px; font-size: 11px; color: rgba(255,255,255,0.4); text-align: center; }
</style>
</head>
<body>
<div class="page">

  <!-- Header -->
  <table width="100%" bgcolor="#000000" cellpadding="0" cellspacing="0" style="background:#000000;">
    <tr>
      <td style="padding:0;">
        <img src="https://cpmccammack.github.io/CornerstoneHE/logo.png" alt="Cornerstone" width="100%" style="display:block;width:100%;border:0;">
      </td>
    </tr>
  </table>

  <!-- Quote ID -->
  <div style="padding:14px 36px 0;font-size:12px;color:#aaa;">Quote <strong style="color:#333;">${leadId}</strong> &nbsp;·&nbsp; ${todayStr}</div>

  <!-- Bill to / from -->
  <div class="bill">
    <div class="bill-col">
      <h4>Prepared For</h4>
      <strong>${data.name || 'Customer'}</strong>
      <p>${data.phone || ''}</p>
      <p>${data.email || ''}</p>
      <p>${data.address || ''}</p>
      <p class="offer">Offer good until: <strong>${offerStr}</strong></p>
    </div>
    <div class="bill-col" style="text-align:right">
      <h4>From</h4>
      <strong>${COMPANY.rep}</strong>
      <p>${COMPANY.name}</p>
      <p>${COMPANY.address}</p>
      <p>${COMPANY.phone}</p>
      <p>${COMPANY.email}</p>
    </div>
  </div>

  <!-- Scope -->
  <div class="section">
    <div class="section-title">Scope of Work</div>
    <div class="scope-text">${scopeLines}</div>
  </div>

  <!-- Line items -->
  <div class="section">
    <div class="section-title">Services</div>
    <div class="line">
      <div class="line-desc">
        <strong>Forestry Mulching${densityLabel ? ' — ' + densityLabel + ' Density' : ''}</strong>
        <span>${data.acreage ? data.acreage + ' acres · ' : ''}All vegetation cleared, chipped, and spread on site. No hauling required.</span>
      </div>
      <div class="line-price">$${total.toLocaleString()}</div>
    </div>
    <div style="padding-top:14px">
      <div class="total-row grand"><span>Total</span><span>$${total.toLocaleString()}</span></div>
    </div>
  </div>

  <!-- Approve / Decline / Financing -->
  <div class="actions">
    <p>Please review your quote and let us know how you'd like to proceed.</p>
    <div class="btn-row">
      <a class="btn btn-approve" href="${approveUrl}">✓ &nbsp;Approve Quote</a>
      <a class="btn btn-decline" href="${declineUrl}">Decline</a>
    </div>
    <a class="btn-finance" href="${financeUrl}">Interested in financing options?</a>
  </div>

  <!-- Map Image -->
  ${data.mapImageUrl ? `
  <div class="section">
    <div class="section-title">Project Area</div>
    <img src="${data.mapImageUrl}" alt="Project area map" style="width:100%;border-radius:4px;display:block;">
  </div>` : ''}

  <!-- Signature -->
  <div class="sig">
    <div class="sig-name">${COMPANY.rep}</div>
    <div class="sig-title">Cornerstone Hardscape &amp; Excavation</div>
    <div class="sig-contact">${COMPANY.phone} &nbsp;·&nbsp; ${COMPANY.email}</div>
  </div>

  <div class="footer">Thank you for considering Cornerstone. We appreciate your business.</div>
</div>
</body>
</html>`;

  GmailApp.sendEmail(
    data.email,
    `Your Cornerstone Quote — $${total.toLocaleString()} (${leadId})`,
    `Hi ${(data.name || 'there').split(' ')[0]},\n\nYour forestry mulching quote is ready.\n\nTotal: $${total.toLocaleString()}\nTimeline: ${days} ${days == 1 ? 'day' : 'days'}\nOffer good until: ${offerStr}\n\nApprove: ${approveUrl}\nDecline: ${declineUrl}\nFinancing: ${financeUrl}\n\n${COMPANY.rep}\n${COMPANY.phone}`,
    { htmlBody: html, name: COMPANY.name, replyTo: COMPANY.email }
  );

  updateLead({ id: leadId, stage: 'Quote Sent' });
  return { success: true, id: leadId };
}

// ── Send Follow-Up Email ──────────────────────────────────────
function sendFollowUpEmail(data) {
  if (!data.email) return { error: 'No email address' };

  const firstName = (data.name || 'there').split(' ')[0];
  const message = data.message ||
    `Hi ${firstName},\n\nJust following up on the quote we sent over. We'd love to get started on your project!\n\nLet us know if you have any questions.\n\n${COMPANY.rep}\n${COMPANY.phone}`;

  GmailApp.sendEmail(
    data.email,
    data.subject || `Following up — Cornerstone Quote`,
    message,
    { name: COMPANY.name, replyTo: COMPANY.email }
  );

  if (data.leadId) {
    addNote({ id: data.leadId, note: 'Follow-up email sent' });
    updateLead({ id: data.leadId, stage: 'Follow-Up' });
  }

  return { success: true };
}

// ── Send Job Confirmation Email with .ics calendar invite ─────
function sendJobConfirmation(data) {
  if (!data.email) return { error: 'No email address' };
  const firstName = (data.name || 'there').split(' ')[0];
  const dateFmt = data.scheduledDate
    ? Utilities.formatDate(new Date(data.scheduledDate + 'T12:00:00'), Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy')
    : null;

  // Build .ics attachment if a date was provided
  let icsBlob = null;
  if (data.scheduledDate) {
    // DTSTART/DTEND for all-day event: YYYYMMDD
    const startDate = new Date(data.scheduledDate + 'T12:00:00');
    const endDate   = new Date(data.scheduledDate + 'T12:00:00');
    endDate.setDate(endDate.getDate() + 1);
    const fmt = d => Utilities.formatDate(d, 'UTC', 'yyyyMMdd');
    const uid = `cornerstone-${data.leadId || Date.now()}@cornerstonehe.net`;
    const ics = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Cornerstone Hardscape & Excavation//EN',
      'METHOD:REQUEST',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTART;VALUE=DATE:${fmt(startDate)}`,
      `DTEND;VALUE=DATE:${fmt(endDate)}`,
      `SUMMARY:Cornerstone Forestry Mulching${data.address ? ' — ' + data.address : ''}`,
      `DESCRIPTION:Forestry mulching by Cornerstone Hardscape & Excavation.\\n${data.address ? 'Address: ' + data.address + '\\n' : ''}${data.total ? 'Quote Total: $' + Number(data.total).toLocaleString() + '\\n' : ''}Questions? Call ${COMPANY.phone}`,
      `LOCATION:${data.address || ''}`,
      `ORGANIZER;CN=${COMPANY.rep}:mailto:${COMPANY.email}`,
      `ATTENDEE;CN=${data.name || 'Customer'};RSVP=TRUE:mailto:${data.email}`,
      'STATUS:CONFIRMED',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    icsBlob = Utilities.newBlob(ics, 'text/calendar', 'Cornerstone-Job.ics');
  }

  const emailOpts = {
    name: COMPANY.name,
    replyTo: COMPANY.email,
    htmlBody: `<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;color:#222;max-width:520px;margin:0 auto;padding:20px;">
      <table width="100%" bgcolor="#000000" cellpadding="0" cellspacing="0"><tr><td style="padding:0;">
        <img src="https://cpmccammack.github.io/CornerstoneHE/logo.png" alt="Cornerstone" width="100%" style="display:block;width:100%;border:0;">
      </td></tr></table>
      <div style="background:white;border:1px solid #eee;border-top:none;padding:28px 32px;">
        <h2 style="margin:0 0 6px;font-size:20px;">Your Job is Confirmed</h2>
        <p style="color:#666;margin:0 0 24px;font-size:13px;">Hi ${firstName}, here are your confirmed project details. A calendar invite is attached.</p>
        <table style="width:100%;border-collapse:collapse;font-size:13px;">
          ${dateFmt ? `<tr><td style="padding:8px 0;border-bottom:1px solid #f0f0f0;color:#888;width:120px;">Scheduled</td><td style="padding:8px 0;border-bottom:1px solid #f0f0f0;font-weight:700;">${dateFmt}</td></tr>` : ''}
          ${data.address ? `<tr><td style="padding:8px 0;border-bottom:1px solid #f0f0f0;color:#888;">Address</td><td style="padding:8px 0;border-bottom:1px solid #f0f0f0;">${data.address}</td></tr>` : ''}
          ${data.total ? `<tr><td style="padding:8px 0;color:#888;">Quote Total</td><td style="padding:8px 0;font-weight:700;font-size:16px;">$${Number(data.total).toLocaleString()}</td></tr>` : ''}
        </table>
        <p style="margin:20px 0 0;font-size:13px;color:#666;">If this date doesn't work, just reply to this email or call us at <strong>${COMPANY.phone}</strong>.</p>
        <hr style="border:none;border-top:1px solid #eee;margin:24px 0;">
        <p style="margin:0;font-size:13px;color:#444;">${COMPANY.rep}<br><span style="color:#888;">${COMPANY.name} · ${COMPANY.phone}</span></p>
      </div>
    </body></html>`,
  };
  if (icsBlob) emailOpts.attachments = [icsBlob];

  GmailApp.sendEmail(
    data.email,
    dateFmt ? `Cornerstone — Job Confirmed for ${dateFmt}` : `Cornerstone — Your Project is Confirmed`,
    `Hi ${firstName},\n\nYour forestry mulching project with Cornerstone has been confirmed.\n\n${dateFmt ? 'Scheduled Date: ' + dateFmt + '\n' : ''}Address: ${data.address || ''}\n${data.total ? 'Quote Total: $' + Number(data.total).toLocaleString() : ''}\n\nA calendar invite is attached. If this date doesn't work, call us at ${COMPANY.phone}.\n\n${COMPANY.rep}\n${COMPANY.name}`,
    emailOpts
  );

  if (data.leadId) addNote({ id: data.leadId, note: `Job confirmation + calendar invite sent${dateFmt ? ' for ' + dateFmt : ''}` });
  return { success: true };
}

// ── Schedule Job on Google Calendar ──────────────────────────
function scheduleJob(data) {
  const calendar = CalendarApp.getDefaultCalendar();
  const start = new Date(data.startDate);
  const end   = new Date(data.endDate || data.startDate);
  if (start.getTime() === end.getTime()) end.setDate(end.getDate() + 1);

  const title = `Cornerstone — ${data.name || 'Job'} (${data.address || ''})`;

  // Generate PDF and get its Drive URL
  const pdf = createQuotePDF(data);
  const pdfUrl = pdf ? pdf.fileUrl : null;

  const desc = [
    `Customer: ${data.name || ''}`,
    `Phone: ${data.phone || ''}`,
    `Email: ${data.email || ''}`,
    `Address: ${data.address || ''}`,
    data.acreage ? `Area: ${data.acreage} acres (${data.density || ''} density)` : '',
    data.total ? `Quote Total: $${Number(data.total).toLocaleString()}` : '',
    data.leadId ? `Lead ID: ${data.leadId}` : '',
    pdfUrl ? `\nQuote PDF: ${pdfUrl}` : '',
    '',
    'Scope: Forestry mulching — clear, chip, and spread all vegetation on site. No hauling required.',
  ].filter(Boolean).join('\n');

  const opts = { description: desc, location: data.address || '' };

  const event = calendar.createAllDayEvent(title, start, end, opts);

  // Send customer email invite via Gmail (more reliable than addGuest)
  if (data.customerEmail) {
    try {
      const startFmt = Utilities.formatDate(start, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');
      const endFmt   = Utilities.formatDate(end, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');
      GmailApp.sendEmail(
        data.customerEmail,
        `Cornerstone — Job Scheduled: ${startFmt}`,
        `Hi ${(data.name || 'there').split(' ')[0]},\n\nYour forestry mulching project has been scheduled.\n\nStart: ${startFmt}\nEnd: ${endFmt}\nAddress: ${data.address || ''}\n${data.total ? 'Quote Total: $' + Number(data.total).toLocaleString() : ''}\n\nWe'll see you then! If you have any questions, call us at ${COMPANY.phone}.\n\n${COMPANY.rep}\n${COMPANY.name}\n${COMPANY.phone}`,
        {
          name: COMPANY.name,
          replyTo: COMPANY.email,
          htmlBody: `<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;color:#222;max-width:520px;margin:0 auto;padding:20px;">
            <table width="100%" bgcolor="#000000" cellpadding="0" cellspacing="0"><tr><td style="padding:0;">
              <img src="https://cpmccammack.github.io/CornerstoneHE/logo.png" alt="Cornerstone" width="100%" style="display:block;width:100%;border:0;">
            </td></tr></table>
            <div style="background:white;border:1px solid #eee;border-top:none;padding:28px 32px;">
              <h2 style="margin:0 0 6px;font-size:20px;">Your Job Has Been Scheduled</h2>
              <p style="color:#666;margin:0 0 24px;font-size:13px;">Hi ${(data.name || 'there').split(' ')[0]}, here are your project details.</p>
              <table style="width:100%;border-collapse:collapse;font-size:13px;">
                <tr><td style="padding:8px 0;border-bottom:1px solid #f0f0f0;color:#888;width:120px;">Start Date</td><td style="padding:8px 0;border-bottom:1px solid #f0f0f0;font-weight:700;">${startFmt}</td></tr>
                <tr><td style="padding:8px 0;border-bottom:1px solid #f0f0f0;color:#888;">End Date</td><td style="padding:8px 0;border-bottom:1px solid #f0f0f0;font-weight:700;">${endFmt}</td></tr>
                ${data.address ? `<tr><td style="padding:8px 0;border-bottom:1px solid #f0f0f0;color:#888;">Address</td><td style="padding:8px 0;border-bottom:1px solid #f0f0f0;">${data.address}</td></tr>` : ''}
                ${data.total ? `<tr><td style="padding:8px 0;color:#888;">Quote Total</td><td style="padding:8px 0;font-weight:700;font-size:16px;">$${Number(data.total).toLocaleString()}</td></tr>` : ''}
              </table>
              ${pdfUrl ? `<p style="margin:20px 0 0;font-size:12px;color:#888;">Quote PDF: <a href="${pdfUrl}" style="color:#1a6b1a;">${data.name || 'Customer'} Quote</a></p>` : ''}
              <hr style="border:none;border-top:1px solid #eee;margin:24px 0;">
              <p style="margin:0;font-size:13px;color:#444;">${COMPANY.rep}<br><span style="color:#888;">${COMPANY.name} · ${COMPANY.phone}</span></p>
            </div>
          </body></html>`
        }
      );
    } catch(e) { /* email send failed silently */ }
  }

  // Attach PDF to calendar event via Calendar API
  if (pdf) {
    try {
      const token = ScriptApp.getOAuthToken();
      const calId  = encodeURIComponent(calendar.getId());
      const evtId  = event.getId().replace('@google.com', '');
      UrlFetchApp.fetch(
        `https://www.googleapis.com/calendar/v3/calendars/${calId}/events/${evtId}?supportsAttachments=true`,
        {
          method: 'PATCH',
          headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
          payload: JSON.stringify({ attachments: [{ fileId: pdf.fileId, title: pdf.fileName, mimeType: 'application/pdf' }] }),
          muteHttpExceptions: true,
        }
      );
    } catch(e) { /* attachment failed silently — event still created */ }
  }

  if (data.leadId) {
    addNote({ id: data.leadId, note: `Job scheduled: ${Utilities.formatDate(start, Session.getScriptTimeZone(), 'MM/dd/yyyy')}${pdfUrl ? ' · PDF saved to Drive' : ''}` });
  }

  return { success: true, eventId: event.getId(), pdfUrl };
}

// ── Generate Quote PDF → save to Drive ───────────────────────
function createQuotePDF(data) {
  try {
    const leadId = data.leadId || 'DRAFT';
    const name   = data.name || 'Customer';
    const total  = data.total ? `$${Number(data.total).toLocaleString()}` : 'See quote';
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
    const offerDate = new Date(); offerDate.setDate(offerDate.getDate() + 7);
    const offerStr  = Utilities.formatDate(offerDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');

    // Build a clean Google Doc
    const doc  = DocumentApp.create(`Cornerstone Quote — ${leadId}`);
    const body = doc.getBody();
    body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

    const titlePara = body.appendParagraph('CORNERSTONE HARDSCAPE & EXCAVATION');
    titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    const subPara = body.appendParagraph('651 Reed Lane, Simpsonville, KY 40067  ·  502-396-7887  ·  isaacmosko@cornerstonehe.net');
    subPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    subPara.setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 9, [DocumentApp.Attribute.FOREGROUND_COLOR]: '#888888' });

    body.appendParagraph('').setSpacingAfter(6);

    body.appendParagraph(`Quote ${leadId}  ·  ${dateStr}`).setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 10, [DocumentApp.Attribute.FOREGROUND_COLOR]: '#888888' });

    body.appendParagraph('').setSpacingAfter(4);

    const prep = body.appendParagraph(`Prepared for: ${name}`);
    prep.setAttributes({ [DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FONT_SIZE]: 12 });
    if (data.phone)   body.appendParagraph(String(data.phone)).setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 11 });
    if (data.email)   body.appendParagraph(data.email).setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 11 });
    if (data.address) body.appendParagraph(data.address).setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 11 });
    body.appendParagraph(`Offer good until: ${offerStr}`).setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 10, [DocumentApp.Attribute.FOREGROUND_COLOR]: '#cc4444' });

    body.appendParagraph('').setSpacingAfter(8);

    const sow = body.appendParagraph('SCOPE OF WORK');
    sow.setAttributes({ [DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FONT_SIZE]: 10, [DocumentApp.Attribute.FOREGROUND_COLOR]: '#888888' });

    const densityLabel = { light: 'Light', medium: 'Medium', dense: 'Dense' }[data.density] || data.density || '';
    const scopeText = [
      `Cornerstone Hardscape & Excavation will perform forestry mulching services at ${data.address || 'the specified property'}.`,
      data.acreage ? `The area to be cleared is approximately ${data.acreage} acres of ${densityLabel.toLowerCase()} vegetation density.` : '',
      `Our crew will use a professional forestry mulcher to clear, chip, and spread all vegetation on site — leaving a clean, mulched surface with no hauling required.`,
    ].filter(Boolean).join(' ');
    body.appendParagraph(scopeText).setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 11 });

    body.appendParagraph('').setSpacingAfter(8);

    const svc = body.appendParagraph('SERVICES');
    svc.setAttributes({ [DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FONT_SIZE]: 10, [DocumentApp.Attribute.FOREGROUND_COLOR]: '#888888' });

    const svcLine = body.appendParagraph(`Forestry Mulching${densityLabel ? ' — ' + densityLabel + ' Density' : ''}${data.acreage ? '  (' + data.acreage + ' acres)' : ''}`);
    svcLine.setAttributes({ [DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FONT_SIZE]: 12 });

    body.appendParagraph('').setSpacingAfter(4);

    const totalLine = body.appendParagraph(`Total: ${total}`);
    totalLine.setAttributes({ [DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FONT_SIZE]: 16 });

    body.appendParagraph('').setSpacingAfter(12);
    body.appendParagraph('Thank you for choosing Cornerstone. Questions? Call 502-396-7887')
      .setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 10, [DocumentApp.Attribute.FOREGROUND_COLOR]: '#888888', [DocumentApp.Attribute.ITALIC]: true });

    doc.saveAndClose();

    // Export as PDF
    const docId   = doc.getId();
    const pdfBlob = UrlFetchApp.fetch(
      `https://docs.google.com/document/d/${docId}/export?format=pdf`,
      { headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` } }
    ).getBlob();

    const fileName = `Cornerstone-Quote-${leadId}-${name.replace(/\s+/g, '-')}.pdf`;
    pdfBlob.setName(fileName);

    // Save to Drive folder
    const folders = DriveApp.getFoldersByName('Cornerstone Quotes');
    const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder('Cornerstone Quotes');
    const pdfFile = folder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Delete temp Doc
    DriveApp.getFileById(docId).setTrashed(true);

    return { fileId: pdfFile.getId(), fileUrl: pdfFile.getUrl(), fileName };
  } catch(e) {
    return null;
  }
}

// ── Auth triggers: run once each to grant permissions ────────
function testDriveAccess() {
  const folders = DriveApp.getFoldersByName('Cornerstone Quotes');
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Cornerstone Quotes');
  Logger.log('Drive OK: ' + folder.getName());
}

function testDocumentCreate() {
  const doc = DocumentApp.create('Cornerstone Test Doc — DELETE ME');
  doc.getBody().appendParagraph('Test');
  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  Logger.log('DocumentApp OK');
}

// ── Setup: run once ───────────────────────────────────────────
function setupCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Leads'].forEach(name => { const s = ss.getSheetByName(name); if (s) ss.deleteSheet(s); });

  const s = ss.insertSheet('Leads');
  const headers = ['Lead ID','Date Added','Name','Phone','Email','Address','Deal Value','Source','Stage','Last Touch','Next Action','Notes','Acreage','Density','Difficulty %','Estimate Total','Approved','Financing'];
  s.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1a3a1a').setFontColor('white').setFontWeight('bold');
  s.setFrozenRows(1);

  const widths = [70,100,140,110,180,200,90,90,100,100,110,220,70,80,80,100,80,80];
  widths.forEach((w, i) => s.setColumnWidth(i + 1, w));

  SpreadsheetApp.getUi().alert('✅ CRM ready!\n\nNext: Deploy as Web App\nDeploy → New deployment → Web App\nExecute as: Me | Access: Anyone\n\nCopy the URL into your bid tool app.');
}
