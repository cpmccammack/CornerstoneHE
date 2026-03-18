// ============================================================
// Cornerstone HE — Elite CRM
// Paste this entire file into Google Apps Script (script.google.com)
// Then run: setupCRM()
// ============================================================

// ── Config ────────────────────────────────────────────────────
const STAGES = ['New Lead', 'Contacted', 'Quote Sent', 'Follow-Up', 'Won', 'Lost'];
const SOURCES = ['Referral', 'Website', 'Job Site', 'Cold Call', 'Social Media', 'Other'];
const STAGE_WEIGHTS = { 'New Lead': 0.1, 'Contacted': 0.25, 'Quote Sent': 0.6, 'Follow-Up': 0.45, 'Won': 1.0, 'Lost': 0 };
const FOLLOWUP_DAYS = { 'New Lead': 1, 'Contacted': 2, 'Quote Sent': 4, 'Follow-Up': 3, 'Won': 7, 'Lost': 30 };

// ── Setup: run once ───────────────────────────────────────────
function setupCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Delete existing sheets if re-running
  ['Leads', 'Dashboard', 'Settings'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });

  setupSettings(ss);
  setupLeads(ss);
  setupDashboard(ss);

  // Set Leads as active
  ss.setActiveSheet(ss.getSheetByName('Leads'));

  SpreadsheetApp.getUi().alert('✅ CRM setup complete!\n\nSheets created:\n• Leads\n• Dashboard\n• Settings\n\nRun installTriggers() next to enable automation.');
}

// ── Settings sheet ────────────────────────────────────────────
function setupSettings(ss) {
  const s = ss.insertSheet('Settings');
  s.getRange('A1').setValue('STAGES').setFontWeight('bold');
  s.getRange('B1').setValue('SOURCES').setFontWeight('bold');
  STAGES.forEach((v, i) => s.getRange(i + 2, 1).setValue(v));
  SOURCES.forEach((v, i) => s.getRange(i + 2, 2).setValue(v));
  s.hideSheet();
}

// ── Leads sheet ───────────────────────────────────────────────
function setupLeads(ss) {
  const s = ss.insertSheet('Leads');

  const headers = [
    'Lead ID', 'Date Added', 'Contact Name', 'Phone', 'Email',
    'Address', 'Deal Value ($)', 'Source', 'Stage',
    'Last Touch', 'Next Action Date', 'Notes',
    'Days Since Touch', 'Overdue?', 'Priority Score', 'Next Best Action'
  ];

  // Header row
  const headerRange = s.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1a3a1a').setFontColor('white').setFontWeight('bold').setFontSize(10);
  s.setFrozenRows(1);

  // Column widths
  const widths = [70, 90, 140, 110, 170, 200, 100, 100, 100, 90, 110, 200, 90, 75, 100, 200];
  widths.forEach((w, i) => s.setColumnWidth(i + 1, w));

  // Stage dropdown validation
  const stageRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName('Settings').getRange('A2:A7'), true)
    .setAllowInvalid(false).build();
  s.getRange('I2:I1000').setDataValidation(stageRule);

  // Source dropdown validation
  const sourceRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName('Settings').getRange('B2:B7'), true)
    .setAllowInvalid(false).build();
  s.getRange('H2:H1000').setDataValidation(sourceRule);

  // Auto-formulas (cols M–P) — only apply to row 2 as template
  // Days Since Touch (M)
  s.getRange('M2').setFormula('=IF(J2="","",TODAY()-J2)');
  // Overdue? (N)
  s.getRange('N2').setFormula(
    '=IF(OR(I2="Won",I2="Lost",K2=""),"—",IF(TODAY()>K2,"OVERDUE","OK"))'
  );
  // Priority Score (O) — 0-100
  s.getRange('O2').setFormula(
    '=IF(OR(I2="",G2=""),"",ROUND(MIN(100,' +
    'IF(I2="New Lead",10,IF(I2="Contacted",25,IF(I2="Quote Sent",60,IF(I2="Follow-Up",45,IF(I2="Won",100,0)))))' +
    '*MAX(0.1,1-(MAX(0,M2-3)/30))' +
    '*(G2/5000)' +
    '),0))'
  );
  // Next Best Action (P)
  s.getRange('P2').setFormula(
    '=IF(I2="New Lead","Call to introduce",IF(I2="Contacted",IF(M2>=2,"Follow-up call","Wait — reached out recently"),IF(I2="Quote Sent",IF(M2>=4,"Follow up on quote","Wait — quote is recent"),IF(I2="Follow-Up","Check in — push for decision",IF(I2="Won","Collect deposit & schedule work",IF(I2="Lost","Check back in 30 days","—"))))))'
  );

  // Conditional formatting — Overdue = red, OK = green
  const overdueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('OVERDUE').setBackground('#ffcccc').setFontColor('#cc0000')
    .setRanges([s.getRange('N2:N1000')]).build();
  const okRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('OK').setBackground('#d9ead3').setFontColor('#274e13')
    .setRanges([s.getRange('N2:N1000')]).build();
  s.setConditionalFormatRules([overdueRule, okRule]);

  // Stage color coding
  const stageColors = {
    'New Lead': ['#fff2cc', '#7d6608'],
    'Contacted': ['#cfe2f3', '#1c4587'],
    'Quote Sent': ['#d9d2e9', '#4a235a'],
    'Follow-Up': ['#fce5cd', '#7f4f24'],
    'Won': ['#d9ead3', '#274e13'],
    'Lost': ['#f4cccc', '#660000'],
  };
  const stageRules = Object.entries(stageColors).map(([stage, [bg, fg]]) =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(stage).setBackground(bg).setFontColor(fg)
      .setRanges([s.getRange('I2:I1000')]).build()
  );
  s.setConditionalFormatRules([overdueRule, okRule, ...stageRules]);

  // Add one sample row
  addSampleLead(s);
}

function addSampleLead(s) {
  const today = new Date();
  const nextAction = new Date(today); nextAction.setDate(today.getDate() + 2);
  s.getRange(2, 1, 1, 12).setValues([[
    'L-001', today, 'Joel Shine', '502-356-6420', 'joelshine@hotmail.com',
    'Creekwood Dr, Mount Washington KY', 3500, 'Referral', 'Quote Sent',
    today, nextAction, 'Forestry mulching — Lot 5, 1 acre, light density'
  ]]);
  // Apply formula row
  ['M2','N2','O2','P2'].forEach(cell => {
    const formula = s.getRange(cell).getFormula();
    // formulas already set above
  });
}

// ── Dashboard sheet ───────────────────────────────────────────
function setupDashboard(ss) {
  const s = ss.insertSheet('Dashboard');
  s.setTabColor('#274e13');

  const title = s.getRange('A1');
  title.setValue('CORNERSTONE HE — PIPELINE DASHBOARD')
    .setFontSize(14).setFontWeight('bold').setFontColor('#274e13');

  s.getRange('A2').setValue('Auto-refreshes when you open the sheet').setFontColor('#999').setFontSize(9);

  // Section headers
  const sections = [
    ['A4', 'PIPELINE OVERVIEW'],
    ['A12', 'DEALS BY STAGE'],
    ['D12', 'CONTACT TODAY'],
  ];
  sections.forEach(([cell, label]) => {
    s.getRange(cell).setValue(label).setFontWeight('bold').setFontSize(10)
      .setFontColor('#1a3a1a').setBorder(false, false, true, false, false, false, '#1a3a1a', SpreadsheetApp.BorderStyle.SOLID);
  });

  // Pipeline metrics (formulas referencing Leads sheet)
  const metrics = [
    ['A5', 'Total Pipeline Value', "=SUMIF(Leads!I:I,\"<>Lost\",Leads!G:G)"],
    ['A6', 'Active Deals', "=COUNTIFS(Leads!I:I,\"<>Won\",Leads!I:I,\"<>Lost\",Leads!A:A,\"<>\")"],
    ['A7', 'Overdue Follow-Ups', "=COUNTIF(Leads!N:N,\"OVERDUE\")"],
    ['A8', 'Won This Month', "=COUNTIFS(Leads!I:I,\"Won\",Leads!J:J,\">=\"&DATE(YEAR(TODAY()),MONTH(TODAY()),1))"],
    ['A9', 'Close Rate', "=IFERROR(COUNTIF(Leads!I:I,\"Won\")/(COUNTIF(Leads!I:I,\"Won\")+COUNTIF(Leads!I:I,\"Lost\")),0)"],
    ['A10', 'Avg Deal Size', "=AVERAGEIF(Leads!G:G,\">0\",Leads!G:G)"],
  ];

  metrics.forEach(([cell, label, formula]) => {
    const row = parseInt(cell.slice(1));
    s.getRange('A' + row).setValue(label).setFontColor('#555').setFontSize(10);
    s.getRange('B' + row).setFormula(formula).setFontWeight('bold').setFontSize(11);
    if (cell === 'A5' || cell === 'A10') s.getRange('B' + row).setNumberFormat('$#,##0');
    if (cell === 'A9') s.getRange('B' + row).setNumberFormat('0%');
  });

  // Deals by stage
  STAGES.forEach((stage, i) => {
    const row = 13 + i;
    s.getRange('A' + row).setValue(stage).setFontSize(10);
    s.getRange('B' + row).setFormula(`=COUNTIF(Leads!I:I,"${stage}")`).setFontSize(10);
    s.getRange('C' + row).setFormula(`=SUMIF(Leads!I:I,"${stage}",Leads!G:G)`).setNumberFormat('$#,##0').setFontSize(10);
  });

  s.getRange('A13:A18').setBackground('#f3f3f3');

  // Contact today list header
  s.getRange('D13').setValue('Name').setFontWeight('bold').setFontSize(9);
  s.getRange('E13').setValue('Deal $').setFontWeight('bold').setFontSize(9);
  s.getRange('F13').setValue('Next Action').setFontWeight('bold').setFontSize(9);

  // Array formula for who to contact today
  s.getRange('D14').setFormula(
    '=IFERROR(FILTER(Leads!C:C,Leads!K:K<=TODAY(),Leads!I:I<>"Won",Leads!I:I<>"Lost",Leads!C:C<>"Contact Name"),"")'
  );
  s.getRange('E14').setFormula(
    '=IFERROR(FILTER(Leads!G:G,Leads!K:K<=TODAY(),Leads!I:I<>"Won",Leads!I:I<>"Lost",Leads!C:C<>"Contact Name"),"")'
  );
  s.getRange('F14').setFormula(
    '=IFERROR(FILTER(Leads!P:P,Leads!K:K<=TODAY(),Leads!I:I<>"Won",Leads!I:I<>"Lost",Leads!C:C<>"Contact Name"),"")'
  );

  s.setColumnWidth(1, 180);
  s.setColumnWidth(2, 120);
  s.setColumnWidth(3, 120);
  s.setColumnWidth(4, 160);
  s.setColumnWidth(5, 90);
  s.setColumnWidth(6, 200);
}

// ── onEdit trigger: auto-stamp dates, extend formulas ─────────
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Leads') return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;

  // Auto-stamp Last Touch when Stage changes (col 9)
  if (col === 9) {
    sheet.getRange(row, 10).setValue(new Date()); // Last Touch = today
    // Auto-set Next Action Date based on stage
    const stage = e.value;
    const daysOut = FOLLOWUP_DAYS[stage] || 3;
    const nextDate = new Date();
    nextDate.setDate(nextDate.getDate() + daysOut);
    sheet.getRange(row, 11).setValue(nextDate);
  }

  // Auto-generate Lead ID if new row
  if (col === 3 && !sheet.getRange(row, 1).getValue()) {
    const lastId = sheet.getRange(row - 1, 1).getValue();
    const num = lastId ? parseInt(lastId.toString().replace('L-', '')) + 1 : 1;
    sheet.getRange(row, 1).setValue('L-' + String(num).padStart(3, '0'));
    sheet.getRange(row, 2).setValue(new Date()); // Date Added
  }

  // Extend formulas to new rows
  if (row > 2 && col <= 12) {
    const formulaCells = ['M', 'N', 'O', 'P'];
    formulaCells.forEach(col => {
      const templateFormula = sheet.getRange(col + '2').getFormula();
      if (templateFormula) {
        const newFormula = templateFormula.replace(/2/g, row.toString());
        sheet.getRange(col + row).setFormula(newFormula);
      }
    });
  }
}

// ── Web App: receive leads from Bid Tool ──────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Leads');
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const newRow = lastRow + 1;

    // Auto Lead ID
    const lastId = sheet.getRange(lastRow, 1).getValue();
    const num = lastId ? parseInt(lastId.toString().replace('L-', '')) + 1 : 1;
    const leadId = 'L-' + String(num).padStart(3, '0');

    const today = new Date();
    const nextAction = new Date(today); nextAction.setDate(today.getDate() + 1);

    sheet.getRange(newRow, 1, 1, 12).setValues([[
      leadId,
      today,
      data.name || '',
      data.phone || '',
      data.email || '',
      data.address || '',
      data.dealValue || 0,
      data.source || 'Website',
      'New Lead',
      today,
      nextAction,
      `Forestry mulching · ${data.acreage || '?'} ac · ${data.density || '?'} · ${data.difficulty || 0}% difficulty`
    ]]);

    // Extend formulas
    ['M','N','O','P'].forEach(c => {
      const tmpl = sheet.getRange(c + '2').getFormula();
      if (tmpl) sheet.getRange(c + newRow).setFormula(tmpl.replace(/2/g, newRow));
    });

    return ContentService.createTextOutput(JSON.stringify({ success: true, leadId }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Install triggers ──────────────────────────────────────────
function installTriggers() {
  // Remove existing
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // onEdit
  ScriptApp.newTrigger('onEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();

  // Daily overdue check at 8am
  ScriptApp.newTrigger('sendOverdueDigest').timeBased().everyDays(1).atHour(8).create();

  SpreadsheetApp.getUi().alert('✅ Triggers installed!\n• onEdit — auto-stamps dates and IDs\n• Daily 8am digest — overdue follow-up summary');
}

// ── Daily digest: email overdue leads ─────────────────────────
function sendOverdueDigest() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leads');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('Contact Name');
  const overdueCol = headers.indexOf('Overdue?');
  const actionCol = headers.indexOf('Next Best Action');
  const dealCol = headers.indexOf('Deal Value ($)');

  const overdue = data.slice(1).filter(r => r[overdueCol] === 'OVERDUE');
  if (!overdue.length) return;

  let body = `Good morning! You have ${overdue.length} overdue follow-up${overdue.length > 1 ? 's' : ''}:\n\n`;
  overdue.forEach(r => {
    body += `• ${r[nameCol]} — $${Number(r[dealCol]).toLocaleString()} — ${r[actionCol]}\n`;
  });
  body += `\nLog in to your CRM to take action today.`;

  GmailApp.sendEmail(Session.getActiveUser().getEmail(), '🔔 Cornerstone CRM — Overdue Follow-Ups', body);
}
