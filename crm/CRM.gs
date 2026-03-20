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
      @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
      *{box-sizing:border-box;margin:0;padding:0;}
      body{font-family:'Plus Jakarta Sans',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
        background:#F8FAFC;display:flex;align-items:center;justify-content:center;
        min-height:100vh;padding:20px;}
      .card{background:white;border-radius:16px;padding:48px 40px;max-width:440px;
        width:100%;text-align:center;border:1px solid #E2E8F0;
        box-shadow:0 4px 24px rgba(0,0,0,0.06);}
      .mark{width:56px;height:56px;border-radius:14px;background:#0A0A0A;
        display:flex;align-items:center;justify-content:center;margin:0 auto 24px;}
      h2{font-size:22px;font-weight:700;color:#0F172A;margin-bottom:10px;line-height:1.3;}
      p{color:#64748B;font-size:14px;line-height:1.7;margin:0;}
      .co{font-size:12px;color:#94A3B8;margin-top:28px;padding-top:20px;
        border-top:1px solid #F1F5F9;}
    </style></head><body><div class="card">${body}</div></body></html>
  `);
}

function notifyOwner(subject, body) {
  try {
    GmailApp.sendEmail(COMPANY.email, subject, body, { name: COMPANY.name + ' CRM' });
  } catch(e) {}
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

    // View quote as rendered HTML page
    if (action === 'viewQuote') {
      const lead = getLead(e.parameter.leadId).lead;
      if (!lead) return htmlPage('<h2>Quote not found</h2>');
      const offerDate = new Date(); offerDate.setDate(offerDate.getDate() + 7);
      const offerStr  = Utilities.formatDate(offerDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
      const todayStr  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
      let html;

      // If a custom quote was sent, show that instead of forestry template
      if (lead.lastQuoteJson) {
        try {
          const q = JSON.parse(lead.lastQuoteJson);
          if (q.type === 'custom') {
            const subtotal = q.lineItems.reduce(function(s,l){ return s + (parseFloat(l.qty)||1)*(parseFloat(l.unitPrice)||0); }, 0);
            const linesHtml = q.lineItems.map(function(item) {
              const lineTotal = Math.round((parseFloat(item.qty)||1)*(parseFloat(item.unitPrice)||0));
              return `<div class="line"><div class="line-desc"><strong>${item.service}</strong>${item.description ? '<span>' + item.description + '</span>' : ''}</div><div class="line-price">$${lineTotal.toLocaleString()}</div></div>`;
            }).join('');
            const markupSection = q.markup > 0
              ? `<div class="total-row"><span>Subtotal</span><span>$${Math.round(subtotal).toLocaleString()}</span></div><div class="total-row"><span>Markup (${q.markup}%)</span><span>+$${(q.total - Math.round(subtotal)).toLocaleString()}</span></div>`
              : '';
            html = buildEmailQuoteHtml({ leadId: lead.id, todayStr, offerStr, data: { ...lead, mapImageUrl: '' }, densityLabel: '', days: null, total: q.total, scopeLines: q.notes || '', customLinesHtml: linesHtml + markupSection });
          }
        } catch(err) {}
      }

      // Fallback to forestry mulching template
      if (!html) {
        const densityLabel = { light: 'Light', medium: 'Medium', dense: 'Dense' }[lead.density] || lead.density || '';
        const PROD_DAY = { light: 2, medium: 1, dense: 0.5 };
        const days = lead.acreage && lead.density ? Math.max(1, Math.round((lead.acreage / (PROD_DAY[lead.density]||1)) * 2) / 2) : 1;
        const total = Math.round(parseFloat(lead.estimateTotal) || parseFloat(lead.dealValue) || 0);
        const scopeLines = [
          'Cornerstone Hardscape & Excavation will perform forestry mulching services at ' + (lead.address || 'the property') + '.',
          lead.acreage ? 'The area to be cleared is approximately ' + lead.acreage + ' acres of ' + densityLabel.toLowerCase() + ' vegetation density.' : '',
          'Estimated project duration: ' + days + ' ' + (days == 1 ? 'day' : 'days') + '.',
          'Our crew will use a professional forestry mulcher to clear, chip, and spread all vegetation on site — leaving a clean, mulched surface with no hauling required.',
        ].filter(Boolean).join(' ');
        html = buildEmailQuoteHtml({ leadId: lead.id, todayStr, offerStr, data: lead, densityLabel, days, total, scopeLines });
      }

      return HtmlService.createHtmlOutput(html);
    }

    // ── Invoice page ──────────────────────────────────────────────
    if (action === 'viewInvoice') {
      const lead = getLead(e.parameter.leadId).lead;
      if (!lead) return htmlPage('<h2>Invoice not found.</h2>');
      const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
      const dueDate  = new Date(); dueDate.setDate(dueDate.getDate() + 30);
      const dueDateStr = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
      const total = Math.round(parseFloat(lead.estimateTotal) || parseFloat(lead.dealValue) || 0);
      let lineRowsHtml = '';
      if (lead.lastQuoteJson) {
        try {
          const q = JSON.parse(lead.lastQuoteJson);
          if (q.type === 'custom' && q.lineItems) {
            lineRowsHtml = q.lineItems.map(function(item) {
              const lt = Math.round((parseFloat(item.qty)||1)*(parseFloat(item.unitPrice)||0));
              return '<tr><td style="padding:12px 0;border-bottom:1px solid #f5f5f5;font-size:14px;">' + item.service +
                (item.description ? '<br><span style="font-size:12px;color:#888;">' + item.description + '</span>' : '') +
                '</td><td style="padding:12px 0;border-bottom:1px solid #f5f5f5;text-align:right;font-weight:700;font-size:14px;">$' + lt.toLocaleString() + '</td></tr>';
            }).join('');
          }
        } catch(err) {}
      }
      if (!lineRowsHtml) {
        const dl = {light:'Light',medium:'Medium',dense:'Dense'}[lead.density]||lead.density||'';
        const svcName = 'Forestry Mulching' + (dl ? ' — ' + dl + ' Density' : '');
        lineRowsHtml = '<tr><td style="padding:12px 0;font-size:14px;">' + svcName +
          (lead.acreage ? '<br><span style="font-size:12px;color:#888;">' + lead.acreage + ' acres</span>' : '') +
          '</td><td style="padding:12px 0;text-align:right;font-weight:700;font-size:14px;">$' + total.toLocaleString() + '</td></tr>';
      }
      const invoiceHtml = `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Invoice — ${lead.name || 'Customer'}</title>
<style>*{box-sizing:border-box;}body{margin:0;padding:20px;background:#f0f0f0;font-family:Arial,sans-serif;color:#222;}
.page{max-width:680px;margin:0 auto;background:#fff;border-radius:4px;overflow:hidden;box-shadow:0 2px 16px rgba(0,0,0,0.1);}
.bill{display:flex;justify-content:space-between;padding:24px 36px;border-bottom:1px solid #eee;gap:20px;}
.bill-col h4{margin:0 0 6px;font-size:10px;text-transform:uppercase;letter-spacing:1px;color:#aaa;}
.bill-col p{margin:2px 0;font-size:13px;color:#444;}
.bill-col strong{color:#1a1a1a;font-size:14px;}
.section{padding:20px 36px;border-bottom:1px solid #eee;}
.section-title{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1.5px;color:#aaa;margin-bottom:10px;}
.total-row{display:flex;justify-content:space-between;padding:6px 0;font-size:13px;color:#888;}
.total-row.grand{font-size:18px;font-weight:800;color:#1a1a1a;border-top:2px solid #1a1a1a;margin-top:8px;padding-top:14px;}
.pay{padding:20px 36px;background:#F8FAFC;border-bottom:1px solid #eee;}
.pay h4{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#aaa;margin:0 0 8px;}
.pay p{font-size:13px;color:#444;margin:3px 0;line-height:1.6;}
.footer{background:#000;padding:14px 36px;font-size:11px;color:rgba(255,255,255,0.4);text-align:center;}
</style></head><body><div class="page">
<table width="100%" bgcolor="#000000" cellpadding="0" cellspacing="0"><tr><td style="padding:0;">
  <img src="https://cpmccammack.github.io/CornerstoneHE/logo.png" alt="${COMPANY.name}" width="100%" style="display:block;width:100%;border:0;">
</td></tr></table>
<div style="padding:14px 36px 0;font-size:12px;color:#aaa;">
  Invoice <strong style="color:#333;">${lead.id}</strong> &nbsp;·&nbsp; ${todayStr}
  &nbsp;&nbsp;<span style="color:#DC2626;font-weight:700;">Due: ${dueDateStr}</span>
</div>
<div class="bill">
  <div class="bill-col">
    <h4>Bill To</h4>
    <strong>${lead.name || 'Customer'}</strong>
    <p>${lead.phone || ''}</p><p>${lead.email || ''}</p><p>${lead.address || ''}</p>
  </div>
  <div class="bill-col" style="text-align:right">
    <h4>From</h4>
    <strong>${COMPANY.rep}</strong>
    <p>${COMPANY.name}</p><p>${COMPANY.address}</p><p>${COMPANY.phone}</p><p>${COMPANY.email}</p>
  </div>
</div>
<div class="section">
  <div class="section-title">Services</div>
  <table width="100%" cellpadding="0" cellspacing="0">${lineRowsHtml}</table>
  <div style="padding-top:14px;">
    <div class="total-row grand"><span>Total Due</span><span>$${total.toLocaleString()}</span></div>
  </div>
</div>
<div class="pay">
  <h4>Payment Instructions</h4>
  <p>Please make payment by <strong>${dueDateStr}</strong>.</p>
  <p>Contact us at <strong>${COMPANY.phone}</strong> or <strong>${COMPANY.email}</strong> with any questions.</p>
</div>
<div class="footer">Thank you for your business — ${COMPANY.name}</div>
</div></body></html>`;
      return HtmlService.createHtmlOutput(invoiceHtml);
    }

    // ── Sign & Approve page ───────────────────────────────────────
    if (action === 'signQuote') {
      const id   = e.parameter.leadId;
      const lead = id ? getLead(id).lead : null;
      if (!lead) return htmlPage('<h2>Quote not found.</h2>');

      // Already signed — show confirmation instead of form
      if (lead.approved === 'approved') {
        return htmlPage(`
          <div class="mark"><svg width="24" height="24" fill="none" stroke="white" stroke-width="2" viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"/></svg></div>
          <h2>Already Signed</h2>
          <p>This quote has already been approved and signed. We'll be in touch soon to get your project scheduled.</p>
          <p class="co">${COMPANY.name} &middot; ${COMPANY.phone}</p>
        `);
      }

      // Build quote HTML
      const scriptUrl2 = ScriptApp.getService().getUrl();
      const offerDate2 = new Date(); offerDate2.setDate(offerDate2.getDate() + 7);
      const offerStr2  = Utilities.formatDate(offerDate2, Session.getScriptTimeZone(), 'MM/dd/yyyy');
      const todayStr2  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
      let quoteHtml;
      if (lead.lastQuoteJson) {
        try {
          const q2 = JSON.parse(lead.lastQuoteJson);
          if (q2.type === 'custom') {
            const sub2 = q2.lineItems.reduce(function(s,l){ return s+(parseFloat(l.qty)||1)*(parseFloat(l.unitPrice)||0);},0);
            const lh2 = q2.lineItems.map(function(item){
              const lt=Math.round((parseFloat(item.qty)||1)*(parseFloat(item.unitPrice)||0));
              return `<div class="line"><div class="line-desc"><strong>${item.service}</strong>${item.description?'<span>'+item.description+'</span>':''}</div><div class="line-price">$${lt.toLocaleString()}</div></div>`;
            }).join('');
            const mk2 = q2.markup>0?`<div class="total-row"><span>Subtotal</span><span>$${Math.round(sub2).toLocaleString()}</span></div><div class="total-row"><span>Markup (${q2.markup}%)</span><span>+$${(q2.total-Math.round(sub2)).toLocaleString()}</span></div>`:'';
            quoteHtml = buildEmailQuoteHtml({leadId:id,todayStr:todayStr2,offerStr:offerStr2,data:{...lead,mapImageUrl:''},densityLabel:'',days:null,total:q2.total,scopeLines:q2.notes||'',customLinesHtml:lh2+mk2});
          }
        } catch(err2) {}
      }
      if (!quoteHtml) {
        const dl2 = {light:'Light',medium:'Medium',dense:'Dense'}[lead.density]||lead.density||'';
        const pd2 = {light:2,medium:1,dense:0.5};
        const dy2 = lead.acreage&&lead.density?Math.max(1,Math.round((lead.acreage/(pd2[lead.density]||1))*2)/2):1;
        const tt2 = Math.round(parseFloat(lead.estimateTotal)||parseFloat(lead.dealValue)||0);
        const sc2 = [
          COMPANY.name+' will perform forestry mulching services at '+(lead.address||'the property')+'.',
          lead.acreage?'The area to be cleared is approximately '+lead.acreage+' acres of '+dl2.toLowerCase()+' vegetation density.':'',
          'Estimated project duration: '+dy2+' '+(dy2==1?'day':'days')+'.',
          'Our crew will use a professional forestry mulcher to clear, chip, and spread all vegetation on site — leaving a clean, mulched surface with no hauling required.',
        ].filter(Boolean).join(' ');
        quoteHtml = buildEmailQuoteHtml({leadId:id,todayStr:todayStr2,offerStr:offerStr2,data:lead,densityLabel:dl2,days:dy2,total:tt2,scopeLines:sc2});
      }

      // Append signature form before </body>
      const signForm = `
<div style="max-width:680px;margin:24px auto 40px;background:#fff;border-radius:4px;padding:32px 36px;box-shadow:0 2px 16px rgba(0,0,0,0.08);font-family:Arial,sans-serif;">
  <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1.5px;color:#aaa;margin-bottom:16px;">Sign & Approve</div>
  <p style="font-size:13px;color:#444;line-height:1.7;margin:0 0 20px;">By signing below, you agree to the terms and pricing outlined in this quote and authorize <strong>${COMPANY.name}</strong> to proceed with the work.</p>
  <form method="GET" action="${scriptUrl2}">
    <input type="hidden" name="action" value="submitSignature">
    <input type="hidden" name="leadId" value="${id}">
    <label style="display:block;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#888;margin-bottom:6px;">Full Legal Name</label>
    <input type="text" name="sigName" required placeholder="Type your full name to sign"
      style="width:100%;padding:12px 14px;border:1px solid #ddd;border-radius:6px;font-size:15px;font-family:Arial,sans-serif;margin-bottom:16px;outline:none;">
    <label style="display:flex;align-items:flex-start;gap:10px;font-size:13px;color:#555;line-height:1.5;margin-bottom:20px;cursor:pointer;">
      <input type="checkbox" name="agreed" value="1" required style="margin-top:2px;flex-shrink:0;">
      I have read and agree to the terms of this quote. I authorize ${COMPANY.name} to perform the work as described.
    </label>
    <button type="submit"
      style="width:100%;padding:14px;background:#0A0A0A;color:#fff;border:none;border-radius:8px;font-size:15px;font-weight:700;cursor:pointer;font-family:Arial,sans-serif;letter-spacing:0.3px;">
      Sign &amp; Approve
    </button>
  </form>
</div>`;
      const signedHtml = quoteHtml.replace('</body>', signForm + '</body>');
      return HtmlService.createHtmlOutput(signedHtml);
    }

    // ── Process signature submission ──────────────────────────────
    if (action === 'submitSignature') {
      const id      = e.parameter.leadId;
      const sigName = (e.parameter.sigName || '').trim();
      if (!id || !sigName) return htmlPage('<h2>Missing information.</h2><p>Please go back and complete all fields.</p>');

      const lead = getLead(id).lead;
      if (!lead) return htmlPage('<h2>Quote not found.</h2>');
      if (lead.approved === 'approved') {
        return htmlPage(`
          <div class="mark"><svg width="24" height="24" fill="none" stroke="white" stroke-width="2" viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"/></svg></div>
          <h2>Already Signed</h2>
          <p>This quote was already approved. We'll be in touch shortly.</p>
          <p class="co">${COMPANY.name} &middot; ${COMPANY.phone}</p>
        `);
      }

      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy hh:mm a');
      updateLead({ id, stage: 'Won', approved: 'approved' });
      addNote({ id, note: 'Quote signed by ' + sigName + ' on ' + nowStr });

      const firstName = (lead.name || 'there').split(' ')[0];
      const val = lead.estimateTotal ? '$' + Number(lead.estimateTotal).toLocaleString() : '';

      // Notification to owner
      notifyOwner(
        'Quote Signed — ' + (lead.name || 'Customer') + (val ? ' (' + val + ')' : ''),
        (lead.name || 'Customer') + ' signed and approved their quote.' +
        '\nSigned by: ' + sigName +
        '\nTimestamp: ' + nowStr +
        (lead.address ? '\nAddress: ' + lead.address : '') +
        (val ? '\nTotal: ' + val : '') +
        '\nLead ID: ' + id +
        '\n\nLog in to the CRM to schedule the job.'
      );

      // Landing page IS the confirmation — no email sent to client
      return HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>You're all set — ${COMPANY.name.split(' ')[0]}</title>
</head>
<body style="margin:0;padding:0;background:#f0f0f0;font-family:Arial,sans-serif;color:#222;">
  <div style="max-width:680px;margin:0 auto;">
    <table width="100%" bgcolor="#000000" cellpadding="0" cellspacing="0"><tr><td style="padding:0;">
      <img src="https://cpmccammack.github.io/CornerstoneHE/logo.png" alt="${COMPANY.name}" width="100%" style="display:block;width:100%;border:0;">
    </td></tr></table>
    <div style="background:#fff;padding:40px 36px;">
      <p style="font-size:22px;font-weight:700;color:#0F172A;margin:0 0 14px;">Hi ${firstName}, you're all set.</p>
      <p style="font-size:14px;color:#444;line-height:1.8;margin:0 0 14px;">
        Thank you for partnering with <strong>${COMPANY.name}</strong>. We're excited to work with you and look forward to delivering an exceptional result.
      </p>
      <p style="font-size:14px;color:#444;line-height:1.8;margin:0 0 28px;">
        Your signed quote${val ? ' for <strong>' + val + '</strong>' : ''} has been recorded and our team will be in touch shortly to confirm your project timeline.
      </p>
      <div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:10px;padding:18px 20px;margin-bottom:28px;">
        <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#94A3B8;margin-bottom:10px;">Signature Confirmation</div>
        <div style="font-size:14px;color:#0F172A;font-weight:600;margin-bottom:4px;">${sigName}</div>
        <div style="font-size:12px;color:#64748B;">Signed on ${nowStr}</div>
      </div>
      <p style="font-size:13px;color:#888;border-top:1px solid #eee;padding-top:20px;margin:0;">
        ${COMPANY.rep} &nbsp;·&nbsp; ${COMPANY.name}<br>
        ${COMPANY.phone} &nbsp;·&nbsp; <a href="mailto:${COMPANY.email}" style="color:#888;">${COMPANY.email}</a>
      </p>
    </div>
    <div style="background:#000;padding:14px 36px;font-size:11px;color:rgba(255,255,255,0.35);text-align:center;">
      Thank you for choosing ${COMPANY.name}. We appreciate your business.
    </div>
  </div>
</body></html>`);
    }

    // Customer-facing approval links (return HTML pages — no JSONP needed)
    if (action === 'approveQuote') {
      const id = e.parameter.leadId;
      if (id) {
        updateLead({ id, stage: 'Won' });
        addNote({ id, note: 'Customer approved quote via email link' });
        markFinancing(id, e.parameter.financing === '1');
        const lead = getLead(id);
        const name = lead ? (lead.name || 'A customer') : 'A customer';
        const addr = lead ? (lead.address || '') : '';
        const val  = lead ? (lead.estimateTotal ? '$' + Number(lead.estimateTotal).toLocaleString() : '') : '';
        notifyOwner(
          'Quote Approved — ' + name + (val ? ' (' + val + ')' : ''),
          name + ' approved their quote.' +
          (addr ? '\nAddress: ' + addr : '') +
          (val  ? '\nTotal: ' + val : '') +
          '\n\nLead ID: ' + id +
          '\n\nLog in to the CRM to schedule the job.'
        );
      }
      return htmlPage(`
        <div class="mark"><svg width="24" height="24" fill="none" stroke="white" stroke-width="2" viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"/></svg></div>
        <h2>Quote Approved</h2>
        <p>Thank you for choosing ${COMPANY.name.split(' ')[0]}. We'll be in touch shortly to get your project scheduled.</p>
        <p class="co">${COMPANY.name} &middot; ${COMPANY.phone}</p>
      `);
    }

    if (action === 'declineQuote') {
      const id = e.parameter.leadId;
      if (id) {
        updateLead({ id, stage: 'Lost' });
        addNote({ id, note: 'Customer declined quote via email link' });
        const lead = getLead(id);
        const name = lead ? (lead.name || 'A customer') : 'A customer';
        notifyOwner(
          'Quote Declined — ' + name,
          name + ' declined their quote.' +
          '\nLead ID: ' + id +
          '\n\nConsider a follow-up or adjusted quote.'
        );
      }
      return htmlPage(`
        <div class="mark" style="background:#F1F5F9;"><svg width="24" height="24" fill="none" stroke="#64748B" stroke-width="2" viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></div>
        <h2>No worries at all.</h2>
        <p>Thanks for considering ${COMPANY.name.split(' ')[0]}. If your plans change or you need anything down the road, we're always here.</p>
        <p class="co">${COMPANY.name} &middot; ${COMPANY.phone}</p>
      `);
    }

    if (action === 'requestFinancing') {
      const id = e.parameter.leadId;
      if (id) {
        addNote({ id, note: 'Customer requested financing info via email link' });
        updateLead({ id, stage: 'Follow-Up' });
        markFinancing(id, true);
        const lead = getLead(id);
        const name = lead ? (lead.name || 'A customer') : 'A customer';
        const val  = lead ? (lead.estimateTotal ? '$' + Number(lead.estimateTotal).toLocaleString() : '') : '';
        notifyOwner(
          'Financing Requested — ' + name + (val ? ' (' + val + ')' : ''),
          name + ' is interested in financing.' +
          (val ? '\nQuote Total: ' + val : '') +
          '\nLead ID: ' + id +
          '\n\nFollow up with financing options.'
        );
      }
      return htmlPage(`
        <div class="mark"><svg width="24" height="24" fill="none" stroke="white" stroke-width="2" viewBox="0 0 24 24"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg></div>
        <h2>Financing Request Received</h2>
        <p>We'll reach out shortly with options that work for your budget. Thanks for your interest!</p>
        <p class="co">${COMPANY.name} &middot; ${COMPANY.phone}</p>
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
    if (action === 'sendCustomQuote')      return corsResponse(sendCustomQuote(data), cb);
    if (action === 'deleteLead')           return corsResponse(deleteLead(data), cb);

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
  approved: 17, financing: 18, mapImageUrl: 19, scheduledDate: 20, lastQuoteJson: 21
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
    mapImageUrl:   row[COLS.mapImageUrl - 1] || '',
    scheduledDate:  row[COLS.scheduledDate - 1]  || '',
    lastQuoteJson:  row[COLS.lastQuoteJson - 1]   || '',
  };
}

// ── Read leads ────────────────────────────────────────────────
function getLeads() {
  const s = getSheet();
  const lastRow = s.getLastRow();
  if (lastRow < 2) return { leads: [] };
  const rows = s.getRange(2, 1, lastRow - 1, 21).getValues();
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

  s.getRange(newRow, 1, 1, 21).setValues([[
    id, today,
    data.name || '', data.phone || '', data.email || '',
    data.address || '', data.dealValue || 0, data.source || 'Bid Tool',
    'New Lead', today, nextDate, data.notes || '',
    data.acreage || '', data.density || '', data.difficulty || 0, data.estimateTotal || 0,
    '', false, data.mapImageUrl || '', '', ''
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

  if (data.stage)         s.getRange(sheetRow, COLS.stage).setValue(data.stage);
  if (data.notes)         s.getRange(sheetRow, COLS.notes).setValue(data.notes);
  if (data.name)          s.getRange(sheetRow, COLS.name).setValue(data.name);
  if (data.phone)         s.getRange(sheetRow, COLS.phone).setValue(data.phone);
  if (data.email)         s.getRange(sheetRow, COLS.email).setValue(data.email);
  if (data.address)       s.getRange(sheetRow, COLS.address).setValue(data.address);
  if (data.dealValue)     s.getRange(sheetRow, COLS.dealValue).setValue(data.dealValue);
  if (data.estimateTotal) s.getRange(sheetRow, COLS.estimateTotal).setValue(data.estimateTotal);
  if (data.archived)      s.getRange(sheetRow, COLS.stage).setValue('Archived');

  if (data.stage) {
    s.getRange(sheetRow, COLS.lastTouch).setValue(new Date());
    const daysOut = FOLLOWUP_DAYS[data.stage] || 3;
    const nextDate = new Date();
    nextDate.setDate(nextDate.getDate() + daysOut);
    s.getRange(sheetRow, COLS.nextAction).setValue(nextDate);
  }

  return { success: true };
}

function deleteLead(data) {
  const s = getSheet();
  if (s.getLastRow() < 2) return { error: 'No leads found' };
  const rows = s.getRange(2, 1, s.getLastRow() - 1, 1).getValues();
  const rowIndex = rows.findIndex(r => r[0] === data.id);
  if (rowIndex === -1) return { error: 'Lead not found' };
  s.deleteRow(rowIndex + 2);
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
  const approveUrl  = `${scriptUrl}?action=signQuote&leadId=${leadId}`;
  const declineUrl  = `${scriptUrl}?action=declineQuote&leadId=${leadId}`;
  const financeUrl  = `${scriptUrl}?action=requestFinancing&leadId=${leadId}`;

  const offerDate = new Date();
  offerDate.setDate(offerDate.getDate() + 7);
  const offerStr = Utilities.formatDate(offerDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');

  const densityLabel = { light: 'Light', medium: 'Medium', dense: 'Dense' }[data.density] || data.density || '';
  const DAY_RATE = 3500;
  const PROD_DAY = { light: 2, medium: 1, dense: 0.5 };
  const days = data.timeline || (data.acreage && data.density ? Math.max(1, Math.round((data.acreage / PROD_DAY[data.density]) * 2) / 2) : 1);
  const calcBase = data.acreage && data.density ? (days * DAY_RATE) : 0;
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
      <a class="btn btn-approve" href="${approveUrl}" style="color:#ffffff !important;">Approve &amp; Sign</a>
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
  const start = new Date((data.startDate || '').replace(/-/g, '/'));
  const end   = new Date((data.endDate || data.startDate || '').replace(/-/g, '/'));
  if (start.getTime() === end.getTime()) end.setDate(end.getDate() + 1);

  const title = `Cornerstone — ${data.name || 'Job'} (${data.address || ''})`;

  // Generate quote view URL (served by this web app — renders perfectly)
  const scriptUrl = ScriptApp.getService().getUrl();
  const pdfUrl = data.leadId ? `${scriptUrl}?action=viewQuote&leadId=${data.leadId}` : null;
  const pdf = null; // no longer saving to Drive

  const desc = [
    `Customer: ${data.name || ''}`,
    `Phone: ${data.phone || ''}`,
    `Email: ${data.email || ''}`,
    `Address: ${data.address || ''}`,
    data.acreage ? `Area: ${data.acreage} acres (${data.density || ''} density)` : '',
    data.total ? `Quote Total: $${Number(data.total).toLocaleString()}` : '',
    data.leadId ? `Lead ID: ${data.leadId}` : '',
    pdfUrl ? `Quote PDF: ${pdfUrl}` : '',
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
    // Store scheduled date for monthly pipeline view
    const s2 = getSheet();
    const rows2 = s2.getRange(2, 1, s2.getLastRow() - 1, 1).getValues();
    const ri = rows2.findIndex(r => r[0] === data.leadId);
    if (ri !== -1) s2.getRange(ri + 2, COLS.scheduledDate).setValue(data.startDate || '');
  }

  return { success: true, eventId: event.getId(), pdfUrl };
}

// ── Shared: save HTML file to Drive (opens in browser, looks like email) ─────
function saveHtmlToDrive(html, leadId, name) {
  try {
    const fileName = 'Cornerstone-Quote-' + leadId + '-' + name.replace(/\s+/g,'-') + '.html';
    const htmlBlob = Utilities.newBlob(html, MimeType.HTML, fileName);

    const folders = DriveApp.getFoldersByName('Cornerstone Quotes');
    const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder('Cornerstone Quotes');
    const file = folder.createFile(htmlBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Return both blob (for email attachment) and file info (for calendar link)
    return { blob: htmlBlob, fileId: file.getId(), fileUrl: file.getUrl(), fileName: fileName };
  } catch(e) {
    return null;
  }
}

// ── Shared: quote PDF HTML template ──────────────────────────
function buildQuoteHtml(opts) {
  const subtotalRow = opts.markup > 0
    ? '<tr><td colspan="2" style="color:#888;font-size:12px;border-bottom:none;padding:6px 0;">Subtotal</td>' +
      '<td style="text-align:right;font-weight:600;color:#888;font-size:12px;border-bottom:none;padding:6px 0;">$' + Math.round(opts.subtotal).toLocaleString() + '</td></tr>' +
      '<tr><td colspan="2" style="color:#888;font-size:12px;border-bottom:none;padding:6px 0;">Markup (' + opts.markup + '%)</td>' +
      '<td style="text-align:right;font-weight:600;color:#888;font-size:12px;border-bottom:none;padding:6px 0;">+$' + (opts.total - Math.round(opts.subtotal)).toLocaleString() + '</td></tr>'
    : '';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    '*{box-sizing:border-box;margin:0;padding:0;}' +
    'body{font-family:Arial,Helvetica,sans-serif;color:#111;background:#fff;font-size:13px;line-height:1.5;}' +
    '</style></head><body>' +

    // Header — typographic style (Drive PDF renderer cannot render background colors)
    '<table width="100%" cellpadding="0" cellspacing="0"><tr><td style="padding:36px 48px 20px;">' +
    '<div style="font-size:22px;font-weight:900;letter-spacing:2px;text-transform:uppercase;color:#000000;">Cornerstone Hardscape &amp; Excavation</div>' +
    '<div style="font-size:10px;color:#888888;margin-top:6px;letter-spacing:0.3px;">651 Reed Lane, Simpsonville, KY 40067 &nbsp;&nbsp;|&nbsp;&nbsp; 502-396-7887 &nbsp;&nbsp;|&nbsp;&nbsp; isaacmosko@cornerstonehe.net</div>' +
    '<div style="border-top:3px solid #000000;margin-top:16px;"></div>' +
    '</td></tr></table>' +

    // Body
    '<div style="padding:24px 48px 40px;">' +

    '<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px;border-bottom:1px solid #eee;padding-bottom:16px;"><tr>' +
    '<td style="font-size:11px;color:#888;">Quote #' + opts.leadId + ' &nbsp;·&nbsp; ' + opts.dateStr + '</td>' +
    '<td style="text-align:right;font-size:11px;color:#cc4444;font-weight:600;">Valid until ' + opts.offerStr + '</td>' +
    '</tr></table>' +

    '<div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#888;margin:0 0 8px;">Prepared For</div>' +
    '<div style="font-size:14px;font-weight:700;margin-bottom:4px;">' + opts.name + '</div>' +
    '<div style="font-size:12px;color:#444;line-height:1.7;">' +
    (opts.phone   ? opts.phone + '<br>' : '') +
    (opts.email   ? opts.email + '<br>' : '') +
    (opts.address ? opts.address        : '') +
    '</div>' +

    '<div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#888;margin:20px 0 8px;">Services</div>' +
    '<table width="100%" cellpadding="0" cellspacing="0">' +
    '<thead><tr>' +
    '<th style="text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:1px;color:#888;padding:8px 0;border-bottom:2px solid #000;">Service</th>' +
    '<th style="text-align:center;font-size:10px;text-transform:uppercase;letter-spacing:1px;color:#888;padding:8px 0;border-bottom:2px solid #000;">Qty</th>' +
    '<th style="text-align:right;font-size:10px;text-transform:uppercase;letter-spacing:1px;color:#888;padding:8px 0;border-bottom:2px solid #000;">Amount</th>' +
    '</tr></thead><tbody>' + opts.lineRows + subtotalRow + '</tbody></table>' +

    '<table width="100%" cellpadding="0" cellspacing="0" style="border-top:2px solid #000;margin-top:10px;padding-top:14px;"><tr>' +
    '<td style="font-size:15px;font-weight:700;">Total</td>' +
    '<td style="text-align:right;font-size:26px;font-weight:700;">$' + opts.total.toLocaleString() + '</td>' +
    '</tr></table>' +

    (opts.scopeText ? '<div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#888;margin:20px 0 8px;">Scope of Work</div>' +
      '<div style="background:#f9f9f9;border-left:3px solid #000;padding:12px 16px;font-size:12px;color:#444;">' + opts.scopeText + '</div>' : '') +
    (opts.notes     ? '<div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#888;margin:20px 0 8px;">Notes</div>' +
      '<div style="background:#f9f9f9;border-left:3px solid #000;padding:12px 16px;font-size:12px;color:#444;">' + opts.notes + '</div>' : '') +

    '<div style="margin-top:36px;padding-top:16px;border-top:1px solid #eee;font-size:11px;color:#888;text-align:center;">' +
    'Thank you for choosing Cornerstone &nbsp;·&nbsp; Questions? Call 502-396-7887</div>' +

    '</div></body></html>';
}

// ── Generate Quote HTML file (forestry mulching) — same HTML as email ────────
function createQuotePDF(data) {
  try {
    const leadId   = data.leadId || 'DRAFT';
    const name     = data.name   || 'Customer';
    const total    = Math.round(parseFloat(data.total) || 0);
    const offerDate = new Date(); offerDate.setDate(offerDate.getDate() + 7);
    const offerStr  = Utilities.formatDate(offerDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    const todayStr  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
    const densityLabel = { light: 'Light', medium: 'Medium', dense: 'Dense' }[data.density] || data.density || '';
    const DAY_RATE = 3500;
    const PROD_DAY = { light: 2, medium: 1, dense: 0.5 };
    const days = data.timeline || (data.acreage && data.density ? Math.max(1, Math.round((data.acreage / (PROD_DAY[data.density]||1)) * 2) / 2) : 1);
    const scopeLines = [
      'Cornerstone Hardscape & Excavation will perform forestry mulching services at ' + (data.address || 'the property') + '.',
      data.acreage ? 'The area to be cleared is approximately ' + data.acreage + ' acres of ' + densityLabel.toLowerCase() + ' vegetation density.' : '',
      'Estimated project duration: ' + days + ' ' + (days == 1 ? 'day' : 'days') + '.',
      'Our crew will use a professional forestry mulcher to clear, chip, and spread all vegetation on site — leaving a clean, mulched surface with no hauling required.',
    ].filter(Boolean).join(' ');

    // Use the same HTML as the email — renders perfectly in a browser
    const html = buildEmailQuoteHtml({ leadId, todayStr, offerStr, data, densityLabel, days, total, scopeLines });
    const result = saveHtmlToDrive(html, leadId, name);
    if (!result) return null;
    return { blob: result.blob, fileId: result.fileId, fileUrl: result.fileUrl, fileName: result.fileName };
  } catch(e) {
    return null;
  }
}

// ── Build email-style quote HTML (shared by email + Drive file) ───────────────
function buildEmailQuoteHtml(opts) {
  const { leadId, todayStr, offerStr, data, densityLabel, days, total, scopeLines, customLinesHtml } = opts;
  return `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  * { box-sizing: border-box; }
  body { margin: 0; padding: 20px; background: #f0f0f0; font-family: Arial, sans-serif; color: #222; }
  .page { max-width: 680px; margin: 0 auto; background: white; border-radius: 4px; overflow: hidden; box-shadow: 0 2px 16px rgba(0,0,0,0.1); }
  .bill { display: flex; justify-content: space-between; padding: 24px 36px; border-bottom: 1px solid #eee; gap: 20px; }
  .bill-col h4 { margin: 0 0 6px; font-size: 10px; text-transform: uppercase; letter-spacing: 1px; color: #aaa; }
  .bill-col p { margin: 2px 0; font-size: 13px; color: #444; }
  .bill-col strong { color: #1a1a1a; font-size: 14px; }
  .offer strong { color: #c84444; }
  .section { padding: 20px 36px; border-bottom: 1px solid #eee; }
  .section-title { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; color: #aaa; margin-bottom: 10px; }
  .line { display: flex; justify-content: space-between; align-items: flex-start; padding: 12px 0; border-bottom: 1px solid #f5f5f5; gap: 16px; }
  .line:last-child { border-bottom: none; }
  .line-desc strong { font-size: 14px; display: block; margin-bottom: 3px; }
  .line-desc span { font-size: 12px; color: #888; }
  .line-price { font-size: 15px; font-weight: 700; color: #1a1a1a; white-space: nowrap; }
  .total-row { display: flex; justify-content: space-between; padding: 6px 0; font-size: 13px; color: #888; }
  .total-row.grand { font-size: 18px; font-weight: 800; color: #1a1a1a; border-top: 2px solid #1a1a1a; margin-top: 8px; padding-top: 14px; }
  .sig { padding: 20px 36px 24px; }
  .footer { background: #000; padding: 14px 36px; font-size: 11px; color: rgba(255,255,255,0.4); text-align: center; }
</style>
</head>
<body>
<div class="page">
  <table width="100%" bgcolor="#000000" cellpadding="0" cellspacing="0"><tr><td bgcolor="#000000" style="padding:0;">
    <img src="https://cpmccammack.github.io/CornerstoneHE/logo.png" alt="Cornerstone" width="100%" style="display:block;width:100%;border:0;">
  </td></tr></table>
  <div style="padding:14px 36px 0;font-size:12px;color:#aaa;">Quote <strong style="color:#333;">${leadId}</strong> &nbsp;·&nbsp; ${todayStr}</div>
  <div class="bill">
    <div class="bill-col">
      <h4>Prepared For</h4>
      <strong>${data.name || 'Customer'}</strong>
      <p>${data.phone || ''}</p><p>${data.email || ''}</p><p>${data.address || ''}</p>
      <p style="font-size:11px;color:#888;margin-top:10px;">Offer good until: <strong style="color:#c84444;">${offerStr}</strong></p>
    </div>
    <div class="bill-col" style="text-align:right">
      <h4>From</h4>
      <strong>${COMPANY.rep}</strong>
      <p>${COMPANY.name}</p><p>${COMPANY.address}</p><p>${COMPANY.phone}</p><p>${COMPANY.email}</p>
    </div>
  </div>
  ${scopeLines ? `<div class="section">
    <div class="section-title">Scope of Work</div>
    <div style="font-size:13px;line-height:1.7;color:#444;">${scopeLines}</div>
  </div>` : ''}
  <div class="section">
    <div class="section-title">Services</div>
    ${customLinesHtml || `<div class="line">
      <div class="line-desc">
        <strong>Forestry Mulching${densityLabel ? ' — ' + densityLabel + ' Density' : ''}</strong>
        <span>${data.acreage ? data.acreage + ' acres · ' : ''}All vegetation cleared, chipped, and spread on site. No hauling required.</span>
      </div>
      <div class="line-price">$${total.toLocaleString()}</div>
    </div>`}
    <div style="padding-top:14px">
      <div class="total-row grand"><span>Total</span><span>$${total.toLocaleString()}</span></div>
    </div>
  </div>
  ${data.mapImageUrl ? `
  <div class="section">
    <div class="section-title">Project Area</div>
    <img src="${data.mapImageUrl}" alt="Project area map" referrerpolicy="no-referrer" style="width:100%;border-radius:4px;display:block;">
  </div>` : ''}
  <div class="sig">
    <div style="font-size:15px;font-weight:700;">${COMPANY.rep}</div>
    <div style="font-size:12px;color:#888;">Cornerstone Hardscape &amp; Excavation</div>
    <div style="font-size:12px;color:#555;margin-top:8px;">${COMPANY.phone} &nbsp;·&nbsp; ${COMPANY.email}</div>
  </div>
  <div class="footer">Thank you for considering Cornerstone. We appreciate your business.</div>
</div>
</body></html>`;
}

// ── Send Custom Multi-Service Quote ──────────────────────────
function sendCustomQuote(data) {
  if (!data.email) return { error: 'No email address on this lead' };

  const name      = data.name || 'Customer';
  const firstName = name.split(' ')[0];
  const lineItems = JSON.parse(data.lineItems || '[]');
  const markup    = parseFloat(data.markup) || 0;
  const total     = Math.round(parseFloat(data.total) || 0);
  const dateStr   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');
  const offerDate = new Date(); offerDate.setDate(offerDate.getDate() + 30);
  const offerStr  = Utilities.formatDate(offerDate, Session.getScriptTimeZone(), 'MMMM d, yyyy');
  const subtotal  = lineItems.reduce(function(s,l){ return s + (parseFloat(l.qty)||1)*(parseFloat(l.unitPrice)||0); }, 0);

  const leadId = data.id || data.leadId || 'DRAFT';
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');

  // Build line items HTML matching the forestry quote style
  const linesHtml = lineItems.map(function(item) {
    const lineTotal = Math.round((parseFloat(item.qty)||1)*(parseFloat(item.unitPrice)||0));
    return `<div class="line">
      <div class="line-desc">
        <strong>${item.service}</strong>
        ${item.description ? '<span>' + item.description + '</span>' : ''}
      </div>
      <div class="line-price">$${lineTotal.toLocaleString()}</div>
    </div>`;
  }).join('');

  const markupSection = markup > 0
    ? `<div class="total-row"><span>Subtotal</span><span>$${Math.round(subtotal).toLocaleString()}</span></div>
       <div class="total-row"><span>Markup (${markup}%)</span><span>+$${(total - Math.round(subtotal)).toLocaleString()}</span></div>`
    : '';

  const htmlBody = buildEmailQuoteHtml({
    leadId, todayStr, offerStr,
    data: { ...data, mapImageUrl: '' },
    densityLabel: '',
    days: null,
    total,
    scopeLines: data.notes || '',
    customLinesHtml: linesHtml + markupSection,
  });

  const subject   = 'Your Quote from Cornerstone' + (data.address ? ' — ' + data.address : '');
  const plainText = 'Hi ' + firstName + ',\n\nPlease see your quote below.\n\nTotal: $' + total.toLocaleString() + '\nOffer valid until: ' + offerStr + '\n\n' + COMPANY.rep + '\n' + COMPANY.name + ' · ' + COMPANY.phone;

  // Save to Drive for record-keeping but do NOT attach — quote is inline in email body
  createCustomQuotePDF(data, lineItems, markup, total);
  const emailOpts = { name: COMPANY.name, replyTo: COMPANY.email, htmlBody: htmlBody };

  GmailApp.sendEmail(data.email, subject, plainText, emailOpts);

  if (leadId) {
    const serviceNames = lineItems.map(function(l){ return l.service; }).join(', ');
    addNote({ id: leadId, note: 'Custom quote sent: ' + serviceNames + ' — $' + total.toLocaleString() });
    updateLead({ id: leadId, stage: 'Quote Sent', estimateTotal: total });
    // Store quote data for viewQuote page
    const s2 = getSheet();
    const rows2 = s2.getRange(2, 1, s2.getLastRow() - 1, 1).getValues();
    const ri = rows2.findIndex(r => r[0] === leadId);
    if (ri !== -1) s2.getRange(ri + 2, COLS.lastQuoteJson).setValue(JSON.stringify({ type: 'custom', lineItems, markup, total, notes: data.notes || '' }));
  }

  return { success: true };
}

// ── Create PDF for Custom Quote ───────────────────────────────
function createCustomQuotePDF(data, lineItems, markup, total) {
  try {
    const leadId   = data.id || data.leadId || 'DRAFT';
    const name     = data.name || 'Customer';
    const dateStr  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
    const offerDate = new Date(); offerDate.setDate(offerDate.getDate() + 30);
    const offerStr  = Utilities.formatDate(offerDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    const subtotal  = lineItems.reduce(function(s,l){ return s + (parseFloat(l.qty)||1)*(parseFloat(l.unitPrice)||0); }, 0);

    const lineRows = lineItems.map(function(item) {
      const lineTotal = Math.round((parseFloat(item.qty)||1)*(parseFloat(item.unitPrice)||0));
      return '<tr><td>' + item.service +
        (item.description ? '<br><span style="font-size:11px;color:#888;">' + item.description + '</span>' : '') + '</td>' +
        '<td class="c">' + item.qty + ' ' + item.unit + '</td>' +
        '<td class="r">$' + lineTotal.toLocaleString() + '</td></tr>';
    }).join('');

    const html = buildQuoteHtml({
      leadId, name,
      phone:   data.phone   || '',
      email:   data.email   || '',
      address: data.address || '',
      dateStr, offerStr, lineRows, subtotal, markup, total,
      scopeText: '', notes: data.notes || ''
    });

    return saveHtmlToDrive(html, leadId, name);
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
