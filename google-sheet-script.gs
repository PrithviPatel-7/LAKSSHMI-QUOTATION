// Laksshmi Tiles & Hardwares - Quotation storage for Google Sheets
// 1) Change the password below. 2) Deploy as a Web app (see steps).
const SECRET = 'CHANGE-THIS-PASSWORD';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName('Quotations');
  if (!s) {
    s = ss.insertSheet('Quotations');
    s.appendRow(['Quotation No','Date','Customer','Phone','Address','Sub-total','Transport','GST','Grand Total','Items','Data (do not edit)','Updated']);
    s.setFrozenRows(1);
    s.getRange('A:A').setNumberFormat('@');
  }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e) {
  if (e.parameter.key !== SECRET) return out_({error: 'denied'});
  const v = sheet_().getDataRange().getValues(), list = [];
  for (let i = 1; i < v.length; i++) { try { list.push(JSON.parse(v[i][10])); } catch (x) {} }
  return out_({quotations: list});
}
function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.key !== SECRET) return out_({error: 'denied'});
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const s = sheet_(), last = s.getLastRow();
    const nos = last > 1 ? s.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
    if (b.action === 'delete') {
      const i = nos.indexOf(String(b.no));
      if (i >= 0) s.deleteRow(i + 2);
      return out_({ok: true});
    }
    const q = b.quote;
    const items = q.items.map(i => i.desc + ' - ' + i.qty + ' ' + i.unit + ' x ' + i.rate + (i.dr ? ' (disc ' + i.dr + ')' : '')).join('; ');
    const row = [q.no, q.date, q.cname, q.cphone, q.caddr, q.sub, q.trans, q.gstOn ? q.gstAmt : 'No GST', q.total, items, JSON.stringify(q), new Date()];
    const i = nos.indexOf(String(q.no));
    if (i >= 0) s.getRange(i + 2, 1, 1, row.length).setValues([row]); else s.appendRow(row);
    return out_({ok: true});
  } finally { lock.releaseLock(); }
}
