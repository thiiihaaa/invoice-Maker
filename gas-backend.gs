const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE'; // ← paste your Sheet ID
const DB_TAB   = 'invoices';
const LOG_TAB  = 'Invoice Log';

/* ═══════════════════════════════════
   ROUTING
═══════════════════════════════════ */
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    const user   = verifyToken(body.token);
    if (!user) return res({ error: 'Unauthorized' });

    if (action === 'saveInvoice') return res(saveInvoice(user, body.invoice));
    if (action === 'logInvoice')  return res(saveInvoice(user, body.invoice)); // alias
    if (action === 'deleteInvoice') return res(deleteInvoice(user, body.id));
    return res({ error: 'Unknown action: ' + action });
  } catch (err) {
    return res({ error: err.message });
  }
}

function doGet(e) {
  try {
    const params = e.parameter;
    const user   = verifyToken(params.token);
    if (!user) return res({ error: 'Unauthorized' });

    if (params.action === 'listInvoices') return res({ invoices: listInvoices(user) });
    if (params.action === 'getInvoice')   return res({ invoice: getInvoice(user, params.id) });
    return res({ error: 'Unknown action: ' + params.action });
  } catch (err) {
    return res({ error: err.message });
  }
}

/* ═══════════════════════════════════
   AUTH
═══════════════════════════════════ */
function verifyToken(token) {
  if (!token) return null;
  try {
    const url  = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + token;
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    const info = JSON.parse(resp.getContentText());
    if (info.aud !== '88796990745-40f00aregrbv35t70f9jignl9u3gej16.apps.googleusercontent.com') return null;
    return info.email;
  } catch (_) { return null; }
}

/* ═══════════════════════════════════
   SHEET HELPERS
═══════════════════════════════════ */
function getTab(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh   = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) sh.appendRow(headers);
  }
  return sh;
}

function getDbSheet() {
  return getTab(DB_TAB, ['id','user','data','updatedAt']);
}

function getLogSheet() {
  return getTab(LOG_TAB, ['Invoice #','Client','Date','Due','Items','Subtotal','Discount','Total','Status','Saved At','Sheet Tab']);
}

/* ═══════════════════════════════════
   SAVE INVOICE
   Frontend field names:
   invNum, invDate, invDue, clientName, clientPhone,
   bizName, bizAddr, notes, subtotal, discount,
   grandTotal, status, items[]{desc, qty, price}
═══════════════════════════════════ */
function saveInvoice(user, data) {
  if (!data) return { error: 'No data provided' };

  const ss      = SpreadsheetApp.openById(SHEET_ID);
  const db      = getDbSheet();
  const now     = new Date().toISOString();
  const id      = data.id || data.invNum || ('INV-' + Date.now());
  data.id       = id;
  data.updatedAt = now;

  // Upsert in DB tab
  const rows = db.getDataRange().getValues();
  let found  = false;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id && rows[i][1] === user) {
      db.getRange(i + 1, 3, 1, 2).setValues([[JSON.stringify(data), now]]);
      found = true;
      break;
    }
  }
  if (!found) db.appendRow([id, user, JSON.stringify(data), now]);

  // Create / overwrite dedicated detail sheet tab
  const tabName = makeTabName(id, data.clientName || '');
  createDetailSheet(ss, tabName, data, now);

  // Update Invoice Log summary row
  logSummary(id, data, now, tabName);

  return { success: true, id: id, tab: tabName };
}

/* ── Build a safe tab name (max 31 chars) ── */
function makeTabName(id, client) {
  const safe = (client || 'Client').replace(/[\/\\?\*\[\]:]/g, '').substring(0, 18).trim();
  return (id + ' · ' + safe).substring(0, 31);
}

/* ── Create / overwrite dedicated invoice detail sheet ── */
function createDetailSheet(ss, tabName, d, savedAt) {
  const existing = ss.getSheetByName(tabName);
  if (existing) ss.deleteSheet(existing);
  const sh = ss.insertSheet(tabName);

  const PINK  = '#FF85BB';
  const PALE  = '#FFF0F7';
  const ALT   = '#FFF8FC';
  const TOTAL = '#FFE4F3';

  const subtotal = parseFloat(d.subtotal)   || 0;
  const discount = parseFloat(d.discount)   || 0;
  const total    = parseFloat(d.grandTotal) || 0;

  // ROW 1: Banner
  sh.getRange(1,1,1,6).merge()
    .setValue('INVOICE')
    .setBackground(PINK).setFontColor('#fff')
    .setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 48);

  // ROW 2: Invoice # / Status
  sh.getRange(2,1).setValue('Invoice #').setFontWeight('bold');
  sh.getRange(2,2).setValue(d.invNum || '—').setFontSize(12).setFontWeight('bold');
  sh.getRange(2,4).setValue('Status').setFontWeight('bold');
  sh.getRange(2,5).setValue((d.status||'unpaid').toUpperCase())
    .setFontWeight('bold')
    .setFontColor((d.status||'').toLowerCase()==='paid' ? '#059669' : '#e0609a');

  // ROW 3: Dates
  sh.getRange(3,1).setValue('Invoice Date');
  sh.getRange(3,2).setValue(d.invDate || '—');
  sh.getRange(3,4).setValue('Due Date');
  sh.getRange(3,5).setValue(d.invDue || '—');

  // ROW 4: spacer
  sh.setRowHeight(4, 10);

  // ROW 5: Section headers
  sh.getRange(5,1,1,2).merge().setValue('FROM (Business)').setBackground(PALE).setFontWeight('bold');
  sh.getRange(5,4,1,3).merge().setValue('TO (Client)').setBackground(PALE).setFontWeight('bold');

  // ROWS 6-8: Business / Client info
  sh.getRange(6,1,1,2).merge().setValue(d.bizName  || '—');
  sh.getRange(6,4,1,3).merge().setValue(d.clientName || '—').setFontWeight('bold');
  sh.getRange(7,1,1,2).merge().setValue(d.bizAddr  || '');
  sh.getRange(7,4,1,3).merge().setValue(d.clientPhone || '');

  // ROW 9: spacer
  sh.setRowHeight(9, 10);

  // ROW 10: Items header
  var headers = [['#','Description','Qty','Unit Price','Amount','']];
  sh.getRange(10,1,1,6).setValues(headers)
    .setBackground(PINK).setFontColor('#fff').setFontWeight('bold');
  sh.getRange(10,1).setHorizontalAlignment('center');
  sh.getRange(10,3).setHorizontalAlignment('center');
  sh.getRange(10,4).setHorizontalAlignment('right');
  sh.getRange(10,5).setHorizontalAlignment('right');

  // ROWS 11+: Line items
  var items = d.items || [];
  var r = 11;
  items.forEach(function(item, idx) {
    var bg     = idx % 2 === 0 ? '#ffffff' : ALT;
    var qty    = parseFloat(item.qty)   || 0;
    var price  = parseFloat(item.price) || 0;
    var amount = qty * price;
    sh.getRange(r,1).setValue(idx+1).setBackground(bg).setHorizontalAlignment('center');
    sh.getRange(r,2).setValue(item.desc||'—').setBackground(bg);
    sh.getRange(r,3).setValue(qty).setBackground(bg).setHorizontalAlignment('center');
    sh.getRange(r,4).setValue(price.toFixed(2)).setBackground(bg).setHorizontalAlignment('right');
    sh.getRange(r,5).setValue(amount.toFixed(2)).setBackground(bg).setHorizontalAlignment('right').setFontWeight('bold');
    sh.getRange(r,6).setBackground(bg);
    r++;
  });

  sh.setRowHeight(r, 8); r++;

  // Totals
  function totalRow(label, value, isFinal) {
    var bg  = isFinal ? PINK  : TOTAL;
    var col = isFinal ? '#fff' : '#333';
    sh.getRange(r,3,1,2).merge().setValue(label).setFontWeight('bold')
      .setBackground(bg).setFontColor(col).setHorizontalAlignment('right');
    sh.getRange(r,5).setValue(value).setFontWeight('bold')
      .setBackground(bg).setFontColor(col).setHorizontalAlignment('right');
    if (isFinal) sh.setRowHeight(r, 28);
    r++;
  }

  totalRow('Subtotal', subtotal.toFixed(2));
  if (discount > 0) totalRow('Discount', '- ' + discount.toFixed(2));
  totalRow('TOTAL', total.toFixed(2), true);

  // Notes
  if (d.notes) {
    r++;
    sh.getRange(r,1,1,6).merge().setValue('Notes').setBackground(PALE).setFontWeight('bold');
    r++;
    sh.getRange(r,1,1,6).merge().setValue(d.notes).setWrap(true).setBackground('#fffdf9');
    sh.setRowHeight(r, 60);
    r++;
  }

  // Footer
  r++;
  sh.getRange(r,1,1,6).merge()
    .setValue('Saved: ' + savedAt + '  ·  Kuphoria Invoice Maker')
    .setFontColor('#aaa').setFontSize(9).setHorizontalAlignment('right');

  // Column widths
  sh.setColumnWidth(1, 35);
  sh.setColumnWidth(2, 260);
  sh.setColumnWidth(3, 55);
  sh.setColumnWidth(4, 110);
  sh.setColumnWidth(5, 110);
  sh.setColumnWidth(6, 20);
  sh.setFrozenRows(1);
}

/* ── Summary row in Invoice Log ── */
function logSummary(id, d, savedAt, tabName) {
  const log  = getLogSheet();
  const rows = log.getDataRange().getValues();
  const newRow = [
    d.invNum      || id,
    d.clientName  || '',
    d.invDate     || '',
    d.invDue      || '',
    (d.items||[]).length,
    d.subtotal    || 0,
    d.discount    || 0,
    d.grandTotal  || 0,
    d.status      || 'unpaid',
    savedAt,
    tabName,
  ];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === (d.invNum || id)) {
      log.getRange(i+1, 1, 1, newRow.length).setValues([newRow]);
      return;
    }
  }
  log.appendRow(newRow);
}

/* ═══════════════════════════════════
   LIST / GET / DELETE
═══════════════════════════════════ */
function listInvoices(user) {
  const db   = getDbSheet();
  const rows = db.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === user) {
      try {
        const d = JSON.parse(rows[i][2]);
        list.push({
          id:        rows[i][0],
          updatedAt: rows[i][3],
          invNum:    d.invNum,
          clientName: d.clientName,
          invDate:   d.invDate,
          grandTotal: d.grandTotal,
          status:    d.status,
        });
      } catch (_) {}
    }
  }
  return list.reverse();
}

function getInvoice(user, id) {
  const db   = getDbSheet();
  const rows = db.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id && rows[i][1] === user) {
      return JSON.parse(rows[i][2]);
    }
  }
  return null;
}

function deleteInvoice(user, id) {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const db   = getDbSheet();
  const log  = getLogSheet();
  const rows = db.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id && rows[i][1] === user) {
      let tabName = null;
      try {
        const d = JSON.parse(rows[i][2]);
        tabName = makeTabName(id, d.clientName || '');
      } catch(_) {}

      db.deleteRow(i+1);

      const logRows = log.getDataRange().getValues();
      for (let j = 1; j < logRows.length; j++) {
        if (logRows[j][0] === id) { log.deleteRow(j+1); break; }
      }

      if (tabName) {
        const detailSh = ss.getSheetByName(tabName);
        if (detailSh) ss.deleteSheet(detailSh);
      }

      return { success: true };
    }
  }
  return { error: 'Not found' };
}

/* ═══════════════════════════════════
   RESPONSE HELPER
═══════════════════════════════════ */
function res(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
