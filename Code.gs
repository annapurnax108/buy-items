// Paste into Extensions > Apps Script of your Google Sheet. Run setup() once, then Deploy > Web app.
const ITEMS = 'Items', LOG = 'Log';
const I_HEAD = ['id', 'item', 'store', 'category', 'type', 'qty', 'unit', 'price', 'need', 'last_bought', 'times_bought', 'avg_days', 'notes', 'updated_at', 'one_time'];
const L_HEAD = ['id', 'date', 'item_id', 'item', 'store', 'qty', 'price', 'total'];
const STORES = ['Costco', 'HMart', "Trader Joe's", 'NPL/In-store', 'Other']; // keep in sync with STORES in index.html
const TYPES = ['staple', 'occasional'];
const NCOL = I_HEAD.length, UPDATED_COL = 14; // one_time (col 15) is last, so updated_at is no longer the final column
const SHEET_ID = '1ZjIlfJmwzSMag_B1yx41pSzfXrFriOAma3RdotcVT14'; // the buy-list spreadsheet (this script need not be bound to it)
const ss = () => SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();

function setup() {
  const book = ss();
  const it = book.getSheetByName(ITEMS) || book.insertSheet(ITEMS);
  it.getRange(1, 1, 1, NCOL).setValues([I_HEAD]).setFontWeight('bold');
  it.setFrozenRows(1);
  it.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STORES, true).build());
  it.getRange('E2:E').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(TYPES, true).build());
  it.getRange('I2:I').setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  it.getRange('O2:O').setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build()); // one_time: removed from Items once bought
  it.getRange('H2:H').setNumberFormat('$0.00');
  it.getRange('J2:J').setNumberFormat('@'); // last_bought stays plain yyyy-mm-dd text
  const lg = book.getSheetByName(LOG) || book.insertSheet(LOG);
  lg.getRange(1, 1, 1, L_HEAD.length).setValues([L_HEAD]).setFontWeight('bold');
  lg.setFrozenRows(1);
  lg.getRange('B2:B').setNumberFormat('@');
  lg.getRange('G2:H').setNumberFormat('$0.00');
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TOKEN')) props.setProperty('TOKEN', Utilities.getUuid() + Utilities.getUuid());
  Logger.log('Your secret token (paste into the app): ' + props.getProperty('TOKEN'));
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Buy list').addItem('Recalculate frequency stats', 'recomputeAll').addToUi();
}

// Fills id / updated_at for rows edited by hand in Sheets.
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== ITEMS || e.range.getRow() < 2 || e.range.getColumn() > 9) return;
  const r = e.range.getRow();
  if (!sh.getRange(r, 1).getValue() && sh.getRange(r, 2).getValue()) sh.getRange(r, 1).setValue(Utilities.getUuid());
  sh.getRange(r, UPDATED_COL).setValue(new Date().toISOString());
}

function doPost(e) {
  const out = (o) => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'bad json' }); }
  const stored = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!req || !stored || typeof req.token !== 'string' || req.token !== stored) return out({ ok: false, error: 'unauthorized' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const book = ss(), sh = book.getSheetByName(ITEMS), lg = book.getSheetByName(LOG);
    fillIds(sh);
    if (req.action === 'upsert') upsert(sh, req.item);
    else if (req.action === 'delete') remove(sh, req.id);
    else if (req.action === 'buy') buy(sh, lg, req);
    else if (req.action === 'unbuy') unbuy(sh, lg, req);
    return out({ ok: true, items: list(sh) });
  } finally { lock.releaseLock(); }
}

// rows typed straight into the sheet get an id the first time the app talks to it (onEdit only fires for a bound script)
function fillIds(sh) {
  const n = lastItemRow(sh) - 1;
  if (n < 1) return;
  const rng = sh.getRange(2, 1, n, 2), v = rng.getValues();
  let changed = false;
  v.forEach(r => { if (!r[0] && r[1]) { r[0] = Utilities.getUuid(); changed = true; } });
  if (changed) rng.setValues(v);
}

// last row of Items that holds an item. sh.getLastRow() can't be used: the checkbox column is pre-filled down the sheet,
// so it reports row ~1000 and new items would land far below the header
function lastItemRow(sh) {
  const v = sh.getRange(1, 1, sh.getMaxRows(), 2).getValues();
  for (let i = v.length - 1; i >= 1; i--) if (v[i][0] || v[i][1]) return i + 1;
  return 1;
}

const fmtDate = (d) => d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(d || '');

function list(sh) {
  const n = lastItemRow(sh) - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, NCOL).getValues()
    .filter(r => r[0] || r[1])
    .map(r => ({
      id: r[0] || '', item: r[1], store: r[2], category: r[3], type: r[4] || 'staple', qty: r[5] || 1, unit: r[6],
      price: r[7], need: r[8] === true || String(r[8]).toLowerCase() === 'true',
      last_bought: fmtDate(r[9]), times_bought: Number(r[10]) || 0, avg_days: r[11], notes: r[12], updated_at: String(r[13] || ''),
      one_time: r[14] === true || String(r[14]).toLowerCase() === 'true'
    }));
}

function findRow(sh, id) {
  const n = lastItemRow(sh) - 1;
  if (n < 1 || !id) return -1;
  const i = sh.getRange(2, 1, n, 1).getValues().map(r => r[0]).indexOf(id);
  return i < 0 ? -1 : i + 2;
}

// stats columns (last_bought, times_bought, avg_days) belong to the Log, so the app never overwrites them
function upsert(sh, t) {
  const row = findRow(sh, t.id);
  const stats = row < 0 ? ['', '', ''] : sh.getRange(row, 10, 1, 3).getValues()[0];
  const price = t.price === '' || t.price == null ? '' : Number(t.price);
  const vals = [t.id, t.item, t.store, t.category || '', t.type || 'staple', Number(t.qty) || 1, t.unit || '', price, !!t.need,
    stats[0], stats[1], stats[2], t.notes || '', new Date().toISOString(), !!t.one_time];
  const r = row < 0 ? lastItemRow(sh) + 1 : row; // new items go in the first empty row under the last item
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
  sh.getRange(r, 1, 1, NCOL).setValues([vals]);
  sh.getRange(r, 10).setNumberFormat('@').setValue(stats[0]);
  if (row < 0) {
    sh.getRange(r, 3).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STORES, true).build());
    sh.getRange(r, 5).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(TYPES, true).build());
    sh.getRange(r, 9).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    sh.getRange(r, 15).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    sh.getRange(r, 8).setNumberFormat('$0.00');
  }
}

function remove(sh, id) {
  const row = findRow(sh, id);
  if (row > 0) sh.deleteRow(row);
}

// log a purchase (idempotent on log_id, since the app retries queued changes) and take the item off the list
function buy(sh, lg, q) {
  const row = findRow(sh, q.id);
  const n = lg.getLastRow() - 1;
  const seen = n > 0 && lg.getRange(2, 1, n, 1).getValues().some(r => r[0] === q.log_id);
  if (!seen) {
    const qty = Number(q.qty) || 1, price = q.price === '' || q.price == null ? '' : Number(q.price);
    lg.appendRow([q.log_id, q.date, q.id, q.item, q.store, qty, price, price === '' ? '' : qty * price]);
    const r = lg.getLastRow();
    lg.getRange(r, 2).setNumberFormat('@').setValue(q.date);
    lg.getRange(r, 7, 1, 2).setNumberFormat('$0.00');
  }
  if (row < 0) return;
  // a one-off item leaves the catalog once bought (the Log row keeps its name, price and date)
  if (sh.getRange(row, 15).getValue() === true) { sh.deleteRow(row); return; }
  sh.getRange(row, 9).setValue(false); recompute(sh, lg, row);
}

function unbuy(sh, lg, q) {
  const n = lg.getLastRow() - 1;
  if (n > 0) {
    const i = lg.getRange(2, 1, n, 1).getValues().map(r => r[0]).indexOf(q.log_id);
    if (i >= 0) lg.deleteRow(i + 2);
  }
  const row = findRow(sh, q.id);
  if (row > 0) { sh.getRange(row, 9).setValue(true); recompute(sh, lg, row); }
}

const dayNum = (iso) => Date.parse(iso + 'T00:00:00Z') / 86400000;

// last_bought / times_bought / avg_days (mean gap between distinct purchase dates) for one Items row
function recompute(sh, lg, row, logRows) {
  const id = sh.getRange(row, 1).getValue();
  const n = lg.getLastRow() - 1;
  const rows = logRows || (n > 0 ? lg.getRange(2, 1, n, 3).getValues() : []);
  const dates = rows.filter(r => r[2] === id).map(r => fmtDate(r[1])).filter(Boolean).sort();
  const uniq = dates.filter((d, i) => d !== dates[i - 1]);
  const avg = uniq.length > 1 ? Math.round((dayNum(uniq[uniq.length - 1]) - dayNum(uniq[0])) / (uniq.length - 1)) : '';
  sh.getRange(row, 10).setNumberFormat('@');
  sh.getRange(row, 10, 1, 3).setValues([[uniq.length ? uniq[uniq.length - 1] : '', dates.length, avg]]);
}

// menu action: rebuild stats for every item (use after editing or deleting Log rows by hand)
function recomputeAll() {
  const book = ss(), sh = book.getSheetByName(ITEMS), lg = book.getSheetByName(LOG);
  const n = lg.getLastRow() - 1, logRows = n > 0 ? lg.getRange(2, 1, n, 3).getValues() : [];
  for (let r = 2; r <= lastItemRow(sh); r++) if (sh.getRange(r, 1).getValue()) recompute(sh, lg, r, logRows);
}

// one-time fix for items that ended up far below the header: delete the empty rows between the header and the first item
function compactItems() {
  const sh = ss().getSheetByName(ITEMS);
  const v = sh.getRange(1, 1, sh.getMaxRows(), 2).getValues();
  let first = -1;
  for (let i = 1; i < v.length; i++) if (v[i][0] || v[i][1]) { first = i + 1; break; }
  if (first > 2) sh.deleteRows(2, first - 2);
  Logger.log('first item row was ' + first + '; now ' + (first > 2 ? 2 : first));
}
