/**
 * Dokan Khata on Google Apps Script.
 * - Serves the app (Index.html) as a web app.
 * - Keeps every shop's products, sales, income/expenses and settings in this Google Sheet.
 * - Checks access codes against the "Codes" sheet, so a code can be switched off or extended here.
 * - Saves plan payments (bKash, Nagad, Rocket) to the "Payments" sheet and emails the owner.
 * Bind this script to a Google Sheet (Extensions > Apps Script) and deploy it as a web app.
 */

var PLANS = { basic: 'Shuru (৳350)', standard: 'Byabsa (৳500)', pro: 'Pro (৳1000)' };
var SHEETS = {
  codes:    ['Code', 'Business', 'Phone', 'Plan', 'Expires', 'Status', 'Created', 'Note'],
  payments: ['Received', 'Plan', 'Method', 'Transaction ID', 'Shop', 'Phone', 'Status', 'Code'],
  data:     ['Code', 'Store', 'ID', 'Updated', 'JSON']
};
var SHEET_NAMES = { codes: 'Codes', payments: 'Payments', data: 'Data' };
var STORES = ['products', 'sales', 'entries', 'meta'];

/* ---------- web app ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Dokan Khata')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setFaviconUrl('https://shehransufyani.github.io/Dokan-Khata/icons/icon-192.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- sheet helpers ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(key) {
  var ss = ss_(), name = SHEET_NAMES[key], sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SHEETS[key].length).setValues([SHEETS[key]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function setup() { Object.keys(SHEETS).forEach(sheet_); }
function norm_(code) { return String(code || '').trim().toUpperCase(); }

/* ---------- access codes ---------- */
function findCode_(code) {
  code = norm_(code);
  if (!code) return null;
  var sh = sheet_('codes'), last = sh.getLastRow();
  if (last < 2) return null;
  var rows = sh.getRange(2, 1, last - 1, 8).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (norm_(rows[i][0]) === code) {
      var r = rows[i], exp = r[4] instanceof Date ? r[4].getTime() : new Date(r[4]).getTime();
      return { row: i + 2, code: code, biz: String(r[1] || ''), tier: String(r[3] || '').toLowerCase(),
               exp: exp, active: String(r[5] || 'active').toLowerCase() === 'active' };
    }
  }
  return null;
}
/** Called by the app. Returns the plan for a code, or null. Expired codes come back with their end date so the app can say so. */
function checkCode(code) {
  var c = findCode_(code);
  if (!c || !c.active || !PLANS[c.tier] || !c.exp) return null;
  return { id: c.code, biz: c.biz, tier: c.tier, exp: c.exp, code: c.code, server: true };
}
function usable_(code) { var c = checkCode(code); return c && c.exp >= Date.now() ? c : null; }

function makeCode_() {
  var chars = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789', s = 'DK-';
  for (var i = 0; i < 8; i++) { if (i === 4) s += '-'; s += chars.charAt(Math.floor(Math.random() * chars.length)); }
  return findCode_(s) ? makeCode_() : s;
}
function addCode_(biz, phone, tier, months, note) {
  var exp = new Date(); exp.setMonth(exp.getMonth() + months); exp.setHours(23, 59, 59, 0);
  var code = makeCode_();
  sheet_('codes').appendRow([code, biz, phone || '', tier, exp, 'active', new Date(), note || '']);
  return { code: code, exp: exp, link: appLink_(code) };
}
function appLink_(code) {
  var url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  return url ? url + '?code=' + encodeURIComponent(code) : '(deploy the web app first to get a link)';
}

/* ---------- shop data ---------- */
/** Returns everything stored for this code: {products:[], sales:[], entries:[], meta:[]} */
function loadAll(code) {
  var c = checkCode(code);
  if (!c) throw new Error('Access code not valid');
  var out = { products: [], sales: [], entries: [], meta: [] };
  var sh = sheet_('data'), last = sh.getLastRow();
  if (last < 2) return out;
  var rows = sh.getRange(2, 1, last - 1, 5).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (norm_(rows[i][0]) !== c.code || !out[rows[i][1]]) continue;
    try { out[rows[i][1]].push(JSON.parse(rows[i][4])); } catch (e) {}
  }
  return out;
}
/** ops: [{op:'put'|'del'|'clear', store, id, obj}] applied in order. Writes need a code that has not expired. */
function saveBatch(code, ops) {
  var c = usable_(code);
  if (!c) throw new Error('Access code not valid or plan ended');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_('data'), last = sh.getLastRow();
    var keys = last > 1 ? sh.getRange(2, 1, last - 1, 3).getValues() : [];
    var index = {}, deleteRows = [], appends = [];
    for (var i = 0; i < keys.length; i++) index[norm_(keys[i][0]) + '|' + keys[i][1] + '|' + keys[i][2]] = i + 2;
    (ops || []).forEach(function (o) {
      if (STORES.indexOf(o.store) < 0) return;
      if (o.op === 'clear') {
        Object.keys(index).forEach(function (k) {
          if (k.indexOf(c.code + '|' + o.store + '|') !== 0) return;
          if (typeof index[k] === 'string') appends[+index[k].slice(1)] = null; else deleteRows.push(index[k]);
          delete index[k];
        });
        return;
      }
      var id = String(o.op === 'put' ? (o.obj && o.obj.id) : o.id), key = c.code + '|' + o.store + '|' + id;
      if (o.op === 'del') {
        if (typeof index[key] === 'string') appends[+index[key].slice(1)] = null; else if (index[key]) deleteRows.push(index[key]);
        delete index[key]; return;
      }
      var json = JSON.stringify(o.obj);
      if (json.length > 49000 && o.obj.photo) { var copy = JSON.parse(json); copy.photo = ''; json = JSON.stringify(copy); }
      var row = [c.code, o.store, id, new Date(), json];
      var at = index[key];
      if (typeof at === 'string') appends[+at.slice(1)] = row;          // appended earlier in this batch
      else if (at) sh.getRange(at, 1, 1, 5).setValues([row]);
      else { index[key] = 'a' + appends.length; appends.push(row); }
    });
    appends = appends.filter(function (r) { return r; });
    if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, 5).setValues(appends);
    deleteRows.sort(function (a, b) { return b - a; }).forEach(function (r) { if (r > 1) sh.deleteRow(r); });
  } finally { lock.releaseLock(); }
  return true;
}

/* ---------- plan payments ---------- */
function submitPayment(p) {
  p = p || {};
  var tier = String(p.tier || ''), method = String(p.method || '');
  if (!PLANS[tier] || ['bkash', 'nagad', 'rocket'].indexOf(method) < 0) throw new Error('Choose a plan and a payment method');
  var trx = String(p.trx || '').slice(0, 60), shop = String(p.shop || '').slice(0, 120), phone = String(p.phone || '').slice(0, 30);
  if (!trx || !shop || !phone) throw new Error('Transaction ID, shop name and phone are needed');
  sheet_('payments').appendRow([new Date(), tier, method, trx, shop, phone, 'waiting', '']);
  try {
    var owner = Session.getEffectiveUser().getEmail();
    if (owner) MailApp.sendEmail(owner, 'Dokan Khata payment: ' + shop + ' (' + PLANS[tier] + ')',
      'New plan payment to check.\n\nPlan: ' + PLANS[tier] + '\nPaid with: ' + method + '\nTransaction ID: ' + trx +
      '\nShop: ' + shop + '\nPhone: ' + phone + '\n\nCheck it in your ' + method + ' app, then open the Sheet, select the row in "Payments" and use Dokan Khata > Approve selected payment.\n' + ss_().getUrl());
  } catch (e) {}
  return true;
}

/* ---------- owner menu in the Sheet ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Dokan Khata')
    .addItem('Create access code', 'menuCreateCode')
    .addItem('Approve selected payment', 'menuApprovePayment')
    .addItem('Create my master code', 'menuMasterCode')
    .addSeparator()
    .addItem('Show app link', 'menuAppLink')
    .addItem('Set up sheets', 'setup')
    .addToUi();
}
function ask_(ui, title, prompt) {
  var r = ui.prompt(title, prompt, ui.ButtonSet.OK_CANCEL);
  return r.getSelectedButton() === ui.Button.OK ? r.getResponseText().trim() : null;
}
function askPlan_(ui) {
  var t = ask_(ui, 'Plan', 'Type 1 for Shuru (৳350), 2 for Byabsa (৳500) or 3 for Pro (৳1000)');
  return t === null ? null : ({ '1': 'basic', '2': 'standard', '3': 'pro' })[t] || null;
}
function showCode_(ui, res, biz) {
  ui.alert('Code created',
    'Business: ' + biz + '\nCode: ' + res.code + '\nEnds: ' + Utilities.formatDate(res.exp, Session.getScriptTimeZone(), 'd MMM yyyy') +
    '\n\nSend this link on WhatsApp:\n' + res.link, ui.ButtonSet.OK);
}
function menuCreateCode() {
  var ui = SpreadsheetApp.getUi();
  var biz = ask_(ui, 'New access code', 'Business name'); if (!biz) return;
  var phone = ask_(ui, 'New access code', 'Their WhatsApp number (optional)'); if (phone === null) return;
  var tier = askPlan_(ui); if (!tier) return;
  var months = parseInt(ask_(ui, 'Length', 'How many months? (1, 3, 6 or 12)'), 10) || 1;
  showCode_(ui, addCode_(biz, phone, tier, months, ''), biz);
}
function menuApprovePayment() {
  var ui = SpreadsheetApp.getUi(), sh = ss_().getActiveSheet();
  if (sh.getName() !== SHEET_NAMES.payments || sh.getActiveRange().getRow() < 2) { ui.alert('Open the Payments sheet and click a payment row first.'); return; }
  var r = sh.getActiveRange().getRow(), v = sh.getRange(r, 1, 1, 8).getValues()[0];
  if (v[7]) { ui.alert('This payment already has a code: ' + v[7]); return; }
  var months = parseInt(ask_(ui, 'Approve ' + v[4], 'Payment checked? How many months to give? (usually 1)'), 10);
  if (!months) return;
  var res = addCode_(String(v[4]), String(v[5]), String(v[1]), months, 'TrxID ' + v[3] + ' via ' + v[2]);
  sh.getRange(r, 7, 1, 2).setValues([['approved', res.code]]);
  showCode_(ui, res, String(v[4]));
}
function menuMasterCode() {
  var ui = SpreadsheetApp.getUi();
  var exp = new Date(2099, 11, 31, 23, 59, 59), code = makeCode_();
  sheet_('codes').appendRow([code, 'Owner (master)', '', 'pro', exp, 'active', new Date(), 'Master code, keep private']);
  showCode_(ui, { code: code, exp: exp, link: appLink_(code) }, 'Owner (master)');
}
function menuAppLink() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('App link', appLink_('').replace('?code=', '') || 'Deploy the web app first (Deploy > New deployment > Web app).', ui.ButtonSet.OK);
}
