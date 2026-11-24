/**
 * Store Flow on Google Apps Script.
 * - Serves the app (Index.html) as a web app.
 * - Shop owners sign up with an email or phone number and a password ("Accounts" sheet).
 * - Keeps every shop's products, sales, income/expenses and settings in this Google Sheet.
 * - Checks access codes against the "Codes" sheet, so a code can be switched off or extended here.
 * - Saves plan payments (bKash, Nagad, Rocket) to the "Payments" sheet and emails the owner.
 * Bind this script to a Google Sheet (Extensions > Apps Script) and deploy it as a web app.
 */

var PLANS = { basic: 'Shuru (৳350)', standard: 'Bepsha (৳500)', pro: 'Pro (৳1000)' };
var SHEETS = {
  codes:    ['Code', 'Business', 'Phone', 'Plan', 'Expires', 'Status', 'Created', 'Note', 'Account'],
  payments: ['Paid at', 'Package', 'Method', 'Transaction ID', 'Store name', 'Phone', 'Status', 'Access code', 'Account'],
  data:     ['Code', 'Store', 'ID', 'Updated', 'JSON'],
  accounts: ['ID', 'Login', 'Shop', 'Created', 'Trial started', 'Code', 'Last login', 'Hash', 'Salt', 'Reset']
};
var SHEET_NAMES = { codes: 'Codes', payments: 'Payments', data: 'Data', accounts: 'Accounts' };
var TRIAL_MS = 7 * 24 * 60 * 60 * 1000; // 7-day free trial
var STORES = ['products', 'sales', 'entries', 'orders', 'meta'];
var PRICES = { basic: 350, standard: 500, pro: 1000 };

/* ---------- web app ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Store Flow')
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
  } else if (sh.getRange(1, 1, 1, SHEETS[key].length).getValues()[0].join('|') !== SHEETS[key].join('|')) {
    sh.getRange(1, 1, 1, SHEETS[key].length).setValues([SHEETS[key]]).setFontWeight('bold');
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
  var rows = sh.getRange(2, 1, last - 1, 9).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (norm_(rows[i][0]) === code) {
      var r = rows[i], exp = r[4] instanceof Date ? r[4].getTime() : new Date(r[4]).getTime();
      return { row: i + 2, code: code, biz: String(r[1] || ''), tier: String(r[3] || '').toLowerCase(),
               exp: exp, active: String(r[5] || 'active').toLowerCase() === 'active',
               master: /^master/i.test(String(r[7] || '')), account: String(r[8] || '') };
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
function addCode_(biz, phone, tier, months, note, account) {
  var exp = new Date(); exp.setMonth(exp.getMonth() + months); exp.setHours(23, 59, 59, 0);
  var code = makeCode_();
  sheet_('codes').appendRow([code, biz, phone ? "'" + phone : '', tier, exp, 'active', new Date(), note || '', account || '']);
  return { code: code, exp: exp, link: appLink_(code) };
}
function appLink_(code) {
  var url = PropertiesService.getScriptProperties().getProperty('APP_URL') || '';
  if (!url) { try { url = ScriptApp.getService().getUrl(); } catch (e) {} }
  return url ? url + '?code=' + encodeURIComponent(code) : '(deploy the web app first to get a link)';
}

/* ---------- shop data ---------- */
/** Returns everything stored for this account: {products:[], sales:[], entries:[], meta:[]} */
function loadAll(token) {
  var c = { code: auth_(token).id };
  var out = { products: [], sales: [], entries: [], orders: [], meta: [] };
  var sh = sheet_('data'), last = sh.getLastRow();
  if (last < 2) return out;
  var rows = sh.getRange(2, 1, last - 1, 5).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (norm_(rows[i][0]) !== c.code || !out[rows[i][1]]) continue;
    try { out[rows[i][1]].push(JSON.parse(rows[i][4])); } catch (e) {}
  }
  return out;
}
/** ops: [{op:'put'|'del'|'clear', store, id, obj}] applied in order. Writes need a running trial or a plan that has not ended. */
function saveBatch(token, ops) {
  var a = auth_(token);
  if (!trialOn_(a) && !usable_(a.code)) throw new Error('LOCKED');
  var c = { code: a.id };
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
function submitPayment(token, p) {
  var a = auth_(token);
  p = p || {};
  var tier = String(p.tier || ''), method = String(p.method || '');
  if (!PLANS[tier] || ['bkash', 'nagad', 'rocket'].indexOf(method) < 0) throw new Error('Choose a plan and a payment method');
  var trx = String(p.trx || '').slice(0, 60), shop = String(p.shop || '').slice(0, 120), phone = String(p.phone || '').slice(0, 30);
  if (!trx || !shop || !phone) throw new Error('Transaction ID, shop name and phone are needed');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  var res;
  try {
    var ps = sheet_('payments'), last = ps.getLastRow();
    var used = last > 1 ? ps.getRange(2, 4, last - 1, 1).getValues() : [];
    for (var i = 0; i < used.length; i++) if (norm_(used[i][0]) === norm_(trx)) throw new Error('DUP_TRX');
    // A unique code for this payment, already tied to the paying account and plan. Send it once the money has arrived.
    res = addCode_(shop, phone, tier, 1, 'TrxID ' + trx + ' via ' + method + ' for ' + a.login, a.id);
    ps.appendRow([new Date(), PLANS[tier], method, "'" + trx, shop, "'" + phone, 'check money, then send code', res.code, a.login]);
  } finally { lock.releaseLock(); }
  try {
    var owner = Session.getEffectiveUser().getEmail();
    if (owner) MailApp.sendEmail(owner, 'Store Flow payment: ' + shop + ' (' + PLANS[tier] + ')',
      'New plan payment to check.\n\nPlan: ' + PLANS[tier] + '\nPaid with: ' + method + '\nTransaction ID: ' + trx +
      '\nShop: ' + shop + '\nPhone: ' + phone + '\nAccount: ' + a.login + '\n\nAccess code for this payment: ' + res.code + '\n\nCheck the money arrived in your ' + method + ' app, then send this code to ' + phone + '. If the money did not arrive, set the code to "off" in the Codes sheet.\n' + ss_().getUrl());
  } catch (e) {}
  return true;
}

/* ---------- owner menu in the Sheet ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Store Flow')
    .addItem('Create access code', 'menuCreateCode')
    .addItem('Show code for selected payment', 'menuApprovePayment')
    .addItem('Create my master code', 'menuMasterCode')
    .addItem('Reset password of selected account', 'menuResetPassword')
    .addSeparator()
    .addItem('Update owner dashboard', 'menuDashboard')
    .addItem('Send renewal reminders now', 'menuReminders')
    .addSeparator()
    .addItem('Show app link', 'menuAppLink')
    .addItem('Set app link', 'menuSetAppLink')
    .addItem('Set up sheets', 'setup')
    .addToUi();
}
function ask_(ui, title, prompt) {
  var r = ui.prompt(title, prompt, ui.ButtonSet.OK_CANCEL);
  return r.getSelectedButton() === ui.Button.OK ? r.getResponseText().trim() : null;
}
function askPlan_(ui) {
  var t = ask_(ui, 'Plan', 'Type 1 for Shuru (৳350), 2 for Bepsha (৳500) or 3 for Pro (৳1000)');
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
  var r = sh.getActiveRange().getRow(), v = sh.getRange(r, 1, 1, 9).getValues()[0];
  if (v[7]) { sh.getRange(r, 7).setValue('code sent'); ui.alert('Code for ' + v[4], 'Send this code to ' + v[5] + ':\n\n' + v[7] + '\n\nThe row is now marked "code sent".', ui.ButtonSet.OK); return; }
  var months = parseInt(ask_(ui, 'Approve ' + v[4], 'Payment checked? How many months to give? (usually 1)'), 10);
  if (!months) return;
  var tier = PLANS[String(v[1])] ? String(v[1]) : Object.keys(PLANS).filter(function (k) { return PLANS[k] === String(v[1]); })[0];
  if (!tier) { ui.alert('Unknown package in this row: ' + v[1]); return; }
  var res = addCode_(String(v[4]), String(v[5]), tier, months, 'TrxID ' + v[3] + ' via ' + v[2] + (v[8] ? ' for ' + v[8] : ''));
  sh.getRange(r, 7, 1, 2).setValues([['approved', res.code]]);
  showCode_(ui, res, String(v[4]));
}
function menuMasterCode() {
  var ui = SpreadsheetApp.getUi();
  var exp = new Date(2099, 11, 31, 23, 59, 59), code = makeCode_();
  sheet_('codes').appendRow([code, 'Owner (master)', '', 'pro', exp, 'active', new Date(), 'Master code, keep private']);
  showCode_(ui, { code: code, exp: exp, link: appLink_(code) }, 'Owner (master)');
}
function menuSetAppLink() {
  var ui = SpreadsheetApp.getUi();
  var url = ask_(ui, 'Set app link', 'Paste the Web app URL from Deploy > Manage deployments (it ends in /exec).');
  if (!url) return;
  if (!/^https:\/\/script\.google\.com\/.*\/exec$/.test(url)) { ui.alert('That does not look like a Web app URL ending in /exec.'); return; }
  PropertiesService.getScriptProperties().setProperty('APP_URL', url);
  ui.alert('Saved. New access codes will use this link.');
}
function menuAppLink() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('App link', appLink_('').replace('?code=', '') || 'Deploy the web app first (Deploy > New deployment > Web app).', ui.ButtonSet.OK);
}


/* ---------- accounts: sign up, log in, trial, codes ---------- */
function secret_() {
  var pr = PropertiesService.getScriptProperties(), k = pr.getProperty('TOKEN_SECRET');
  if (!k) { k = Utilities.getUuid() + Utilities.getUuid(); pr.setProperty('TOKEN_SECRET', k); }
  return k;
}
function b64_(bytes) { return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, ''); }
function hash_(pw, salt) {
  var h = String(pw);
  for (var i = 0; i < 150; i++) h = b64_(Utilities.computeHmacSha256Signature(h, salt));
  return h;
}
function cleanLogin_(raw) {
  var s = String(raw || '').trim().toLowerCase();
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s)) return s;
  var d = s.replace(/[^\d]/g, '');
  if (d.indexOf('880') === 0) d = d.slice(2);
  if (/^01[3-9]\d{8}$/.test(d)) return d;
  return null;
}
function findAccount_(by, value) {
  var sh = sheet_('accounts'), last = sh.getLastRow();
  if (last < 2) return null;
  var rows = sh.getRange(2, 1, last - 1, 10).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(by === 'id' ? r[0] : r[1]) !== String(value)) continue;
    var ts = r[4] instanceof Date ? r[4].getTime() : (r[4] ? new Date(r[4]).getTime() : 0);
    return { row: i + 2, id: String(r[0]), login: String(r[1]), store: String(r[2] || ''), trialStart: ts || 0,
             code: norm_(r[5]), hash: String(r[7]), salt: String(r[8]), reset: String(r[9] || '') };
  }
  return null;
}
function token_(id) {
  var exp = Date.now() + 180 * 864e5, body = id + '.' + exp;
  return body + '.' + b64_(Utilities.computeHmacSha256Signature(body, secret_()));
}
function auth_(token) {
  var p = String(token || '').split('.');
  if (p.length !== 3 || +p[1] < Date.now()) throw new Error('AUTH');
  if (b64_(Utilities.computeHmacSha256Signature(p[0] + '.' + p[1], secret_())) !== p[2]) throw new Error('AUTH');
  var a = findAccount_('id', p[0]);
  if (!a) throw new Error('AUTH');
  return a;
}
function trialOn_(a) { return a.trialStart && Date.now() < a.trialStart + TRIAL_MS + 60000; }
function state_(a, token) {
  return { token: token || null, id: a.id, login: a.login, store: a.store, trialStart: a.trialStart,
           access: a.code ? checkCode(a.code) : null, now: Date.now() };
}
function tooMany_(login) {
  var c = CacheService.getScriptCache(), k = 'fail:' + login, n = +(c.get(k) || 0);
  return n >= 8;
}
function failed_(login) {
  var c = CacheService.getScriptCache(), k = 'fail:' + login;
  c.put(k, String(+(c.get(k) || 0) + 1), 900);
}
/** Called by the app. */
function signUp(p) {
  p = p || {};
  var login = cleanLogin_(p.login), pw = String(p.password || '');
  if (!login) throw new Error('LOGIN_FORMAT');
  if (pw.length < 6) throw new Error('PASSWORD_SHORT');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (findAccount_('login', login)) throw new Error('EXISTS');
    var id = 'A-' + Utilities.getUuid().slice(0, 8).toUpperCase(), salt = Utilities.getUuid();
    sheet_('accounts').appendRow([id, login, '', new Date(), '', '', new Date(), hash_(pw, salt), salt, '']);
  } finally { lock.releaseLock(); }
  var a = findAccount_('id', id);
  return state_(a, token_(id));
}
function logIn(p) {
  p = p || {};
  var login = cleanLogin_(p.login);
  if (!login) throw new Error('LOGIN_FORMAT');
  if (tooMany_(login)) throw new Error('TOO_MANY');
  var a = findAccount_('login', login);
  if (!a || hash_(String(p.password || ''), a.salt) !== a.hash) { failed_(login); throw new Error('WRONG'); }
  sheet_('accounts').getRange(a.row, 7).setValue(new Date());
  return state_(a, token_(a.id));
}
function session(token) { var s = state_(auth_(token)); daily_(); return s; }
function setStore(token, name) {
  var a = auth_(token); name = String(name || '').trim().slice(0, 80);
  if (!name) throw new Error('STORE');
  sheet_('accounts').getRange(a.row, 3).setValue(name);
  a.store = name; return state_(a);
}
function startTrial(token) {
  var a = auth_(token);
  if (!a.trialStart) { a.trialStart = Date.now(); sheet_('accounts').getRange(a.row, 5).setValue(new Date(a.trialStart)); }
  return state_(a);
}
function redeemCode(token, code) {
  var a = auth_(token), c = findCode_(code);
  if (!c || !c.active || !PLANS[c.tier] || !c.exp) throw new Error('BAD_CODE');
  if (c.account && c.account !== a.id && !c.master) throw new Error('CODE_USED');
  if (!c.account && !c.master) sheet_('codes').getRange(c.row, 9).setValue(a.id);
  sheet_('accounts').getRange(a.row, 6).setValue(c.code);
  if (!a.store && c.biz && !c.master) sheet_('accounts').getRange(a.row, 3).setValue(c.biz);
  return checkCode(c.code);
}
function resetStart(raw) {
  var login = cleanLogin_(raw);
  if (!login) throw new Error('LOGIN_FORMAT');
  if (login.indexOf('@') < 0) return { via: 'owner' };
  var a = findAccount_('login', login);
  if (a) {
    var code = String(Math.floor(100000 + Math.random() * 900000));
    sheet_('accounts').getRange(a.row, 10).setValue(hash_(code, a.salt) + '|' + (Date.now() + 30 * 60000));
    MailApp.sendEmail(login, 'Store Flow password reset code', 'Your Store Flow reset code is ' + code + '\n\nIt works for 30 minutes. If you did not ask for this, ignore this email.');
  }
  return { via: 'email' };
}
function resetFinish(p) {
  p = p || {};
  var login = cleanLogin_(p.login), pw = String(p.password || '');
  if (!login) throw new Error('LOGIN_FORMAT');
  if (pw.length < 6) throw new Error('PASSWORD_SHORT');
  if (tooMany_(login)) throw new Error('TOO_MANY');
  var a = findAccount_('login', login), r = a ? a.reset.split('|') : [];
  if (!a || r.length !== 2 || +r[1] < Date.now() || hash_(String(p.code || '').trim(), a.salt) !== r[0]) { failed_(login); throw new Error('BAD_RESET'); }
  var salt = Utilities.getUuid();
  sheet_('accounts').getRange(a.row, 8, 1, 3).setValues([[hash_(pw, salt), salt, '']]);
  return state_(a, token_(a.id));
}
function menuResetPassword() {
  var ui = SpreadsheetApp.getUi(), sh = ss_().getActiveSheet();
  if (sh.getName() !== SHEET_NAMES.accounts || sh.getActiveRange().getRow() < 2) { ui.alert('Open the Accounts sheet and click the account row first.'); return; }
  var r = sh.getActiveRange().getRow(), login = sh.getRange(r, 2).getValue();
  var chars = 'abcdefghjkmnpqrstuvwxyz23456789', pw = '';
  for (var i = 0; i < 8; i++) pw += chars.charAt(Math.floor(Math.random() * chars.length));
  var salt = Utilities.getUuid();
  sh.getRange(r, 8, 1, 3).setValues([[hash_(pw, salt), salt, '']]);
  ui.alert('Password reset', 'Account: ' + login + '\nNew password: ' + pw + '\n\nSend this to the shop owner. They can log in with it straight away.', ui.ButtonSet.OK);
}


/* ---------- owner dashboard and renewal reminders ---------- */
function tz_() { return Session.getScriptTimeZone() || 'Asia/Dhaka'; }
function fmt_(d) { return Utilities.formatDate(new Date(d), tz_(), 'd MMM yyyy'); }
function waUrl_(phone, text) {
  var d = String(phone || '').replace(/[^\d]/g, '');
  if (d.indexOf('01') === 0) d = '88' + d;
  else if (/^1\d{9}$/.test(d)) d = '880' + d; // the Sheet drops the leading 0 of numbers like 01712345678
  if (!/^8801\d{9}$/.test(d)) return '';
  return 'https:' + '//api.whatsapp.com/send?phone=' + d + '&text=' + encodeURIComponent(text);
}
/** A Bangladesh mobile number as people write it (01XXXXXXXXX), even when the Sheet stored it as a number. */
function phone_(v) {
  var d = String(v || '').replace(/[^\d]/g, '');
  if (/^1\d{9}$/.test(d)) return '0' + d;
  if (/^8801\d{9}$/.test(d)) return d.slice(2);
  return String(v || '');
}
function waCell_(phone, text) { var u = waUrl_(phone, text); return u ? '=HYPERLINK("' + u.replace(/"/g, '""') + '","Send on WhatsApp")' : ''; }
function rows_(key, cols) { var sh = sheet_(key), last = sh.getLastRow(); return last > 1 ? sh.getRange(2, 1, last - 1, cols).getValues() : []; }
function price_(pkg) { var m = String(pkg || '').replace(/,/g, '').match(/(\d{3,5})/); if (m) return +m[1]; var k = String(pkg || '').toLowerCase(); return PRICES[k] || 0; }

/** Collects everything the owner needs: shops, plans ending, trials ending, money. */
function ownerData_() {
  var now = Date.now(), day = 864e5;
  var accounts = rows_('accounts', 10).map(function (r) {
    var ts = r[4] instanceof Date ? r[4].getTime() : (r[4] ? new Date(r[4]).getTime() : 0);
    return { id: String(r[0]), login: String(r[1]).indexOf('@') < 0 ? phone_(r[1]) : String(r[1]), shop: String(r[2] || ''), created: r[3], trialStart: ts, code: norm_(r[5]), last: r[6] };
  });
  var byId = {}; accounts.forEach(function (a) { byId[a.id] = a; });
  var codes = rows_('codes', 9).map(function (r) {
    var exp = r[4] instanceof Date ? r[4].getTime() : new Date(r[4]).getTime();
    return { code: norm_(r[0]), biz: String(r[1] || ''), phone: phone_(r[2]), tier: String(r[3] || '').toLowerCase(), exp: exp,
             active: String(r[5] || 'active').toLowerCase() === 'active', master: /^master/i.test(String(r[7] || '')), account: String(r[8] || '') };
  });
  var paying = [], ending = [];
  accounts.forEach(function (a) {
    var c = null; codes.forEach(function (x) { if (x.code === a.code) c = x; });
    if (!c || !c.active || c.master || !(c.exp > now)) return;
    var item = { shop: a.shop || c.biz, login: a.login, phone: c.phone || (a.login.indexOf('@') < 0 ? a.login : ''), tier: c.tier, exp: c.exp, code: c.code, last: a.last, daysLeft: Math.ceil((c.exp - now) / day) };
    paying.push(item);
    if (c.exp - now <= 7 * day) ending.push(item);
  });
  var trials = [], trialsEnding = [];
  accounts.forEach(function (a) {
    if (!a.trialStart) return;
    var hasPlan = paying.some(function (p) { return p.login === a.login; });
    var end = a.trialStart + TRIAL_MS;
    if (hasPlan || end < now) return;
    var item = { shop: a.shop, login: a.login, phone: a.login.indexOf('@') < 0 ? a.login : '', end: end, hoursLeft: Math.ceil((end - now) / 36e5) };
    trials.push(item);
    if (end - now <= 2 * day) trialsEnding.push(item);
  });
  var month = new Date(); month.setDate(1); month.setHours(0, 0, 0, 0);
  var revMonth = 0, revAll = 0, waiting = 0, byPlan = {};
  rows_('payments', 9).forEach(function (r) {
    var st = String(r[6] || '').toLowerCase(), amt = price_(r[1]);
    if (st.indexOf('check') === 0) { waiting++; return; }
    if (st === 'off' || st === 'refunded' || st === 'rejected') return;
    revAll += amt;
    var at = r[0] instanceof Date ? r[0].getTime() : new Date(r[0]).getTime();
    if (at >= month.getTime()) { revMonth += amt; byPlan[r[1]] = (byPlan[r[1]] || 0) + amt; }
  });
  paying.sort(function (a, b) { return a.exp - b.exp; });
  ending.sort(function (a, b) { return a.exp - b.exp; });
  trialsEnding.sort(function (a, b) { return a.end - b.end; });
  return { accounts: accounts, paying: paying, ending: ending, trials: trials, trialsEnding: trialsEnding, revMonth: revMonth, revAll: revAll, waiting: waiting, byPlan: byPlan };
}
function renewMsg_(p) {
  return 'Assalamu alaikum! Your Store Flow ' + (PLANS[p.tier] || p.tier) + ' plan for ' + p.shop + ' ends on ' + fmt_(p.exp) +
    '. To keep using it, send the plan fee by bKash, Nagad or Rocket and enter the Transaction ID in the app (Plans page). Thank you!';
}
function trialMsg_(t) {
  return 'Assalamu alaikum! Your Store Flow free trial for ' + (t.shop || 'your shop') + ' ends on ' + fmt_(t.end) +
    '. Choose a plan in the app (Plans page) and pay by bKash, Nagad or Rocket to keep everything you added. Thank you!';
}
/** Writes the "Dashboard" tab: totals, plans ending, trials ending, all paying shops. */
function buildDashboard_() {
  var d = ownerData_(), ss = ss_(), sh = ss.getSheetByName('Dashboard');
  if (!sh) { sh = ss.insertSheet('Dashboard', 0); }
  sh.clear(); sh.clearConditionalFormatRules();
  var out = [], fmtRows = [];
  function row(a) { while (a.length < 7) a.push(''); out.push(a); }
  function head(title) { row(['']); fmtRows.push(['h', out.length + 1]); row([title]); }
  function cols(a) { fmtRows.push(['c', out.length + 1]); row(a); }
  row(['Store Flow: owner dashboard', '', '', '', '', '', 'Updated ' + Utilities.formatDate(new Date(), tz_(), 'd MMM yyyy, h:mm a')]);
  row(['']);
  cols(['Paying shops', 'On free trial', 'All accounts', 'Revenue this month (৳)', 'Revenue all time (৳)', 'Payments to check', 'Plans ending in 7 days']);
  row([d.paying.length, d.trials.length, d.accounts.length, d.revMonth, d.revAll, d.waiting, d.ending.length]);
  head('Plans ending in the next 7 days (send a reminder)');
  cols(['Shop', 'Login', 'Plan', 'Ends', 'Days left', 'Phone', 'Remind']);
  if (!d.ending.length) row(['None right now.']);
  d.ending.forEach(function (p) { row([p.shop, p.login, PLANS[p.tier] || p.tier, fmt_(p.exp), p.daysLeft, p.phone ? "'" + p.phone : '', waCell_(p.phone, renewMsg_(p))]); });
  head('Free trials ending in the next 2 days');
  cols(['Shop', 'Login', 'Trial ends', 'Hours left', '', 'Phone', 'Remind']);
  if (!d.trialsEnding.length) row(['None right now.']);
  d.trialsEnding.forEach(function (t) { row([t.shop, t.login, fmt_(t.end), t.hoursLeft, '', t.phone ? "'" + t.phone : '', waCell_(t.phone, trialMsg_(t))]); });
  head('Revenue this month by plan');
  cols(['Plan', 'Revenue (৳)']);
  var plans = Object.keys(d.byPlan); if (!plans.length) row(['No paid plans yet this month.']);
  plans.forEach(function (k) { row([k, d.byPlan[k]]); });
  head('All paying shops');
  cols(['Shop', 'Login', 'Plan', 'Ends', 'Days left', 'Code', 'Last login']);
  if (!d.paying.length) row(['No paying shops yet.']);
  d.paying.forEach(function (p) { row([p.shop, p.login, PLANS[p.tier] || p.tier, fmt_(p.exp), p.daysLeft, p.code, p.last ? fmt_(p.last) : '']); });
  sh.getRange(1, 1, out.length, 7).setValues(out);
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  sh.getRange(3, 1, 2, 7).setHorizontalAlignment('center');
  sh.getRange(4, 1, 1, 7).setFontSize(18).setFontWeight('bold').setFontColor('#1E9E63');
  fmtRows.forEach(function (f) {
    var r = sh.getRange(f[1], 1, 1, 7);
    if (f[0] === 'h') r.setFontWeight('bold').setFontSize(13).setBackground('#E8F7EF');
    else r.setFontWeight('bold').setFontColor('#555555');
  });
  sh.setColumnWidths(1, 7, 170); sh.setFrozenRows(0);
  PropertiesService.getScriptProperties().setProperty('DASH_AT', String(Date.now()));
  return d;
}
/** Emails shops whose plan ends in 3 days or 1 day (email logins), and sends the owner one summary. Each reminder goes once. */
function sendReminders_(d) {
  d = d || ownerData_();
  var pr = PropertiesService.getScriptProperties(), sent = 0, ownerLines = [];
  d.ending.forEach(function (p) {
    var stage = p.daysLeft <= 1 ? '1' : p.daysLeft <= 3 ? '3' : '';
    if (!stage) return;
    var key = 'REM|' + p.code + '|' + p.exp + '|' + stage;
    if (pr.getProperty(key)) return;
    if (p.login.indexOf('@') > 0) {
      try {
        MailApp.sendEmail(p.login, 'Your Store Flow plan ends ' + (stage === '1' ? 'tomorrow' : 'in 3 days'),
          renewMsg_(p) + '\n\nআপনার Store Flow প্ল্যান ' + fmt_(p.exp) + ' তারিখে শেষ হবে। চালু রাখতে বিকাশ, নগদ বা রকেটে পেমেন্ট করে অ্যাপের Plans পেজে Transaction ID দিন। ধন্যবাদ!');
        sent++;
      } catch (e) {}
    }
    ownerLines.push('• ' + p.shop + ' (' + (PLANS[p.tier] || p.tier) + ') ends ' + fmt_(p.exp) + ' · ' + (p.phone ? p.phone + ' · ' + waUrl_(p.phone, renewMsg_(p)) : p.login));
    pr.setProperty(key, String(Date.now()));
  });
  d.trialsEnding.forEach(function (t) {
    var key = 'TRI|' + t.login + '|' + t.end;
    if (t.hoursLeft > 24 || pr.getProperty(key)) return;
    ownerLines.push('• Trial: ' + (t.shop || t.login) + ' ends ' + fmt_(t.end) + ' · ' + (t.phone ? t.phone + ' · ' + waUrl_(t.phone, trialMsg_(t)) : t.login));
    pr.setProperty(key, String(Date.now()));
  });
  if (ownerLines.length) {
    try {
      var owner = Session.getEffectiveUser().getEmail();
      if (owner) MailApp.sendEmail(owner, 'Store Flow: ' + ownerLines.length + ' shop(s) to remind today',
        'These shops have a plan or free trial ending soon. Tap a WhatsApp link to send them a ready-made reminder.\n\n' + ownerLines.join('\n') + '\n\nOwner dashboard: ' + ss_().getUrl());
    } catch (e) {}
  }
  return { emailed: sent, listed: ownerLines.length };
}
/** Runs once a day, the first time anyone opens the app that day. No extra setup or permissions needed. */
function daily_() {
  try {
    var pr = PropertiesService.getScriptProperties(), today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
    if (pr.getProperty('DAILY_AT') === today) return;
    var lock = LockService.getScriptLock(); if (!lock.tryLock(100)) return;
    try {
      if (pr.getProperty('DAILY_AT') === today) return;
      pr.setProperty('DAILY_AT', today);
      sendReminders_(buildDashboard_());
    } finally { lock.releaseLock(); }
  } catch (e) {}
}
function menuDashboard() { buildDashboard_(); ss_().setActiveSheet(ss_().getSheetByName('Dashboard')); }
function menuReminders() {
  var r = sendReminders_(buildDashboard_());
  SpreadsheetApp.getUi().alert('Renewal reminders', r.listed ? ('Emailed ' + r.emailed + ' shop(s). ' + r.listed + ' reminder(s) are in your email and on the Dashboard tab with WhatsApp links.') : 'No plans or trials are ending soon. Nothing to send today.', SpreadsheetApp.getUi().ButtonSet.OK);
}
