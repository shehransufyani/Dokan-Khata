/**
 * Store Flow on Google Apps Script.
 * - Serves the app (Index.html) as a web app.
 * - Shop owners sign up with an email or phone number and a password ("Accounts" sheet).
 * - Keeps every shop's products, sales, income/expenses and settings in this Google Sheet.
 * - Checks access codes against the "Codes" sheet, so a code can be switched off or extended here.
 * - Saves plan payments (bKash, Nagad, Rocket) to the "Payments" sheet and emails the owner.
 * Bind this script to a Google Sheet (Extensions > Apps Script) and deploy it as a web app.
 */

var PLANS = { basic: 'Shuru (৳500)', standard: 'Bepsha (৳800)', pro: 'Pro (৳1200)' };
/** Payment options people can pay for. Pro Yearly gives a 12-month Pro code. */
var PAY_PLANS = { basic: PLANS.basic, standard: PLANS.standard, pro: PLANS.pro, yearly: 'Pro Yearly (৳10000)' };
function payTerms_(opt) { return opt === 'yearly' ? { tier: 'pro', months: 12, note: 'Yearly · ' } : { tier: opt, months: 1, note: '' }; }
var SHEETS = {
  codes:    ['Code', 'Business', 'Phone', 'Plan', 'Expires', 'Status', 'Created', 'Note', 'Account'],
  payments: ['Paid at', 'Package', 'Method', 'Transaction ID', 'Store name', 'Phone', 'Status', 'Access code', 'Account'],
  data:     ['Code', 'Store', 'ID', 'Updated', 'JSON'],
  accounts: ['ID', 'Login', 'Shop', 'Created', 'Trial started', 'Code', 'Last login', 'Hash', 'Salt', 'Reset']
};
var SHEET_NAMES = { codes: 'Codes', payments: 'Payments', data: 'Data', accounts: 'Accounts' };
var TRIAL_MS = 7 * 24 * 60 * 60 * 1000; // 7-day free trial
var STORES = ['products', 'sales', 'entries', 'orders', 'meta'];
var PRICES = { basic: 500, standard: 800, pro: 1200, yearly: 10000 };

/* ---------- web app ---------- */
function doGet(e) {
  if (e && e.parameter && e.parameter.feed) return catalogFeed_(e.parameter.feed);
  if (e && e.parameter && e.parameter.admin) return HtmlService.createHtmlOutputFromFile('Admin').setTitle('Store Flow Admin')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Store Flow')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setFaviconUrl('https://shehransufyani.github.io/StoreFlow/icons/icon-192.png')
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
  if (!PAY_PLANS[tier] || ['bkash', 'nagad', 'rocket'].indexOf(method) < 0) throw new Error('Choose a plan and a payment method');
  var terms = payTerms_(tier);
  var trx = String(p.trx || '').slice(0, 60), shop = String(p.shop || '').slice(0, 120), phone = String(p.phone || '').slice(0, 30);
  if (!trx || !shop || !phone) throw new Error('Transaction ID, shop name and phone are needed');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  var res;
  try {
    var ps = sheet_('payments'), last = ps.getLastRow();
    var used = last > 1 ? ps.getRange(2, 4, last - 1, 1).getValues() : [];
    for (var i = 0; i < used.length; i++) if (norm_(used[i][0]) === norm_(trx)) throw new Error('DUP_TRX');
    // A unique code for this payment, already tied to the paying account and plan. Send it once the money has arrived.
    res = addCode_(shop, phone, terms.tier, terms.months, terms.note + 'TrxID ' + trx + ' via ' + method + ' for ' + a.login, a.id);
    ps.appendRow([new Date(), PAY_PLANS[tier], method, "'" + trx, shop, "'" + phone, 'check money, then send code', res.code, a.login]);
  } finally { lock.releaseLock(); }
  try {
    var owner = Session.getEffectiveUser().getEmail();
    if (owner) MailApp.sendEmail(owner, 'Store Flow payment: ' + shop + ' (' + PAY_PLANS[tier] + ')',
      'New plan payment to check.\n\nPlan: ' + PAY_PLANS[tier] + (terms.months === 12 ? ' (code lasts 12 months)' : '') + '\nPaid with: ' + method + '\nTransaction ID: ' + trx +
      '\nShop: ' + shop + '\nPhone: ' + phone + '\nAccount: ' + a.login + '\n\nAccess code for this payment: ' + res.code + '\n\nCheck the money arrived in your ' + method + ' app, then send this code to ' + phone + '. If the money did not arrive, set the code to "off" in the Codes sheet.\n' + ss_().getUrl());
  } catch (e) {}
  return true;
}

/* ---------- owner menu in the Sheet ---------- */
function onOpen() {
  var m = SpreadsheetApp.getUi().createMenu('Store Flow')
    .addItem('Create access code', 'menuCreateCode')
    .addItem('Show code for selected payment', 'menuApprovePayment')
    .addItem('Create my master code', 'menuMasterCode')
    .addItem('Reset password of selected account', 'menuResetPassword')
    .addSeparator()
    .addItem('Update owner dashboard', 'menuDashboard')
    .addItem('Send renewal reminders now', 'menuReminders')
    .addSeparator()
    .addItem('Open admin panel', 'menuAdmin')
    .addItem('Show app link', 'menuAppLink')
    .addItem('Set app link', 'menuSetAppLink')
    .addItem('Set up sheets', 'setup');
  if (typeof onOpenFlowAI_ === 'function') m = onOpenFlowAI_(m.addSeparator());
  m.addToUi();
}
function ask_(ui, title, prompt) {
  var r = ui.prompt(title, prompt, ui.ButtonSet.OK_CANCEL);
  return r.getSelectedButton() === ui.Button.OK ? r.getResponseText().trim() : null;
}
function askPlan_(ui) {
  var t = ask_(ui, 'Plan', 'Type 1 for Shuru (৳500), 2 for Bepsha (৳800), 3 for Pro (৳1200) or 4 for Pro Yearly (৳10000, 12 months)');
  return t === null ? null : ({ '1': 'basic', '2': 'standard', '3': 'pro', '4': 'yearly' })[t] || null;
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
  if (tier === 'yearly') { showCode_(ui, addCode_(biz, phone, 'pro', 12, 'Yearly'), biz); return; }
  var months = parseInt(ask_(ui, 'Length', 'How many months? (1, 3, 6 or 12)'), 10) || 1;
  showCode_(ui, addCode_(biz, phone, tier, months, ''), biz);
}
function menuApprovePayment() {
  var ui = SpreadsheetApp.getUi(), sh = ss_().getActiveSheet();
  if (sh.getName() !== SHEET_NAMES.payments || sh.getActiveRange().getRow() < 2) { ui.alert('Open the Payments sheet and click a payment row first.'); return; }
  var r = sh.getActiveRange().getRow(), v = sh.getRange(r, 1, 1, 9).getValues()[0];
  if (v[7]) { sh.getRange(r, 7).setValue('code sent'); ui.alert('Code for ' + v[4], 'Send this code to ' + v[5] + ':\n\n' + v[7] + '\n\nThe row is now marked "code sent".', ui.ButtonSet.OK); return; }
  var opt = PAY_PLANS[String(v[1])] ? String(v[1]) : Object.keys(PAY_PLANS).filter(function (k) { return PAY_PLANS[k] === String(v[1]); })[0];
  if (!opt) { ui.alert('Unknown package in this row: ' + v[1]); return; }
  var terms = payTerms_(opt), months = terms.months;
  if (opt !== 'yearly') { months = parseInt(ask_(ui, 'Approve ' + v[4], 'Payment checked? How many months to give? (usually 1)'), 10); if (!months) return; }
  var res = addCode_(String(v[4]), String(v[5]), terms.tier, months, terms.note + 'TrxID ' + v[3] + ' via ' + v[2] + (v[8] ? ' for ' + v[8] : ''));
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
/** The Sheet turns phone logins like 01712345678 into the number 1712345678. Compare logins in one form. */
function loginKey_(v) { var s = String(v || '').trim(); return s.indexOf('@') >= 0 ? s.toLowerCase() : phone_(s); }
function findAccount_(by, value) {
  var sh = sheet_('accounts'), last = sh.getLastRow();
  if (last < 2) return null;
  var rows = sh.getRange(2, 1, last - 1, 10).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (by === 'id' ? String(r[0]) !== String(value) : loginKey_(r[1]) !== loginKey_(value)) continue;
    var ts = r[4] instanceof Date ? r[4].getTime() : (r[4] ? new Date(r[4]).getTime() : 0);
    return { row: i + 2, id: String(r[0]), login: loginKey_(r[1]), store: String(r[2] || ''), trialStart: ts || 0,
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
    sheet_('accounts').appendRow([id, login.indexOf('@') < 0 ? "'" + login : login, '', new Date(), '', '', new Date(), hash_(pw, salt), salt, '']);
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
    sheet_('accounts').getRange(a.row, 10).setValue(hash_(code, a.salt) + '|' + (Date.now() + 10 * 60000));
    MailApp.sendEmail(login, 'Store Flow password reset code', 'Your Store Flow reset code is ' + code + '\n\nIt works for 10 minutes. If you did not ask for this, ignore this email.');
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
function txt_(v) { v = String(v || ''); return /^\d+$/.test(v) ? "'" + v : v; }
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
    return { code: norm_(r[0]), biz: String(r[1] || ''), phone: phone_(r[2]), tier: String(r[3] || '').toLowerCase() === 'pro' && /yearly/i.test(String(r[7] || '')) ? 'yearly' : String(r[3] || '').toLowerCase(), exp: exp,
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
  return 'Assalamu alaikum! Your Store Flow ' + (PAY_PLANS[p.tier] || p.tier) + ' plan for ' + p.shop + ' ends on ' + fmt_(p.exp) +
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
  d.ending.forEach(function (p) { row([p.shop, txt_(p.login), PAY_PLANS[p.tier] || p.tier, fmt_(p.exp), p.daysLeft, p.phone ? "'" + p.phone : '', waCell_(p.phone, renewMsg_(p))]); });
  head('Free trials ending in the next 2 days');
  cols(['Shop', 'Login', 'Trial ends', 'Hours left', '', 'Phone', 'Remind']);
  if (!d.trialsEnding.length) row(['None right now.']);
  d.trialsEnding.forEach(function (t) { row([t.shop, txt_(t.login), fmt_(t.end), t.hoursLeft, '', t.phone ? "'" + t.phone : '', waCell_(t.phone, trialMsg_(t))]); });
  head('Revenue this month by plan');
  cols(['Plan', 'Revenue (৳)']);
  var plans = Object.keys(d.byPlan); if (!plans.length) row(['No paid plans yet this month.']);
  plans.forEach(function (k) { row([k, d.byPlan[k]]); });
  head('All paying shops');
  cols(['Shop', 'Login', 'Plan', 'Ends', 'Days left', 'Code', 'Last login']);
  if (!d.paying.length) row(['No paying shops yet.']);
  d.paying.forEach(function (p) { row([p.shop, txt_(p.login), PAY_PLANS[p.tier] || p.tier, fmt_(p.exp), p.daysLeft, p.code, p.last ? fmt_(p.last) : '']); });
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
    ownerLines.push('• ' + p.shop + ' (' + (PAY_PLANS[p.tier] || p.tier) + ') ends ' + fmt_(p.exp) + ' · ' + (p.phone ? p.phone + ' · ' + waUrl_(p.phone, renewMsg_(p)) : p.login));
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
/* ---------- daily profit summary: free email each morning about yesterday ---------- */
/** Emails every active shop that logs in with an email address a short summary of yesterday. Free (Apps Script email).
    Shops can switch it off in Setup. Phone-number shops see the same summary on their home screen instead. */
function sendDailySummaries_() {
  var accts = rows_('accounts', 10), byId = {};
  accts.forEach(function (r) {
    var login = loginKey_(r[1]); if (login.indexOf('@') < 0) return;
    var a = { id: String(r[0]), login: login, store: String(r[2] || ''), code: norm_(r[5]), trialStart: r[4] ? new Date(r[4]).getTime() : 0 };
    if (!trialOn_(a) && !usable_(a.code)) return;
    byId[a.id] = { a: a, sales: [], entries: [], products: [], meta: {} };
  });
  if (!Object.keys(byId).length) return 0;
  var start = new Date(); start.setHours(0, 0, 0, 0);
  var to = start.getTime(), from = to - 864e5;
  rows_('data', 5).forEach(function (r) {
    var b = byId[norm_(r[0])]; if (!b) return;
    var o; try { o = JSON.parse(r[4]); } catch (e) { return; }
    if (r[1] === 'sales' && o.at >= from && o.at < to) b.sales.push(o);
    else if (r[1] === 'sales' && o.method === 'due' && !o.paidAt) (b.dues = b.dues || []).push(o);
    else if (r[1] === 'entries' && o.at >= from && o.at < to) b.entries.push(o);
    else if (r[1] === 'products') b.products.push(o);
    else if (r[1] === 'meta') b.meta = o;
  });
  var sent = 0, day = fmt_(from);
  Object.keys(byId).forEach(function (id) {
    var b = byId[id], m = b.meta || {};
    if (m.dailyEmail === false || m.sample) return;
    if (!b.sales.length && !b.entries.length) return;
    var sum = function (arr, f) { return arr.reduce(function (s, x) { return s + (+f(x) || 0); }, 0); };
    var sales = sum(b.sales, function (s) { return s.total; });
    var profit = sales - sum(b.sales, function (s) { return sum(s.items || [], function (i) { return i.qty * (i.cost || 0); }); });
    var exp = sum(b.entries.filter(function (e) { return e.type === 'expense' && e.category !== 'purchase'; }), function (e) { return e.amount; });
    var inc = sum(b.entries.filter(function (e) { return e.type === 'income'; }), function (e) { return e.amount; });
    var dueNew = sum(b.sales.filter(function (s) { return s.method === 'due'; }), function (s) { return s.total; });
    var dueAll = dueNew + sum(b.dues || [], function (s) { return s.total; });
    var top = {}; b.sales.forEach(function (s) { (s.items || []).forEach(function (i) { top[i.name] = (top[i.name] || 0) + i.qty * i.price; }); });
    var best = Object.keys(top).sort(function (x, y) { return top[y] - top[x]; }).slice(0, 3);
    var low = b.products.filter(function (p) { return p.stock <= 0 || (p.lowAt && p.stock <= p.lowAt); }).map(function (p) { return p.name; }).slice(0, 8);
    var tk = function (n) { return '৳' + Math.round(n).toLocaleString('en-US'); }, shop = b.a.store || m.shopName || 'your shop';
    var en = [shop + ' — ' + day, '', 'Sales: ' + tk(sales) + ' from ' + b.sales.length + ' bill(s)', 'Profit on goods sold: ' + tk(profit),
      'Running costs: ' + tk(exp) + (inc ? ' · Other income: ' + tk(inc) : ''), 'Net for the day: ' + tk(profit + inc - exp),
      'New baki (due): ' + tk(dueNew) + ' · Total still owed to you: ' + tk(dueAll)];
    if (best.length) en.push('Best sellers: ' + best.join(', '));
    if (low.length) en.push('Low or out of stock: ' + low.join(', '));
    var bn = ['গতকালের হিসাব (' + day + ')', 'বিক্রি: ' + tk(sales) + ' (' + b.sales.length + 'টি বিল)', 'বিক্রিত মালের লাভ: ' + tk(profit), 'খরচ: ' + tk(exp),
      'দিনের নিট: ' + tk(profit + inc - exp), 'মোট বাকি পাওনা: ' + tk(dueAll)];
    try {
      MailApp.sendEmail(b.a.login, 'Store Flow: yesterday ' + tk(sales) + ' sales, ' + tk(profit + inc - exp) + ' net',
        en.join('\n') + '\n\n' + bn.join('\n') + '\n\nTo stop these emails, open Store Flow > Setup and switch off the daily summary.');
      sent++;
    } catch (e) { console.error('summary mail: ' + e); }
  });
  return sent;
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
      try { sendDailySummaries_(); } catch (e) { console.error('daily summaries: ' + e); }
    } finally { lock.releaseLock(); }
  } catch (e) {}
}
function menuDashboard() { buildDashboard_(); ss_().setActiveSheet(ss_().getSheetByName('Dashboard')); }
function menuReminders() {
  var r = sendReminders_(buildDashboard_());
  SpreadsheetApp.getUi().alert('Renewal reminders', r.listed ? ('Emailed ' + r.emailed + ' shop(s). ' + r.listed + ' reminder(s) are in your email and on the Dashboard tab with WhatsApp links.') : 'No plans or trials are ending soon. Nothing to send today.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ---------- Facebook and Instagram catalog feed (free) ---------- */
/* Meta Commerce Manager can fetch this CSV link every day, so prices and stock there stay in step with the app. */
function feedSig_(id) { return b64_(Utilities.computeHmacSha256Signature(id + '|feed', secret_())).slice(0, 16); }
/** Called by the app (Pro). Returns this shop's catalog feed link. */
function catalogLink(token) {
  var a = auth_(token), c = usable_(a.code);
  if (!(c && c.tier === 'pro') && !trialOn_(a)) throw new Error('LOCKED');
  var url = PropertiesService.getScriptProperties().getProperty('APP_URL') || '';
  if (!url) { try { url = ScriptApp.getService().getUrl(); } catch (e) {} }
  return url + '?feed=' + encodeURIComponent(a.id + '.' + feedSig_(a.id));
}
function catalogFeed_(raw) {
  var parts = String(raw).split('.'), id = parts[0], out = ContentService.createTextOutput().setMimeType(ContentService.MimeType.CSV);
  if (!id || parts[1] !== feedSig_(id)) return out.setContent('error,invalid feed link');
  var a = findAccount_('id', id);
  if (!a || (!usable_(a.code) && !trialOn_(a))) return out.setContent('error,plan not active');
  var products = [], meta = {};
  rows_('data', 5).forEach(function (r) {
    if (norm_(r[0]) !== norm_(id)) return;
    try { if (r[1] === 'products') products.push(JSON.parse(r[4])); else if (r[1] === 'meta') meta = JSON.parse(r[4]); } catch (e) {}
  });
  var shop = a.store || meta.shopName || 'Shop', phone = phone_(meta.phone || '');
  var img = 'https://shehransufyani.github.io/StoreFlow/icons/icon-512.png';
  var q = function (v) { return '"' + String(v == null ? '' : v).replace(/"/g, '""').replace(/[\r\n]+/g, ' ') + '"'; };
  var lines = [['id', 'title', 'description', 'availability', 'condition', 'price', 'link', 'image_link', 'brand', 'product_type', 'quantity_to_sell_on_facebook'].join(',')];
  products.filter(function (p) { return p && p.name && +p.price > 0; }).forEach(function (p) {
    var title = p.name + (p.nameBn ? ' (' + p.nameBn + ')' : '');
    var link = (phone && waUrl_(phone, 'Hi ' + shop + ', I want to order: ' + p.name)) || 'https://shehransufyani.github.io/StoreFlow/';
    var photo = /^https?:\/\//.test(String(p.photo || '')) ? p.photo : img;
    lines.push([q(p.id), q(title.slice(0, 150)), q(title + ' from ' + shop + (p.unit ? ', per ' + p.unit : '') + '. Order on WhatsApp.'),
      q(p.stock > 0 ? 'in stock' : 'out of stock'), q('new'), q((+p.price).toFixed(2) + ' BDT'), q(link), q(photo), q(shop), q(p.category || ''),
      q(Math.max(0, Math.floor(+p.stock || 0)))].join(','));
  });
  return out.setContent(lines.join('\n'));
}

/* ---------- admin panel (web page at the app link + ?admin=1) ---------- */
/* Only someone with the admin key can read or change anything. The key lives in Script Properties (ADMIN_KEY);
   the owner sees it from the Sheet menu: Store Flow > Open admin panel. */
function adminKey_() {
  var pr = PropertiesService.getScriptProperties(), k = pr.getProperty('ADMIN_KEY');
  if (!k) { k = Utilities.getUuid().replace(/-/g, '').slice(0, 20); pr.setProperty('ADMIN_KEY', k); }
  return k;
}
function adminAuth_(key) {
  var c = CacheService.getScriptCache(), n = +(c.get('adminfail') || 0);
  if (n >= 10) throw new Error('TOO_MANY');
  if (!key || String(key) !== adminKey_()) { c.put('adminfail', String(n + 1), 900); throw new Error('ADMIN_KEY'); }
}
function adminUrl_() {
  var url = PropertiesService.getScriptProperties().getProperty('APP_URL') || '';
  if (!url) { try { url = ScriptApp.getService().getUrl(); } catch (e) {} }
  return url + '?admin=1';
}
function menuAdmin() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('Admin panel', 'Open this link (bookmark it):\n' + adminUrl_() + '\n\nYour admin key (keep it private, do not share it):\n' + adminKey_() +
    '\n\nThe panel asks for the key once on each device.', ui.ButtonSet.OK);
}
function ms_(v) { return v instanceof Date ? v.getTime() : (v ? new Date(v).getTime() : 0) || 0; }
/** Called by the admin page. Everything the panel shows. */
function adminData(key) {
  adminAuth_(key);
  var now = Date.now(), month = new Date(); month.setDate(1); month.setHours(0, 0, 0, 0);
  var codes = rows_('codes', 9).map(function (r, i) {
    return { row: i + 2, code: norm_(r[0]), biz: String(r[1] || ''), phone: phone_(r[2]), tier: String(r[3] || '').toLowerCase(), exp: ms_(r[4]),
             active: String(r[5] || 'active').toLowerCase() === 'active', created: ms_(r[6]), note: String(r[7] || ''), master: /^master/i.test(String(r[7] || '')), account: String(r[8] || '') };
  });
  var byCode = {}; codes.forEach(function (c) { byCode[c.code] = c; });
  var pays = rows_('payments', 9).map(function (r, i) {
    return { row: i + 2, at: ms_(r[0]), plan: String(r[1] || ''), amount: price_(r[1]), method: String(r[2] || ''), trx: String(r[3] || ''), shop: String(r[4] || ''),
             phone: phone_(r[5]), status: String(r[6] || ''), code: norm_(r[7]), login: loginKey_(r[8]) };
  });
  var shops = rows_('accounts', 10).map(function (r) {
    var login = loginKey_(r[1]), email = login.indexOf('@') > 0 ? login : '';
    var a = { id: String(r[0]), login: login, email: email, shop: String(r[2] || ''), created: ms_(r[3]), trialStart: ms_(r[4]), code: norm_(r[5]), last: ms_(r[6]) };
    var c = byCode[a.code] || null;
    var myPays = pays.filter(function (p) { return (p.login && p.login === login) || (p.code && codes.some(function (x) { return x.code === p.code && x.account === a.id; })); });
    var phone = (c && c.phone) || (myPays[0] && myPays[0].phone) || (email ? '' : login);
    var trialEnd = a.trialStart ? a.trialStart + TRIAL_MS : 0;
    var status = c && c.master ? 'owner' : c && c.active && c.exp > now ? 'paid' : c && !c.active ? 'off' : trialEnd > now && !c ? 'trial' : c ? 'ended' : a.trialStart ? 'trial-ended' : 'new';
    var paid = myPays.filter(function (p) { return !/^(off|refunded|rejected)/i.test(p.status) && !/^check/i.test(p.status); });
    return { id: a.id, shop: a.shop || (c && c.biz) || '', login: login, email: email, phone: phone, created: a.created, last: a.last, trialEnd: trialEnd,
             code: a.code, tier: c ? (c.tier === 'pro' && /yearly/i.test(c.note) ? 'yearly' : c.tier) : '', exp: c ? c.exp : 0, codeActive: c ? c.active : false,
             status: status, payments: myPays, paidTotal: paid.reduce(function (s, p) { return s + p.amount; }, 0), lastPaid: myPays.length ? Math.max.apply(null, myPays.map(function (p) { return p.at; })) : 0 };
  });
  var used = {}; shops.forEach(function (x) { if (x.code) used[x.code] = 1; });
  var spare = codes.filter(function (c) { return !used[c.code] && !c.master; }).map(function (c) {
    return { code: c.code, biz: c.biz, phone: c.phone, tier: c.tier === 'pro' && /yearly/i.test(c.note) ? 'yearly' : c.tier, exp: c.exp, active: c.active, created: c.created, note: c.note, account: c.account };
  });
  var revMonth = 0, revAll = 0;
  pays.forEach(function (p) { if (/^(off|refunded|rejected|check)/i.test(p.status)) return; revAll += p.amount; if (p.at >= month.getTime()) revMonth += p.amount; });
  var waiting = pays.filter(function (p) { return /^check/i.test(p.status); }).sort(function (a, b) { return b.at - a.at; });
  return { now: now, appUrl: adminUrl_().replace('?admin=1', ''), shops: shops, spare: spare, waiting: waiting, revMonth: revMonth, revAll: revAll, prices: PRICES, plans: PAY_PLANS };
}
/** Make a new code. If an account is given, the shop gets it straight away. */
function adminCreateCode(key, o) {
  adminAuth_(key); o = o || {};
  var opt = String(o.plan || ''); if (!PAY_PLANS[opt]) throw new Error('PLAN');
  var terms = payTerms_(opt), months = Math.max(1, Math.min(36, parseInt(o.months, 10) || terms.months));
  var a = o.account ? findAccount_('id', o.account) : null;
  var biz = String(o.biz || (a && a.store) || '').slice(0, 120);
  var res = addCode_(biz, String(o.phone || '').slice(0, 30), terms.tier, months, terms.note + 'Made in admin panel', a ? a.id : '');
  if (a) sheet_('accounts').getRange(a.row, 6).setValue(res.code);
  return { code: res.code, exp: res.exp.getTime(), link: res.link };
}
/** Change a code: plan, end date, or switch it on/off. */
function adminUpdateCode(key, code, o) {
  adminAuth_(key); o = o || {};
  var c = findCode_(code); if (!c) throw new Error('NO_CODE');
  var sh = sheet_('codes');
  if (o.plan) {
    if (!PAY_PLANS[o.plan]) throw new Error('PLAN');
    var tier = o.plan === 'yearly' ? 'pro' : o.plan, note = String(sh.getRange(c.row, 8).getValue() || '');
    note = o.plan === 'yearly' ? (/yearly/i.test(note) ? note : 'Yearly · ' + note) : note.replace(/Yearly · /i, '');
    sh.getRange(c.row, 4).setValue(tier); sh.getRange(c.row, 8).setValue(note);
  }
  if (o.exp) { var d = new Date(+o.exp); d.setHours(23, 59, 59, 0); sh.getRange(c.row, 5).setValue(d); }
  if (o.active === true || o.active === false) sh.getRange(c.row, 6).setValue(o.active ? 'active' : 'off');
  return true;
}
/** Approve or reject a payment waiting to be checked. Approving gives the shop its code straight away. */
function adminPayment(key, row, action) {
  adminAuth_(key);
  var ps = sheet_('payments'); row = parseInt(row, 10);
  if (!(row >= 2) || row > ps.getLastRow()) throw new Error('ROW');
  var v = ps.getRange(row, 1, 1, 9).getValues()[0], c = findCode_(v[7]);
  if (action === 'approve') {
    ps.getRange(row, 7).setValue('approved');
    if (c) {
      sheet_('codes').getRange(c.row, 6).setValue('active');
      var a = c.account ? findAccount_('id', c.account) : (v[8] ? findAccount_('login', v[8]) : null);
      if (a) sheet_('accounts').getRange(a.row, 6).setValue(c.code);
    }
  } else if (action === 'reject') {
    ps.getRange(row, 7).setValue('rejected');
    if (c) sheet_('codes').getRange(c.row, 6).setValue('off');
  } else throw new Error('ACTION');
  return true;
}
/** Email a shop its code. */
function adminEmailCode(key, code, to) {
  adminAuth_(key);
  var c = findCode_(code); if (!c) throw new Error('NO_CODE');
  to = String(to || '').trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) throw new Error('NO_EMAIL');
  var plan = PAY_PLANS[c.tier === 'pro' && /yearly/i.test(String(sheet_('codes').getRange(c.row, 8).getValue())) ? 'yearly' : c.tier] || c.tier;
  MailApp.sendEmail(to, 'Your Store Flow access code: ' + c.code,
    'Assalamu alaikum!\n\nYour Store Flow access code' + (c.biz ? ' for ' + c.biz : '') + ' is: ' + c.code + '\nPlan: ' + plan + '\nValid until: ' + fmt_(c.exp) +
    '\n\nOpen this link to start: ' + appLink_(c.code) + '\nOr open the app, go to Plans and enter the code.\n\n' +
    'আপনার স্টোর ফ্লো অ্যাক্সেস কোড: ' + c.code + '। মেয়াদ ' + fmt_(c.exp) + ' পর্যন্ত। লিংকে চাপ দিন অথবা অ্যাপের Plans পেজে কোডটি দিন।\n\nThank you,\nStore Flow');
  return true;
}

