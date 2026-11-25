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
  SpreadsheetApp.getUi().createMenu('Store Flow')
    .addItem('Create access code', 'menuCreateCode')
    .addItem('Show code for selected payment', 'menuApprovePayment')
    .addItem('Create my master code', 'menuMasterCode')
    .addItem('Reset password of selected account', 'menuResetPassword')
    .addSeparator()
    .addItem('Update owner dashboard', 'menuDashboard')
    .addItem('Send renewal reminders now', 'menuReminders')
    .addSeparator()
    .addItem('Set Flow AI keys', 'menuSetAiKey')
    .addItem('Flow AI usage and cost', 'menuAiUsage')
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

/* ---------- Flow AI: an assistant that reads the shop's numbers ---------- */
/* Providers: Groq, Mistral and OpenRouter. Their keys live in Script Properties (GROQ_API_KEY, MISTRAL_API_KEY,
   OPENROUTER_API_KEY) and never reach the page. The owner sets them from the sheet menu: Store Flow > Set Flow AI keys.
   Flow AI tries the providers in AI_ORDER (default groq, mistral, openrouter) and moves on when one fails.
   GROQ_MODEL, MISTRAL_MODEL and OPENROUTER_MODEL can override the models below. */
var AI_SHEET = 'Flow AI';
var AI_COLS = ['Time', 'Account', 'Shop', 'Question', 'Answer', 'Memory', 'Lang', 'Tokens in', 'Tokens out', 'Provider'];
var AI_LIMITS = { pro: 50, trial: 10 };       // questions per shop per day
function aiSheet_() {
  var ss = ss_(), sh = ss.getSheetByName(AI_SHEET);
  if (!sh) { sh = ss.insertSheet(AI_SHEET); sh.getRange(1, 1, 1, AI_COLS.length).setValues([AI_COLS]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
/** Who may use Flow AI: Pro and Pro Yearly codes, and shops in their free trial (with a smaller daily limit). */
function aiAccess_(a) {
  var c = a.code ? usable_(a.code) : null;
  if (c) return c.tier === 'pro' ? 'pro' : null;
  return trialOn_(a) ? 'trial' : null;
}
function aiUsed_(a, add) {
  var c = CacheService.getScriptCache(), k = 'ai:' + a.id + ':' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd');
  var n = +(c.get(k) || 0);
  if (add) { n++; c.put(k, String(n), 6 * 3600); }
  return n;
}
/** Past questions for this shop, oldest first, plus the latest memory notes. */
function aiPast_(a, max) {
  var sh = aiSheet_(), last = sh.getLastRow(), out = [], memory = '';
  if (last < 2) return { turns: out, memory: memory };
  var from = Math.max(2, last - 1500), rows = sh.getRange(from, 1, last - from + 1, 7).getValues();
  for (var i = rows.length - 1; i >= 0 && out.length < max; i--) {
    if (String(rows[i][1]) !== a.id) continue;
    if (!memory && rows[i][5]) memory = String(rows[i][5]);
    out.unshift({ at: new Date(rows[i][0]).getTime(), q: String(rows[i][3]), a: String(rows[i][4]) });
  }
  return { turns: out, memory: memory };
}
/** Called by the app when Flow AI opens. */
function aiHistory(token) {
  var a = auth_(token), acc = aiAccess_(a);
  var p = acc ? aiPast_(a, 20) : { turns: [] };
  var key = aiProviders_().length > 0;
  return { access: acc, on: key, turns: p.turns, used: acc ? aiUsed_(a) : 0, limit: acc ? AI_LIMITS[acc] : 0 };
}
function aiNum_(n) { return Math.round(n || 0).toLocaleString('en-US'); }
/** Turns the shop's records into a short fact sheet for the model (numbers are worked out here, not by the model). */
function aiFacts_(d, shop) {
  var now = Date.now(), day = 864e5, start = new Date(); start.setHours(0, 0, 0, 0);
  var t0 = start.getTime(), monthStart = new Date(start.getFullYear(), start.getMonth(), 1).getTime();
  var meta = d.meta.filter(function (m) { return m && m.id === 'meta'; })[0] || d.meta[0] || {};
  var products = d.products || [], sales = d.sales || [], entries = d.entries || [], orders = d.orders || [];
  var sum = function (arr, f) { return arr.reduce(function (s, x) { return s + (+f(x) || 0); }, 0); };
  var pname = function (p) { return (p.name || '?') + (p.nameBn ? ' / ' + p.nameBn : ''); };
  function period(from, to) {
    var ss = sales.filter(function (s) { return s.at >= from && s.at < to; });
    var es = entries.filter(function (e) { return e.at >= from && e.at < to; });
    var total = sum(ss, function (s) { return s.total; });
    var cogs = sum(ss, function (s) { return sum(s.items || [], function (i) { return i.qty * (i.cost || 0); }); });
    var exp = sum(es.filter(function (e) { return e.type === 'expense' && e.category !== 'purchase'; }), function (e) { return e.amount; });
    var buy = sum(es.filter(function (e) { return e.type === 'expense' && e.category === 'purchase'; }), function (e) { return e.amount; });
    var inc = sum(es.filter(function (e) { return e.type === 'income'; }), function (e) { return e.amount; });
    return 'sales ৳' + aiNum_(total) + ' from ' + ss.length + ' sales, gross profit ৳' + aiNum_(total - cogs) +
      ', running costs ৳' + aiNum_(exp) + ', other income ৳' + aiNum_(inc) + ', stock bought ৳' + aiNum_(buy) +
      ', net earnings ৳' + aiNum_(total - cogs + inc - exp);
  }
  var L = [];
  L.push('Shop: ' + (shop || meta.shopName || 'unnamed') + '. Currency: Bangladeshi taka (৳). Today: ' + Utilities.formatDate(new Date(), tz_(), 'EEEE d MMM yyyy, HH:mm') + '.');
  if (meta.sample) L.push('NOTE: the shop is still showing the sample data that comes with the app, not real records.');
  L.push('Products: ' + products.length + '. Sales recorded: ' + sales.length + (sales.length ? ' (first on ' + fmt_(Math.min.apply(null, sales.map(function (s) { return s.at; }))) + ')' : '') + '.');
  L.push('', 'TOTALS');
  L.push('- Today: ' + period(t0, now + 1));
  L.push('- Last 7 days: ' + period(t0 - 6 * day, now + 1));
  L.push('- Previous 7 days: ' + period(t0 - 13 * day, t0 - 6 * day));
  L.push('- Last 30 days: ' + period(t0 - 29 * day, now + 1));
  L.push('- The 30 days before that: ' + period(t0 - 59 * day, t0 - 29 * day));
  L.push('- This month so far: ' + period(monthStart, now + 1));

  // per product, last 30 days
  var from30 = t0 - 29 * day, per = {};
  sales.filter(function (s) { return s.at >= from30; }).forEach(function (s) {
    (s.items || []).forEach(function (i) {
      var r = per[i.pid] = per[i.pid] || { qty: 0, amt: 0, profit: 0 };
      r.qty += +i.qty || 0; r.amt += (+i.qty || 0) * (+i.price || 0); r.profit += (+i.qty || 0) * ((+i.price || 0) - (+i.cost || 0));
    });
  });
  var rowsP = products.map(function (p) {
    var r = per[p.id] || { qty: 0, amt: 0, profit: 0 }, perDay = r.qty / 30;
    return { p: p, r: r, daysLeft: perDay > 0 ? Math.floor(Math.max(0, p.stock) / perDay) : null };
  });
  var best = rowsP.filter(function (x) { return x.r.qty > 0; }).sort(function (a, b) { return b.r.amt - a.r.amt; });
  L.push('', 'BEST SELLERS, last 30 days (by money taken)');
  if (!best.length) L.push('- none yet');
  best.slice(0, 15).forEach(function (x) {
    var p = x.p, m = p.price > 0 ? Math.round((p.price - (p.cost || 0)) / p.price * 100) : 0;
    L.push('- ' + pname(p) + ': sold ' + aiNum_(x.r.qty) + ' ' + (p.unit || '') + ' for ৳' + aiNum_(x.r.amt) + ', profit ৳' + aiNum_(x.r.profit) +
      ', price ৳' + p.price + ', cost ৳' + (p.cost || 0) + ' (margin ' + m + '%), stock left ' + aiNum_(p.stock) + (x.daysLeft !== null ? ' (about ' + x.daysLeft + ' days at this pace)' : ''));
  });
  var low = products.filter(function (p) { return p.stock <= 0 || (p.lowAt && p.stock <= p.lowAt); });
  L.push('', 'LOW OR OUT OF STOCK');
  if (!low.length) L.push('- none');
  low.slice(0, 20).forEach(function (p) { var r = per[p.id]; L.push('- ' + pname(p) + ': stock ' + aiNum_(p.stock) + (p.lowAt ? ' (alert at ' + p.lowAt + ')' : '') + ', sold ' + aiNum_(r ? r.qty : 0) + ' in 30 days, cost ৳' + (p.cost || 0)); });
  var slow = rowsP.filter(function (x) { return x.r.qty === 0 && x.p.stock > 0; }).sort(function (a, b) { return b.p.stock * (b.p.cost || 0) - a.p.stock * (a.p.cost || 0); });
  L.push('', 'NOT SOLD IN 30 DAYS but still in stock (money stuck in stock)');
  if (!slow.length) L.push('- none');
  slow.slice(0, 15).forEach(function (x) { L.push('- ' + pname(x.p) + ': ' + aiNum_(x.p.stock) + ' in stock, worth ৳' + aiNum_(x.p.stock * (x.p.cost || 0)) + ' at cost, price ৳' + x.p.price); });
  var thin = products.filter(function (p) { return p.price > 0 && p.cost > 0 && (p.price - p.cost) / p.price < 0.08; });
  if (thin.length) { L.push('', 'VERY LOW OR NEGATIVE MARGIN (price vs cost)'); thin.slice(0, 15).forEach(function (p) { L.push('- ' + pname(p) + ': price ৳' + p.price + ', cost ৳' + p.cost); }); }
  L.push('', 'Stock value at cost: ৳' + aiNum_(sum(products, function (p) { return Math.max(0, p.stock) * (p.cost || 0); })));

  // dues
  var dues = {}, dueOld = {};
  sales.filter(function (s) { return s.method === 'due' && !s.paidAt; }).forEach(function (s) {
    var k = s.customer || 'unknown'; dues[k] = (dues[k] || 0) + (+s.total || 0); dueOld[k] = Math.min(dueOld[k] || s.at, s.at);
  });
  var dk = Object.keys(dues).sort(function (a, b) { return dues[b] - dues[a]; });
  L.push('', 'CUSTOMERS WHO OWE MONEY (due sales not yet paid): total ৳' + aiNum_(sum(dk, function (k) { return dues[k]; })));
  if (!dk.length) L.push('- none');
  dk.slice(0, 20).forEach(function (k) { L.push('- ' + k + ': ৳' + aiNum_(dues[k]) + ', oldest from ' + fmt_(dueOld[k]) + ' (' + Math.floor((now - dueOld[k]) / day) + ' days ago)'); });

  // money in and out
  var byCat = {};
  entries.filter(function (e) { return e.at >= from30 && e.type === 'expense'; }).forEach(function (e) { byCat[e.category || 'other'] = (byCat[e.category || 'other'] || 0) + (+e.amount || 0); });
  L.push('', 'EXPENSES BY TYPE, last 30 days: ' + (Object.keys(byCat).map(function (k) { return k + ' ৳' + aiNum_(byCat[k]); }).join(', ') || 'none'));
  var methods = {};
  sales.filter(function (s) { return s.at >= from30; }).forEach(function (s) { methods[s.method || 'cash'] = (methods[s.method || 'cash'] || 0) + (+s.total || 0); });
  L.push('How customers paid, last 30 days: ' + (Object.keys(methods).map(function (k) { return k + ' ৳' + aiNum_(methods[k]); }).join(', ') || 'none'));

  // when the shop sells
  var wd = [0, 0, 0, 0, 0, 0, 0], hr = {}, names = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  sales.filter(function (s) { return s.at >= t0 - 59 * day; }).forEach(function (s) {
    var dt = new Date(s.at); wd[dt.getDay()] += +s.total || 0;
    var h = +Utilities.formatDate(dt, tz_(), 'H'); hr[h] = (hr[h] || 0) + (+s.total || 0);
  });
  L.push('Sales by weekday, last 60 days: ' + names.map(function (n, i) { return n + ' ৳' + aiNum_(wd[i]); }).join(', '));
  var hk = Object.keys(hr).sort(function (a, b) { return hr[b] - hr[a]; }).slice(0, 4);
  if (hk.length) L.push('Busiest hours, last 60 days: ' + hk.map(function (h) { return h + ':00 (৳' + aiNum_(hr[h]) + ')'; }).join(', '));
  var avg = sales.filter(function (s) { return s.at >= from30; });
  if (avg.length) L.push('Average sale, last 30 days: ৳' + aiNum_(sum(avg, function (s) { return s.total; }) / avg.length) + '. Discounts given: ৳' + aiNum_(sum(avg, function (s) { return s.discount; })) + '.');

  // orders
  var open = orders.filter(function (o) { return o.status === 'pending' || o.status === 'ready' || o.status === 'sent'; });
  var pre = open.filter(function (o) { return o.type === 'pre'; });
  L.push('', 'ORDERS: ' + open.length + ' open (' + pre.length + ' pre-orders), ' + orders.filter(function (o) { return o.status === 'paid' || o.status === 'done' || o.status === 'delivered'; }).length + ' completed, ' + orders.filter(function (o) { return o.status === 'cancelled' || o.status === 'returned'; }).length + ' cancelled or returned.');
  pre.filter(function (o) { return o.pickupAt; }).sort(function (a, b) { return a.pickupAt - b.pickupAt; }).slice(0, 8).forEach(function (o) { L.push('- pre-order for ' + (o.customer || '?') + ' on ' + fmt_(o.pickupAt)); });
  return L.join('\n');
}
var AI_SYSTEM = [
  'You are Flow AI, the business assistant inside Store Flow, an app that small shops in Bangladesh use to track stock, sales, money and dues.',
  'You talk to the shop owner. Your job is to help them understand their shop and make more money: what to restock, what sells best, what is not selling, pricing, costs, who owes money and how to collect it politely, and practical ideas to grow sales.',
  'Use the shop facts given to you. They are worked out from the shop\'s own records and are correct; quote their numbers rather than doing new arithmetic where you can. If the facts do not cover something, say so plainly instead of guessing, and suggest what to record in the app so you can answer next time.',
  'Write for a busy shopkeeper on a phone: short, warm and practical. Lead with the answer. Use at most 5 short bullet points or 2 short paragraphs, then one clear next step. Use ৳ for money. No tables, no headings, no long introductions.',
  'Simple formatting only: **bold** for key numbers or product names, lines starting with "- " for bullets.',
  'Advice must suit a small Bangladeshi shop (bKash/Nagad, WhatsApp customers, local suppliers, festivals such as Eid and Pohela Boishakh). Never suggest anything illegal or dishonest, never shame customers who owe money.',
  'If a question has nothing to do with the shop or business, answer briefly and kindly, then steer back to how you can help the shop.',
  'You also keep short memory notes about this shop between conversations: lasting facts the owner told you (goals, suppliers, plans, preferences) and advice you gave that you should follow up on. Update them each time: keep what is still useful, add what is new, drop what is stale, at most 8 short lines. Do not copy numbers that are already in the shop facts.',
  'Always reply with only a JSON object with two string fields: "answer" (the reply shown to the owner) and "memory" (the updated memory notes, one note per line).'
].join('\n');
var AI_PROVIDERS = {
  groq: { name: 'Groq', key: 'GROQ_API_KEY', url: 'https://api.groq.com/openai/v1/chat/completions', model: 'openai/gpt-oss-120b', price: [0.15, 0.60] },
  mistral: { name: 'Mistral', key: 'MISTRAL_API_KEY', url: 'https://api.mistral.ai/v1/chat/completions', model: 'mistral-large-latest', price: [0.70, 2.10] },
  openrouter: { name: 'OpenRouter', key: 'OPENROUTER_API_KEY', url: 'https://openrouter.ai/api/v1/chat/completions', model: 'deepseek/deepseek-v4.1-flash', price: [0.14, 0.42] }
};
/** The providers that have a key, in the order to try them. */
function aiProviders_() {
  var pr = PropertiesService.getScriptProperties();
  var order = String(pr.getProperty('AI_ORDER') || 'groq,mistral,openrouter').toLowerCase().split(/[\s,]+/);
  return order.filter(function (id) { return AI_PROVIDERS[id] && pr.getProperty(AI_PROVIDERS[id].key); }).map(function (id) {
    var p = AI_PROVIDERS[id];
    return { id: id, name: p.name, url: p.url, key: pr.getProperty(p.key), model: pr.getProperty(id.toUpperCase() + '_MODEL') || p.model };
  });
}
/** One chat request to one provider (all three speak the same OpenAI-style chat format). */
function aiCallOne_(p, messages, maxTokens, json) {
  var body = { model: p.model, messages: messages, max_tokens: maxTokens, temperature: 0.4 };
  if (json) body.response_format = { type: 'json_object' };
  var headers = { Authorization: 'Bearer ' + p.key };
  if (p.id === 'openrouter') { headers['X-Title'] = 'Store Flow'; headers['HTTP-Referer'] = 'https://shehransufyani.github.io/Dokan-Khata/'; }
  var res = UrlFetchApp.fetch(p.url, { method: 'post', contentType: 'application/json', headers: headers, payload: JSON.stringify(body), muteHttpExceptions: true });
  var code = res.getResponseCode(), out = {};
  try { out = JSON.parse(res.getContentText()); } catch (e) {}
  if (code !== 200 || !out.choices || !out.choices[0]) {
    var msg = (out.error && (out.error.message || out.error)) || res.getContentText().slice(0, 300);
    var err = new Error(code === 401 || code === 403 ? 'AI_KEY' : code === 402 || /credit|billing|balance|quota|insufficient/i.test(String(msg)) ? 'AI_BILLING' : code === 429 || code >= 500 ? 'AI_BUSY' : 'AI_FAIL');
    err.detail = p.name + ' ' + code + ': ' + String(msg).slice(0, 300);
    throw err;
  }
  var u = out.usage || {};
  return { text: String(out.choices[0].message && out.choices[0].message.content || ''), provider: p.name, tin: u.prompt_tokens || 0, tout: u.completion_tokens || 0 };
}
/** Tries each provider with a key until one answers. */
function aiCall_(messages, maxTokens, json) {
  var list = aiProviders_(), last = null;
  if (!list.length) throw new Error('AI_OFF');
  for (var i = 0; i < list.length; i++) {
    try { return aiCallOne_(list[i], messages, maxTokens, json); }
    catch (e) { console.error('Flow AI: ' + (e.detail || e.message)); last = e; }
  }
  throw last;
}
function aiParse_(text) {
  var t = String(text || '').trim().replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
  try { return JSON.parse(t); } catch (e) {}
  var m = t.match(/\{[\s\S]*\}/);
  if (m) { try { return JSON.parse(m[0]); } catch (e) {} }
  return { answer: t, memory: '' };
}
/** Called by the app. Answers one question about the shop. */
function askFlow(token, question, lang) {
  var a = auth_(token), acc = aiAccess_(a);
  if (!acc) throw new Error('AI_PLAN');
  if (!aiProviders_().length) throw new Error('AI_OFF');
  question = String(question || '').trim().slice(0, 1200);
  if (!question) throw new Error('AI_EMPTY');
  if (aiUsed_(a) >= AI_LIMITS[acc]) throw new Error('AI_LIMIT');
  var past = aiPast_(a, 6), facts = aiFacts_(loadAll(token), a.store), bn = lang === 'bn';
  var messages = [{ role: 'system', content: AI_SYSTEM }];
  past.turns.forEach(function (t) {
    messages.push({ role: 'user', content: t.q });
    messages.push({ role: 'assistant', content: JSON.stringify({ answer: t.a, memory: '' }) });
  });
  messages.push({ role: 'user', content: [
    '<shop_facts>\n' + facts + '\n</shop_facts>',
    '<memory_notes>\n' + (past.memory || '(none yet)') + '\n</memory_notes>',
    'Reply in ' + (bn ? 'Bangla (বাংলা), using simple everyday words; product names can stay as the owner wrote them.' : 'simple English.') + ' Reply with only the JSON object.',
    '<question>\n' + question + '\n</question>'
  ].join('\n\n') });
  var r = aiCall_(messages, 2500, true), out = aiParse_(r.text);
  var answer = String(out.answer || '').trim(), memory = String(out.memory || past.memory || '').trim().slice(0, 1500);
  if (!answer) throw new Error('AI_FAIL');
  aiUsed_(a, true);
  aiSheet_().appendRow([new Date(), a.id, a.store || '', question, answer, memory, bn ? 'bn' : 'en', r.tin, r.tout, r.provider]);
  return { answer: answer, at: Date.now(), used: aiUsed_(a), limit: AI_LIMITS[acc] };
}
/** Sheet menu: paste a key for Groq, Mistral or OpenRouter. Keys are stored in Script Properties and each is tested. */
function menuSetAiKey() {
  var ui = SpreadsheetApp.getUi(), pr = PropertiesService.getScriptProperties();
  var pick = ask_(ui, 'Set Flow AI keys', 'Which key are you adding? Type 1 for Groq, 2 for Mistral, 3 for OpenRouter.');
  var id = { '1': 'groq', '2': 'mistral', '3': 'openrouter' }[pick];
  if (!id) return;
  var P = AI_PROVIDERS[id];
  var k = ask_(ui, 'Set ' + P.name + ' key', 'Paste your ' + P.name + ' API key. It is kept in this script\'s settings and is never sent to the app pages. Leave empty and press OK to remove the ' + P.name + ' key.');
  if (k === null) return;
  if (!k) { pr.deleteProperty(P.key); ui.alert(P.name + ' key removed.'); return; }
  pr.setProperty(P.key, k);
  var p = aiProviders_().filter(function (x) { return x.id === id; })[0];
  try {
    aiCallOne_(p, [{ role: 'user', content: 'Reply with the word OK.' }], 200, false);
    ui.alert(P.name + ' works. Flow AI is on (' + aiProviders_().map(function (x) { return x.name; }).join(', then ') + ').');
  } catch (e) {
    var why = { AI_KEY: 'The key was not accepted. Check you copied all of it.', AI_BILLING: 'The key works but the account has no credit or free quota left. Add credit on the ' + P.name + ' website.', AI_BUSY: P.name + ' is busy right now. The key is saved; try again in a minute.' }[e.message] || ('The key is saved but a test question failed: ' + (e.detail || e.message));
    ui.alert(why);
  }
}
/** Sheet menu: how many Flow AI questions were asked and roughly what they cost. */
function menuAiUsage() {
  var sh = aiSheet_(), last = sh.getLastRow(), ui = SpreadsheetApp.getUi();
  var rows = last > 1 ? sh.getRange(2, 1, last - 1, 10).getValues() : [], since = Date.now() - 30 * 864e5, n = 0, usd = 0, shops = {}, by = {};
  var price = {}; Object.keys(AI_PROVIDERS).forEach(function (k) { price[AI_PROVIDERS[k].name] = AI_PROVIDERS[k].price; });
  rows.forEach(function (r) {
    if (new Date(r[0]).getTime() < since) return;
    n++; shops[r[1]] = 1; by[r[9] || '?'] = (by[r[9] || '?'] || 0) + 1;
    var pp = price[r[9]] || [0.7, 2.1]; usd += (+r[7] || 0) / 1e6 * pp[0] + (+r[8] || 0) / 1e6 * pp[1];
  });
  ui.alert('Flow AI, last 30 days', n + ' questions from ' + Object.keys(shops).length + ' shop(s).\nAnswered by: ' + (Object.keys(by).map(function (k) { return k + ' ' + by[k]; }).join(', ') || 'none') +
    '\nRough cost: $' + usd.toFixed(2) + ' (about ৳' + Math.round(usd * 122) + ') at default model prices. Exact charges are on each provider\'s website.', ui.ButtonSet.OK);
}
