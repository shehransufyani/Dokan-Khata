/* Flow AI (on hold until the owner adds AI keys). Add this file to the Apps Script project as FlowAI.gs,
   then set CONFIG.ai = true in index.html, rebuild and deploy. It needs the "connect to an external service" permission. */
function onOpenFlowAI_(menu) { return menu.addItem('Set Flow AI keys', 'menuSetAiKey').addItem('Flow AI usage and cost', 'menuAiUsage'); }
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
  if (p.id === 'openrouter') { headers['X-Title'] = 'Store Flow'; headers['HTTP-Referer'] = 'https://shehransufyani.github.io/StoreFlow/'; }
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
