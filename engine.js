/**
 * Store POS: Google Apps Script server.
 * Your data lives in this Google Sheet (one tab per table); seller emails go out from your Gmail.
 *
 * Setup (once): in the Sheet, Extensions -> Apps Script, paste this whole file, Save.
 * Reload the Sheet, then menu POS -> "Set up". Type your store password in the Settings tab, next to store_password.
 * Then Deploy -> New deployment -> Web app (Execute as: Me, Who has access: Anyone) and copy the Web app URL.
 */
const TZ = "Asia/Manila";
const DEFAULT_NOTIFY_DAYS = 3;  // used when the seller gave no notify date
const SHEETS = {
  Products: ["id", "barcode", "name", "price", "stock", "seller_id", "expiry", "mfg_date", "notify_on", "photo_at"],
  Sellers: ["id", "name", "email", "phone", "shelf", "rent", "low_stock", "rent_start", "renew_answer", "renew_for", "renew_at"],
  Sales: ["id", "uid", "ts", "total", "cash"],
  SaleItems: ["sale_id", "product_id", "name", "seller_id", "qty", "price"],
  Alerts: ["product_id", "kind", "ref"],
  Emails: ["id", "to_addr", "subject", "body", "created", "sent", "error", "seller_id"],
  Settings: ["key", "value"],
  Photos: ["product_id", "photo", "photo_at"],  // small JPEG per product; Products.photo_at says which version is current
};
const NUMERIC = new Set(["id", "price", "stock", "seller_id", "rent", "low_stock", "total", "cash", "sale_id", "product_id", "qty"]);
const DATES = new Set(["expiry", "mfg_date", "notify_on", "rent_start", "renew_for"]);  // yyyy-MM-dd
const TIMES = new Set(["ts", "created", "sent", "photo_at", "renew_at"]);  // yyyy-MM-ddTHH:mm:ss

// ---------- web app entry ----------

// Links in the rent reminder email land here: a renter says if they renew their shelf (see rentEmail_)
function doGet(e) {
  const q = (e && e.parameter) || {};
  if (q.a !== "renew") return ContentService.createTextOutput("Store POS server is running.");
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return renewPage_(q); } finally { lock.releaseLock(); }
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return json_({status: 400, error: "Bad request"}); }
  const pw = storePassword_();
  if (pw.length < 10) {
    return json_({status: 403, error: "Set the store password first: in your Google Sheet, Settings tab, type it next to store_password (10+ characters)"});
  }
  if (String(req.pin || "") !== pw) {
    Utilities.sleep(1500);  // slows password guessing (Apps Script can't see who is asking, so no per-device lockout)
    return json_({status: 401, error: "Wrong password"});
  }
  const fn = ROUTES[req.path];
  if (!fn) return json_({status: 404, error: "Not found"});
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);  // one request at a time: two tablets never overwrite each other
    const data = fn(req.body || {});
    return json_({status: 200, data: data === undefined ? {ok: true} : data});
  } catch (x) {
    return json_({status: x.status || 500, error: x.status ? x.message : "Server error: " + x.message});
  } finally {
    lock.releaseLock();
  }
}

const ROUTES = {
  "/api/state": () => state_(),
  "/api/report": report_,
  "/api/sales": salesList_,
  "/api/seller-report": sellerReport_,
  "/api/product-report": productReport_,
  "/api/pull": pull_,
  "/api/products": saveProduct_,
  "/api/products/delete": deleteProduct_,
  "/api/restock": restock_,
  "/api/photo": savePhoto_,
  "/api/photos": photos_,
  "/api/sellers": saveSeller_,
  "/api/sellers/delete": deleteSeller_,
  "/api/sellers/report": stockReport_,
  "/api/sellers/renew": renewSeller_,
  "/api/sale": sale_,
  "/api/settings": saveSettings_,
  "/api/outbox/delete": d => removeById_("Emails", d.id),
  "/api/outbox/readdress": readdress_,
  "/api/outbox/test": testEmail_,
  "/api/send-now": () => { checkAlerts_(); flush_(); return state_(); },
};

// ---------- Sheet menu + hourly job ----------

function onOpen() {
  SpreadsheetApp.getUi().createMenu("POS")
    .addItem("Set up", "setup")
    .addItem("Check seller alerts now", "hourly")
    .addToUi();
}

function setup() {
  // No pop-up password box (Google stops a script after 6 minutes): the password lives in a Settings cell instead
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(sheet_);
  const blank = ss.getSheetByName("Sheet1");
  if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);
  if (!rows_("Settings").some(r => r.key === "store_password")) insert_("Settings", {key: "store_password", value: ""});
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "hourly")) {
    ScriptApp.newTrigger("hourly").timeBased().everyHours(1).create();
  }
  ss.setActiveSheet(sheet_("Settings"));
  SpreadsheetApp.getUi().alert("Done. In this Settings tab, type your store password in the value cell next to store_password " +
    "(10+ characters). Every tablet and computer uses it once. You can change it here anytime.");
}

function storePassword_() {
  return String(settings_().store_password || "").trim();
}

function hourly() {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { checkAlerts_(); flush_(); } finally { lock.releaseLock(); }
}

// ---------- sheet helpers ----------

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange("A:Z").setNumberFormat("@");  // plain text: dates and barcodes stay exactly as typed
    sh.appendRow(SHEETS[name]);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < SHEETS[name].length) {
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);  // a newer version added columns
  }
  return sh;
}

function rows_(name) {
  const values = sheet_(name).getDataRange().getValues();
  const head = values.shift() || [];
  return values.map((r, i) => {
    const o = {_row: i + 2};
    head.forEach((k, j) => { o[k] = clean_(k, r[j]); });
    return o;
  });
}

// Google turns text that looks like a date or number into its own types (and drops a barcode's leading zeros).
// Writing text with a leading ' keeps it exactly as typed; reading turns any converted dates back into our format.
const cell_ = v => v === null || v === undefined ? "" : typeof v === "string" ? "'" + v : v;

function clean_(k, v) {
  if (v === "" || v === null || v === undefined) return null;
  if (typeof v === "string") v = v.replace(/^'/, "");  // in case the sheet kept the ' as a character
  if (NUMERIC.has(k)) return Number(v);
  if (DATES.has(k) || TIMES.has(k)) {
    const s = String(v), iso = DATES.has(k) ? /^\d{4}-\d{2}-\d{2}$/ : /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/;
    if (iso.test(s)) return s;
    const d = v instanceof Date ? v : new Date(s.replace(/^(\d{4}-\d{2}-\d{2}) (\d):/, "$1 0$2:"));
    if (!isNaN(d)) return Utilities.formatDate(d, TZ, DATES.has(k) ? "yyyy-MM-dd" : "yyyy-MM-dd'T'HH:mm:ss");
    return s;
  }
  return v instanceof Date ? Utilities.formatDate(v, TZ, "yyyy-MM-dd") : String(v);
}
const line_ = (name, o) => SHEETS[name].map(k => cell_(o[k]));

function insert_(name, o) { sheet_(name).appendRow(line_(name, o)); }

function insertMany_(name, list) {
  if (!list.length) return;
  const sh = sheet_(name);
  sh.getRange(sh.getLastRow() + 1, 1, list.length, SHEETS[name].length).setValues(list.map(o => line_(name, o)));
}

function update_(name, o) { sheet_(name).getRange(o._row, 1, 1, SHEETS[name].length).setValues([line_(name, o)]); }

function removeById_(name, id) {
  const r = rows_(name).find(x => x.id === Number(id));
  if (r) sheet_(name).deleteRow(r._row);
}

const nextId_ = list => list.reduce((m, r) => Math.max(m, r.id || 0), 0) + 1;
// Products and sellers can be created on any device, even offline, so their ids must not collide between devices
function uniqueId_(list) {
  let id = Date.now() * 10 + Math.floor(Math.random() * 10);
  while (list.some(r => r.id === id)) id++;
  return id;
}
const strip_ = o => { const c = Object.assign({}, o); delete c._row; return c; };
const byName_ = (a, b) => String(a.name).localeCompare(String(b.name));

function settings_() {
  const s = {};
  rows_("Settings").forEach(r => { s[r.key] = r.value || ""; });
  return s;
}

function setSetting_(key, value) {
  const r = rows_("Settings").find(x => x.key === key);
  if (r) update_("Settings", Object.assign(r, {value: String(value)}));
  else insert_("Settings", {key: key, value: String(value)});
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function bad_(msg) { const e = new Error(msg); e.status = 400; throw e; }

const now_ = () => Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
const today_ = () => now_().slice(0, 10);
const isDate_ = s => /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(Date.parse(s));
const addDays_ = (iso, n) => new Date(Date.parse(iso) + n * 864e5).toISOString().slice(0, 10);
const daysFrom_ = (a, b) => Math.round((Date.parse(b) - Date.parse(a)) / 864e5);

function num_(v, label) {
  const n = Number(v);
  if (v === "" || v === null || v === undefined || !isFinite(n)) bad_(label + " must be a number");
  return n;
}

// ---------- products / sellers ----------

function nextCode_(products) {
  // In-store barcode: 20 + 6 digits (the 20-29 prefix is reserved for in-store use, never on factory goods)
  let n = Number(settings_().last_code || 0), code;
  do { n++; code = "20" + String(n).padStart(6, "0"); } while (products.some(p => p.barcode === code));
  setSetting_("last_code", n);
  return code;
}

function saveProduct_(d) {
  const p = {
    id: d.id ? Number(d.id) : null, name: String(d.name || "").trim(), barcode: String(d.barcode || "").trim(),
    price: num_(d.price, "Price"), stock: num_(d.stock, "Stock"), seller_id: d.seller_id ? Number(d.seller_id) : null,
    expiry: d.expiry || null, mfg_date: d.mfg_date || null, notify_on: d.notify_on || null,
  };
  if (!p.name) bad_("Name is required");
  if (!Number.isInteger(p.stock)) bad_("Stock must be a whole number");
  if (p.price < 0 || p.stock < 0) bad_("Price and stock cannot be negative");
  ["expiry", "mfg_date", "notify_on"].forEach(k => { if (p[k] && !isDate_(p[k])) bad_("Bad date: " + p[k]); });
  if (p.mfg_date && p.mfg_date > today_()) bad_("Manufacturing date cannot be in the future");
  if (p.mfg_date && p.expiry && p.mfg_date > p.expiry) bad_("Manufacturing date is after the expiry date");
  if (p.notify_on && !p.expiry) bad_("Set the expiry date too, or clear the notify date");
  if (p.notify_on && p.notify_on > p.expiry) bad_("Notify date must be on or before the expiry date");
  if (p.barcode && !/^[\x20-\x7e]+$/.test(p.barcode)) bad_("Barcode must be plain letters/numbers");
  const all = rows_("Products");
  if (!p.barcode) p.barcode = nextCode_(all);
  if (all.some(r => r.barcode === p.barcode && r.id !== p.id)) bad_("That barcode is already used by another product");
  const old = p.id && all.find(r => r.id === p.id);
  if (old) {
    update_("Products", Object.assign(p, {_row: old._row, photo_at: old.photo_at}));  // the photo is saved separately
  } else {
    p.id = p.id || uniqueId_(all);  // a device that made it offline already chose its id
    insert_("Products", p);
  }
  return {id: p.id, barcode: p.barcode};
}

// Restock: ADD pieces to the stock Google has now (not overwrite it), so sales made meanwhile on other devices still
// count. A new batch can bring a new expiry date (and notify date).
function restock_(d) {
  const add = num_(d.add, "Pieces to add");
  if (!Number.isInteger(add) || add <= 0) bad_("Pieces to add must be a whole number above 0");
  const p = rows_("Products").find(r => r.id === Number(d.id));
  if (!p) bad_("Product not found (was it deleted?)");
  p.stock += add;
  if ("expiry" in d) {
    const e = d.expiry || null, n = d.notify_on || null;
    [e, n].forEach(v => { if (v && !isDate_(v)) bad_("Bad date: " + v); });
    if (n && !e) bad_("Set the expiry date too, or clear the notify date");
    if (n && n > e) bad_("Notify date must be on or before the expiry date");
    if (e && p.mfg_date && p.mfg_date > e) bad_("Manufacturing date is after the expiry date");
    p.expiry = e; p.notify_on = n;
  }
  update_("Products", p);
  return {stock: p.stock};
}

function deleteProduct_(d) {
  removeById_("Products", d.id);
  const ph = rows_("Photos").find(r => r.product_id === Number(d.id));
  if (ph) sheet_("Photos").deleteRow(ph._row);
}

// Product photo: a small JPEG made by the device (data URL). An empty photo removes it.
// `at` comes from the device, so its own copy and Google's carry the same version.
function savePhoto_(d) {
  const id = Number(d.id), photo = String(d.photo || "");
  if (photo && (photo.length > 45000 || !/^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(photo))) bad_("The photo is too big or not a picture");
  const p = rows_("Products").find(r => r.id === id);
  if (!p) bad_("Product not found (was it deleted?)");
  const at = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(d.at || "") ? d.at : now_();
  const old = rows_("Photos").find(r => r.product_id === id);
  if (old && photo) update_("Photos", Object.assign(old, {photo: photo, photo_at: at}));
  else if (old) sheet_("Photos").deleteRow(old._row);
  else if (photo) insert_("Photos", {product_id: id, photo: photo, photo_at: at});
  update_("Products", Object.assign(p, {photo_at: photo ? at : null}));
  return {at: at};
}

function photos_(q) {
  const ids = q.ids ? new Set(q.ids.map(Number)) : null;
  return rows_("Photos").filter(r => !ids || ids.has(r.product_id)).map(strip_);
}

function saveSeller_(d) {
  const s = {
    id: d.id ? Number(d.id) : null, name: String(d.name || "").trim(), email: String(d.email || "").trim(),
    phone: String(d.phone || ""), shelf: String(d.shelf || ""), rent: Number(d.rent) || 0,
    low_stock: d.low_stock === "" || d.low_stock == null ? 5 : num_(d.low_stock, "Low-stock alert"),
    rent_start: d.rent_start || null,
  };
  if (s.rent_start && !isDate_(s.rent_start)) bad_("Bad date: " + s.rent_start);
  if (!s.name) bad_("Name is required");
  if (s.email && s.email.indexOf("@") < 0) bad_("Email address looks wrong");
  const all = rows_("Sellers");
  const old = s.id && all.find(r => r.id === s.id);
  if (old) {  // the renter's renewal answer is kept (it is changed with renewSeller_ or their email link)
    update_("Sellers", Object.assign(s, {_row: old._row, renew_answer: old.renew_answer, renew_for: old.renew_for, renew_at: old.renew_at}));
  } else {
    s.id = s.id || uniqueId_(all);
    insert_("Sellers", s);
  }
  return {id: s.id};
}

function deleteSeller_(d) {
  if (rows_("Products").some(p => p.seller_id === Number(d.id))) {
    bad_("Seller still has products on the shelf. Remove or reassign them first.");
  }
  removeById_("Sellers", d.id);
}

// ---------- selling ----------

function sale_(d) {
  // Every sale is finished on the device first: the customer has left with the goods, so record it as it happened
  // (its prices and time; stock may go negative). The device checks stock and expiry before selling.
  // `uid` (made by the device) makes retries safe.
  const uid = d.uid || null;
  const sales = rows_("Sales");
  if (uid) {
    const old = sales.find(s => s.uid === uid);
    if (old) return {id: old.id, duplicate: true};
  }
  const byId = new Map();
  (d.items || []).forEach(it => {
    const id = Number(it.id), l = byId.get(id) || {qty: 0, price: it.price, name: it.name};
    l.qty += parseInt(it.qty, 10);
    byId.set(id, l);
  });
  if (!byId.size) bad_("Cart is empty");
  const products = rows_("Products"), lines = [];
  let total = 0;
  byId.forEach((it, pid) => {
    const p = products.find(r => r.id === pid);  // may be gone if it was deleted meanwhile
    if (!(it.qty > 0)) bad_("Bad quantity");
    const line = {p: p, id: pid, name: it.name || (p && p.name) || "?", price: Number(it.price), seller_id: p ? p.seller_id : null, qty: it.qty};
    if (!isFinite(line.price)) bad_("Bad price");
    total += line.price * it.qty;
    lines.push(line);
  });
  total = Math.round(total * 100) / 100;
  const cash = Number(d.cash) || total;
  const ts = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(d.ts || "") ? d.ts.slice(0, 19) : now_();
  const id = nextId_(sales);
  insert_("Sales", {id: id, uid: uid, ts: ts, total: total, cash: cash});
  insertMany_("SaleItems", lines.map(l => ({sale_id: id, product_id: l.id, name: l.name, seller_id: l.seller_id, qty: l.qty, price: l.price})));
  lines.forEach(l => { if (l.p) update_("Products", Object.assign(l.p, {stock: l.p.stock - l.qty})); });
  return {id: id, ts: ts, total: total, cash: cash, change: Math.round((cash - total) * 100) / 100,
          items: lines.map(l => ({name: l.name, qty: l.qty, price: l.price}))};
}

function report_(q) {
  const today = today_(), from = q.from || today, to = (q.to || today) + "T99";
  const sales = rows_("Sales").filter(s => s.ts >= from && s.ts <= to);
  const ids = new Set(sales.map(s => s.id));
  const items = rows_("SaleItems").filter(i => ids.has(i.sale_id));
  const sellers = rows_("Sellers");
  const group = (key, label) => {
    const m = new Map();
    items.forEach(i => {
      const k = key(i), g = m.get(k) || Object.assign(label(i), {qty: 0, amount: 0});
      g.qty += i.qty; g.amount += i.qty * i.price;
      m.set(k, g);
    });
    return Array.from(m.values()).sort((a, b) => b.amount - a.amount);
  };
  return {
    summary: {n: sales.length, total: sales.reduce((s, x) => s + x.total, 0)},
    by_owner: group(i => i.seller_id || 0, i => {
      const s = sellers.find(x => x.id === i.seller_id);
      return {owner: s ? s.name : "Store", shelf: s ? s.shelf : null};
    }),
    by_item: group(i => i.product_id + "|" + i.name, i => ({id: i.product_id, name: i.name})),
  };
}

function salesList_(q) {
  // Transaction history: newest first, with each sale's items (max 500 per request)
  const today = today_(), from = q.from || today, to = (q.to || today) + "T99";
  const sales = rows_("Sales").filter(s => s.ts >= from && s.ts <= to).sort((a, b) => b.ts.localeCompare(a.ts)).slice(0, 500);
  const ids = new Set(sales.map(s => s.id)), items = rows_("SaleItems").filter(i => ids.has(i.sale_id));
  return sales.map(s => Object.assign(strip_(s), {
    items: items.filter(i => i.sale_id === s.id).map(i => ({name: i.name, qty: i.qty, price: i.price})),
  }));
}

function sellerReport_(q) {
  // Shelf renters: per seller, what sold (every line), what they earned, rent, stock left, and the emails we sent them
  const today = today_(), from = q.from || today, to = (q.to || today) + "T99";
  const sales = new Map(rows_("Sales").filter(s => s.ts >= from && s.ts <= to).map(s => [s.id, s]));
  const items = rows_("SaleItems").filter(i => i.seller_id && sales.has(i.sale_id));
  const products = rows_("Products"), emails = rows_("Emails");
  return rows_("Sellers").sort(byName_).map(s => {
    const mine = items.filter(i => i.seller_id === s.id);
    const lines = mine.map(i => {
      const x = sales.get(i.sale_id);
      return {ts: x.ts, receipt: receiptNo_(x), name: i.name, qty: i.qty, price: i.price, amount: i.qty * i.price};
    }).sort((a, b) => b.ts.localeCompare(a.ts));
    const shelf = products.filter(p => p.seller_id === s.id).sort(byName_).map(p => {
      const sold = mine.filter(i => i.product_id === p.id);
      return {name: p.name, barcode: p.barcode, price: p.price, stock: p.stock, expiry: p.expiry,
              sold: sold.reduce((n, i) => n + i.qty, 0), amount: sold.reduce((n, i) => n + i.qty * i.price, 0)};
    });
    const mails = emails.filter(m => (m.seller_id ? m.seller_id === s.id : s.email && m.to_addr === s.email) && m.created >= from && m.created <= to)
      .map(m => ({created: m.created, sent: m.sent, subject: m.subject, error: m.error})).reverse();
    const pcs = lines.reduce((n, l) => n + l.qty, 0), amount = lines.reduce((n, l) => n + l.amount, 0), rent = s.rent || 0;
    return {seller: strip_(s), summary: {pcs: pcs, amount: amount, rent: rent, payout: amount - rent}, products: shelf, sales: lines, emails: mails};
  });
}

const receiptNo_ = s => s.uid ? s.uid.slice(-6).toUpperCase() : String(s.id);  // same number as on the printed receipt

function productReport_(q) {
  // Every product (or one, with q.id): stock now, price, and what sold in the period; for one product also each sale line
  const today = today_(), from = q.from || today, to = (q.to || today) + "T99", one = q.id ? Number(q.id) : null;
  const sales = new Map(rows_("Sales").filter(s => s.ts >= from && s.ts <= to).map(s => [s.id, s]));
  const items = rows_("SaleItems").filter(i => sales.has(i.sale_id) && (!one || i.product_id === one));
  const sellers = rows_("Sellers"), products = rows_("Products").filter(p => !one || p.id === one).sort(byName_);
  if (one && !products.length) bad_("Product not found (was it deleted?)");
  const list = products.map(p => {
    const mine = items.filter(i => i.product_id === p.id), s = sellers.find(x => x.id === p.seller_id);
    return {id: p.id, name: p.name, barcode: p.barcode, price: p.price, stock: p.stock, value: p.price * p.stock,
            owner: s ? s.name : "Store", shelf: s ? s.shelf : null, expiry: p.expiry, mfg_date: p.mfg_date,
            sold: mine.reduce((n, i) => n + i.qty, 0), amount: mine.reduce((n, i) => n + i.qty * i.price, 0), sales: new Set(mine.map(i => i.sale_id)).size};
  });
  if (!one) {  // items sold in the period whose product was deleted since, so the totals match the sales report
    const gone = new Map();
    items.filter(i => !products.some(p => p.id === i.product_id)).forEach(i => {
      const k = i.product_id + "|" + i.name, s = sellers.find(x => x.id === i.seller_id);
      const g = gone.get(k) || {id: i.product_id, name: i.name, deleted: true, barcode: "", price: i.price, stock: null, value: 0,
        owner: s ? s.name : "Store", shelf: s ? s.shelf : null, expiry: null, mfg_date: null, sold: 0, amount: 0, ids: new Set()};
      g.sold += i.qty; g.amount += i.qty * i.price; g.ids.add(i.sale_id);
      gone.set(k, g);
    });
    gone.forEach(g => { g.sales = g.ids.size; delete g.ids; list.push(g); });
  }
  const lines = !one ? [] : items.map(i => {
    const x = sales.get(i.sale_id);
    return {ts: x.ts, receipt: receiptNo_(x), qty: i.qty, price: i.price, amount: i.qty * i.price};
  }).sort((a, b) => b.ts.localeCompare(a.ts));
  return {products: list, lines: lines};
}

function pull_(q) {
  // For devices keeping their own copy: everything current, plus the sales they haven't seen yet (in batches)
  const after = Number(q.after) || 0, st = settings_(), settings = {};
  Object.keys(st).forEach(k => { if (k !== "store_password" && k !== "renew_key") settings[k] = st[k]; });  // secrets stay on Google
  const newer = rows_("Sales").filter(s => s.id > after).sort((a, b) => a.id - b.id);
  const batch = newer.slice(0, 1500), ids = new Set(batch.map(s => s.id));
  const items = rows_("SaleItems").filter(i => ids.has(i.sale_id));
  return {
    products: rows_("Products").map(strip_), sellers: rows_("Sellers").map(strip_), settings: settings,
    emails: rows_("Emails").slice(-200).map(m => { const c = strip_(m); delete c.body; return c; }),
    sales: batch.map(s => Object.assign(strip_(s), {items: items.filter(i => i.sale_id === s.id).map(strip_)})),
    more: newer.length > batch.length, quota: MailApp.getRemainingDailyQuota(),
  };
}

// ---------- seller emails ----------

function queueEmail_(to, subject, body, sellerId) {
  insert_("Emails", {id: nextId_(rows_("Emails")), to_addr: to, subject: subject, body: body, created: now_(), seller_id: sellerId || null});
}

function stockList_(items) {
  return items.map(p => {
    const d = [];
    if (p.mfg_date) d.push("made " + p.mfg_date);
    if (p.expiry) d.push("expires " + p.expiry);
    return "  - " + p.name + ": " + p.stock + " left" + (d.length ? " (" + d.join(", ") + ")" : "");
  }).join("\n") || "  (no items on your shelf)";
}

function sellerEmail_(s, items, alerts) {
  const store = settings_().store_name || "Urdaneta Store", shelf = s.shelf ? " " + s.shelf : "";
  const body = "Hi " + s.name + ",\n\nThis is " + store + ". Update on your rented shelf" + shelf + ":\n\n" +
    (alerts.length ? alerts.join("\n") + "\n\n" : "") +
    "Stock on your shelf as of " + Utilities.formatDate(new Date(), TZ, "MMM dd, yyyy") + ":\n" + stockList_(items) +
    "\n\nThank you!\n" + store;
  queueEmail_(s.email, store + ": " + (alerts.length ? "ACTION NEEDED - " : "") + "your shelf" + shelf + " stock update", body, s.id);
}

function checkAlerts_(today) {
  // One email per seller when an item hits its notify date, expires, or runs low.
  // Each alert fires once; a new expiry date or a restock re-arms it.
  today = today || today_();
  const products = rows_("Products"), alerts = rows_("Alerts"), clear = [];
  const add = (pid, kind, ref) => {
    if (alerts.some(a => a.product_id === pid && a.kind === kind && (a.ref || "") === ref)) return false;
    insert_("Alerts", {product_id: pid, kind: kind, ref: ref});
    alerts.push({product_id: pid, kind: kind, ref: ref});
    return true;
  };
  let queued = 0;
  rows_("Sellers").filter(s => s.email).forEach(s => {
    const items = products.filter(p => p.seller_id === s.id).sort(byName_), lines = [];
    items.forEach(p => {
      if (p.expiry && p.stock > 0) {
        const notify = p.notify_on || addDays_(p.expiry, -DEFAULT_NOTIFY_DAYS);
        if (today > p.expiry) {
          if (add(p.id, "expired", p.expiry)) lines.push("!! EXPIRED on " + p.expiry + ": " + p.name + " - " + p.stock + " pcs. Please pull out.");
        } else if (notify <= today && add(p.id, "expiring", p.expiry)) {
          const left = daysFrom_(today, p.expiry);
          lines.push("! Expires " + (left === 0 ? "TODAY" : "in " + left + " day(s)") + " (" + p.expiry + "): " + p.name + " - " + p.stock + " pcs left.");
        }
      }
      if (p.stock <= s.low_stock) {
        if (add(p.id, "low", "")) lines.push("! Low stock: " + p.name + " - only " + p.stock + " left. Please restock.");
      } else {
        alerts.filter(a => a.product_id === p.id && a.kind === "low" && a._row).forEach(a => clear.push(a._row));
      }
    });
    if (lines.length) { sellerEmail_(s, items, lines); queued++; }
  });
  clear.sort((a, b) => b - a).forEach(r => sheet_("Alerts").deleteRow(r));  // bottom-up so row numbers stay valid
  // Rent: a few days before each monthly due date the renter is asked if they renew, and the owner gets a reminder.
  // Once per due date (the Alerts row uses the seller id); not when the renter already answered for that date.
  const notice = Number(settings_().rent_notice_days) || 3;
  rows_("Sellers").filter(s => s.rent_start).forEach(s => {
    const due = nextDue_(s.rent_start, today), left = daysFrom_(today, due);
    if (left > notice || (s.renew_for === due && s.renew_answer) || !add(s.id, "rent", due)) return;
    if (s.email) rentEmail_(s, due, left);
    ownerEmail_(s, "rent due " + dateText_(due), "The monthly rent for " + s.name + (s.shelf ? " (shelf " + s.shelf + ")" : "") +
      " is due on " + dateText_(due) + " (" + daysText_(left) + "), PHP " + Number(s.rent || 0).toFixed(2) + ".\n\n" +
      (s.email ? "We emailed " + s.name + " to ask if they renew; their answer will show in the POS (Sellers & Shelves)."
               : s.name + " has no email address in the POS, so please ask them yourself."));
    queued++;
  });
  return queued;
}

// ---------- shelf rent ----------

// Rent is due every month on the day the rental started (the last day of shorter months). Next due date, today or later.
function nextDue_(start, today) {
  if (!start) return null;
  const p = start.split("-").map(Number);
  for (let i = 1; i < 1200; i++) {
    const m = p[1] - 1 + i, y = p[0] + Math.floor(m / 12), mo = m % 12, last = new Date(Date.UTC(y, mo + 1, 0)).getUTCDate();
    const due = y + "-" + String(mo + 1).padStart(2, "0") + "-" + String(Math.min(p[2], last)).padStart(2, "0");
    if (due >= today) return due;
  }
  return null;
}
const dateText_ = iso => Utilities.formatDate(new Date(iso + "T12:00:00+08:00"), TZ, "MMM dd, yyyy");
const daysText_ = n => n === 0 ? "today" : n === 1 ? "tomorrow" : "in " + n + " days";
const htmlEsc_ = s => String(s).replace(/[&<>"]/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"}[c]));

// A short signature in each answer link, so nobody can answer for another renter
function renewKey_(id, due, answer) {
  let key = settings_().renew_key;
  if (!key) { key = Utilities.getUuid(); setSetting_("renew_key", key); }
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(id + "|" + due + "|" + answer, key)).replace(/=+$/, "").slice(0, 22);
}
function renewLink_(s, due, answer) {
  const base = settings_().web_app_url || ScriptApp.getService().getUrl();
  return base + "?a=renew&s=" + s.id + "&d=" + due + "&v=" + answer + "&k=" + renewKey_(s.id, due, answer);
}

function rentEmail_(s, due, left) {
  const store = settings_().store_name || "Urdaneta Store", shelf = s.shelf ? " " + s.shelf : "";
  const body = "Hi " + s.name + ",\n\nThis is " + store + ". Your current month on shelf" + shelf + " ends on " + dateText_(due) +
    " (" + daysText_(left) + "). The monthly rent is PHP " + Number(s.rent || 0).toFixed(2) + ".\n\n" +
    "Would you like to keep the shelf for another month? Please tap one:\n\n" +
    "  YES, I will renew:\n  " + renewLink_(s, due, "yes") + "\n\n" +
    "  NO, I will end my rental:\n  " + renewLink_(s, due, "no") + "\n\n" +
    "You can also simply reply to this email.\n\nThank you!\n" + store;
  queueEmail_(s.email, store + ": your shelf" + shelf + " rent is due " + dateText_(due) + " - will you renew?", body, s.id);
}

function ownerEmail_(s, what, body) {
  const store = settings_().store_name || "Urdaneta Store";
  queueEmail_(Session.getEffectiveUser().getEmail(), store + ": " + s.name + " - " + what, body + "\n\n" + store + " POS", null);
}

// The renter's answer: the link opens a page asking to confirm (so a mail scanner opening links can't answer for them)
function renewPage_(q) {
  const page = (title, msg, extra) => HtmlService.createHtmlOutput('<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<div style="font:17px/1.5 Arial,sans-serif;max-width:460px;margin:48px auto;padding:0 20px;text-align:center;color:#0f0f0f">' +
    '<h2 style="font-size:22px">' + title + "</h2><p>" + msg + "</p>" + (extra || "") + "</div>").setTitle("Shelf rent");
  const s = rows_("Sellers").find(r => r.id === Number(q.s));
  const valid = s && (q.v === "yes" || q.v === "no") && /^\d{4}-\d{2}-\d{2}$/.test(q.d || "") && q.k === renewKey_(s.id, q.d, q.v);
  if (!valid) return page("Link not valid", "This link doesn't work any more. Please contact the store.");
  const store = htmlEsc_(settings_().store_name || "Urdaneta Store"), shelf = s.shelf ? " " + htmlEsc_(s.shelf) : "";
  const yes = q.v === "yes";
  if (q.c !== "1") {
    const btn = '<a target="_top" href="' + htmlEsc_(renewLink_(s, q.d, q.v) + "&c=1") + '" style="display:inline-block;margin-top:12px;padding:14px 28px;border-radius:24px;' +
      'background:' + (yes ? "#0f0f0f" : "#cc0000") + ';color:#fff;text-decoration:none;font-weight:bold">' + (yes ? "Yes, renew my shelf" : "No, end my rental") + "</a>";
    return page(store, htmlEsc_(s.name) + ", please confirm for shelf" + shelf + " (rent due " + dateText_(q.d) + "):", btn);
  }
  if (!(s.renew_for === q.d && s.renew_answer === q.v)) {
    update_("Sellers", Object.assign(s, {renew_answer: q.v, renew_for: q.d, renew_at: now_()}));
    ownerEmail_(s, yes ? "will RENEW" : "will NOT renew", s.name + " answered " + (yes ? "YES, they will renew" : "NO, they will end their rental") +
      (s.shelf ? " for shelf " + s.shelf : "") + " (rent due " + dateText_(q.d) + ").");
    flush_();
  }
  return page("Thank you, " + htmlEsc_(s.name) + "!", yes ? "We noted that you will keep shelf" + shelf + " for another month. See you at " + store + "."
    : "We noted that you will end your rental of shelf" + shelf + ". " + store + " will contact you about collecting your items.");
}

// The store records an answer itself (the renter replied by text or in person); "" clears it
function renewSeller_(d) {
  const s = rows_("Sellers").find(r => r.id === Number(d.id));
  if (!s) bad_("Seller not found (was it deleted?)");
  if (!s.rent_start) bad_("Set the rental start date first");
  const answer = String(d.answer || "");
  if (answer && answer !== "yes" && answer !== "no") bad_("Answer must be yes or no");
  update_("Sellers", Object.assign(s, {renew_answer: answer || null, renew_for: answer ? nextDue_(s.rent_start, today_()) : null, renew_at: answer ? now_() : null}));
}

function stockReport_(d) {
  const sellers = rows_("Sellers").filter(s => s.email && (!d.id || s.id === Number(d.id)));
  const products = rows_("Products");
  sellers.forEach(s => sellerEmail_(s, products.filter(p => p.seller_id === s.id).sort(byName_), []));
  flush_();
  return {queued: sellers.length};
}

function flush_() {
  const cfg = settings_(), name = cfg.smtp_from || cfg.store_name || "Store";
  rows_("Emails").filter(m => !m.sent).forEach(m => {
    try {
      MailApp.sendEmail({to: m.to_addr, subject: m.subject, body: m.body, name: name});
      update_("Emails", Object.assign(m, {sent: now_(), error: null}));
    } catch (x) {
      const msg = String(x.message || x);
      const bad = /invalid email/i.test(msg);  // a wrong address won't fix itself: stop retrying
      update_("Emails", Object.assign(m, {sent: bad ? now_() : null, error: (bad ? "failed: " : "") + msg.slice(0, 300)}));
    }
  });
}

function readdress_(d) {
  const to = String(d.to_addr || "").trim();
  if (to.indexOf("@") < 0 || / /.test(to)) bad_("Email address looks wrong");
  const m = rows_("Emails").find(x => x.id === Number(d.id));
  if (!m || m.sent) bad_("That email was already sent and can't be changed");
  update_("Emails", Object.assign(m, {to_addr: to, error: null}));
}

function testEmail_() {
  const store = settings_().store_name || "Urdaneta Store";
  queueEmail_(Session.getEffectiveUser().getEmail(), store + ": test email",
              "This is a test from your " + store + " POS.\n\nIf you can read this, seller alerts will be delivered the same way.");
  flush_();
}

// ---------- settings / state ----------

const PUBLIC_SETTINGS = ["store_name", "smtp_from", "receipt_address", "receipt_phone", "receipt_footer", "rent_notice_days", "web_app_url"];

function saveSettings_(d) {
  PUBLIC_SETTINGS.forEach(k => { if (k in d) setSetting_(k, String(d[k] || "").trim()); });
}

function state_() {
  const st = settings_();
  return {
    products: rows_("Products").sort(byName_).map(strip_),
    sellers: rows_("Sellers").sort(byName_).map(strip_),
    settings: PUBLIC_SETTINGS.reduce((o, k) => { o[k] = st[k] || ""; return o; }, {}),  // never the password
    outbox: rows_("Emails").slice(-20).reverse().map(m => { const c = strip_(m); delete c.body; return c; }),
    quota: MailApp.getRemainingDailyQuota(),
  };
}
