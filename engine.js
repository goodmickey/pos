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
  Products: ["id", "barcode", "name", "price", "stock", "seller_id", "expiry", "mfg_date", "notify_on", "photo_at", "cost"],  // cost: what the store paid per piece
  Sellers: ["id", "name", "email", "phone", "shelf", "rent", "low_stock", "rent_start", "renew_answer", "renew_for", "renew_at", "commission", "deleted"],  // commission: % the store keeps of their sales
  Sales: ["id", "uid", "ts", "total", "cash", "pay_via", "epay", "epay_ref", "refund_of", "reason", "credit_to"],  // pay_via/epay: GCash, Maya...; refund_of: uid
  SaleItems: ["sale_id", "product_id", "name", "seller_id", "qty", "price", "cost"],  // cost at the time of the sale
  Alerts: ["product_id", "kind", "ref"],
  Emails: ["id", "to_addr", "subject", "body", "created", "sent", "error", "seller_id"],
  Settings: ["key", "value"],
  Photos: ["product_id", "photo", "photo_at"],
  Drawer: ["uid", "ts", "kind", "amount", "note", "expected", "counted", "detail"],  // cash drawer: open, in, out, close
  CreditPays: ["uid", "ts", "sale_uid", "name", "amount", "via", "note"],  // money received later for a Pay later sale
  Texts: ["id", "to_phone", "body", "created", "sent", "error", "seller_id", "uid"],  // SMS: sent by the tablet with the SIM (Google can't)
  Payouts: ["uid", "ts", "seller_id", "name", "from_ts", "to_ts", "sales", "commission", "rent", "paid", "via", "note"],  // money settled with a renter
  Adjustments: ["uid", "ts", "product_id", "name", "seller_id", "qty", "price", "reason", "note", "cost"],  // stock removed (qty < 0) and why  // small JPEG per product; Products.photo_at says which version is current
};
const NUMERIC = new Set(["id", "price", "stock", "seller_id", "rent", "low_stock", "total", "cash", "epay", "sale_id", "product_id", "qty", "amount", "expected", "counted", "cost", "commission", "sales", "paid"]);
const PAY_VIA = ["GCash", "Maya", "Bank / QR"];  // paid by phone: the exact amount, checked by its reference number
const DATES = new Set(["expiry", "mfg_date", "notify_on", "rent_start", "renew_for"]);  // yyyy-MM-dd
const TIMES = new Set(["ts", "created", "sent", "photo_at", "renew_at", "from_ts", "to_ts", "deleted"]);  // yyyy-MM-ddTHH:mm:ss

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
  "/api/adjust": adjust_,
  "/api/photo": savePhoto_,
  "/api/photos": photos_,
  "/api/sellers": saveSeller_,
  "/api/sellers/delete": deleteSeller_,
  "/api/sellers/report": stockReport_,
  "/api/sellers/renew": renewSeller_,
  "/api/payouts": payoutState_,
  "/api/payout": payout_,
  "/api/credits": credits_,
  "/api/credit-pay": creditPay_,
  "/api/texts": () => rows_("Texts").filter(t => !t.sent && !t.error).map(strip_),
  "/api/text-sent": textSent_,
  "/api/status-text": statusText_,
  "/api/sms-preview": smsPreview_,
  "/api/text": sendText_,
  "/api/sale": sale_,
  "/api/drawer": drawerMove_,
  "/api/drawer-state": drawerState_,
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

function nextCode_(products, id) {
  // In-store barcode: 2 + 7 digits (the 20-29 prefix is reserved for in-store use, never on factory goods). Taken from
  // the product's own id, so two devices adding products offline never hand out the same code.
  let n = Number(String(id).slice(-7)), code;
  do { code = "2" + String(n % 1e7).padStart(7, "0"); n++; } while (products.some(p => p.barcode === code));
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
  const cost = d.cost === undefined || d.cost === null || d.cost === "" ? null : num_(d.cost, "Cost");
  if (cost !== null && cost < 0) bad_("Cost cannot be negative");
  ["expiry", "mfg_date", "notify_on"].forEach(k => { if (p[k] && !isDate_(p[k])) bad_("Bad date: " + p[k]); });
  if (p.mfg_date && p.mfg_date > today_()) bad_("Manufacturing date cannot be in the future");
  if (p.mfg_date && p.expiry && p.mfg_date > p.expiry) bad_("Manufacturing date is after the expiry date");
  if (p.notify_on && !p.expiry) bad_("Set the expiry date too, or clear the notify date");
  if (p.notify_on && p.notify_on > p.expiry) bad_("Notify date must be on or before the expiry date");
  if (p.barcode && !/^[\x20-\x7e]+$/.test(p.barcode)) bad_("Barcode must be plain letters/numbers");
  const all = rows_("Products");
  // A retry of an edit that changed stock: already applied (its uid is in Adjustments)
  if (d.uid && rows_("Adjustments").some(a => a.uid === String(d.uid))) return {id: p.id, barcode: p.barcode};
  if (!p.barcode) p.barcode = nextCode_(all, p.id || (p.id = uniqueId_(all)));
  if (all.some(r => r.barcode === p.barcode && r.id !== p.id)) bad_("That barcode is already used by another product");
  const old = p.id && all.find(r => r.id === p.id);
  if (!old && d.stock_was !== undefined && d.stock_was !== "") bad_("Product not found (was it deleted?)");  // an edit, not a new product
  if (old) {
    // The edit form sends the stock it showed when opened (stock_was): only the change typed there is applied, so
    // sales made meanwhile on this or another device still count. (Old app versions send no stock_was: stock as typed.)
    if (d.stock_was !== undefined && d.stock_was !== "") {
      const change = p.stock - num_(d.stock_was, "Stock");
      p.stock = (Number(old.stock) || 0) + change;
      if (change && d.uid) insert_("Adjustments", {uid: String(d.uid), ts: now_(), product_id: old.id, name: p.name, seller_id: p.seller_id, qty: change, price: p.price, reason: "Edit", note: null, cost: null});
    }
    update_("Products", Object.assign(p, {_row: old._row, photo_at: old.photo_at, cost: "cost" in d ? cost : old.cost}));  // the photo is saved separately
  } else {
    p.id = p.id || uniqueId_(all);  // a device that made it offline already chose its id
    p.cost = cost;
    insert_("Products", p);
  }
  return {id: p.id, barcode: p.barcode};
}

// Restock: ADD pieces to the stock Google has now (not overwrite it), so sales made meanwhile on other devices still
// count. A new batch can bring a new expiry date (and notify date).
function restock_(d) {
  const uid = d.uid ? String(d.uid) : null;
  if (uid && rows_("Adjustments").some(a => a.uid === uid)) return {stock: (rows_("Products").find(r => r.id === Number(d.id)) || {}).stock, duplicate: true};
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
  if (uid) insert_("Adjustments", {uid: uid, ts: now_(), product_id: p.id, name: p.name, seller_id: p.seller_id, qty: add, price: p.price, reason: "Restock", note: null, cost: p.cost === undefined ? null : p.cost});
  return {stock: p.stock};
}

// Stock taken off without a sale: spoiled, damaged, lost... Like restock it is a change (-n), so sales made at the same
// time on other devices still count; each one is kept (with the price then) for the "Stock removed" report.
const ADJUST_REASONS = ["Spoiled / expired", "Damaged", "Lost / stolen", "Store use", "Returned to renter", "Count correction"];
function adjust_(d) {
  const uid = d.uid ? String(d.uid) : null;
  if (uid && rows_("Adjustments").some(a => a.uid === uid)) return {duplicate: true};  // a retry: already recorded
  const qty = num_(d.remove, "Pieces to remove");
  if (!Number.isInteger(qty) || qty <= 0) bad_("Pieces to remove must be a whole number above 0");
  const reason = String(d.reason || "");
  if (ADJUST_REASONS.indexOf(reason) < 0) bad_("Choose why the stock is removed");
  const p = rows_("Products").find(r => r.id === Number(d.id));
  if (!p) bad_("Product not found (was it deleted?)");
  p.stock -= qty;
  update_("Products", p);
  const ts = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(d.ts || "") ? d.ts.slice(0, 19) : now_();
  insert_("Adjustments", {uid: uid, ts: ts, product_id: p.id, name: p.name, seller_id: p.seller_id, qty: -qty, price: p.price,
    reason: reason, note: String(d.note || "").trim().slice(0, 200) || null, cost: p.cost === undefined ? null : p.cost});
  return {stock: p.stock};
}

// Stock removed in a period: by reason, and every removal (value at the selling price then)
function removed_(from, to, sellers) {
  const lines = rows_("Adjustments").filter(a => a.qty < 0 && a.reason !== "Edit" && a.ts >= from && a.ts <= to).sort((a, b) => b.ts.localeCompare(a.ts)).map(a => {
    const s = sellers.find(x => x.id === a.seller_id);
    return {ts: a.ts, id: a.product_id, name: a.name, owner: s ? sellerName_(s) : "Store", qty: -a.qty, reason: a.reason, note: a.note, value: r2_(-a.qty * a.price),
            cost: a.cost === null || a.cost === undefined ? null : r2_(-a.qty * a.cost)};
  });
  const m = {};
  lines.forEach(l => {
    const g = m[l.reason] = m[l.reason] || {reason: l.reason, qty: 0, value: 0, cost: 0};
    g.qty += l.qty; g.value = r2_(g.value + l.value); g.cost = r2_(g.cost + (l.cost || 0));
  });
  return {by_reason: Object.keys(m).map(k => m[k]).sort((a, b) => b.value - a.value), lines: lines.slice(0, 300)};
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
  const commission = d.commission === undefined || d.commission === null || d.commission === "" ? 0 : num_(d.commission, "Commission");
  if (commission < 0 || commission > 100) bad_("Commission must be from 0 to 100 percent");
  if (s.rent_start && !isDate_(s.rent_start)) bad_("Bad date: " + s.rent_start);
  if (!s.name) bad_("Name is required");
  if (s.email && s.email.indexOf("@") < 0) bad_("Email address looks wrong");
  const all = rows_("Sellers");
  const old = s.id && all.find(r => r.id === s.id);
  if (old) {  // the renter's renewal answer is kept (it is changed with renewSeller_ or their email link)
    update_("Sellers", Object.assign(s, {_row: old._row, renew_answer: old.renew_answer, renew_for: old.renew_for, renew_at: old.renew_at,
      commission: "commission" in d ? commission : old.commission, deleted: old.deleted}));
  } else {
    s.id = s.id || uniqueId_(all);
    s.commission = commission;
    insert_("Sellers", s);
  }
  return {id: s.id};
}

// A deleted seller leaves the POS with their products, but their row stays (marked deleted): past sales and payouts
// keep their name in Reports instead of turning into the store's own sales.
const liveSellers_ = () => rows_("Sellers").filter(s => !s.deleted);
const sellerName_ = s => s.deleted ? s.name + " (deleted)" : s.name;
function deleteSeller_(d) {
  const id = Number(d.id), s = rows_("Sellers").find(x => x.id === id);
  if (!s || s.deleted) return;  // a retry: already done
  rows_("Products").filter(p => p.seller_id === id).reverse().forEach(p => deleteProduct_({id: p.id}));
  update_("Sellers", Object.assign(s, {deleted: now_()}));
  rows_("Texts").filter(t => t.seller_id === id && !t.sent && !t.error).forEach(t => update_("Texts", Object.assign(t, {error: "Seller deleted"})));
}

// ---------- selling ----------

function sale_(d) {
  // Every sale is finished on the device first: the customer has left with the goods, so record it as it happened
  // (its prices and time; stock may go negative). The device checks stock and expiry before selling.
  // `uid` (made by the device) makes retries safe.
  // A refund (refund_of = the sale's uid) is a sale with negative quantities: every report, renter payout and the
  // cash drawer then net it out by themselves. Its prices come from the original sale.
  const uid = d.uid || null;
  const sales = rows_("Sales");
  if (uid) {
    const old = sales.find(s => s.uid === uid);
    if (old) return {id: old.id, duplicate: true};
  }
  const refundOf = d.refund_of ? String(d.refund_of) : null, reason = refundOf ? String(d.reason || "").trim().slice(0, 200) : null;
  let origItems = [], already = new Map();  // already refunded: product id -> pieces
  let creditTo = String(d.credit_to || "").trim().slice(0, 80) || null;  // Pay later: who will pay
  let creditBalance = null;  // a refund of a Pay later sale that is still owed: how much
  if (refundOf) {
    const orig = sales.find(s => s.uid === refundOf);
    if (!orig) bad_("The sale to refund was not found");
    if (orig.refund_of) bad_("A refund cannot be refunded");
    if (!reason) bad_("Say why the sale is refunded");
    creditTo = orig.credit_to || null;  // items brought back from a Pay later sale: no money moves, they just owe less
    if (creditTo) {  // ...while they still owe. Paid already: a normal refund (money goes back)
      const owed = orig.total + sales.filter(x => x.refund_of === refundOf).reduce((n, x) => n + x.total, 0)
        - rows_("CreditPays").filter(x => x.sale_uid === refundOf).reduce((n, x) => n + x.amount, 0);
      if (owed <= 0.001) creditTo = null;
      else creditBalance = owed;
    }
    const prior = new Set(sales.filter(s => s.refund_of === refundOf).map(s => s.id));
    rows_("SaleItems").forEach(i => {
      if (i.sale_id === orig.id) origItems.push(i);
      if (prior.has(i.sale_id)) already.set(i.product_id, (already.get(i.product_id) || 0) - i.qty);
    });
  }
  const byId = new Map();
  (d.items || []).forEach(it => {
    const id = Number(it.id), l = byId.get(id) || {qty: 0, price: it.price, name: it.name, seller_id: it.seller_id, cost: it.cost};
    l.qty += parseInt(it.qty, 10);
    byId.set(id, l);
  });
  if (!byId.size) bad_("Cart is empty");
  const products = rows_("Products"), lines = [];
  let total = 0;
  byId.forEach((it, pid) => {
    const p = products.find(r => r.id === pid);  // may be gone if it was deleted meanwhile
    if (refundOf ? !(it.qty < 0) : !(it.qty > 0)) bad_("Bad quantity");
    const line = {p: p, id: pid, name: it.name || (p && p.name) || "?", price: Number(it.price), seller_id: p ? p.seller_id : Number(it.seller_id) || null, qty: it.qty,
                  cost: p && p.cost !== null && p.cost !== undefined ? p.cost : it.cost === undefined || it.cost === null || it.cost === "" ? null : Number(it.cost)};
    if (refundOf) {
      const o = origItems.filter(i => i.product_id === pid)[0];
      if (!o) bad_("That item is not on the sale being refunded");
      const left = o.qty - (already.get(pid) || 0);
      if (-it.qty > left) bad_("Only " + left + " of " + o.name + " can still be refunded");
      Object.assign(line, {name: o.name, price: o.price, seller_id: o.seller_id, cost: o.cost === undefined ? null : o.cost});  // as it was sold (the renter it was sold for)
    }
    if (!isFinite(line.price)) bad_("Bad price");
    total += line.price * it.qty;
    lines.push(line);
  });
  total = Math.round(total * 100) / 100;
  if (creditBalance !== null && -total > creditBalance + 0.001) bad_("Only " + smsMoney_(creditBalance) + " is still owed on this sale. Refund up to that now; refund the rest after they pay.");
  // Paid by GCash / Maya / bank QR (all, or part with the rest in cash): the amount and its reference number.
  // A refund goes back all in cash, or all by the e-wallet (the store sends it).
  const via = creditTo ? null : String(d.pay_via || "").trim() || null;
  const epay = via ? (refundOf ? total : Math.round(Number(d.epay) * 100) / 100) : 0;
  const ref = via ? String(d.epay_ref || "").trim().slice(0, 40) : null;
  if (via && PAY_VIA.indexOf(via) < 0) bad_("Unknown payment: " + via);
  if (via && !refundOf && !(epay > 0 && epay <= total)) bad_(via + " amount must be more than 0 and at most the total");
  if (via && !refundOf && ref.length < 4) bad_("Enter the " + via + " reference number (at least the last 4 digits)");
  const cash = refundOf || creditTo ? 0 : via ? Number(d.cash) || 0 : Number(d.cash) || total;  // cash received (older apps always sent it)
  const ts = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(d.ts || "") ? d.ts.slice(0, 19) : now_();
  const id = nextId_(sales);
  insert_("Sales", {id: id, uid: uid, ts: ts, total: total, cash: cash, pay_via: via, epay: epay || null, epay_ref: ref || null, refund_of: refundOf, reason: reason, credit_to: creditTo});
  insertMany_("SaleItems", lines.map(l => ({sale_id: id, product_id: l.id, name: l.name, seller_id: l.seller_id, qty: l.qty, price: l.price, cost: l.cost})));
  const restock = !refundOf || d.restock !== false;  // refunded goods go back on the shelf unless spoiled / damaged
  lines.forEach(l => { if (l.p && restock) update_("Products", Object.assign(l.p, {stock: l.p.stock - l.qty})); });
  return {id: id, ts: ts, total: total, cash: cash, pay_via: via, epay: epay, epay_ref: ref, refund_of: refundOf, reason: reason, credit_to: creditTo,
          change: creditTo ? 0 : Math.round((cash + epay - total) * 100) / 100,  // for a refund: the cash given back
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
    summary: {n: sales.filter(x => !x.refund_of).length, total: r2_(sales.reduce((s, x) => s + x.total, 0)),  // total is after refunds
              refunds: {n: sales.filter(x => x.refund_of).length, amount: r2_(-sales.filter(x => x.refund_of).reduce((s, x) => s + x.total, 0))}},
    by_pay: payments_(sales),
    removed: removed_(from, to, sellers),
    collected: rows_("CreditPays").filter(p => p.ts >= from && p.ts <= to).reduce((m, p) => {  // Pay later money received in the period
      const g = m.filter(x => x.via === p.via)[0] || (m.push({via: p.via, n: 0, amount: 0}), m[m.length - 1]); g.n++; g.amount = r2_(g.amount + p.amount); return m;
    }, []),
    by_owner: group(i => i.seller_id || 0, i => {
      const s = sellers.find(x => x.id === i.seller_id);
      return {owner: s ? sellerName_(s) : "Store", shelf: s ? s.shelf : null};
    }),
    by_item: group(i => i.product_id + "|" + i.name, i => ({id: i.product_id, name: i.name})),
    profit: profit_(items, rows_("Products")),
  };
}

// Gross profit on the store's own goods (renters' goods earn the store their rent instead): sales - cost.
// Each line uses its cost at the time of the sale, or the product's cost now for sales made before costs were entered.
function profit_(items, products) {
  const m = new Map(), missing = new Set();
  let sales = 0, cost = 0;
  items.filter(i => !i.seller_id).forEach(i => {
    const p = products.find(x => x.id === i.product_id);
    const c = i.cost !== null && i.cost !== undefined ? i.cost : p && p.cost !== null && p.cost !== undefined ? p.cost : null;
    if (c === null) { missing.add(i.name); return; }
    const g = m.get(i.product_id) || {id: i.product_id, name: i.name, qty: 0, sales: 0, cost: 0};
    g.qty += i.qty; g.sales += i.qty * i.price; g.cost += i.qty * c;
    m.set(i.product_id, g);
    sales += i.qty * i.price; cost += i.qty * c;
  });
  const list = Array.from(m.values()).map(g => Object.assign(g, {sales: r2_(g.sales), cost: r2_(g.cost), profit: r2_(g.sales - g.cost)}))
    .sort((a, b) => b.profit - a.profit);
  return {sales: r2_(sales), cost: r2_(cost), profit: r2_(sales - cost), items: list, missing: Array.from(missing).sort()};
}

// How the money came in: cash kept (what was paid minus the e-wallet part), then each e-wallet
function payments_(sales) {
  const m = {Cash: {via: "Cash", n: 0, amount: 0}};
  sales.forEach(s => {
    if (s.credit_to) { const g = m["Pay later"] = m["Pay later"] || {via: "Pay later", n: 0, amount: 0}; g.n++; g.amount += s.total; return; }  // no money yet
    const e = s.pay_via ? s.epay || 0 : 0, cash = Math.round((s.total - e) * 100) / 100;
    if (cash !== 0) { m.Cash.n++; m.Cash.amount += cash; }  // refunds count negative
    if (e !== 0) { const g = m[s.pay_via] = m[s.pay_via] || {via: s.pay_via, n: 0, amount: 0}; g.n++; g.amount += e; }
  });
  return Object.keys(m).map(k => m[k]).filter(g => g.n).map(g => Object.assign(g, {amount: r2_(g.amount)}));
}
const r2_ = n => Math.round(n * 100) / 100;

// ---------- cash drawer: starting cash, money in / out, end-of-day count ----------
// A shift runs from "open" (starting cash counted) to "close" (cash counted). Expected cash = starting cash + the cash
// part of every sale in between + money in - money out; e-wallet payments never reach the drawer.
// ponytail: one drawer for the whole store (all devices' sales count); per-device drawers if a second counter opens.
function drawerState_() {
  const moves = rows_("Drawer").map(strip_).filter(m => m.ts).sort((a, b) => a.ts.localeCompare(b.ts)), shifts = [];  // a row cleared by hand is skipped
  moves.forEach(m => {
    const cur = shifts.length && !shifts[shifts.length - 1].close ? shifts[shifts.length - 1] : null;
    if (m.kind === "open") shifts.push({open: m, moves: [], close: null});
    else if (cur && m.kind === "close") cur.close = m;
    else if (cur) cur.moves.push(m);
  });
  const sales = rows_("Sales"), last = shifts[shifts.length - 1];
  return {
    used: shifts.length > 0,
    current: last && !last.close ? shiftSummary_(last, sales) : null,
    past: shifts.filter(x => x.close).slice(-60).reverse().map(x => shiftSummary_(x, sales)),
  };
}

function shiftSummary_(sh, sales) {
  const to = sh.close ? sh.close.ts : "9999", mine = sales.filter(x => x.ts >= sh.open.ts && x.ts <= to);
  const pay = payments_(mine), cashSales = (pay.filter(p => p.via === "Cash")[0] || {amount: 0}).amount;
  const sum = kind => r2_(sh.moves.filter(m => m.kind === kind).reduce((n, m) => n + m.amount, 0));
  const ins = sum("in"), outs = sum("out"), expected = r2_(sh.open.amount + cashSales + ins - outs);
  const c = sh.close;
  return {
    opened: sh.open.ts, start: sh.open.amount, closed: c ? c.ts : null, moves: sh.moves,
    n: mine.length, total: r2_(mine.reduce((n, x) => n + x.total, 0)), by_pay: pay, cash_sales: cashSales, ins: ins, outs: outs,
    expected: c ? c.expected : expected,  // a closed day keeps the figure the cashier counted against
    counted: c ? c.counted : null, diff: c ? r2_(c.counted - c.expected) : null, note: c ? c.note : null, detail: c ? c.detail : null,
  };
}

function drawerMove_(d) {
  const kind = String(d.kind || ""), uid = String(d.uid || "") || null;
  if (["open", "in", "out", "close"].indexOf(kind) < 0) bad_("Unknown drawer action");
  if (uid && rows_("Drawer").some(m => m.uid === uid)) return {duplicate: true};  // a retry: already recorded
  const amount = r2_(Number(kind === "close" ? d.counted : d.amount));
  if (!(amount >= 0) || ((kind === "in" || kind === "out") && !(amount > 0))) bad_("Enter the amount");
  const note = String(d.note || "").trim().slice(0, 200) || null;
  if ((kind === "in" || kind === "out") && !note) bad_("Say what the money is for");
  const cur = drawerState_().current;
  if (kind === "open" && cur) bad_("The drawer is already open (since " + cur.opened.replace("T", " ").slice(0, 16) + ")");
  if (kind !== "open" && !cur) bad_("Open the drawer first: count the starting cash");
  const ts = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(d.ts || "") ? d.ts.slice(0, 19) : now_();
  const row = {uid: uid, ts: ts, kind: kind, amount: kind === "close" ? null : amount, note: note, expected: null, counted: null, detail: null};
  if (kind === "close") {
    // judged against what the cashier was shown (the device sends it), so the device and Google agree
    row.expected = d.expected !== undefined && d.expected !== null && d.expected !== "" && isFinite(Number(d.expected)) ? r2_(Number(d.expected)) : cur.expected;
    row.counted = amount;
    row.detail = d.detail ? String(d.detail).slice(0, 500) : null;
  }
  insert_("Drawer", row);
  return kind === "close" ? drawerState_().past[0] : {ok: true};
}

// ---------- Pay later: goods taken now, money received later (a reseller at the end of the day) ----------
function credits_() {  // sales with money still to receive, oldest first, each with what was brought back and paid
  const all = rows_("Sales"), pays = rows_("CreditPays"), refunds = all.filter(s => s.refund_of);
  const open = all.filter(s => s.credit_to && !s.refund_of), items = rows_("SaleItems"), back = {};
  const refundOf = new Map(refunds.map(r => [r.id, r.refund_of]));
  items.forEach(i => { const u = refundOf.get(i.sale_id); if (u) { back[u] = back[u] || {}; back[u][i.product_id] = (back[u][i.product_id] || 0) - i.qty; } });
  return open.map(s => {
    const returned = r2_(-refunds.filter(r => r.refund_of === s.uid).reduce((n, r) => n + r.total, 0));
    const paid = r2_(pays.filter(p => p.sale_uid === s.uid).reduce((n, p) => n + p.amount, 0));
    return {id: receiptNo_(s), uid: s.uid, ts: s.ts, name: s.credit_to, total: s.total, returned: returned, paid: paid, balance: r2_(s.total - returned - paid), credit_to: s.credit_to, cash: 0,
      items: items.filter(i => i.sale_id === s.id).map(i => ({id: i.product_id, name: i.name, qty: i.qty, price: i.price, refunded: (back[s.uid] && back[s.uid][i.product_id]) || 0}))};
  }).filter(c => c.balance > 0).sort((a, b) => a.ts.localeCompare(b.ts));
}

function creditPay_(d) {
  const uid = d.uid ? String(d.uid) : null;
  if (uid && rows_("CreditPays").some(p => p.uid === uid)) return {duplicate: true};  // a retry: already recorded
  const c = credits_().filter(x => x.uid === String(d.sale_uid))[0];
  if (!c) bad_("That sale is already fully paid");
  const amount = r2_(num_(d.amount, "Amount")), via = String(d.via || "Cash");
  if (!(amount > 0)) bad_("Enter the amount received");
  if (amount > c.balance) bad_("That is more than the " + smsMoney_(c.balance) + " still to pay");
  if (via !== "Cash" && PAY_VIA.indexOf(via) < 0) bad_("Unknown payment: " + via);
  const ts = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(d.ts || "") ? d.ts.slice(0, 19) : now_();
  insert_("CreditPays", {uid: uid, ts: ts, sale_uid: c.uid, name: c.name, amount: amount, via: via, note: String(d.note || "").trim().slice(0, 200) || null});
  if (via === "Cash" && drawerState_().current) insert_("Drawer", {uid: uid ? uid + "-d" : null, ts: ts, kind: "in", amount: amount, note: "Pay later: " + c.name});
  return {balance: r2_(c.balance - amount)};
}

function salesList_(q) {
  // Transaction history: newest first, with each sale's items (max 500 per request) and what was refunded of each
  const today = today_(), from = q.from || today, to = (q.to || today) + "T99", all = rows_("Sales");
  const sales = all.filter(s => s.ts >= from && s.ts <= to).sort((a, b) => b.ts.localeCompare(a.ts)).slice(0, 500);
  const ids = new Set(sales.map(s => s.id)), refundOf = new Map(all.filter(s => s.refund_of).map(s => [s.id, s.refund_of]));
  const items = [], back = {};  // back[uid][product id] = pieces refunded
  rows_("SaleItems").forEach(i => {
    if (ids.has(i.sale_id)) items.push(i);
    const u = refundOf.get(i.sale_id);
    if (u) { back[u] = back[u] || {}; back[u][i.product_id] = (back[u][i.product_id] || 0) - i.qty; }
  });
  return sales.map(s => Object.assign(strip_(s), {
    items: items.filter(i => i.sale_id === s.id).map(i => ({id: i.product_id, name: i.name, qty: i.qty, price: i.price,
      refunded: (s.uid && back[s.uid] && back[s.uid][i.product_id]) || 0})),
  }));
}

function sellerReport_(q) {
  // Shelf renters: per seller, what sold (every line), what they earned, rent, stock left, and the emails we sent them
  const today = today_(), from = q.from || today, to = (q.to || today) + "T99";
  const sales = new Map(rows_("Sales").filter(s => s.ts >= from && s.ts <= to).map(s => [s.id, s]));
  const items = rows_("SaleItems").filter(i => i.seller_id && sales.has(i.sale_id));
  const products = rows_("Products"), emails = rows_("Emails");
  return rows_("Sellers").sort(byName_).map(s => {  // deleted sellers only when they have something in this period
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
    // Rent is monthly, so it is settled at each payout (see payoutState_), never taken from a period of any length here
    const pct = s.commission || 0, commission = r2_(amount * pct / 100);
    const paid = rows_("Payouts").filter(p => p.seller_id === s.id && p.ts >= from && p.ts <= to).reverse().map(strip_);
    return {seller: Object.assign(strip_(s), {name: sellerName_(s)}), summary: {pcs: pcs, amount: amount, rent: rent, pct: pct, commission: commission, payout: r2_(amount - commission)},
            products: shelf, sales: lines, emails: mails, payouts: paid};
  }).filter(r => !r.seller.deleted || r.sales.length || r.payouts.length);
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
            owner: s ? sellerName_(s) : "Store", shelf: s ? s.shelf : null, expiry: p.expiry, mfg_date: p.mfg_date,
            sold: mine.reduce((n, i) => n + i.qty, 0), amount: mine.reduce((n, i) => n + i.qty * i.price, 0), sales: new Set(mine.map(i => i.sale_id)).size};
  });
  if (!one) {  // items sold in the period whose product was deleted since, so the totals match the sales report
    const gone = new Map();
    items.filter(i => !products.some(p => p.id === i.product_id)).forEach(i => {
      const k = i.product_id + "|" + i.name, s = sellers.find(x => x.id === i.seller_id);
      const g = gone.get(k) || {id: i.product_id, name: i.name, deleted: true, barcode: "", price: i.price, stock: null, value: 0,
        owner: s ? sellerName_(s) : "Store", shelf: s ? s.shelf : null, expiry: null, mfg_date: null, sold: 0, amount: 0, ids: new Set()};
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
    drawer: rows_("Drawer").slice(-500).map(strip_),
    adjustments: rows_("Adjustments").slice(-2000).map(strip_),
    payouts: rows_("Payouts").slice(-1000).map(strip_),
    credit_pays: rows_("CreditPays").slice(-2000).map(strip_),
    texts: rows_("Texts").slice(-300).map(strip_),
    more: newer.length > batch.length, quota: MailApp.getRemainingDailyQuota(),
  };
}

// ---------- seller emails ----------

// Text messages wait here for the tablet that sends them (Settings > This device). No links: the networks block them.
function queueText_(phone, body, sellerId, uid) {
  const d = String(phone || "").replace(/\D/g, ""), to = d.indexOf("63") === 0 && d.length === 12 ? "0" + d.slice(2) : d;
  insert_("Texts", {id: nextId_(rows_("Texts")), to_phone: to, body: smsChars_(body), created: now_(), seller_id: sellerId || null, uid: uid || null,
    error: /^09\d{9}$/.test(to) ? null : "Phone number looks wrong"});
}

// One character outside the phone alphabet (peso sign, curly quote, a-acute) shrinks a text from 160 to 70 letters and
// multiplies its cost, so everything is turned into plain letters first. n-tilde and e-acute are in the alphabet.
function smsChars_(t) {
  return String(t).replace(/\u20b1/g, "P").replace(/[\u2018\u2019]/g, "'").replace(/[\u201c\u201d]/g, '"').replace(/[\u2013\u2014\u2212]/g, "-").replace(/\u2026/g, "...")
    .replace(/[^\x0a\x20-\x7e\u00f1\u00d1\u00e9]/g, c => c.normalize("NFD").replace(/[^\x20-\x7e]/g, ""));
}

// The wording of each text: editable in Settings > Text messages. {words} are filled in for each seller.
const SMS_LINK = /(https?:|www\.|\.com\b|\.ph\b)/i;
const SMS_DEFAULTS = {
  sms_status: "{from}: Hi {name}. {items} {sales} {payout} {rent} {request} {call}",
  sms_rent: "{from}: Hi {name}, your rent of {amount} is due {date}. Will you renew for another month? Please reply YES or NO to this number.",
  sms_payout: "{from}: Hi {name}, your statement to {date}: sales {sales}, commission {commission}, rent {rent}. {result}",
};
function smsFill_(key, vals, template) {
  const t = template || settings_()[key] || SMS_DEFAULTS[key];
  return smsChars_(t.replace(/\{(\w+)\}/g, (m, k) => k in vals ? vals[k] : m)).replace(/[ \t]{2,}/g, " ").trim();
}
const smsFrom_ = s => (settings_().store_name || "Urdaneta Store") + (s.shelf ? ", Shelf " + s.shelf : "");
const smsName_ = s => s.name.trim().split(/\s+/)[0];

function smsPreview_(q) {  // sample data, so the page can show a text before any seller gets it
  const from = (settings_().store_name || "Urdaneta Store") + ", Shelf 12";
  const samples = {
    sms_status: {from: from, name: "Nena", items: "Polvoron expires Nov 12, 9 left. Low stock: Banana Chips 2, Peanut Brittle 3. 14 other items OK.", sales: "Nov sales: 48 pcs.",
      payout: "For payout: P2,592 after commission.", rent: "Rent P1,500 due Nov 15.", request: "Please pull out the Polvoron by Nov 12.", call: settings_().receipt_phone ? "Questions? Call " + settings_().receipt_phone + "." : ""},
    sms_rent: {from: from, name: "Nena", amount: "P1,500", date: "Nov 15, 2026"},
    sms_payout: {from: from, name: "Nena", date: "Nov 06, 2026", sales: "P2,880", commission: "P288", rent: "P1,500", result: "Paid to you (GCash): P1,092."},
  };
  return {text: smsFill_(q.key, samples[q.key] || {}, q.template || SMS_DEFAULTS[q.key])};
}

function smsMoney_(n) {  // P2,592 or P2,592.50
  const v = Math.abs(n), t = v % 1 ? v.toFixed(2) : String(v), p = t.split(".");
  return "P" + p[0].replace(/\B(?=(\d{3})+$)/g, ",") + (p[1] ? "." + p[1] : "");
}

// The status text for one seller: who it is from, what needs action, the money, one request. Named items are capped so
// it stays within 2 texts; everything else is a count (sender first, one request).
function statusText_(q) {
  const s = liveSellers_().find(x => x.id === Number(q.id));
  if (!s) bad_("Seller not found");
  const st = settings_(), today = today_(), short = n => n.length > 18 ? n.slice(0, 17).trim() + "." : n, day = d => dateText_(d).slice(0, 6);
  const items = rows_("Products").filter(p => p.seller_id === s.id), limit = s.low_stock === null || s.low_stock === undefined ? 5 : s.low_stock;
  const expired = items.filter(p => p.stock > 0 && p.expiry && p.expiry < today);
  const soon = items.filter(p => p.stock > 0 && p.expiry && p.expiry >= today && ((p.notify_on && p.notify_on <= today) || p.expiry <= addDays_(today, 7)));
  const dated = expired.concat(soon), low = items.filter(p => p.stock <= limit && dated.indexOf(p) < 0).sort((a, b) => a.stock - b.stock);
  const pay = payoutState_({id: s.id})[0], month = now_().slice(0, 8) + "01T00:00:00", sold = renterSales_(s.id, month, now_()).reduce((n, i) => n + i.qty, 0);
  const due = nextDue_(s.rent_start, today), rentSoon = s.rent && due && due <= addDays_(today, 7);
  const parts = named => {
    const out = [], first = smsName_(s), left = named, ds = dated.slice(0, left);
    ds.forEach(p => out.push(short(p.name) + (p.expiry < today ? " expired " : " expires ") + day(p.expiry) + ", " + p.stock + " left."));
    if (dated.length > ds.length) out.push("+" + (dated.length - ds.length) + " more expiring.");
    const ls = low.slice(0, Math.max(left - ds.length, 0));
    if (ls.length) out.push("Low stock: " + ls.map(p => short(p.name) + " " + Math.max(p.stock, 0)).join(", ") + (low.length > ls.length ? ", +" + (low.length - ls.length) + " more." : "."));
    else if (low.length) out.push(low.length + (low.length === 1 ? " item low." : " items low."));
    const fine = items.length - dated.length - low.length;
    return {first: first, list: out, fine: fine};
  };
  const money = [];
  if (pay.sales || pay.rent) money.push(pay.to_pay < 0 ? "Balance owed to store: " + smsMoney_(pay.to_pay) + "." :
    "For payout: " + smsMoney_(pay.to_pay) + (pay.commission && pay.rent ? " after commission and rent." : pay.commission ? " after commission." : pay.rent ? " after rent." : "."));
  if (rentSoon) money.push("Rent " + smsMoney_(s.rent) + " due " + day(due) + ".");
  const ask = expired.length + soon.length ? (expired.length ? "Please pull out expired items." : "Please pull out the " + short(soon[0].name) + " by " + day(soon[0].expiry) + ".")
    : low.length ? "Please restock soon." : "";
  const call = st.receipt_phone ? "Questions? Call " + st.receipt_phone + "." : "";
  const write = (named, withSales) => {
    const p = parts(named), okAll = !dated.length && !low.length;
    const line = okAll ? [items.length === 1 ? "Your 1 item is OK." : items.length ? "All " + items.length + " items OK." : "No items on your shelf."]
      : p.list.concat(p.fine ? [p.fine + (p.fine === 1 ? " other item OK." : " other items OK.")] : []);
    return smsFill_("sms_status", {from: smsFrom_(s), name: p.first, items: line.join(" "), sales: withSales && sold ? day(today).slice(0, 3) + " sales: " + sold + " pcs." : "",
      payout: money.filter(m => m.indexOf("Rent ") !== 0).join(" "), rent: money.filter(m => m.indexOf("Rent ") === 0).join(" "), request: ask, call: call});
  };
  let text = "";
  for (const [n, sales] of [[3, true], [2, true], [2, false], [1, false], [0, false]]) { text = write(n, sales); if (text.length <= 306) break; }
  return {to: s.phone || "", text: text};
}

function sendText_(d) {  // a text written by staff, checked on screen first
  const uid = d.uid ? String(d.uid) : null;
  if (uid && rows_("Texts").some(t => t.uid === uid)) return {duplicate: true};
  const s = rows_("Sellers").find(x => x.id === Number(d.id));
  if (!s) bad_("Seller not found");
  if (!s.phone) bad_("Add the seller's phone number first");
  const body = String(d.body || "").trim();
  if (!body) bad_("The message is empty");
  if (SMS_LINK.test(body)) bad_("Take out the web link: the phone networks block texts with links");
  queueText_(s.phone, body.slice(0, 480), s.id, uid);
  return {ok: true};
}

function textSent_(d) {  // the tablet's result; with `retry` the text goes back in the queue
  const t = rows_("Texts").find(x => x.id === Number(d.id));
  if (!t) return {ok: true};  // cleared from the Sheet by hand
  update_("Texts", Object.assign(t, d.retry ? {sent: null, error: null} : {sent: d.error ? null : now_(), error: d.error ? String(d.error).slice(0, 200) : null}));
  return {ok: true};
}

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
  liveSellers_().filter(s => s.email).forEach(s => {
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
  liveSellers_().filter(s => s.rent_start).forEach(s => {
    const due = nextDue_(s.rent_start, today), left = daysFrom_(today, due);
    if (left > notice || (s.renew_for === due && s.renew_answer) || !add(s.id, "rent", due)) return;
    if (s.email) rentEmail_(s, due, left);
    if (s.phone) queueText_(s.phone, smsFill_("sms_rent", {from: smsFrom_(s), name: smsName_(s), amount: smsMoney_(s.rent || 0), date: dateText_(due)}), s.id);
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
  const s = liveSellers_().find(r => r.id === Number(q.s));
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
// ---------- renter payouts: sales since the last payout, minus commission, minus rent due since then ----------
// ponytail: a sale synced late with a time before a payout falls outside it; only matters with several offline counters.
function renterSales_(id, from, to) {
  const ids = new Map(rows_("Sales").filter(s => s.ts > from && s.ts <= to).map(s => [s.id, s]));
  return rows_("SaleItems").filter(i => i.seller_id === id && ids.has(i.sale_id));
}

function rentDates_(s, from, to) {  // the rent due dates after `from`, up to and including `to`
  const out = [];
  let d = s.rent ? nextDue_(s.rent_start, addDays_(from.slice(0, 10), 1)) : null;
  while (d && d <= to.slice(0, 10) && out.length < 60) { out.push(d); d = nextDue_(s.rent_start, addDays_(d, 1)); }
  return out;
}

function payoutState_(q) {
  q = q || {};
  const now = now_(), payouts = rows_("Payouts");
  return liveSellers_().filter(s => !q.id || s.id === Number(q.id)).map(s => {
    const mine = payouts.filter(p => p.seller_id === s.id).sort((a, b) => a.to_ts.localeCompare(b.to_ts)), last = mine[mine.length - 1];
    // First payout: from the 1st of this month (older sales were settled by hand); the pop-up can change the date
    const since = q.from ? q.from + "T00:00:00" : last ? last.to_ts : now.slice(0, 8) + "01T00:00:00";
    const lines = renterSales_(s.id, since, now), sales = r2_(lines.reduce((n, i) => n + i.qty * i.price, 0));
    const pct = s.commission || 0, commission = r2_(sales * pct / 100), dates = rentDates_(s, since, now), rent = r2_(dates.length * (s.rent || 0));
    return {id: s.id, since: since, until: now, sales: sales, pcs: lines.reduce((n, i) => n + i.qty, 0), pct: pct, commission: commission,
            rent: rent, rent_dates: dates, to_pay: r2_(sales - commission - rent), last: last ? strip_(last) : null};
  });
}

function payout_(d) {
  const uid = d.uid ? String(d.uid) : null, all = rows_("Payouts");
  if (uid && all.some(p => p.uid === uid)) return {duplicate: true};  // a retry: already recorded
  const s = rows_("Sellers").find(x => x.id === Number(d.id));
  if (!s) bad_("Seller not found");
  const via = String(d.via || "Cash");
  if (via !== "Cash" && PAY_VIA.indexOf(via) < 0) bad_("Unknown payment: " + via);
  const ts = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(d.ts || "") ? d.ts.slice(0, 19) : now_();
  const from = String(d.from_ts || "");
  if (!/^\d{4}-\d{2}-\d{2}T/.test(from) || from > ts) bad_("Bad start date");
  // The amounts the owner saw and agreed with the renter, so the device and Google keep the same record
  const sales = r2_(num_(d.sales, "Sales")), commission = r2_(num_(d.commission, "Commission")), rent = r2_(num_(d.rent, "Rent"));
  if (commission < 0 || rent < 0) bad_("Commission and rent cannot be negative");
  const paid = r2_(sales - commission - rent), note = String(d.note || "").trim().slice(0, 200) || null;
  const until = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(d.to_ts || "") && d.to_ts <= ts ? d.to_ts : ts;  // sales after it go to the next payout
  const row = {uid: uid, ts: ts, seller_id: s.id, name: s.name, from_ts: from, to_ts: until, sales: sales, commission: commission, rent: rent, paid: paid, via: via, note: note};
  insert_("Payouts", row);
  // Cash goes through the drawer when it is open; a renter who owes the store pays in
  if (via === "Cash" && paid !== 0 && drawerState_().current) {
    insert_("Drawer", {uid: uid ? uid + "-d" : null, ts: ts, kind: paid > 0 ? "out" : "in", amount: Math.abs(paid), note: "Renter payout: " + s.name});
  }
  if (s.email) {
    const store = settings_().store_name || "Urdaneta Store", php = n => "PHP " + Number(n).toFixed(2);
    queueEmail_(s.email, store + ": your payout statement, " + dateText_(ts.slice(0, 10)),
      "Hi " + s.name + ",\n\nThis is " + store + ". Your statement from " + dateText_(from.slice(0, 10)) + " to " + dateText_(ts.slice(0, 10)) + ":\n\n" +
      "  Sales: " + php(sales) + "\n  Commission: " + php(commission) + "\n  Shelf rent: " + php(rent) + "\n  " +
      (paid >= 0 ? "Paid to you (" + via + "): " + php(paid) : "You paid the store (" + via + "): " + php(-paid)) + "\n" +
      (note ? "\nNote: " + note + "\n" : "") + "\nThank you!\n" + store, s.id);
  }
  if (s.phone) queueText_(s.phone, smsFill_("sms_payout", {from: smsFrom_(s), name: smsName_(s), date: dateText_(ts.slice(0, 10)), sales: smsMoney_(sales), commission: smsMoney_(commission), rent: smsMoney_(rent),
    result: (paid >= 0 ? "Paid to you" : "You paid the store") + " (" + via + "): " + smsMoney_(paid) + "."}), s.id);
  return strip_(row);
}

function renewSeller_(d) {
  const s = rows_("Sellers").find(r => r.id === Number(d.id));
  if (!s) bad_("Seller not found (was it deleted?)");
  if (!s.rent_start) bad_("Set the rental start date first");
  const answer = String(d.answer || "");
  if (answer && answer !== "yes" && answer !== "no") bad_("Answer must be yes or no");
  update_("Sellers", Object.assign(s, {renew_answer: answer || null, renew_for: answer ? nextDue_(s.rent_start, today_()) : null, renew_at: answer ? now_() : null}));
}

function stockReport_(d) {
  const sellers = liveSellers_().filter(s => s.email && (!d.id || s.id === Number(d.id)));
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

const PUBLIC_SETTINGS = ["store_name", "smtp_from", "receipt_address", "receipt_phone", "receipt_footer", "rent_notice_days", "web_app_url", "receipt_logo", "sms_status", "sms_rent", "sms_payout"];  // receipt_logo: "width,height,base64" 1-bit picture

function saveSettings_(d) {
  if (String(d.receipt_logo || "").length > 20000) bad_("The logo is too big");  // a Sheet cell holds 50,000 characters
  ["sms_status", "sms_rent", "sms_payout"].forEach(k => {
    if (String(d[k] || "").length > 480) bad_("The message is too long");
    if (SMS_LINK.test(d[k] || "")) bad_("Take out the web link: the phone networks block texts with links");
  });
  PUBLIC_SETTINGS.forEach(k => { if (k in d) setSetting_(k, String(d[k] || "").trim()); });
}

function state_() {
  const st = settings_();
  return {
    products: rows_("Products").sort(byName_).map(strip_),
    sellers: liveSellers_().sort(byName_).map(strip_),
    settings: PUBLIC_SETTINGS.reduce((o, k) => { o[k] = st[k] || ""; return o; }, {}),  // never the password
    outbox: rows_("Emails").slice(-20).reverse().map(m => { const c = strip_(m); delete c.body; return c; }),
    texts: rows_("Texts").slice(-20).reverse().map(strip_),
    quota: MailApp.getRemainingDailyQuota(),
  };
}
