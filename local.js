/* The POS's own database on this device (local-first).
 * engine.js (the same file as the Google script, google/Code.gs) runs here too, against tables kept in this device's
 * storage (IndexedDB), so the screens never wait for the internet. Every change is also queued and sent to Google in
 * the background; then Google's copy, with the other devices' sales, is pulled back.
 * Works in the browser, the Mac/Windows app and the Android app. Also runs under Node for google/test_sync.js. */
(function (G) {
  if (typeof SHEETS === "undefined") return;  // engine.js not loaded: the POS talks to pos.py instead

  // Requests that only Google can do (they send email)
  const CLOUD = new Set(["/api/send-now", "/api/outbox/test", "/api/sellers/report", "/api/outbox/delete", "/api/outbox/readdress"]);
  const WRITE = new Set(["/api/products", "/api/products/delete", "/api/photo", "/api/sellers", "/api/sellers/delete", "/api/sale", "/api/settings"]);

  let data = {}, ops = [], meta = {}, failed = [];

  // ---------- stand-ins for the Google services the device-side routes use ----------
  const text = v => v === null || v === undefined ? "" : typeof v === "string" && v[0] === "'" ? v.slice(1) : v;
  function sheet(n) {
    const d = data[n];
    return {
      getRange: (a) => typeof a === "string" ? {setNumberFormat() {}} : {setValues: rows => rows.forEach((r, i) => { d[a - 1 + i] = r.map(text); })},
      appendRow: r => { d.push(r.map(text)); },
      getDataRange: () => ({getValues: () => d.map(r => r.slice())}),
      deleteRow: r => { d.splice(r - 1, 1); },
      getLastRow: () => d.length,
      getLastColumn: () => (d[0] || []).length,
      setFrozenRows() {},
    };
  }
  G.SpreadsheetApp = {getActiveSpreadsheet: () => ({
    getSheetByName: n => data[n] ? sheet(n) : null,
    insertSheet: n => { data[n] = []; return sheet(n); },
  })};
  G.Utilities = {
    formatDate(d, tz, f) {
      if (f === "MMM dd, yyyy") return new Intl.DateTimeFormat("en-US", {timeZone: tz, month: "short", day: "2-digit", year: "numeric"}).format(d);
      const p = Object.fromEntries(new Intl.DateTimeFormat("en-US", {timeZone: tz, year: "numeric", month: "2-digit", day: "2-digit",
        hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23"}).formatToParts(d).map(x => [x.type, x.value]));
      return f === "yyyy-MM-dd" ? `${p.year}-${p.month}-${p.day}` : `${p.year}-${p.month}-${p.day}T${p.hour}:${p.minute}:${p.second}`;
    },
  };
  G.MailApp = {getRemainingDailyQuota: () => meta.quota ?? null};  // emails themselves are only sent by Google (CLOUD)

  // ---------- saving on the device ----------
  let dbp = null;
  const openDb = () => dbp || (dbp = new Promise((ok, no) => {
    const r = G.indexedDB.open("pos-local", 1);
    r.onupgradeneeded = () => r.result.createObjectStore("kv");
    r.onsuccess = () => ok(r.result);
    r.onerror = () => no(r.error);
  }));
  async function persist() {
    const db = await openDb();
    await new Promise((ok, no) => {
      const tx = db.transaction("kv", "readwrite"), st = tx.objectStore("kv");
      st.put({data, ops, meta, failed}, "state");
      tx.oncomplete = ok; tx.onerror = () => no(tx.error);
    });
  }
  async function open() {
    const db = await openDb();
    const saved = await new Promise((ok, no) => {
      const r = db.transaction("kv").objectStore("kv").get("state");
      r.onsuccess = () => ok(r.result); r.onerror = () => no(r.error);
    });
    if (saved) ({data, ops, meta, failed} = saved);
    failed = failed || [];
  }

  // ---------- requests from the POS screens ----------
  async function call(path, body) {
    const fn = ROUTES[path];
    if (!fn) throw new Error("This works only with internet");
    const copy = JSON.parse(JSON.stringify(body || {}));
    const out = fn(JSON.parse(JSON.stringify(copy)));  // same rules as Google; throws on bad input
    if (WRITE.has(path)) {
      // Send the same change to Google later, with the ids this device chose
      if (out && out.id && (path === "/api/products" || path === "/api/sellers")) copy.id = out.id;
      if (out && out.barcode) copy.barcode = out.barcode;
      ops.push({path, body: copy, at: Date.now()});
    }
    await persist();
    return out === undefined ? {ok: true} : out;
  }

  // ---------- sync with Google ----------
  const toRow = (name, o) => SHEETS[name].map(k => o[k] === null || o[k] === undefined ? "" : o[k]);
  const replace = (name, list) => { data[name] = [SHEETS[name].slice(), ...list.map(o => toRow(name, o))]; };

  function apply(r) {
    replace("Products", r.products);
    replace("Sellers", r.sellers);
    replace("Emails", r.emails);
    replace("Settings", Object.keys(r.settings).map(k => ({key: k, value: r.settings[k]})));
    if (!data.Sales) replace("Sales", []);
    if (!data.SaleItems) replace("SaleItems", []);
    meta.quota = r.quota;
    // Sales: add the ones this device doesn't have yet (other devices', or older history)
    const have = new Set(rows_("Sales").map(s => s.uid).filter(Boolean));
    let next = nextId_(rows_("Sales"));
    for (const s of r.sales) {
      meta.cloudSaleId = Math.max(meta.cloudSaleId || 0, s.id);
      if (s.uid && have.has(s.uid)) continue;
      const id = next++;
      data.Sales.push(toRow("Sales", {...s, id}));
      for (const i of s.items) data.SaleItems.push(toRow("SaleItems", {...i, sale_id: id}));
    }
  }

  // remote(path, body) calls Google; it throws an Error with .offline when there is no connection, .status when Google refused
  async function sync(remote) {
    while (ops.length) {
      try {
        await remote(ops[0].path, ops[0].body);
      } catch (e) {
        if (e.offline || !(e.status >= 400 && e.status < 500)) throw e;  // no internet or Google busy: try again later
        // Google refused this change (e.g. a barcode another device already used): keep it aside and carry on
        failed.push({...ops[0], error: e.message});
      }
      ops.shift();
      await persist();
    }
    let more = true;
    while (more) {
      const r = await remote("/api/pull", {after: meta.cloudSaleId || 0});
      if (ops.length) return;  // a sale was made while we were downloading: send it first, next round
      apply(r);
      more = r.more;
      await persist();
    }
    // Photos travel on their own (they are big): download only the ones this device doesn't have in their current version
    const have = new Map(rows_("Photos").map(r => [r.product_id, r.photo_at]));
    const want = rows_("Products").filter(p => p.photo_at && have.get(p.id) !== p.photo_at).map(p => p.id);
    for (let i = 0; i < want.length; i += 25) {
      const got = await remote("/api/photos", {ids: want.slice(i, i + 25)});
      if (ops.length) return;  // a change was made meanwhile (maybe a new photo): it goes first, next round
      const keep = rows_("Photos").filter(r => !got.some(g => g.product_id === r.product_id));
      replace("Photos", keep.concat(got));
      await persist();
    }
    // Forget photos of deleted products and removed photos
    const current = new Set(rows_("Products").filter(p => p.photo_at).map(p => p.id));
    replace("Photos", rows_("Photos").filter(r => current.has(r.product_id)));
    meta.lastSync = Date.now();
    await persist();
  }

  G.Local = {
    open, call, sync,
    cloudOnly: path => CLOUD.has(path),
    pending: () => ops.length,
    pendingSaleUids: () => new Set(ops.filter(o => o.path === "/api/sale").map(o => o.body.uid)),
    failed: () => failed,
    clearFailed: async () => { failed = []; await persist(); },
    lastSync: () => meta.lastSync || null,
    isEmpty: () => !meta.lastSync && !(data.Products && data.Products.length > 1),
    addOp: async (path, body) => { ops.push({path, body, at: Date.now()}); await persist(); },
  };
})(typeof window !== "undefined" ? window : globalThis);
