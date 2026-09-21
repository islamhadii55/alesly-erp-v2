(function () {
  const DB_NAME = "alasly-offline";
  const DB_VERSION = 1;
  const cfg = window.ALASLY_SYNC || {};
  let dbp = null;

  function openDb() {
    if (dbp) return dbp;
    dbp = new Promise(function (resolve, reject) {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = function () {
        const db = req.result;
        if (!db.objectStoreNames.contains("queue")) {
          db.createObjectStore("queue", { keyPath: "op_uuid" });
        }
        if (!db.objectStoreNames.contains("catalog")) {
          db.createObjectStore("catalog", { keyPath: "key" });
        }
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
    return dbp;
  }

  function txStore(name, mode) {
    return openDb().then(function (db) {
      return db.transaction(name, mode).objectStore(name);
    });
  }

  function putQueue(op) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        const tx = db.transaction("queue", "readwrite");
        tx.objectStore("queue").put(op);
        tx.oncomplete = function () { resolve(op); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  function allQueue() {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        const req = db.transaction("queue").objectStore("queue").getAll();
        req.onsuccess = function () { resolve(req.result || []); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function delQueue(id) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        const tx = db.transaction("queue", "readwrite");
        tx.objectStore("queue").delete(id);
        tx.oncomplete = function () { resolve(); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  function saveCatalog(data) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        const tx = db.transaction("catalog", "readwrite");
        tx.objectStore("catalog").put({ key: "main", data: data, saved_at: Date.now() });
        tx.oncomplete = function () { resolve(); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  function loadCatalog() {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        const req = db.transaction("catalog").objectStore("catalog").get("main");
        req.onsuccess = function () { resolve(req.result ? req.result.data : null); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function uuid() {
    if (crypto && crypto.randomUUID) return crypto.randomUUID();
    return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, function (c) {
      const r = Math.random() * 16 | 0;
      const v = c === "x" ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
  }

  function setBadge(online, pending) {
    const el = document.getElementById("offline-badge");
    if (!el) return;
    if (!online) {
      el.className = "badge b-void";
      el.textContent = pending ? "بدون إنترنت · " + pending + " معلّق" : "بدون إنترنت";
    } else if (pending) {
      el.className = "badge b-credit";
      el.textContent = "متصل · مزامنة " + pending;
    } else {
      el.className = "badge b-paid";
      el.textContent = "متصل ومزامَن";
    }
  }

  async function refreshBadge() {
    const ops = await allQueue();
    const pending = ops.filter(function (o) { return o.status !== "synced"; }).length;
    setBadge(navigator.onLine, pending);
    return pending;
  }

  async function enqueue(opType, payload) {
    const op = {
      op_uuid: uuid(),
      op_type: opType,
      payload: payload,
      created_at: new Date().toISOString().slice(0, 19).replace("T", " "),
      status: "pending"
    };
    await putQueue(op);
    await refreshBadge();
    if (navigator.onLine) syncNow();
    return op;
  }

  async function syncNow() {
    if (!navigator.onLine) {
      await refreshBadge();
      return { ok: false, reason: "offline" };
    }
    const ops = (await allQueue()).filter(function (o) { return o.status !== "synced"; });
    try {
      const catRes = await fetch("/api/offline-catalog");
      if (catRes.ok) {
        const cat = await catRes.json();
        await saveCatalog(cat);
      }
    } catch (e) { /* keep last catalog */ }
    if (!ops.length) {
      await refreshBadge();
      return { ok: true, applied: 0 };
    }
    try {
      const res = await fetch("/api/sync/push", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          device_id: cfg.device_id,
          device_name: cfg.device_name,
          platform: window.ALASLY_PLATFORM || "web",
          ops: ops
        })
      });
      if (!res.ok) throw new Error("sync failed");
      const data = await res.json();
      const errIds = {};
      (data.errors || []).forEach(function (e) { errIds[e.op_uuid] = e.error; });
      for (let i = 0; i < ops.length; i++) {
        const op = ops[i];
        if (errIds[op.op_uuid]) {
          op.status = "error";
          op.error = errIds[op.op_uuid];
          await putQueue(op);
        } else {
          await delQueue(op.op_uuid);
        }
      }
      await refreshBadge();
      return data;
    } catch (err) {
      await refreshBadge();
      return { ok: false, error: String(err) };
    }
  }

  function interceptSearch() {
    const origFetch = window.fetch;
    window.fetch = function (input, init) {
      const url = typeof input === "string" ? input : (input && input.url) || "";
      if (url.indexOf("/api/search") === 0 || url.indexOf("/api/barcode") === 0 || url.indexOf("/api/next-sku") === 0 || url.indexOf("/api/product/") === 0) {
        return origFetch.apply(this, arguments).catch(function () {
          return loadCatalog().then(function (cat) {
            if (!cat) return new Response("[]", { headers: { "Content-Type": "application/json" } });
            if (url.indexOf("/api/next-sku") === 0) {
              return new Response(JSON.stringify({ sku: "OFF-" + Date.now().toString().slice(-6) }), { headers: { "Content-Type": "application/json" } });
            }
            if (url.indexOf("/api/barcode") === 0) {
              const code = (new URL(url, location.origin).searchParams.get("code") || "").toLowerCase();
              const p = (cat.products || []).find(function (x) {
                return String(x.barcode || "").toLowerCase() === code || String(x.sku || "").toLowerCase() === code;
              }) || {};
              return new Response(JSON.stringify(p), { headers: { "Content-Type": "application/json" } });
            }
            if (url.indexOf("/api/product/") === 0) {
              const id = parseInt(url.split("/").pop(), 10);
              const p = (cat.products || []).find(function (x) { return x.id === id; }) || {};
              return new Response(JSON.stringify(p), { headers: { "Content-Type": "application/json" } });
            }
            const u = new URL(url, location.origin);
            const q = (u.searchParams.get("q") || "").toLowerCase();
            const type = u.searchParams.get("type") || "product";
            let rows = [];
            if (type === "product") rows = cat.products || [];
            else if (type === "customer") rows = cat.customers || [];
            else if (type === "supplier") rows = cat.suppliers || [];
            else if (type === "employee") rows = cat.employees || [];
            const out = rows.filter(function (r) {
              return JSON.stringify(r).toLowerCase().indexOf(q) !== -1;
            }).slice(0, 12).map(function (r) {
              if (type === "product") {
                return Object.assign({}, r, {
                  label: (r.sku || "") + " — " + (r.name || ""),
                  hint: [r.brand, r.car_model, "كمية " + r.qty, r.location].filter(Boolean).join(" · ")
                });
              }
              return Object.assign({}, r, { label: r.name, hint: r.phone || "" });
            });
            return new Response(JSON.stringify(out), { headers: { "Content-Type": "application/json" } });
          });
        });
      }
      return origFetch.apply(this, arguments);
    };
  }

  function interceptForms() {
    document.addEventListener("submit", function (e) {
      if (navigator.onLine) return;
      const form = e.target;
      if (!form || form.method.toLowerCase() !== "post") return;
      if (form.hasAttribute("data-online-only")) return;
      const action = (form.getAttribute("action") || location.pathname).toLowerCase();
      const fd = new FormData(form);
      const obj = {};
      fd.forEach(function (v, k) {
        if (obj[k] === undefined) obj[k] = v;
        else if (Array.isArray(obj[k])) obj[k].push(v);
        else obj[k] = [obj[k], v];
      });
      function packItems() {
        return {
          "item_product_id[]": fd.getAll("item_product_id[]"),
          "item_desc[]": fd.getAll("item_desc[]"),
          "item_qty[]": fd.getAll("item_qty[]"),
          "item_price[]": fd.getAll("item_price[]"),
          "item_cost[]": fd.getAll("item_cost[]"),
          party_id: fd.get("party_id"),
          party_name: fd.get("party_name") || fd.get("party_search") || "عميل نقدي",
          payment_method: fd.get("payment_method") || "نقدي",
          discount: fd.get("discount") || 0,
          tax: fd.get("tax") || 0,
          coupon_code: fd.get("coupon_code") || "",
          notes: (fd.get("notes") || "") + " [حفظ بدون اتصال]"
        };
      }
      let opType = null;
      let payload = obj;
      if (action.indexOf("/pos") !== -1 || action.indexOf("/invoices/") !== -1) {
        opType = "sale";
        payload = packItems();
        if (action.indexOf("/purchase") !== -1) payload.kind = "purchase";
        else if (action.indexOf("/maintenance") !== -1) payload.kind = "maintenance";
        else payload.kind = "sale";
      } else if (action.indexOf("/expenses") !== -1) {
        opType = "expense";
      } else if (action.indexOf("/customers") !== -1) {
        opType = "customer";
      } else if (action.indexOf("/inventory") !== -1 || action.indexOf("/products") !== -1) {
        opType = "product";
      } else if (action.indexOf("/jobs/new") !== -1) {
        opType = "job";
      }
      if (!opType) return;
      e.preventDefault();
      enqueue(opType, payload).then(function () {
        alert("لا يوجد اتصال. حُفظت العملية محلياً وستُزامَن تلقائياً عند عودة الإنترنت.");
        location.reload();
      });
    }, true);
  }

  function registerSw() {
    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.register("/sw.js").catch(function () {});
    }
  }

  window.AlaslyOffline = { enqueue: enqueue, syncNow: syncNow, loadCatalog: loadCatalog, refreshBadge: refreshBadge };

  document.addEventListener("DOMContentLoaded", function () {
    interceptSearch();
    interceptForms();
    registerSw();
    refreshBadge();
    if (navigator.onLine) syncNow();
    window.addEventListener("online", function () { syncNow(); });
    window.addEventListener("offline", function () { refreshBadge(); });
    setInterval(function () { if (navigator.onLine) syncNow(); }, 20000);
    const btn = document.getElementById("sync-now-btn");
    if (btn) btn.addEventListener("click", function (e) {
      e.preventDefault();
      syncNow().then(function (r) {
        alert(r && r.ok ? "تمت المزامنة" : "تعذر الاتصال — ستُعاد المحاولة تلقائياً");
      });
    });
  });
})();
