(function () {
  function el(html) {
    const t = document.createElement("template");
    t.innerHTML = html.trim();
    return t.content.firstElementChild;
  }

  function debounce(fn, ms) {
    let t;
    return function () {
      const ctx = this, args = arguments;
      clearTimeout(t);
      t = setTimeout(function () { fn.apply(ctx, args); }, ms);
    };
  }

  function attachSuggest(input, options) {
    const wrap = document.createElement("div");
    wrap.className = "suggest-wrap";
    input.parentNode.insertBefore(wrap, input);
    wrap.appendChild(input);
    const box = document.createElement("div");
    box.className = "suggest-box";
    wrap.appendChild(box);
    let items = [];
    let active = -1;

    function hide() { box.classList.remove("open"); box.innerHTML = ""; active = -1; }

    function render() {
      if (!items.length) { hide(); return; }
      box.innerHTML = "";
      items.forEach(function (item, i) {
        const row = document.createElement("div");
        row.className = "suggest-item" + (i === active ? " active" : "");
        row.innerHTML = "<strong>" + (item.label || item.name || "") + "</strong>" +
          (item.hint ? "<small>" + item.hint + "</small>" : "");
        row.addEventListener("mousedown", function (e) {
          e.preventDefault();
          choose(item);
        });
        box.appendChild(row);
      });
      box.classList.add("open");
    }

    function choose(item) {
      if (options.onSelect) options.onSelect(item, input);
      else input.value = item.label || item.name || "";
      hide();
    }

    const run = debounce(function () {
      const q = input.value.trim();
      if (!q) { hide(); return; }
      const type = options.type || input.getAttribute("data-suggest") || "product";
      fetch("/api/search?type=" + encodeURIComponent(type) + "&q=" + encodeURIComponent(q))
        .then(function (r) { return r.json(); })
        .then(function (data) { items = data || []; active = items.length ? 0 : -1; render(); })
        .catch(function () { hide(); });
    }, 80);

    input.addEventListener("input", run);
    input.addEventListener("focus", function () { if (input.value.trim()) run(); });
    input.addEventListener("blur", function () { setTimeout(hide, 180); });
    input.addEventListener("keydown", function (e) {
      if (!box.classList.contains("open")) return;
      if (e.key === "ArrowDown") { e.preventDefault(); active = Math.min(active + 1, items.length - 1); render(); }
      else if (e.key === "ArrowUp") { e.preventDefault(); active = Math.max(active - 1, 0); render(); }
      else if (e.key === "Enter" && active >= 0) { e.preventDefault(); choose(items[active]); }
      else if (e.key === "Escape") hide();
    });
  }

  function fillProductFields(tr, p) {
    const idInput = tr.querySelector('[name="item_product_id[]"]');
    const desc = tr.querySelector('[name="item_desc[]"]');
    const cost = tr.querySelector('[name="item_cost[]"]');
    const price = tr.querySelector('[name="item_price[]"]');
    if (idInput) idInput.value = p.id;
    if (desc) desc.value = p.name;
    if (cost) cost.value = p.cost;
    if (price) price.value = p.price;
    updateLineTotal(tr);
  }

  function updateLineTotal(row) {
    if (!row) return;
    const qty = parseFloat((row.querySelector('[name="item_qty[]"]') || {}).value) || 0;
    const price = parseFloat((row.querySelector('[name="item_price[]"]') || {}).value) || 0;
    const total = row.querySelector('.line-total');
    if (total) total.textContent = (qty * price).toFixed(2);
  }

  function bindLineTotal(row) {
    if (!row || row.dataset.totalBound) return;
    row.dataset.totalBound = "1";
    row.querySelectorAll('[name="item_qty[]"], [name="item_price[]"]').forEach(function (input) {
      input.addEventListener("input", function () { updateLineTotal(row); });
      input.addEventListener("change", function () { updateLineTotal(row); });
    });
    updateLineTotal(row);
  }

  function setField(name, value) {
    if (value === undefined || value === null || value === "") return;
    const field = document.querySelector("form.form-grid [name='" + name + "']");
    if (field && !field.readOnly) field.value = value;
  }

  function fillProductForm(item) {
    const form = document.querySelector("form.form-grid");
    if (!form || !form.querySelector("[name='name']")) return;
    const editing = !!form.querySelector("[name='id']");
    const nameInput = form.querySelector("[name='name']");
    if (nameInput) nameInput.value = item.name || "";
    setField("category", item.category);
    setField("brand", item.brand);
    setField("car_model", item.car_model);
    setField("year_from", item.year_from);
    setField("year_to", item.year_to);
    setField("unit", item.unit);
    setField("item_type", item.item_type);
    setField("cost", item.cost);
    setField("price", item.price);
    setField("min_qty", item.min_qty);
    setField("warehouse", item.warehouse);
    setField("aisle", item.aisle);
    setField("shelf", item.shelf);
    setField("bin", item.bin);
    if (editing) {
      setField("sku", item.sku);
      setField("barcode", item.barcode || item.sku);
      setField("qty", item.qty);
      setField("notes", item.notes);
    }
    const loc = item.location || [item.warehouse, item.aisle, item.shelf, item.bin].filter(Boolean).join(" / ");
    const hint = document.getElementById("sku-hint");
    if (hint) {
      hint.textContent = editing
        ? "تم تحميل بيانات الصنف من المخزن"
        : "بيانات مرتبطة من المخزن" + (item.sku ? " (" + item.sku + ")" : "") +
          (loc ? " — " + loc : "") + " — كمية " + (item.qty || 0) + " — الكود الجديد يُولَّد تلقائياً";
    }
  }

  window.addProductRow = function () {
    const tpl = document.getElementById("row-tpl");
    if (!tpl) return;
    const node = tpl.content.cloneNode(true);
    document.querySelector("#items tbody").appendChild(node);
    const row = document.querySelector("#items tbody tr:last-child");
    bindLineTotal(row);
    const search = row.querySelector(".item-search");
    if (search) {
      attachSuggest(search, {
        type: "product",
        onSelect: function (item) {
          search.value = item.sku + " — " + item.name;
          fillProductFields(row, item);
        }
      });
    }
  };

  function bindExistingRows() {
    document.querySelectorAll("#items tbody tr").forEach(function (row) {
      bindLineTotal(row);
      const search = row.querySelector(".item-search");
      if (search && !search.dataset.bound) {
        search.dataset.bound = "1";
        attachSuggest(search, {
          type: "product",
          onSelect: function (item) {
            search.value = item.sku + " — " + item.name;
            fillProductFields(row, item);
          }
        });
      }
    });
  }

  document.addEventListener("DOMContentLoaded", function () {
    document.querySelectorAll("[data-suggest]").forEach(function (input) {
      const type = input.getAttribute("data-suggest");
      const target = input.getAttribute("data-target");
      attachSuggest(input, {
        type: type,
        onSelect: function (item) {
          if (type === "product" && input.name === "name" && !input.hasAttribute("data-goto-product")) {
            fillProductForm(item);
            return;
          }
          input.value = item.label || item.name || "";
          if (target) {
            const hidden = document.querySelector(target);
            if (hidden) hidden.value = item.id;
          }
          if (type === "customer" || type === "supplier") {
            const nameField = document.getElementById("party-name");
            if (nameField) nameField.value = item.name || item.label || "";
          }
          if (type === "invoice" && item.url) window.location = item.url;
          if (input.hasAttribute("data-goto-product") && item.id) {
            window.location = "/inventory?q=" + encodeURIComponent(item.sku || item.name);
          }
        }
      });
    });
    const liveFilter = document.querySelector(".global-search input");
    if (liveFilter) {
      liveFilter.addEventListener("input", function () {
        const q = liveFilter.value.trim();
        document.querySelectorAll("table.filter-table tr").forEach(function (tr) {
          if (tr.querySelector("th")) return;
          tr.style.display = !q || tr.textContent.indexOf(q) !== -1 ? "" : "none";
        });
      });
    }
    bindExistingRows();
    if (document.getElementById("items") && !document.querySelector("#items tbody tr")) {
      window.addProductRow();
    }

    const skuInput = document.querySelector("input[name='sku']");
    const catInput = document.querySelector("input[name='category']");
    const skuHint = document.getElementById("sku-hint");
    function refreshSku() {
      if (!skuInput || skuInput.readOnly === false && skuInput.value && skuInput.dataset.locked === "1") return;
      if (!skuInput || (skuInput.dataset.edit === "1")) return;
      const cat = catInput ? catInput.value : "";
      fetch("/api/next-sku?category=" + encodeURIComponent(cat))
        .then(function (r) { return r.json(); })
        .then(function (d) {
          if (!skuInput.value || skuInput.dataset.auto === "1") {
            skuInput.value = d.sku;
            skuInput.dataset.auto = "1";
            if (skuHint) skuHint.textContent = "سيتم الحفظ بالكود: " + d.sku;
          }
        });
    }
    if (skuInput && skuInput.dataset.edit !== "1") {
      skuInput.placeholder = "يُولَّد تلقائياً";
      skuInput.removeAttribute("required");
      skuInput.addEventListener("input", function () { skuInput.dataset.auto = skuInput.value ? "0" : "1"; });
      if (catInput) catInput.addEventListener("input", refreshSku);
      refreshSku();
    }
  });
})();
