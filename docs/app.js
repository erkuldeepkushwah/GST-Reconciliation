/* GST Reconciliation — browser UI */
(function () {
  "use strict";

  var C = window.GSTCore;
  var $ = function (id) { return document.getElementById(id); };
  var esc = function (s) {
    return String(s === null || s === undefined ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  };

  /* ------------------------------------------------------------------ *
   * IndexedDB storage
   * ------------------------------------------------------------------ */
  var DB_NAME = "gst_reco_db", DB_VER = 3, _db = null;

  function openDB() {
    if (_db) return Promise.resolve(_db);
    return new Promise(function (resolve, reject) {
      var req = indexedDB.open(DB_NAME, DB_VER);
      req.onupgradeneeded = function (e) {
        var db = e.target.result;
        if (!db.objectStoreNames.contains("rows")) {
          var s = db.createObjectStore("rows", { keyPath: "id", autoIncrement: true });
          s.createIndex("gstin_type", "_gstin_type");
        }
        if (!db.objectStoreNames.contains("imports")) {
          var i = db.createObjectStore("imports", { keyPath: "id", autoIncrement: true });
          i.createIndex("gstin", "gstin");
        }
      };
      req.onsuccess = function () { _db = req.result; resolve(_db); };
      req.onerror = function () { reject(req.error); };
    });
  }

  function tx(store, mode) {
    return openDB().then(function (db) { return db.transaction(store, mode).objectStore(store); });
  }

  function addAll(store, items) {
    return tx(store, "readwrite").then(function (os) {
      return new Promise(function (resolve, reject) {
        items.forEach(function (it) { os.add(it); });
        os.transaction.oncomplete = function () { resolve(items.length); };
        os.transaction.onerror = function () { reject(os.transaction.error); };
      });
    });
  }

  function rowsFor(gstin, returnType) {
    return tx("rows", "readonly").then(function (os) {
      return new Promise(function (resolve, reject) {
        var req = os.index("gstin_type").getAll(gstin + "|" + returnType);
        req.onsuccess = function () { resolve(req.result || []); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function importsFor(gstin) {
    return tx("imports", "readonly").then(function (os) {
      return new Promise(function (resolve, reject) {
        var req = os.index("gstin").getAll(gstin);
        req.onsuccess = function () {
          var list = (req.result || []).sort(function (a, b) { return b.id - a.id; });
          resolve(list);
        };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function clearEverything() {
    _db = null;
    return new Promise(function (resolve) {
      var req = indexedDB.deleteDatabase(DB_NAME);
      req.onsuccess = req.onerror = req.onblocked = function () { resolve(); };
    });
  }

  /* ------------------------------------------------------------------ *
   * File reading
   * ------------------------------------------------------------------ */
  function decodeText(buffer) {
    try {
      return new TextDecoder("utf-8", { fatal: true }).decode(buffer).replace(/^\uFEFF/, "");
    } catch (e) {
      try { return new TextDecoder("windows-1252").decode(buffer); }
      catch (e2) { return new TextDecoder("utf-8").decode(buffer).replace(/^\uFEFF/, ""); }
    }
  }

  function parseCSV(text) {
    var firstLine = text.split(/\r?\n/)[0] || "";
    var counts = [[",", 0], [";", 0], ["\t", 0], ["|", 0]];
    var inQuotes = false;
    for (var i = 0; i < firstLine.length; i++) {
      var ch = firstLine[i];
      if (ch === '"') inQuotes = !inQuotes;
      if (!inQuotes) counts.forEach(function (c) { if (ch === c[0]) c[1]++; });
    }
    counts.sort(function (a, b) { return b[1] - a[1]; });
    var delim = counts[0][1] > 0 ? counts[0][0] : ",";

    var rows = [], field = "", row = [], q = false;
    for (var p = 0; p < text.length; p++) {
      var c = text[p];
      if (q) {
        if (c === '"') {
          if (text[p + 1] === '"') { field += '"'; p++; }
          else q = false;
        } else field += c;
      } else if (c === '"') q = true;
      else if (c === delim) { row.push(field); field = ""; }
      else if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; }
      else if (c === "\r") { /* skip */ }
      else field += c;
    }
    if (field.length || row.length) { row.push(field); rows.push(row); }
    return rows.filter(function (r) { return r.some(function (v) { return String(v).trim() !== ""; }); });
  }

  function dedupeHeaders(headers) {
    var seen = {}, out = [];
    headers.forEach(function (h, idx) {
      var name = String(h === null || h === undefined ? "" : h).trim();
      if (!name) name = "Column " + (idx + 1);
      if (seen[name]) { seen[name]++; name = name + " (" + seen[name] + ")"; }
      else seen[name] = 1;
      out.push(name);
    });
    return out;
  }

  function gridToTable(grid) {
    if (!grid.length) throw new Error("The file appears to be empty.");
    var headers = dedupeHeaders(grid[0]);
    var rows = [];
    for (var i = 1; i < grid.length; i++) {
      var cells = grid[i];
      if (!cells || !cells.some(function (v) { return String(v === null || v === undefined ? "" : v).trim() !== ""; })) continue;
      var obj = {};
      headers.forEach(function (h, idx) { obj[h] = cells[idx] === undefined ? "" : cells[idx]; });
      rows.push(obj);
    }
    return { headers: headers, rows: rows };
  }

  function readFile(file) {
    var name = (file.name || "").toLowerCase();
    if (name.endsWith(".xls") && !name.endsWith(".xlsx") && !name.endsWith(".xlsm")) {
      return Promise.reject(new Error("Old .xls files are not supported. In Excel, use Save As → Excel Workbook (*.xlsx) or CSV."));
    }
    return file.arrayBuffer().then(function (buf) {
      if (name.endsWith(".csv")) {
        var text = decodeText(buf);
        return gridToTable(parseCSV(text));
      }
      if (name.endsWith(".xlsx") || name.endsWith(".xlsm")) {
        var wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true, cellNF: false, cellText: false });
        var sheetName = wb.SheetNames[0];
        if (!sheetName) throw new Error("No worksheet found in this workbook.");
        var sheet = wb.Sheets[sheetName];
        var grid = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "", blankrows: false });
        var table = gridToTable(grid);
        table.sheetName = sheetName;
        return table;
      }
      throw new Error("Unsupported file type. Upload .xlsx, .xlsm or .csv.");
    });
  }

  /* ------------------------------------------------------------------ *
   * Period selection
   * ------------------------------------------------------------------ */
  var MONTH_NAMES = ["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"];
  var MONTH_NUM = { April: 4, May: 5, June: 6, July: 7, August: 8, September: 9, October: 10, November: 11, December: 12, January: 1, February: 2, March: 3 };

  function renderPeriodExtra() {
    var mode = $("periodMode").value, box = $("periodExtra"), fy = $("fy").value;
    if (mode === "Month-wise") {
      box.innerHTML = '<label class="field"><span>Month</span><select id="monthSel">' +
        MONTH_NAMES.map(function (m) { return "<option>" + m + "</option>"; }).join("") + "</select></label>";
    } else if (mode === "Quarter-wise") {
      box.innerHTML = '<label class="field"><span>Quarter</span><select id="quarterSel">' +
        ["Q1", "Q2", "Q3", "Q4"].map(function (q) { return "<option>" + q + "</option>"; }).join("") + "</select></label>";
    } else if (mode === "Custom Date Range") {
      var d = C.fyDates(fy);
      box.innerHTML =
        '<label class="field"><span>Start date</span><input id="startDate" type="date" value="' + d.start + '"></label>' +
        '<label class="field"><span>End date</span><input id="endDate" type="date" value="' + d.end + '"></label>';
    } else {
      box.innerHTML = "";
    }
    updatePeriodCaption();
  }

  function currentPeriod() {
    var mode = $("periodMode").value, fy = $("fy").value, start, end;
    if (mode === "Month-wise") {
      var m = MONTH_NUM[($("monthSel") || {}).value || "April"];
      var r = C.monthDates(fy, m); start = r.start; end = r.end;
    } else if (mode === "Quarter-wise") {
      var q = ($("quarterSel") || {}).value || "Q1";
      var r2 = C.quarterDates(fy, q); start = r2.start; end = r2.end;
    } else if (mode === "Custom Date Range") {
      start = ($("startDate") || {}).value || C.fyDates(fy).start;
      end = ($("endDate") || {}).value || C.fyDates(fy).end;
      if (end < start) end = start;
    } else {
      var r3 = C.fyDates(fy); start = r3.start; end = r3.end;
    }
    var importPeriod = start.slice(5, 7) + start.slice(0, 4);
    return { start: start, end: end, importPeriod: importPeriod };
  }

  function fmtD(iso) {
    if (!iso) return "—";
    var p = iso.split("-");
    return p[2] + "-" + p[1] + "-" + p[0];
  }

  var MONTH_LABELS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

  function periodLabel() {
    var mode = $("periodMode").value, p = currentPeriod(), fy = $("fy").value;
    if (mode === "Month-wise") {
      return MONTH_LABELS[parseInt(p.start.slice(5, 7), 10) - 1] + " " + p.start.slice(0, 4);
    }
    if (mode === "Quarter-wise") return "FY " + fy + " · " + (($("quarterSel") || {}).value || "Q1");
    if (mode === "Financial Year") return "FY " + fy;
    return fmtD(p.start).slice(0, 5) + " – " + fmtD(p.end).slice(0, 5);
  }

  function updateSetupChip() {
    var gstin = ($("gstin").value || "").trim().toUpperCase();
    var g = $("setupGstin"), per = $("setupPeriod");
    if (g) g.textContent = gstin.length ? gstin : "Add GSTIN";
    if (per) per.textContent = periodLabel();
  }

  function updatePeriodCaption() {
    var p = currentPeriod();
    $("periodCaption").textContent = "Selected dates: " + fmtD(p.start) + " to " + fmtD(p.end) + "  ·  import period " + p.importPeriod;
    var hint = $("importHint");
    if (hint) hint.textContent = "Rows will be tagged with period " + p.importPeriod + ".";
    updateSetupChip();
  }

  /* ---- mobile setup sheet ---- */
  function openSheet() {
    var bar = $("sidebar"), scrim = $("scrim");
    bar.classList.add("open");
    scrim.hidden = false;
    requestAnimationFrame(function () { scrim.classList.add("show"); });
    document.body.classList.add("no-scroll");
    $("setupBtn").setAttribute("aria-expanded", "true");
  }

  function closeSheet() {
    var bar = $("sidebar"), scrim = $("scrim");
    bar.classList.remove("open");
    scrim.classList.remove("show");
    document.body.classList.remove("no-scroll");
    $("setupBtn").setAttribute("aria-expanded", "false");
    setTimeout(function () { if (!bar.classList.contains("open")) scrim.hidden = true; }, 230);
  }

  /* ------------------------------------------------------------------ *
   * Tabs
   * ------------------------------------------------------------------ */
  function showPanel(name) {
    Array.prototype.forEach.call(document.querySelectorAll(".tab"), function (t) {
      t.classList.toggle("active", t.dataset.panel === name);
      t.setAttribute("aria-selected", t.dataset.panel === name ? "true" : "false");
    });
    Array.prototype.forEach.call(document.querySelectorAll(".panel"), function (p) {
      p.classList.toggle("active", p.id === "panel-" + name);
    });
    closeSheet();
    window.scrollTo({ top: 0, behavior: "smooth" });
    if (name === "history") renderHistory();
  }

  /* ------------------------------------------------------------------ *
   * Import flow
   * ------------------------------------------------------------------ */
  var state = { table: null, fileName: "", mapping: null };

  function renderPreview(table, fileName) {
    var head = "<table><thead><tr>" + table.headers.map(function (h) {
      return "<th>" + esc(h) + "</th>";
    }).join("") + "</tr></thead><tbody>";
    var body = table.rows.slice(0, 10).map(function (r) {
      return "<tr>" + table.headers.map(function (h) {
        var v = r[h];
        if (v instanceof Date) v = C.parseDateValue(v);
        return "<td>" + esc(v) + "</td>";
      }).join("") + "</tr>";
    }).join("");
    $("previewTable").innerHTML = head + body + "</tbody></table>";
    $("previewBadge").textContent = table.rows.length.toLocaleString("en-IN") + " rows · " + table.headers.length + " columns" +
      (table.sheetName ? " · sheet: " + table.sheetName : "");
    $("fileInfo").textContent = fileName + " — " + table.rows.length.toLocaleString("en-IN") + " rows detected.";
  }

  function renderMapping(headers, detected) {
    var options = ['<option value="(Not mapped)">(Not mapped)</option>'].concat(
      headers.map(function (h) { return '<option value="' + esc(h) + '">' + esc(h) + "</option>"; })
    ).join("");
    $("mapGrid").innerHTML = C.FIELDS.map(function (f, i) {
      var det = detected[f] || "(Not mapped)";
      return '<label class="field"><span>' + esc(f) + "</span>" +
        '<select data-field="' + esc(f) + '" class="' + (det !== "(Not mapped)" ? "mapped" : "") + '">' +
        options.replace('value="' + esc(det) + '"', 'value="' + esc(det) + '" selected') +
        "</select></label>";
    }).join("");
    Array.prototype.forEach.call($("mapGrid").querySelectorAll("select"), function (sel) {
      sel.addEventListener("change", function () {
        sel.classList.toggle("mapped", sel.value !== "(Not mapped)");
      });
    });
    readMapping();
  }

  function readMapping() {
    var m = {};
    Array.prototype.forEach.call($("mapGrid").querySelectorAll("select"), function (sel) {
      m[sel.dataset.field] = sel.value;
    });
    state.mapping = m;
    return m;
  }

  function dbTypeFor(label) {
    if (label === "Books - Purchase Register") return "BOOKS-PURCHASE";
    if (label === "Books - Sales Register") return "BOOKS-SALES";
    return label;
  }

  function handleFile(file) {
    if (!file) return;
    state.fileName = file.name;
    $("fileInfo").textContent = "Reading " + file.name + "…";
    readFile(file).then(function (table) {
      if (!table.rows.length) throw new Error("No data rows found in this file.");
      state.table = table;
      renderPreview(table, file.name);
      $("previewCard").classList.remove("hidden");
      $("mappingCard").classList.remove("hidden");
      renderMapping(table.headers, C.autoMapColumns(table.headers));
      $("importMsg").innerHTML = "";
    }).catch(function (err) {
      state.table = null;
      $("previewCard").classList.add("hidden");
      $("mappingCard").classList.add("hidden");
      $("fileInfo").innerHTML = '<span class="chip bad">' + esc(err.message || String(err)) + "</span>";
    });
  }

  function doImport() {
    var gstin = $("gstin").value.trim().toUpperCase();
    var label = $("returnType").value;
    var dbType = dbTypeFor(label);
    var mapping = readMapping();
    var msg = $("importMsg");

    if (gstin.length !== 15) {
      msg.innerHTML = '<div class="notice warn">Enter your own valid 15-character GSTIN in the sidebar.</div>';
      return;
    }
    if (!state.table) {
      msg.innerHTML = '<div class="notice warn">Choose a file to import first.</div>';
      return;
    }
    if (dbType !== "GSTR-3B" && mapping["Invoice No"] === "(Not mapped)") {
      msg.innerHTML = '<div class="notice warn">Invoice No must be mapped for invoice-level reconciliation.</div>';
      return;
    }

    var p = currentPeriod();
    var rows = C.standardize(state.table.rows, mapping, {
      gstin: gstin, period: p.importPeriod, returnType: dbType, sourceName: state.fileName
    });
    rows.forEach(function (r) { r._gstin_type = gstin + "|" + dbType; });

    $("importBtn").disabled = true;
    msg.innerHTML = '<div class="hint">Saving ' + rows.length.toLocaleString("en-IN") + " rows…</div>";

    addAll("rows", rows).then(function () {
      return addAll("imports", [{
        gstin: gstin, period: p.importPeriod, return_type: dbType,
        filename: state.fileName, row_count: rows.length,
        imported_at: new Date().toISOString().replace("T", " ").slice(0, 19)
      }]);
    }).then(function () {
      msg.innerHTML = '<div class="notice" style="background:var(--ok-soft);border:1px solid #bfe3dc;color:#0f5c55">' +
        "Imported " + rows.length.toLocaleString("en-IN") + " rows as <strong>" + esc(dbType) +
        "</strong> for period " + esc(p.importPeriod) + ".</div>";
      updateStorageInfo();
    }).catch(function (err) {
      msg.innerHTML = '<div class="notice warn">Could not save rows: ' + esc(err && err.message ? err.message : String(err)) + "</div>";
    }).then(function () {
      $("importBtn").disabled = false;
    });
  }

  /* ------------------------------------------------------------------ *
   * Reconciliation
   * ------------------------------------------------------------------ */
  var lastResult = null;

  function statusChip(status) {
    var cls = status === "Matched" ? "ok" : (status.indexOf("mismatch") !== -1 ? "bad" : "warn");
    return '<span class="chip ' + cls + '">' + esc(status) + "</span>";
  }

  function runReco() {
    var gstin = $("gstin").value.trim().toUpperCase();
    var choice = $("recoChoice").value;
    var tol = parseFloat($("tolerance").value) || 0;
    var msg = $("recoMsg");
    var results = $("recoResults");

    if (gstin.length !== 15) {
      msg.innerHTML = '<span class="chip bad">Enter a valid 15-character GSTIN in the sidebar first.</span>';
      return;
    }
    var pair = choice === "GSTR-2A vs Books Purchase" ? ["GSTR-2A", "BOOKS-PURCHASE"]
      : choice === "GSTR-1 vs Books Sales" ? ["GSTR-1", "BOOKS-SALES"]
        : ["GSTR-2B", "BOOKS-PURCHASE"];

    var p = currentPeriod();
    msg.innerHTML = '<span class="hint">Running reconciliation…</span>';

    Promise.all([rowsFor(gstin, pair[0]), rowsFor(gstin, pair[1])]).then(function (res) {
      var left = C.filterRows(res[0], p.start, p.end, pair[0]);
      var right = C.filterRows(res[1], p.start, p.end, pair[1]);
      if (!left.length || !right.length) {
        results.classList.add("hidden");
        msg.innerHTML = '<span class="chip warn">Both datasets are needed. Found ' + left.length +
          " rows for " + esc(pair[0]) + " and " + right.length + " rows for " + esc(pair[1]) +
          " in " + fmtD(p.start) + " – " + fmtD(p.end) + ".</span>";
        return;
      }
      var out = C.compareFrames(left, right, tol);
      lastResult = out;
      renderReco(out, left.length, right.length);
      msg.innerHTML = "";
    }).catch(function (err) {
      msg.innerHTML = '<span class="chip bad">' + esc(err && err.message ? err.message : String(err)) + "</span>";
    });
  }

  function renderReco(rows, nLeft, nRight) {
    var matched = rows.filter(function (r) { return r.Status === "Matched"; }).length;
    var onlyPortal = rows.filter(function (r) { return r.Status === "Only in Portal"; }).length;
    var onlyBooks = rows.filter(function (r) { return r.Status === "Only in Books"; }).length;
    var exceptions = rows.length - matched;

    $("recoMetrics").innerHTML = [
      { k: "Matched", v: matched, c: "ok" },
      { k: "Only in Portal", v: onlyPortal, c: "warn" },
      { k: "Only in Books", v: onlyBooks, c: "warn" },
      { k: "Exceptions", v: exceptions, c: exceptions ? "bad" : "ok" }
    ].map(function (m) {
      return '<div class="metric ' + m.c + '"><div class="k">' + m.k + '</div><div class="v">' +
        m.v.toLocaleString("en-IN") + "</div></div>";
    }).join("");

    $("recoBadge").textContent = nLeft + " portal rows · " + nRight + " books rows · " + rows.length + " invoice keys";

    var head = "<table><thead><tr>" +
      "<th>Invoice No</th><th>Party GSTIN</th>" +
      '<th class="num">Portal taxable</th><th class="num">Portal tax</th>' +
      '<th class="num">Books taxable</th><th class="num">Books tax</th>' +
      "<th>Status</th></tr></thead><tbody>";
    var body = rows.map(function (r) {
      return "<tr><td>" + esc(r["Invoice Key"]) + "</td><td>" + esc(r["GSTIN"]) + "</td>" +
        '<td class="num">' + C.inr(r["Taxable Value_Portal"]) + "</td>" +
        '<td class="num">' + C.inr(r["Total Tax_Portal"]) + "</td>" +
        '<td class="num">' + C.inr(r["Taxable Value_Books"]) + "</td>" +
        '<td class="num">' + C.inr(r["Total Tax_Books"]) + "</td>" +
        "<td>" + statusChip(r.Status) + "</td></tr>";
    }).join("");
    $("recoTable").innerHTML = head + body + "</tbody></table>";
    $("recoResults").classList.remove("hidden");
  }

  function downloadXlsx(rows, sheetName, fileName) {
    var ws = XLSX.utils.json_to_sheet(rows);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, fileName);
  }

  /* ------------------------------------------------------------------ *
   * GSTR-3B summary
   * ------------------------------------------------------------------ */
  var THREE_B = [
    ["Outward taxable value", "outward"],
    ["IGST liability", "igst"],
    ["CGST liability", "cgst"],
    ["SGST/UTGST liability", "sgst"],
    ["Cess liability", "cess"],
    ["Eligible ITC", "itc"],
    ["RCM liability", "rcm"]
  ];

  function renderSummary() {
    var head = "<table><thead><tr><th>Particulars</th>" +
      '<th class="num">Portal (₹)</th><th class="num">Books (₹)</th>' +
      '<th class="num">Difference (Books − Portal)</th></tr></thead><tbody>';
    var body = THREE_B.map(function (f) {
      return "<tr><td>" + esc(f[0]) + "</td>" +
        '<td class="num"><input type="number" step="0.01" min="0" data-p="' + f[1] + '" value="0" style="text-align:right"></td>' +
        '<td class="num"><input type="number" step="0.01" min="0" data-b="' + f[1] + '" value="0" style="text-align:right"></td>' +
        '<td class="num" id="diff-' + f[1] + '">0.00</td></tr>';
    }).join("");
    $("summaryTable").innerHTML = head + body + "</tbody></table>";
    Array.prototype.forEach.call($("summaryTable").querySelectorAll("input"), function (inp) {
      inp.addEventListener("input", updateSummaryDiffs);
    });
    updateSummaryDiffs();
  }

  function summaryValues() {
    return THREE_B.map(function (f) {
      var p = $("summaryTable").querySelector('input[data-p="' + f[1] + '"]');
      var b = $("summaryTable").querySelector('input[data-b="' + f[1] + '"]');
      return {
        Particulars: f[0],
        Portal: p ? (parseFloat(p.value) || 0) : 0,
        Books: b ? (parseFloat(b.value) || 0) : 0
      };
    });
  }

  function updateSummaryDiffs() {
    summaryValues().forEach(function (r, i) {
      var key = THREE_B[i][1];
      var cell = $("diff-" + key);
      if (!cell) return;
      var d = Math.round((r.Books - r.Portal) * 100) / 100;
      cell.textContent = C.inr(d);
      cell.style.color = Math.abs(d) > 0.009 ? "var(--bad)" : "var(--muted)";
      cell.style.fontWeight = Math.abs(d) > 0.009 ? "650" : "400";
    });
  }

  /* ------------------------------------------------------------------ *
   * History + storage
   * ------------------------------------------------------------------ */
  function renderHistory() {
    var gstin = $("gstin").value.trim().toUpperCase();
    var box = $("historyTable");
    if (gstin.length !== 15) {
      box.innerHTML = '<div class="empty">Enter your GSTIN in the sidebar to see import history.</div>';
      return;
    }
    importsFor(gstin).then(function (list) {
      if (!list.length) {
        box.innerHTML = '<div class="empty">No imports yet for ' + esc(gstin) + ".</div>";
        return;
      }
      var head = "<table><thead><tr><th>Period</th><th>Type</th><th>File</th>" +
        '<th class="num">Rows</th><th>Imported at</th></tr></thead><tbody>';
      var body = list.map(function (r) {
        return "<tr><td>" + esc(r.period) + "</td><td>" + esc(r.return_type) + "</td><td>" + esc(r.filename) +
          '</td><td class="num">' + Number(r.row_count).toLocaleString("en-IN") + "</td><td>" + esc(r.imported_at) + "</td></tr>";
      }).join("");
      box.innerHTML = head + body + "</tbody></table>";
    }).catch(function (err) {
      box.innerHTML = '<div class="empty">Could not read history: ' + esc(err && err.message ? err.message : String(err)) + "</div>";
    });
  }

  function updateStorageInfo() {
    if (!navigator.storage || !navigator.storage.estimate) {
      $("storageInfo").textContent = "Stored locally in this browser.";
      return;
    }
    navigator.storage.estimate().then(function (est) {
      var mb = (est.usage || 0) / 1048576;
      $("storageInfo").textContent = "Using " + mb.toFixed(2) + " MB in this browser profile.";
    }).catch(function () {
      $("storageInfo").textContent = "Stored locally in this browser.";
    });
  }

  /* ------------------------------------------------------------------ *
   * Wire up
   * ------------------------------------------------------------------ */
  function init() {
    $("tabs").addEventListener("click", function (e) {
      var t = e.target.closest(".tab");
      if (t) showPanel(t.dataset.panel);
    });

    $("gstin").addEventListener("input", function () {
      this.value = this.value.toUpperCase().replace(/[^0-9A-Z]/g, "");
      updatePeriodCaption();
    });

    ["periodMode", "fy"].forEach(function (id) {
      $(id).addEventListener("change", renderPeriodExtra);
    });
    $("periodExtra").addEventListener("change", updatePeriodCaption);

    $("setupBtn").addEventListener("click", function () {
      if ($("sidebar").classList.contains("open")) closeSheet(); else openSheet();
    });
    $("closeSheet").addEventListener("click", closeSheet);
    $("scrim").addEventListener("click", closeSheet);
    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape") closeSheet();
    });

    $("fileInput").addEventListener("change", function (e) {
      handleFile(e.target.files && e.target.files[0]);
    });

    $("returnType").addEventListener("change", function () {
      $("threeBWarn").classList.toggle("hidden", this.value !== "GSTR-3B");
    });

    $("resetMapping").addEventListener("click", function () {
      if (!state.table) return;
      renderMapping(state.table.headers, C.autoMapColumns(state.table.headers));
    });

    $("importBtn").addEventListener("click", doImport);
    $("runReco").addEventListener("click", runReco);

    $("downloadReco").addEventListener("click", function () {
      if (!lastResult || !lastResult.length) return;
      downloadXlsx(lastResult, "Reconciliation", "GST_Reconciliation_Report.xlsx");
    });

    $("downloadSummary").addEventListener("click", function () {
      var rows = summaryValues().map(function (r) {
        return {
          Particulars: r.Particulars, Portal: r.Portal, Books: r.Books,
          "Difference (Books - Portal)": Math.round((r.Books - r.Portal) * 100) / 100
        };
      });
      downloadXlsx(rows, "3B Summary", "GSTR3B_Summary_Reconciliation.xlsx");
    });

    $("refreshHistory").addEventListener("click", renderHistory);

    $("clearData").addEventListener("click", function () {
      if (!confirm("Delete all imported rows and import history stored in this browser? This cannot be undone.")) return;
      clearEverything().then(function () {
        $("importMsg").innerHTML = "";
        $("recoResults").classList.add("hidden");
        renderHistory();
        updateStorageInfo();
      });
    });

    renderPeriodExtra();
    renderSummary();
    updateStorageInfo();
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
