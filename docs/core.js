/* GST Reconciliation — core logic (no DOM, testable in Node) */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.GSTCore = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var FIELDS = ["GSTIN", "Party Name", "Invoice No", "Invoice Date", "Invoice Type",
    "Taxable Value", "IGST", "CGST", "SGST", "Cess", "Total Tax"];

  var TEXT_FIELDS = ["GSTIN", "Party Name", "Invoice No", "Invoice Date", "Invoice Type"];
  var VALUE_FIELDS = ["Taxable Value", "IGST", "CGST", "SGST", "Cess", "Total Tax"];

  var ALIASES = {
    "Party Name": ["party name", "supplier name", "customer name", "vendor name", "trade name", "legal name", "name of supplier", "name of recipient", "receiver name", "recipient name", "buyer name", "party"],
    "Invoice Type": ["invoice type", "document type", "doc type", "voucher type", "supply type", "type of invoice"],
    "GSTIN": ["gstin", "gstin of supplier", "supplier gstin", "gstin of recipient", "recipient gstin", "ctin", "customer gstin", "party gstin", "gst number", "gstin uin of recipient", "gstin uin of supplier"],
    "Invoice No": ["invoice no", "invoice number", "inum", "bill no", "bill number", "document number", "doc no", "voucher no", "voucher number", "inv no", "invoice"],
    "Invoice Date": ["invoice date", "date", "idt", "bill date", "document date", "voucher date"],
    "Taxable Value": ["taxable value", "taxable amount", "taxable", "assessable value", "taxable amt"],
    "IGST": ["igst", "integrated tax", "igst amount", "integrated tax amount"],
    "CGST": ["cgst", "central tax", "cgst amount", "central tax amount"],
    "SGST": ["sgst", "utgst", "state tax", "sgst amount", "state tax amount", "utgst amount"],
    "Cess": ["cess", "cess amount"],
    "Total Tax": ["total tax", "tax amount", "total gst", "gst amount", "total tax amount"]
  };

  var TOKEN_EXPAND = {
    amt: "amount", amnt: "amount", amts: "amount",
    dt: "date", inv: "invoice", invs: "invoice",
    no: "number", num: "number", nos: "number", numbers: "number",
    val: "value", vals: "value", desc: "description",
    taxbl: "taxable", tot: "total", qty: "quantity"
  };

  /* Normalise a heading, and expand common abbreviations so that
     "GST Amt" matches "gst amount" and "Voucher No" matches "voucher number". */
  function normCol(value) {
    var base = String(value === null || value === undefined ? "" : value)
      .trim().toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
    if (!base) return "";
    return base.split(" ").map(function (t) { return TOKEN_EXPAND[t] || t; }).join(" ");
  }

  /* Map detected source columns onto the 11 standard fields. */
  function autoMapColumns(headers) {
    var mapped = {};
    var normalized = headers.map(function (h) { return { col: h, nc: normCol(h) }; });
    FIELDS.forEach(function (field) {
      var aliasNorms = (ALIASES[field] || []).map(normCol);
      var found = null;
      normalized.forEach(function (n) {
        if (found === null && aliasNorms.indexOf(n.nc) !== -1) found = n.col;
      });
      if (found === null) {
        normalized.forEach(function (n) {
          if (found !== null) return;
          var hit = aliasNorms.some(function (a) {
            return a.length >= 4 && (n.nc.indexOf(a) !== -1 || a.indexOf(n.nc) !== -1);
          });
          if (hit) found = n.col;
        });
      }
      mapped[field] = found === null ? "(Not mapped)" : found;
    });
    return mapped;
  }

  var BLANKS = ["", "nan", "none", "null", "-", "na", "n/a"];

  function toNumber(value) {
    if (value === null || value === undefined) return 0;
    if (typeof value === "number") return isFinite(value) ? value : 0;
    var s = String(value).replace(/,/g, "").replace(/\u20b9/g, "").replace(/\u00a0/g, " ").trim();
    if (BLANKS.indexOf(s.toLowerCase()) !== -1) return 0;
    var neg = /^\(.*\)$/.test(s);
    if (neg) s = s.slice(1, -1);
    var n = parseFloat(s.replace(/[^0-9eE+\-.]/g, ""));
    if (!isFinite(n)) return 0;
    return neg ? -n : n;
  }

  function normalizeInvoice(value) {
    if (value === null || value === undefined) return "";
    var s = String(value).replace(/\s+/g, "").trim().toUpperCase();
    if (BLANKS.indexOf(s.toLowerCase()) !== -1) return "";
    return s;
  }

  function pad2(n) { return (n < 10 ? "0" : "") + n; }

  function fmtDate(d) {
    return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
  }

  var MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12 };

  /* Day-first parsing, matching the original pandas dayfirst=True behaviour. */
  function parseDateValue(value) {
    if (value === null || value === undefined) return "";
    if (value instanceof Date && !isNaN(value.getTime())) return fmtDate(value);
    if (typeof value === "number" && isFinite(value)) {
      if (value > 20000 && value < 80000) {
        var ms = Math.round((value - 25569) * 86400000);
        var d = new Date(ms);
        if (!isNaN(d.getTime())) {
          return d.getUTCFullYear() + "-" + pad2(d.getUTCMonth() + 1) + "-" + pad2(d.getUTCDate());
        }
      }
      return "";
    }
    var s = String(value).trim();
    if (!s || BLANKS.indexOf(s.toLowerCase()) !== -1) return "";
    var iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (iso) return iso[1] + "-" + pad2(+iso[2]) + "-" + pad2(+iso[3]);
    var dmy = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
    if (dmy) {
      var day = +dmy[1], mon = +dmy[2], yr = +dmy[3];
      if (yr < 100) yr += yr < 70 ? 2000 : 1900;
      if (mon >= 1 && mon <= 12 && day >= 1 && day <= 31) return yr + "-" + pad2(mon) + "-" + pad2(day);
    }
    var named = s.match(/^(\d{1,2})[\-\s]([A-Za-z]{3,9})[\-\s](\d{2,4})$/);
    if (named) {
      var m2 = MONTHS[named[2].slice(0, 3).toLowerCase()];
      var y2 = +named[3]; if (y2 < 100) y2 += y2 < 70 ? 2000 : 1900;
      if (m2) return y2 + "-" + pad2(m2) + "-" + pad2(+named[1]);
    }
    var fallback = new Date(s);
    if (!isNaN(fallback.getTime())) return fmtDate(fallback);
    return "";
  }

  function cleanText(v) {
    if (v === null || v === undefined) return "";
    var s = String(v).trim();
    if (BLANKS.indexOf(s.toLowerCase()) !== -1) return "";
    return s;
  }

  /* Turn raw source rows into the standard schema. */
  function standardize(rows, mapping, ctx) {
    ctx = ctx || {};
    return rows.map(function (row) {
      var out = {};
      FIELDS.forEach(function (field) {
        var col = mapping[field];
        var raw = (col && col !== "(Not mapped)") ? row[col] : "";
        if (TEXT_FIELDS.indexOf(field) !== -1) out[field] = cleanText(raw);
        else out[field] = toNumber(raw);
      });
      out["GSTIN"] = out["GSTIN"].toUpperCase();
      out["Invoice Date"] = parseDateValue((mapping["Invoice Date"] && mapping["Invoice Date"] !== "(Not mapped)") ? row[mapping["Invoice Date"]] : "");
      var taxSum = out["IGST"] + out["CGST"] + out["SGST"] + out["Cess"];
      if (out["Total Tax"] === 0 && taxSum !== 0) out["Total Tax"] = taxSum;
      out["Invoice Key"] = normalizeInvoice(out["Invoice No"]);
      out["App GSTIN"] = ctx.gstin || "";
      out["Period"] = ctx.period || "";
      out["Return Type"] = ctx.returnType || "";
      out["Source File"] = ctx.sourceName || "";
      return out;
    });
  }

  function monthPeriods(startISO, endISO) {
    var s = startISO.split("-").map(Number), e = endISO.split("-").map(Number);
    var y = s[0], m = s[1], out = [];
    while (y < e[0] || (y === e[0] && m <= e[1])) {
      out.push(pad2(m) + y);
      if (m === 12) { y += 1; m = 1; } else { m += 1; }
      if (out.length > 240) break;
    }
    return out;
  }

  function lastDay(y, m) { return new Date(y, m, 0).getDate(); }

  function fyDates(fy) {
    var startYear = parseInt(String(fy).split("-")[0], 10);
    return { start: startYear + "-04-01", end: (startYear + 1) + "-03-31" };
  }

  function quarterDates(fy, quarter) {
    var startYear = parseInt(String(fy).split("-")[0], 10);
    var map = { Q1: [4, 6], Q2: [7, 9], Q3: [10, 12], Q4: [1, 3] };
    var q = map[quarter] || map.Q1;
    var year = quarter === "Q4" ? startYear + 1 : startYear;
    return {
      start: year + "-" + pad2(q[0]) + "-01",
      end: year + "-" + pad2(q[1]) + "-" + pad2(lastDay(year, q[1]))
    };
  }

  function monthDates(fy, monthNumber) {
    var startYear = parseInt(String(fy).split("-")[0], 10);
    var year = monthNumber <= 3 ? startYear + 1 : startYear;
    return {
      start: year + "-" + pad2(monthNumber) + "-01",
      end: year + "-" + pad2(monthNumber) + "-" + pad2(lastDay(year, monthNumber))
    };
  }

  function filterRows(rows, startISO, endISO, returnType) {
    var books = returnType === "BOOKS-PURCHASE" || returnType === "BOOKS-SALES";
    if (books) {
      return rows.filter(function (r) {
        return r["Invoice Date"] && r["Invoice Date"] >= startISO && r["Invoice Date"] <= endISO;
      });
    }
    var periods = monthPeriods(startISO, endISO);
    return rows.filter(function (r) { return periods.indexOf(r["Period"]) !== -1; });
  }

  /* Aggregate by invoice key + party GSTIN, then outer-join the two sides. */
  function compareFrames(left, right, tolerance) {
    tolerance = tolerance || 0;
    function aggregate(rows) {
      var map = new Map();
      rows.forEach(function (r) {
        var key = String(r["Invoice Key"] || "").trim().toUpperCase();
        if (!key) return;
        var gstin = String(r["GSTIN"] || "").trim().toUpperCase();
        var id = key + "\u0001" + gstin;
        var acc = map.get(id);
        if (!acc) {
          acc = { "Invoice Key": key, "GSTIN": gstin };
          VALUE_FIELDS.forEach(function (f) { acc[f] = 0; });
          map.set(id, acc);
        }
        VALUE_FIELDS.forEach(function (f) { acc[f] += Number(r[f]) || 0; });
      });
      return map;
    }
    var l = aggregate(left), r = aggregate(right);
    var ids = [];
    l.forEach(function (_, id) { ids.push(id); });
    r.forEach(function (_, id) { if (l.has(id) === false) ids.push(id); });

    var out = [];
    ids.forEach(function (id) {
      var a = l.get(id), b = r.get(id);
      var row = {
        "Invoice Key": (a || b)["Invoice Key"],
        "GSTIN": (a || b)["GSTIN"]
      };
      VALUE_FIELDS.forEach(function (f) {
        row[f + "_Portal"] = a ? a[f] : 0;
        row[f + "_Books"] = b ? b[f] : 0;
      });
      if (a && !b) row["Status"] = "Only in Portal";
      else if (!a && b) row["Status"] = "Only in Books";
      else {
        var diffs = VALUE_FIELDS.filter(function (f) { return Math.abs(a[f] - b[f]) > tolerance; });
        row["Status"] = diffs.length === 0 ? "Matched" : "Value mismatch: " + diffs.join(", ");
      }
      out.push(row);
    });
    out.sort(function (x, y) {
      var sx = x["Status"] === "Matched" ? 1 : 0, sy = y["Status"] === "Matched" ? 1 : 0;
      if (sx !== sy) return sx - sy;
      return String(x["Invoice Key"]).localeCompare(String(y["Invoice Key"]));
    });
    return out;
  }

  function inr(n) {
    var v = Number(n) || 0;
    return v.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  return {
    FIELDS: FIELDS, TEXT_FIELDS: TEXT_FIELDS, VALUE_FIELDS: VALUE_FIELDS, ALIASES: ALIASES,
    normCol: normCol, autoMapColumns: autoMapColumns, toNumber: toNumber,
    normalizeInvoice: normalizeInvoice, parseDateValue: parseDateValue,
    standardize: standardize, monthPeriods: monthPeriods, fyDates: fyDates,
    quarterDates: quarterDates, monthDates: monthDates, filterRows: filterRows,
    compareFrames: compareFrames, inr: inr, pad2: pad2
  };
});
