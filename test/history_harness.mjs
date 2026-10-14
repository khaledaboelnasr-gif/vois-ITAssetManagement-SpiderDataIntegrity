// Test bench for history_store_officescript.ts and reference_extract_officescript.ts (no real data).
// Usage: node history_harness.mjs ../history_store_officescript.ts ../asset_audit_officescript.ts ../reference_extract_officescript.ts synthetic.json
import { stripTypeScriptTypes } from "node:module";
import fs from "node:fs";
import assert from "node:assert/strict";
const [storeSrc, mainSrc, extractSrc, dataPath] = process.argv.slice(2);
const load = (p) => new Function("ExcelScript", stripTypeScriptTypes(fs.readFileSync(p, "utf8")) + "\nreturn { main };");
const ExcelScript = new Proxy({}, { get: () => new Proxy({}, { get: (t, q) => String(q) }) });
const store = load(storeSrc)(ExcelScript).main;
const any = new Proxy(function () {}, { get: (t, p) => p === Symbol.toPrimitive ? () => "x" : any, apply: () => any });

// ---- minimal mock: one sheet = sparse cell map; a table = header address + row count
class Sheet {
  constructor(name, values) { this.name = name; this.cells = new Map(); this.textCols = new Set(); (values || []).forEach((r, i) => r.forEach((v, j) => this.cells.set(`${i},${j}`, v))); }
  getName() { return this.name; }
  get(r, c) { const v = this.cells.get(`${r},${c}`); return v === undefined ? "" : v; }
  size() { let h = 0, w = 0; for (const k of this.cells.keys()) { const [r, c] = k.split(",").map(Number); h = Math.max(h, r + 1); w = Math.max(w, c + 1); } return [h, w]; }
  getUsedRange() { const [h, w] = this.size(); return h ? new Range(this, 0, 0, h, w) : undefined; }
  getRangeByIndexes(r, c, h, w) { return new Range(this, r, c, h, w); }
}
class Range {
  constructor(ws, r, c, h, w) { if (h < 1 || w < 1) throw new Error(`empty range ${h}x${w}`); Object.assign(this, { ws, r, c, h, w }); }
  getValues() { return Array.from({ length: this.h }, (_, i) => Array.from({ length: this.w }, (_, j) => this.ws.get(this.r + i, this.c + j))); }
  setValues(v) {
    if (v.length !== this.h || v.some(x => x.length !== this.w)) throw new Error("setValues dims");
    v.forEach((row, i) => row.forEach((x, j) => {
      if (x === undefined || x === null) throw new Error("bad value");
      // Excel turns "2026-09" into a date serial unless the cell is text-formatted
      const col = this.c + j;
      const y = typeof x === "string" && /^\d{4}-\d{2}$/.test(x) && !this.ws.textCols.has(col)
        ? Math.round(Date.UTC(+x.slice(0, 4), +x.slice(5) - 1, 1) / 86400000) + 25569 : x;
      this.ws.cells.set(`${this.r + i},${col}`, y);
    }));
  }
  setNumberFormat(f) { if (f === "@") for (let j = 0; j < this.w; j++) this.ws.textCols.add(this.c + j); }
  getRowIndex() { return this.r; } getColumnIndex() { return this.c; } getColumnCount() { return this.w; } getRowCount() { return this.h; }
}
class Table {
  constructor(ws, r, c, w, rows) { Object.assign(this, { ws, r, c, w, rows }); }
  setName(n) { this.name = n; } getName() { return this.name; }
  getWorksheet() { return this.ws; }
  getRowCount() { return this.rows; }
  getHeaderRowRange() { return new Range(this.ws, this.r, this.c, 1, this.w); }
  getRangeBetweenHeaderAndTotal() { return new Range(this.ws, this.r + 1, this.c, Math.max(this.rows, 1), this.w); }
  resize(rng) { assert.equal(rng.r, this.r); this.w = rng.w; this.rows = rng.h - 1; }
}
function workbook(sheets) {
  return { sheets: sheets || [], tables: [],
    getWorksheets() { return this.sheets; },
    getWorksheet(n) { return this.sheets.find(s => s.name === n); },
    addWorksheet(n) { const s = new Sheet(n); this.sheets.push(s); return s; },
    getTable(n) { return this.tables.find(t => t.name === n); },
    // Like Excel: a header-only range becomes a table with one blank data row
    addTable(rng) { const t = new Table(rng.ws, rng.r, rng.c, rng.w, Math.max(rng.h - 1, 1)); this.tables.push(t); return t; } };
}

// ---- 1. audit script on the synthetic export -> historyRows (harness.mjs-style mock for the main script)
const data = JSON.parse(fs.readFileSync(dataPath, "utf8"));
const mainWb = (names) => {
  const sheets = [];
  const mk = (name, vals) => {
    const w = vals ? Math.max(...vals.map(r => r.length)) : 0;
    const v = vals ? vals.map(r => r.concat(Array(w - r.length).fill(""))) : null;
    const s = { name, getName: () => name,
      getUsedRange: () => v ? { getRowCount: () => v.length, getColumnCount: () => w, getRowIndex: () => 0, getColumnIndex: () => 0 } : undefined,
      getRangeByIndexes: (r, c, h, ww) => new Proxy({ getValues: () => v ? v.slice(r, r + h).map(x => x.slice(c, c + ww)) : [], setValues() {} },
        { get: (t, p) => p in t ? t[p] : any }),
      addTable: () => any, delete() { sheets.splice(sheets.indexOf(s), 1); } };
    return new Proxy(s, { get: (t, p) => p in t ? t[p] : any });
  };
  names.forEach(n => sheets.push(mk(n, data[n])));
  return new Proxy({ getWorksheets: () => sheets, getWorksheet: n => sheets.find(s => s.name === n),
    addWorksheet: n => { const s = mk(n); sheets.push(s); return s; } }, { get: (t, p) => p in t ? t[p] : any });
};
const CUR = "2026-10 Current Export";
const runMain = (date, historyJson, mappingsJson, usersJson, devicesJson) => load(mainSrc)(ExcelScript).main(mainWb([CUR]),
  date, undefined, historyJson || undefined, undefined, undefined, false, mappingsJson || undefined, usersJson, devicesJson);

// ---- 2. reference extract: users and devices workbooks -> compact JSON -> audit script
const extract = load(extractSrc)(ExcelScript).main;
const ux = extract(workbook([new Sheet("Sheet1", [["Report: Entra users"], []].concat(data["Users (Entra)"]))]), "users");
const dx = extract(workbook([new Sheet("Intune", data["Intune Devices"])]), "devices");
console.log("extract users:", ux.rows, ux.columns.join("; "), "| devices:", dx.rows, dx.columns.join("; "), "| chars", ux.chars + dx.chars);
assert.equal(ux.rows, data["Users (Entra)"].length - 1);
assert.equal(dx.rows, data["Intune Devices"].length - 1);
assert.equal(JSON.parse(ux.json).filter(r => r[2] === "0").length, data["Users (Entra)"].filter(r => r[2] === "False").length, "inactive users kept");
assert.throws(() => extract(workbook([new Sheet("X", [["a", "b"]])]), "users"), /No users sheet/);

const r1 = runMain("2026-09-30", "", "", ux.json, dx.json);
const sum = (rs, metric) => rs.filter(r => r.metric === metric).reduce((a, r) => a + r.value, 0);
console.log("audit Sep: period", r1.period, "| historyRows", r1.historyRows.length, "records", sum(r1.historyRows, "#records"), "| leavers", r1.leaversHoldingAssets);
assert.ok(r1.leaversHoldingAssets > 0, "leavers found through usersJson");

// ---- 3. empty AuditHistory.xlsx: read creates the table and returns nothing
const hx = workbook();
let res = store(hx, "read");
assert.equal(res.rows, 0); assert.equal(res.historyJson, "[]"); assert.equal(res.mappingsJson, "");
const readBack = (...a) => JSON.parse(store(hx, "read", undefined, undefined, ...a).historyJson);

// ---- 4. save, then read back
res = store(hx, "save", JSON.stringify(r1.historyRows), "20260930-090000_a_run1", undefined, undefined, "a@vois.com", "https://link/1");
assert.equal(res.added, r1.historyRows.length);
let back = readBack();
assert.equal(back.length, r1.historyRows.length); assert.equal(sum(back, "#records"), sum(r1.historyRows, "#records"));
assert.ok(back.every(r => /^\d{4}-\d{2}$/.test(r.period)), "periods stay YYYY-MM text");

// ---- 5. a retried flow run (same RunId) adds nothing; a second person's run is kept but counted once
assert.equal(store(hx, "save", JSON.stringify(r1.historyRows), "20260930-090000_a_run1").added, 0);
store(hx, "save", JSON.stringify(r1.historyRows), "20260930-100000_b_run2", undefined, undefined, "b@vois.com");
res = store(hx, "read");
assert.equal(sum(JSON.parse(res.historyJson), "#records"), sum(r1.historyRows, "#records"), "not double counted");
assert.ok(res.used.every(u => u.endsWith("run2")), "newest complete run is used");

// ---- 6. a newer partial run (a third of one market) is kept but not used
const mk1 = r1.historyRows[0].market;
store(hx, "save", JSON.stringify(r1.historyRows.filter(r => r.market === mk1).map(r => ({ ...r, value: Math.floor(r.value / 3) }))), "20260930-120000_c_partial");
res = store(hx, "read");
assert.equal(sum(JSON.parse(res.historyJson).filter(r => r.market === mk1), "#records"), sum(r1.historyRows.filter(r => r.market === mk1), "#records"));

// ---- 7. Runs log: one line per saved run, with the integrity score
const log = hx.getTable("Runs");
assert.equal(log.getRowCount(), 3);
const logRows = log.getWorksheet().getRangeByIndexes(1, 0, 3, 8).getValues();
console.log("runs log:", logRows.map(r => [r[0], r[2], r[4], r[5], r[6]].join(" | ")));
assert.equal(logRows[0][2], "a@vois.com"); assert.equal(logRows[0][6], r1.integrityScore); assert.equal(logRows[0][7], "https://link/1");

// ---- 8. Mappings table -> mappingsJson
const ms = hx.addWorksheet("Mappings");
ms.getRangeByIndexes(0, 0, 3, 4).setValues([["Kind", "Value", "MapsTo", "Note"], ["Check", "NOT_VERIFIED", "Off", "no audit dates yet"], ["", "", "", ""]]);
const mt = hx.addTable(ms.getRangeByIndexes(0, 0, 3, 4)); mt.setName("Mappings");
res = store(hx, "read");
assert.equal(JSON.parse(res.mappingsJson).length, 1);

// ---- 9. next month's audit compares with September
const r2 = runMain("2026-10-14", res.historyJson, res.mappingsJson, ux.json, dx.json);
console.log("audit Oct:", r2.runInfo.filter(x => /History periods|Mappings/.test(x)).join(" || "));
assert.ok(r2.runInfo.some(x => /History periods available: 2026-09, 2026-10/.test(x)));
assert.ok(/Previous = each market's last audit/.test(r2.agentInput), "agent brief has the comparison");
assert.ok(!r2.topChecks.some(c => /physically verified/.test(c)), "mapping switched NOT_VERIFIED off");
store(hx, "save", JSON.stringify(r2.historyRows), "20261014-090000_a_oct");
assert.deepEqual(store(hx, "read").periods, ["2026-09", "2026-10"]);
assert.equal(store(hx, "read", undefined, undefined, 1).periods.length, 1, "months filter");
console.log("history store + extract: all checks passed");
