// Local test bench: runs asset_audit_officescript.ts in Node 22.6+ against a mock ExcelScript API (no real data).
// Usage: node harness.mjs ../asset_audit_officescript.ts synthetic.json [history.json]
// Env: AUDIT=YYYY-MM-DD (default 2026-10-08)  DUMP="Sheet:rows,..."  SHARE=1  MARKET="VOIS EG"  MAPPINGS=<json>
//      NOREF=1 (drop the Users and Intune tabs)  REFJSON=1 (pass them as usersJson / devicesJson instead of tabs)
//      ONLYCUR=1 (drop the previous export)  SNAPSHOT=1 (previous month as a snapshot from an earlier run, not a tab)
//      NOPO=1 (drop the PO tab)  POJSON=1 (pass the PO rows as poJson)  MARK=1 (fill the Physical Check sheet and re-run)
//      HISTOUT=<file> (save historyRows)  BRIEFOUT=<file> (save agentInput)  REPORTOUT=<file> (save the email and PowerPoint prompt)  MAXCALLS=<n> (fail above n Excel calls)
//      TABNAMES=1 (tabs named "previous month export" / "this month export")
import { stripTypeScriptTypes } from "node:module";
import fs from "node:fs";
import assert from "node:assert/strict";
const [src, dataPath, histPath] = process.argv.slice(2);
const js = stripTypeScriptTypes(fs.readFileSync(src, "utf8"));
// Every call into the mock Excel API is counted by method name: each one is a round trip to Excel in the real thing,
// so a count that grows with the number of rows means an Excel call inside a per-row loop (very slow in Excel)
const apiByName = new Map();
const count = name => apiByName.set(name, (apiByName.get(name) || 0) + 1);
let lastName = "";
const any = new Proxy(function () {}, { get: (t, p) => { if (p === Symbol.toPrimitive) return () => "x"; lastName = String(p); return any; },
  apply: () => { count(lastName); return any; } });
const ExcelScript = new Proxy({}, { get: () => new Proxy({}, { get: (t2, q) => String(q) }) });
const wrap = o => new Proxy(o, { get: (t, p) => {
  if (!(p in t)) { lastName = String(p); return any; }
  if (typeof t[p] !== "function") return t[p];
  return (...args) => { count(String(p)); return t[p].apply(t, args); };
} });
const colIdx = a => { let n = 0; for (const ch of a) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
let calls = 0;
// Rough formula check (Excel's own check is stricter): starts with =, balanced quotes and brackets, and every
// 'Sheet'! reference names a sheet that exists
function checkFormula(x, where) {
  const bad = m => { throw new Error(`invalid formula on ${where}: ${m}: ${x}`); };
  if (typeof x !== "string" || !x.startsWith("=")) bad("no leading =");
  if ((x.match(/"/g) || []).length % 2) bad("unbalanced quotes");
  const noStr = x.replace(/"[^"]*"/g, ""), depth = [...noStr].reduce((d, ch) => d < 0 ? d : d + (ch === "(" ? 1 : ch === ")" ? -1 : 0), 0);
  if (depth !== 0) bad("unbalanced brackets");
  for (const m of noStr.matchAll(/'((?:[^']|'')+)'!/g)) { const n = m[1].replace(/''/g, "'"); if (!wb.sheets.some(s => s.name === n)) bad(`no sheet '${n}'`); }
}
class Range {
  constructor(ws, r, c, h, w) { Object.assign(this, { ws, r, c, h, w }); }
  setValues(v) {
    calls++;
    if (v.length !== this.h || v.some(x => x.length !== this.w)) throw new Error(`setValues dims ${v.length}x${v[0]?.length} vs ${this.h}x${this.w} on ${this.ws.name}`);
    v.forEach((row, i) => row.forEach((x, j) => {
      if (x === undefined || x === null || (typeof x === "number" && !isFinite(x))) throw new Error(`bad value ${x} on ${this.ws.name} r${this.r + i} c${this.c + j}`);
      this.ws.cells.set(`${this.r + i},${this.c + j}`, x);
    }));
  }
  setValue(x) { this.setValues([[x]]); }
  getValues() { calls++; const v = this.ws.values.length ? this.ws.values : this.ws.grid(); return v.slice(this.r, this.r + this.h).map(r => r.slice(this.c, this.c + this.w).concat(Array(Math.max(0, this.w - (r.length - this.c))).fill(""))); }
  getRowIndex() { return this.r; } getColumnIndex() { return this.c; } getRowCount() { return this.h; } getColumnCount() { return this.w; }
  getAddress() { const L = c => { let s = ""; c++; while (c) { s = String.fromCharCode(65 + (c - 1) % 26) + s; c = Math.floor((c - 1) / 26); } return s; };
    return `'${this.ws.name.replace(/'/g, "''")}'!${L(this.c)}${this.r + 1}:${L(this.c + this.w - 1)}${this.r + this.h}`; }
  getColumn(i) { return wrap(new Range(this.ws, this.r, this.c + i, this.h, 1)); }
  setNumberFormat(f) { if (typeof f !== "string") throw new Error("setNumberFormat needs string"); this.each(k => this.ws.style(k).nf = f); }
  // Formatting is recorded per cell, for RENDER (a picture of each tab)
  each(fn) { for (let i = 0; i < this.h; i++) for (let j = 0; j < this.w; j++) fn(`${this.r + i},${this.c + j}`); }
  getFormat() {
    const set = (prop, v) => this.each(k => this.ws.style(k)[prop] = v);
    const font = { setBold: v => set("bold", v), setItalic: v => set("italic", v), setColor: v => set("color", v), setSize: v => set("size", v),
      setUnderline: v => set("underline", v), setName() {} };
    return wrap({ getFont: () => wrap(font), getFill: () => wrap({ setColor: v => set("fill", v) }), setWrapText: v => set("wrap", v),
      setColumnWidth: v => { for (let j = 0; j < this.w; j++) this.ws.colW[this.c + j] = v; },
      setVerticalAlignment() {}, setHorizontalAlignment: v => set("align", v) });
  }
  addConditionalFormat() { return any; }
  setFormulas(f) { f.forEach((row, i) => row.forEach((x, j) => { checkFormula(x, this.ws.name); this.ws.style(`${this.r + i},${this.c + j}`).formula = x; })); }
  setHyperlink(h) { this.each(k => { this.ws.style(k).link = h.documentReference || h.address; }); if (h.textToDisplay) this.ws.cells.set(`${this.r},${this.c}`, h.textToDisplay); }
}
class Sheet {
  constructor(name, values) { this.name = name; this.values = values || []; this.cells = new Map(); this.tables = []; this.fmt = new Map();
    this.colW = {}; this.tableRanges = []; this.pivots = []; }
  style(k) { let s = this.fmt.get(k); if (!s) { s = {}; this.fmt.set(k, s); } return s; }
  getName() { return this.name; }
  size() { let h = 0, w = 0; for (const k of this.cells.keys()) { const [r, c] = k.split(",").map(Number); h = Math.max(h, r + 1); w = Math.max(w, c + 1); } return [h, w]; }
  grid() { const [h, w] = this.size(); return Array.from({ length: h }, (_, r) => Array.from({ length: w }, (_, c) => { const v = this.cells.get(`${r},${c}`); return v === undefined ? "" : v; })); }
  getUsedRange() { if (this.values.length) return wrap(new Range(this, 0, 0, this.values.length, this.values[0].length)); const [h, w] = this.size(); return h ? wrap(new Range(this, 0, 0, h, w)) : undefined; }
  getRangeByIndexes(r, c, h, w) { if (h < 1 || w < 1) throw new Error(`empty range ${h}x${w} on ${this.name}`); return wrap(new Range(this, r, c, h, w)); }
  getRange(a) { const m = a.match(/^([A-Z]+)(\d+)$/); return wrap(new Range(this, +m[2] - 1, colIdx(m[1]), 1, 1)); }
  addTable(rng) {
    const self = this, box = { r: rng.r, c: rng.c, h: rng.h, w: rng.w, name: "" };
    // Excel refuses a table that overlaps another table
    this.tableRanges.forEach(t => { if (box.r < t.r + t.h && t.r < box.r + box.h && box.c < t.c + t.w && t.c < box.c + box.w) throw new Error(`table overlaps ${t.name} on ${this.name}`); });
    this.tableRanges.push(box);
    return wrap({ setName(n) { box.name = n; if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(n) || n.length > 255) throw new Error("bad table name " + n); if (wb.tableNames.has(n.toLowerCase())) throw new Error("dup table " + n); wb.tableNames.add(n.toLowerCase()); self.tables.push(n.toLowerCase()); }, setPredefinedTableStyle() {} }); }
  addPivotTable(name, table, dest) {
    const pv = { name, table: table.getName(), r: dest.r, c: dest.c, rows: [], value: "", sum: false };
    this.pivots.push(pv);
    return wrap({ getHierarchy: h => ({ h }), addRowHierarchy: x => { pv.rows.push(x.h); }, addDataHierarchy: x => { pv.value = x.h; return wrap({ setSummarizeBy: f => { pv.sum = f === "sum"; } }); } });
  }
  getFreezePanes() { return any; } setPosition() {} setShowGridlines() {} setTabColor() {} activate() {}
  delete() { wb.sheets = wb.sheets.filter(s => s.name !== this.name); this.tables.forEach(n => wb.tableNames.delete(n)); }
}
const data = JSON.parse(fs.readFileSync(dataPath, "utf8"));
const pad = rows => { const w = Math.max(...rows.map(r => r.length)); return rows.map(r => r.concat(Array(w - r.length).fill(""))); };
const refTabs = ["Users (Entra)", "Intune Devices"];
const PREV_TAB = "2026-09 Previous Export", CUR_TAB = "2026-10 Current Export", PO_TAB = "PO Validation";
const drop = n => ((process.env.NOREF || process.env.REFJSON) && refTabs.includes(n)) || ((process.env.ONLYCUR || process.env.SNAPSHOT) && n === PREV_TAB) ||
  ((process.env.NOPO || process.env.POJSON) && n === PO_TAB);
// TABNAMES=1: the team's own tab names ("previous month export" / "this month export")
const tabName = n => process.env.TABNAMES ? ({ [PREV_TAB]: "previous month export", [CUR_TAB]: "this month export" }[n] || n) : n;
const wb = { sheets: Object.keys(data).filter(n => !drop(n)).map(n => wrap(new Sheet(tabName(n), pad(data[n])))),
  tableNames: new Set(),
  getWorksheets() { return this.sheets; }, getWorksheet(n) { return this.sheets.find(s => s.name === n); },
  getApplication() { return wrap({ getCalculationMode: () => "automatic", setCalculationMode() {} }); },
  getTable(n) {
    if (!this.tableNames.has(String(n).toLowerCase())) return undefined;
    const ws = this.sheets.find(s => s.tableRanges.some(t => t.name.toLowerCase() === String(n).toLowerCase()));
    const t = ws.tableRanges.find(x => x.name.toLowerCase() === String(n).toLowerCase());
    return wrap({ getName: () => n, getRowCount: () => t.h - 1, getHeaderRowRange: () => ws.getRangeByIndexes(t.r, t.c, 1, t.w),
      getRangeBetweenHeaderAndTotal: () => ws.getRangeByIndexes(t.r + 1, t.c, t.h - 1, t.w) });
  },
  addWorksheet(n) { if (this.getWorksheet(n)) throw new Error("dup sheet " + n); const s = wrap(new Sheet(n)); this.sheets.push(s); return s; } };
// REFJSON=1: compact JSON like reference_extract_officescript.ts produces
const usersJson = process.env.REFJSON ? JSON.stringify(data["Users (Entra)"].slice(1).map(r => [r[1], "", r[2] === "False" ? "0" : "1", r[3]])) : undefined;
const devicesJson = process.env.REFJSON ? JSON.stringify(data["Intune Devices"].slice(1).map(r => [r[0], r[1], r[2].slice(0, 10), r[3], ""])) : undefined;
const poJson = process.env.POJSON ? JSON.stringify(data[PO_TAB].slice(1).map(r => Object.fromEntries(data[PO_TAB][0].map((h, i) => [h, r[i]])))) : undefined;
const load = () => new Function("ExcelScript", js + "\nreturn { main };")(ExcelScript).main;
// SNAPSHOT=1: an earlier run on the previous export alone returns the snapshot this run compares with
let snapshot;
if (process.env.SNAPSHOT) {
  const pwb = { sheets: [wrap(new Sheet(PREV_TAB, pad(data[PREV_TAB])))], tableNames: new Set(), getWorksheets() { return this.sheets; },
    getApplication() { return wrap({ getCalculationMode: () => "automatic", setCalculationMode() {} }); }, getTable() { return undefined; },
    getWorksheet(n) { return this.sheets.find(s => s.name === n); }, addWorksheet(n) { const s = wrap(new Sheet(n)); this.sheets.push(s); return s; } };
  snapshot = load()(pwb, "2026-09-08", undefined, undefined, undefined, undefined, false, undefined, undefined, undefined, undefined, undefined, true).snapshot;
  console.log("snapshot chars", snapshot.length);
  wb.tableNames.clear();  // the earlier run used its own workbook; its table names are not in this one
}
const run = () => load()(wb, process.env.AUDIT || "2026-10-08", undefined, histPath ? fs.readFileSync(histPath, "utf8") : undefined,
  process.env.MARKET, "leavers, stock", process.env.SHARE === "1", process.env.MAPPINGS, usersJson, devicesJson, poJson, snapshot, false);
apiByName.clear();
let t0 = Date.now();
const res = run();
const apiTotal = Array.from(apiByName.values()).reduce((a, b) => a + b, 0);
console.log("excel calls", apiTotal, "| most used:", Array.from(apiByName.entries()).sort((a, b) => b[1] - a[1]).slice(0, Number(process.env.TOPCALLS || 8)).map(e => e.join(" ")).join(", "));
if (process.env.MAXCALLS && apiTotal > Number(process.env.MAXCALLS)) throw new Error(`Too many Excel calls: ${apiTotal} > ${process.env.MAXCALLS} (an Excel call inside a per-row loop?)`);
console.log("run ms", Date.now() - t0, "| api calls", calls, "| sheets", wb.sheets.map(s => s.name + ":" + (s.values.length ? "src" : s.cells.size)).join(", "));
const { agentInput, historyRows, runInfo, reconciliation, qualityChecks, snapshot: _s, ...head } = res;
console.log(head);
console.log(runInfo.join("\n"));
console.log("--- reconciliation\n" + reconciliation.join("\n") + "\n--- quality checks not passed\n" + qualityChecks.join("\n"));
console.log("agentInput chars", agentInput.length, "| historyRows", historyRows.length);
fs.writeFileSync("agent_input.md", agentInput);
if (process.env.HISTOUT) fs.writeFileSync(process.env.HISTOUT, JSON.stringify(historyRows));
if (process.env.BRIEFOUT) fs.writeFileSync(process.env.BRIEFOUT, agentInput.replace(/^.*(2026\.10\.14\.[0-9]+|Run time).*$/gm, ""));
if (process.env.REPORTOUT) fs.writeFileSync(process.env.REPORTOUT, [res.emailSubject, "", res.emailBody, "", "--- PowerPoint prompt (" + res.pptPrompt.replace(/\n/g, " ").length + " characters)", res.pptPrompt].join("\n"));
if (process.env.MARK) {
  const pc = wb.getWorksheet("Physical Check");
  if (!pc) console.log("MARK skipped: no Physical Check tab (REPORT_ONLY mode)");
  for (let r = 4; pc && pc.cells.get(`${r},0`) !== undefined; r++) {
    pc.cells.set(`${r},10`, r % 10 === 0 ? "N" : "Y"); pc.cells.set(`${r},11`, r % 7 === 0 ? "N" : "Y"); pc.cells.set(`${r},12`, "Y");
  }
}
// POFILL=1: complete the to-do rows on POs Validated as a user would, then the re-run must read them as PO data
let poFilled = 0;
if (process.env.POFILL) {
  const ps = wb.getWorksheet("POs Validated"), t = ps.tableRanges.find(x => x.name === "tblPOValidated");
  for (let i = 1; i < t.h; i++) {
    const at = c => `${t.r + i},${t.c + c}`;
    if (ps.cells.get(at(1)) && !ps.cells.get(at(3))) {
      ps.cells.set(at(0), 46300); ps.cells.set(at(3), 5); ps.cells.set(at(4), "USER-0009"); ps.cells.set(at(5), "Yes"); ps.cells.set(at(6), "Yes"); poFilled++;
    }
    // an edit on POs Validated (Pending -> Yes) must win over the PO tab it came from
    [5, 6].forEach(c => { if (ps.cells.get(at(c)) === "Pending") ps.cells.set(at(c), "Yes"); });
  }
  console.log("POFILL: completed", poFilled, "to-do row(s)");
}
if (process.env.SHARE !== "1" && !process.env.NORERUN) { // re-running on the same workbook must work (our sheets are skipped)
  const res2 = run();
  console.log("RERUN ok | assets", res2.assets, "| score", res2.integrityScore);
  assert.equal(res2.assets, res.assets, "re-run reads the same assets");
  const poCount = x => +((x.reconciliation.find(l => l.startsWith("PO validation records processed")) || "").match(/: (\d+)/) || [0, 0])[1];
  console.log("PO rows: first run", poCount(res), "| re-run", poCount(res2));
  if (!process.env.POJSON) assert.equal(poCount(res2), poCount(res) + poFilled, "re-run reads the user's PO rows (and the completed to-do rows) once");
  if (process.env.POFILL) assert.ok(!res2.reconciliation.some(l => /Pending/.test(l)), "edits on POs Validated win over the PO tab");
  if (process.env.MARK) console.log("physical:", res2.runInfo.length, (res2.agentInput.match(/## 11\. Physical check\n(.*)/) || [])[1]);
}
const dump = (name, maxR, maxLen = 70) => { const w = wb.getWorksheet(name); console.log("=== " + name); if (!w) return;
  for (let r = 0; r < maxR; r++) { const out = []; for (let c = 0; c < 18; c++) { const v = w.cells.get(`${r},${c}`); if (v !== undefined && v !== "") out.push(String(v).slice(0, maxLen)); } if (out.length) console.log(r, out.join(" | ")); } };
(process.env.DUMP || "").split(",").filter(x => x).forEach(x => { const [n, k] = x.split(":"); dump(n, +k || 25); });

// RENDER=<dir>: an HTML picture of each result tab (formatting, links, tables, PivotTables); screenshot.mjs turns them into PNGs
if (process.env.RENDER) {
  const esc = v => String(v).replace(/&/g, "&amp;").replace(/</g, "&lt;");
  const serialDate = (v, withTime) => { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v * 86400000));
    const m = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][d.getUTCMonth()];
    return `${String(d.getUTCDate()).padStart(2, "0")}-${m}-${d.getUTCFullYear()}` + (withTime ? ` ${String(d.getUTCHours()).padStart(2, "0")}:${String(d.getUTCMinutes()).padStart(2, "0")}` : ""); };
  const show = (v, nf) => {
    if (typeof v !== "number" || !nf || nf === "@") return { t: v, color: "" };
    if (/yyyy/.test(nf)) return { t: serialDate(v, /hh/.test(nf)), color: "" };
    if (/%/.test(nf)) return { t: (v * 100).toFixed(1) + "%", color: "" };
    const dec = /0\.0/.test(nf) ? 1 : 0, abs = Math.abs(v).toLocaleString("en-US", { minimumFractionDigits: dec, maximumFractionDigits: dec });
    if (/;/.test(nf)) { const pos = nf.split(";")[0]; const posRed = /^\[Red\]/.test(pos);
      if (v > 0) return { t: "+" + abs, color: posRed ? "#C00000" : "#00843D" }; if (v < 0) return { t: "-" + abs, color: posRed ? "#00843D" : "#C00000" }; return { t: "0", color: "" }; }
    return { t: (v < 0 ? "-" : "") + abs, color: "" };
  };
  // A PivotTable as Excel shows it (compact layout): nested row labels with counts or sums, then Grand Total
  const pivotCells = (ws, pv) => {
    const src = wb.sheets.map(x => [x, x.tableRanges.find(t => t.name === pv.table)]).find(x => x[1]);
    const out = new Map(); if (!src) return out;
    const [tws, t] = src, head = []; for (let j = 0; j < t.w; j++) head.push(tws.cells.get(`${t.r},${t.c + j}`));
    const idx = h => head.indexOf(h), rows = [];
    for (let i = 1; i < t.h; i++) { const r = []; for (let j = 0; j < t.w; j++) r.push(tws.cells.get(`${t.r + i},${t.c + j}`)); rows.push(r); }
    const val = r => pv.sum ? (Number(r[idx(pv.value)]) || 0) : (r[idx(pv.value)] !== "" && r[idx(pv.value)] !== undefined ? 1 : 0);
    let line = pv.r;
    const put = (c, v, st) => { out.set(`${line},${c}`, { v, st }); };
    put(pv.c, "Row Labels", "ph"); put(pv.c + 1, (pv.sum ? "Sum of " : "Count of ") + pv.value, "ph"); line++;
    const walk = (list, level) => {
      const groups = new Map(); list.forEach(r => { const k = String(r[idx(pv.rows[level])] || "(blank)"); if (!groups.has(k)) groups.set(k, []); groups.get(k).push(r); });
      [...groups.keys()].sort().forEach(k => { put(pv.c, " ".repeat(4 * level) + k, level < pv.rows.length - 1 ? "pg" : ""); put(pv.c + 1, groups.get(k).reduce((a, r) => a + val(r), 0), level < pv.rows.length - 1 ? "pg" : ""); line++;
        if (level < pv.rows.length - 1) walk(groups.get(k), level + 1); });
    };
    walk(rows, 0); put(pv.c, "Grand Total", "ph"); put(pv.c + 1, rows.reduce((a, r) => a + val(r), 0), "ph");
    return out;
  };
  fs.mkdirSync(process.env.RENDER, { recursive: true });
  wb.sheets.filter(ws => !ws.values.length).forEach(ws => {
    const piv = new Map(); ws.pivots.forEach(pv => pivotCells(ws, pv).forEach((x, k) => piv.set(k, x)));
    let H = 0, W = 0; for (const k of [...ws.cells.keys(), ...piv.keys()]) { const [r, c] = k.split(",").map(Number); H = Math.max(H, r + 1); W = Math.max(W, c + 1); }
    H = Math.min(H, +(process.env.RENDER_ROWS || 70));
    const inTable = (r, c) => ws.tableRanges.find(t => r >= t.r && r < t.r + t.h && c >= t.c && c < t.c + t.w);
    let html = `<html><head><meta charset="utf-8"><style>body{margin:0;font:11pt Calibri,Arial,sans-serif;background:#fff}table{border-collapse:collapse;table-layout:fixed}
td{height:20px;padding:1px 4px;white-space:nowrap;overflow:visible;vertical-align:bottom;font-size:11pt}td.rh{background:#f3f3f3;color:#666;text-align:center;font-size:9pt;border:1px solid #ddd}
th{background:#f3f3f3;color:#666;font-weight:normal;font-size:9pt;border:1px solid #ddd}.tb{border:1px solid #BFBFBF}.ph{background:#DDEBF7;font-weight:bold;border-top:1px solid #9BC2E6}.pg{font-weight:bold}</style></head><body>
<div style="padding:4px 8px;background:#217346;color:#fff;font:bold 12px Arial">${esc(ws.name)}</div><table style="width:${32 + Array.from({ length: W }, (_, c) => Math.round((ws.colW[c] || 48) * 1.33)).reduce((x, y) => x + y, 0)}px"><tr><th style="width:32px"></th>`;
    for (let c = 0; c < W; c++) html += `<th style="width:${Math.round((ws.colW[c] || 48) * 1.33)}px">${String.fromCharCode(65 + (c % 26))}</th>`;
    html += "</tr>";
    for (let r = 0; r < H; r++) {
      html += `<tr><td class="rh">${r + 1}</td>`;
      for (let c = 0; c < W; c++) {
        const k = `${r},${c}`, st = ws.fmt.get(k) || {}, p = piv.get(k), t = inTable(r, c);
        let v = ws.cells.get(k); if (v === undefined) v = ""; if (typeof v === "string" && v.startsWith("'")) v = v.slice(1);
        const shown = p ? { t: typeof p.v === "number" ? p.v.toLocaleString("en-US") : p.v, color: "" } : show(v, st.nf);
        const css = [];
        if (t) { css.push("border:1px solid #BFBFBF"); if (r === t.r) css.push("font-weight:bold"); else if ((r - t.r) % 2 === 0) css.push("background:#F2F2F2"); }
        if (st.fill) css.push(`background:${st.fill}`);
        if (st.bold) css.push("font-weight:bold"); if (st.italic) css.push("font-style:italic");
        if (st.color) css.push(`color:${st.color}`); if (shown.color) css.push(`color:${shown.color}`);
        if (st.size) css.push(`font-size:${st.size}pt`); if (st.underline || st.link) css.push("text-decoration:underline");
        if (st.wrap && t && r !== t.r) css.push("white-space:normal");
        if (st.wrap && t && r === t.r) css.push("white-space:normal");
        if (typeof v === "number" || p && typeof p.v === "number") css.push("text-align:right");
        const cls = p && p.st ? ` class="${p.st}"` : "";
        html += `<td${cls} style="${css.join(";")}"${st.formula ? ` title="${esc(st.formula)}"` : ""}>${esc(shown.t)}${st.formula ? '<sup style="color:#999;font-size:7pt">ƒ</sup>' : ""}</td>`;
      }
      html += "</tr>";
    }
    fs.writeFileSync(`${process.env.RENDER}/${ws.name.replace(/[^A-Za-z0-9]+/g, "_")}.html`, html + "</table></body></html>");
  });
  console.log("rendered to", process.env.RENDER);
}
if (process.env.TABLES) wb.sheets.forEach(s => s.tableRanges.forEach(t => console.log("table", s.name, t.name, `rows ${t.r + 1}-${t.r + t.h}`, `cols ${t.c}-${t.c + t.w - 1}`)));
