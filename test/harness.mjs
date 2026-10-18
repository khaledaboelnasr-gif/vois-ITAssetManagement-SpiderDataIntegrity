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
  getAddress() { return `${this.ws.name}!R${this.r}C${this.c}`; }
  setNumberFormat(f) { if (typeof f !== "string") throw new Error("setNumberFormat needs string"); }
  getFormat() { return any; } addConditionalFormat() { return any; }
}
class Sheet {
  constructor(name, values) { this.name = name; this.values = values || []; this.cells = new Map(); this.tables = []; }
  getName() { return this.name; }
  size() { let h = 0, w = 0; for (const k of this.cells.keys()) { const [r, c] = k.split(",").map(Number); h = Math.max(h, r + 1); w = Math.max(w, c + 1); } return [h, w]; }
  grid() { const [h, w] = this.size(); return Array.from({ length: h }, (_, r) => Array.from({ length: w }, (_, c) => { const v = this.cells.get(`${r},${c}`); return v === undefined ? "" : v; })); }
  getUsedRange() { if (this.values.length) return wrap(new Range(this, 0, 0, this.values.length, this.values[0].length)); const [h, w] = this.size(); return h ? wrap(new Range(this, 0, 0, h, w)) : undefined; }
  getRangeByIndexes(r, c, h, w) { if (h < 1 || w < 1) throw new Error(`empty range ${h}x${w} on ${this.name}`); return wrap(new Range(this, r, c, h, w)); }
  getRange(a) { const m = a.match(/^([A-Z]+)(\d+)$/); return wrap(new Range(this, +m[2] - 1, colIdx(m[1]), 1, 1)); }
  addTable(rng) { const self = this; return wrap({ setName(n) { if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(n) || n.length > 255) throw new Error("bad table name " + n); if (wb.tableNames.has(n.toLowerCase())) throw new Error("dup table " + n); wb.tableNames.add(n.toLowerCase()); self.tables.push(n.toLowerCase()); }, setPredefinedTableStyle() {} }); }
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
  getTable(n) { return this.tableNames.has(String(n).toLowerCase()) ? wrap({ getName: () => n }) : undefined; },
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
if (process.env.SHARE !== "1" && !process.env.NORERUN) { // re-running on the same workbook must work (our sheets are skipped)
  const res2 = run();
  console.log("RERUN ok | assets", res2.assets, "| score", res2.integrityScore);
  assert.equal(res2.assets, res.assets, "re-run reads the same assets");
  if (process.env.MARK) console.log("physical:", res2.runInfo.length, (res2.agentInput.match(/## 11\. Physical check\n(.*)/) || [])[1]);
}
const dump = (name, maxR, maxLen = 70) => { const w = wb.getWorksheet(name); console.log("=== " + name); if (!w) return;
  for (let r = 0; r < maxR; r++) { const out = []; for (let c = 0; c < 18; c++) { const v = w.cells.get(`${r},${c}`); if (v !== undefined && v !== "") out.push(String(v).slice(0, maxLen)); } if (out.length) console.log(r, out.join(" | ")); } };
(process.env.DUMP || "").split(",").filter(x => x).forEach(x => { const [n, k] = x.split(":"); dump(n, +k || 25); });
