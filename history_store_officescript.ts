/**
 * VOIS Asset Audit - History store (Office Script for SharePoint "AuditHistory.xlsx")
 *
 * One script, two modes, so the flow needs two "Run script" actions instead of hundreds of row actions.
 *
 *   mode = "read"  returns historyJson (for the audit script's historyJson) and mappingsJson (central Mappings table)
 *   mode = "save"  adds this run's historyRows and one line to the Runs log. Nothing is ever overwritten:
 *                  every run keeps its own rows, tagged with its RunId.
 *
 * When several runs cover the same period + market (two people, or a re-run after fixes), "read" uses ONE of them:
 * the newest run that is complete, i.e. it has at least 90% of the asset records of the biggest run for that
 * period + market. So a partial export never replaces a full audit, and nothing is counted twice.
 *
 * Parameters:
 *   mode       "read" or "save"
 *   rowsJson   save: JSON array of {period, market, assetClass, metric, value} (the audit script's historyRows)
 *   runId      save: unique id of this flow run (also in the result file name)
 *   months     read: only return the latest N periods (default 13 = this month + a year back); 0 = all
 *   market     read: only return this market; empty = all markets
 *   requester  save: who ran it (email), for the Runs log
 *   fileLink   save: link to the result workbook, for the Runs log
 *
 * Tables (created on first run if missing):
 *   History   Period | Market | AssetClass | Metric | Value | UpdatedOn | RunId
 *   Runs      RunId | UpdatedOn | Requester | Markets | Periods | Assets | IntegrityScore | File
 *   Mappings  Kind | Value | MapsTo | Note   (optional, created by hand; empty = built-in rules only)
 */
function main(workbook: ExcelScript.Workbook, mode: string, rowsJson?: string, runId?: string,
  months?: number, market?: string, requester?: string, fileLink?: string): StoreResult {
  const history = ensureTable(workbook, HISTORY_TABLE, HEADERS);
  const m = (mode || "").trim().toLowerCase();
  if (m === "read") {
    return readHistory(workbook, history, months === undefined || months === null ? 13 : months, market || "");
  }
  if (m === "save") {
    if (!rowsJson) { throw new Error("save needs rowsJson (the audit script's historyRows)."); }
    if (!runId) { throw new Error("save needs runId, so every run's rows stay separate."); }
    const incoming: HistoryRow[] = JSON.parse(rowsJson);
    return save(workbook, history, incoming, runId, requester || "", fileLink || "");
  }
  throw new Error(`Unknown mode '${mode}'. Use "read" or "save".`);
}

// ============================================================================ settings
const HISTORY_TABLE = "History";
const RUNS_TABLE = "Runs";
const MAPPINGS_TABLE = "Mappings";
const HEADERS = ["Period", "Market", "AssetClass", "Metric", "Value", "UpdatedOn", "RunId"];
const RUNS_HEADERS = ["RunId", "UpdatedOn", "Requester", "Markets", "Periods", "Assets", "IntegrityScore", "File"];
const RECORDS_METRIC = "#records";
const CHUNK = 4000;          // rows per read / write call (stays well under the ~5 MB request limit)
const COMPLETE_SHARE = 0.9;  // a run is "complete" for a period + market with at least 90% of the biggest run's records
const STORE_VERSION = "2026.10.14";

interface HistoryRow { period: string; market: string; assetClass: string; metric: string; value: number; }
interface StoredRow extends HistoryRow { updated: string; runId: string; }
interface StoreResult {
  mode: string; historyJson: string; mappingsJson: string; rows: number; periods: string[];
  used: string[]; added: number; version: string;
}
type Cell = string | number | boolean;

// ============================================================================ tables
function ensureTable(wb: ExcelScript.Workbook, name: string, headers: string[]): ExcelScript.Table {
  const t = wb.getTable(name);
  if (t) { return t; }
  let ws = wb.getWorksheet(name);
  if (!ws) { ws = wb.addWorksheet(name); }
  ws.getRangeByIndexes(0, 0, 1, headers.length).setValues([headers]);
  const nt = wb.addTable(ws.getRangeByIndexes(0, 0, 1, headers.length), true);
  nt.setName(name);
  return nt;
}

function bodyValues(t: ExcelScript.Table): Cell[][] {
  const n = t.getRowCount();
  if (n === 0) { return []; }
  const body = t.getRangeBetweenHeaderAndTotal();
  const ws = t.getWorksheet();
  const top = body.getRowIndex(), left = body.getColumnIndex(), w = body.getColumnCount();
  let out: Cell[][] = [];
  for (let r = 0; r < n; r += CHUNK) {
    out = out.concat(ws.getRangeByIndexes(top + r, left, Math.min(CHUNK, n - r), w).getValues() as Cell[][]);
  }
  // An empty table still has one blank row
  return out.filter(row => row.some(c => String(c).trim() !== ""));
}

function columns(t: ExcelScript.Table, expected: string[]): Map<string, number> {
  const h: Cell[] = t.getHeaderRowRange().getValues()[0];
  const idx = new Map<string, number>();
  h.forEach((c, i) => idx.set(String(c).trim().toLowerCase(), i));
  const missing = expected.filter(e => !idx.has(e.toLowerCase()));
  if (missing.length) { throw new Error(`Table '${t.getName()}' is missing columns: ${missing.join(", ")}`); }
  return idx;
}

// Adds rows under the existing ones (never rewrites them), then fits the table to the new size
function appendRows(t: ExcelScript.Table, rows: Cell[][], textCols: number[]) {
  if (!rows.length) { return; }
  const ws = t.getWorksheet();
  const head = t.getHeaderRowRange();
  const top = head.getRowIndex() + 1, left = head.getColumnIndex(), w = head.getColumnCount();
  const n = t.getRowCount();
  // Excel's placeholder row in an empty table is reused
  const start = n === 1 && bodyValues(t).length === 0 ? 0 : n;
  // Text format first, so Excel does not turn "2026-09" into a date
  textCols.forEach(c => ws.getRangeByIndexes(top + start, left + c, rows.length, 1).setNumberFormat("@"));
  for (let r = 0; r < rows.length; r += CHUNK) {
    const part = rows.slice(r, r + CHUNK).map(row => row.slice(0, w));
    ws.getRangeByIndexes(top + start + r, left, part.length, w).setValues(part);
  }
  t.resize(ws.getRangeByIndexes(top - 1, left, start + rows.length + 1, w));
}

// ============================================================================ read
function readStored(t: ExcelScript.Table): StoredRow[] {
  const idx = columns(t, HEADERS);
  const ci = (h: string) => idx.get(h.toLowerCase()) as number;
  return bodyValues(t).filter(r => String(r[ci("Period")]).trim() !== "").map(r => ({
    period: periodText(r[ci("Period")]), market: String(r[ci("Market")]), assetClass: String(r[ci("AssetClass")]),
    metric: String(r[ci("Metric")]), value: Number(r[ci("Value")]) || 0,
    updated: String(r[ci("UpdatedOn")]), runId: String(r[ci("RunId")])
  }));
}

function readHistory(wb: ExcelScript.Workbook, t: ExcelScript.Table, months: number, market: string): StoreResult {
  let rows = readStored(t);
  if (market) { rows = rows.filter(r => r.market.toLowerCase() === market.toLowerCase()); }
  const chosen = pickRuns(rows);
  rows = rows.filter(r => chosen.get(r.period + " | " + r.market) === r.runId);
  let periods = Array.from(new Set<string>(rows.map(r => r.period))).sort();
  if (months > 0 && periods.length > months) {
    periods = periods.slice(periods.length - months);
    const keep = new Set<string>(periods);
    rows = rows.filter(r => keep.has(r.period));
  }
  const out: HistoryRow[] = rows.map(r => ({ period: r.period, market: r.market, assetClass: r.assetClass, metric: r.metric, value: r.value }));
  const used: string[] = [];
  chosen.forEach((run, key) => used.push(`${key}: ${run}`));
  return result("read", { historyJson: JSON.stringify(out), mappingsJson: readMappings(wb), rows: out.length, periods, used: used.sort() });
}

// For each period + market: the newest run with at least 90% of the biggest run's asset records
function pickRuns(rows: StoredRow[]): Map<string, string> {
  const runs = new Map<string, Map<string, { total: number; updated: string }>>();
  rows.forEach(r => {
    const key = r.period + " | " + r.market;
    const byRun = runs.get(key) || new Map<string, { total: number; updated: string }>();
    const x = byRun.get(r.runId) || { total: 0, updated: r.updated };
    if (r.metric === RECORDS_METRIC) { x.total += r.value; }
    if (r.updated > x.updated) { x.updated = r.updated; }
    byRun.set(r.runId, x);
    runs.set(key, byRun);
  });
  const chosen = new Map<string, string>();
  runs.forEach((byRun, key) => {
    const list = Array.from(byRun.keys()).map(id => ({ id, total: (byRun.get(id) as { total: number }).total,
      updated: (byRun.get(id) as { updated: string }).updated }));
    const max = list.reduce((a, x) => Math.max(a, x.total), 0);
    const complete = list.filter(x => x.total >= max * COMPLETE_SHARE)
      .sort((a, b) => b.updated.localeCompare(a.updated) || b.id.localeCompare(a.id));
    chosen.set(key, complete[0].id);
  });
  return chosen;
}

function readMappings(wb: ExcelScript.Workbook): string {
  const t = wb.getTable(MAPPINGS_TABLE);
  if (!t || t.getRowCount() === 0) { return ""; }
  const head: Cell[] = t.getHeaderRowRange().getValues()[0];
  const h = head.map(c => String(c).trim());
  const rows = bodyValues(t);
  if (!rows.length) { return ""; }
  return JSON.stringify(rows.map(r => {
    const o: Record<string, string | number> = {};
    h.forEach((name, i) => { o[name] = typeof r[i] === "number" ? (r[i] as number) : String(r[i]); });
    return o;
  }));
}

// Excel may turn "2026-09" into a date serial; bring it back to YYYY-MM
function periodText(v: Cell): string {
  if (typeof v === "number" && v > 20000) {
    const d = new Date(Math.round((v - 25569) * 86400000));
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
  }
  return String(v).trim().replace(/^'/, "");
}

// ============================================================================ save
function save(wb: ExcelScript.Workbook, t: ExcelScript.Table, incoming: HistoryRow[], runId: string,
  requester: string, fileLink: string): StoreResult {
  const idx = columns(t, HEADERS);
  const ci = (h: string) => idx.get(h.toLowerCase()) as number;
  const width = t.getHeaderRowRange().getColumnCount();
  // A retried flow run must not add its rows twice
  if (readStored(t).some(r => r.runId === runId)) {
    return result("save", { rows: t.getRowCount(), used: [`Run ${runId} was already saved - nothing added`] });
  }
  const stamp = new Date().toISOString().slice(0, 19).replace("T", " ") + " UTC";
  const rows: Cell[][] = incoming.map(r => {
    const row: Cell[] = new Array(width).fill("");
    row[ci("Period")] = String(r.period);
    row[ci("Market")] = String(r.market);
    row[ci("AssetClass")] = String(r.assetClass || "");
    row[ci("Metric")] = String(r.metric);
    row[ci("Value")] = Number(r.value) || 0;
    row[ci("UpdatedOn")] = stamp;
    row[ci("RunId")] = runId;
    return row;
  });
  appendRows(t, rows, [ci("Period")]);

  // One line per run: who, when, what, the score, and where the result is
  const log = ensureTable(wb, RUNS_TABLE, RUNS_HEADERS);
  const li = columns(log, RUNS_HEADERS);
  const lc = (h: string) => li.get(h.toLowerCase()) as number;
  const line: Cell[] = new Array(log.getHeaderRowRange().getColumnCount()).fill("");
  const periods = Array.from(new Set<string>(incoming.map(r => String(r.period)))).sort();
  const total = (metric: string) => incoming.filter(r => r.metric === metric).reduce((a, r) => a + (Number(r.value) || 0), 0);
  line[lc("RunId")] = runId;
  line[lc("UpdatedOn")] = stamp;
  line[lc("Requester")] = requester;
  line[lc("Markets")] = Array.from(new Set<string>(incoming.map(r => String(r.market)))).sort().join(", ");
  line[lc("Periods")] = periods.join(", ");
  line[lc("Assets")] = total(RECORDS_METRIC);
  line[lc("IntegrityScore")] = total("#inscope") ? (100 * (1 - total("#issues") / total("#inscope"))).toFixed(1) + "%" : "";
  line[lc("File")] = fileLink;
  appendRows(log, [line], [lc("Periods")]);
  return result("save", { rows: t.getRowCount(), periods, added: rows.length });
}

function result(mode: string, part: Partial<StoreResult>): StoreResult {
  const r: StoreResult = {
    mode, historyJson: "", mappingsJson: "", rows: 0, periods: [], used: [], added: 0, version: STORE_VERSION
  };
  return Object.assign(r, part);
}
