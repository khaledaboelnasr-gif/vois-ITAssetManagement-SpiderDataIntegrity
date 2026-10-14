/**
 * VOIS Asset Audit - Reference extract (Office Script)
 *
 * Runs on a reference workbook kept in SharePoint (a Users export from HR / Entra ID, or a device export from
 * Intune / SCCM) and returns it as compact JSON for the audit script's usersJson / devicesJson parameters.
 * The flow cannot copy tabs between workbooks with standard connectors, so the reference travels as text.
 *
 * Parameters:
 *   kind   "users" or "devices"
 *   market optional: devices only - keep devices whose company / country column maps to this text (contains)
 *
 * Output (only the columns the audit needs, so it stays far below the ~5 MB limit):
 *   users    [[email, employeeId, active 1/0, leaveDate YYYY-MM-DD], ...]
 *   devices  [[deviceName, serial, lastSeen YYYY-MM-DD, primaryUser, company], ...]
 *
 * Column names are matched ignoring case, spaces and punctuation; keep the lists in step with USER_COLS and
 * DEVICE_COLS in asset_audit_officescript.ts.
 */
function main(workbook: ExcelScript.Workbook, kind: string, market?: string): ExtractResult {
  const k = (kind || "").trim().toLowerCase();
  if (k !== "users" && k !== "devices") { throw new Error(`Unknown kind '${kind}'. Use "users" or "devices".`); }
  const cols = k === "users" ? USER_COLS : DEVICE_COLS;
  const need = k === "users" ? ["email", "empId"] : ["name", "serial"];
  let best: { ws: ExcelScript.Worksheet | null; row: number; pos: Record<string, number>; n: number } = { ws: null, row: -1, pos: {}, n: 0 };
  workbook.getWorksheets().forEach((ws: ExcelScript.Worksheet) => {
    const used = ws.getUsedRange(true);
    if (!used) { return; }
    const head = ws.getRangeByIndexes(used.getRowIndex(), used.getColumnIndex(), Math.min(20, used.getRowCount()),
      used.getColumnCount()).getValues() as Raw[][];
    head.forEach((r, i) => {
      const pos = mapColumns(r.map(v => headerKey(v)), cols), n = Object.keys(pos).length;
      if (n > best.n && need.some(f => pos[f] !== undefined)) { best = { ws, row: i, pos, n }; }
    });
  });
  if (!best.ws) {
    throw new Error(`No ${k} sheet found. Expected a header row with ` + need.map(f => cols[f].slice(0, 3).join(" / ")).join(" or ") + ".");
  }
  const ws = best.ws as ExcelScript.Worksheet, pos = best.pos;
  const all = readValues(ws), header = all[best.row], values = all.slice(best.row + 1);
  const get = (r: Raw[], f: string) => pos[f] === undefined ? "" : String(r[pos[f]] === undefined || r[pos[f]] === null ? "" : r[pos[f]]).trim();
  const date = (r: Raw[], f: string) => pos[f] === undefined ? "" : dayText(parseDate(r[pos[f]]));
  let rows: string[][];
  if (k === "users") {
    rows = values.filter(r => get(r, "email") || get(r, "empId")).map(r => {
      const act = activeValue(get(r, "enabled"));
      return [get(r, "email").toLowerCase(), get(r, "empId"), act === false ? "0" : "1", date(r, "leaveDate")];
    });
  } else {
    rows = values.filter(r => get(r, "name") || get(r, "serial"))
      .filter(r => !market || get(r, "company").toLowerCase().indexOf(market.toLowerCase()) >= 0 || !get(r, "company"))
      .map(r => [get(r, "name"), get(r, "serial"), date(r, "lastSeen"), get(r, "user").toLowerCase(), get(r, "company")]);
  }
  const json = JSON.stringify(rows);
  return {
    kind: k, sheet: ws.getName(), rows: rows.length, chars: json.length,
    columns: Object.keys(pos).map(f => `${f} = '${String(header[pos[f]])}'`),
    json
  };
}

type Raw = string | number | boolean;
interface ExtractResult { kind: string; sheet: string; rows: number; chars: number; columns: string[]; json: string; }

const USER_COLS: Record<string, string[]> = {
  email: ["Email", "Mail", "User Principal Name", "UserPrincipalName", "UPN", "Email Address", "Work Email", "Business Email"],
  empId: ["Employee ID", "EmployeeId", "Employee Number", "Staff ID", "Personnel Number", "Emp ID"],
  name: ["Display Name", "Full Name", "Employee Name", "Name"],
  enabled: ["Account Enabled", "AccountEnabled", "Enabled", "Active", "Employee Status", "Employment Status", "User Status", "Status"],
  leaveDate: ["Leaving Date", "Termination Date", "Last Working Day", "Exit Date", "Leave Date", "End Date", "Employee Leave Date Time", "Contract End Date"]
};
const DEVICE_COLS: Record<string, string[]> = {
  name: ["Device Name", "DeviceName", "Computer Name", "Hostname", "Host Name", "Machine Name", "Netbios Name", "Name"],
  serial: ["Serial Number", "SerialNumber", "Serial", "Serial No"],
  lastSeen: ["Last Check-in", "Last check in", "Last Sync", "Last Sync Date Time", "LastSyncDateTime", "Last Contact",
    "Last Seen", "Last Logon", "Last Logon Date", "Last Logon Timestamp", "Last Activity", "Last Active Time", "Last Seen Time"],
  user: ["Primary User UPN", "Primary user email address", "User Principal Name", "UserPrincipalName", "Primary User",
    "User Email", "Email", "Last Logged On User"],
  company: ["Company", "Country", "Market", "Legal Entity"]
};

function headerKey(v: Raw): string {
  return String(v === undefined || v === null ? "" : v).toLowerCase().replace(/[^a-z0-9]/g, "");
}
function mapColumns(header: string[], cols: Record<string, string[]>): Record<string, number> {
  const pos: Record<string, number> = {};
  const taken = new Set<number>();
  Object.keys(cols).forEach(f => {
    for (const name of cols[f]) {
      const i = header.indexOf(headerKey(name));
      if (i >= 0 && !taken.has(i)) { pos[f] = i; taken.add(i); break; }
    }
  });
  return pos;
}
function readValues(ws: ExcelScript.Worksheet): Raw[][] {
  const used = ws.getUsedRange(true);
  if (!used) { return []; }
  const rows = used.getRowCount(), cols = used.getColumnCount();
  const r0 = used.getRowIndex(), c0 = used.getColumnIndex();
  const CHUNK = Math.max(200, Math.floor(400000 / Math.max(1, cols)));
  let out: Raw[][] = [];
  for (let i = 0; i < rows; i += CHUNK) {
    out = out.concat(ws.getRangeByIndexes(r0 + i, c0, Math.min(CHUNK, rows - i), cols).getValues() as Raw[][]);
  }
  return out;
}
function activeValue(v: string): boolean | null {
  const s = v.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
  if (!s) { return null; }
  if (/^(true|yes|y|1|active|enabled|employed|current|on)$/.test(s)) { return true; }
  if (/^(false|no|n|0|inactive|disabled|terminated|left|leaver|withdrawn|resigned|exited|separated|off)$/.test(s)) { return false; }
  return null;
}
function dayText(d: Date | null): string {
  return d ? `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}` : "";
}
// Same rules as the audit script: Excel serials, ISO, day-first d/m/y, and "14-Oct-2026"
function parseDate(v: Raw): Date | null {
  if (typeof v === "number") {
    return v > 1000 && v < 80000 ? new Date(Date.UTC(1899, 11, 30) + Math.round(v * 86400000)) : null;
  }
  const s = String(v || "").trim();
  if (!s) { return null; }
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) { return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])); }
  m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/);
  if (m) {
    const a = +m[1], b = +m[2];
    const dayFirst = a > 12 || s.indexOf(".") > 0 || (b <= 12 && a <= 12);
    const month = dayFirst ? b : a, day = dayFirst ? a : b;
    return month >= 1 && month <= 12 ? new Date(Date.UTC(+m[3], month - 1, day)) : null;
  }
  m = s.match(/^(\d{1,2})[ \-]([A-Za-z]{3})[A-Za-z]*[ \-,]+(\d{4})/);
  if (m) {
    const mi = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"].indexOf(m[2].toLowerCase());
    return mi >= 0 ? new Date(Date.UTC(+m[3], mi, +m[1])) : null;
  }
  return null;
}
