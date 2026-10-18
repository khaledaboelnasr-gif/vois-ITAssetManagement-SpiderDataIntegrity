/**
 * VOIS Asset Audit - Office Script (Excel on the web > Automate)
 *
 * Monthly asset reconciliation and data-integrity audit of the VCAMS (Spider) asset export. It:
 *   1. reconciles the previous and the current export on Asset no.: new assets, removed assets, status changes,
 *      the Scrapped / Lost operational list, and the user's PO validation (POs Validated + PO reconciliation);
 *   2. builds the monthly summary by asset type (High value / Low value / Other) with PivotTables and quality checks;
 *   3. audits the current export's data integrity (completeness, validity, duplicates, consistency, timeliness),
 *      optionally against a user directory (leavers) and device management (Intune / SCCM);
 *   4. measures in-stock utilisation (stock age, months of cover, idle stock, moves between sites) and lifecycle;
 *   5. returns a brief (agentInput) for the Copilot Studio agent, and history rows for the month-on-month trend.
 *
 * The workbook it runs on holds (raw tabs are never changed; all output goes to new sheets):
 *   - the current VCAMS export, and for the reconciliation the previous one (any tab with an Asset no. / serial
 *     column and a status or asset type column). Name them e.g. "2026-09 Previous Export" and "2026-10 Current
 *     Export"; without "previous" / "current" in the names, the older one (by the date in the name, else by the
 *     newest update date in the data) is the previous export.
 *   - optional: a PO validation tab (Purchase Date | Purchase Number | Model | Quantity | Raised By |
 *     Physical Security Validation | Spider Validation), supplied by the user - never invented;
 *   - optional: a Users tab (HR / Entra ID) and a Devices tab (Intune / SCCM).
 *
 * Parameters (all optional - used by the Power Automate flow, leave empty when running by hand):
 *   auditDate        "YYYY-MM-DD" the current export represents; default = today (UTC). Its month = history period
 *   previousDate     "YYYY-MM-DD" the previous export represents; default = from the tab name / snapshot / a month before
 *   historyJson      JSON array of {period, market, assetClass, metric, value} from the History table (trend)
 *   marketFilter     keep only this market, e.g. "VOIS EG"; empty = all markets
 *   focus            what the user asked to focus on (echoed to the agent)
 *   shareSafe        true = deterministic tokens for people, asset numbers, serials, hostnames, MACs, POs, function
 *                    units and locations (USER-0001, AST-000001, ...) and remove the raw tabs, for wider sharing
 *   mappingsJson     JSON array of {Kind, Value, MapsTo} from the central Mappings table; empty = 'Mappings' sheet
 *   usersJson        compact users reference from reference_extract (overrides a Users tab)
 *   devicesJson      compact devices reference from reference_extract (overrides a Devices tab)
 *   poJson           JSON array of the user's PO validation rows (overrides a PO validation tab)
 *   previousSnapshot the snapshot text a previous run returned, used when the workbook has only the current export
 *   returnSnapshot   true = return a compact snapshot of the current export (the flow stores it for next month)
 */
function main(workbook: ExcelScript.Workbook, auditDate?: string, previousDate?: string, historyJson?: string,
  marketFilter?: string, focus?: string, shareSafe?: boolean, mappingsJson?: string, usersJson?: string,
  devicesJson?: string, poJson?: string, previousSnapshot?: string, returnSnapshot?: boolean): AuditResult {
  // Excel recalculates every formula after each write; with 200,000 rows that makes the run much slower. Pause it
  // while the script runs and put the user's setting back at the end (also when the run fails).
  const calcMode = workbook.getApplication().getCalculationMode();
  workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
  try {
    return runAudit(workbook, auditDate, previousDate, historyJson, marketFilter, focus, shareSafe, mappingsJson, usersJson, devicesJson, poJson, previousSnapshot, returnSnapshot);
  } finally {
    workbook.getApplication().setCalculationMode(calcMode);
    // Work out the Pivot & Summary formulas now: otherwise the saved file can show 0 until Excel recalculates
    progress("Calculating the summary formulas");
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  }
}
function runAudit(workbook: ExcelScript.Workbook, auditDate?: string, previousDate?: string, historyJson?: string,
  marketFilter?: string, focus?: string, shareSafe?: boolean, mappingsJson?: string, usersJson?: string,
  devicesJson?: string, poJson?: string, previousSnapshot?: string, returnSnapshot?: boolean): AuditResult {
  CLOCK.start = 0;
  clearMemo();
  progress("Started (script " + SCRIPT_VERSION + ")");
  const asOf = (auditDate ? parseDate(auditDate) : null) || utcToday();
  const period = monthOf(asOf);
  const maps = loadMappings(workbook, mappingsJson);
  const tabs = findTabs(workbook, maps);
  if (!tabs.assets.length) {
    throw new Error("No VCAMS asset export found. The export needs a header row with an Asset no., asset tag or serial " +
      "number column and a status or asset type column.");
  }
  const roles = splitExports(workbook, tabs.assets, maps);
  progress(`Reading the current export (${roles.cur.join(", ")})`);
  const cur = loadAssets(workbook, roles.cur, maps);
  progress(`Current export read: ${cur.rows} rows`);
  const prevLoad = roles.prev.length ? loadAssets(workbook, roles.prev, maps) : null;
  if (prevLoad) { progress(`Previous export read (${roles.prev.join(", ")}): ${prevLoad.rows} rows`); }
  const reconciling = !!prevLoad || !!previousSnapshot;
  if (reconciling) {
    [cur].concat(prevLoad ? [prevLoad] : []).forEach((l, i) => {
      const missing = REQUIRED_FIELDS.filter(f => !l.cols.has(f));
      if (missing.length) {
        throw new Error(`The ${i ? "previous" : "current"} export has no ${missing.map(f => STANDARD_HEADER[f]).join(", ")} column ` +
          `(accepted names: ${missing.map(f => ASSET_COLS[f].slice(0, 4).join(" / ")).join("; ")}). Add it, or add a Mappings row ` +
          "Kind 'Column', Value = the standard header, MapsTo = your column name.");
      }
    });
  }
  const snapshot = returnSnapshot ? makeSnapshot(cur.assets, period) : "";
  const marketsInFile = keysByCount(countBy(cur.assets.map(a => a.market)));
  const inMarket = (a: Asset) => !marketFilter || a.market.toLowerCase() === marketFilter.toLowerCase();
  const assets = cur.assets.filter(a => inMarket(a));
  if (!assets.length) {
    throw new Error(marketFilter ? `No assets for market '${marketFilter}'. Markets in the export: ${marketsInFile.join(", ")}`
      : "The asset sheet has a header row but no asset rows.");
  }
  const snap = !prevLoad && previousSnapshot ? parseSnapshot(previousSnapshot, maps) : null;
  const prevAssets = prevLoad ? prevLoad.assets.filter(a => inMarket(a)) : snap ? snap.assets.filter(a => inMarket(a)) : null;
  const prevPeriod = previousDate && parseDate(previousDate) ? monthOf(parseDate(previousDate) as Date) :
    roles.prevPeriod || (snap ? snap.period : "") || monthOf(new Date(Date.UTC(asOf.getUTCFullYear(), asOf.getUTCMonth() - 1, 1)));
  const users = loadUsers(workbook, tabs.users, usersJson, asOf);
  const devices = loadDevices(workbook, tabs.devices, devicesJson, maps);
  const markets = keysByCount(countBy(assets.map(a => a.market)));

  const ctx: Ctx = {
    assets, asOf, period, markets, users, devices, cols: cur.cols, maps, focus: focus || "", shareSafe: !!shareSafe,
    dupGroups: [], manyDevices: [], orphans: [], runInfo: [], history: [], prevOf: new Map<string, string>(),
    physical: { sample: [], marks: new Map<string, PhysicalMark>(), found: 0, missing: 0, locOk: 0, locChecked: 0,
      userOk: 0, userChecked: 0 },
    stockAgeSource: cur.cols.has("statusDate") ? "status date" : cur.cols.has("updated") ? "asset update date (approximate)" :
      cur.cols.has("purchase") ? "purchase date (no status date column)" : "",
    deploySource: "", recon: emptyRecon(), po: loadPo(workbook, tabs.po, poJson), formulas: !shareSafe && !marketFilter,
    curTabs: cur.tabs, prevTabs: prevLoad ? prevLoad.tabs : [], tokens: []
  };
  if (prevAssets) {
    ctx.recon = reconcile(ctx, cur, prevLoad, prevAssets, assets, prevLoad ? "previous export tab " +
      roles.prev.map(n => `'${n}'`).join(", ") : "snapshot from the previous run", prevPeriod);
    // The dates of the two exports, for the monthly email: given, else from the tab names
    ctx.recon.prevDay = (previousDate ? parseDate(previousDate) : null) || (roles.prev.length ? tabDay(roles.prev[0]) : null);
    ctx.recon.curDay = auditDate ? asOf : tabDay(roles.cur[0]);
  }
  progress("Comparison with the previous month done");
  ctx.deploySource = ctx.recon.available ? "status changes into In operation since the previous export (exact)" :
    cur.cols.has("deployDate") ? "deployment date (last 90 days / 3)" : cur.cols.has("statusDate") ?
      "status date of in-use assets (last 90 days / 3)" : "";
  runChecks(ctx);
  progress("Data checks done");
  ctx.orphans = devices.list.filter(d => !d.matched && (!marketFilter || !d.market ||
    d.market.toLowerCase() === marketFilter.toLowerCase()));
  // Masking runs after the checks and the reconciliation, so matching uses the real values
  if (shareSafe) { applyShareSafe(ctx); }
  const current = historyFromAssets(ctx);
  ctx.history = mergeHistory(historyJson, current, period);
  markets.forEach(m => ctx.prevOf.set(m, previousPeriod(ctx.history, m, period)));
  ctx.physical = readPhysical(workbook, ctx);

  const periods = Array.from(new Set<string>(ctx.history.map(r => r.period))).sort();
  const colsText = (l: AssetLoad) => Array.from(l.found.keys()).map(f => `${STANDARD_HEADER[f] || f} = '${l.found.get(f)}'`).join("; ");
  ctx.runInfo = [
    ["Script version", SCRIPT_VERSION],
    ["Current export", `${roles.cur.join(", ")} - as of ${fmtDay(asOf)}` +
      (auditDate ? " (given)" : " (today)") + ` - history period ${period}`],
    ["Previous export", prevLoad ? `${roles.prev.join(", ")} - ${prevPeriod}` : snap ?
      `Snapshot of ${snap.period} (${snap.assets.length} assets; asset no., status, type, model, PO only)` :
      "None - reconciliation (new / removed / status changes) skipped. Add the previous export as a tab named '... Previous Export'"],
    ["How the exports were told apart", roles.note],
    ["Rows read (current)", loadText(cur)],
    ["Rows read (previous)", prevLoad ? loadText(prevLoad) : snap ? `${snap.assets.length} assets from the snapshot` : "-"],
    ["Columns used", colsText(cur)],
    ["Columns with more than one match", cur.ambiguous.length ? cur.ambiguous.join("; ") +
      " - to choose another, add a Mappings row: Column | <standard header> | <your column>" : "None"],
    ["Columns not found", (ASSET_FIELDS.filter(f => !cur.cols.has(f)).map(f => STANDARD_HEADER[f] || f).join(", ") || "None") +
      " - checks that need them are skipped (see Glossary > Needs)"],
    ["PO validation", ctx.po.source ? `${ctx.po.rows.length} rows from ${ctx.po.source}` + (ctx.po.problems.length ?
      `; ${ctx.po.problems.length} need attention: ${ctx.po.problems.slice(0, 5).join("; ")}` : "") :
      "Not supplied - PO reconciliation lists the new assets' POs as 'Missing validation'. Never invented."],
    ["Users reference", users.source ? `${users.source}: ${users.count} users (${users.inactive} inactive or left)` :
      "None - leaver and user-directory checks skipped (add a Users tab or usersJson)"],
    ["Devices reference", devices.source ? `${devices.source}: ${devices.list.length} devices` :
      "None - Intune / SCCM checks skipped; 'Last scan' from the export is used for last-seen checks"],
    ["Market filter", marketFilter || "None (all markets)"],
    ["Markets", markets.map(m => `${m} (${assets.filter(a => a.market === m).length})`).join(", ")],
    ["Stock age from", ctx.stockAgeSource || "Not available (no status / update / purchase date column)"],
    ["Deployments from", ctx.deploySource || "Not available"],
    ["History periods available", periods.filter(p => p <= period).join(", ") || "None - this audit becomes the baseline"],
    ["Mappings", maps.source === "None" ? "None (built-in rules only)" : `${maps.applied} applied from ${maps.source}` +
      (maps.problems.length ? `; ${maps.problems.length} ignored: ${maps.problems.slice(0, 5).join("; ")}` : "")],
    ["Anonymisation", shareSafe ? "On - tokens (deterministic in this run, the same in both exports): " + ctx.tokens.join(", ") +
      ". The mapping is not stored. Raw export and reference tabs removed." :
      "Off - real values shown; share with the requester and the service team only"]
  ];
  progress("Writing the result sheets");
  buildWorkbook(workbook, ctx);
  progress("Result sheets written");
  if (shareSafe) {
    roles.cur.concat(roles.prev, tabs.users, tabs.devices).forEach(n => {
      const w = workbook.getWorksheet(n);
      if (w) { w.delete(); }
    });
  }
  const k = kpiValues(metricsFor(current, ""));
  const flagCount = flagCounts(assets);
  const rc = ctx.recon;
  const report = rc.available ? monthlyReport(ctx) : null;
  return {
    auditDate: fmtDay(asOf), period, previousPeriod: rc.available ? rc.prevPeriod : "", assets: assets.length, inScope: k.inScope,
    markets, reconciliation: rc.available ? reconText(ctx) : [], integrityScore: k.score === "" ? "" : pct(k.score as number, 1),
    recordsWithIssues: k.issues, highSeverity: k.high,
    topChecks: keysByCount(flagCount).filter(c => FLAG_INFO[c].dim !== LIFECYCLE).slice(0, 8)
      .map(c => `${FLAG_INFO[c].name} (${flagCount.get(c)})`),
    leaversHoldingAssets: flagCount.get("LEAVER_HOLDS_ASSET") || 0,
    stockUtilisation: k.util === "" ? "" : pct(k.util as number, 1),
    qualityChecks: rc.available ? qualityChecks(ctx).filter(q => q[3] !== "Pass").map(q => `${q[3]}: ${q[0]} (${q[2]})`) : [],
    runInfo: ctx.runInfo.map(r => `${r[0]}: ${r[1]}`),
    agentInput: buildAgentInput(ctx),
    historyRows: current,
    snapshot,
    emailSubject: report ? report.subject : "",
    emailBody: report ? report.email.join("\n") : "",
    pptPrompt: report ? report.ppt.join("\n") : ""
  };
}
// Progress in the Output tab: seconds since the start and what the script is doing
const CLOCK = { start: 0 };
function progress(text: string): void {
  if (!CLOCK.start) { CLOCK.start = Date.now(); }
  console.log(`${((Date.now() - CLOCK.start) / 1000).toFixed(0)} s - ${text}`);
}
function loadText(l: AssetLoad): string {
  return `${l.rows} data rows; ${l.unique} unique asset numbers; ${l.blankIds} rows without an asset number; ` +
    `${l.dupIds} asset numbers on more than one row` + (l.crossTabDups ? `; ${l.crossTabDups} identical rows repeated across tabs counted once` : "");
}

// ============================================================================ settings
const SETTINGS: Record<string, number> = {
  notVerifiedDays: 365,     // in use / in stock and not physically verified for this many days -> NOT_VERIFIED
  staleSeenDays: 90,        // in use but last scanned / seen longer ago than this -> STALE_DEVICE (VCAMS sets Disabled
                            // after 90 days without a network connection, so an "In operation" asset past 90 days was missed)
  recentSeenDays: 30,       // stock / inactive / scrapped / lost but seen within this many days -> ONLINE_NOT_IN_USE
  stockAgedDays: 180,       // in stock longer than this -> LONG_IN_STOCK (and "aged stock" in Stock & Utilisation)
  repairMaxDays: 30,        // in repair longer than this -> LONG_IN_REPAIR
  transitMaxDays: 60,       // ordered / in transit longer than this -> LONG_IN_TRANSIT
  minSerialChars: 5,        // shorter serial numbers count as invalid
  systemicShare: 0.25,      // a check on more than this share of records = process-wide finding
  marketRateMin: 50,        // markets with fewer records are shown but not ranked
  coverMaxMonths: 3,        // more months of stock cover than this = overstocked
  coverMinMonths: 0.5,      // fewer = low stock
  shortageMinDeployed: 2,   // a type / site that deploys this many a month with no stock left = shortage
  siteMinRecords: 20,       // sites with fewer records are left out of the site score ranking
  physicalSample: 50,       // assets in the Physical Check sample (0 = no sample)
  exceptionsMax: 20000,     // rows on the Exceptions sheet (most severe first) - keeps big exports inside the time limit
  allAssetsMax: 20000,      // above this the All Assets sheet is left empty (the raw export tab has every asset)
  samplesPerCheck: 3        // example assets per check in the agent brief
};
// Hostname naming standard, as a regular expression (empty = not checked). Per market via Mappings
// (Kind 'HostnamePattern', Value = market, MapsTo = pattern), e.g. ^EG[LD]\d{6}$
const HOSTNAME_PATTERN: string = "";

const SCRIPT_VERSION = "2026.10.18.2";
const MAPPINGS_SHEET = "Mappings", PHYSICAL_SHEET = "Physical Check";
const PHYSICAL_COLS = ["Found? (Y/N)", "Right location? (Y/N)", "Right user? (Y/N)", "Comment"];

// Lifecycle stages (used for the audit logic; the reports show the VCAMS status itself)
const IN_USE = "In Use", STOCK = "In Stock", REPAIR = "In Repair", TRANSIT = "On Order / In Transit", INACTIVE = "Inactive",
  RETIRED = "Retired", LOST = "Lost", DISPOSED = "Scrapped / Disposed", UNKNOWN = "Unknown status";
const STAGES = [IN_USE, STOCK, REPAIR, TRANSIT, INACTIVE, RETIRED, LOST, DISPOSED, UNKNOWN];
// The VCAMS controlled statuses, in the capitalisation the reports use, and the stage each one means (confirmed by the
// asset team 2026-10-14): Marketed = donated to charity; Building = being prepared; Disabled = no network connection for
// more than 90 days; Authorised inactive = the user is on long leave and keeps the asset; Ordered = on order.
// Any other status is flagged for review (and read with STAGE_RULES so it still counts somewhere).
const VCAMS_STATUSES: [string, string][] = [
  ["In operation", IN_USE], ["In stock", STOCK], ["Ordered", TRANSIT], ["Reserved", STOCK], ["Ready for use", STOCK],
  ["To install", STOCK], ["Building", STOCK], ["Marketed", DISPOSED], ["Disabled", INACTIVE],
  ["Authorised inactive", INACTIVE], ["Scrapped", DISPOSED], ["Lost", LOST], ["Inactive", INACTIVE], ["Stolen", LOST]
];
// Statuses with the same meaning written another way (compared without case, spaces or punctuation)
const STATUS_SPELLINGS: Record<string, string> = { "authorizedinactive": "Authorised inactive", "inoperations": "In operation" };
// Current statuses that go on the operational "Scrap & Inactive Assets" list when they are new this month
const OPERATIONAL_STATUSES = ["Scrapped", "Lost"];
const STAGE_RULES: [string, RegExp][] = [
  [LOST, /\b(lost|stolen|missing|theft)\b/],
  [DISPOSED, /dispos|scrap|\bsold\b|marketed|donat|recycl|written ?off|write ?off|destroy|returned to (vendor|lessor|supplier)|lease return/],
  [INACTIVE, /inactive|disabled|blocked|suspended/],
  [RETIRED, /retir|decommission|end of life|\beol\b|obsolete|archiv|pending (disposal|write)|awaiting (disposal|collection)|to be disposed/],
  [REPAIR, /repair|\brma\b|faulty|broken|defect|maintenance|damag|warranty claim/],
  [TRANSIT, /transit|shipp|\bordered\b|on order|in delivery|dispatch|pending receipt/],
  [STOCK, /stock|spare|available|storage|warehouse|\bstore\b|\bnew\b|received|unassigned|unallocated|ready|\bpool\b|buffer|reserved|returned|to install|building/],
  [IN_USE, /use|operation|deploy|assign|\bactive\b|allocat|issued|production|\blive\b|in service/]
];
// Stages that are in scope for the integrity score (scrapped / disposed assets are history)
const SCOPE_STAGES = new Set<string>([IN_USE, STOCK, REPAIR, TRANSIT, INACTIVE, RETIRED, LOST, UNKNOWN]);
const NOT_IN_USE_STAGES = new Set<string>([STOCK, INACTIVE, RETIRED, DISPOSED, LOST]);

// The VCAMS asset types: [type as reported, value group, kind]. The kind drives the audit rules (which types need
// a user, a serial, device management, a lifecycle). Types not listed are reported under "Unlisted" and flagged.
const HIGH_VALUE = "High value", LOW_VALUE = "Low value", OTHER_VALUE = "Other", UNLISTED = "Unlisted";
const VALUE_GROUPS = [HIGH_VALUE, LOW_VALUE, OTHER_VALUE, UNLISTED];
const VCAMS_TYPES: [string, string, string][] = [
  ["Notebook", HIGH_VALUE, "Laptop"], ["Desktop", HIGH_VALUE, "Desktop"], ["Monitor", HIGH_VALUE, "Monitor"],
  ["Zubehör-Client", HIGH_VALUE, "Accessory"], ["Headset", HIGH_VALUE, "Headset"],
  ["Bag", LOW_VALUE, "Peripheral"], ["Mouse", LOW_VALUE, "Peripheral"], ["Mobile", LOW_VALUE, "Mobile Phone"],
  ["Keyboard", LOW_VALUE, "Peripheral"],
  ["IP Phone", OTHER_VALUE, "IP Phone"], ["VC Device", OTHER_VALUE, "Meeting Room"], ["Display unit", OTHER_VALUE, "Meeting Room"],
  ["Printer", OTHER_VALUE, "Printer"], ["Wireless sharing device", OTHER_VALUE, "Meeting Room"], ["Connectors", OTHER_VALUE, "Peripheral"],
  ["Scanner", OTHER_VALUE, "Printer"], ["External Drive", OTHER_VALUE, "Peripheral"], ["Virtual Machine", OTHER_VALUE, "Virtual"]
];

// Kinds for types that are not in the list, matched on the asset type and then on the model
const OTHER = "Other";
const CLASS_RULES: [string, RegExp][] = [
  ["Headset", /headset|headphone|earphone|earbud|jabra|plantronics|\bpoly\b|speakerphone/],
  ["Docking Station", /dock/],
  ["Laptop", /laptop|notebook|ultrabook|portable computer|latitude|elitebook|thinkpad|probook|zbook|macbook|surface (laptop|book|pro)|vostro|inspiron|travelmate|lifebook/],
  ["Desktop", /desktop|workstation|\bpc\b|optiplex|elitedesk|prodesk|thinkcentre|\btower\b|mini ?pc|imac|mac mini|thin client|all ?in ?one/],
  ["Monitor", /monitor|display|screen|\blcd\b|\bled\b/],
  ["Tablet", /tablet|ipad|galaxy tab/],
  ["Mobile Phone", /mobile|smart ?phone|iphone|galaxy|android|cell ?phone|handset/],
  ["Printer", /printer|\bmfp\b|\bmfd\b|scanner|copier|plotter|multifunction/],
  ["Network Device", /switch|router|firewall|access point|\bwlan\b|wireless|modem|gateway/],
  ["Server", /server|\bsan\b|\bnas\b|\brack\b|blade|storage array/],
  ["Virtual", /virtual/],
  ["Peripheral", /keyboard|mouse|webcam|camera|adapter|charger|cable|\bbag\b|speaker|token|yubikey|usb|connector|drive/]
];
// Laptop / desktop model families, to spot an asset type that does not match the model
const MODEL_FAMILY: [string, RegExp][] = [
  ["Laptop", /latitude|elitebook|thinkpad|probook|zbook|macbook|surface (laptop|book)|vostro|travelmate|lifebook/],
  ["Desktop", /optiplex|elitedesk|prodesk|thinkcentre|imac|mac mini/]
];
// End-user computing devices: checked against device management and per-user limits
const EUC_CLASSES = new Set<string>(["Laptop", "Desktop"]);
// Assigned to a place, not a person: no user needed when in use
const LOCATION_CLASSES = new Set<string>(["Monitor", "Docking Station", "Accessory", "Printer", "Network Device", "Server",
  "Peripheral", "IP Phone", "Meeting Room", "Virtual"]);
// No serial number expected / not a physical device
const NO_SERIAL_CLASSES = new Set<string>(["Peripheral", "Virtual"]);
const VIRTUAL_CLASSES = new Set<string>(["Virtual"]);
// Most of each kind one user should hold while in use (above = review). Tune via Mappings (Kind 'Limit')
const PER_USER_LIMIT: Record<string, number> = { "Laptop": 1, "Desktop": 1, "Mobile Phone": 1, "Tablet": 1 };
// Expected life in years (purchase date + years = refresh due). Tune via Mappings (Kind 'Lifecycle', by type or kind)
const LIFECYCLE_YEARS: Record<string, number> = {
  "Laptop": 4, "Desktop": 5, "Monitor": 7, "Mobile Phone": 3, "Tablet": 3, "Docking Station": 5, "Accessory": 5, "Headset": 3,
  "Printer": 6, "Network Device": 7, "Server": 6, "IP Phone": 7, "Meeting Room": 6
};

const MARKET_MAP: [string, string][] = [
  ["^ve$", "VOIS EG"], ["india", "VOIS IN"], ["egypt", "VOIS EG"], ["romania", "VOIS EU - Romania"], ["albania", "VOIS EU - Albania"],
  ["hungary", "VOIS EU - Hungary"], ["spain", "VOIS EU - Spain"], ["turk", "VOIS EU - Turkey"], ["czech", "VOIS EU - Czech"],
  ["portugal", "VOIS EU - Portugal"], ["ireland", "VOIS EU - Ireland"], ["germany", "VOIS EU - Germany"], ["\\buk\\b", "VOIS UK"]
];
// Matched on the company / Legal Entity value (lower case, punctuation as spaces). Legal Entity "VE" = Egypt.
// Add other entity codes with Mappings rows: Market | <code> | <market label> (matches the whole value).
// Manufacturer spellings -> one name (only for the "written several ways" report and model checks)
const MAKE_RULES: [string, RegExp][] = [
  ["HP", /hewlett|^hp\b|^hpe?\b|hp inc/], ["Dell", /^dell/], ["Lenovo", /lenovo/], ["Apple", /apple/],
  ["Microsoft", /microsoft/], ["Samsung", /samsung/], ["Cisco", /cisco/], ["Jabra", /jabra|^gn\b/],
  ["Poly", /\bpoly\b|plantronics/], ["Logitech", /logitech/], ["Philips", /philips/], ["LG", /^lg\b/],
  ["Asus", /asus/], ["Acer", /acer/], ["Fujitsu", /fujitsu/], ["Dynabook", /toshiba|dynabook/], ["Xerox", /xerox/],
  ["Canon", /canon/], ["Ricoh", /ricoh/], ["Kyocera", /kyocera/], ["Brother", /brother/], ["Epson", /epson/]
];
const PLACEHOLDER = /^(n ?a|none|null|nil|unknown|tbd|tbc|to be filled.*|default string|system serial number|chassis serial number|not (available|specified|applicable)|serial|invalid|0+|1234567890?|123456789|x+|-+|\.+|\?+)$/;
const BLANK_USER = /^(|n ?a|none|null|unknown|unassigned|not assigned|spare|stock|it stock|-+|\.+)$/;

// PO validation (user-supplied, never invented)
const PO_HEADERS = ["Purchase Date", "Purchase Number", "Model", "Quantity", "Raised By", "Physical Security Validation", "Spider Validation"];
const PO_LABELS = ["Yes", "No", "Pending", "Not Applicable"];
const PO_RESULTS = { matched: "Matched", variance: "Quantity variance", missing: "Missing validation", notFound: "PO not found in new assets",
  pending: "Pending approval", blank: "Blank PO number" };

// ============================================================================ checks
const COMPLETENESS = "Completeness", VALIDITY = "Validity", UNIQUENESS = "Uniqueness", CONSISTENCY = "Consistency",
  TIMELINESS = "Timeliness", RECONCILIATION = "Reconciliation", LIFECYCLE = "Lifecycle";
// The integrity score counts these; Reconciliation depends on which references were supplied and Lifecycle is not a
// data error, so both are reported on their own
const SCORED_DIMS = [COMPLETENESS, VALIDITY, UNIQUENESS, CONSISTENCY, TIMELINESS];
const ALL_DIMS = SCORED_DIMS.concat([RECONCILIATION, LIFECYCLE]);
const SEVERITY = ["High", "Medium", "Low"];
interface FlagInfo { name: string; dim: string; sev: string; why: string; fix: string; owner: string; }
const FLAG_INFO: Record<string, FlagInfo> = {
  NO_SERIAL: { name: "Serial number missing", dim: COMPLETENESS, sev: "High",
    why: "The asset cannot be matched to the device, the warranty or device management", fix: "Read the serial from the device or the delivery note and update VCAMS", owner: "Asset team" },
  NO_TAG: { name: "Asset tag missing", dim: COMPLETENESS, sev: "Medium",
    why: "The asset cannot be scanned or found in a physical audit", fix: "Label the asset and record the tag", owner: "Asset team" },
  NO_MODEL: { name: "Manufacturer or model missing", dim: COMPLETENESS, sev: "Low",
    why: "Refresh planning, stock and warranty reports cannot group it", fix: "Fill manufacturer and model from the device or the order", owner: "Asset team" },
  NO_CLASS: { name: "Asset type missing or not in the standard list", dim: COMPLETENESS, sev: "Medium",
    why: "The asset falls outside the High / Low value / Other summary and the per-type rules", fix: "Use a standard asset type (or add a Group mapping for a new type)", owner: "Asset team / VCAMS owner" },
  NEW_NO_PO: { name: "New this month without a purchase number", dim: COMPLETENESS, sev: "Medium",
    why: "The new asset cannot be reconciled to a purchase order", fix: "Add the PurchaseNo from the order / delivery note", owner: "Asset team / Procurement" },
  NO_USER: { name: "In use but no user assigned", dim: COMPLETENESS, sev: "High",
    why: "Nobody is accountable for the device; it cannot be recovered when the user leaves", fix: "Find the user (device management, last logon) and assign them, or move the asset to stock", owner: "Asset team / Service desk" },
  NO_LOCATION: { name: "Location missing", dim: COMPLETENESS, sev: "Medium",
    why: "The asset cannot be found for audit, repair or collection", fix: "Record the site / building", owner: "Asset team" },
  NO_COST_CENTER: { name: "Cost center missing", dim: COMPLETENESS, sev: "Low",
    why: "The cost cannot be charged back", fix: "Take the cost center from the user's HR record", owner: "Asset team / Finance" },
  NO_PURCHASE_DATE: { name: "Purchase date missing", dim: COMPLETENESS, sev: "Low",
    why: "Age, refresh and warranty cannot be worked out", fix: "Fill from the purchase order or the vendor warranty lookup", owner: "Asset team / Procurement" },
  BAD_SERIAL: { name: "Serial number is a placeholder or too short", dim: VALIDITY, sev: "High",
    why: "Placeholder serials ('N/A', 'To be filled by O.E.M.', '0') make duplicates and false matches", fix: "Replace with the real serial from the device BIOS / label", owner: "Asset team" },
  BAD_EMAIL: { name: "User email not valid", dim: VALIDITY, sev: "Medium",
    why: "The user cannot be matched to HR or device management", fix: "Correct the email (name@domain)", owner: "Asset team" },
  UNKNOWN_STATUS: { name: "Status not in the controlled list", dim: VALIDITY, sev: "Medium",
    why: "Non-standard statuses split the reports and may be counted in the wrong stage", fix: "Use a controlled status (or add a Status mapping)", owner: "Asset team / VCAMS owner" },
  BAD_DATES: { name: "Impossible dates", dim: VALIDITY, sev: "Medium",
    why: "Dates in the future or in the wrong order break age, warranty and audit reports", fix: "Correct the dates", owner: "Asset team" },
  BAD_HOSTNAME: { name: "Hostname does not follow the naming standard", dim: VALIDITY, sev: "Low",
    why: "Non-standard names are hard to match to device management and the market", fix: "Rename the device or correct the record", owner: "EUC / Asset team" },
  DUP_SERIAL: { name: "Same serial number on more than one record", dim: UNIQUENESS, sev: "High",
    why: "The same device is counted twice, or two devices share a wrong serial", fix: "Check both records; merge or correct, keep one", owner: "Asset team" },
  DUP_TAG: { name: "Same asset tag on more than one record", dim: UNIQUENESS, sev: "High",
    why: "A scanned tag finds the wrong asset", fix: "Re-label one asset and correct the record", owner: "Asset team" },
  DUP_ASSET_ID: { name: "Same asset ID on more than one record", dim: UNIQUENESS, sev: "High",
    why: "The record key is not unique; updates go to the wrong record", fix: "Raise with the VCAMS owner to merge the records", owner: "VCAMS owner" },
  DUP_HOSTNAME: { name: "Same hostname on more than one record", dim: UNIQUENESS, sev: "Medium",
    why: "Device management and tickets match the wrong asset", fix: "Check which device has the name; correct the other record", owner: "Asset team" },
  DUP_MAC: { name: "Same MAC address on more than one record", dim: UNIQUENESS, sev: "Medium",
    why: "Network matching finds the wrong asset", fix: "Correct the MAC on the wrong record", owner: "Asset team" },
  USER_ON_NOT_IN_USE: { name: "User still assigned to a stock / retired / disposed asset", dim: CONSISTENCY, sev: "Medium",
    why: "The user looks like they still hold it; their leaver check and stock counts are wrong", fix: "Clear the user, or set the status to In Use if they really have it", owner: "Asset team" },
  STATUS_REVERSED: { name: "Back in use or stock after being scrapped / marketed / lost", dim: CONSISTENCY, sev: "High",
    why: "A scrapped or sold asset cannot come back; a found asset needs checking", fix: "Check the asset physically; correct the status or the asset number", owner: "Asset team" },
  CLASS_MODEL_MISMATCH: { name: "Asset type does not match the model", dim: CONSISTENCY, sev: "Low",
    why: "Laptops recorded as desktops (or the other way round) distort refresh and stock planning", fix: "Correct the asset type", owner: "Asset team" },
  MANY_DEVICES: { name: "User holds more devices of a type than allowed", dim: CONSISTENCY, sev: "Medium",
    why: "Old devices are not returned, or the record was not updated after a swap", fix: "Recover the extra device or record why the user needs it", owner: "Asset team / Line manager" },
  NOT_VERIFIED: { name: "Not physically verified within the audit cycle", dim: TIMELINESS, sev: "Medium",
    why: "Nobody has seen the asset recently; it may be lost", fix: "Include it in the next physical audit / ask the user to confirm", owner: "Asset team" },
  LONG_IN_STOCK: { name: "In stock longer than the stock limit", dim: TIMELINESS, sev: "Low",
    why: "Idle capital; the device ages on the shelf", fix: "Redeploy before buying new, or move to a site that needs it", owner: "Asset team / Procurement" },
  LONG_IN_REPAIR: { name: "In repair longer than the repair limit", dim: TIMELINESS, sev: "Medium",
    why: "The repair is stuck or the record was not updated", fix: "Chase the vendor / update the status", owner: "Asset team" },
  LONG_IN_TRANSIT: { name: "In transit longer than the transit limit", dim: TIMELINESS, sev: "Medium",
    why: "The delivery may be lost or was received without an update", fix: "Confirm receipt and update the status", owner: "Asset team / Logistics" },
  LEAVER_HOLDS_ASSET: { name: "Assigned to a user who has left", dim: RECONCILIATION, sev: "High",
    why: "Unrecovered asset and data risk; often the biggest audit finding", fix: "Recover the device through the leaver process; update VCAMS", owner: "Asset team / HR leaver process" },
  USER_NOT_FOUND: { name: "Assigned user not found in the user directory", dim: RECONCILIATION, sev: "Medium",
    why: "The user may have left, or the email / ID is wrong", fix: "Check the user; correct the email / ID or recover the device", owner: "Asset team" },
  NOT_SEEN: { name: "In use but not in device management", dim: RECONCILIATION, sev: "High",
    why: "The device may be lost, not enrolled, or the serial / hostname is wrong", fix: "Find the device; enrol it or correct the record", owner: "EUC / Asset team" },
  STALE_DEVICE: { name: "In use but not seen online recently", dim: RECONCILIATION, sev: "Medium",
    why: "The device may be lost, broken or unused", fix: "Contact the user; recover unused devices to stock", owner: "Asset team" },
  ONLINE_NOT_IN_USE: { name: "Online recently but recorded as stock / retired / disposed / lost", dim: RECONCILIATION, sev: "High",
    why: "Someone is using a device VCAMS says is not in use - a security and audit risk", fix: "Find the user; correct the status or recover the device", owner: "Asset team / Security" },
  USER_MISMATCH: { name: "Different user in device management", dim: RECONCILIATION, sev: "Medium",
    why: "The device changed hands without an update", fix: "Confirm who has it and update the assignment", owner: "Asset team" },
  WARRANTY_EXPIRED: { name: "In use after warranty / lease end", dim: LIFECYCLE, sev: "Low",
    why: "Repairs are paid; leased devices may incur penalties", fix: "Plan the refresh or extend cover", owner: "Asset team / Procurement" },
  OVER_AGE: { name: "In use beyond its expected life", dim: LIFECYCLE, sev: "Low",
    why: "Older devices fail more and slow users down", fix: "Add to the refresh plan", owner: "Asset team / EUC" },
  STOCK_OVER_AGE: { name: "In stock but already beyond its expected life", dim: LIFECYCLE, sev: "Low",
    why: "Deploying it gives a user an old device", fix: "Dispose or use only as a short-term loan", owner: "Asset team" }
};

// ============================================================================ columns
// VCAMS / Spider export column names, matched ignoring case, spaces and punctuation. The first name in each list is
// the standard header; the others are accepted alternatives. Add a name when a new export layout appears, or map a
// column without changing the script with a Mappings row: Column | <standard header> | <your column>.
const ASSET_COLS: Record<string, string[]> = {
  assetId: ["Asset no.", "Asset No", "Asset Number", "Asset ID", "AssetID", "Inventory Number", "Inventory No", "Inventory ID", "Inventarnummer"],
  serial: ["Serial no.", "Serial No", "Serial Number", "Serial", "SerialNumber", "S/N", "Service Tag", "Seriennummer"],
  status: ["Asset status", "Status", "Current Status", "Asset State", "Lifecycle Status", "State"],
  cls: ["Asset Type", "Type", "Category", "Asset Class", "Device Type", "Asset Category", "Product Type", "Hardware Type"],
  model: ["Model", "Asset Model", "Device Model", "Model Name", "Product", "Product Name", "Modell"],
  purchase: ["Purchase date", "PO Date", "Acquisition Date", "Date Purchased", "Delivery Date", "Received Date", "Kaufdatum"],
  po: ["PurchaseNo", "Purchase No", "Purchase Number", "PO Number", "PO No", "Purchase Order", "Order Number"],
  lastSeen: ["Last scan", "Last Scan Date", "Last Seen", "Last Discovered", "Last Contact", "Last Check-in", "Last Sync"],
  scanAge: ["Scandate Age in Days", "Scan Age", "Scan Age Days", "Scan Age in Days"],
  functionUnit: ["Function Unit", "Functional Unit", "FunctionUnit", "Function"],
  created: ["AssetCreationDate", "Asset Creation Date", "Creation Date", "Created Date", "Created On"],
  createdAge: ["Creation Age in Days", "Creation Age", "Age in Days"],
  updated: ["AssetUpdateDate", "Asset Update Date", "Update Date", "Last Modified", "Modified Date", "Last Updated", "Updated"],
  tag: ["Asset Tag", "Tag", "Tag Number", "Asset Label", "Barcode", "Label Number"],
  hostname: ["Hostname", "Host Name", "Device Name", "Computer Name", "Machine Name", "Network Name"],
  mac: ["MAC Address", "MAC", "Ethernet MAC", "WiFi MAC", "Wireless MAC"],
  make: ["Manufacturer", "Make", "Vendor", "Brand", "Hersteller"],
  user: ["Assigned To", "Assigned User", "User", "User Name", "Employee Name", "Employee", "Used By", "Custodian", "Owner", "Benutzer"],
  email: ["User Email", "Email", "E-mail", "Email Address", "Assigned To Email", "Owner Email", "Employee Email", "UPN", "User Principal Name"],
  empId: ["Employee ID", "Employee Number", "Staff ID", "Personnel Number", "Emp ID", "User ID", "Personalnummer"],
  costCenter: ["Cost Center", "Cost Centre", "Cost Center Code", "CostCenter", "Kostenstelle"],
  location: ["Location", "Site", "Building", "Office", "Site Name", "Location Name", "Standort"],
  company: ["Company", "Legal Entity", "Market", "Country", "Organisation", "Organization", "Entity", "OpCo", "Business Unit"],
  managedBy: ["Managed By", "Support Group", "Asset Manager", "Owner Group", "Responsible Team", "Asset Owner Group"],
  warranty: ["Warranty End", "Warranty End Date", "Warranty Expiry", "Warranty Expiration Date", "Warranty Until", "Lease End", "Lease End Date", "Contract End"],
  verified: ["Last Audit Date", "Last Inventory Date", "Last Verified", "Last Verified Date", "Audit Date", "Inventory Date", "Last Physical Audit"],
  statusDate: ["Status Date", "Status Change Date", "Status Changed", "Last Status Change"],
  deployDate: ["Deployment Date", "Deployed Date", "Assigned Date", "Assignment Date", "Issued Date", "Issue Date", "Allocation Date", "Handover Date", "Asset Assignment Date"],
  cost: ["Cost", "Purchase Price", "Price", "Purchase Cost", "Value", "Asset Value", "Net Book Value", "Book Value"],
  comment: ["Comment", "Comments", "Spider Comment", "Asset Comment", "Remarks", "Remark", "Notes", "Note"],
  statusComment: ["Status comment", "Status Comments", "Status Remark", "Status Remarks", "Status Note"]
};
const ASSET_FIELDS = Object.keys(ASSET_COLS);
// Needed to reconcile two exports: the run stops with a clear message when one is missing
const REQUIRED_FIELDS = ["assetId", "status", "cls"];
const STANDARD_HEADER: Record<string, string> = {
  assetId: "Asset no.", serial: "Serial no.", status: "Asset status", cls: "Asset Type", model: "Model", purchase: "Purchase date",
  po: "PurchaseNo", lastSeen: "Last scan", scanAge: "Scandate Age in Days", functionUnit: "Function Unit",
  created: "AssetCreationDate", createdAge: "Creation Age in Days", updated: "AssetUpdateDate", tag: "Asset Tag",
  hostname: "Hostname", mac: "MAC Address", make: "Manufacturer", user: "Assigned To", email: "User Email",
  empId: "Employee ID", costCenter: "Cost Center", location: "Location", company: "Company / market", managedBy: "Managed By",
  warranty: "Warranty End", verified: "Last Audit Date", statusDate: "Status Date", deployDate: "Deployment Date", cost: "Cost", comment: "Comment", statusComment: "Status comment"
};
const FIELD_LABEL = STANDARD_HEADER;
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
const PO_COLS: Record<string, string[]> = {
  date: ["Purchase Date", "PO Date"], number: ["Purchase Number", "PurchaseNo", "Purchase No", "PO Number", "PO"],
  model: ["Model"], quantity: ["Quantity", "Qty"], raisedBy: ["Raised By", "Requested By"],
  physical: ["Physical Security Validation", "Physical Security"], spider: ["Spider Validation", "Spider", "VCAMS Validation"]
};
const USER_TAB = /\b(users?|hr|employees?|people|staff|leavers?|entra|azure ad|active directory|directory|workday)\b/;
const DEVICE_TAB = /\b(intune|sccm|mecm|defender|devices?|discovery|endpoints?|computers?|jamf|tanium)\b/;

function headerKey(v: Raw): string {
  return String(v === undefined || v === null ? "" : v).toLowerCase().replace(/[^a-z0-9äöüß]/g, "");
}
interface ColumnMatch { pos: Record<string, number>; amb: string[]; }
// Field -> column position; each column is used once, fields in the order of the map. A field that more than one
// column could fill is reported in amb (the first match in the list is used). forced: field -> header key.
function mapColumns(raw: Raw[], cols: Record<string, string[]>, forced?: Map<string, string>): ColumnMatch {
  const header = raw.map(v => headerKey(v));
  const pos: Record<string, number> = {};
  const amb: string[] = [];
  const taken = new Set<number>();
  if (forced) {
    forced.forEach((key, f) => {
      const i = header.indexOf(key);
      if (cols[f] && i >= 0) { pos[f] = i; taken.add(i); }
    });
  }
  Object.keys(cols).forEach(f => {
    if (pos[f] !== undefined) { return; }
    const hits: number[] = [];
    cols[f].forEach(name => {
      const i = header.indexOf(headerKey(name));
      if (i >= 0 && !taken.has(i) && hits.indexOf(i) < 0) { hits.push(i); }
    });
    if (!hits.length) { return; }
    pos[f] = hits[0];
    taken.add(hits[0]);
    if (hits.length > 1) {
      amb.push(`${STANDARD_HEADER[f] || f}: used '${String(raw[hits[0]]).trim()}', also matched ` +
        hits.slice(1).map(i => `'${String(raw[i]).trim()}'`).join(", "));
    }
  });
  return { pos, amb };
}
// The row in the first 20 with the most recognised columns
function bestHeader(rows: Raw[][], cols: Record<string, string[]>, forced?: Map<string, string>): { row: number; pos: Record<string, number>; amb: string[] } {
  let best = -1, bestPos: Record<string, number> = {}, bestAmb: string[] = [], bestN = 1;
  rows.slice(0, 20).forEach((r, i) => {
    const m = mapColumns(r, cols, forced), n = Object.keys(m.pos).length;
    if (n > bestN) { best = i; bestPos = m.pos; bestAmb = m.amb; bestN = n; }
  });
  return { row: best, pos: bestPos, amb: bestAmb };
}
function has(pos: Record<string, number>, f: string): boolean {
  return pos[f] !== undefined;
}

// ============================================================================ types
type Raw = string | number | boolean;
type Cell = string | number;
interface Flag { code: string; detail: string; }
interface Asset {
  key: string; sheet: string; row: number; assetId: string; tag: string; serial: string; serialN: string;
  hostname: string; hostN: string; mac: string; macN: string; clsRaw: string; assetClass: string; group: string; kind: string;
  makeRaw: string; make: string; model: string; statusRaw: string; status: string; standardStatus: boolean; stage: string;
  user: string; email: string; empId: string; costCenter: string; location: string; company: string; market: string;
  managedBy: string; cost: number; po: string; functionUnit: string; comment: string; statusComment: string; scanAge: Cell; createdAge: Cell;
  purchase: Date | null; warranty: Date | null; verified: Date | null; lastSeen: Date | null; statusDate: Date | null;
  deployDate: Date | null; created: Date | null; updated: Date | null; badDates: string[]; flags: Flag[];
  device: DeviceRef | null; prev: Asset | null; isNew: boolean; deployedNow: boolean;
}
interface UserRef { email: string; empId: string; name: string; active: boolean; leaveDate: Date | null; }
interface Users { source: string; byEmail: Map<string, UserRef>; byEmp: Map<string, UserRef>; count: number; inactive: number; }
interface DeviceRef { name: string; nameN: string; serial: string; serialN: string; lastSeen: Date | null; user: string; market: string; matched: boolean; }
interface Devices { source: string; list: DeviceRef[]; bySerial: Map<string, DeviceRef>; byName: Map<string, DeviceRef>; }
interface Mappings {
  status: Map<string, string>; types: Map<string, [string, string]>; market: [string, string][]; limits: Map<string, number>;
  lifecycle: Map<string, number>; hostname: Map<string, RegExp>; columns: Map<string, string>; off: Set<string>;
  source: string; applied: number; problems: string[];
}
interface TabInfo { name: string; headerAbs: number; lastAbs: number; colAbs: Record<string, number>; rows: number; }
interface AssetLoad {
  assets: Asset[]; rows: number; unique: number; blankIds: number; dupIds: number; crossTabDups: number;
  cols: Set<string>; found: Map<string, string>; ambiguous: string[]; tabs: TabInfo[];
}
interface DupGroup { on: string; value: string; assets: Asset[]; }
interface UserGroup { user: string; email: string; assetClass: string; assets: Asset[]; }
interface HistoryRow { period: string; market: string; assetClass: string; metric: string; value: number; }
interface PhysicalMark { found: string; loc: string; user: string; comment: string; }
interface Physical {
  sample: Asset[]; marks: Map<string, PhysicalMark>; found: number; missing: number; locOk: number; locChecked: number;
  userOk: number; userChecked: number;
}
interface PoRow { date: Cell; number: string; model: string; quantity: Cell; raisedBy: string; physical: string; spider: string; problems: string[]; }
interface PoData { source: string; rows: PoRow[]; problems: string[]; }
interface StatusChange { cur: Asset; prev: Asset; transition: string; }
interface SerialChange { cur: Asset; prev: Asset; change: string; }
interface Recon {
  available: boolean; source: string; prevPeriod: string; curPeriod: string; prevDay: Date | null; curDay: Date | null;
  prevRows: number; curRows: number;
  prevUnique: number; curUnique: number; prevBlank: number; curBlank: number; prevDup: string[]; curDup: string[];
  serialChanges: SerialChange[]; serialChecked: boolean;
  prevByType: Map<string, number>; curByType: Map<string, number>; newAssets: Asset[]; removed: Asset[]; existing: number;
  changes: StatusChange[]; operational: StatusChange[]; unknownStatuses: Map<string, number>; unlistedTypes: Map<string, number>;
  fromSnapshot: boolean;
}
interface Ctx {
  assets: Asset[]; asOf: Date; period: string; markets: string[]; users: Users; devices: Devices; cols: Set<string>;
  maps: Mappings; focus: string; shareSafe: boolean; dupGroups: DupGroup[]; manyDevices: UserGroup[]; orphans: DeviceRef[];
  runInfo: string[][]; history: HistoryRow[]; prevOf: Map<string, string>; physical: Physical;
  stockAgeSource: string; deploySource: string; recon: Recon; po: PoData; formulas: boolean; curTabs: TabInfo[];
  prevTabs: TabInfo[]; tokens: string[];
}
interface Kpis {
  records: number; inScope: number; stages: Record<string, number>; issues: number; high: number; score: Cell;
  util: Cell; dup: number; stockAged: number; deployed: number; cover: Cell; newAssets: number; removed: number; changes: number;
  serialChanges: number; scrapped: number; lost: number; checks: Map<string, number>;
}
interface AuditResult {
  auditDate: string; period: string; previousPeriod: string; assets: number; inScope: number; markets: string[];
  reconciliation: string[]; integrityScore: string; recordsWithIssues: number; highSeverity: number; topChecks: string[];
  leaversHoldingAssets: number; stockUtilisation: string; qualityChecks: string[]; runInfo: string[]; agentInput: string;
  historyRows: HistoryRow[]; snapshot: string; emailSubject: string; emailBody: string; pptPrompt: string;
}

// ============================================================================ helpers
const DAY = 86400000;
function norm(s: string): string {
  return (s || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}
function countBy(items: string[]): Map<string, number> {
  const m = new Map<string, number>();
  items.forEach(i => m.set(i, (m.get(i) || 0) + 1));
  return m;
}
function keysByCount(m: Map<string, number>): string[] {
  return Array.from(m.keys()).sort((a, b) => (m.get(b) || 0) - (m.get(a) || 0) || a.localeCompare(b));
}
function topList(items: string[], n: number): string {
  const m = countBy(items);
  return keysByCount(m).slice(0, n).map(k => `${k} (${m.get(k)})`).join("; ");
}
function groupBy(items: Asset[], keyOf: (a: Asset) => string): Map<string, Asset[]> {
  const m = new Map<string, Asset[]>();
  items.forEach(a => {
    const k = keyOf(a);
    if (!k) { return; }
    const list = m.get(k);
    if (list) { list.push(a); } else { m.set(k, [a]); }
  });
  return m;
}
function pct(a: number, b: number): string {
  return b ? (100 * a / b).toFixed(1) + "%" : "0%";
}
function ratio(a: number, b: number): Cell {
  return b ? a / b : "";
}
function short(s: string, n: number): string {
  const t = (s || "").replace(/\s+/g, " ").trim();
  return t.length <= n ? t : t.slice(0, n - 1) + "…";
}
function utcToday(): Date {
  const d = new Date();
  return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
}
function monthOf(d: Date): string {
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
}
function fmtDay(d: Date | null): string {
  return d ? `${monthOf(d)}-${String(d.getUTCDate()).padStart(2, "0")}` : "";
}
function daysSince(d: Date | null, asOf: Date): number {
  return d ? Math.floor((asOf.getTime() - d.getTime()) / DAY) : -1;
}
function addYears(d: Date, years: number): Date {
  return new Date(Date.UTC(d.getUTCFullYear() + Math.floor(years), d.getUTCMonth() + Math.round((years % 1) * 12), d.getUTCDate()));
}
function parseDate(v: Raw): Date | null {
  if (typeof v === "number") {
    return v > 1000 && v < 80000 ? new Date(Date.UTC(1899, 11, 30) + Math.round(v * DAY)) : null;
  }
  const s = String(v || "").trim();
  if (!s) { return null; }
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) {
    return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)));
  }
  m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?/i);
  if (m) {
    const a = +m[1], b = +m[2];
    const dayFirst = a > 12 || s.indexOf(".") > 0 || (b <= 12 && a <= 12);  // VOIS markets write day first
    const month = dayFirst ? b : a, day = dayFirst ? a : b;
    let h = +(m[4] || 0);
    if (m[7] && m[7].toUpperCase() === "PM" && h < 12) { h += 12; }
    if (m[7] && m[7].toUpperCase() === "AM" && h === 12) { h = 0; }
    return month >= 1 && month <= 12 ? new Date(Date.UTC(+m[3], month - 1, day, h, +(m[5] || 0), +(m[6] || 0))) : null;
  }
  m = s.match(/^(\d{1,2})[ \-]([A-Za-z]{3})[A-Za-z]*[ \-,]+(\d{4})/);
  if (m) {
    const mi = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"].indexOf(m[2].toLowerCase());
    return mi >= 0 ? new Date(Date.UTC(+m[3], mi, +m[1])) : null;
  }
  return null;
}
function marketOf(company: string, maps: Mappings): string {
  const c = norm(company);
  for (const [key, label] of maps.market.concat(MARKET_MAP)) {
    if (new RegExp(key).test(c)) { return label; }
  }
  return (company || "").trim() || "(no market)";
}
function compactKey(s: string): string {
  return (s || "").toLowerCase().replace(/[^a-z0-9äöüß]/g, "");
}
const STATUS_INDEX = new Map<string, [string, string]>(VCAMS_STATUSES.map(x => [compactKey(x[0]), x] as [string, [string, string]]));
const TYPE_INDEX = new Map<string, [string, string, string]>(VCAMS_TYPES.map(x => [compactKey(x[0]), x] as [string, [string, string, string]]));
// VCAMS status -> the controlled label (trimmed, standard capitalisation), whether it is a controlled status, and its stage
function statusOf(raw: string, maps: Mappings): { label: string; standard: boolean; stage: string } {
  const label = (raw || "").replace(/\s+/g, " ").trim(), k = compactKey(label);
  const mapped = maps.status.get(k);
  if (mapped) { return { label, standard: true, stage: mapped }; }
  const known = STATUS_INDEX.get(k) || STATUS_INDEX.get(compactKey(STATUS_SPELLINGS[k] || ""));
  if (known) { return { label: known[0], standard: true, stage: known[1] }; }
  const s = norm(label);
  if (!s) { return { label: "", standard: false, stage: UNKNOWN }; }
  for (const [stage, re] of STAGE_RULES) {
    if (re.test(s)) { return { label, standard: false, stage }; }
  }
  return { label, standard: false, stage: UNKNOWN };
}
// VCAMS asset type -> the type as reported, its value group (High / Low / Other / Unlisted) and its kind
function typeOf(raw: string, model: string, maps: Mappings): { label: string; group: string; kind: string } {
  const label = (raw || "").replace(/\s+/g, " ").trim(), k = compactKey(label);
  const mapped = maps.types.get(k);
  const known = TYPE_INDEX.get(k);
  let kind = mapped && mapped[1] ? mapped[1] : known ? known[2] : "";
  if (!kind) {
    kind = OTHER;
    for (const text of [norm(label), norm(model)]) {
      const hit = text ? CLASS_RULES.find(r => r[1].test(text)) : undefined;
      if (hit) { kind = hit[0]; break; }
    }
  }
  return { label: known ? known[0] : label || "(no type)", group: mapped && mapped[0] ? mapped[0] : known ? known[1] : UNLISTED, kind };
}
function makeOf(raw: string): string {
  const s = norm(raw);
  for (const [name, re] of MAKE_RULES) {
    if (re.test(s)) { return name; }
  }
  return (raw || "").trim();
}
function serialKey(s: string): string {
  return (s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}
function hostKey(s: string): string {
  return (s || "").toLowerCase().split(".")[0].replace(/[^a-z0-9\-]/g, "");
}
function isPlaceholder(s: string): boolean {
  return PLACEHOLDER.test(norm(s)) || PLACEHOLDER.test((s || "").trim().toLowerCase());
}
function addFlag(a: Asset, code: string, detail: string, maps: Mappings): void {
  if (!maps.off.has(code) && !a.flags.some(f => f.code === code)) { a.flags.push({ code, detail }); }
}
function hasFlag(a: Asset, code: string): boolean {
  return a.flags.some(f => f.code === code);
}
// Counts per check over all assets (push, not concat: concat in a reduce is quadratic on 100,000 assets)
function flagCounts(assets: Asset[]): Map<string, number> {
  const m = new Map<string, number>();
  assets.forEach(a => a.flags.forEach(f => m.set(f.code, (m.get(f.code) || 0) + 1)));
  return m;
}
function flatten(lists: string[][]): string[] {
  const out: string[] = [];
  lists.forEach(l => l.forEach(x => out.push(x)));
  return out;
}
function uniqueCodes(a: Asset): string[] {
  return a.flags.map(f => f.code);
}
function sevRank(code: string): number {
  return SEVERITY.indexOf(FLAG_INFO[code].sev);
}
// Records with no issue at all (scored dimensions) and with no issue in each dimension, in one pass over the list
function cleanCounts(list: Asset[]): { clean: number; dims: number[] } {
  const dims = ALL_DIMS.map(() => 0);
  let clean = 0;
  list.forEach(a => {
    if (!a.flags.length) { clean++; dims.forEach((n, i) => { dims[i] = n + 1; }); return; }
    const hit = ALL_DIMS.map(() => false);
    let scored = false;
    a.flags.forEach(f => {
      const d = FLAG_INFO[f.code].dim, i = ALL_DIMS.indexOf(d);
      if (i >= 0) { hit[i] = true; }
      if (SCORED_DIMS.indexOf(d) >= 0) { scored = true; }
    });
    if (!scored) { clean++; }
    hit.forEach((h, i) => { if (!h) { dims[i]++; } });
  });
  return { clean, dims };
}
function dataIssues(a: Asset): Flag[] {
  return a.flags.filter(f => SCORED_DIMS.indexOf(FLAG_INFO[f.code].dim) >= 0);
}
function hasUser(a: Asset): boolean {
  return !BLANK_USER.test(norm(a.user)) || !!a.email || !!a.empId;
}
function userKey(a: Asset): string {
  return a.email ? a.email.toLowerCase() : a.empId ? "id:" + a.empId.toLowerCase() : BLANK_USER.test(norm(a.user)) ? "" : "name:" + norm(a.user);
}
function userLabel(a: Asset): string {
  return a.user || a.email || a.empId;
}
function inScope(a: Asset): boolean {
  return SCOPE_STAGES.has(a.stage);
}
function plainFlags(a: Asset): string {
  return a.flags.slice().sort((x, y) => sevRank(x.code) - sevRank(y.code)).map(f => FLAG_INFO[f.code].name).join("; ");
}
function worstSeverity(a: Asset): string {
  const r = a.flags.reduce((m, f) => Math.min(m, sevRank(f.code)), 9);
  return r < 9 ? SEVERITY[r] : "";
}
function lifecycleYears(a: Asset, maps: Mappings): number {
  return maps.lifecycle.get(compactKey(a.assetClass)) || maps.lifecycle.get(compactKey(a.kind)) || LIFECYCLE_YEARS[a.kind] || 0;
}
function dateSerial(d: Date | null): Cell {
  return d ? Math.round((d.getTime() - Date.UTC(1899, 11, 30)) / 60000) / 1440 : "";
}
function colLetter(i: number): string {
  let s = "", n = i + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function monthName(period: string): string {
  const m = period.match(/^(\d{4})-(\d{2})$/);
  return m ? ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November",
    "December"][+m[2] - 1] + " " + m[1] : period;
}

// ============================================================================ loading
// Reads a block in row chunks sized from a sample, so each response stays far below the ~5 MB limit even for
// exports of 100,000+ rows
// Each read is one round trip to Excel, so read as much as safely fits in one response (~5 MB limit)
const READ_BYTES = 3000000, READ_MAX_ROWS = 110000;
function readBlock(ws: ExcelScript.Worksheet, r0: number, c0: number, rows: number, cols: number, label?: string): Raw[][] {
  const out: Raw[][] = [];
  if (rows < 1 || cols < 1) { return out; }
  const first = Math.min(rows, 200);
  const sample = ws.getRangeByIndexes(r0, c0, first, cols).getValues() as Raw[][];
  sample.forEach(r => out.push(r));
  let total = 0;
  sample.forEach(r => r.forEach(v => { total += String(v).length + 8; }));
  const bytes = total / Math.max(1, sample.length);
  const chunk = Math.max(200, Math.min(READ_MAX_ROWS, Math.floor(READ_BYTES / Math.max(1, bytes))));
  for (let i = first; i < rows; i += chunk) {
    const part = ws.getRangeByIndexes(r0 + i, c0, Math.min(chunk, rows - i), cols).getValues() as Raw[][];
    part.forEach(r => out.push(r));
    if (label && rows > 5000) { progress(`${label}: ${Math.min(rows, i + chunk).toLocaleString("en-US")} of ${rows.toLocaleString("en-US")} rows`); }
  }
  return out;
}
function readValues(ws: ExcelScript.Worksheet): Raw[][] {
  const used = ws.getUsedRange(true);
  if (!used) { return []; }
  return readBlock(ws, used.getRowIndex(), used.getColumnIndex(), used.getRowCount(), used.getColumnCount());
}
function headRows(ws: ExcelScript.Worksheet): Raw[][] {
  const used = ws.getUsedRange(true);
  if (!used) { return []; }
  return ws.getRangeByIndexes(used.getRowIndex(), used.getColumnIndex(), Math.min(20, used.getRowCount()),
    used.getColumnCount()).getValues() as Raw[][];
}

// Sheet NAMES, never sheet objects: Office Scripts does not allow storing Excel objects (look them up by name instead)
interface Tabs { assets: string[]; users: string[]; devices: string[]; po: string[]; }
// Sorts the workbook's tabs into asset exports, PO validation, user references and device references. Our own
// result sheets are skipped (so re-running on the same workbook works), except POs Validated, which echoes the
// user's PO data and is read back when there is no other PO tab.
function findTabs(wb: ExcelScript.Workbook, maps: Mappings): Tabs {
  const out: Tabs = { assets: [], users: [], devices: [], po: [] };
  let ownPo = "";
  wb.getWorksheets().forEach((w: ExcelScript.Worksheet) => {
    // Each Excel call is a round trip to Excel, so the name is asked for once
    const tab = w.getName();
    const own = SHEETS.indexOf(tab) >= 0 || tab === MAPPINGS_SHEET;
    if (own && tab !== PO_SHEET) { return; }
    const head = headRows(w);
    if (!head.length) { return; }
    const p = bestHeader(head, PO_COLS).pos;
    if (has(p, "number") && has(p, "quantity")) {
      if (own) { ownPo = tab; } else { out.po.push(tab); }
      return;
    }
    if (own) { return; }
    const name = tab.toLowerCase();
    const a = bestHeader(head, ASSET_COLS, maps.columns).pos, u = bestHeader(head, USER_COLS).pos, d = bestHeader(head, DEVICE_COLS).pos;
    const idCols = ["assetId", "tag", "serial"].filter(f => has(a, f)).length;
    const isAsset = idCols > 0 && (has(a, "status") || has(a, "cls")) && Object.keys(a).length >= 3;
    const isDevice = has(d, "lastSeen") && (has(d, "serial") || has(d, "name"));
    const isUser = (has(u, "email") || has(u, "empId")) && (has(u, "enabled") || has(u, "leaveDate"));
    // Device-management exports also have serial and model columns: an asset tab needs a status or an asset no. / tag
    const deviceLike = isDevice && !has(a, "assetId") && !has(a, "tag") && !has(a, "status");
    if (DEVICE_TAB.test(name) && isDevice && !(isAsset && !deviceLike && /vcams|asset|inventory|export/.test(name))) {
      out.devices.push(tab);
    } else if (USER_TAB.test(name) && isUser && !isAsset) {
      out.users.push(tab);
    } else if (isAsset && !deviceLike) {
      out.assets.push(tab);
    } else if (isDevice) {
      out.devices.push(tab);
    } else if (isUser) {
      out.users.push(tab);
    }
  });
  if (!out.po.length && ownPo) { out.po.push(ownPo); }
  return out;
}

// A date in a tab name: "2026-09", "2026-09-08", "8-9-2026" / "8.9.2026" (day first), "Sep 2026" / "September 2026"
function tabDate(name: string): Date | null {
  let m = name.match(/(\d{4})[-_. ](\d{1,2})(?:[-_. ](\d{1,2}))?/);
  if (m && +m[2] >= 1 && +m[2] <= 12) { return new Date(Date.UTC(+m[1], +m[2] - 1, +(m[3] || 1))); }
  m = name.match(/(\d{1,2})[-_.](\d{1,2})[-_.](\d{4})/);
  if (m) { return parseDate(`${m[1]}/${m[2]}/${m[3]}`); }
  m = name.match(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[ \-_]*(\d{4})\b/i);
  if (m) {
    return new Date(Date.UTC(+m[2], ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"].indexOf(m[1].toLowerCase()), 1));
  }
  return null;
}
// The day in a tab name ("10-14-2026 Export", "2026-10-14 Export"); null when the name has no day
function tabDay(name: string): Date | null {
  const m = name.match(/(\d{4})[-_. ](\d{1,2})[-_. ](\d{1,2})/);
  if (m && +m[2] >= 1 && +m[2] <= 12) { return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])); }
  return /(\d{1,2})[-_.](\d{1,2})[-_.](\d{4})/.test(name) ? tabDate(name) : null;
}
// The newest AssetUpdateDate (else creation date, else last scan) in a tab: the older export has older data
function latestDate(wb: ExcelScript.Workbook, name: string, maps: Mappings): number {
  const ws = wb.getWorksheet(name);
  if (!ws) { return 0; }
  const used = ws.getUsedRange(true);
  if (!used) { return 0; }
  const h = bestHeader(headRows(ws), ASSET_COLS, maps.columns);
  const f = ["updated", "created", "lastSeen", "statusDate"].find(x => has(h.pos, x));
  if (h.row < 0 || !f) { return 0; }
  const values = readBlock(ws, used.getRowIndex() + h.row + 1, used.getColumnIndex() + h.pos[f], used.getRowCount() - h.row - 1, 1);
  return values.reduce((mx: number, r) => { const d = parseDate(r[0]); return d && d.getTime() > mx ? d.getTime() : mx; }, 0);
}
// Which export tab is the previous month and which the current one
function splitExports(wb: ExcelScript.Workbook, tabs: string[], maps: Mappings): { cur: string[]; prev: string[]; prevPeriod: string; note: string } {
  if (tabs.length === 1) { return { cur: tabs, prev: [], prevPeriod: "", note: "One export tab" }; }
  const PREV = /\b(prev|previous|last|old|before|prior)\b/i, CUR = /\b(current|curr|this|latest|now)\b/i;
  const period = (n: string) => { const d = tabDate(n); return d ? monthOf(d) : ""; };
  const prev = tabs.filter(n => PREV.test(n)), cur = tabs.filter(n => !PREV.test(n));
  if (prev.length && cur.length && (cur.some(n => CUR.test(n)) || cur.length === 1)) {
    return { cur, prev, prevPeriod: period(prev[0]), note: "By tab name: 'previous' = " + prev.join(", ") };
  }
  if (tabs.length === 2) {
    const d = tabs.map(n => tabDate(n));
    if (d[0] && d[1] && d[0].getTime() !== d[1].getTime()) {
      const older = d[0].getTime() < d[1].getTime() ? 0 : 1;
      return { cur: [tabs[1 - older]], prev: [tabs[older]], prevPeriod: monthOf(d[older] as Date),
        note: `By the dates in the tab names: '${tabs[older]}' is the older export` };
    }
    const latest = tabs.map(n => latestDate(wb, n, maps));
    if (latest[0] && latest[1] && latest[0] !== latest[1]) {
      const older = latest[0] < latest[1] ? 0 : 1;
      return { cur: [tabs[1 - older]], prev: [tabs[older]], prevPeriod: period(tabs[older]),
        note: `By the newest update date in each tab: '${tabs[older]}' is older (${fmtDay(new Date(latest[older]))} vs ` +
          `${fmtDay(new Date(latest[1 - older]))}). Name the tabs '... Previous Export' / '... Current Export' to be sure.` };
    }
  }
  return { cur: tabs, prev: [], prevPeriod: "", note: `${tabs.length} export tabs combined as one export. To reconcile two months, ` +
    "name them '... Previous Export' and '... Current Export'." };
}

function blankAsset(): Asset {
  return {
    key: "", sheet: "", row: 0, assetId: "", tag: "", serial: "", serialN: "", hostname: "", hostN: "", mac: "", macN: "",
    clsRaw: "", assetClass: "", group: UNLISTED, kind: OTHER, makeRaw: "", make: "", model: "", statusRaw: "", status: "",
    standardStatus: false, stage: UNKNOWN, user: "", email: "", empId: "", costCenter: "", location: "", company: "", market: "",
    managedBy: "", cost: 0, po: "", functionUnit: "", comment: "", statusComment: "", scanAge: "", createdAge: "", purchase: null, warranty: null, verified: null,
    lastSeen: null, statusDate: null, deployDate: null, created: null, updated: null, badDates: [], flags: [], device: null,
    prev: null, isNew: false, deployedNow: false
  };
}
// Status, type, market and the other derived fields, from the raw values
// An export repeats the same few statuses, types, companies and makes on every row, so each distinct value is
// worked out once and remembered (this keeps 100,000+ rows fast)
const MEMO = {
  status: new Map<string, { label: string; standard: boolean; stage: string }>(),
  type: new Map<string, { label: string; group: string; kind: string }>(),
  market: new Map<string, string>(),
  make: new Map<string, string>(),
  date: new Map<string, Date | null>()
};
function clearMemo(): void {
  MEMO.status.clear(); MEMO.type.clear(); MEMO.market.clear(); MEMO.make.clear(); MEMO.date.clear();
}
function memoDate(v: Raw): Date | null {
  if (typeof v !== "string") { return parseDate(v); }
  if (MEMO.date.has(v)) { return MEMO.date.get(v) || null; }
  const d = parseDate(v);
  if (MEMO.date.size < 50000) { MEMO.date.set(v, d); }
  return d;
}
function classify(a: Asset, maps: Mappings): void {
  let st = MEMO.status.get(a.statusRaw);
  if (!st) { st = statusOf(a.statusRaw, maps); MEMO.status.set(a.statusRaw, st); }
  const tk = a.clsRaw + "\u0001" + a.model;
  let ty = MEMO.type.get(tk);
  if (!ty) { ty = typeOf(a.clsRaw, a.model, maps); MEMO.type.set(tk, ty); }
  a.status = st.label; a.standardStatus = st.standard; a.stage = st.stage;
  a.assetClass = ty.label; a.group = ty.group; a.kind = ty.kind;
  let market = MEMO.market.get(a.company);
  if (market === undefined) { market = marketOf(a.company, maps); MEMO.market.set(a.company, market); }
  a.market = market;
  let make = MEMO.make.get(a.makeRaw);
  if (make === undefined) { make = makeOf(a.makeRaw); MEMO.make.set(a.makeRaw, make); }
  a.make = make;
  a.serialN = isPlaceholder(a.serial) ? "" : serialKey(a.serial);
  a.hostN = isPlaceholder(a.hostname) ? "" : hostKey(a.hostname);
  a.macN = a.mac.toUpperCase().replace(/[^0-9A-F]/g, "").replace(/^0+$/, "");
  a.key = a.assetId || a.tag || a.serial || a.hostname;
}
function idKey(s: string): string {
  return (s || "").trim().toUpperCase();
}

const DATE_FIELDS = ["purchase", "warranty", "verified", "lastSeen", "statusDate", "deployDate", "created", "updated"];
function loadAssets(wb: ExcelScript.Workbook, sheets: string[], maps: Mappings): AssetLoad {
  const out: AssetLoad = { assets: [], rows: 0, unique: 0, blankIds: 0, dupIds: 0, crossTabDups: 0, cols: new Set<string>(),
    found: new Map<string, string>(), ambiguous: [], tabs: [] };
  const seenEarlier = new Set<string>();
  sheets.forEach(sheetName => {
    const ws = wb.getWorksheet(sheetName);
    if (!ws) { return; }
    const used = ws.getUsedRange(true);
    if (!used) { return; }
    const head = headRows(ws), h = bestHeader(head, ASSET_COLS, maps.columns);
    if (h.row < 0) { return; }
    const fields = Object.keys(h.pos);
    fields.forEach(f => {
      out.cols.add(f);
      if (!out.found.has(f)) { out.found.set(f, String(head[h.row][h.pos[f]]).trim()); }
    });
    h.amb.forEach(x => { if (out.ambiguous.indexOf(x) < 0) { out.ambiguous.push(x); } });
    // Read only the columns the audit uses: the span between the first and the last of them, or column by column
    // when that span is mostly other columns (wide exports)
    // Each Excel call is a round trip to Excel: ask for the used range's position once
    const uRow = used.getRowIndex(), uCol = used.getColumnIndex(), uRows = used.getRowCount(), uCols = used.getColumnCount();
    const top = uRow + h.row + 1, n = uRows - h.row - 1;
    const colAbs: Record<string, number> = {};
    fields.forEach(f => { colAbs[f] = uCol + h.pos[f]; });
    let c0 = Math.min(...fields.map(f => h.pos[f]));
    const c1 = Math.max(...fields.map(f => h.pos[f]));
    let values: Raw[][];
    progress(`'${sheetName}': ${n.toLocaleString("en-US")} rows x ${uCols} columns; reading the ${fields.length} columns the audit uses`);
    if (c1 - c0 + 1 > 1.5 * fields.length) {
      const pos: Record<string, number> = {};
      const cols = fields.map((f, i) => {
        pos[f] = i;
        return readBlock(ws, top, uCol + h.pos[f], n, 1, `'${sheetName}' column ${i + 1} of ${fields.length} (${STANDARD_HEADER[f] || f})`);
      });
      values = [];
      for (let i = 0; i < n; i++) { values.push(cols.map(c => c[i] ? c[i][0] : "")); }
      Object.keys(pos).forEach(f => { h.pos[f] = pos[f]; });
      c0 = 0;
    } else {
      values = readBlock(ws, top, uCol + c0, n, c1 - c0 + 1, `'${sheetName}'`);
    }
    out.tabs.push({ name: sheetName, headerAbs: top - 1, lastAbs: top + n - 1, colAbs, rows: n });
    const at = (r: Raw[], f: string): Raw => h.pos[f] === undefined ? "" : r[h.pos[f] - c0];
    const get = (r: Raw[], f: string) => { const v = at(r, f); return String(v === undefined || v === null ? "" : v).trim(); };
    const num = (r: Raw[], f: string): Cell => { const v = at(r, f); return typeof v === "number" ? v : get(r, f) && isFinite(Number(get(r, f))) ? Number(get(r, f)) : ""; };
    const seenHere = new Set<string>();
    progress(`'${sheetName}': building the asset records`);
    for (let i = 0; i < values.length; i++) {
      if (i > 0 && i % 20000 === 0) { progress(`'${sheetName}': building records ${i.toLocaleString("en-US")} of ${values.length.toLocaleString("en-US")}`); }
      const r = values[i];
      if (!r.some(v => String(v).trim())) { continue; }
      out.rows++;
      const a = blankAsset();
      a.assetId = get(r, "assetId"); a.tag = get(r, "tag"); a.serial = get(r, "serial"); a.hostname = get(r, "hostname");
      if (!a.assetId && !a.tag && !a.serial && !a.hostname) { out.blankIds++; continue; }
      if (sheets.length > 1) {
        // The same export split over tabs may repeat rows at the joins: an identical row on an earlier tab is counted once
        const sig = fields.map(f => String(at(r, f))).join("\u0001");
        if (seenEarlier.has(sig)) { out.crossTabDups++; out.rows--; continue; }
        seenHere.add(sig);
      }
      a.sheet = sheetName; a.row = top + i + 1;
      a.mac = get(r, "mac"); a.clsRaw = get(r, "cls"); a.model = get(r, "model"); a.statusRaw = get(r, "status");
      a.makeRaw = get(r, "make"); a.user = get(r, "user"); a.email = get(r, "email").toLowerCase(); a.empId = get(r, "empId");
      a.costCenter = get(r, "costCenter"); a.location = get(r, "location"); a.company = get(r, "company");
      a.managedBy = get(r, "managedBy"); a.po = get(r, "po"); a.functionUnit = get(r, "functionUnit"); a.comment = get(r, "comment"); a.statusComment = get(r, "statusComment");
      a.scanAge = num(r, "scanAge"); a.createdAge = num(r, "createdAge");
      const costText = get(r, "cost").replace(/[^0-9.\-]/g, "");
      a.cost = costText ? Number(costText) || 0 : 0;
      a.purchase = memoDate(at(r, "purchase")); a.warranty = memoDate(at(r, "warranty")); a.verified = memoDate(at(r, "verified"));
      a.lastSeen = memoDate(at(r, "lastSeen")); a.statusDate = memoDate(at(r, "statusDate")); a.deployDate = memoDate(at(r, "deployDate"));
      a.created = memoDate(at(r, "created")); a.updated = memoDate(at(r, "updated"));
      // A date column with text that is not a date is an invalid date, not a missing one
      DATE_FIELDS.forEach(f => {
        const v = at(r, f);
        const t = get(r, f);
        if (t && !memoDate(v) && !isPlaceholder(t)) { a.badDates.push(`${FIELD_LABEL[f]} '${short(t, 20)}' is not a date`); }
      });
      if (BLANK_USER.test(norm(a.user)) && !a.email && !a.empId) { a.user = ""; }
      if (isPlaceholder(a.email)) { a.email = ""; }
      if (!a.assetId && h.pos.assetId !== undefined) { out.blankIds++; }
      classify(a, maps);
      out.assets.push(a);
    }
    seenHere.forEach(s => seenEarlier.add(s));
    progress(`'${sheetName}': ${values.length.toLocaleString("en-US")} records built`);
  });
  const ids = countBy(out.assets.filter(a => a.assetId).map(a => idKey(a.assetId)));
  out.unique = ids.size;
  ids.forEach(n => { if (n > 1) { out.dupIds++; } });
  return out;
}

function activeValue(v: string): boolean | null {
  const s = norm(v);
  if (!s) { return null; }
  if (/^(true|yes|y|1|active|enabled|employed|current|on)$/.test(s)) { return true; }
  if (/^(false|no|n|0|inactive|disabled|terminated|left|leaver|withdrawn|resigned|exited|separated|off)$/.test(s)) { return false; }
  return null;
}
function addUser(users: Users, u: UserRef, asOf: Date): void {
  if (u.leaveDate && u.leaveDate.getTime() <= asOf.getTime()) { u.active = false; }
  // A person can have several rows (rehire, second account): active wins
  const put = (m: Map<string, UserRef>, k: string) => {
    if (!k) { return; }
    const old = m.get(k);
    if (!old || (!old.active && u.active)) { m.set(k, u); }
  };
  put(users.byEmail, u.email.toLowerCase());
  put(users.byEmp, u.empId.toLowerCase());
  users.count++;
  if (!u.active) { users.inactive++; }
}
// Users: compact JSON [[email, employeeId, active 1/0/"", leaveDate YYYY-MM-DD], ...] or a Users tab
function loadUsers(wb: ExcelScript.Workbook, sheets: string[], json: string | undefined, asOf: Date): Users {
  const users: Users = { source: "", byEmail: new Map<string, UserRef>(), byEmp: new Map<string, UserRef>(), count: 0, inactive: 0 };
  if (json) {
    const rows: string[][] = JSON.parse(json);
    users.source = "usersJson (central reference)";
    rows.forEach(r => addUser(users, { email: String(r[0] || ""), empId: String(r[1] || ""), name: "",
      active: String(r[2]) !== "0", leaveDate: parseDate(String(r[3] || "")) }, asOf));
    return users;
  }
  sheets.forEach(sheetName => {
    const ws = wb.getWorksheet(sheetName);
    if (!ws) { return; }
    const values = readValues(ws), h = bestHeader(values, USER_COLS);
    if (h.row < 0) { return; }
    users.source = (users.source ? users.source + ", " : "") + `tab '${sheetName}'`;
    const get = (r: Raw[], f: string) => h.pos[f] === undefined ? "" : String(r[h.pos[f]] === undefined ? "" : r[h.pos[f]]).trim();
    for (let i = h.row + 1; i < values.length; i++) {
      const r = values[i], email = get(r, "email"), empId = get(r, "empId");
      if (!email && !empId) { continue; }
      const act = activeValue(get(r, "enabled"));
      addUser(users, { email, empId, name: get(r, "name"), active: act !== false,
        leaveDate: h.pos.leaveDate === undefined ? null : parseDate(r[h.pos.leaveDate]) }, asOf);
    }
  });
  return users;
}
// Devices: compact JSON [[name, serial, lastSeen YYYY-MM-DD, user, market], ...] or an Intune / SCCM tab
function loadDevices(wb: ExcelScript.Workbook, sheets: string[], json: string | undefined, maps: Mappings): Devices {
  const out: Devices = { source: "", list: [], bySerial: new Map<string, DeviceRef>(), byName: new Map<string, DeviceRef>() };
  const add = (d: DeviceRef) => {
    if (!d.nameN && !d.serialN) { return; }
    out.list.push(d);
    // The same device can be listed twice (re-enrolled): keep the most recently seen
    const keep = (m: Map<string, DeviceRef>, k: string) => {
      if (!k) { return; }
      const old = m.get(k);
      if (!old || (d.lastSeen && (!old.lastSeen || d.lastSeen.getTime() > old.lastSeen.getTime()))) { m.set(k, d); }
    };
    keep(out.bySerial, d.serialN);
    keep(out.byName, d.nameN);
  };
  const ref = (name: string, serial: string, seen: Date | null, user: string, market: string): DeviceRef => ({
    name, nameN: isPlaceholder(name) ? "" : hostKey(name), serial, serialN: isPlaceholder(serial) ? "" : serialKey(serial),
    lastSeen: seen, user: user.toLowerCase(), market, matched: false
  });
  if (json) {
    const rows: string[][] = JSON.parse(json);
    out.source = "devicesJson (central reference)";
    rows.forEach(r => add(ref(String(r[0] || ""), String(r[1] || ""), parseDate(String(r[2] || "")), String(r[3] || ""), String(r[4] || ""))));
    return out;
  }
  sheets.forEach(sheetName => {
    const ws = wb.getWorksheet(sheetName);
    if (!ws) { return; }
    const values = readValues(ws), h = bestHeader(values, DEVICE_COLS);
    if (h.row < 0) { return; }
    out.source = (out.source ? out.source + ", " : "") + `tab '${sheetName}'`;
    const get = (r: Raw[], f: string) => h.pos[f] === undefined ? "" : String(r[h.pos[f]] === undefined ? "" : r[h.pos[f]]).trim();
    for (let i = h.row + 1; i < values.length; i++) {
      const r = values[i];
      add(ref(get(r, "name"), get(r, "serial"), h.pos.lastSeen === undefined ? null : parseDate(r[h.pos.lastSeen]),
        get(r, "user"), get(r, "company")));
    }
  });
  out.list.forEach(d => { if (d.market) { d.market = marketOf(d.market, maps); } });
  return out;
}


// PO validation rows supplied by the user (poJson, or a tab with Purchase Number and Quantity). Values are kept
// exactly as supplied; problems are reported, never fixed.
function loadPo(wb: ExcelScript.Workbook, sheets: string[], json: string | undefined): PoData {
  const out: PoData = { source: "", rows: [], problems: [] };
  let raw: Raw[][] = [];
  if (json) {
    const parsed: Record<string, string | number>[] = JSON.parse(json);
    const pick = (o: Record<string, string | number>, names: string[]): Raw => {
      const k = Object.keys(o).find(x => names.some(n => headerKey(n) === headerKey(x)));
      return k === undefined || o[k] === null ? "" : o[k];
    };
    raw = parsed.map(o => Object.keys(PO_COLS).map(f => pick(o, PO_COLS[f])));
    out.source = "the PO data supplied in the chat (poJson)";
  } else if (sheets.length && wb.getWorksheet(sheets[0])) {
    const ws = wb.getWorksheet(sheets[0]) as ExcelScript.Worksheet;
    const values = readValues(ws), h = bestHeader(values, PO_COLS);
    out.source = `tab '${sheets[0]}'`;
    // Rows up to the first empty row (the POs Validated sheet has the reconciliation section below)
    for (let i = h.row + 1; i < values.length && values[i].some(v => String(v).trim()); i++) {
      raw.push(Object.keys(PO_COLS).map(f => h.pos[f] === undefined ? "" : values[i][h.pos[f]]));
    }
  }
  const text = (v: Raw) => String(v === undefined || v === null ? "" : v).trim();
  raw.forEach((r, i) => {
    const [date, number, model, quantity, raisedBy, physical, spider] = r;
    const row: PoRow = { date: typeof date === "number" ? date : text(date), number: text(number), model: text(model),
      quantity: typeof quantity === "number" ? quantity : text(quantity), raisedBy: text(raisedBy), physical: text(physical),
      spider: text(spider), problems: [] };
    const q = Number(String(row.quantity).replace(/,/g, ""));
    if (!row.number) { row.problems.push("Purchase Number missing"); }
    if (row.quantity === "" || !isFinite(q) || q < 0) { row.problems.push(`Quantity '${row.quantity}' is not a number of 0 or more`); }
    else if (typeof row.quantity !== "number") { row.quantity = q; }
    [["Physical Security Validation", row.physical], ["Spider Validation", row.spider]].forEach(x => {
      if (!x[1]) { row.problems.push(`${x[0]} missing`); }
      else if (!PO_LABELS.some(l => l.toLowerCase() === x[1].toLowerCase())) { row.problems.push(`${x[0]} '${x[1]}' is not one of ${PO_LABELS.join(", ")}`); }
    });
    if (typeof row.date === "string" && row.date && parseDate(row.date)) { row.date = dateSerial(parseDate(row.date)); }
    if (!row.raisedBy) { row.problems.push("Raised By missing"); }
    row.problems.forEach(p => out.problems.push(`PO ${row.number || "(row " + (i + 1) + ")"}: ${p}`));
    out.rows.push(row);
  });
  return out;
}

function emptyMappings(): Mappings {
  return { status: new Map<string, string>(), types: new Map<string, [string, string]>(), market: [], limits: new Map<string, number>(),
    lifecycle: new Map<string, number>(), hostname: new Map<string, RegExp>(), columns: new Map<string, string>(), off: new Set<string>(),
    source: "None", applied: 0, problems: [] };
}
// Mappings rows {Kind, Value, MapsTo} from the central table (mappingsJson) or a 'Mappings' sheet in this workbook
function loadMappings(wb: ExcelScript.Workbook, json: string | undefined): Mappings {
  const maps = emptyMappings();
  let rows: string[][] = [];
  if (json) {
    const parsed: Record<string, string | number>[] = JSON.parse(json);
    const pick = (o: Record<string, string | number>, k: string) => {
      const key = Object.keys(o).find(x => x.toLowerCase().replace(/[^a-z]/g, "") === k);
      return key === undefined ? "" : String(o[key]).trim();
    };
    rows = parsed.map(o => [pick(o, "kind"), pick(o, "value"), pick(o, "mapsto")]);
    maps.source = "the central Mappings table";
  } else {
    const ws = wb.getWorksheet(MAPPINGS_SHEET);
    if (ws) {
      const values = readValues(ws);
      const hr = values.findIndex(r => r.map(v => headerKey(v)).indexOf("kind") >= 0 && r.map(v => headerKey(v)).indexOf("mapsto") >= 0);
      if (hr >= 0) {
        const hk = values[hr].map(v => headerKey(v));
        const ki = hk.indexOf("kind"), vi = hk.indexOf("value"), mi = hk.indexOf("mapsto");
        rows = values.slice(hr + 1).map(r => [String(r[ki]).trim(), String(r[vi]).trim(), String(r[mi]).trim()]);
        maps.source = `the '${MAPPINGS_SHEET}' sheet`;
      }
    }
  }
  rows.filter(r => r[0] || r[1]).forEach((r, i) => {
    const kind = norm(r[0]), value = r[1], to = r[2], line = `row ${i + 1} (${r[0]} / ${value})`;
    const num = Number(to), vk = compactKey(value);
    if (kind === "status") {
      const known = STATUS_INDEX.get(compactKey(to));
      const stage = STAGES.find(s => norm(s) === norm(to)) || (known ? known[1] : "");
      if (stage) { maps.status.set(vk, stage); maps.applied++; }
      else { maps.problems.push(`${line}: MapsTo must be a controlled status or one of ${STAGES.join(", ")}`); }
    } else if (kind === "class" || kind === "asset type" || kind === "kind") {
      const old = maps.types.get(vk) || ["", ""];
      if (to) { maps.types.set(vk, [old[0], to]); maps.applied++; } else { maps.problems.push(`${line}: no kind`); }
    } else if (kind === "group") {
      const g = VALUE_GROUPS.find(x => norm(x) === norm(to));
      const old = maps.types.get(vk) || ["", ""];
      if (g) { maps.types.set(vk, [g, old[1]]); maps.applied++; } else { maps.problems.push(`${line}: group must be ${VALUE_GROUPS.slice(0, 3).join(", ")}`); }
    } else if (kind === "column") {
      const f = ASSET_FIELDS.find(x => x.toLowerCase() === value.toLowerCase() || headerKey(STANDARD_HEADER[x]) === headerKey(value) ||
        ASSET_COLS[x].some(nm => headerKey(nm) === headerKey(value)));
      if (f && to) { maps.columns.set(f, headerKey(to)); maps.applied++; } else { maps.problems.push(`${line}: unknown field or no column name`); }
    } else if (kind === "market") {
      if (to) { maps.market.push(["^" + norm(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "$", to]); maps.applied++; }
    } else if (kind === "setting") {
      if (SETTINGS[value] !== undefined && isFinite(num) && to !== "") { SETTINGS[value] = num; maps.applied++; }
      else { maps.problems.push(`${line}: unknown setting or not a number`); }
    } else if (kind === "limit" || kind === "lifecycle") {
      if (isFinite(num) && num > 0) { (kind === "limit" ? maps.limits : maps.lifecycle).set(vk, num); maps.applied++; }
      else { maps.problems.push(`${line}: needs a number`); }
    } else if (kind === "hostnamepattern" || kind === "hostname pattern") {
      try { maps.hostname.set(value.toLowerCase() || "*", new RegExp(to, "i")); maps.applied++; }
      catch (e) { maps.problems.push(`${line}: not a valid pattern`); }
    } else if (kind === "check") {
      const code = value.toUpperCase();
      if (!FLAG_INFO[code]) { maps.problems.push(`${line}: unknown check`); }
      else if (/^(off|no|disable|disabled)$/i.test(to)) { maps.off.add(code); maps.applied++; }
      else if (SEVERITY.indexOf(to) >= 0) { FLAG_INFO[code].sev = to; maps.applied++; }
      else { maps.problems.push(`${line}: use Off or High / Medium / Low`); }
    } else {
      maps.problems.push(`${line}: unknown kind`);
    }
  });
  return maps;
}

// ---------------------------------------------------------------------------- snapshot (previous month without its file)
// A compact copy of the export (asset no., status, type, model, PO, market) that the flow stores and passes back
// next month as previousSnapshot. Values are dictionary-coded: about 25-35 characters per asset, so 100,000 assets take
// about 3 MB - stay under the ~5 MB request limit (run per market, or keep the previous export in the workbook, if larger).
function makeSnapshot(assets: Asset[], period: string): string {
  const dict: Record<string, string[]> = { s: [], t: [], m: [], p: [], k: [] };
  const idx = new Map<string, number>();
  const code = (d: string, v: string) => {
    if (!v) { return ""; }
    const k = d + "\u0001" + v;
    let i = idx.get(k);
    if (i === undefined) { i = dict[d].length; dict[d].push(v); idx.set(k, i); }
    return String(i);
  };
  const rows = assets.filter(a => a.assetId).map(a => [a.assetId.replace(/[|~]/g, " "), code("s", a.statusRaw), code("t", a.clsRaw),
    code("m", a.model), code("p", a.po), code("k", a.company)].join("|"));
  return JSON.stringify({ v: 1, period, count: rows.length, dict, rows: rows.join("~") });
}
function parseSnapshot(text: string, maps: Mappings): { period: string; assets: Asset[] } {
  const snap: { v: number; period: string; dict: Record<string, string[]>; rows: string } = JSON.parse(text);
  const look = (d: string, i: string) => i === "" || i === undefined ? "" : (snap.dict[d] || [])[Number(i)] || "";
  const assets = snap.rows ? snap.rows.split("~").map(line => {
    const p = line.split("|"), a = blankAsset();
    a.assetId = p[0]; a.statusRaw = look("s", p[1]); a.clsRaw = look("t", p[2]); a.model = look("m", p[3]); a.po = look("p", p[4]);
    a.company = look("k", p[5]); a.sheet = "snapshot " + snap.period;
    classify(a, maps);
    return a;
  }) : [];
  return { period: snap.period, assets };
}

// ---------------------------------------------------------------------------- anonymisation
// Deterministic tokens within this run, the same in both exports, so matching and counts still work. The mapping
// is never written anywhere.
function applyShareSafe(ctx: Ctx): void {
  const maps = new Map<string, Map<string, string>>();
  const tok = (kind: string, v: string, make: (n: number) => string) => {
    if (!v) { return ""; }
    let m = maps.get(kind);
    if (!m) { m = new Map<string, string>(); maps.set(kind, m); }
    let t = m.get(v);
    if (!t) { t = make(m.size + 1); m.set(v, t); }
    return t;
  };
  const pad = (n: number, w: number) => String(n).padStart(w, "0");
  const user = (k: string) => tok("user", k, n => `USER-${pad(n, 4)}`);
  const mac = (n: number) => pad(n, 12).replace(/(..)(?=.)/g, "$1:");
  const emailToken = new Map<string, string>();
  const prevAll: Asset[] = ctx.recon.removed.concat(ctx.recon.changes.map(c => c.prev));
  const all = ctx.assets.concat(prevAll);
  all.forEach(a => { const k = userKey(a); if (k && a.email) { emailToken.set(a.email, user(k) + "@example.invalid"); } });
  const scrub = (s: string) => s.replace(/[\w.+-]+@[\w-]+\.[\w.]+/g, m => emailToken.get(m.toLowerCase()) || "USER@example.invalid");
  ctx.manyDevices.forEach(g => { g.user = user(userKey(g.assets[0])); g.email = ""; });
  all.forEach(a => {
    const k = userKey(a), alias = k ? user(k) : "", label = userLabel(a);
    a.flags.forEach(f => {
      f.detail = f.code === "BAD_SERIAL" ? "Placeholder or too short" : f.code === "BAD_EMAIL" ? "Not a valid email" :
        f.code === "BAD_HOSTNAME" ? "" : scrub(label && alias ? f.detail.split(label).join(alias) : f.detail);
    });
    a.user = alias; a.email = alias ? alias + "@example.invalid" : ""; a.empId = "";
    a.assetId = tok("asset", idKey(a.assetId), n => `AST-${pad(n, 6)}`);
    a.serial = tok("serial", a.serial, n => `SN-${pad(n, 6)}`);
    a.tag = tok("tag", a.tag, n => `TAG-${pad(n, 6)}`);
    a.hostname = tok("host", a.hostname.toLowerCase(), n => `HOST-${pad(n, 6)}`);
    a.mac = tok("mac", a.macN, mac);
    a.po = tok("po", a.po, n => `PO-${pad(n, 6)}`);
    a.functionUnit = tok("unit", a.functionUnit, n => `UNIT-${pad(n, 3)}`);
    a.location = tok("site", a.location, n => `SITE-${pad(n, 3)}`);
    a.comment = ""; a.statusComment = "";  // free text can hold names
    a.costCenter = tok("cc", a.costCenter, n => `CC-${pad(n, 3)}`);
    a.managedBy = tok("team", a.managedBy, n => `TEAM-${pad(n, 3)}`);
    a.company = a.market;
    a.key = a.assetId || a.tag || a.serial || a.hostname;
  });
  // Duplicate details named the other records: rebuild them with the tokens
  ctx.dupGroups.forEach(g => {
    const code = DUP_CODES[g.on];
    g.value = g.on === "Asset no." ? g.assets[0].assetId : g.on === "Serial" ? g.assets[0].serial : g.on === "Asset tag" ? g.assets[0].tag :
      g.on === "Hostname" ? g.assets[0].hostname : g.assets[0].mac;
    g.assets.forEach(a => a.flags.filter(f => f.code === code).forEach(f => {
      f.detail = `${g.on} ${g.value} also on ` + g.assets.filter(x => x !== a).slice(0, 3).map(x => x.key).join(", ");
    }));
  });
  ctx.devices.list.forEach(d => {
    d.user = d.user ? (emailToken.get(d.user) || "USER@example.invalid") : "";
    d.name = tok("host", d.name.toLowerCase(), n => `HOST-${pad(n, 6)}`);
    d.serial = tok("serial", d.serial, n => `SN-${pad(n, 6)}`);
  });
  ctx.po.rows.forEach(r => {
    r.number = tok("po", r.number, n => `PO-${pad(n, 6)}`);
    r.raisedBy = r.raisedBy ? user("name:" + norm(r.raisedBy)) : "";
    r.problems = r.problems.map(p => p.replace(/^PO \S+/, "PO"));
  });
  ctx.po.problems = ctx.po.problems.map(p => p.replace(/^PO \S+:/, "PO:"));
  const names: Record<string, string> = { user: "people (names, emails, employee IDs)", asset: "asset numbers", serial: "serial numbers",
    tag: "asset tags", host: "hostnames", mac: "MAC addresses", po: "purchase numbers", unit: "function units", site: "locations",
    cc: "cost centers", team: "teams" };
  ctx.tokens = Array.from(maps.keys()).map(k => `${names[k] || k} (${(maps.get(k) as Map<string, string>).size})`);
}
const DUP_CODES: Record<string, string> = { "Asset no.": "DUP_ASSET_ID", "Serial": "DUP_SERIAL", "Asset tag": "DUP_TAG",
  "Hostname": "DUP_HOSTNAME", "MAC": "DUP_MAC" };

// ============================================================================ monthly reconciliation
function emptyRecon(): Recon {
  return { available: false, source: "", prevPeriod: "", curPeriod: "", prevDay: null, curDay: null, prevRows: 0, curRows: 0, prevUnique: 0, curUnique: 0,
    prevBlank: 0, curBlank: 0, prevDup: [], curDup: [], serialChanges: [], serialChecked: false, prevByType: new Map<string, number>(), curByType: new Map<string, number>(),
    newAssets: [], removed: [], existing: 0, changes: [], operational: [], unknownStatuses: new Map<string, number>(),
    unlistedTypes: new Map<string, number>(), fromSnapshot: false };
}
const SERIAL_ADDED = "Serial added (was blank or a placeholder)", SERIAL_REMOVED = "Serial removed (now blank or a placeholder)",
  SERIAL_CHANGED = "Serial changed";
// Serial as compared month to month: case, spaces and punctuation ignored; placeholders ("N/A", "0") count as blank
function serialCompare(s: string): string {
  return isPlaceholder(s) ? "" : serialKey(s);
}
// Previous vs current export on Asset no. (trimmed, case-insensitive). Never by row, serial, model or PO.
function reconcile(ctx: Ctx, cur: AssetLoad, prevLoad: AssetLoad | null, prevAssets: Asset[], curAssets: Asset[],
  source: string, prevPeriod: string): Recon {
  const rc = emptyRecon();
  rc.available = true; rc.source = source; rc.prevPeriod = prevPeriod; rc.curPeriod = ctx.period; rc.fromSnapshot = !prevLoad;
  rc.serialChecked = !!prevLoad && prevLoad.cols.has("serial") && cur.cols.has("serial");
  const index = (list: Asset[]) => {
    const m = new Map<string, Asset>();
    list.forEach(a => { const k = idKey(a.assetId); if (k && !m.has(k)) { m.set(k, a); } });
    return m;
  };
  const dups = (list: Asset[]) => {
    const c = countBy(list.filter(a => a.assetId).map(a => idKey(a.assetId)));
    return Array.from(c.keys()).filter(k => (c.get(k) || 0) > 1);
  };
  const prevMap = index(prevAssets), curMap = index(curAssets);
  rc.prevRows = prevAssets.length; rc.curRows = curAssets.length;
  rc.prevUnique = prevMap.size; rc.curUnique = curMap.size;
  rc.prevBlank = prevAssets.filter(a => !a.assetId).length; rc.curBlank = curAssets.filter(a => !a.assetId).length;
  rc.prevDup = dups(prevAssets); rc.curDup = dups(curAssets);
  rc.prevByType = countBy(prevAssets.map(a => a.assetClass));
  rc.curByType = countBy(curAssets.map(a => a.assetClass));
  curMap.forEach((a, k) => {
    const p = prevMap.get(k);
    if (!p) { a.isNew = true; rc.newAssets.push(a); return; }
    rc.existing++;
    a.prev = p;
    // Status compared after trimming and standard capitalisation, so " Reserved " = "Reserved"
    if (compactKey(p.status) !== compactKey(a.status)) {
      const ch: StatusChange = { cur: a, prev: p, transition: `${p.status || "(blank)"} → ${a.status || "(blank)"}` };
      rc.changes.push(ch);
      if (OPERATIONAL_STATUSES.some(s => compactKey(s) === compactKey(a.status))) { rc.operational.push(ch); }
    }
    // Serial number changes (the snapshot has no serials, so only with the previous export tab)
    if (rc.serialChecked && serialCompare(p.serial) !== serialCompare(a.serial)) {
      rc.serialChanges.push({ cur: a, prev: p, change: !serialCompare(p.serial) ? SERIAL_ADDED : !serialCompare(a.serial) ? SERIAL_REMOVED : SERIAL_CHANGED });
    }
    // Deployed this month: moved into use from any other stage
    if (a.stage === IN_USE && p.stage !== IN_USE) { a.deployedNow = true; }
  });
  rc.newAssets.forEach(a => { if (a.stage === IN_USE) { a.deployedNow = true; } });
  prevMap.forEach((p, k) => { if (!curMap.has(k)) { rc.removed.push(p); } });
  prevAssets.concat(curAssets).filter(a => !a.standardStatus).forEach(a => rc.unknownStatuses.set(a.status || "(blank)",
    (rc.unknownStatuses.get(a.status || "(blank)") || 0) + 1));
  prevAssets.concat(curAssets).filter(a => a.group === UNLISTED).forEach(a => rc.unlistedTypes.set(a.assetClass,
    (rc.unlistedTypes.get(a.assetClass) || 0) + 1));
  // Office Scripts only accept inline arrow functions as array callbacks (not a named comparator)
  rc.operational.sort((x, y) => x.cur.status.localeCompare(y.cur.status) || x.cur.assetClass.localeCompare(y.cur.assetClass) ||
    x.cur.assetId.localeCompare(y.cur.assetId));
  rc.serialChanges.sort((x, y) => x.change.localeCompare(y.change) || x.cur.assetClass.localeCompare(y.cur.assetClass) ||
    x.cur.assetId.localeCompare(y.cur.assetId));
  return rc;
}
// [check, expected, actual, result] - the controls the monthly process requires before reporting completion
function qualityChecks(ctx: Ctx): Cell[][] {
  const rc = ctx.recon, out: Cell[][] = [];
  const check = (name: string, expected: Cell, actual: Cell, ok: boolean, review?: boolean) =>
    out.push([name, expected, actual, ok ? "Pass" : review ? "Review" : "Fail"]);
  if (!rc.available) { return out; }
  check("Previous export: rows without an Asset no.", 0, rc.prevBlank, rc.prevBlank === 0, true);
  check("Previous export: asset numbers on more than one row", 0, rc.prevDup.length, rc.prevDup.length === 0, true);
  check("Current export: rows without an Asset no.", 0, rc.curBlank, rc.curBlank === 0, true);
  check("Current export: asset numbers on more than one row", 0, rc.curDup.length, rc.curDup.length === 0, true);
  check("Current unique assets = existing + new", rc.curUnique, rc.existing + rc.newAssets.length, rc.curUnique === rc.existing + rc.newAssets.length);
  check("Previous unique assets = existing + removed", rc.prevUnique, rc.existing + rc.removed.length, rc.prevUnique === rc.existing + rc.removed.length);
  const prevIds = new Set<string>(rc.removed.map(a => idKey(a.assetId)));
  const wrongNew = rc.newAssets.filter(a => a.prev !== null).length;
  check("No existing asset listed as new", 0, wrongNew, wrongNew === 0);
  check("No removed asset is in the current export", 0, ctx.assets.filter(a => prevIds.has(idKey(a.assetId))).length,
    !ctx.assets.some(a => prevIds.has(idKey(a.assetId))));
  const badOps = rc.operational.filter(c => !c.prev || compactKey(c.prev.status) === compactKey(c.cur.status) ||
    !OPERATIONAL_STATUSES.some(s => compactKey(s) === compactKey(c.cur.status))).length;
  check("Every Scrap & Inactive row: in both exports, a real change, now Scrapped or Lost", 0, badOps, badOps === 0);
  if (rc.serialChecked) {
    const ids = new Set<string>(rc.serialChanges.map(c => idKey(c.cur.assetId)));
    const bad = rc.serialChanges.filter(c => !c.prev || serialCompare(c.prev.serial) === serialCompare(c.cur.serial)).length;
    check("Every Serial Number Changes row: in both exports, a real serial change", 0, bad, bad === 0);
    check("Serial number changes listed once per asset", rc.serialChanges.length, ids.size, ids.size === rc.serialChanges.length);
    const blanked = rc.serialChanges.filter(c => c.change === SERIAL_REMOVED).length;
    check("Serial numbers removed (now blank or a placeholder)", 0, blanked, blanked === 0, true);
  } else {
    check("Serial number changes checked", "Yes", rc.fromSnapshot ? "No - the snapshot has no serials" : "No - Serial no. column missing", false, true);
  }
  check("Operational + excluded status changes = all status changes", rc.changes.length,
    rc.operational.length + (rc.changes.length - rc.operational.length), true);
  const byPo = countBy(rc.newAssets.map(a => a.po || ""));
  const poTotal = Array.from(byPo.values()).reduce((s, n) => s + n, 0);
  check("New assets by PurchaseNo (incl. blank) add up to the new assets", rc.newAssets.length, poTotal, poTotal === rc.newAssets.length);
  check("New assets with a blank PurchaseNo", 0, byPo.get("") || 0, !(byPo.get("") || 0), true);
  const sumType = (m: Map<string, number>) => Array.from(m.values()).reduce((s, n) => s + n, 0);
  check("Summary: previous total by type = previous rows", rc.prevRows, sumType(rc.prevByType), sumType(rc.prevByType) === rc.prevRows);
  check("Summary: current total by type = current rows", rc.curRows, sumType(rc.curByType), sumType(rc.curByType) === rc.curRows);
  check("Statuses outside the controlled list (both exports)", 0, Array.from(rc.unknownStatuses.values()).reduce((s, n) => s + n, 0),
    rc.unknownStatuses.size === 0, true);
  check("Asset types outside the standard list (both exports)", 0, Array.from(rc.unlistedTypes.values()).reduce((s, n) => s + n, 0),
    rc.unlistedTypes.size === 0, true);
  const po = poReconRows(ctx);
  check("PO validation rows supplied by the user", "1 or more", ctx.po.rows.length, ctx.po.rows.length > 0, true);
  check("PO validation rows with missing or non-standard values", 0, ctx.po.problems.length, ctx.po.problems.length === 0, true);
  const open = po.filter(r => r[7] !== PO_RESULTS.matched).length;
  check("PO reconciliation lines not 'Matched'", 0, open, open === 0, true);
  return out;
}
// PO reconciliation: [Purchase Number, Model, Export Asset Count, User-Provided Quantity, Quantity Variance,
//                     Physical Security Validation, Spider Validation, Reconciliation Result]
function poReconRows(ctx: Ctx): Cell[][] {
  const rows: Cell[][] = [];
  const byPo = new Map<string, Asset[]>();
  ctx.recon.newAssets.forEach(a => { const k = a.po.trim(), l = byPo.get(k); if (l) { l.push(a); } else { byPo.set(k, [a]); } });
  const supplied = new Map<string, PoRow[]>();
  ctx.po.rows.forEach(r => { if (r.number) { supplied.set(r.number.trim(), (supplied.get(r.number.trim()) || []).concat([r])); } });
  const label = (v: string) => PO_LABELS.find(l => l.toLowerCase() === v.toLowerCase()) || v;
  supplied.forEach((list, po) => {
    const found = byPo.get(po) || [];
    const qty = list.reduce((s, r) => s + (typeof r.quantity === "number" ? r.quantity : 0), 0);
    const val = (f: (r: PoRow) => string) => Array.from(new Set<string>(list.map(r => f(r)))).join(" / ");
    const phys = val(r => r.physical), spid = val(r => r.spider);
    const labels = [phys, spid].map(v => label(v));
    let result = PO_RESULTS.matched;
    if (!found.length) { result = PO_RESULTS.notFound; }
    else if (!phys || !spid) { result = PO_RESULTS.missing; }
    else if (labels.some(l => l !== "Yes" && l !== "Not Applicable")) { result = PO_RESULTS.pending; }
    else if (qty !== found.length) { result = PO_RESULTS.variance; }
    rows.push([po, val(r => r.model) || topList(found.map(a => a.model), 2), found.length, qty, qty - found.length, phys, spid, result]);
  });
  byPo.forEach((list, po) => {
    if (supplied.has(po)) { return; }
    rows.push(po ? [po, topList(list.map(a => a.model), 2), list.length, "", "", "", "", PO_RESULTS.missing]
      : ["(empty)", topList(list.map(a => a.model), 2), list.length, "", "", "", "", PO_RESULTS.blank]);
  });
  return rows.sort((a, b) => (a[0] === "(empty)" ? 1 : 0) - (b[0] === "(empty)" ? 1 : 0) || String(a[0]).localeCompare(String(b[0])));
}
// The completion report lines (also returned to the flow)
function reconText(ctx: Ctx): string[] {
  const rc = ctx.recon, po = poReconRows(ctx);
  const scrapped = rc.operational.filter(c => compactKey(c.cur.status) === "scrapped").length;
  const lost = rc.operational.filter(c => compactKey(c.cur.status) === "lost").length;
  const blankPo = rc.newAssets.filter(a => !a.po).length;
  return [
    `Previous-month asset total (${monthName(rc.prevPeriod)}): ${rc.prevRows.toLocaleString("en-US")}`,
    `Current-month asset total (${monthName(rc.curPeriod)}): ${rc.curRows.toLocaleString("en-US")}`,
    `Net movement: ${signed(rc.curRows - rc.prevRows)}`,
    `New assets: ${rc.newAssets.length}` + (rc.newAssets.length ? ` - by type: ${topList(rc.newAssets.map(a => a.assetClass), 10)}` : ""),
    `New assets by PO: ${topList(rc.newAssets.filter(a => a.po).map(a => a.po), 10) || "none"}`,
    `New assets with blank PurchaseNo: ${blankPo}`,
    `Removed assets (in the previous export, not in the current one - review, not automatically scrapped): ${rc.removed.length}`,
    `Total status changes: ${rc.changes.length}`,
    `Scrapped (new this month): ${scrapped}; Lost (new this month): ${lost}`,
    `Other status changes excluded from the operational table: ${rc.changes.length - rc.operational.length}`,
    rc.serialChecked ? `Serial number changes: ${rc.serialChanges.length}` + (rc.serialChanges.length ? ` (${mapList(countBy(rc.serialChanges.map(c => c.change)), 3)}; ` +
      `by type: ${mapList(countBy(rc.serialChanges.map(c => c.cur.assetClass)), 5)})` : "") :
      "Serial number changes: not checked (" + (rc.fromSnapshot ? "previous month came from the snapshot, which has no serials" : "no Serial no. column") + ")",
    `PO validation records processed: ${ctx.po.rows.length}` + (ctx.po.source ? ` (${ctx.po.source})` : " (none supplied)"),
    "PO export count vs supplied quantity: " + (po.length ? po.map(r => `${r[0] || "(blank PO)"}: export ${r[2]}, supplied ${r[3] === "" ? "-" : r[3]}, ` +
      `variance ${r[4] === "" ? "-" : signed(r[4] as number)} - ${r[7]}`).join("; ") : "no new-asset POs"),
    `Missing or pending PO validations: ${openValidations(ctx).length}` + (openValidations(ctx).length ? ` (${openValidations(ctx).join("; ")})` : ""),
    `Duplicate asset numbers: previous ${rc.prevDup.length}, current ${rc.curDup.length}; rows without an asset number: previous ${rc.prevBlank}, current ${rc.curBlank}`,
    `Statuses outside the controlled list: ${mapList(rc.unknownStatuses, 10) || "none"}`,
    `Asset types outside the standard list: ${mapList(rc.unlistedTypes, 10) || "none"}`
  ];
}
// POs whose validation is not complete: a supplied row with a blank, Pending or No validation, or a new-asset PO with
// no supplied row
function openValidations(ctx: Ctx): string[] {
  const ok = (v: string) => /^(yes|not applicable)$/i.test(v.trim());
  const out = ctx.po.rows.filter(r => !ok(r.physical) || !ok(r.spider))
    .map(r => `${r.number || "(no PO number)"}: Physical Security ${r.physical || "blank"}, Spider ${r.spider || "blank"}`);
  const supplied = new Set<string>(ctx.po.rows.map(r => r.number.trim()));
  Array.from(new Set<string>(ctx.recon.newAssets.map(a => a.po.trim()).filter(p => p && !supplied.has(p))))
    .forEach(p => out.push(`${p}: no validation supplied`));
  return out;
}
function mapList(m: Map<string, number>, n: number): string {
  return keysByCount(m).slice(0, n).map(k => `${k} (${m.get(k)})`).join("; ");
}
function signed(n: number): string {
  return (n > 0 ? "+" : "") + n.toLocaleString("en-US");
}

// ============================================================================ monthly email and PowerPoint prompt
// The monthly email holds only what the team already sends: totals, the summary by type, new / removed assets, serial
// number changes, Scrapped / Lost, PO validation and the points to review. No names or emails.
interface MonthlyReport { subject: string; attachment: string; email: string[]; heads: number[]; ppt: string[]; }
// Status-change and serial groups up to this size are listed asset by asset (larger ones as a count)
const EMAIL_DETAIL_MAX = 10;
// Where the team adds what the investigation found for each listed asset
const EMAIL_NOTE = "[add validation note]";
const COPILOT_PROMPT_LIMIT = 2000;
function fmtNum(n: number): string {
  return n.toLocaleString("en-US");
}
function countText(m: Map<string, number>, n: number): string {
  const keys = keysByCount(m);
  return keys.slice(0, n).map(k => `${k} ${fmtNum(m.get(k) || 0)}`).join(", ") + (keys.length > n ? `, and ${keys.length - n} more` : "");
}
// Asset types of one value group, in the standard order: [type, previous, current]
function typeRows(rc: Recon, group: string, maps: Mappings): [string, number, number][] {
  const types = VCAMS_TYPES.filter(t => t[1] === group).map(t => t[0]);
  if (group === UNLISTED) {
    Array.from(rc.prevByType.keys()).concat(Array.from(rc.curByType.keys())).forEach(t => {
      if (types.indexOf(t) < 0 && !TYPE_INDEX.has(compactKey(t)) && typeOf(t, "", maps).group === UNLISTED) { types.push(t); }
    });
  }
  return types.map(t => [t, rc.prevByType.get(t) || 0, rc.curByType.get(t) || 0] as [string, number, number]);
}
// One line per PO for the email: what the PO contains (user-supplied), what is already on Spider (new assets with that
// PO), and what is still pending. Supplied models are matched to asset types ("Bags" = Bag, "Mouses" / "Mice" = Mouse).
// Stock utilization for the email: Notebooks and Headsets (Headset type + headset models under Zubehör-Client)
function stockLines(ctx: Ctx): string[] {
  const p = (x: StockRow) => utilOf(x) === "" ? "-" : (100 * (utilOf(x) as number)).toFixed(1) + "%";
  return stockSummary(ctx).map((x, i) => (i < 2 ? "• " : "      o ") + `${x.label.trim()}: ${fmtNum(x.use)} in use, ${fmtNum(x.stock)} in stock - ${p(x)} utilized` +
    (ctx.recon.available ? `; ${fmtNum(x.deployed)} deployed this month` + (x.deployed ? `, stock covers about ${(x.stock / x.deployed).toFixed(1)} months` : "") : ""));
}
// The Spider comments of an asset for the email: the status comment first, then the general comment if it differs
function spiderComment(a: Asset): string {
  const sc = a.statusComment.trim(), c = a.comment.trim();
  return (sc ? ` - Spider Comment: ${short(sc, 200)}` : "") + (c && c !== sc ? ` - ${sc ? "Comment" : "Spider Comment"}: ${short(c, 200)}` : "");
}
function poLines(ctx: Ctx): string[] {
  const rc = ctx.recon, out: string[] = [];
  const typeKey = (model: string) => {
    const k = compactKey(model), all = rc.newAssets.map(a => a.assetClass);
    const hit = all.find(t => [k, k.replace(/s$/, ""), k.replace(/es$/, ""), k === "mice" ? "mouse" : ""].indexOf(compactKey(t)) >= 0);
    return hit || "";
  };
  const ok = (v: string) => /^(yes|not applicable)$/i.test(v.trim());
  const supplied = new Map<string, PoRow[]>();
  ctx.po.rows.forEach(r => { const k = r.number.trim(); if (k) { supplied.set(k, (supplied.get(k) || []).concat([r])); } });
  supplied.forEach((rows, po) => {
    const onSpider = rc.newAssets.filter(a => a.po.trim() === po);
    const added: string[] = [], pending: string[] = [], extra: string[] = [];
    const matched = new Set<string>();
    rows.forEach(r => {
      const qty = typeof r.quantity === "number" ? r.quantity : 0, t = typeKey(r.model), label = (r.model || "items").toLowerCase();
      const n = t ? onSpider.filter(a => a.assetClass === t).length : 0;
      if (t) { matched.add(t); }
      if (n >= qty && qty > 0) { added.push(`${fmtNum(qty)} ${label}`); if (n > qty) { extra.push(`${fmtNum(n - qty)} more ${label} on Spider than on the PO`); } }
      else if (n > 0) { added.push(`${fmtNum(n)} of ${fmtNum(qty)} ${label}`); pending.push(`${fmtNum(qty - n)} ${label}`); }
      else { pending.push(`${fmtNum(qty)} ${label}`); }
    });
    const other = onSpider.filter(a => !matched.has(a.assetClass));
    if (other.length) { extra.push(`${fmtNum(other.length)} assets on Spider with this PO not in the supplied lines (${countText(countBy(other.map(a => a.assetClass)), 3)})`); }
    let t = `PO: ${po} - ` + (added.length ? `Contains ${added.join(" and ")} which are already added` +
      (pending.length ? `, however there are ${pending.join(" and ")} still pending to be added on Spider` : "") :
      `Contains ${pending.join(" and ")}, all still pending to be added on Spider`) + ".";
    if (extra.length) { t += " " + extra.join("; ") + "."; }
    const open = rows.filter(r => !ok(r.physical) || !ok(r.spider));
    if (open.length) { t += ` Validation: Physical Security ${Array.from(new Set<string>(open.map(r => r.physical || "blank"))).join(" / ")}, ` +
      `Spider ${Array.from(new Set<string>(open.map(r => r.spider || "blank"))).join(" / ")}.`; }
    out.push(t);
  });
  const unsupplied = countBy(rc.newAssets.filter(a => a.po.trim() && !supplied.has(a.po.trim())).map(a => a.po.trim()));
  keysByCount(unsupplied).forEach(po => out.push(`PO: ${po} - ${fmtNum(unsupplied.get(po) || 0)} assets added on Spider (` +
    countText(countBy(rc.newAssets.filter(a => a.po.trim() === po).map(a => a.assetClass)), 3) + "); no PO validation supplied."));
  if (!out.length) { out.push("No POs to validate this month."); }
  return out;
}
function monthlyReport(ctx: Ctx): MonthlyReport {
  const rc = ctx.recon, prevName = monthName(rc.prevPeriod), curName = monthName(rc.curPeriod);
  const email: string[] = [], heads: number[] = [];
  const head = (t: string) => { heads.push(email.length); email.push(t); };
  const line = (t: string) => email.push(t);
  const scrapped = rc.operational.filter(c => compactKey(c.cur.status) === "scrapped");
  const lost = rc.operational.filter(c => compactKey(c.cur.status) === "lost");
  const blankPo = rc.newAssets.filter(a => !a.po).length;
  const byPo = countBy(rc.newAssets.filter(a => a.po).map(a => a.po));
  const po = poReconRows(ctx), open = openValidations(ctx);
  const groups = VALUE_GROUPS.map(g => {
    const rows = typeRows(rc, g, ctx.maps);
    return { g, rows, p: rows.reduce((s, r) => s + r[1], 0), c: rows.reduce((s, r) => s + r[2], 0) };
  }).filter(x => x.g !== UNLISTED || x.p || x.c);
  const review: string[] = [];
  if (rc.removed.length) { review.push(`${fmtNum(rc.removed.length)} removed assets to review (they are not treated as scrapped)`); }
  if (blankPo) { review.push(`${fmtNum(blankPo)} new assets with a blank PurchaseNo`); }
  if (open.length) { review.push(`PO validations missing or pending: ${open.length}`); }
  po.filter(r => r[7] === PO_RESULTS.variance).forEach(r => review.push(`PO ${r[0]}: quantity variance ${signed(r[4] as number)}`));
  if (rc.prevDup.length || rc.curDup.length) { review.push(`Duplicate asset numbers: ${rc.prevDup.length} last month, ${rc.curDup.length} this month`); }
  if (rc.prevBlank || rc.curBlank) { review.push(`Rows without an asset number: ${rc.prevBlank} last month, ${rc.curBlank} this month`); }
  if (rc.unknownStatuses.size) { review.push(`Statuses outside the standard list (both exports): ${countText(rc.unknownStatuses, 5)}`); }
  if (rc.unlistedTypes.size) { review.push(`Asset types outside the standard list (both exports): ${countText(rc.unlistedTypes, 5)}`); }
  const subject = "Asset Governance Task - Spider Data Integrity";
  const day = (d: Date | null) => d ? `${d.getUTCDate()} ${["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][d.getUTCMonth()]} ${d.getUTCFullYear()}` : "";
  const plural = (t: string) => /s$|-/i.test(t) ? t : t + "s";
  const typeCounts = (list: Asset[]) => keysByCount(countBy(list.map(a => a.assetClass)));
  const L1 = "• ", L2 = "      o ", L3 = "            ▪ ", L4 = "                  • ";

  line("Hi All,");
  line("");
  line("Please find attached excel sheet and below findings from Spider extraction analysis from " +
    (rc.prevDay && rc.curDay ? `${day(rc.prevDay)} to ${day(rc.curDay)}` : `${prevName} to ${curName}`) + ".");
  line("");
  line("Excel sheet have pivot table and summary with all the details.");
  line("");
  head("1. Overall Summary");
  line(`${L1}Previous-month total assets: ${fmtNum(rc.prevRows)}`);
  line(`${L1}Current-month total assets: ${fmtNum(rc.curRows)}`);
  const newTypes = typeCounts(rc.newAssets);
  line(`${L1}New assets: ${fmtNum(rc.newAssets.length)}` + (rc.newAssets.length ? ` (${newTypes.length === 1 ? plural(newTypes[0]) :
    countText(countBy(rc.newAssets.map(a => plural(a.assetClass))), 5)})` : ""));
  if (rc.removed.length) { line(`${L1}Removed assets (in the previous extraction, not in the current one - to review): ${fmtNum(rc.removed.length)}`); }
  line("");
  head("2. New Assets Validation");
  if (!rc.newAssets.length) { line(`${L1}No new assets`); }
  newTypes.forEach(t => {
    const list = rc.newAssets.filter(a => a.assetClass === t), byPoT = countBy(list.map(a => a.po));
    line(`${L1}${plural(t)}: ${fmtNum(list.length)}`);
    keysByCount(byPoT).filter(k => k).slice(0, 8).forEach(k => line(`${L2}${fmtNum(byPoT.get(k) || 0)} ${plural(t).toLowerCase()} with PO ${k}`));
    if (byPoT.get("")) { line(`${L2}${fmtNum(byPoT.get("") || 0)} ${plural(t)} with Empty PO on Spider`); }
  });
  line("");
  head("2.1. P.Os Validation");
  poLines(ctx).forEach(t => line(`${L1}${t}`));
  line("");
  head("3. Serial number validation:");
  if (!rc.serialChecked) {
    line(`${L1}Not checked (` + (rc.fromSnapshot ? "the previous month came from the stored snapshot, which has no serial numbers" : "no Serial no. column") + ")");
  } else if (!rc.serialChanges.length) {
    line(`${L1}No Changes`);
  } else {
    [SERIAL_CHANGED, SERIAL_ADDED, SERIAL_REMOVED].forEach(kind => {
      const list = rc.serialChanges.filter(c => c.change === kind);
      if (!list.length) { return; }
      line(`${L1}${kind}: ${fmtNum(list.length)}`);
      keysByCount(countBy(list.map(c => c.cur.assetClass))).forEach(t => {
        const ofType = list.filter(c => c.cur.assetClass === t);
        line(`${L2}${plural(t)}: ${fmtNum(ofType.length)}`);
        if (list.length <= EMAIL_DETAIL_MAX) {
          ofType.forEach(c => line(`${L3}Asset no: ${c.cur.assetId} - Serial number: ${c.prev.serial || "(blank)"} → ${c.cur.serial || "(blank)"}`));
        }
      });
    });
  }
  line("");
  // Status changes this month by the status they ended in (only real changes; spelling differences are not changes)
  const endedIn = (test: (st: string) => boolean) => rc.changes.filter(c => test(compactKey(c.cur.status)));
  const groupsByStatus: [string, StatusChange[]][] = [
    ["Scrapped", endedIn(k => k === "scrapped")],
    ["Inactive", endedIn(k => k === "inactive")],
    ["Lost", endedIn(k => k === "lost")],
    ["Stolen", endedIn(k => k === "stolen")]
  ];
  head("4. Status changes validation:");
  groupsByStatus.forEach(g => line(`${L1}${g[0]}: ${fmtNum(g[1].length)}`));
  line("");
  head("4.1. Status-change details by Asset Type and Model:");
  const shown = groupsByStatus.filter(g => g[1].length);
  if (!shown.length) { line(`${L1}No status changes to Scrapped, Inactive, Lost or Stolen`); }
  shown.forEach(g => {
    line(`${L1}${g[0]}`);
    typeCounts(g[1].map(c => c.cur)).forEach(t => {
      const ofType = g[1].filter(c => c.cur.assetClass === t);
      // Large groups as a count; small ones by model, with each asset and its Spider comment to complete
      if (ofType.length > EMAIL_DETAIL_MAX) { line(`${L2}${plural(t)}: ${fmtNum(ofType.length)}`); return; }
      line(`${L2}${t}`);
      keysByCount(countBy(ofType.map(c => c.cur.model || "(no model)"))).forEach(m => {
        const ofModel = ofType.filter(c => (c.cur.model || "(no model)") === m);
        line(`${L3}${m}: ${ofModel.length}`);
        ofModel.forEach(c => line(`${L4}Asset no: ${c.cur.assetId}` + (c.cur.serial ? ` - Serial number: ${c.cur.serial}` : "") +
          spiderComment(c.cur) + " - " + EMAIL_NOTE));
      });
    });
  });
  line("");
  head("5. Stock utilization:");
  stockLines(ctx).forEach(t => line(t));
  line("");
  line("Best regards,");

  // The model column of an unsupplied PO lists models with counts, e.g. "Bag (50)": the counts are dropped here
  const model = (r: Cell[]) => String(r[1]).replace(/ \(\d+\)/g, "");
  // Copilot in PowerPoint: 3 slides at most, numbers given exactly, the open VOIS template's look
  const ppt: string[] = [];
  const top = (m: Map<string, number>, n: number) => keysByCount(m).slice(0, n).map(k => `${k} ${fmtNum(m.get(k) || 0)}`).join(", ");
  const moved = flatten(groups.map(x => x.rows.filter(r => r[1] !== r[2]).map(r => `${r[0]}\u0001${r[2] - r[1]}`)))
    .map(t => t.split("\u0001")).sort((a, b) => Math.abs(+b[1]) - Math.abs(+a[1])).slice(0, 4).map(t => `${t[0]} ${signed(+t[1])}`).join(", ");
  const serialShort = (c: string) => c === SERIAL_CHANGED ? "changed" : c === SERIAL_ADDED ? "added" : "removed";
  ppt.push(`Create exactly 3 slides titled VCAMS Monthly Asset Reconciliation - ${curName}, using this VOIS template's layouts, ` +
    "colours and fonts. No title slide, no agenda, no closing slide. Use only the numbers below, exactly as written; do not add, " +
    "round or estimate numbers. Keep text short: at most 4 bullets per slide. Add the charts asked for as simple bar or column charts.");
  ppt.push(`Slide 1, Overview ${prevName} vs ${curName}: 3 large figures: ${prevName} ${fmtNum(rc.prevRows)} assets, ` +
    `${curName} ${fmtNum(rc.curRows)} assets, net movement ${signed(rc.curRows - rc.prevRows)}. Clustered column chart by value group, ` +
    `${prevName} vs ${curName}: ` + groups.map(x => `${x.g} ${fmtNum(x.p)} vs ${fmtNum(x.c)}`).join("; ") +
    `. One line: biggest changes ${moved || "none"}.`);
  ppt.push(`Slide 2, Movements this month: bar chart with New ${fmtNum(rc.newAssets.length)}, Removed ${fmtNum(rc.removed.length)}, ` +
    `Scrapped ${fmtNum(scrapped.length)}, Lost ${fmtNum(lost.length)}` +
    (rc.serialChecked ? `, Serial number changes ${fmtNum(rc.serialChanges.length)}` : "") + ". Bullets: " +
    (rc.newAssets.length ? `new by type ${top(countBy(rc.newAssets.map(a => a.assetClass)), 3)}; by PO ${byPo.size ? top(byPo, 2) : "none"}, ` +
      `blank PurchaseNo ${fmtNum(blankPo)}` : "no new assets") +
    (rc.operational.length ? `; scrapped or lost by type ${top(countBy(rc.operational.map(c => c.cur.assetClass)), 3)}` : "") +
    (rc.serialChecked && rc.serialChanges.length ? `; serials ${top(countBy(rc.serialChanges.map(c => serialShort(c.change))), 3)}` : "") +
    (rc.removed.length ? "; removed assets are to be reviewed, not treated as scrapped" : "") + "; stock utilization " +
    stockSummary(ctx).slice(0, 2).map(x => `${x.label.replace(" (all)", "")} ${utilOf(x) === "" ? "-" : (100 * (utilOf(x) as number)).toFixed(1) + "%"}`).join(", ") + ".");
  ppt.push("Slide 3, PO validation and points for review: a small table with columns PO, Model, In export, Supplied, Variance, Result: " +
    (po.length ? po.slice(0, 4).map(r => `${r[0] === "(empty)" ? "Empty PurchaseNo" : r[0]}, ${model(r) || "-"}, ${fmtNum(r[2] as number)}, ` +
      `${r[3] === "" ? "-" : fmtNum(r[3] as number)}, ${r[4] === "" ? "-" : signed(r[4] as number)}, ${r[7]}`).join("; ") +
      (po.length > 4 ? `; and ${po.length - 4} more POs in the workbook` : "") : "no new-asset POs") +
    `. Then up to 3 bullets for review: ${review.length ? review.slice(0, 3).join("; ") : "none"}.`);
  const d = rc.curDay || ctx.asOf, n = d.getUTCDate();
  const th = n % 10 === 1 && n !== 11 ? "st" : n % 10 === 2 && n !== 12 ? "nd" : n % 10 === 3 && n !== 13 ? "rd" : "th";
  const attachment = `Spider Data Integrity - ${n}${th} of ${monthName(monthOf(d)).split(" ")[0]}`;
  return { subject, attachment, email, heads, ppt };
}

// ============================================================================ checks
function runChecks(ctx: Ctx): void {
  const { assets, asOf, maps, cols, users, devices } = ctx;
  const days = (d: Date | null) => daysSince(d, asOf);
  // Device management match: serial first, then hostname
  assets.forEach(a => {
    const d = (a.serialN && devices.bySerial.get(a.serialN)) || (a.hostN && devices.byName.get(a.hostN)) || null;
    if (d) { a.device = d; d.matched = true; }
  });

  assets.forEach(a => {
    const active = a.stage !== DISPOSED, physical = !VIRTUAL_CLASSES.has(a.kind);
    const changed = a.statusDate || a.updated;
    // --- completeness
    if (active && cols.has("serial") && !a.serial && !NO_SERIAL_CLASSES.has(a.kind)) { addFlag(a, "NO_SERIAL", "", maps); }
    if (active && cols.has("tag") && !a.tag) { addFlag(a, "NO_TAG", "", maps); }
    if (active && (cols.has("make") || cols.has("model")) && ((cols.has("make") && !a.makeRaw) || (cols.has("model") && !a.model))) {
      addFlag(a, "NO_MODEL", !a.makeRaw && !a.model ? "Both missing" : !a.makeRaw ? "Manufacturer missing" : "Model missing", maps);
    }
    if (active && cols.has("cls") && a.group === UNLISTED) {
      addFlag(a, "NO_CLASS", a.clsRaw ? `'${a.clsRaw}' is not in the standard list` + (a.kind !== OTHER ? ` (read as ${a.kind})` : "") : "Empty", maps);
    }
    if (a.isNew && cols.has("po") && !a.po) { addFlag(a, "NEW_NO_PO", `New since ${ctx.recon.prevPeriod}, ${a.status}`, maps); }
    if (a.stage === IN_USE && !hasUser(a) && !LOCATION_CLASSES.has(a.kind) && (cols.has("user") || cols.has("email") || cols.has("empId"))) {
      addFlag(a, "NO_USER", "", maps);
    }
    if (cols.has("location") && physical && !a.location && [IN_USE, STOCK, REPAIR].indexOf(a.stage) >= 0) { addFlag(a, "NO_LOCATION", "", maps); }
    if (cols.has("costCenter") && !a.costCenter && a.stage === IN_USE) { addFlag(a, "NO_COST_CENTER", "", maps); }
    if (active && cols.has("purchase") && !a.purchase && !a.badDates.some(x => x.indexOf(FIELD_LABEL.purchase) === 0)) {
      addFlag(a, "NO_PURCHASE_DATE", "", maps);
    }
    // --- validity
    if (active && a.serial && (isPlaceholder(a.serial) || serialKey(a.serial).length < SETTINGS.minSerialChars)) {
      addFlag(a, "BAD_SERIAL", `'${short(a.serial, 30)}'`, maps);
    }
    if (a.email && !/^[^@\s]+@[^@\s]+\.[a-z]{2,}$/i.test(a.email)) { addFlag(a, "BAD_EMAIL", `'${short(a.email, 40)}'`, maps); }
    if (!a.standardStatus) {
      addFlag(a, "UNKNOWN_STATUS", a.status ? `'${a.status}'` + (a.stage !== UNKNOWN ? ` (read as ${a.stage})` : " (not recognised)") : "Empty", maps);
    }
    const bad = a.badDates.slice();
    const future = (d: Date | null, label: string) => { if (d && d.getTime() > asOf.getTime() + DAY) { bad.push(`${label} in the future (${fmtDay(d)})`); } };
    future(a.purchase, "Purchase date"); future(a.verified, "Last verified"); future(a.lastSeen, "Last seen"); future(a.deployDate, "Deployment date");
    future(a.created, "Creation date"); future(a.updated, "Update date");
    if (a.purchase && a.purchase.getUTCFullYear() < 1995) { bad.push(`Purchase date ${fmtDay(a.purchase)} is too early`); }
    if (a.purchase && a.warranty && a.warranty.getTime() < a.purchase.getTime()) { bad.push("Warranty ends before purchase"); }
    if (a.purchase && a.deployDate && a.deployDate.getTime() < a.purchase.getTime() - 31 * DAY) { bad.push("Deployed before purchase"); }
    if (bad.length) { addFlag(a, "BAD_DATES", bad.join("; "), maps); }
    const pattern = maps.hostname.get(a.market.toLowerCase()) || maps.hostname.get("*") || (HOSTNAME_PATTERN ? new RegExp(HOSTNAME_PATTERN, "i") : null);
    if (pattern && a.hostname && EUC_CLASSES.has(a.kind) && active && !pattern.test(a.hostname)) {
      addFlag(a, "BAD_HOSTNAME", `'${a.hostname}'`, maps);
    }
    // --- consistency
    if ([STOCK, RETIRED, DISPOSED].indexOf(a.stage) >= 0 && hasUser(a) && !LOCATION_CLASSES.has(a.kind)) {
      addFlag(a, "USER_ON_NOT_IN_USE", `${a.stage}, user ${userLabel(a)}`, maps);
    }
    const family = MODEL_FAMILY.find(f => f[1].test(norm(a.model)));
    if (family && EUC_CLASSES.has(a.kind) && family[0] !== a.kind) {
      addFlag(a, "CLASS_MODEL_MISMATCH", `Recorded as ${a.assetClass}, model '${a.model}' is a ${family[0].toLowerCase()}`, maps);
    }
    // --- timeliness
    if (cols.has("verified") && (a.stage === IN_USE || a.stage === STOCK) && (!a.verified || days(a.verified) > SETTINGS.notVerifiedDays)) {
      addFlag(a, "NOT_VERIFIED", a.verified ? `Last verified ${fmtDay(a.verified)} (${days(a.verified)} days)` : "Never verified", maps);
    }
    // Days in the current status: status date, else the asset's last update (approximate)
    const inStage = days(changed), since = `${inStage} days since ${fmtDay(changed)}` + (a.statusDate ? "" : " (last update)");
    if (a.stage === STOCK && physical && inStage > SETTINGS.stockAgedDays) { addFlag(a, "LONG_IN_STOCK", since, maps); }
    if (a.stage === REPAIR && inStage > SETTINGS.repairMaxDays) { addFlag(a, "LONG_IN_REPAIR", since, maps); }
    if (a.stage === TRANSIT && inStage > SETTINGS.transitMaxDays) { addFlag(a, "LONG_IN_TRANSIT", since, maps); }
    if (a.prev && [DISPOSED, LOST].indexOf(a.prev.stage) >= 0 && [IN_USE, STOCK, REPAIR, TRANSIT].indexOf(a.stage) >= 0) {
      addFlag(a, "STATUS_REVERSED", `${a.prev.status} in ${ctx.recon.prevPeriod}, now ${a.status}`, maps);
    }
    // --- reconciliation: users
    if (users.source && (a.stage === IN_USE || a.stage === REPAIR) && (a.email || a.empId)) {
      const u = (a.email && users.byEmail.get(a.email)) || (a.empId && users.byEmp.get(a.empId.toLowerCase())) || null;
      if (!u) {
        addFlag(a, "USER_NOT_FOUND", a.email || a.empId, maps);
      } else if (!u.active) {
        addFlag(a, "LEAVER_HOLDS_ASSET", u.leaveDate ? `Left on ${fmtDay(u.leaveDate)} (${days(u.leaveDate)} days ago)` : "Account disabled / inactive", maps);
      }
    }
    // --- reconciliation: device management (or the export's own last-seen date)
    const seen = a.device && a.device.lastSeen ? a.device.lastSeen : a.lastSeen;
    if (devices.source && EUC_CLASSES.has(a.kind) && a.stage === IN_USE && !a.device) {
      addFlag(a, "NOT_SEEN", a.serialN || a.hostN ? "No device with this serial or hostname" : "No serial or hostname to match", maps);
    }
    if (a.stage === IN_USE && seen && days(seen) > SETTINGS.staleSeenDays) {
      addFlag(a, "STALE_DEVICE", `Last seen ${fmtDay(seen)} (${days(seen)} days)`, maps);
    }
    // Seen after the status changed (a scan just before scrapping is normal)
    if (NOT_IN_USE_STAGES.has(a.stage) && seen && days(seen) <= SETTINGS.recentSeenDays && days(seen) >= -1 &&
      (!changed || seen.getTime() > changed.getTime())) {
      addFlag(a, "ONLINE_NOT_IN_USE", `${a.stage} in VCAMS, seen ${fmtDay(seen)}` + (a.device && a.device.user ? ` by ${a.device.user}` : ""), maps);
    }
    if (a.device && a.device.user && a.email && a.stage === IN_USE && a.device.user.indexOf("@") > 0 && a.device.user !== a.email) {
      addFlag(a, "USER_MISMATCH", `VCAMS ${a.email}, device management ${a.device.user}`, maps);
    }
    // --- lifecycle (information, not data errors)
    const life = lifecycleYears(a, maps);
    if (a.stage === IN_USE && a.warranty && a.warranty.getTime() < asOf.getTime()) {
      addFlag(a, "WARRANTY_EXPIRED", `Ended ${fmtDay(a.warranty)}`, maps);
    }
    if (life && a.purchase && addYears(a.purchase, life).getTime() < asOf.getTime()) {
      const age = (days(a.purchase) / 365.25).toFixed(1);
      if (a.stage === IN_USE) { addFlag(a, "OVER_AGE", `${age} years old (life ${life})`, maps); }
      if (a.stage === STOCK) { addFlag(a, "STOCK_OVER_AGE", `${age} years old (life ${life})`, maps); }
    }
  });

  // --- uniqueness: same identifier on several records that are not disposed
  const live = assets.filter(a => a.stage !== DISPOSED);
  const dup = (code: string, on: string, keyOf: (a: Asset) => string, show: (a: Asset) => string) => {
    groupBy(live, keyOf).forEach((list, k) => {
      if (list.length < 2) { return; }
      ctx.dupGroups.push({ on, value: show(list[0]), assets: list });
      list.forEach(a => addFlag(a, code, `${on} ${show(a)} also on ` + list.filter(x => x !== a).slice(0, 3).map(x => x.key).join(", ") +
        (list.length > 4 ? ` +${list.length - 4} more` : ""), maps));
    });
  };
  if (cols.has("assetId")) { dup("DUP_ASSET_ID", "Asset no.", a => idKey(a.assetId), a => a.assetId); }
  if (cols.has("serial")) { dup("DUP_SERIAL", "Serial", a => a.serialN.length >= SETTINGS.minSerialChars ? a.serialN : "", a => a.serial); }
  if (cols.has("tag")) { dup("DUP_TAG", "Asset tag", a => isPlaceholder(a.tag) ? "" : a.tag.toUpperCase().replace(/\s+/g, ""), a => a.tag); }
  if (cols.has("hostname")) { dup("DUP_HOSTNAME", "Hostname", a => a.hostN, a => a.hostname); }
  if (cols.has("mac")) { dup("DUP_MAC", "MAC", a => a.macN.length === 12 ? a.macN : "", a => a.mac); }

  // --- users holding more devices of one type than allowed
  groupBy(assets.filter(a => a.stage === IN_USE), a => { const k = userKey(a); return k ? k + "|" + a.kind : ""; }).forEach((list, k) => {
    const cls = list[0].kind, limit = ctx.maps.limits.get(compactKey(list[0].assetClass)) || ctx.maps.limits.get(compactKey(cls)) || PER_USER_LIMIT[cls] || 0;
    if (!limit || list.length <= limit) { return; }
    ctx.manyDevices.push({ user: userLabel(list[0]), email: list[0].email, assetClass: list[0].assetClass, assets: list });
    list.forEach(a => addFlag(a, "MANY_DEVICES", `${userLabel(a)} holds ${list.length} ${cls.toLowerCase()}s (limit ${limit})`, maps));
  });
  ctx.manyDevices.sort((a, b) => b.assets.length - a.assets.length);
}

// ============================================================================ stock and utilisation
// Days an in-stock asset has been in stock: status date, else last update, else purchase date
function stockAge(a: Asset, ctx: Ctx): number {
  if (ctx.cols.has("statusDate") || ctx.cols.has("updated")) { return daysSince(a.statusDate || a.updated, ctx.asOf); }
  return daysSince(a.purchase, ctx.asOf);
}
// Deployed (became in use) within the last N days, from dates (used when there is no previous export)
function deployedWithin(a: Asset, ctx: Ctx, n: number): boolean {
  if (a.stage !== IN_USE) { return false; }
  const d = ctx.cols.has("deployDate") ? a.deployDate : a.statusDate;
  const age = daysSince(d, ctx.asOf);
  return age >= 0 && age <= n;
}
// Deployments per month: exact from the reconciliation (moved into In operation since the previous export, or new and
// in operation), else the last 90 days of deployment / status dates divided by 3
function deployRate(list: Asset[], ctx: Ctx): number {
  if (ctx.recon.available) { return list.filter(a => a.deployedNow).length; }
  return list.filter(a => deployedWithin(a, ctx, 90)).length / 3;
}
const AGE_BUCKETS: [string, number][] = [["0-30 days", 30], ["31-90", 90], ["91-180", 180], ["181-365", 365], ["Over 1 year", 1e9]];
function ageBucket(days: number): string {
  if (days < 0) { return "Age unknown"; }
  for (const [label, max] of AGE_BUCKETS) {
    if (days <= max) { return label; }
  }
  return "Age unknown";
}
function coverMonths(stock: number, perMonth: number): Cell {
  return perMonth ? stock / perMonth : "";
}
function stockVerdict(stock: number, perMonth: number, aged: number, ctx: Ctx): string {
  if (!stock && perMonth >= SETTINGS.shortageMinDeployed) { return "Shortage - no stock left for steady demand"; }
  if (!stock) { return ""; }
  if (!perMonth) {
    return `Idle - no deployments ${ctx.recon.available ? "since the previous export" : "in 90 days"}; stop buying, redeploy or move ${stock} to where they are needed`;
  }
  const c = stock / perMonth;
  if (c > SETTINGS.coverMaxMonths) { return `Overstocked - ${c.toFixed(1)} months of cover; redeploy before buying` + (aged ? ` (${aged} aged)` : ""); }
  if (c < SETTINGS.coverMinMonths) { return `Low stock - ${(c * 30).toFixed(0)} days of cover; reorder or move stock in`; }
  return aged ? `Healthy cover; ${aged} aged units - deploy these first` : "Healthy";
}
const UTIL_COLS = ["In Use", "In Stock", "Utilisation"].concat(AGE_BUCKETS.map(b => "Stock " + b[0]),
  ["Stock Age Unknown", "Deployed / Month", "Months of Cover", "Stock Beyond Life", "Idle Stock Value", "What It Means"]);
// [group..., in use, in stock, utilisation, age buckets..., age unknown, deployed / month, months cover, beyond life, idle value, verdict]
function utilisationRows(ctx: Ctx, scope: Asset[], keyOf: (a: Asset) => string[]): Cell[][] {
  const groups = groupBy(scope.filter(a => (a.stage === IN_USE || a.stage === STOCK) && !VIRTUAL_CLASSES.has(a.kind)), a => keyOf(a).join("\u0001"));
  const rows: Cell[][] = [];
  groups.forEach((list, k) => {
    const use = list.filter(a => a.stage === IN_USE), stock = list.filter(a => a.stage === STOCK), pm = deployRate(use, ctx);
    if (!stock.length && pm < SETTINGS.shortageMinDeployed) { return; }
    const ages = stock.map(a => ageBucket(stockAge(a, ctx)));
    const aged = stock.filter(a => stockAge(a, ctx) > SETTINGS.stockAgedDays).length;
    rows.push((k.split("\u0001") as Cell[]).concat([use.length, stock.length, ratio(use.length, use.length + stock.length)],
      AGE_BUCKETS.map(b => ages.filter(x => x === b[0]).length), [ages.filter(x => x === "Age unknown").length, Math.round(pm * 10) / 10,
        coverMonths(stock.length, pm), stock.filter(a => hasFlag(a, "STOCK_OVER_AGE")).length,
        stock.reduce((s, a) => s + a.cost, 0), stockVerdict(stock.length, pm, aged, ctx)]));
  });
  const ki = scope.length ? keyOf(scope[0]).length : 0;
  return rows.sort((a, b) => (b[ki + 1] as number) - (a[ki + 1] as number));
}
// Sites with surplus vs sites with a shortage of the same type in the same market -> move stock
function moveSuggestions(ctx: Ctx): Cell[][] {
  const out: Cell[][] = [];
  const scope = ctx.assets.filter(a => (a.stage === IN_USE || a.stage === STOCK) && a.location && !VIRTUAL_CLASSES.has(a.kind));
  groupBy(scope, a => a.market + "\u0001" + a.assetClass).forEach((list, k) => {
    const [market, cls] = k.split("\u0001");
    const sites = Array.from(groupBy(list, a => a.location).entries()).map(e => {
      const stock = e[1].filter(a => a.stage === STOCK).length, pm = deployRate(e[1].filter(a => a.stage === IN_USE), ctx);
      return { site: e[0], stock, pm, cover: pm ? stock / pm : (stock ? 99 : 0) };
    });
    const surplus = sites.filter(s => s.stock > 0 && s.cover > SETTINGS.coverMaxMonths).sort((a, b) => b.stock - a.stock);
    const short = sites.filter(s => s.pm >= SETTINGS.shortageMinDeployed && s.cover < SETTINGS.coverMinMonths).sort((a, b) => b.pm - a.pm);
    short.forEach(s => {
      const from = surplus[0];
      const need = Math.max(1, Math.ceil(s.pm - s.stock));
      out.push([market, cls, s.site, s.stock, Math.round(s.pm * 10) / 10, need, from ? from.site : "", from ? from.stock : "",
        from ? `Move up to ${Math.min(need, from.stock)} from ${from.site} before ordering` : "Order or move stock from another market"]);
    });
  });
  return out;
}

// ============================================================================ history and comparison
const ALL_CLASSES = "(all)";
// Additive metrics per market + asset type; scores and rates are worked out from them, so they compare across runs
function historyFromAssets(ctx: Ctx): HistoryRow[] {
  const rows: HistoryRow[] = [];
  groupBy(ctx.assets, a => a.market + "\u0001" + a.assetClass).forEach((list, k) => {
    const [market, assetClass] = k.split("\u0001");
    const m = new Map<string, number>();
    const add = (metric: string, n: number) => { if (n) { m.set(metric, (m.get(metric) || 0) + n); } };
    add("#records", list.length);
    const scope = list.filter(a => inScope(a));
    add("#inscope", scope.length);
    STAGES.forEach(s => add("#stage:" + s, list.filter(a => a.stage === s).length));
    add("#issues", scope.filter(a => dataIssues(a).length > 0).length);
    add("#high", scope.filter(a => dataIssues(a).some(f => FLAG_INFO[f.code].sev === "High")).length);
    add("#dup", list.filter(a => a.flags.some(f => f.code.indexOf("DUP_") === 0)).length);
    add("#stockaged", list.filter(a => a.stage === STOCK && stockAge(a, ctx) > SETTINGS.stockAgedDays).length);
    add("#deployed", deployRate(list.filter(a => a.stage === IN_USE), ctx));
    add("#new", list.filter(a => a.isNew).length);
    const ch = list.filter(a => a.prev && compactKey(a.prev.status) !== compactKey(a.status));
    add("#changes", ch.length);
    add("#serialchanges", list.filter(a => a.prev && ctx.recon.serialChecked && serialCompare(a.prev.serial) !== serialCompare(a.serial)).length);
    add("#scrapped", ch.filter(a => compactKey(a.status) === "scrapped").length);
    add("#lost", ch.filter(a => compactKey(a.status) === "lost").length);
    list.forEach(a => a.flags.forEach(f => add(f.code, 1)));
    m.forEach((value, metric) => rows.push({ period: ctx.period, market, assetClass, metric, value }));
  });
  groupBy(ctx.recon.removed, a => a.market + "\u0001" + a.assetClass).forEach((list, k) => {
    const [market, assetClass] = k.split("\u0001");
    rows.push({ period: ctx.period, market, assetClass, metric: "#removed", value: list.length });
  });
  return rows;
}
function mergeHistory(json: string | undefined, current: HistoryRow[], period: string): HistoryRow[] {
  const markets = new Set<string>(current.map(r => r.market));
  let old: HistoryRow[] = [];
  if (json) {
    const parsed: HistoryRow[] = JSON.parse(json);
    old = parsed.map(r => ({ period: String(r.period), market: String(r.market), assetClass: String(r.assetClass || ""),
      metric: String(r.metric), value: Number(r.value) || 0 }));
  }
  // This run replaces any earlier run of the same period + market
  return old.filter(r => !(r.period === period && markets.has(r.market))).concat(current);
}
function previousPeriod(history: HistoryRow[], market: string, period: string): string {
  return history.filter(r => r.market === market && r.period < period && r.metric === "#records")
    .reduce((p, r) => r.period > p ? r.period : p, "");
}
// Metric totals for one market ("" = all markets) in one period; per-market previous periods when period = "prev"
function metricsFor(rows: HistoryRow[], market: string, cls?: string): Map<string, number> {
  const m = new Map<string, number>();
  rows.filter(r => (!market || r.market === market) && (!cls || r.assetClass === cls))
    .forEach(r => m.set(r.metric, (m.get(r.metric) || 0) + r.value));
  return m;
}
function prevRows(ctx: Ctx, market: string): HistoryRow[] {
  const ms = market ? [market] : ctx.markets;
  return ctx.history.filter(r => ms.indexOf(r.market) >= 0 && ctx.prevOf.get(r.market) === r.period && r.period !== "");
}
function hasPrevious(ctx: Ctx, market: string): boolean {
  return (market ? [market] : ctx.markets).some(m => !!ctx.prevOf.get(m));
}
function kpiValues(m: Map<string, number>): Kpis {
  const g = (k: string) => m.get(k) || 0;
  const stages: Record<string, number> = {};
  STAGES.forEach(s => { stages[s] = g("#stage:" + s); });
  const use = stages[IN_USE], stock = stages[STOCK];
  const checks = new Map<string, number>();
  Object.keys(FLAG_INFO).forEach(c => checks.set(c, g(c)));
  return {
    records: g("#records"), inScope: g("#inscope"), stages, issues: g("#issues"), high: g("#high"),
    score: g("#inscope") ? 1 - g("#issues") / g("#inscope") : "", util: ratio(use, use + stock), dup: g("#dup"), stockAged: g("#stockaged"),
    deployed: g("#deployed"), cover: coverMonths(stock, g("#deployed")), newAssets: g("#new"), removed: g("#removed"),
    changes: g("#changes"), serialChanges: g("#serialchanges"), scrapped: g("#scrapped"), lost: g("#lost"), checks
  };
}
// The monthly KPI list: [KPI, value, kind] where kind says how to show and compare it
function kpiList(k: Kpis, ctx: Ctx): [string, Cell, string][] {
  const c = (code: string) => k.checks.get(code) || 0;
  const list: [string, Cell, string][] = [
    ["Assets recorded", k.records, "n"], ["In scope (not scrapped / disposed)", k.inScope, "n"]];
  if (ctx.recon.available) {
    list.push(["New assets since the previous export", k.newAssets, "n"], ["Removed since the previous export (review)", k.removed, "bad"],
      ["Status changes since the previous export", k.changes, "n"], ["Newly Scrapped", k.scrapped, "n"], ["Newly Lost", k.lost, "bad"],
      ["Serial number changes", k.serialChanges, "n"],
      ["New assets with a blank PurchaseNo", c("NEW_NO_PO"), "bad"], ["Back from scrapped / lost", c("STATUS_REVERSED"), "bad"]);
  }
  STAGES.filter(s => s !== UNKNOWN).forEach(s => list.push([s, k.stages[s], "n"]));
  list.push(["Status not recognised (no stage)", k.stages[UNKNOWN], "bad"],
    ["Integrity score (in-scope records with no data issue)", k.score, "pct"], ["Records with a data issue", k.issues, "bad"],
    ["Records with a high-severity issue", k.high, "bad"],
    ["Records sharing an asset no. / serial / tag / hostname / MAC", k.dup, "bad"],
    ["Statuses outside the controlled list", c("UNKNOWN_STATUS"), "bad"], ["Asset types outside the standard list", c("NO_CLASS"), "bad"]);
  if (ctx.cols.has("user") || ctx.cols.has("email")) {
    list.push(["In use with no user", c("NO_USER"), "bad"], ["User on stock / scrapped asset", c("USER_ON_NOT_IN_USE"), "bad"]);
  }
  if (ctx.cols.has("verified")) { list.push([`Not verified in ${SETTINGS.notVerifiedDays} days`, c("NOT_VERIFIED"), "bad"]); }
  if (ctx.users.source) { list.push(["Assets held by leavers", c("LEAVER_HOLDS_ASSET"), "bad"], ["Assigned user not in directory", c("USER_NOT_FOUND"), "bad"]); }
  if (ctx.devices.source) { list.push(["In use but not in device management", c("NOT_SEEN"), "bad"]); }
  if (ctx.devices.source || ctx.cols.has("lastSeen")) {
    list.push([`In use, not seen for ${SETTINGS.staleSeenDays}+ days`, c("STALE_DEVICE"), "bad"],
      ["Online but recorded as not in use", c("ONLINE_NOT_IN_USE"), "bad"]);
  }
  list.push(["Stock utilisation (in use / in use + in stock)", k.util, "pct"],
    [`Stock older than ${SETTINGS.stockAgedDays} days`, k.stockAged, "bad"]);
  if (ctx.deploySource) { list.push(["Deployments per month", k.deployed, "num"], ["Months of stock cover", k.cover, "num"]); }
  list.push(["In stock beyond expected life", c("STOCK_OVER_AGE"), "bad"], ["In use beyond expected life", c("OVER_AGE"), "bad"],
    ["In use after warranty / lease end", c("WARRANTY_EXPIRED"), "bad"]);
  return list;
}
function changeLabel(now: Cell, prev: Cell, kind: string): Cell {
  if (now === "" || prev === "") { return ""; }
  const d = (now as number) - (prev as number);
  if (kind === "pct") { return Math.round(d * 1000) / 10; }
  return kind === "num" ? Math.round(d * 10) / 10 : d;
}

// ============================================================================ physical check
// A random, market-balanced sample of in-use and in-stock assets to check on site. Marks typed into an existing
// Physical Check sheet are kept, and the same sample is reused until it is cleared.
function seededRandom(seedText: string): () => number {
  let h = 2166136261;
  for (let i = 0; i < seedText.length; i++) { h = Math.imul(h ^ seedText.charCodeAt(i), 16777619); }
  return () => {
    h = Math.imul(h ^ (h >>> 15), 2246822507); h = Math.imul(h ^ (h >>> 13), 3266489909);
    return ((h ^= h >>> 16) >>> 0) / 4294967296;
  };
}
function readPhysical(wb: ExcelScript.Workbook, ctx: Ctx): Physical {
  const p: Physical = { sample: [], marks: new Map<string, PhysicalMark>(), found: 0, missing: 0, locOk: 0, locChecked: 0, userOk: 0, userChecked: 0 };
  const byKey = new Map<string, Asset>();
  ctx.assets.forEach(a => { if (!byKey.has(a.key)) { byKey.set(a.key, a); } });
  const ws = wb.getWorksheet(PHYSICAL_SHEET);
  const keep: string[] = [];
  if (ws) {
    const values = readValues(ws);
    const hr = values.findIndex(r => String(r[0]) === "Asset" && r.map(v => String(v)).indexOf(PHYSICAL_COLS[0]) >= 0);
    if (hr >= 0) {
      const ci = PHYSICAL_COLS.map(c => values[hr].map(v => String(v)).indexOf(c));
      values.slice(hr + 1).forEach(r => {
        const k = String(r[0]).trim();
        if (!k) { return; }
        keep.push(k);
        const v = (i: number) => String(r[ci[i]] === undefined ? "" : r[ci[i]]).trim().toUpperCase().slice(0, 1);
        const mark = { found: v(0), loc: v(1), user: v(2), comment: String(r[ci[3]] === undefined ? "" : r[ci[3]]) };
        if (mark.found || mark.loc || mark.user || mark.comment) { p.marks.set(k, mark); }
      });
    }
  }
  if (p.marks.size) {
    p.sample = keep.map(k => byKey.get(k)).filter(a => a !== undefined) as Asset[];
    p.marks.forEach(m => {
      if (m.found === "Y") { p.found++; } else if (m.found === "N") { p.missing++; }
      if (m.found === "Y" && (m.loc === "Y" || m.loc === "N")) { p.locChecked++; if (m.loc === "Y") { p.locOk++; } }
      if (m.found === "Y" && (m.user === "Y" || m.user === "N")) { p.userChecked++; if (m.user === "Y") { p.userOk++; } }
    });
    return p;
  }
  if (!SETTINGS.physicalSample) { return p; }
  const rnd = seededRandom(ctx.period + "|" + ctx.markets.join(","));
  const pool = ctx.assets.filter(a => (a.stage === IN_USE || a.stage === STOCK) && !VIRTUAL_CLASSES.has(a.kind));
  const n = Math.min(SETTINGS.physicalSample, pool.length);
  groupBy(pool, a => a.market).forEach(list => {
    const take = Math.max(1, Math.round(n * list.length / pool.length));
    const shuffled = list.map(a => ({ a, r: rnd() })).sort((x, y) => x.r - y.r).slice(0, take).map(x => x.a);
    p.sample = p.sample.concat(shuffled);
  });
  p.sample = p.sample.slice(0, Math.max(n, ctx.markets.length));
  return p;
}
function physicalText(p: Physical): string {
  const checked = p.found + p.missing;
  if (!checked) { return ""; }
  return `${p.found} of ${checked} sampled assets found (${pct(p.found, checked)})` +
    (p.locChecked ? `; location right for ${pct(p.locOk, p.locChecked)}` : "") +
    (p.userChecked ? `; user right for ${pct(p.userOk, p.userChecked)}` : "");
}

// ============================================================================ findings
interface Finding { sev: string; text: string; impact: string; action: string; owner: string; }
function buildFindings(ctx: Ctx): Finding[] {
  const scope = ctx.assets.filter(a => inScope(a));
  const out: Finding[] = [];
  const rc = ctx.recon;
  if (rc.available) {
    if (rc.removed.length) {
      out.push({ sev: "High", text: `${rc.removed.length} assets were in the ${monthName(rc.prevPeriod)} export but are missing from this one`,
        impact: "An asset should never disappear from VCAMS; absence is not the same as scrapped", action: "Review each one on the Removed Assets sheet; restore the record or document the disposal",
        owner: "Asset team / VCAMS owner" });
    }
    const lost = rc.operational.filter(c => compactKey(c.cur.status) === "lost").length;
    if (lost) {
      out.push({ sev: "High", text: `${lost} assets became Lost this month`, impact: "Possible data-security incident and write-off",
        action: "Confirm a security / loss report exists for each; block the devices", owner: "Asset team / Security" });
    }
    const po = poReconRows(ctx).filter(r => r[7] !== PO_RESULTS.matched);
    if (po.length) {
      out.push({ sev: "Medium", text: `${po.length} PO reconciliation lines are not matched: ` + topList(po.map(r => String(r[7])), 5),
        impact: "New assets that cannot be tied to a validated purchase order", action: "Supply or correct the PO validation (POs Validated sheet); fix PurchaseNo on new assets",
        owner: "Asset team / Procurement" });
    }
  }
  const byCode = new Map<string, Asset[]>();
  ctx.assets.forEach(a => a.flags.forEach(f => {
    const l = byCode.get(f.code);
    if (l) { l.push(a); } else { byCode.set(f.code, [a]); }
  }));
  Array.from(byCode.keys()).sort((a, b) => sevRank(a) - sevRank(b) || (byCode.get(b) as Asset[]).length - (byCode.get(a) as Asset[]).length)
    .forEach(code => {
      const list = byCode.get(code) as Asset[], info = FLAG_INFO[code];
      const wide = list.length > SETTINGS.systemicShare * scope.length && info.dim !== LIFECYCLE;
      const mk = countBy(list.map(a => a.market));
      const where = ctx.markets.length > 1 ? " - most in " + keysByCount(mk).slice(0, 2).map(m => `${m} (${mk.get(m)})`).join(", ") : "";
      out.push({
        sev: info.sev, text: `${list.length} assets (${pct(list.length, scope.length)}): ${info.name}${where}` +
          (wide ? ". Process-wide" : ""),
        impact: info.why,
        action: wide ? `Fix the process or the VCAMS form so the field is captured, then clean up in bulk. ${info.fix}` : info.fix,
        owner: info.owner
      });
    });
  const util = utilisationRows(ctx, ctx.assets, a => [a.market, a.assetClass]);
  util.filter(r => /^(Idle|Overstocked)/.test(String(r[r.length - 1]))).slice(0, 5).forEach(r => out.push({
    sev: "Medium", text: `${r[0]} ${String(r[1]).toLowerCase()}: ${r[3]} in stock - ${r[r.length - 1]}`,
    impact: "Money tied up in idle stock; devices age on the shelf", action: "Redeploy before buying; move stock between sites; dispose units past their life",
    owner: "Asset team / Procurement"
  }));
  util.filter(r => /^(Shortage|Low stock)/.test(String(r[r.length - 1]))).slice(0, 3).forEach(r => out.push({
    sev: "Medium", text: `${r[0]} ${String(r[1]).toLowerCase()}: ${r[r.length - 1]}`,
    impact: "Joiners and replacements wait for devices", action: "Reorder or move stock from a site with surplus (see Stock & Utilisation)",
    owner: "Asset team / Procurement"
  }));
  const pt = physicalText(ctx.physical);
  if (pt && ctx.physical.missing) {
    out.push({ sev: "High", text: `Physical check: ${pt}`, impact: "Records do not match what is on site",
      action: "Investigate the missing assets; repeat the check on the worst site", owner: "Asset team" });
  }
  return out.sort((a, b) => SEVERITY.indexOf(a.sev) - SEVERITY.indexOf(b.sev));
}

// ============================================================================ workbook output
const BRAND = "#E60000", DARK = "#333333", GREY = "#595959";
// REPORT_ONLY: the workbook holds only what the team sends every month (the reconciliation tabs and the email), plus
// Stock & Utilisation. false = the full audit (integrity score, exceptions, leavers, lifecycle, ...).
const REPORT_ONLY = true;
const REPORT_SHEETS = ["Monthly Email", "Pivot & Summary", "New Assets", "Removed Assets", "Serial Number Changes",
  "Scrap & Inactive Assets", "POs Validated", "Stock & Utilisation"];
const SHEETS = ["Start Here", "Monthly Email", "Executive Summary", "Monthly KPIs", "Pivot & Summary", "New Assets", "Removed Assets",
  "Serial Number Changes", "Scrap & Inactive Assets", "POs Validated", "Quality Checks", "Integrity Score", "Exceptions", "Duplicates",
  "Users & Leavers", "Reconciliation", "Stock & Utilisation", "Lifecycle", "Since Last Audit", "Inventory Overview",
  "Physical Check", "All Assets", "Glossary"];

function newSheet(wb: ExcelScript.Workbook, name: string, title: string, subtitle: string): ExcelScript.Worksheet {
  const old = wb.getWorksheet(name);
  if (old) { old.delete(); }
  const ws = wb.addWorksheet(name);
  ws.setShowGridlines(false);
  ws.setTabColor(SHEETS.indexOf(name) >= 0 && SHEETS.indexOf(name) < 4 ? BRAND : name === MAPPINGS_SHEET ? "#2E7D32" : DARK);
  const t = ws.getRange("A1");
  t.setValue(title);
  t.getFormat().getFont().setSize(16);
  t.getFormat().getFont().setBold(true);
  t.getFormat().getFont().setColor(BRAND);
  const s = ws.getRange("A2");
  s.setValue(subtitle);
  s.getFormat().getFont().setItalic(true);
  s.getFormat().getFont().setColor(GREY);
  return ws;
}
function safe(v: Cell): Cell {
  if (typeof v === "number") { return isFinite(v) ? v : 0; }
  const s = String(v === undefined || v === null ? "" : v).slice(0, 32000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function putValues(ws: ExcelScript.Worksheet, row: number, col: number, values: Cell[][]): void {
  // Excel on the web rejects requests over ~5 MB, so write in chunks of ~1 MB of text and
  // force a sync (getAddress) after each so queued writes are not merged into one request
  const MAX_CHARS = 1000000, MAX_ROWS = 2000;
  let start = 0;
  while (start < values.length) {
    let end = start, size = 0;
    while (end < values.length && end - start < MAX_ROWS && (end === start || size < MAX_CHARS)) {
      values[end].forEach(v => { size += String(v === null || v === undefined ? "" : v).length + 4; });
      end++;
    }
    const part = values.slice(start, end).map(r => r.map(v => safe(v)));
    const rg = ws.getRangeByIndexes(row + start, col, part.length, part[0].length);
    rg.setValues(part);
    if (values.length > 200) { rg.getAddress(); }
    start = end;
  }
}
interface TableOpts {
  widths: number[]; wrap?: number[]; pct?: number[]; bars?: number[]; change?: number[]; num?: number[]; freeze?: boolean;
  text?: number[]; date?: number[]; datetime?: number[]; count?: number[]; diff?: number[]; blankRow?: boolean; col?: number;
}
const TABLES = { n: 0 };
// Writes a formatted Excel table with its header at 0-based row `row` (and column `opts.col`, default A); returns the
// next free row. A name starting with "tbl" is used as it is (fixed names, used by formulas and PivotTables); others
// get a unique suffix. Header: VOIS red, white bold text.
function addTable(ws: ExcelScript.Worksheet, row: number, name: string, header: string[], rows: Cell[][], opts: TableOpts): number {
  const text = opts.text || [], c0 = opts.col || 0;
  const body: Cell[][] = rows.length ? rows.map(r => r.concat(new Array<Cell>(Math.max(0, header.length - r.length)).fill("")).slice(0, header.length)
    .map((v, i) => text.indexOf(i) >= 0 && typeof v === "number" ? String(v) : v))
    : [header.map((h, i) => i === 0 && !opts.blankRow ? "None found" : "")];
  putValues(ws, row, c0, [header]);
  // Identifiers as text before writing, so long numbers and leading zeros survive
  text.forEach(c => ws.getRangeByIndexes(row + 1, c0 + c, body.length, 1).setNumberFormat("@"));
  putValues(ws, row + 1, c0, body);
  const table = ws.addTable(ws.getRangeByIndexes(row, c0, body.length + 1, header.length), true);
  TABLES.n++;
  table.setName(/^tbl/.test(name) ? name : `t${name}${TABLES.n}`);
  table.setPredefinedTableStyle("TableStyleLight15");
  const head = ws.getRangeByIndexes(row, c0, 1, header.length).getFormat();
  head.getFill().setColor(BRAND);
  head.getFont().setColor("#FFFFFF");
  head.getFont().setBold(true);
  head.setWrapText(true);
  head.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  if (opts.freeze) { ws.getFreezePanes().freezeRows(row + 1); }
  opts.widths.forEach((w, c) => ws.getRangeByIndexes(0, c0 + c, 1, 1).getFormat().setColumnWidth(w));
  ws.getRangeByIndexes(row + 1, c0, body.length, header.length).getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
  const col = (c: number) => ws.getRangeByIndexes(row + 1, c0 + c, body.length, 1);
  (opts.wrap || []).forEach(c => col(c).getFormat().setWrapText(true));
  (opts.pct || []).forEach(c => col(c).setNumberFormat("0.0%"));
  (opts.num || []).forEach(c => col(c).setNumberFormat("#,##0.0"));
  (opts.bars || []).forEach(c => {
    const cf = col(c).addConditionalFormat(ExcelScript.ConditionalFormatType.dataBar);
    const bar = cf.getDataBar();
    if (bar) { bar.getPositiveFormat().setFillColor("#F4B6B6"); }
  });
  // Red = worse (more issues), green = better: a colour number format (conditional font colours fail on the web)
  (opts.change || []).forEach(c => col(c).setNumberFormat("[Red]+#,##0.0;[Color10]-#,##0.0;0"));
  (opts.date || []).forEach(c => col(c).setNumberFormat(FMT_DATE));
  (opts.datetime || []).forEach(c => col(c).setNumberFormat(FMT_DATETIME));
  (opts.count || []).forEach(c => col(c).setNumberFormat(FMT_COUNT));
  (opts.diff || []).forEach(c => col(c).setNumberFormat(FMT_DIFF));
  return row + body.length + 3;
}
// A link to another tab (or a cell): `target` is a sheet name, or "Sheet!A10"
function link(ws: ExcelScript.Worksheet, row: number, col: number, text: string, target: string): void {
  const parts = target.split("!"), ref = `'${parts[0].replace(/'/g, "''")}'!${parts[1] || "A1"}`;
  ws.getRangeByIndexes(row, col, 1, 1).setHyperlink({ documentReference: ref, textToDisplay: text, screenTip: `Go to ${parts[0]}` });
  ws.getRangeByIndexes(row, col, 1, 1).getFormat().getFont().setColor("#0563C1");
  ws.getRangeByIndexes(row, col, 1, 1).getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.single);
}
// Row 3 of the detail tabs: back to the summary
function backLink(ws: ExcelScript.Worksheet): void {
  link(ws, 2, 0, "‹ Back to Pivot & Summary", SUMMARY_SHEET);
}
function sectionTitle(ws: ExcelScript.Worksheet, row: number, text: string, note?: string, col?: number): number {
  const c = col || 0;
  putValues(ws, row, c, [[text]]);
  const f = ws.getRangeByIndexes(row, c, 1, 1).getFormat().getFont();
  f.setBold(true); f.setColor(BRAND); f.setSize(12);
  if (note) {
    putValues(ws, row + 1, c, [[note]]);
    ws.getRangeByIndexes(row + 1, c, 1, 1).getFormat().getFont().setItalic(true);
    ws.getRangeByIndexes(row + 1, c, 1, 1).getFormat().getFont().setColor(GREY);
    return row + 2;
  }
  return row + 1;
}
function textRows(ws: ExcelScript.Worksheet, row: number, lines: string[][]): number {
  if (lines.length) { putValues(ws, row, 0, lines.map(l => [l[0], l.length > 1 ? l[1] : ""])); }
  lines.forEach((l, i) => { if (l.length > 1 && l[0]) { ws.getRangeByIndexes(row + i, 0, 1, 1).getFormat().getFont().setBold(true); } });
  return row + lines.length + 1;
}

function assetCells(a: Asset): Cell[] {
  return [a.key, a.market, a.assetClass, a.status, a.stage, a.serial, a.tag, a.hostname, a.model, a.user, a.email, a.location, a.managedBy];
}
const ASSET_HEADER = ["Asset", "Market", "Asset Type", "VCAMS Status", "Stage", "Serial", "Asset Tag", "Hostname", "Model",
  "Assigned User", "User Email", "Location", "Managed By"];
const ASSET_WIDTHS = [110, 110, 90, 90, 80, 110, 90, 100, 130, 130, 170, 120, 110];

// Stock utilisation covers Notebooks and Headsets only (user decision, 2026-10-18). Headsets are the "Headset" type plus
// the headset models filed under "Zubehör-Client" (e.g. "Headset_Jabra Evolve 20_VE", "Headset Blackwire 520").
const HEADSET_MODEL = /headset|jabra|plantronics|blackwire|evolve|epos|sennheiser|voyager|poly\b/i;
const STOCK_NOTEBOOKS = "Notebooks", STOCK_HEADSETS = "Headsets";
function stockGroup(a: Asset): string {
  const t = compactKey(a.assetClass);
  if (t === "notebook") { return STOCK_NOTEBOOKS; }
  if (t === "headset") { return STOCK_HEADSETS; }
  if (t === compactKey("Zubehör-Client") && HEADSET_MODEL.test(a.model)) { return STOCK_HEADSETS; }
  return "";
}
interface StockRow { label: string; use: number; stock: number; deployed: number; }
// [Notebooks, Headsets (all), Headsets - Headset type, Headsets - Zubehör-Client headset models]
function stockSummary(ctx: Ctx): StockRow[] {
  const row = (label: string, keep: (a: Asset) => boolean): StockRow => {
    const list = ctx.assets.filter(a => keep(a)), use = list.filter(a => a.stage === IN_USE);
    return { label, use: use.length, stock: list.filter(a => a.stage === STOCK).length, deployed: deployRate(use, ctx) };
  };
  return [
    row(STOCK_NOTEBOOKS, a => stockGroup(a) === STOCK_NOTEBOOKS),
    row(STOCK_HEADSETS + " (all)", a => stockGroup(a) === STOCK_HEADSETS),
    row("   of which type Headset", a => stockGroup(a) === STOCK_HEADSETS && compactKey(a.assetClass) === "headset"),
    row("   of which under Zubehör-Client", a => stockGroup(a) === STOCK_HEADSETS && compactKey(a.assetClass) !== "headset")
  ];
}
function utilOf(r: StockRow): Cell {
  return r.use + r.stock ? r.use / (r.use + r.stock) : "";
}
function buildStockSheet(wb: ExcelScript.Workbook, ctx: Ctx): void {
  progress("Writing Stock & Utilisation");
  const monthly = ctx.recon.available;
  const ws = newSheet(wb, "Stock & Utilisation", "Stock Utilization - Notebooks and Headsets - " + monthName(ctx.period),
    "Utilization = in use / (in use + in stock). In stock = In stock, Reserved, Ready for use, To install, Building. " +
    "Headsets = type Headset + headset models under Zubehör-Client.");
  backLink(ws);
  const sum = stockSummary(ctx);
  let r = sectionTitle(ws, 4, "1. Summary", monthly ? "Deployed this month = moved to In operation since the previous export; " +
    "months of cover = in stock / deployed this month" : "Deployments need the previous export");
  const coverOf = (x: StockRow): Cell => x.deployed ? x.stock / x.deployed : "";
  r = addTable(ws, r, "tblStockSummary", ["Category", "In Use", "In Stock", "Utilization", "Deployed This Month", "Months of Cover"],
    sum.map(x => [x.label, x.use, x.stock, utilOf(x), monthly ? x.deployed : "", monthly ? coverOf(x) : ""]),
    { widths: [260, 110, 90, 90, 110, 100, 100], count: [1, 2, 4], pct: [3], num: [5], bars: [3] });
  ws.getRangeByIndexes(r - 3 - sum.length + 2, 0, 1, 6).getFormat().getFont().setBold(true);
  ws.getRangeByIndexes(r - 3 - sum.length + 1, 0, 1, 6).getFormat().getFont().setBold(true);
  // Per model, most stock first
  const byModel = (group: string): Cell[][] => Array.from(groupBy(ctx.assets.filter(a => stockGroup(a) === group &&
    (a.stage === IN_USE || a.stage === STOCK)), a => a.assetClass + "\u0001" + (a.model || "(no model)")).entries()).map(e => {
    const [cls, model] = e[0].split("\u0001"), use = e[1].filter(a => a.stage === IN_USE), st = e[1].filter(a => a.stage === STOCK).length;
    const dep = deployRate(use, ctx);
    return [model, cls, use.length, st, use.length + st ? use.length / (use.length + st) : "", monthly ? dep : "", monthly && dep ? st / dep : ""] as Cell[];
  }).sort((x, y) => (y[3] as number) - (x[3] as number) || (y[2] as number) - (x[2] as number)).slice(0, 100);
  const modelHeader = ["Model", "Asset Type", "In Use", "In Stock", "Utilization", "Deployed This Month", "Months of Cover"];
  const modelOpts: TableOpts = { widths: [260, 110, 90, 90, 110, 100, 100], count: [2, 3, 5], pct: [4], num: [6] };
  r = sectionTitle(ws, r, "2. Notebooks by model", "Most stock first");
  r = addTable(ws, r, "tblStockNotebookModels", modelHeader, byModel(STOCK_NOTEBOOKS), modelOpts);
  r = sectionTitle(ws, r, "3. Headsets by model", "Type Headset and headset models under Zubehör-Client; most stock first");
  addTable(ws, r, "tblStockHeadsetModels", modelHeader, byModel(STOCK_HEADSETS), modelOpts);
}

// Pivot & Summary: key figures (linked to the detail tabs), the summary by type, status changes, new-asset breakdowns
// and the three PivotTables, each under its own numbered heading.
//   Columns A-E: 1 Key figures, 2 Assets by type, 3 Status changes | G-H: 4-7 New assets | J.., M.., P..: 8-10 PivotTables
function buildSummarySheet(wb: ExcelScript.Workbook, ctx: Ctx, notes: string[]): void {
  progress("Writing Pivot & Summary");
  const rc = ctx.recon, prevName = monthName(rc.prevPeriod), curName = monthName(rc.curPeriod);
  const live = ctx.formulas && ctx.prevTabs.length === 1 && ctx.curTabs.length === 1 && ctx.prevTabs[0].colAbs.cls !== undefined &&
    ctx.curTabs[0].colAbs.cls !== undefined;
  const ws = newSheet(wb, SUMMARY_SHEET, `Monthly Asset Summary - ${prevName} vs ${curName}`, "Key figures link to the detail tabs " +
    "(click Open ›). " + (live ? "Counts by type are live formulas on the two export tabs." : "Counts calculated by the script."));
  [[0, 300], [1, 120], [2, 110], [3, 120], [4, 110], [5, 24], [6, 200], [7, 110], [8, 24], [9, 200], [10, 110], [11, 24],
    [12, 260], [13, 110], [14, 24], [15, 180], [16, 110]].forEach(w => ws.getRangeByIndexes(0, w[0], 1, 1).getFormat().setColumnWidth(w[1]));
  const at = (row: number) => `${SUMMARY_SHEET}!A${row + 1}`;

  // --- 2 Assets by type (built first: the key figures point at its grand total); placed after section 1
  const types = new Map<string, string>();
  VCAMS_TYPES.forEach(t => types.set(t[0], t[1]));
  Array.from(rc.prevByType.keys()).concat(Array.from(rc.curByType.keys())).forEach(t => {
    if (!types.has(t)) { const ty = typeOf(t, "", ctx.maps); types.set(t, ty.group); }
  });
  const rows: Cell[][] = [], kinds: string[] = [];
  VALUE_GROUPS.forEach(g => {
    const list = Array.from(types.keys()).filter(t => types.get(t) === g);
    if (!list.length || (g === UNLISTED && !list.some(t => rc.prevByType.get(t) || rc.curByType.get(t)))) { return; }
    rows.push([`${g.toUpperCase()} ASSETS`, "", "", "", ""]); kinds.push("group");
    list.forEach(t => {
      const p = rc.prevByType.get(t) || 0, c = rc.curByType.get(t) || 0;
      rows.push([t, p, c - p, c, live ? "OK" : ""]); kinds.push("type");
    });
    const p = list.reduce((sm, t) => sm + (rc.prevByType.get(t) || 0), 0), c = list.reduce((sm, t) => sm + (rc.curByType.get(t) || 0), 0);
    rows.push([`${g} total`, p, c - p, c, ""]); kinds.push("total");
  });
  rows.push(["Grand total", rc.prevRows, rc.curRows - rc.prevRows, rc.curRows, live ? "OK" : ""]); kinds.push("grand");

  // --- 1 Key figures: [label, value, formula (or ""), number format, link target (or "")]
  const scrapped = rc.operational.filter(c => compactKey(c.cur.status) === "scrapped").length;
  const ended = (st: string) => rc.changes.filter(c => compactKey(c.cur.status) === st).length;
  const po = poReconRows(ctx), sum = stockSummary(ctx);
  const keyTop = 5, sumTitle = keyTop + 18, sumTop = sumTitle + 2, grandRow = sumTop + 1 + rows.length;  // 0-based
  const trTitle = grandRow + 3;
  const key: [string, Cell, string, string, string][] = [
    [`Previous-month total assets (${prevName})`, rc.prevRows, `=B${grandRow}`, FMT_COUNT, at(grandRow - 1)],
    [`Current-month total assets (${curName})`, rc.curRows, `=D${grandRow}`, FMT_COUNT, at(grandRow - 1)],
    ["Net movement", rc.curRows - rc.prevRows, `=B${keyTop + 3}-B${keyTop + 2}`, FMT_DIFF, at(grandRow - 1)],
    ["New assets", rc.newAssets.length, `=COUNTIF(tblNewAssets[Asset Type],"?*")`, FMT_COUNT, "New Assets"],
    ["   of which with an empty PurchaseNo", rc.newAssets.filter(a => !a.po).length,
      `=COUNTIFS(tblNewAssets[Asset Type],"?*",tblNewAssets[PurchaseNo],"")`, FMT_COUNT, "New Assets"],
    ["Removed assets (to review)", rc.removed.length, rc.removed.length ? `=COUNTIF(tblRemovedAssets[Asset Type],"?*")` : "", FMT_COUNT,
      rc.removed.length ? "Removed Assets" : ""],
    ["Serial number changes", rc.serialChecked ? rc.serialChanges.length : "Not checked",
      rc.serialChecked ? `=COUNTIF(tblSerialChanges[Change],"?*")` : "", FMT_COUNT, rc.serialChecked ? "Serial Number Changes" : ""],
    ["All status changes", rc.changes.length, "=SUM(tblStatusTransitions[Assets])", FMT_COUNT, at(trTitle)],
    ["Scrapped (new this month)", scrapped, `=COUNTIF(tblStatusChanges[Current Asset Status],"Scrapped")`, FMT_COUNT, "Scrap & Inactive Assets"],
    ["Inactive (new this month)", ended("inactive"), `=SUMIFS(tblStatusTransitions[Assets],tblStatusTransitions[Status Change],"*→ Inactive")`, FMT_COUNT, at(trTitle)],
    ["Lost (new this month)", ended("lost"), `=COUNTIF(tblStatusChanges[Current Asset Status],"Lost")`, FMT_COUNT, "Scrap & Inactive Assets"],
    ["Stolen (new this month)", ended("stolen"), `=SUMIFS(tblStatusTransitions[Assets],tblStatusTransitions[Status Change],"*→ Stolen")`, FMT_COUNT, at(trTitle)],
    ["PO lines not 'Matched'", po.filter(x => x[7] !== PO_RESULTS.matched).length,
      `=COUNTIF(tblPOReconciliation[Reconciliation Result],"?*")-COUNTIF(tblPOReconciliation[Reconciliation Result],"Matched")`, FMT_COUNT, PO_SHEET],
    ["Notebook stock utilization", utilOf(sum[0]), "=INDEX(tblStockSummary[Utilization],1)", "0.0%", "Stock & Utilisation"],
    ["Headset stock utilization", utilOf(sum[1]), "=INDEX(tblStockSummary[Utilization],2)", "0.0%", "Stock & Utilisation"]
  ];
  let r = sectionTitle(ws, 3, "1. Key figures", "Live: each value is a formula on the detail tab it links to");
  addTable(ws, r, "tblKeyFigures", ["Measure", "Value", "Details"], key.map(k => [k[0], k[1], ""]), { widths: [] });
  key.forEach((k, i) => {
    const cell = ws.getRangeByIndexes(r + 1 + i, 1, 1, 1);
    cell.setNumberFormat(k[3]);
    if (k[2] && ctx.formulas) { cell.setFormulas([[k[2]]]); }
    if (k[4]) { link(ws, r + 1 + i, 2, k[4].indexOf("!") > 0 ? "See below ›" : `Open ${k[4]} ›`, k[4]); }
  });
  ws.getRangeByIndexes(r + 1, 1, key.length, 1).getFormat().getFont().setBold(true);

  // --- 2 Assets by type
  r = sectionTitle(ws, sumTitle, `2. Assets by type - ${prevName} vs ${curName}`, "High value / Low value / Other, every listed type even when zero" +
    (live ? "; 'Check' compares the live count with the script's count" : ""));
  addTable(ws, r, "tblAssetSummary", ["Asset Type", `${prevName} Assets`, "Difference", `${curName} Assets`, "Check"], rows,
    { widths: [], count: [1, 3], diff: [2] });
  const top = r + 1;
  kinds.forEach((k, i) => {
    const line = ws.getRangeByIndexes(top + i, 0, 1, 5).getFormat();
    if (k !== "type") { line.getFont().setBold(true); }
    if (k === "group") { line.getFill().setColor("#F2F2F2"); }
    if (k === "total") { line.getFill().setColor("#FCE4E4"); }
    if (k === "grand") { line.getFill().setColor("#F4B6B6"); }
  });
  // Live formulas on the raw tabs (one tab per export, all markets, real values only)
  if (live) {
    const rangeOf = (t: TabInfo) => `'${t.name.replace(/'/g, "''")}'!$${colLetter(t.colAbs.cls)}$${t.headerAbs + 2}:$${colLetter(t.colAbs.cls)}$${Math.max(t.lastAbs + 1, t.headerAbs + 2)}`;
    const pr = rangeOf(ctx.prevTabs[0]), cr = rangeOf(ctx.curTabs[0]);
    let groupStart = -1;
    kinds.forEach((k, i) => {
      const row = top + i + 1;
      if (k === "group") { groupStart = row + 1; return; }
      if (k === "type") {
        const v = rows[i];
        ws.getRangeByIndexes(top + i, 1, 1, 4).setFormulas([[`=SUMPRODUCT(--(TRIM(${pr})=$A${row}))`, `=D${row}-B${row}`,
          `=SUMPRODUCT(--(TRIM(${cr})=$A${row}))`, `=IF(AND(B${row}=${v[1]},D${row}=${v[3]}),"OK","Differs")`]]);
      } else if (k === "total") {
        ws.getRangeByIndexes(top + i, 1, 1, 3).setFormulas([[`=SUM(B${groupStart}:B${row - 1})`, `=D${row}-B${row}`, `=SUM(D${groupStart}:D${row - 1})`]]);
      } else if (k === "grand") {
        const totals = kinds.map((x, j) => x === "total" ? top + j + 1 : 0).filter(x => x);
        ws.getRangeByIndexes(top + i, 1, 1, 4).setFormulas([[`=${totals.map(x => "B" + x).join("+")}`, `=D${row}-B${row}`,
          `=${totals.map(x => "D" + x).join("+")}`, `=IF(AND(B${row}=${rc.prevRows},D${row}=${rc.curRows}),"OK","Differs")`]]);
      }
    });
    notes.push("Pivot & Summary uses live SUMPRODUCT formulas on the raw export tabs");
  }

  // --- 3 Status changes
  r = sectionTitle(ws, trTitle, `3. Status changes since ${prevName}`, `${fmtNum(rc.changes.length)} in total; only changes ending in ` +
    `${OPERATIONAL_STATUSES.join(" or ")} are listed on Scrap & Inactive Assets (${fmtNum(rc.operational.length)})`);
  const tr = countBy(rc.changes.map(c => c.transition));
  addTable(ws, r, "tblStatusTransitions", ["Status Change", "Assets", "On Scrap & Inactive Assets?"], keysByCount(tr).map(t =>
    [t, tr.get(t) || 0, OPERATIONAL_STATUSES.some(st => t.endsWith("→ " + st)) ? "Yes" : "No"]), { widths: [], count: [1] });

  // --- 4-7 New assets (columns G-H)
  const by = (f: (a: Asset) => string) => { const m = countBy(rc.newAssets.map(a => f(a))); return keysByCount(m).map(k => [k, m.get(k) || 0] as Cell[]); };
  r = sectionTitle(ws, 3, "4. New assets by type", `${fmtNum(rc.newAssets.length)} new assets`, 6);
  r = addTable(ws, r, "tblNewByType", ["Asset Type", "New Assets"], by(a => a.assetClass), { widths: [], count: [1], col: 6 });
  r = sectionTitle(ws, r, "5. New assets by PurchaseNo", "", 6);
  r = addTable(ws, r, "tblNewByPO", ["PurchaseNo", "New Assets"], by(a => a.po || "(empty)"), { widths: [], text: [0], count: [1], col: 6 });
  r = sectionTitle(ws, r, "6. PurchaseNo filled or empty", "", 6);
  r = addTable(ws, r, "tblNewByPOFilled", ["PurchaseNo", "New Assets"], by(a => a.po ? "Filled" : "Empty"), { widths: [], count: [1], col: 6 });
  r = sectionTitle(ws, r, "7. New assets by model", "Top 100", 6);
  addTable(ws, r, "tblNewByModel", ["Model", "New Assets"], by(a => a.model || "(empty)").slice(0, 100), { widths: [], count: [1], col: 6 });

  // --- 8-10 PivotTables (rebuilt every run; skipped, with a note, when the source table is empty). The source table is
  // looked up by name (Office Scripts does not allow keeping Excel objects in variables for later)
  const pivot = (title: string, name: string, tableName: string, n: number, col: number, rowsBy: string[], value: string, sum: boolean) => {
    sectionTitle(ws, 3, title, `From ${tableName}; right-click > Refresh after editing it`, col);
    const table = wb.getTable(tableName);
    if (!table || !n) {
      putValues(ws, 5, col, [["Nothing to show this month (the table is empty)."]]);
      notes.push(`${name} not created - its table is empty`);
      return;
    }
    try {
      const pt = ws.addPivotTable(name, table, ws.getRangeByIndexes(5, col, 1, 1));
      rowsBy.forEach(h => { const hier = pt.getHierarchy(h); if (hier) { pt.addRowHierarchy(hier); } });
      const valueHier = pt.getHierarchy(value);
      if (!valueHier) { notes.push(`${name}: column '${value}' not found`); return; }
      const dh = pt.addDataHierarchy(valueHier);
      dh.setSummarizeBy(sum ? ExcelScript.AggregationFunction.sum : ExcelScript.AggregationFunction.count);
      notes.push(`${name} created on ${SUMMARY_SHEET}!${colLetter(col)}6`);
    } catch (e) {
      putValues(ws, 5, col, [["PivotTable could not be created; the tables on the left hold the same counts."]]);
      notes.push(`${name} could not be created (${String(e).slice(0, 120)})`);
    }
  };
  pivot("8. PivotTable - new assets by type", "NewAssetsPivot", "tblNewAssets", rc.newAssets.length, 9, ["Asset Type"], "Asset no.", false);
  pivot("9. PivotTable - Scrapped / Lost by type and model", "StatusChangesPivot", "tblStatusChanges", rc.operational.length, 12,
    ["Current Asset Status", "Asset Type", "Model"], "Asset no.", false);
  pivot("10. PivotTable - PO quantities", "POValidationsPivot", "tblPOValidated", ctx.po.rows.length, 15, ["Purchase Number", "Model"], "Quantity", true);
}

function buildWorkbook(wb: ExcelScript.Workbook, ctx: Ctx): void {
  TABLES.n = 0;
  const assets = ctx.assets, scope = assets.filter(a => inScope(a)), date = fmtDay(ctx.asOf);
  const kNow = kpiValues(metricsFor(ctx.history.filter(r => r.period === ctx.period && ctx.markets.indexOf(r.market) >= 0), ""));
  const kPrev = kpiValues(metricsFor(prevRows(ctx, ""), ""));
  const prevAny = hasPrevious(ctx, "");
  const marketScope = ctx.markets.join(", ");
  // PivotTables point at the reconciliation tables: remove their sheet first so the tables can be rebuilt
  const oldSummary = wb.getWorksheet(SUMMARY_SHEET);
  if (oldSummary) { oldSummary.delete(); }
  const reconSheets = ["New Assets", "Removed Assets", "Serial Number Changes", "Scrap & Inactive Assets", "Quality Checks"];
  // Order matters: the summary's formulas and PivotTables point at the detail and stock tables, so those come first
  if (ctx.recon.available) {
    const notes = buildReconSheets(wb, ctx);
    buildStockSheet(wb, ctx);
    buildSummarySheet(wb, ctx, notes);
    if (notes.length) { ctx.runInfo.push(["PivotTables and formulas", notes.join("; ")]); }
    buildEmailSheet(wb, ctx);
  } else {
    buildStockSheet(wb, ctx);
    reconSheets.concat([EMAIL_SHEET]).forEach(n => { const w = wb.getWorksheet(n); if (w) { w.delete(); } });
    // The user's PO data is kept on its sheet even without a previous export
    if (ctx.po.rows.length) {
      const ws0 = newSheet(wb, PO_SHEET, "POs Validated", "Supplied by the user. PO reconciliation needs the previous export (new assets).");
      addTable(ws0, 3, "tblPOValidated", PO_HEADERS, ctx.po.rows.map(p => [p.date, p.number, p.model, p.quantity, p.raisedBy, p.physical, p.spider]),
        { widths: [110, 150, 150, 90, 130, 150, 130], text: [1], date: [0], count: [3] });
    }
  }
  if (REPORT_ONLY) {
    // Tabs that earlier (full audit) runs added are removed; the export tabs and Mappings stay
    SHEETS.filter(n => REPORT_SHEETS.indexOf(n) < 0).forEach(n => { const w = wb.getWorksheet(n); if (w) { w.delete(); } });
    REPORT_SHEETS.forEach((name, i) => { const w = wb.getWorksheet(name); if (w) { w.setPosition(i); } });
    const first = wb.getWorksheet(ctx.recon.available ? EMAIL_SHEET : "Stock & Utilisation");
    if (first) { first.activate(); }
    return;
  }

  // --- 1 Start Here
  progress("Writing Start Here");
  let ws = newSheet(wb, "Start Here", "VCAMS Asset Audit - " + date, `Markets: ${marketScope} | ${assets.length} asset records`);
  let r = textRows(ws, 3, [
    ["How to use this workbook", ""],
    ["1.", "Executive Summary and Monthly KPIs give the headline for the monthly report. Quality Checks must all pass " +
      "(or be reviewed) before the report is final."],
    ["", "Monthly reconciliation: Pivot & Summary (counts by type, High / Low value / Other), New Assets, Removed Assets, " +
      "Serial Number Changes, Scrap & Inactive Assets (new Scrapped / Lost only), POs Validated (your PO data + PO reconciliation)."],
    ["2.", "Exceptions is the work list: one row per problem on an asset, with the fix and the owner. Filter by market, " +
      "location or Managed By and hand it out."],
    ["3.", "Stock & Utilisation shows idle and aged stock, months of cover and where to move stock before buying."],
    ["4.", "Physical Check is a random sample to verify on site: type Y/N in the yellow columns and run the script again."],
    ["5.", "Mappings (if shown) lets you teach the script your VCAMS statuses and asset types without changing the code."],
    ["", ""],
    ["Raw exports", "Kept unchanged on their own tabs; every result is on a new sheet. No links to other workbooks."],
    ["Real names and emails", ctx.shareSafe ? "Replaced by tokens (share-safe mode)." : "Shown. Share this workbook with the requester and the service team only."]
  ]);
  r = sectionTitle(ws, r, "Run information");
  r = addTable(ws, r, "RunInfo", ["Item", "Value"], ctx.runInfo, { widths: [190, 900], wrap: [1] });
  r = sectionTitle(ws, r, "Tabs in this workbook");
  addTable(ws, r, "Tabs", ["Tab", "What it shows"], [
    ["Executive Summary", "Monthly reconciliation results, headline per market, and the findings with impact, action and owner"],
    ["Monthly KPIs", "The monthly report KPIs: this audit, previous audit, change, per market"],
    ["Pivot & Summary", "Assets by type: previous month, difference, this month, with High / Low value / Other totals; " +
      "status changes; new assets by type / PO / model; NewAssetsPivot, StatusChangesPivot, POValidationsPivot"],
    ["New Assets", "tblNewAssets: in the current export but not in the previous one (Asset no.)"],
    ["Removed Assets", "tblRemovedAssets: in the previous export but missing now - review required (only when there are any)"],
    ["Serial Number Changes", "tblSerialChanges: assets in both exports whose serial number was added, changed or removed"],
    ["Scrap & Inactive Assets", "tblStatusChanges: assets whose status changed to Scrapped or Lost this month"],
    ["POs Validated", "tblPOValidated (user-supplied, never invented) and the PO reconciliation against new assets"],
    ["Quality Checks", "The reconciliation controls (Pass / Review / Fail) and the completion report"],
    ["Integrity Score", "Share of records with no issue, per market, data-quality dimension, asset type, site and team"],
    ["Exceptions", "Every problem found, most severe first, with the fix and the owner"],
    ["Duplicates", "Records that share a serial number, asset tag, asset ID, hostname or MAC"],
    ["Users & Leavers", "Assets held by leavers, users with too many devices, users not in the directory"],
    ["Reconciliation", "VCAMS vs device management (Intune / SCCM): not seen, online but not in use, missing from VCAMS"],
    ["Stock & Utilisation", "Utilisation, stock age, deployments, months of cover, idle stock value, moves between sites"],
    ["Lifecycle", "Over-age and out-of-warranty assets and the refresh forecast"],
    ["Since Last Audit", "Each check this audit vs the previous one, per market"],
    ["Inventory Overview", "Counts by stage, type and market; status and type mappings used; names written several ways"],
    ["Physical Check", "Random sample to verify on site"],
    ["All Assets", "Every record with its stage, type and problems"],
    ["Glossary", "What each check means, why it matters, how to fix it"]], { widths: [190, 900], wrap: [1] });

  // --- 2 Executive Summary
  progress("Writing Executive Summary");
  ws = newSheet(wb, "Executive Summary", "Executive Summary - " + date,
    "Integrity score = in-scope (not disposed) records with no completeness, validity, uniqueness, consistency or timeliness issue");
  r = 3;
  if (ctx.recon.available) {
    r = sectionTitle(ws, r, `Monthly reconciliation - ${monthName(ctx.recon.prevPeriod)} vs ${monthName(ctx.recon.curPeriod)}`);
    r = textRows(ws, r, reconText(ctx).map(t => [t]));
    const qc = qualityChecks(ctx).filter(q => q[3] !== "Pass");
    r = textRows(ws, r - 1, [["Quality checks", qc.length ? `${qc.length} not passed: ` + qc.map(q => `${q[3]} - ${q[0]} (${q[2]})`).join("; ") :
      "All passed"]]);
  }
  r = sectionTitle(ws, r, "Headline by market");
  const headline = (label: string, market: string): Cell[] => {
    const k = kpiValues(metricsFor(ctx.history.filter(x => x.period === ctx.period && (market ? x.market === market : ctx.markets.indexOf(x.market) >= 0)), ""));
    const p = hasPrevious(ctx, market) ? kpiValues(metricsFor(prevRows(ctx, market), "")) : null;
    return [label, k.records, k.inScope, k.stages[IN_USE], k.stages[STOCK], k.score, p ? p.score : "",
      p && k.score !== "" && p.score !== "" ? Math.round(((k.score as number) - (p.score as number)) * 1000) / 10 : "",
      k.issues, k.high, k.checks.get("LEAVER_HOLDS_ASSET") || 0, k.checks.get("NOT_SEEN") || 0, k.util,
      k.records < SETTINGS.marketRateMin ? `Low volume (under ${SETTINGS.marketRateMin}) - not ranked` : ""];
  };
  const hl: Cell[][] = ctx.markets.map(m => headline(m, m));
  if (ctx.markets.length > 1) { hl.push(headline("All markets", "")); }
  r = addTable(ws, r, "Headline", ["Market", "Assets", "In Scope", "In Use", "In Stock", "Integrity Score", "Previous Score",
    "Change (pts)", "Records with Issues", "High Severity", "Held by Leavers", "Not in Device Mgmt", "Stock Utilisation", "Note"],
  hl, { widths: [150, 70, 70, 70, 70, 80, 80, 70, 80, 80, 80, 90, 80, 200], pct: [5, 6, 12], change: [7] });
  const pt = physicalText(ctx.physical);
  if (pt) { r = textRows(ws, r - 1, [["Physical check", pt]]); }
  r = sectionTitle(ws, r, "Findings", "Most severe first. 'Process-wide' = on more than " +
    `${Math.round(SETTINGS.systemicShare * 100)}% of records: fix the process, then clean up in bulk`);
  addTable(ws, r, "Findings", ["Severity", "What we found", "Impact", "Recommended action", "Owner"],
    buildFindings(ctx).map(f => [f.sev, f.text, f.impact, f.action, f.owner]), { widths: [70, 380, 280, 330, 160], wrap: [1, 2, 3] });

  // --- 3 Monthly KPIs
  progress("Writing Monthly KPIs");
  ws = newSheet(wb, "Monthly KPIs", "Monthly KPIs - " + date,
    prevAny ? "Previous = each market's last audit in History" : "No previous audit in History - this audit is the baseline");
  const nowList = kpiList(kNow, ctx), prevList = kpiList(kPrev, ctx);
  const perMarket = ctx.markets.map(m => kpiList(kpiValues(metricsFor(ctx.history.filter(x => x.period === ctx.period && x.market === m), "")), ctx));
  r = addTable(ws, 3, "Kpis", ["KPI", "This Audit", "Previous Audit", "Change"].concat(ctx.markets.length > 1 ? ctx.markets : []),
    nowList.map((k, i) => [k[0], k[1], prevAny ? prevList[i][1] : "", prevAny ? changeLabel(k[1], prevList[i][1], k[2]) : ""]
      .concat(ctx.markets.length > 1 ? perMarket.map(p => p[i][1]) : [])),
    { widths: [330, 90, 90, 80].concat(ctx.markets.map(() => 100)), change: [3] });
  // Percent KPIs: format their cells as %
  const marketCols = ctx.markets.length > 1 ? ctx.markets.length : 0;
  nowList.forEach((k, i) => {
    const f = k[2] === "pct" ? "0.0%" : k[2] === "num" ? "#,##0.0" : "";
    if (!f) { return; }
    ws.getRangeByIndexes(4 + i, 1, 1, 2).setNumberFormat(f);
    if (marketCols) { ws.getRangeByIndexes(4 + i, 4, 1, marketCols).setNumberFormat(f); }
  });
  textRows(ws, r, [["Note", "Change for percentages is in percentage points. Counts compare with each market's previous audit."]]);

  // --- 4 Integrity Score
  progress("Writing Integrity Score");
  ws = newSheet(wb, "Integrity Score", "Integrity Score - " + date, "Share of in-scope records with no issue in each dimension (higher is better)");
  const dimRow = (label: string, list: Asset[]): Cell[] => {
    const c = cleanCounts(list);
    return [label as Cell, list.length, ratio(c.clean, list.length)].concat(c.dims.map(n => ratio(n, list.length)));
  };
  const scopeByMarket = groupBy(scope, a => a.market), scopeByType = groupBy(scope, a => a.assetClass);
  const dimHeader = (first: string) => [first, "Records", "Integrity Score"].concat(ALL_DIMS.map(d => d + (SCORED_DIMS.indexOf(d) < 0 ? " (not in score)" : "")));
  const dimOpts: TableOpts = { widths: [170, 70, 80, 90, 90, 90, 90, 90, 110, 110], pct: [2, 3, 4, 5, 6, 7, 8, 9], bars: [2] };
  r = sectionTitle(ws, 3, "By market");
  r = addTable(ws, r, "ScoreMarket", dimHeader("Market"), ctx.markets.map(m => dimRow(m, scopeByMarket.get(m) || []))
    .concat(ctx.markets.length > 1 ? [dimRow("All markets", scope)] : []), dimOpts);
  r = sectionTitle(ws, r, "By asset type");
  r = addTable(ws, r, "ScoreType", dimHeader("Asset Type"), keysByCount(countBy(scope.map(a => a.assetClass))).map(c => dimRow(c, scopeByType.get(c) || [])), dimOpts);
  if (ctx.cols.has("location")) {
    r = sectionTitle(ws, r, "By site (lowest score first)", `Sites with at least ${SETTINGS.siteMinRecords} records`);
    r = addTable(ws, r, "ScoreSite", dimHeader("Market | Site"), Array.from(groupBy(scope, a => a.market + " | " + (a.location || "(no location)")).entries())
      .filter(e => e[1].length >= SETTINGS.siteMinRecords).map(e => dimRow(e[0], e[1])).sort((a, b) => (a[2] as number) - (b[2] as number)).slice(0, 40), dimOpts);
  }
  if (ctx.cols.has("managedBy")) {
    r = sectionTitle(ws, r, "By team (Managed By)");
    addTable(ws, r, "ScoreTeam", dimHeader("Managed By"), Array.from(groupBy(scope, a => a.managedBy || "(not set)").entries())
      .filter(e => e[1].length >= SETTINGS.siteMinRecords).map(e => dimRow(e[0], e[1])).sort((a, b) => (a[2] as number) - (b[2] as number)), dimOpts);
  }

  // --- 5 Exceptions
  progress("Writing Exceptions");
  ws = newSheet(wb, "Exceptions", "Exceptions - " + date, "One row per problem on an asset, most severe first. Filter by Market, Location, Managed By or Check and hand out.");
  const exc: Cell[][] = [];
  assets.forEach(a => a.flags.forEach(f => {
    const info = FLAG_INFO[f.code];
    exc.push(([info.sev, info.name, info.dim, f.detail, info.fix, a.managedBy || info.owner] as Cell[]).concat(assetCells(a)));
  }));
  exc.sort((a, b) => SEVERITY.indexOf(String(a[0])) - SEVERITY.indexOf(String(b[0])) || String(a[1]).localeCompare(String(b[1])) ||
    String(a[7]).localeCompare(String(b[7])));
  const excShown = exc.slice(0, SETTINGS.exceptionsMax);
  if (exc.length > excShown.length) { putValues(ws, 2, 0, [[`Showing the ${excShown.length} most severe of ${exc.length} exceptions (SETTINGS.exceptionsMax).`]]); }
  addTable(ws, 3, "Exceptions", ["Severity", "Check", "Dimension", "Detail", "How to fix", "Fix Owner"].concat(ASSET_HEADER), excShown,
    { widths: [65, 230, 95, 260, 260, 130].concat(ASSET_WIDTHS), freeze: true });

  // --- 6 Duplicates
  progress("Writing Duplicates");
  ws = newSheet(wb, "Duplicates", "Duplicates - " + date, "Records (not disposed) that share an identifier. Keep one record per physical asset.");
  const dupRows: Cell[][] = ctx.dupGroups.sort((a, b) => b.assets.length - a.assets.length).map(g => [g.on, g.value, g.assets.length,
    g.assets.map(a => a.key).join(", "), Array.from(new Set<string>(g.assets.map(a => a.market))).join(", "),
    g.assets.map(a => a.stage).join(", "), g.assets.map(a => userLabel(a) || "-").join(", "), g.assets.map(a => a.model || "-").join(", ")]);
  addTable(ws, 3, "Dups", ["Duplicate On", "Value", "Records", "Assets", "Markets", "Stages", "Users", "Models"], dupRows,
    { widths: [90, 150, 65, 260, 140, 200, 260, 220], wrap: [3, 5, 6, 7], freeze: true });

  // --- 7 Users & Leavers
  progress("Writing Users & Leavers");
  ws = newSheet(wb, "Users & Leavers", "Users & Leavers - " + date, ctx.users.source ? `User directory: ${ctx.users.source}` :
    "No user directory supplied: add a Users tab (HR / Entra ID export) or usersJson to find assets held by leavers");
  r = sectionTitle(ws, 3, "Assets held by users who have left");
  r = addTable(ws, r, "Leavers", ["Left / Disabled"].concat(ASSET_HEADER), assets.filter(a => hasFlag(a, "LEAVER_HOLDS_ASSET"))
    .map(a => [(a.flags.find(f => f.code === "LEAVER_HOLDS_ASSET") as Flag).detail as Cell].concat(assetCells(a))), { widths: [200].concat(ASSET_WIDTHS) });
  r = sectionTitle(ws, r, "Users holding more devices of a type than allowed", "Limits per type: " +
    Object.keys(PER_USER_LIMIT).map(c => `${c} ${ctx.maps.limits.get(compactKey(c)) || PER_USER_LIMIT[c]}`).join(", ") + " (tune with Mappings, Kind 'Limit')");
  r = addTable(ws, r, "Many", ["User", "Email", "Asset Type", "Devices", "Assets", "Models", "Markets"], ctx.manyDevices.map(g => [g.user, g.email,
    g.assetClass, g.assets.length, g.assets.map(a => a.key).join(", "), g.assets.map(a => a.model).join(", "),
    Array.from(new Set<string>(g.assets.map(a => a.market))).join(", ")]), { widths: [160, 200, 90, 65, 260, 260, 120], wrap: [4, 5] });
  r = sectionTitle(ws, r, "Assigned users not found in the user directory");
  addTable(ws, r, "NotFound", ["User / Email in VCAMS", "Assets", "Asset List", "Markets"], Array.from(groupBy(assets.filter(a => hasFlag(a, "USER_NOT_FOUND")),
    a => userLabel(a)).entries()).map(e => [e[0], e[1].length, e[1].map(a => a.key).join(", "), Array.from(new Set<string>(e[1].map(a => a.market))).join(", ")] as Cell[])
    .sort((a, b) => (b[1] as number) - (a[1] as number)), { widths: [240, 65, 300, 140], wrap: [2] });

  // --- 8 Reconciliation
  progress("Writing Reconciliation");
  ws = newSheet(wb, "Reconciliation", "Reconciliation with device management - " + date, ctx.devices.source ? `Devices: ${ctx.devices.source} ` +
    `(${ctx.devices.list.length} devices; ${ctx.devices.list.filter(d => d.matched).length} matched to VCAMS)` :
    "No device-management export supplied: add an Intune / SCCM tab or devicesJson. Last-seen checks use the export's own 'Last Seen' column if it has one.");
  const recon = (code: string) => assets.filter(a => hasFlag(a, code)).map(a => [(a.flags.find(f => f.code === code) as Flag).detail as Cell].concat(assetCells(a)));
  r = sectionTitle(ws, 3, "Online recently but VCAMS says stock / retired / disposed / lost", "Security and audit risk: someone is using these devices");
  r = addTable(ws, r, "Online", ["Detail"].concat(ASSET_HEADER), recon("ONLINE_NOT_IN_USE"), { widths: [260].concat(ASSET_WIDTHS) });
  r = sectionTitle(ws, r, "In use in VCAMS but not found in device management (laptops and desktops)");
  r = addTable(ws, r, "NotSeen", ["Detail"].concat(ASSET_HEADER), recon("NOT_SEEN"), {widths: [260].concat(ASSET_WIDTHS) });
  r = sectionTitle(ws, r, `In use but not seen for more than ${SETTINGS.staleSeenDays} days`);
  r = addTable(ws, r, "Stale", ["Detail"].concat(ASSET_HEADER), recon("STALE_DEVICE"), { widths: [260].concat(ASSET_WIDTHS) });
  r = sectionTitle(ws, r, "Different user in device management");
  r = addTable(ws, r, "UserMismatch", ["Detail"].concat(ASSET_HEADER), recon("USER_MISMATCH"), { widths: [260].concat(ASSET_WIDTHS) });
  r = sectionTitle(ws, r, "In device management but not in VCAMS", "Devices to register (or old enrolments to clean up). " +
    `Seen in the last ${SETTINGS.staleSeenDays} days first`);
  addTable(ws, r, "Orphans", ["Device Name", "Serial", "Last Seen", "Days Ago", "User", "Market"], ctx.orphans
    .map(d => [d.name, d.serial, fmtDay(d.lastSeen), d.lastSeen ? daysSince(d.lastSeen, ctx.asOf) : "", d.user, d.market] as Cell[])
    .sort((a, b) => (a[3] === "" ? 1e9 : a[3] as number) - (b[3] === "" ? 1e9 : b[3] as number)).slice(0, 20000), { widths: [150, 140, 90, 70, 220, 120] });

  // --- 9 Stock & Utilisation: written above, before the summary

  // --- 10 Lifecycle
  progress("Writing Lifecycle");
  ws = newSheet(wb, "Lifecycle", "Lifecycle - " + date, "Expected life (years): " +
    Object.keys(LIFECYCLE_YEARS).map(c => `${c} ${ctx.maps.lifecycle.get(compactKey(c)) || LIFECYCLE_YEARS[c]}`).join(", ") + " (tune with Mappings, Kind 'Lifecycle')");
  const inUse = assets.filter(a => a.stage === IN_USE);
  const due = (a: Asset): Date | null => a.purchase && lifecycleYears(a, ctx.maps) ? addYears(a.purchase, lifecycleYears(a, ctx.maps)) : null;
  const dueIn = (a: Asset, from: number, to: number) => {
    const d = due(a);
    if (!d) { return false; }
    const days = (d.getTime() - ctx.asOf.getTime()) / DAY;
    return days > from && days <= to;
  };
  r = sectionTitle(ws, 3, "Refresh forecast (in-use assets)", "Due = purchase date + expected life");
  r = addTable(ws, r, "Refresh", ["Market", "Asset Type", "In Use", "Purchase Date Known", "Already Overdue", "Due in 0-3 Months", "Due in 4-6 Months",
    "Due in 7-12 Months", "Out of Warranty", "Average Age (years)"], Array.from(groupBy(inUse, a => a.market + "\u0001" + a.assetClass).entries()).map(e => {
      const [m, c] = e[0].split("\u0001"), list = e[1], known = list.filter(a => a.purchase);
      return [m, c, list.length, known.length, list.filter(a => hasFlag(a, "OVER_AGE")).length, list.filter(a => dueIn(a, 0, 91)).length,
        list.filter(a => dueIn(a, 91, 183)).length, list.filter(a => dueIn(a, 183, 366)).length, list.filter(a => hasFlag(a, "WARRANTY_EXPIRED")).length,
        known.length ? Math.round(known.reduce((s, a) => s + daysSince(a.purchase, ctx.asOf), 0) / known.length / 36.525) / 10 : ""] as Cell[];
    }).sort((a, b) => (b[4] as number) - (a[4] as number)), { widths: [140, 100, 65, 80, 75, 75, 75, 75, 80, 80], num: [9] });
  r = sectionTitle(ws, r, "Age profile of in-use assets (years since purchase)");
  const ageBands: [string, number][] = [["Under 1", 1], ["1-2", 2], ["2-3", 3], ["3-4", 4], ["4-5", 5], ["5-6", 6], ["Over 6", 999]];
  addTable(ws, r, "AgeProfile", ["Asset Type"].concat(ageBands.map(b => b[0]), ["Unknown"]), keysByCount(countBy(inUse.map(a => a.assetClass))).map(c => {
    const list = inUse.filter(a => a.assetClass === c);
    const yrs = list.map(a => a.purchase ? daysSince(a.purchase, ctx.asOf) / 365.25 : -1);
    let lo = 0;
    return ([c] as Cell[]).concat(ageBands.map(b => { const n = yrs.filter(y => y >= lo && y < b[1]).length; lo = b[1]; return n; }), [yrs.filter(y => y < 0).length]);
  }), { widths: [120, 65, 65, 65, 65, 65, 65, 65, 70] });

  // --- 11 Since Last Audit
  progress("Writing Since Last Audit");
  ws = newSheet(wb, "Since Last Audit", "Since Last Audit - " + date, prevAny ? "Previous = each market's last audit in History (" +
    ctx.markets.map(m => `${m}: ${ctx.prevOf.get(m) || "none"}`).join("; ") + ")" : "No previous audit in History - this audit is the baseline");
  const checkRows = (market: string): Cell[][] => {
    const now = kpiValues(metricsFor(ctx.history.filter(x => x.period === ctx.period && (market ? x.market === market : ctx.markets.indexOf(x.market) >= 0)), ""));
    const prev = kpiValues(metricsFor(prevRows(ctx, market), ""));
    const has = hasPrevious(ctx, market);
    return Object.keys(FLAG_INFO).filter(c => (now.checks.get(c) || 0) || (has && (prev.checks.get(c) || 0))).map(c => {
      const n = now.checks.get(c) || 0, p = prev.checks.get(c) || 0;
      return [FLAG_INFO[c].name, FLAG_INFO[c].sev, FLAG_INFO[c].dim, has ? p : "", n, has ? n - p : "",
        !has ? "Baseline" : !p ? "New" : !n ? "Resolved" : n < p ? "Improved" : n > p ? "Worse" : "No change"];
    }).sort((a, b) => SEVERITY.indexOf(String(a[1])) - SEVERITY.indexOf(String(b[1])) || (b[4] as number) - (a[4] as number));
  };
  r = 3;
  (ctx.markets.length > 1 ? [""].concat(ctx.markets) : ctx.markets).forEach(m => {
    r = sectionTitle(ws, r, m || "All markets");
    r = addTable(ws, r, "Since", ["Check", "Severity", "Dimension", "Previous", "This Audit", "Change", "Trend"], checkRows(m),
      { widths: [330, 70, 100, 75, 75, 70, 90], change: [5] });
  });

  // --- 12 Inventory Overview
  progress("Writing Inventory Overview");
  ws = newSheet(wb, "Inventory Overview", "Inventory Overview - " + date, "What is in VCAMS and how the script read it");
  r = sectionTitle(ws, 3, "Assets by market and stage");
  r = addTable(ws, r, "StageMarket", ["Market"].concat(STAGES, ["Total"]), ctx.markets.map(m => {
    const ms = assets.filter(a => a.market === m);
    return ([m] as Cell[]).concat(STAGES.map(s => ms.filter(a => a.stage === s).length), [ms.length]);
  }), { widths: [150, 70, 70, 70, 70, 70, 80, 70, 80, 70] });
  r = sectionTitle(ws, r, "Assets by type and stage");
  r = addTable(ws, r, "StageType", ["Asset Type"].concat(STAGES, ["Total"]), keysByCount(countBy(assets.map(a => a.assetClass))).map(c => {
    const cs = assets.filter(a => a.assetClass === c);
    return ([c] as Cell[]).concat(STAGES.map(s => cs.filter(a => a.stage === s).length), [cs.length]);
  }), { widths: [150, 70, 70, 70, 70, 70, 80, 70, 80, 70] });
  r = sectionTitle(ws, r, "How each VCAMS status was read", "Not controlled? Confirm its meaning, then add a Mappings row: Status | <status> | <controlled status or stage>");
  r = addTable(ws, r, "StatusMap", ["VCAMS Status (as in the export)", "Reported As", "Controlled Status?", "Stage", "Assets"],
    Array.from(groupBy(assets, a => a.statusRaw || "(empty)").entries()).map(e => [e[0], e[1][0].status || "(blank)", e[1][0].standardStatus ? "Yes" : "No - review",
      e[1][0].stage, e[1].length] as Cell[]).sort((a, b) => (b[4] as number) - (a[4] as number)), { widths: [200, 160, 110, 150, 70] });
  r = sectionTitle(ws, r, "How each VCAMS asset type was read", "Not listed? Add Mappings rows: Group | <type> | High value / Low value / Other, and Class | <type> | <kind>");
  r = addTable(ws, r, "ClassMap", ["VCAMS Asset Type (as in the export)", "Reported As", "Value Group", "Kind (audit rules)", "Assets", "Example Models"],
    Array.from(groupBy(assets, a => a.clsRaw || "(empty)").entries()).map(e => [e[0], e[1][0].assetClass, e[1][0].group, e[1][0].kind, e[1].length,
      topList(e[1].map(a => a.model).filter(x => x), 3)] as Cell[]).sort((a, b) => (b[4] as number) - (a[4] as number)).slice(0, 200),
  { widths: [200, 140, 90, 120, 70, 400], wrap: [5] });
  r = sectionTitle(ws, r, "Top models", "In use and in stock");
  r = addTable(ws, r, "Models", ["Asset Type", "Manufacturer", "Model", "In Use", "In Stock", "Total"], Array.from(groupBy(assets.filter(a => a.stage === IN_USE || a.stage === STOCK),
    a => [a.assetClass, a.make, a.model || "(no model)"].join("\u0001")).entries()).map(e => (e[0].split("\u0001") as Cell[]).concat([
      e[1].filter(a => a.stage === IN_USE).length, e[1].filter(a => a.stage === STOCK).length, e[1].length])).sort((a, b) => (b[5] as number) - (a[5] as number)).slice(0, 50),
  { widths: [110, 110, 220, 65, 65, 65] });
  r = sectionTitle(ws, r, "Names written several ways", "Standardise these values in VCAMS (pick one spelling) - they split reports and break matching");
  const variants: Cell[][] = [];
  const addVariants = (field: string, keyOf: (a: Asset) => string, valueOf: (a: Asset) => string) => {
    groupBy(assets.filter(a => valueOf(a)), keyOf).forEach((list, k) => {
      const spellings = countBy(list.map(a => valueOf(a)));
      if (spellings.size > 1) { variants.push([field, keysByCount(spellings)[0], spellings.size, list.length, keysByCount(spellings).map(s => `${s} (${spellings.get(s)})`).join("; ")]); }
    });
  };
  addVariants("Manufacturer", a => a.make.toLowerCase(), a => a.makeRaw);
  addVariants("Model", a => norm(a.model).replace(/ /g, ""), a => a.model);
  addVariants("Location", a => norm(a.location).replace(/ /g, ""), a => a.location);
  addVariants("Status", a => norm(a.status).replace(/ /g, ""), a => a.status);
  addVariants("Asset type", a => norm(a.clsRaw).replace(/ /g, ""), a => a.clsRaw);
  addTable(ws, r, "Variants", ["Field", "Most Used Spelling", "Spellings", "Assets", "All Spellings"], variants.sort((a, b) => (b[3] as number) - (a[3] as number)).slice(0, 200),
    { widths: [100, 200, 70, 70, 600], wrap: [4] });

  // --- 13 Physical Check
  progress("Writing Physical Check");
  const ph = ctx.physical;
  if (ph.sample.length) {
    ws = newSheet(wb, PHYSICAL_SHEET, "Physical Check - " + date, physicalText(ph) ||
      "Check each asset on site: type Y or N in the yellow columns, save, and run the script again to get the result. Clear the sheet to draw a new sample.");
    const rows = ph.sample.map(a => {
      const m = ph.marks.get(a.key);
      return [a.key, a.market, a.assetClass, a.model, a.serial, a.tag, a.hostname, a.user, a.location, a.stage,
        m ? m.found : "", m ? m.loc : "", m ? m.user : "", m ? m.comment : ""] as Cell[];
    });
    addTable(ws, 3, "Physical", ["Asset", "Market", "Asset Type", "Model", "Serial", "Asset Tag", "Hostname", "Assigned User", "Location", "Stage"].concat(PHYSICAL_COLS),
      rows, { widths: [110, 110, 90, 130, 110, 90, 100, 140, 130, 70, 75, 85, 75, 220], freeze: true });
    ws.getRangeByIndexes(4, 10, rows.length, 4).getFormat().getFill().setColor("#FFF2CC");
  } else {
    const old = wb.getWorksheet(PHYSICAL_SHEET);
    if (old) { old.delete(); }
  }

  // --- 14 All Assets
  progress("Writing All Assets");
  // Big exports: not listed again (the raw export tab has every asset and Exceptions every problem); keeps the run fast
  const listed = assets.length > SETTINGS.allAssetsMax ? [] : assets;
  ws = newSheet(wb, "All Assets", "All Assets - " + date, listed.length < assets.length ?
    `${assets.length} records is above SETTINGS.allAssetsMax (${SETTINGS.allAssetsMax}), so they are not listed again here: use the raw ` +
    "export tab for every asset and Exceptions for every problem (raise allAssetsMax in Mappings to list them)" :
    "Every record with its stage, type and problems");
  addTable(ws, 3, "All", ASSET_HEADER.concat(["PurchaseNo", "Purchase Date", "Last Verified", "Last Scan / Seen", "Sheet Row", "Problems", "Worst Severity", "Problems Found"]),
    listed.map(a => assetCells(a).concat([a.po, dateSerial(a.purchase), dateSerial(a.verified), dateSerial(a.device && a.device.lastSeen ? a.device.lastSeen : a.lastSeen),
      `${a.sheet}!${a.row}`, a.flags.length, worstSeverity(a), plainFlags(a)])),
    { widths: ASSET_WIDTHS.concat([110, 95, 95, 120, 110, 65, 70, 500]), text: [0, 5, 6, 13], date: [14, 15], datetime: [16], freeze: true });

  // --- 15 Glossary
  progress("Writing Glossary");
  ws = newSheet(wb, "Glossary", "Glossary", "What each check means and how to tune the script");
  r = addTable(ws, 3, "Checks", ["Check Code", "Check", "Dimension", "Severity", "Why It Matters", "How to Fix", "Owner", "Needs"],
    Object.keys(FLAG_INFO).map(c => [c, FLAG_INFO[c].name, FLAG_INFO[c].dim, FLAG_INFO[c].sev, FLAG_INFO[c].why, FLAG_INFO[c].fix, FLAG_INFO[c].owner,
      CHECK_NEEDS[c] || ""]), { widths: [150, 260, 100, 65, 300, 300, 150, 220], wrap: [4, 5, 7] });
  r = sectionTitle(ws, r, "Controlled VCAMS statuses", "Any other status is flagged for review; it is still read into a stage by keywords");
  r = addTable(ws, r, "Statuses", ["Status", "Stage", "On Scrap & Inactive Assets when new this month"], VCAMS_STATUSES.map(x => [x[0], x[1],
    OPERATIONAL_STATUSES.indexOf(x[0]) >= 0 ? "Yes" : ""]), { widths: [160, 170, 260] });
  r = sectionTitle(ws, r, "Standard asset types");
  r = addTable(ws, r, "Types", ["Asset Type", "Value Group", "Kind (audit rules)", "Needs a user in use", "Needs a serial", "Expected Life (years)"],
    VCAMS_TYPES.map(x => [x[0], x[1], x[2], LOCATION_CLASSES.has(x[2]) ? "No" : "Yes", NO_SERIAL_CLASSES.has(x[2]) ? "No" : "Yes",
      LIFECYCLE_YEARS[x[2]] || ""]), { widths: [170, 100, 130, 110, 100, 110] });
  r = sectionTitle(ws, r, "Mappings (central table or a 'Mappings' sheet)", "Kind | Value | MapsTo");
  addTable(ws, r, "MapKinds", ["Kind", "Value", "MapsTo", "Example"], [
    ["Status", "A VCAMS status", "A controlled status or a stage: " + STAGES.join(", "), "Status | Awaiting Collection | In stock"],
    ["Group", "A VCAMS asset type", "High value, Low value or Other (adds it to the standard list)", "Group | Docking Station | High value"],
    ["Class", "A VCAMS asset type", "Kind for the audit rules: Laptop, Desktop, Monitor, Mobile Phone, Headset, Peripheral, ...", "Class | Docking Station | Docking Station"],
    ["Column", "A standard header (e.g. Asset status)", "Your export's column name, when several columns match", "Column | Asset status | Current Status"],
    ["Market", "Text in the company / country column", "Market label", "Market | vodafone egypt | VOIS EG"],
    ["Setting", "A SETTINGS name", "Number", "Setting | stockAgedDays | 120"],
    ["Limit", "Asset type", "Most per user while in use", "Limit | Laptop | 2"],
    ["Lifecycle", "Asset type", "Expected life in years", "Lifecycle | Laptop | 5"],
    ["HostnamePattern", "Market (empty = all)", "Regular expression", "HostnamePattern | VOIS EG | ^EG[LD]\\d{6}$"],
    ["Check", "A check code", "Off, or High / Medium / Low", "Check | NO_COST_CENTER | Off"]], { widths: [120, 200, 380, 320], wrap: [2] });
  if (!wb.getWorksheet(MAPPINGS_SHEET) && ctx.maps.source === "None") {
    ws = newSheet(wb, MAPPINGS_SHEET, "Mappings", "Teach the script your VCAMS values. Add rows, then run the script again. See Glossary > Mappings.");
    addTable(ws, 3, "Mappings", ["Kind", "Value", "MapsTo", "Note"], [], { widths: [120, 220, 220, 300] });
  }

  SHEETS.concat([MAPPINGS_SHEET]).forEach((name, i) => { const w = wb.getWorksheet(name); if (w) { w.setPosition(i); } });
  wb.getWorksheet(SHEETS[0])?.activate();
}
const CHECK_NEEDS: Record<string, string> = {
  NO_TAG: "Asset tag column", NO_LOCATION: "Location column", NO_COST_CENTER: "Cost center column", NO_PURCHASE_DATE: "Purchase date column",
  BAD_HOSTNAME: "HostnamePattern mapping", NOT_VERIFIED: "Last verified / audit date column", LONG_IN_STOCK: "Status date or AssetUpdateDate column",
  LONG_IN_REPAIR: "Status date or AssetUpdateDate column", LONG_IN_TRANSIT: "Status date or AssetUpdateDate column",
  NEW_NO_PO: "Previous export and PurchaseNo column", STATUS_REVERSED: "Previous export", LEAVER_HOLDS_ASSET: "Users reference",
  USER_NOT_FOUND: "Users reference", NOT_SEEN: "Devices reference", STALE_DEVICE: "Last scan column or Devices reference",
  ONLINE_NOT_IN_USE: "Last scan column or Devices reference", USER_MISMATCH: "Devices reference with primary user",
  WARRANTY_EXPIRED: "Warranty end column", OVER_AGE: "Purchase date column", STOCK_OVER_AGE: "Purchase date column",
  DUP_ASSET_ID: "Asset no. column", DUP_MAC: "MAC column", DUP_HOSTNAME: "Hostname column"
};

// ============================================================================ reconciliation sheets
const PO_SHEET = "POs Validated", SUMMARY_SHEET = "Pivot & Summary", EMAIL_SHEET = "Monthly Email";
// Monthly Email: the email to copy into Outlook and the prompt to copy into Copilot in PowerPoint, one line per row
// (copying several rows gives clean lines; one cell with line breaks would be copied with quotes around it)
function buildEmailSheet(wb: ExcelScript.Workbook, ctx: Ctx): void {
  progress("Writing Monthly Email");
  const rep = monthlyReport(ctx);
  const ws = newSheet(wb, EMAIL_SHEET, "Monthly Email and PowerPoint prompt", "Built from this run's numbers. Copy, paste, then " +
    "add your name. Check the numbers against Pivot & Summary before sending.");
  ws.getRangeByIndexes(0, 0, 1, 1).getFormat().setColumnWidth(30);
  ws.getRangeByIndexes(0, 1, 1, 1).getFormat().setColumnWidth(900);
  link(ws, 2, 0, "Open Pivot & Summary ›", SUMMARY_SHEET);
  let r = 4;
  const block = (title: string, how: string, lines: string[], bold: number[]): number => {
    putValues(ws, r, 0, [[title, ""], ["", how]]);
    const f = ws.getRangeByIndexes(r, 0, 1, 1).getFormat().getFont();
    f.setBold(true); f.setColor(BRAND); f.setSize(12);
    ws.getRangeByIndexes(r + 1, 1, 1, 1).getFormat().getFont().setItalic(true);
    ws.getRangeByIndexes(r + 1, 1, 1, 1).getFormat().getFont().setColor(GREY);
    const first = r + 2;
    putValues(ws, first, 1, lines.map(t => [t]));
    ws.getRangeByIndexes(first, 1, lines.length, 1).setNumberFormat("@");
    bold.forEach(i => ws.getRangeByIndexes(first + i, 1, 1, 1).getFormat().getFont().setBold(true));
    return first + lines.length + 2;
  };
  r = block("Email subject", "Copy the cell below into the email subject.", [rep.subject], []);
  r = block("Attachment name", "Save the workbook (or the zip) with this name.", [rep.attachment], []);
  const emailFirst = r + 2;
  r = block("Email body", `Select B${emailFirst + 1}:B${emailFirst + rep.email.length}, copy (Ctrl+C) and paste into the email. ` +
    `If Outlook shows a grid, use Paste Options > Keep Text Only. Replace each ${EMAIL_NOTE} with what the investigation found.`, rep.email, rep.heads);
  const prompt = rep.ppt.join(" ");
  const pptFirst = r + 2;
  r = block("Copilot in PowerPoint prompt", "Open a new presentation from the VOIS template, open Copilot, then select " +
    `B${pptFirst + 1}:B${pptFirst + rep.ppt.length}, copy and paste it as the prompt. ${prompt.length.toLocaleString("en-US")} characters` +
    (prompt.length > COPILOT_PROMPT_LIMIT ? ` - longer than ${COPILOT_PROMPT_LIMIT.toLocaleString("en-US")}: if Copilot cuts it, paste the first two rows, then ask it to add slide 3 with the last row.` : "."),
  rep.ppt, []);
}
const NEW_HEADER = ["Asset no.", "Serial no.", "Asset status", "Asset Type", "Model", "Purchase date", "PurchaseNo", "Last scan",
  "Scandate Age in Days", "Function Unit", "AssetCreationDate", "Creation Age in Days", "AssetUpdateDate"];
const REMOVED_HEADER = ["Asset no.", "Serial no.", "Previous Asset Status", "Asset Type", "Model", "Purchase date", "PurchaseNo",
  "Last scan", "Function Unit", "AssetCreationDate", "AssetUpdateDate", "Review Status", "Review Notes"];
const CHANGE_HEADER = ["Asset no.", "Serial no.", "Current Asset Status", "Previous Asset Status", "Status Change", "Asset Type",
  "Model", "Purchase date", "PurchaseNo", "Last scan", "Function Unit", "AssetCreationDate", "AssetUpdateDate"];
const PO_RECON_HEADER = ["Purchase Number", "Model", "Export Asset Count", "User-Provided Quantity", "Quantity Variance",
  "Physical Security Validation", "Spider Validation", "Reconciliation Result"];
const FMT_DATE = "dd-mmm-yyyy", FMT_DATETIME = "dd-mmm-yyyy hh:mm", FMT_COUNT = "#,##0", FMT_DIFF = "[Color10]+#,##0;[Red]-#,##0;0";

function newAssetRow(a: Asset): Cell[] {
  return [a.assetId, a.serial, a.status, a.assetClass, a.model, dateSerial(a.purchase), a.po, dateSerial(a.lastSeen), a.scanAge,
    a.functionUnit, dateSerial(a.created), a.createdAge, dateSerial(a.updated)];
}
// Optional fields stay blank when the export has no such column - never zero, N/A or invented
function buildReconSheets(wb: ExcelScript.Workbook, ctx: Ctx): string[] {
  const rc = ctx.recon, notes: string[] = [];
  const prevName = monthName(rc.prevPeriod), curName = monthName(rc.curPeriod);
  const widths13 = [110, 110, 110, 110, 150, 95, 110, 120, 80, 100, 120, 80, 120];

  // --- New Assets / tblNewAssets
  progress("Writing New Assets");
  let ws = newSheet(wb, "New Assets", `New Assets - ${curName}`, `In the ${curName} export but not in the ${prevName} export ` +
    `(matched on Asset no.): ${rc.newAssets.length}. ${rc.newAssets.filter(a => !a.po).length} have a blank PurchaseNo.`);
  backLink(ws);
  const newRows = rc.newAssets.slice().sort((a, b) => a.assetClass.localeCompare(b.assetClass) || a.po.localeCompare(b.po) ||
    a.assetId.localeCompare(b.assetId)).map(a => newAssetRow(a));
  addTable(ws, 3, "tblNewAssets", NEW_HEADER, newRows, { widths: widths13, text: [0, 1, 6, 9], date: [5], datetime: [7, 10, 12],
    count: [8, 11], freeze: true });

  // --- Removed Assets / tblRemovedAssets (only when there are any)
  progress("Writing Removed Assets");
  if (rc.removed.length) {
    ws = newSheet(wb, "Removed Assets", `Removed Assets - ${curName}`, `In the ${prevName} export but missing from the ${curName} ` +
      `export: ${rc.removed.length}. Not automatically scrapped - each needs review.`);
    backLink(ws);
    addTable(ws, 3, "tblRemovedAssets", REMOVED_HEADER, rc.removed.map(a => [a.assetId, a.serial, a.status, a.assetClass, a.model,
      dateSerial(a.purchase), a.po, dateSerial(a.lastSeen), a.functionUnit, dateSerial(a.created), dateSerial(a.updated), "Review required",
      "Missing from current export" + (rc.fromSnapshot ? " (details beyond status, type, model and PO are in the previous month's file)" : "")]),
    { widths: [110, 110, 120, 110, 150, 95, 110, 120, 100, 120, 120, 110, 260], text: [0, 1, 6, 8], date: [5], datetime: [7, 9, 10], freeze: true });
  } else {
    const old = wb.getWorksheet("Removed Assets");
    if (old) { old.delete(); }
  }

  // --- Serial Number Changes / tblSerialChanges
  progress("Writing Serial Number Changes");
  if (rc.serialChecked) {
    ws = newSheet(wb, "Serial Number Changes", `Serial Number Changes - ${curName}`, `Assets in both exports whose Serial no. differs ` +
      `from ${prevName} (case, spaces and punctuation ignored; placeholders such as N/A count as blank): ${rc.serialChanges.length}. ` +
      mapList(countBy(rc.serialChanges.map(c => c.change)), 3));
    backLink(ws);
    addTable(ws, 3, "tblSerialChanges", ["Asset no.", "Previous Serial no.", "Current Serial no.", "Change", "Asset status", "Asset Type",
      "Model", "PurchaseNo", "Function Unit", "AssetUpdateDate"], rc.serialChanges.map(c => [c.cur.assetId, c.prev.serial, c.cur.serial,
      c.change, c.cur.status, c.cur.assetClass, c.cur.model, c.cur.po, c.cur.functionUnit, dateSerial(c.cur.updated)]),
    { widths: [110, 130, 130, 170, 110, 110, 150, 110, 100, 120], text: [0, 1, 2, 7, 8], datetime: [9], freeze: true });
  } else {
    const old = wb.getWorksheet("Serial Number Changes");
    if (old) { old.delete(); }
  }

  // --- Scrap & Inactive Assets / tblStatusChanges
  progress("Writing Scrap & Inactive Assets");
  const scrapped = rc.operational.filter(c => compactKey(c.cur.status) === "scrapped").length;
  ws = newSheet(wb, "Scrap & Inactive Assets", `Scrap & Inactive Assets - ${curName}`, `Assets whose status changed to ` +
    `${OPERATIONAL_STATUSES.join(" or ")} since ${prevName}: ${rc.operational.length} (Scrapped ${scrapped}, Lost ` +
    `${rc.operational.length - scrapped}). ${rc.changes.length - rc.operational.length} other status changes are excluded (see Pivot & Summary).`);
  backLink(ws);
  addTable(ws, 3, "tblStatusChanges", CHANGE_HEADER, rc.operational.map(c => [c.cur.assetId, c.cur.serial, c.cur.status, c.prev.status,
    c.transition, c.cur.assetClass, c.cur.model, dateSerial(c.cur.purchase), c.cur.po, dateSerial(c.cur.lastSeen), c.cur.functionUnit,
    dateSerial(c.cur.created), dateSerial(c.cur.updated)]), { widths: [110, 110, 110, 120, 200, 110, 150, 95, 110, 120, 100, 120, 120],
    text: [0, 1, 8, 10], date: [7], datetime: [9, 11, 12], freeze: true });

  // --- POs Validated / tblPOValidated + reconciliation section
  progress("Writing POs Validated");
  ws = newSheet(wb, PO_SHEET, `POs Validated - ${curName}`, ctx.po.source ? `Supplied by the user (${ctx.po.source}); values kept exactly as ` +
    "supplied. Edit the table and run the script again to re-check." : "No PO validation supplied yet. Type the user's rows into this table " +
    "(or give them to the agent) and run again. Nothing here is ever filled in automatically.");
  backLink(ws);
  let r = sectionTitle(ws, 4, "1. PO validation rows (supplied by you)", "Your data, kept exactly as supplied; edit it here and run the script again");
  r = addTable(ws, r, "tblPOValidated", PO_HEADERS, ctx.po.rows.map(p => [p.date, p.number, p.model, p.quantity, p.raisedBy, p.physical, p.spider]),
    { widths: [110, 150, 150, 90, 130, 150, 130], text: [1], date: [0], count: [3], blankRow: true });
  r = sectionTitle(ws, r, "2. PO reconciliation", "Quantity Variance = User-Provided Quantity - Export Asset Count (new assets this month)");
  const poRows = poReconRows(ctx);
  r = addTable(ws, r, "tblPOReconciliation", PO_RECON_HEADER, poRows, { widths: [150, 150, 90, 100, 90, 150, 130, 190], text: [0],
    count: [2, 3], diff: [4] });
  const flags: string[][] = [];
  poRows.filter(x => x[7] === PO_RESULTS.notFound).forEach(x => flags.push(["PO not found among new assets", String(x[0])]));
  poRows.filter(x => x[7] === PO_RESULTS.missing).forEach(x => flags.push(["PO without complete validation", String(x[0])]));
  poRows.filter(x => x[7] === PO_RESULTS.variance).forEach(x => flags.push(["Quantity difference", `${x[0]}: ${signed(x[4] as number)}`]));
  openValidations(ctx).filter(x => x.indexOf("no validation supplied") < 0).forEach(x => flags.push(["Approval pending or missing", x]));
  poRows.filter(x => x[7] === PO_RESULTS.blank).forEach(x => flags.push(["New assets with empty PurchaseNo", String(x[2])]));
  ctx.po.problems.forEach(p => flags.push(["PO row to complete or correct", p]));
  r = sectionTitle(ws, r, "3. Flags for manual review", flags.length ? `${flags.length} to review` : "Nothing to review");
  addTable(ws, r, "tblPOFlags", ["Flag", "Details"], flags, { widths: [220], wrap: [1] });

  // --- Quality Checks (full audit only; the results are also in the flow result)
  if (REPORT_ONLY) { return notes; }
  progress("Writing Quality Checks");
  ws = newSheet(wb, "Quality Checks", `Quality Checks - ${curName}`, "Every control must pass (or be reviewed) before the monthly report is final");
  r = addTable(ws, 3, "tblQualityChecks", ["Check", "Expected", "Actual", "Result"], qualityChecks(ctx), { widths: [480, 110, 110, 80] });
  r = sectionTitle(ws, r, "Completion report");
  r = textRows(ws, r, reconText(ctx).map(t => [t]));
  if (rc.prevDup.length || rc.curDup.length) {
    r = sectionTitle(ws, r, "Asset numbers on more than one row");
    r = addTable(ws, r, "tblDupAssetNo", ["Export", "Asset no."], rc.prevDup.slice(0, 500).map(k => ["Previous", k] as Cell[])
      .concat(rc.curDup.slice(0, 500).map(k => ["Current", k] as Cell[])), { widths: [100, 160], text: [1] });
  }
  if (rc.unknownStatuses.size || rc.unlistedTypes.size) {
    r = sectionTitle(ws, r, "Values outside the controlled lists (both exports)", "Kept as they are; add a Mappings row (Status / Group) once confirmed");
    addTable(ws, r, "tblUnknownValues", ["Field", "Value", "Rows"], keysByCount(rc.unknownStatuses).map(k => ["Asset status", k, rc.unknownStatuses.get(k) || 0] as Cell[])
      .concat(keysByCount(rc.unlistedTypes).map(k => ["Asset Type", k, rc.unlistedTypes.get(k) || 0] as Cell[])), { widths: [120, 220, 80], count: [2] });
  }
  return notes;
}

// ============================================================================ agent input (markdown)
function mdTable(header: string[], rows: Cell[][]): string {
  const cell = (v: Cell) => String(v).replace(/\|/g, "/").replace(/\n/g, " ");
  return ["| " + header.join(" | ") + " |", "|" + header.map(() => "---|").join("")]
    .concat(rows.map(r => "| " + r.map(v => cell(v)).join(" | ") + " |")).join("\n");
}
function fmtCell(v: Cell, kind: string): Cell {
  if (v === "") { return ""; }
  if (kind === "pct") { return pct(Math.round((v as number) * 1000), 1000); }
  if (kind === "num") { return (v as number).toFixed(1); }
  return v;
}
function buildAgentInput(ctx: Ctx): string {
  const assets = ctx.assets, scope = assets.filter(a => inScope(a));
  const nowRows = ctx.history.filter(r => r.period === ctx.period && ctx.markets.indexOf(r.market) >= 0);
  const kNow = kpiValues(metricsFor(nowRows, "")), kPrev = kpiValues(metricsFor(prevRows(ctx, ""), ""));
  const prevAny = hasPrevious(ctx, "");
  const L: string[] = [`# Asset Audit Input - ${fmtDay(ctx.asOf)}`,
    `Source: VCAMS export | Markets: ${ctx.markets.join(", ")} | Asset records: ${assets.length} | In scope (not disposed): ${scope.length}`,
    ctx.focus ? `User's focus: ${ctx.focus}` : "",
    "All counts below are exact (calculated by the script). Do not recalculate, round differently or invent numbers.",
    "Integrity score = share of in-scope records with no completeness, validity, uniqueness, consistency or timeliness issue. " +
    "Reconciliation (users / device management) and lifecycle are reported separately and are not in the score.",
    `Users reference: ${ctx.users.source || "none (no leaver checks)"}. Devices reference: ${ctx.devices.source || "none (no device-management checks)"}.`,
    ctx.shareSafe ? "Share-safe mode: users appear as aliases (User-0001); never try to identify them." :
      "Real names and emails are included: share results with the requester and the service team only.",
    "", "## 0. Run information"].concat(ctx.runInfo.map(r => `- ${r[0]}: ${r[1]}`)).concat([""]);

  const rc = ctx.recon;
  if (rc.available) {
    L.push(`## A. Monthly reconciliation - ${monthName(rc.prevPeriod)} vs ${monthName(rc.curPeriod)} (completion report)`,
      `Matched on Asset no. only. Previous export: ${rc.source}.`);
    reconText(ctx).forEach(t => L.push("- " + t));
    L.push("", "Quality checks (report any Fail or Review; do not call the reconciliation complete while a check fails):",
      mdTable(["Check", "Expected", "Actual", "Result"], qualityChecks(ctx)), "");
    L.push("Assets by type (previous / difference / current):", mdTable(["Value group", "Asset Type", monthName(rc.prevPeriod), "Difference", monthName(rc.curPeriod)],
      Array.from(new Set<string>(Array.from(rc.prevByType.keys()).concat(Array.from(rc.curByType.keys())))).map(t => {
        const p = rc.prevByType.get(t) || 0, c = rc.curByType.get(t) || 0;
        return [typeOf(t, "", ctx.maps).group, t, p, signed(c - p), c] as Cell[];
      }).sort((a, b) => VALUE_GROUPS.indexOf(String(a[0])) - VALUE_GROUPS.indexOf(String(b[0])) || (b[4] as number) - (a[4] as number))), "");
    const tr = countBy(rc.changes.map(c => c.transition));
    L.push("Status changes by transition (only those ending in Scrapped or Lost are operational):",
      mdTable(["Status change", "Assets", "Operational"], keysByCount(tr).slice(0, 25).map(t => [t, tr.get(t) || 0,
        OPERATIONAL_STATUSES.some(s => t.endsWith("→ " + s)) ? "Yes" : "Excluded"])), "");
    if (rc.serialChanges.length) {
      L.push("Serial number changes by type (full list on the Serial Number Changes tab):", mdTable(["Asset Type", "Serial added", "Serial changed", "Serial removed"],
        keysByCount(countBy(rc.serialChanges.map(c => c.cur.assetClass))).map(t => [t as Cell].concat([SERIAL_ADDED, SERIAL_CHANGED, SERIAL_REMOVED]
          .map(ch => rc.serialChanges.filter(c => c.cur.assetClass === t && c.change === ch).length)))), "");
    }
    L.push("PO reconciliation (PO data is supplied by the user - never infer, invent or approve it):",
      mdTable(PO_RECON_HEADER, poReconRows(ctx)), ctx.po.problems.length ? "PO data to complete: " + ctx.po.problems.slice(0, 15).join("; ") : "", "");
  } else {
    L.push("## A. Monthly reconciliation", "Not run: only one export was supplied. Ask the user for the previous month's export to " +
      "list new, removed and newly scrapped / lost assets.", "");
  }

  L.push("## 1. Monthly KPIs", prevAny ? "Previous = each market's last audit." : "No previous audit: this is the baseline.");
  const nl = kpiList(kNow, ctx), pl = kpiList(kPrev, ctx);
  const per = ctx.markets.map(m => kpiList(kpiValues(metricsFor(nowRows.filter(x => x.market === m), "")), ctx));
  L.push(mdTable(["KPI", "This audit", "Previous", "Change"].concat(ctx.markets.length > 1 ? ctx.markets : []),
    nl.map((k, i) => [k[0], fmtCell(k[1], k[2]), prevAny ? fmtCell(pl[i][1], k[2]) : "",
      prevAny ? changeLabel(k[1], pl[i][1], k[2]) + (k[2] === "pct" ? " pts" : "") : ""]
      .concat(ctx.markets.length > 1 ? per.map(p => fmtCell(p[i][1], p[i][2])) : []))), "");

  L.push("## 2. Integrity by data-quality dimension (share of in-scope records with no issue)",
    mdTable(["Market", "Records", "Score"].concat(ALL_DIMS), ctx.markets.concat(ctx.markets.length > 1 ? [""] : []).map(m => {
      const list = m ? scope.filter(a => a.market === m) : scope, c = cleanCounts(list);
      return [m || "All markets", list.length, pct(c.clean, list.length)].concat(c.dims.map(n => pct(n, list.length)));
    })), "");

  L.push("## 3. Checks (all markets, most severe first)");
  const byCode = flagCounts(assets);
  // The market of every flagged asset, per check, in one pass (instead of one pass over all assets per check)
  const marketsByCode = new Map<string, string[]>();
  assets.forEach(a => a.flags.forEach(f => {
    const list = marketsByCode.get(f.code);
    if (list) { list.push(a.market); } else { marketsByCode.set(f.code, [a.market]); }
  }));
  const prevChecks = kPrev.checks;
  L.push(mdTable(["Check", "Severity", "Dimension", "Assets", "Share of in scope", "Previous", "Most affected markets", "Process-wide", "Fix", "Owner"],
    Array.from(byCode.keys()).sort((a, b) => sevRank(a) - sevRank(b) || (byCode.get(b) || 0) - (byCode.get(a) || 0)).map(c => {
      const n = byCode.get(c) || 0;
      return [FLAG_INFO[c].name, FLAG_INFO[c].sev, FLAG_INFO[c].dim, n, pct(n, scope.length), prevAny ? (prevChecks.get(c) || 0) : "",
        topList(marketsByCode.get(c) || [], 3), n > SETTINGS.systemicShare * scope.length && FLAG_INFO[c].dim !== LIFECYCLE ? "Yes" : "",
        FLAG_INFO[c].fix, FLAG_INFO[c].owner];
    })), "");
  if (ctx.markets.length > 1) {
    L.push("## 4. Checks by market", mdTable(["Check"].concat(ctx.markets), keysByCount(byCode).map(c => {
      const perMarket = countBy(marketsByCode.get(c) || []);
      return [FLAG_INFO[c].name as Cell].concat(ctx.markets.map(m => perMarket.get(m) || 0));
    })), "");
  }
  if (ctx.cols.has("location")) {
    const sites = Array.from(groupBy(scope, a => a.market + " | " + (a.location || "(no location)")).entries())
      .filter(e => e[1].length >= SETTINGS.siteMinRecords).map(e => [e[0], e[1].length, e[1].filter(a => !dataIssues(a).length).length / e[1].length,
        topList(flatten(e[1].map(a => dataIssues(a).map(f => FLAG_INFO[f.code].name))), 2)] as Cell[])
      .sort((a, b) => (a[2] as number) - (b[2] as number)).slice(0, 10);
    L.push("## 5. Lowest-scoring sites", mdTable(["Market | Site", "Records", "Score", "Main issues"], sites.map(s => [s[0], s[1], pct(Math.round((s[2] as number) * 1000), 1000), s[3]])), "");
  }
  L.push("## 6. Duplicates", `${ctx.dupGroups.length} groups of records share an identifier.`,
    mdTable(["On", "Value", "Records", "Assets", "Stages"], ctx.dupGroups.slice(0, 15).map(g => [g.on, g.value, g.assets.length,
      g.assets.map(a => a.key).join(", "), g.assets.map(a => a.stage).join(", ")])), "");
  const leavers = assets.filter(a => hasFlag(a, "LEAVER_HOLDS_ASSET"));
  L.push("## 7. Users and leavers", ctx.users.source ? `${leavers.length} assets are held by ${new Set<string>(leavers.map(a => userKey(a))).size} users who have left; ` +
    `${assets.filter(a => hasFlag(a, "USER_NOT_FOUND")).length} assets have a user not in the directory.` : "No user directory supplied.",
  leavers.length ? mdTable(["Asset", "Market", "Type", "User", "Left"], leavers.slice(0, 25).map(a => [a.key, a.market, a.assetClass, userLabel(a),
    (a.flags.find(f => f.code === "LEAVER_HOLDS_ASSET") as Flag).detail])) : "",
  `${ctx.manyDevices.length} users hold more devices of a type than allowed` + (ctx.manyDevices.length ? ": " + ctx.manyDevices.slice(0, 10)
    .map(g => `${g.user} (${g.assets.length} ${g.assetClass.toLowerCase()}s)`).join(", ") : "."), "");
  L.push("## 8. Reconciliation with device management", ctx.devices.source ? `Devices: ${ctx.devices.list.length}; matched to VCAMS: ` +
    `${ctx.devices.list.filter(d => d.matched).length}; in device management but not in VCAMS: ${ctx.orphans.length} ` +
    `(${ctx.orphans.filter(d => d.lastSeen && daysSince(d.lastSeen, ctx.asOf) <= SETTINGS.staleSeenDays).length} seen in the last ${SETTINGS.staleSeenDays} days).` :
    "No device-management export supplied.",
  ["NOT_SEEN", "STALE_DEVICE", "ONLINE_NOT_IN_USE", "USER_MISMATCH"].map(c => `- ${FLAG_INFO[c].name}: ${byCode.get(c) || 0}`).join("\n"),
  assets.filter(a => hasFlag(a, "ONLINE_NOT_IN_USE")).length ? "Online but recorded as not in use (security risk): " + assets.filter(a => hasFlag(a, "ONLINE_NOT_IN_USE"))
    .slice(0, 10).map(a => `${a.key} (${a.market}, ${(a.flags.find(f => f.code === "ONLINE_NOT_IN_USE") as Flag).detail})`).join("; ") : "", "");

  L.push("## 9. Stock and utilisation", `Stock age from ${ctx.stockAgeSource || "not available"}; deployments from ${ctx.deploySource || "not available"}. ` +
    "Cover = in stock / deployments per month.");
  L.push(mdTable(["Market", "Type", "In use", "In stock", "Utilisation"].concat(AGE_BUCKETS.map(b => "Stock " + b[0]), ["Deployed / month", "Months cover", "Beyond life", "Idle value", "What it means"]),
    utilisationRows(ctx, assets, a => [a.market, a.assetClass]).slice(0, 30).map(r => [r[0], r[1], r[2], r[3], fmtCell(r[4], "pct")]
      .concat(r.slice(5, 10), [r[11], fmtCell(r[12], "num"), r[13], r[14], r[15]]))));
  const moves = ctx.cols.has("location") ? moveSuggestions(ctx) : [];
  if (moves.length) { L.push("Stock moves before buying:", mdTable(["Market", "Type", "Short site", "Deployed / month", "Suggestion"], moves.slice(0, 15).map(m => [m[0], m[1], m[2], m[4], m[8]]))); }
  L.push("");
  const inUse = assets.filter(a => a.stage === IN_USE);
  const oldStock = countBy(assets.filter(a => hasFlag(a, "STOCK_OVER_AGE")).map(a => a.market + "\u0001" + a.assetClass));
  L.push("## 10. Lifecycle", mdTable(["Market", "Type", "In use", "Beyond expected life", "Out of warranty", "Stock beyond life"],
    Array.from(groupBy(inUse, a => a.market + "\u0001" + a.assetClass).entries()).map(e => {
      const [m, c] = e[0].split("\u0001");
      return [m, c, e[1].length, e[1].filter(a => hasFlag(a, "OVER_AGE")).length, e[1].filter(a => hasFlag(a, "WARRANTY_EXPIRED")).length,
        oldStock.get(e[0]) || 0] as Cell[];
    }).filter(r => (r[3] as number) + (r[4] as number) + (r[5] as number) > 0).sort((a, b) => (b[3] as number) - (a[3] as number)).slice(0, 20)), "");
  const pt = physicalText(ctx.physical);
  L.push("## 11. Physical check", pt || (ctx.physical.sample.length ? `A sample of ${ctx.physical.sample.length} assets is ready on the Physical Check tab; not yet checked.` : "Not used."), "");
  L.push("## 12. Example assets per check (for illustration; the full list is in the Exceptions tab)");
  keysByCount(byCode).filter(c => FLAG_INFO[c].sev !== "Low").slice(0, 10).forEach(c => {
    L.push(`- ${FLAG_INFO[c].name}: ` + assets.filter(a => hasFlag(a, c)).slice(0, SETTINGS.samplesPerCheck)
      .map(a => `${a.key} (${a.market}, ${a.assetClass}, ${a.stage}${(a.flags.find(f => f.code === c) as Flag).detail ? ", " + short((a.flags.find(f => f.code === c) as Flag).detail, 80) : ""})`).join("; "));
  });
  L.push("", "## 13. How to present this",
    "1. Monthly reconciliation first (section A): totals, net movement, new / removed assets, status changes, Scrapped and Lost, PO results, quality checks.",
    "2. Executive summary: integrity score and change, the 3-5 most important findings (high severity first), and the stock position.",
    "3. Monthly report: use the KPI table in section 1 as-is (this audit, previous, change, per market).",
    "4. Action plan: one line per finding with owner and the tab that lists the assets (Exceptions, Users & Leavers, Reconciliation, Stock & Utilisation).",
    "5. Never show names, emails or Raised By in the summary; refer to the tabs for detail.",
    "6. Say which checks were skipped because a column or reference was missing (section 0) - never imply a check passed when it did not run.");
  return L.filter((x, i) => x !== "" || L[i - 1] !== "").join("\n");
}
