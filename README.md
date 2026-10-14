# VOIS Asset Audit (VCAMS / Spider data integrity)

An Excel Office Script that runs the **monthly asset reconciliation** and a **data-integrity audit** of the VCAMS (Spider)
asset inventory, and builds a formatted workbook for the asset team plus a brief for a Copilot Studio agent.

The workbook covers:

- **Monthly reconciliation:** the previous and current exports compared on Asset no., giving:
  - new assets;
  - removed assets;
  - serial number changes;
  - status changes;
  - the operational list of assets that became **Scrapped** or **Lost**;
  - the user-supplied **PO validation** and the PO reconciliation;
  - **Pivot & Summary** by asset type (High value / Low value / Other), with PivotTables;
  - the **quality checks** and completion report.
- **Data integrity:** completeness, validity, duplicates, consistency and timeliness, with an integrity score.
  - With a user directory, it also finds leavers who still hold assets.
  - With an Intune / SCCM export, it also finds devices that are missing from device management, and devices that are online although VCAMS says they are scrapped.
- **In-stock utilisation:**
  - utilisation rate and stock ageing;
  - deployments per month and months of cover;
  - idle stock and moves between sites;
  - stock that is past its life.
- **Lifecycle:** over-age and out-of-warranty assets, and the refresh forecast.
- **Monthly KPIs and trend:** this audit compared with the previous one, per market (from a SharePoint history).

It follows the same pattern as [vois-ticket-review](https://github.com/khaledaboelnasr-gif/vois-ticket-review):

1. An Office Script is the engine.
2. A Power Automate agent flow runs it with standard connectors only.
3. A Copilot Studio agent writes the summary from the script's exact numbers.

## Files

| Path | What it is |
|---|---|
| `asset_audit_officescript.ts` | **The engine.** Paste it into Excel on the web > Automate > New script. |
| `history_store_officescript.ts` | Reads and saves the month-on-month history and run log in SharePoint `AuditHistory.xlsx`, plus the central `Mappings` table. |
| `reference_extract_officescript.ts` | Turns a Users (HR / Entra ID) or Devices (Intune / SCCM) workbook into compact JSON for the engine. |
| `docs/NOTES.md` | Constraints, decisions, how the monthly-process spec maps to the build, open questions, next steps. **Read this first.** |
| `docs/FLOW_GUIDE.md` | SharePoint layout and the Power Automate agent flow, step by step. |
| `docs/AGENT_GUIDE.md` | Copilot Studio agent: instructions, topic, and how it handles PO data. |
| `docs/SCRIPT_OVERVIEW.txt` | Plain-language overview of the tabs and checks, for stakeholders. |
| `test/gen_synthetic.py` | Generates fake previous and current exports, a PO tab, users and Intune data (no real data). |
| `test/harness.mjs` | Local test bench: runs the engine in Node 22.6+ against a mock ExcelScript API. |
| `test/history_harness.mjs` | Test bench for the history store and the reference extract, chained with the engine. |
| `test/excelscript.d.ts` | Minimal ExcelScript type stub for `tsc` checks. |

## Run it by hand

1. Put the exports in one workbook:
   - the previous export on a tab named, for example, `2026-09 Previous Export`;
   - the current export on a tab named, for example, `2026-10 Current Export`;
   - optionally, a `PO Validation` tab with the user's PO rows.
2. In Excel on the web, go to **Automate > New script**, paste `asset_audit_officescript.ts`, and run it.
3. Start with **Start Here**, then **Executive Summary** and **Quality Checks**.

## Test locally

```bash
cd test
python3 gen_synthetic.py > synthetic.json            # 5,000 assets; add a number for more, --minimal for the 13 standard columns only
node harness.mjs ../asset_audit_officescript.ts synthetic.json
DUMP="Pivot & Summary:60,Quality Checks:45" node harness.mjs ../asset_audit_officescript.ts synthetic.json
node history_harness.mjs ../history_store_officescript.ts ../asset_audit_officescript.ts ../reference_extract_officescript.ts synthetic.json
```

`harness.mjs` reads these environment options:

| Option | Effect |
|---|---|
| `DUMP="Sheet:rows,..."` | Prints the first rows of the named sheets. |
| `ONLYCUR=1` | Drops the previous export. |
| `SNAPSHOT=1` | Passes the previous month as a snapshot from an earlier run, instead of as a tab. |
| `NOPO=1` / `POJSON=1` | Drops the PO tab, or passes the PO rows as `poJson`. |
| `NOREF=1` / `REFJSON=1` | Drops the Users and Intune tabs, or passes them as JSON. |
| `SHARE=1` | Runs in share-safe mode (deterministic tokens). |
| `MARKET="VOIS EG"` | Filters the run to one market. |
| `MAPPINGS=<json>` | Passes custom mappings. |
| `MARK=1` | Fills the Physical Check sample and re-runs. |
| `HISTOUT=<file>` | Saves `historyRows` to a file. |
| `NORERUN=1` | Skips the re-run on the same workbook. |

An optional third argument is a history JSON file.

## Type-check

```bash
cd test
npm install
cp ../asset_audit_officescript.ts script.ts && npx tsc --noEmit --strict --target es2017 script.ts excelscript.d.ts
cp ../history_store_officescript.ts store.ts && npx tsc --noEmit --strict --target es2017 store.ts excelscript.d.ts
cp ../reference_extract_officescript.ts extract.ts && npx tsc --noEmit --strict --target es2017 extract.ts excelscript.d.ts
```

All three must type-check with no errors (strict, no `any`).

Then run the Office Scripts rules that `tsc` does not check. Array method callbacks must be inline arrow functions, and explicit `any` is not allowed:

```bash
node lint_officescript.mjs ../asset_audit_officescript.ts ../history_store_officescript.ts ../reference_extract_officescript.ts
```

## Data rules

- Real VCAMS, HR and Intune exports contain personal data. **Never commit them, and never paste them into chat.**
- `.gitignore` blocks `*.xlsx`, `*.csv`, `data/`, `real/` and the generated test JSON.
- Test only with the synthetic data in this repo.
