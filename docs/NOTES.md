# Project notes

## Goal

A Copilot Studio agent that runs the **monthly (or any-time) audit of the VCAMS asset inventory** for the VOIS asset team.

It automates the monthly asset reconciliation the team already does by hand:

- previous vs current export;
- new, removed and newly Scrapped / Lost assets;
- PO validation;
- the summary by asset type with PivotTables.

It adds the data-integrity checks, in-stock utilisation and lifecycle views that help in the same report.

The design follows [vois-ticket-review](https://github.com/khaledaboelnasr-gif/vois-ticket-review):

1. An Office Script is the engine.
2. A Power Automate agent flow runs it.
3. The Copilot Studio agent writes the narrative from the script's exact numbers (`agentInput`) and never recalculates.

## Hard constraints (VIC workstation and tenant)

- **Standard Power Automate connectors only.** No premium connectors, AI Builder, HTTP, custom connectors or Azure.
- **No code interpreter in Copilot Studio.**
- **Office Scripts are available.** They are the engine and run through the Excel Online "Run script" action.
- **Python is stdlib only.** No `pip install` without IT approval.
- **Office Script limits:**
  - no `any`;
  - array callbacks must be arrow functions;
  - about 120 s run time;
  - about 5 MB per request, so reads and writes are chunked.
- **Data:**
  - Exports contain PII; never commit them or paste them into chat.
  - Copilot Studio is VOIS-approved, so masking is not required there.
  - **No anonymisation** (user decision, 2026-10-12): the flow runs with `shareSafe=false`, so real names, emails and phones stay. Results are shared with the requester only, plus the service team. `shareSafe=true` remains available for reports that go wider.

## The monthly process spec, and where it is built

The spec is the "Monthly Asset Reconciliation Copilot Agent" instructions, written from last month's workbook (2026-10-14). Every section is implemented in `asset_audit_officescript.ts`. Real figures from that workbook (100,047 to 104,047 assets, 4,000 new Bags, 587 Scrapped / Lost) are only reference results and are **not** hard-coded.

| Spec section | Where |
|---|---|
| 2. Information to request | Agent topic (`docs/AGENT_GUIDE.md`): both exports or a stored snapshot, the dates, the PO validation rows. |
| 3. Preserve raw exports | Raw tabs are never written. All output goes to new sheets. IDs, serials, POs and hostnames are written as text. |
| 4. Header matching | `ASSET_COLS` (standard header first, then the accepted alternatives; case, spaces and punctuation ignored). A field matched by several columns is reported in "Columns with more than one match", and a Mappings row `Column` chooses one. A missing required field (Asset no., Asset status, Asset Type) stops the run with a clear message. |
| 5–6. Standard headers and formats | Detail tables use the exact 13 headers. Dates are `dd-mmm-yyyy`, date/times `dd-mmm-yyyy hh:mm`, counts `#,##0`, differences `+#,##0;-#,##0;0`. Blanks stay blank. |
| 7. Initial validation | Rows, unique asset numbers, blank and duplicate asset numbers per export ("Rows read", Quality Checks). |
| 8. Status normalisation | `VCAMS_STATUSES` (the 12 controlled statuses). Spacing and case are normalised for the comparison only; unknown statuses are kept and flagged. |
| 9. New assets | `New Assets` / `tblNewAssets`; blank PurchaseNo flagged (`NEW_NO_PO`); summaries by type, PO, model, blank / populated. |
| 10. Removed assets | `Removed Assets` / `tblRemovedAssets` (only when there are any), "Review required"; never treated as scrapped. |
| *(Team's workbook tab, not in the spec)* Serial Number Changes | `Serial Number Changes` / `tblSerialChanges`: assets in both exports whose serial was changed, added or removed. Case, spaces and punctuation are ignored, and placeholders such as N/A count as blank. Needs the previous export tab (the snapshot has no serials). |
| 11–12. Status comparison, Scrapped / Lost list | All changes are counted; `Scrap & Inactive Assets` / `tblStatusChanges` holds only genuine changes ending in Scrapped or Lost. The other changes are reported as excluded (transition table on Pivot & Summary). |
| 13–14. PO validation and reconciliation | `POs Validated` / `tblPOValidated`, holding user data only (poJson or a PO tab), plus the PO reconciliation (8 columns, the 6 approved results) and the flags. Nothing is inferred or approved by the script. |
| 15. Monthly summary | `Pivot & Summary`: every listed type (even when zero), High / Low value / Other totals, grand total. Live `SUMPRODUCT` formulas on the raw tabs, with a Check column against the script's count. |
| 16. PivotTables | `NewAssetsPivot`, `StatusChangesPivot`, `POValidationsPivot`, rebuilt every run. Skipped, with a note, when the source table is empty. |
| 17. Quality controls | `Quality Checks` (Pass / Review / Fail) and the completion report. Also in `agentInput` and the flow result. |
| 18. No external links | Everything is computed inside the workbook; no links to SharePoint workbooks. |
| 20–23. Anonymisation | `shareSafe=true`: deterministic tokens (USER-0001@example.invalid, AST-000001, SN-, HOST-, MAC, PO-, UNIT-, SITE-), the same in both exports, the mapping never stored, raw tabs removed. Off by default (user decision above). |
| 24. Completion report | `reconText()`: Quality Checks, Executive Summary, `agentInput` section A, the `reconciliation` flow output. |

## Key decisions

- **PO validation is managed on the POs Validated tab (2026-10-18).** Table 1 (tblPOValidated) is the input: dropdowns
  for the two validations, a whole-number check on Quantity, yellow while a cell is empty, and a to-do row (Purchase
  Number + asset type from Spider only) for every new-asset PO with no validation yet. The script never fills Quantity,
  Raised By, Purchase Date or a validation. Next run, the table is read back and merged with any PO tab (same Purchase
  Number + Model counts once; POs Validated wins); unfilled to-do rows are not PO data. poJson from the chat overrides
  both. Table 2 = result per PO (green matched, amber missing / pending, red variance / not found); table 3 = what to do.
- **Email 2.1** lists, under each PO, its models on Spider (type - model: count) and the PO lines not on Spider yet.

- **Stock utilisation = Notebooks and Headsets only (user decision, 2026-10-18).** Headsets = type "Headset" plus
  headset models filed under "Zubehör-Client" (model matches Headset / Jabra / Plantronics / Blackwire / Evolve / EPOS /
  Sennheiser / Voyager / Poly). In use = stage In Use; in stock = In stock, Reserved, Ready for use, To install, Building.
- **Pivot & Summary layout (2026-10-18).** Numbered headings on every table: 1 Key figures (each value a live formula on
  the detail tab it links to), 2 Assets by type, 3 Status changes (columns A-E); 4-7 New assets by type / PO / filled
  or empty / model (G-H); 8-10 the PivotTables side by side (J, M, P). Every table header is VOIS red with white text;
  every detail tab has a "Back to Pivot & Summary" link in row 3.

- **Report only (user decision, 2026-10-18).** The script makes what the team already makes and sends each month,
  plus stock utilisation, and nothing else: Monthly Email, Pivot & Summary, New Assets, Removed Assets (when any),
  Serial Number Changes, Scrap & Inactive Assets, POs Validated, Stock & Utilisation. Tabs from earlier full-audit
  runs are removed. The full audit (Executive Summary, Integrity Score, Exceptions, leavers, lifecycle, ...) is
  still in the code: set `REPORT_ONLY = false` to get it back. The quality checks still run and are returned to
  the flow, but are not written as a tab.

- **Real export layout (2026-10-14 run).** The current Spider export has 22 columns in its own order (ID, Asset no.,
  Asset Type, Model, Status, ..., Comment, Status comment, ..., Updated, Legal Entity); the previous one has the 13
  standard headers. Both are matched by header name. `Updated` = AssetUpdateDate, `Asset Assignment Date` =
  Deployment Date, `Status comment` / `Comment` feed the email's "Spider Comment".
- **Inactive and Stolen** are VCAMS statuses (they are in the team's monthly email), so they are in the controlled
  list (stages Inactive and Lost). Damaged and Delivered are still reported for review until confirmed.
- **Monthly email** follows the team's sent email (2026-09-15) section by section. The script never writes the
  investigation result: each listed asset ends with "[add validation note]".
- **Formulas**: calculation is paused during the run and a full recalculation runs at the end, so Pivot & Summary
  never shows 0 in the saved file.

- **Asset no. is the only reconciliation key.** It is never matched by row, serial, model or PO. Trimmed and case-insensitive.
- **Status stages.** Controlled statuses map to stages for the audit logic. The reports always show the VCAMS status.

  | Stage | Statuses |
  |---|---|
  | In Use | In operation |
  | In Stock | In stock, Reserved, Ready for use, To install, Building |
  | On Order | Ordered |
  | Inactive | Disabled, Authorised inactive |
  | Scrapped / Disposed | Scrapped, Marketed |
  | Lost | Lost |

  Meanings confirmed by the asset team (2026-10-14):
  - **Marketed:** donated to charity.
  - **Building:** being prepared.
  - **Disabled:** no network connection for more than 90 days, so VCAMS disables it.
  - **Authorised inactive:** the user is on long leave and keeps the asset.
  - **Ordered:** on order.

  Because of the 90-day Disabled rule, "not seen" (STALE_DEVICE) is 90 days. An *In operation* asset not scanned for longer should have been disabled.
- **Market** comes from the **Legal Entity** column. "VE" is Egypt (VOIS EG). Other entity codes are shown as they are until a Mappings row names them (`Market | <code> | <label>`, matching the whole value).
- **Data sources: Spider export only** (decision 2026-10-14). No HR, Entra or Intune data, so the leaver and device-management checks stay off. "Last scan" from the export drives the last-seen checks.
- **Previous month:** normally both exports in one workbook. The snapshot is only the fallback when a user uploads the current export alone.
- **Asset types** come from the spec's lists.
  - High value: Notebook, Desktop, Monitor, Zubehör-Client, Headset.
  - Low value: Bag, Mouse, Mobile, Keyboard.
  - Other: the remaining nine.
  - Each type also has a *kind* that drives the audit rules: whether it needs a user, a serial, device management, and what its lifecycle is.
  - Types that are not listed are shown under "Unlisted" and flagged.
- **Integrity score** is the share of in-scope (not scrapped) records with no completeness, validity, uniqueness, consistency or timeliness issue. Reconciliation (users, Intune) and lifecycle are reported separately, so the score stays comparable month to month.
- **Deployments per month** are exact when there is a previous export: assets that moved into *In operation* since then, plus new assets that are in operation. Without a previous export they are estimated from dates. Months of cover = in stock / deployments per month.
- **"Online but recorded as not in use"** only counts a scan *after* the status change, so a last scan just before scrapping is not flagged.
- **Two ways to give the previous month:**
  1. Both exports in one workbook (recommended; full detail).
  2. A compact snapshot that the flow stored last month: asset no., status, type, model, PO and market, about 30 characters per asset.
- **Scale.** Columns are read in chunks sized from a sample, and column by column for wide exports.
  - Exceptions are capped at 20,000 rows, most severe first.
  - All Assets is not repeated above 20,000 assets.
  - Quadratic patterns were removed: 107k assets plus 100k previous take about 11 s of script time locally.
  - Excel on the web adds read and write time. See FLOW_GUIDE section 6.

## Status (2026-10-14)

- **Engine 2026.10.14.3:** adds the Serial Number Changes sheet (the team's existing tab), the Legal Entity "VE" market, 90-day "not seen", and the team's tab names ("previous month export" / "this month export").
- **Engine 2026.10.14.2:** monthly reconciliation and the integrity audit, stock utilisation and lifecycle, KPIs and trend, PivotTables, quality checks, agent brief, snapshot in and out, share-safe tokens.
- **History store 2026.10.14:** History, Runs (with integrity score) and Mappings tables, append-only, newest complete run per month and market.
- **Reference extract:** Users and Devices workbooks to compact JSON.
- **Tests:** all three scripts type-check strict. The harness passes all modes:
  - tab and snapshot previous month;
  - PO as a tab, as JSON, or missing;
  - share-safe, with no real IDs or emails in the brief;
  - market filter;
  - wide and minimal exports;
  - 107k assets.

  The history harness passes too.
- **Not yet run in Excel on the web.** The next step is a run on the real September / October workbook.

## Open questions

Answered on 2026-10-14 in the build tracker: status meanings, market (Legal Entity), previous-month method, data sources (Spider only), and the monthly report tabs.

Still open:

1. **Extra columns.** Which columns does the export have beyond the 13 standard ones (assigned user / email, hostname, location, cost center)? They switch on the user and location checks. Tracker step 0.2 has no header list yet.
2. **Other Legal Entity codes.** What are the codes for IN and the EU markets? "VE" is mapped; add the others as Mappings rows.
3. **Date format** in the export. The script reads Excel dates, ISO, `dd-mmm-yyyy`, and day-first `d/m/yyyy`.

## Next steps

1. Run the engine by hand on the real September / October workbook in Excel on the web. Check the following against the reference figures, and note the run time:
   - Pivot & Summary;
   - New Assets;
   - Scrap & Inactive Assets;
   - Quality Checks.
2. Answer the open questions. Add Mappings rows for any status or type the script flags.
3. Build the SharePoint store and the agent flow (`docs/FLOW_GUIDE.md`) and the Copilot Studio agent (`docs/AGENT_GUIDE.md`).
4. Pilot with one market, get the IT data-handling check (DLP, SharePoint access, retention), then publish to Teams.

## Script entry point

```ts
main(workbook, auditDate?, previousDate?, historyJson?, marketFilter?, focus?, shareSafe?, mappingsJson?,
     usersJson?, devicesJson?, poJson?, previousSnapshot?, returnSnapshot?): AuditResult
```

`AuditResult` includes:

- `period` and `previousPeriod`;
- `reconciliation` (the completion report lines);
- `qualityChecks` (the checks that did not pass);
- `integrityScore`, `topChecks` and `stockUtilisation`;
- `runInfo`;
- `agentInput` (the markdown brief for the agent);
- `historyRows`;
- `snapshot`;
- `emailSubject`, `emailBody` and `pptPrompt` (the Monthly Email tab: the monthly email text and the Copilot in PowerPoint prompt; empty when there is no previous month).
