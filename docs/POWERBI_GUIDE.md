# Monthly Spider Data Integrity: from script to Power BI, step by step

The target: each month you put the export in SharePoint, a flow runs the script and emails you the draft, and the
Power BI report (and the slides built on it) update by themselves. No PowerPoint built by hand, no Copilot needed.

```
 You                 Power Automate (standard)          SharePoint                         Power BI
 ───                 ─────────────────────────          ──────────                         ────────
 put the workbook ─▶ runs the Office Script        ─▶  02 Monthly Results/                ─▶ report refreshes:
 in 01 Input         saves "…- 2026-10.xlsx"           Spider Data Integrity - 2026-10.xlsx    4 pages + trends
                     emails you the draft email         (has the BI Data tab)                   (+ live slides in PowerPoint)
```

| Phase | What | Effort | Done when |
|---|---|---|---|
| 1 | SharePoint folders | 15 min | 3 folders exist |
| 2 | Run once by hand, save the result | 15 min | October's result file is in `02 Monthly Results` |
| 3 | Power BI report | 1-2 days | the 4 pages match the mockup |
| 4 | Publish, refresh, share, PowerPoint | 1 hour | managers open a link; slides update |
| 5 | Power Automate flow | half a day | putting a file in `01 Input` does everything |
| 6 | (Optional) Copilot agent for questions | later | only if people ask for it |

Start with phases 1-3 by hand: the report is what people see. The flow (phase 5) only removes clicks, so it comes after.

The mockup of the 4 pages is in `powerbi/mockup/` (`page1.png` … `page4.png`, sample numbers).

---

## Phase 1: SharePoint folders

In your team's SharePoint site, document library **Documents**, create:

```
Spider Data Integrity/
├── 01 Input/               the month's workbook (both exports) before the run      (team only - PII)
├── 02 Monthly Results/     one result file per month: "Spider Data Integrity - 2026-10.xlsx"   (team only - PII)
└── 03 Power BI/            the .pbix file and the theme
```

- Results hold names and emails (detail tabs), so keep both folders to the team. Power BI reads only the **BI Data** tab
  (counts, no names), so what managers see in the report has no PII.
- One file per month, always the same name pattern. Running a month again replaces that month's file, and Power BI
  also keeps only the newest file per month, so nothing is counted twice.

## Phase 2: one manual run (October)

1. Run the script as you do today (version **2026.10.19.1** or later: check the Glossary).
2. Check there is a **BI Data** tab with one table (`tblBIFacts`).
3. Save a copy as `02 Monthly Results/Spider Data Integrity - 2026-10.xlsx`.

Want trends from day one? Run the script on earlier months' workbooks too (e.g. August vs July, September vs August)
and save each as `Spider Data Integrity - 2026-08.xlsx`, `- 2026-09.xlsx`. Every month you add is one more point on
the trend lines.

## Phase 3: build the Power BI report

You need **Power BI Desktop** (free, from the Microsoft Store or Company Portal).

### 3.1 Get the data

1. **Home > Get data > Blank query**, then **Advanced editor**.
2. Paste `powerbi/Facts_query.pq`. Change `SiteUrl` to your site (e.g. `https://vodafone.sharepoint.com/sites/AssetTeam`)
   and keep `FolderName` = `02 Monthly Results`.
3. Sign in with your work account (**Organizational account**) when asked.
4. Rename the query **Facts**, then **Close & Apply**.

You get one table: Month, Month Date, Section, Category, Item, Metric, Value. Every number in the report comes from it.

### 3.2 Model and measures

1. **Modeling > New table**: paste the `Months` table from `powerbi/Measures.dax`.
   - Select `Months[Month]` > **Column tools > Sort by column > Month Sort**.
2. **Model view**: drag `Months[Month Date]` onto `Facts[Month Date]` (one to many, single direction).
3. **Home > Enter data**: create an empty table named `_Measures` (one column, then hide the column).
4. Add every measure from `powerbi/Measures.dax` (**New measure**, paste, Enter). Set the formats:
   - `%` measures (Utilization, Validated %, PO Match %…): **Percentage, 1 decimal**.
   - `Months Of Cover`: **Decimal, 1 decimal**.
   - counts: **Whole number, thousands separator**.

How the measures work: every measure shows **the selected month**, or the latest month when nothing is selected. On a
chart with `Months[Month Short]` on the axis, each point shows its own month. Utilization, cover and % validated are
worked out from the counts, never stored.

### 3.3 Theme and canvas

1. **View > Themes > Browse for themes** > `powerbi/VOIS_Executive_theme.json`.
   - VOIS red for the main numbers, teal / orange / violet only when a chart needs more series, grey for "in stock" and "last
     month". Green / amber / dark red only for status (validated, under investigation, not validated).
   - If Desktop says a property is not supported, it ignores it: the rest still applies.
2. **Format page > Canvas settings**: 16:9, **1280 × 720**.
3. Layout grid used on every page (View > Gridlines and Snap to grid help):
   - margins 24 px left / right; 12 px between visuals;
   - header 0-64 px; page tabs 72-100; KPI row 112-230; charts below.

### 3.4 The header (build once on page 1, then copy to every page)

| Item | Visual | Settings |
|---|---|---|
| Red bar | Shape (rectangle) | 6 × 36 px, fill #E60000, rounded 3, at x 24 y 14 |
| Page message | **Card** with the page's `Title …` measure | category label off, callout 20 pt Segoe UI Semibold #262626, left aligned, no background / border |
| Subtitle | Text box | 12 pt #595959, one line on what the page shows |
| Month | **Slicer** on `Months[Month]` | Dropdown, single select; **sync** it to all 4 pages (View > Sync slicers) |
| Logo | Image | the VOIS logo from your PowerPoint template |
| Page tabs | **Insert > Buttons > Navigator > Page navigator** | selected state: red text, white fill; others grey text |

### 3.5 Page 1: Overview (mockup `page1.png`)

| Area | Visual | Fields / measures |
|---|---|---|
| KPI 1 (red top line) | Card (new) | `Total Assets`, reference label `Net Change Label` |
| KPI 2 | Card (new) | `New Assets` |
| KPI 3 | Card (new) | `Status Changes`, reference label `Needs Attention 1` |
| KPI 4 | Card (new) | `PO Label`, reference label `PO Match %` |
| KPI 5 | Card (new) | `Notebook Utilization`, reference label `Cover Label` |
| Total assets, trend | Line chart | X: `Months[Month Short]`; Y: `Total Assets`; data label on the last point only; **Edit interactions**: the month slicer does **not** filter it |
| Change by value group | Clustered bar | Y: `Facts[Category]`; X: `Assets Change`; filter `Facts[Section]` = Assets by type |
| Biggest movers | Clustered bar | Y: `Facts[Item]`; X: `Assets Change`; filter Top N 3 by `Assets Change`; teal |
| Needs attention | 3 Cards | `Needs Attention 1`, `2`, `3` |

### 3.6 Page 2: New assets & POs (`page2.png`)

| Area | Visual | Fields / measures |
|---|---|---|
| KPIs | 4 Cards | `New Assets`; `POs`; `PO Label` + `PO Match %`; `New Without PO` |
| New assets by type | Clustered bar | Y: `Facts[Item]`; X: `New Assets By Type`; sorted by value |
| New assets per month | Clustered column | X: `Months[Month Short]`; Y: `New Assets`; slicer does not filter it |
| PO validation | Table | `Facts[Item]` (rename "PO"), `PO Quantity` ("On PO"), `PO Assets On Spider` ("On Spider"), `Facts[Category]` ("Result"); conditional formatting on Result: green text for Matched, amber for Pending / Missing, red for Quantity variance |

### 3.7 Page 3: Scrapped & inactive (`page3.png`)

| Area | Visual | Fields / measures |
|---|---|---|
| KPIs | 4 Cards | `Scrapped` (+ change vs `Scrapped Last Month`); `Inactive`; `Lost Stolen`; `Validation Label` + `Validated %` |
| Scrapped by asset type | Clustered bar | Y: `Facts[Item]`; X: `Status Changes By Type`; filter `Facts[Category]` = Scrapped |
| Main reasons | Clustered bar (teal) | Y: `Facts[Item]`; X: `Change Reasons` |
| Validation progress | Donut | values `Validated`, `Under Investigation`, `Not Validated`, `Not Checked Yet`; colours green / amber / dark red / grey |
| Per month | Stacked column | X: `Months[Month Short]`; Y: `Scrapped`, `Inactive`, `Lost Stolen`; slicer does not filter it |

### 3.8 Page 4: Stock utilization (`page4.png`)

| Area | Visual | Fields / measures |
|---|---|---|
| KPIs | 4 Cards | `Notebook Utilization` (+ `Cover Label`); `Wired Utilization`; `Wireless Utilization`; `Notebooks In Stock` (+ `Inactive With User` for Notebooks) |
| In use vs in stock | 100% stacked bar | Y: `Facts[Category]` (filter: Notebooks, Wired headsets, Wireless headsets); X: `In Use` (red), `In Stock` (grey) |
| Utilization trend | Line chart | X: `Months[Month Short]`; Y: `Utilization`; Legend: `Facts[Category]` (same 3); slicer does not filter it |
| By model | Table | `Facts[Item]` (Model), `Model In Use`, `Model In Stock`, `Model Utilization` (data bars, red), `Model Deployed From Stock`; filter `Facts[Category]` = Notebooks (a second copy, or a slicer, for headsets) |

### 3.9 Design rules (what makes it look executive)

- **One message per page**: the page title is the finding ("Notebook utilization 71.3%…"), not a label ("Stock").
- **4-5 KPIs at the top, 2-3 charts below, nothing else.** If a chart doesn't change a decision, drop it.
- **Numbers on the bars, no value axes, no gridlines** (the theme does this). Sorted largest first.
- **Colour has a job**: red = this month / the main number; grey = last month or in stock; green / amber / dark red =
  status only, always with a word ("Matched", "Open").
- **Same place on every page**: header, tabs, KPI row, charts. People learn where to look.
- **No more than 5 bars**: the rest goes in "Other".
- Turn off **visual headers** for viewers (Format > Visual header) on the published version for a clean look.

## Phase 4: publish, refresh, share, PowerPoint

1. **Publish** (Home > Publish) to your team's workspace.
2. In the Power BI service: the semantic model > **Settings > Data source credentials** > OAuth2 with your work account;
   **Scheduled refresh**: on, daily (it reads SharePoint Online, so no gateway is needed).
3. **Share**: create an **App** from the workspace (or share the report) with your managers. Viewers need a Power BI Pro
   licence, or the workspace on Premium / Fabric capacity.
4. **Row-level security**: not needed - the report holds counts only.
5. **PowerPoint**: in your VOIS template, **Insert > Add-ins > Power BI**, paste the page link. Each page becomes a live
   slide that refreshes when opened. Build the deck once (title slide + the 4 pages); every month, open it and present.
   (Or, in the report: **Export > PowerPoint > Embed live data**.)

## Phase 5: the Power Automate flow (standard connectors only)

Trigger: a file arrives in `01 Input`. Connectors: SharePoint, Excel Online (Business), Office 365 Outlook.

1. **Save the script**: Excel on the web > Automate > New script > paste `asset_audit_officescript.ts` > save as
   `asset_audit` (in your OneDrive `Documents/Office Scripts`, or a SharePoint `Scripts` folder).
2. **Create the flow**: Power Automate > Create > Automated cloud flow.
   - Trigger: **SharePoint - When a file is created in a folder**. Site: your site. Folder: `/Shared Documents/Spider Data Integrity/01 Input`.
3. **Excel Online (Business) - Run script** (or **Run script from SharePoint library** if the script is in SharePoint):
   - Location: your site; Document library: Documents; File: the trigger's **x-ms-file-id** (identifier).
   - Script: `asset_audit`. Leave the parameters empty except **shareSafe = No**.
   - Settings > **Retry policy**: exponential, 4 times, PT10S (a just-uploaded file can be locked for a few seconds).
4. Save the result under the month's name:
   - **SharePoint - Get file content**: the trigger file (it now holds the script's tabs).
   - **SharePoint - Create file**: folder `/Shared Documents/Spider Data Integrity/02 Monthly Results`, name
     `Spider Data Integrity - @{outputs('Run_script')?['body/result/period']}.xlsx`, content: the file content.
     Running a month again replaces that month's file. (If your site refuses to overwrite, add **Delete file** for that
     name before it, with "Configure run after" set to also run when it fails.)
5. **Office 365 Outlook - Send an email (V2)** to yourself:
   - Subject: `@{outputs('Run_script')?['body/result/emailSubject']}`
   - Body: `@{replace(outputs('Run_script')?['body/result/emailBody'], decodeUriComponent('%0A'), '<br>')}`
   - Add the link to the result file. You review, add the validation notes, and forward to the team.
6. **Test** with October's workbook: the result file appears in `02 Monthly Results`, the email arrives, and after the
   next Power BI refresh the report shows October.

Limit to know: **Run script** stops after about **120 seconds**. Locally the script needs ~10-15 s for 100k assets;
Excel on the web is slower. Test with the real workbook. If it times out, keep running the script by hand in Excel and
let the flow only copy, rename and email.

## Phase 6 (optional): Copilot agent

Only if managers want to ask questions in Teams ("how many wireless headsets in stock?"). Point a Copilot Studio agent
at the Power BI semantic model or the `02 Monthly Results` folder. The numbers stay the script's; the agent only reads
them. Leave this until the report is in use.

---

## Your monthly routine (once everything is set up)

1. Build the workbook with last month's and this month's exports (as today), and put it in `01 Input`.
2. Wait for the email (a few minutes). If you fill in PO or scrapped-asset validation, do it in the month's file in
   `02 Monthly Results` and run the script there again (it holds both exports), then send the email from the Monthly
   Email tab.
3. Open the Power BI report / the PowerPoint: it shows the new month.

## Questions

- **A month looks wrong in Power BI?** Open that month's result file > BI Data tab: Power BI shows exactly those numbers.
- **Run a month again?** Just run it: the month's file is replaced and Power BI keeps the newest.
- **Add a number to the report?** Tell whoever maintains the script: it is one line in `buildBiSheet` (a new Metric),
  then one measure in Power BI.
