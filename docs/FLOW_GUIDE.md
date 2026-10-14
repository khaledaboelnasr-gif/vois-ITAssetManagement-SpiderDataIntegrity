# Build guide: SharePoint store and Power Automate agent flow

This guide sets up the cloud path:

1. The asset team member attaches the VCAMS workbook in the Copilot Studio agent and gives the PO validation rows in chat.
2. The agent calls an agent flow.
3. The flow runs the Office Script, stores the history and the snapshot, shares the workbook, and emails it.
4. The agent writes the completion report and the summary from the brief (`agentInput`).

**Standard connectors only:** SharePoint, Excel Online (Business) and Office 365 Outlook.

---

## 1. Before you start

- **The workbook (recommended):** one `.xlsx` holding:
  - the previous export, on a tab named e.g. `2026-09 Previous Export`;
  - the current export, e.g. `2026-10 Current Export`.

  This is what the team already builds each month.
  - Without "previous" / "current" in the names, the older tab is the previous export. "Older" means the date in the tab name, else the newest AssetUpdateDate in the data.
  - With **only the current export**, the flow uses the snapshot it saved last month (section 4, step 6). The snapshot covers asset no., status, type, model, PO and market only, so removed assets show fewer details.
- **PO validation:** supplied by the user. Either:
  - a `PO Validation` tab in the workbook with the headers *Purchase Date | Purchase Number | Model | Quantity | Raised By | Physical Security Validation | Spider Validation*; or
  - rows typed in the chat, which the agent passes as `PoJson`.

  The flow never invents or approves PO data.
- **Optional references**, stored in SharePoint and refreshed by the team:
  - `Reference/Users.xlsx`: HR / Entra ID export with email, enabled / status and leaving date. Enables leaver checks.
  - `Reference/Devices.xlsx`: Intune / SCCM export with device name, serial, last check-in and primary user. Enables device-management checks.
- **Flow owner:** a service account or shared owner with edit access to the site.
- **Data handling sign-off (IT):** a DLP policy allowing SharePoint + Excel Online (Business) + Office 365 Outlook, plus SharePoint permissions and retention. See section 6.

## 2. SharePoint layout

```
Asset Audit/
├── Scripts/      asset_audit.osts, history_store.osts, reference_extract.osts   (service team: edit)
├── Runs/         raw uploads while a run is in progress                          (service team only; PII)
├── Results/      <YYYY-MM>/<market>/AssetAudit_<market>_<YYYY-MM>_<RunId>.xlsx    (service team: edit)
│                 <YYYY-MM>/<market>/AssetAudit_<...>_<RunId>_brief.md            (the agent brief, for "show my audit")
├── Snapshots/    <YYYY-MM>_<RunId>.json                                          (service team only)
├── Reference/    Users.xlsx, Devices.xlsx                                        (service team only; optional)
└── AuditHistory.xlsx   tables "History", "Runs" and "Mappings"                   (service team only)
```

Break permission inheritance on `Runs/`, `Snapshots/`, `Reference/` and `AuditHistory.xlsx`. Requesters get a view link to their own result only.

**Nothing is ever overwritten.**

- Every run has its own **RunId**: time, requester and flow-run code. It appears in every file name and every History row.
- When several runs cover the same month and market, the trend uses the newest complete one. "Complete" means it has at least 90% of the asset records of the biggest run.

### AuditHistory.xlsx

Create an empty workbook. The `history_store` script creates the tables on its first run.

| Table | Columns | Notes |
|---|---|---|
| `History` | Period, Market, AssetClass, Metric, Value, UpdatedOn, RunId | One row per run + month + market + asset type + metric. Metrics are counts (`#records`, `#issues`, `#new`, `#scrapped`, `#deployed`, and one per check), so scores and rates are recomputed, never averaged. Period is stored as text. |
| `Runs` | RunId, UpdatedOn, Requester, Markets, Periods, Assets, IntegrityScore, File | One line per run. |
| `Mappings` (optional, by hand) | Kind, Value, MapsTo, Note | Teaches the script your values without code changes. Examples follow. |

Example `Mappings` rows:

```
Status  | Marketed            | Scrapped / Disposed
Group   | Docking Station     | High value
Class   | Docking Station     | Docking Station
Column  | Asset status        | Current Status
Setting | stockAgedDays       | 120
Check   | NOT_VERIFIED        | Off
```

The engine's Glossary tab lists every kind.

## 3. Save the three scripts

1. In Excel on the web, go to **Automate > New script**. Paste `asset_audit_officescript.ts` and save it as **asset_audit**.
2. Do the same for `history_store_officescript.ts` (**history_store**) and `reference_extract_officescript.ts` (**reference_extract**).
3. Copy the three `.osts` files from the owner's OneDrive `Documents/Office Scripts` to `Asset Audit/Scripts/`.

When a script changes, repeat the steps for that script and overwrite its `.osts` file.

## 4. Build the agent flow

In Copilot Studio, go to the agent > **Tools > Add a tool > New agent flow**. Rename each action exactly as in **bold**, because the expressions refer to those names.

> **Time.**
> - Copilot Studio waits about 100 s for **Respond to the agent**.
> - A full VCAMS workbook (two exports of ~100k rows) can take longer than that in Excel on the web.
> - So this flow **responds first** ("started") and emails the result when it is ready.
> - The agent then fetches the brief with the second flow (section 5).
> - For one small market you can move **Respond to the agent** after step 9 and return `agentInput` directly.

### Trigger: When an agent calls the flow

| Input | Type | Required | Filled by the agent with |
|---|---|---|---|
| ExportFile | File | Yes | The workbook from the user |
| AuditDate | Text | No | `YYYY-MM-DD` the current export represents |
| PreviousDate | Text | No | `YYYY-MM-DD` the previous export represents |
| Market | Text | No | e.g. `VOIS EG`; empty = all |
| PoJson | Text | No | The PO rows from the chat as JSON (see AGENT_GUIDE); empty = use the workbook's PO tab |
| Focus | Text | No | What the user wants to focus on |
| RequesterEmail | Text | Yes | `System.User.Email` |

### Step 1: Compose **RunId**

```
concat(formatDateTime(utcNow(), 'yyyyMMdd-HHmmss'), '_',
       toLower(first(split(<RequesterEmail>, '@'))), '_',
       substring(workflow()?['run']?['name'], sub(length(workflow()?['run']?['name']), 6), 6))
```

### Step 2: Condition **Is xlsx**

The condition is `endsWith(toLower(<ExportFile name>), '.xlsx')`.

- **No branch:** **Respond to the agent** with `error` = "Please upload the VCAMS workbook as .xlsx", then **Terminate**.
- **Yes branch:** a **Scope** named **Try** with steps 3–10.

### Step 3: SharePoint, Create file → **Save upload**

`/Asset Audit/Runs/@{outputs('RunId')}.xlsx`, content `<ExportFile contentBytes>` (or `base64ToBinary(...)`).

### Step 4: **Respond to the agent** (early)

| Output | Value |
|---|---|
| status | `started` |
| runId | `outputs('RunId')` |
| message | "The audit has started. You'll get an email with the workbook in a few minutes; then ask me to show the result." |

### Step 5: Excel Online (Business), Run script from SharePoint library → **Read history**

- Workbook: `AuditHistory.xlsx`. Script: `Scripts/history_store.osts`.
- Inputs: `mode` = `read`, `months` = `13`, `market` = `<Market>`.

### Step 6 (optional): previous-month snapshot → **Latest snapshot**

Use this only when users upload the current export alone.

1. **SharePoint, Get files (properties only)** in `Snapshots/` → **Snapshot files**: Order By `Created desc`, Top Count `24`.
2. **Filter array** (Data Operations) → **Latest snapshot**:
   - From: `body('Snapshot_files')?['value']`;
   - Condition: `@less(substring(item()?['{Name}'], 0, 7), substring(coalesce(<AuditDate>, utcNow()), 0, 7))`.

   This keeps only months before the one being audited. The first item is the newest.
3. **Get file content** of `first(body('Latest_snapshot'))?['{Identifier}']` → **Snapshot content**. Wrap it in a Condition `not(empty(body('Latest_snapshot')))`.
   - Expression for the script: `if(empty(body('Latest_snapshot')), '', base64ToString(body('Snapshot_content')?['$content']))`.
   - The script ignores the snapshot when the workbook already has a previous export tab.

### Step 7 (optional): references → **Users ref**, **Devices ref**

Run `Scripts/reference_extract.osts`:

- on `Reference/Users.xlsx` with `kind` = `users`;
- on `Reference/Devices.xlsx` with `kind` = `devices` and `market` = `<Market>`.

Use the `json` output of each. Keep the total parameters under ~4 MB (section 6). For a full-company run with a snapshot, leave the references out or run per market.

### Step 8: Excel Online (Business), Run script from SharePoint library → **Run audit**

Workbook: the Id from **Save upload**. Script: `Scripts/asset_audit.osts`.

| Parameter | Value |
|---|---|
| auditDate | `<AuditDate>` |
| previousDate | `<PreviousDate>` |
| historyJson | `body('Read_history')?['result']?['historyJson']` |
| marketFilter | `<Market>` |
| focus | `<Focus>` |
| shareSafe | `No` |
| mappingsJson | `body('Read_history')?['result']?['mappingsJson']` |
| usersJson / devicesJson | the `json` outputs of step 7, or empty |
| poJson | `<PoJson>` |
| previousSnapshot | the expression from step 6, or empty |
| returnSnapshot | `Yes` |

**Retry policy:** Exponential, count 4, interval PT10S, because a just-created file can be locked.

### Step 9: Save and share the result

1. **Compose MarketLabel:** `if(equals(length(body('Run_audit')?['result']?['markets']), 1), first(body('Run_audit')?['result']?['markets']), 'All markets')`
2. **Compose ResultName:** `concat('AssetAudit_', outputs('MarketLabel'), '_', body('Run_audit')?['result']?['period'], '_', outputs('RunId'))`
3. Save the files:
   - **Get file content** (Save upload Id) → **Result content**.
   - **Create file → Save result:** `/Asset Audit/Results/@{body('Run_audit')?['result']?['period']}/@{outputs('MarketLabel')}/@{outputs('ResultName')}.xlsx`.
   - **Create file → Save brief:** same folder, `@{outputs('ResultName')}_brief.md`, content `body('Run_audit')?['result']?['agentInput']`.
   - **Create file → Save snapshot:** `/Asset Audit/Snapshots/@{body('Run_audit')?['result']?['period']}_@{outputs('RunId')}.json`, content `body('Run_audit')?['result']?['snapshot']`.
4. **Grant access** to the result for `<RequesterEmail>` (Can view, no notification).
5. **Get file properties → Share link** of **Save result**.

### Step 10: History, email, clean-up

1. Run `history_store` with:
   - `mode` = `save`
   - `rowsJson` = `string(body('Run_audit')?['result']?['historyRows'])`
   - `runId`
   - `requester`
   - `fileLink` = the share link

   A retry is safe: the same RunId is never saved twice.
2. **Send an email (V2)** to `<RequesterEmail>`.
   - Subject: `Asset audit @{body('Run_audit')?['result']?['period']} – @{outputs('MarketLabel')}`
   - Body: the link, then `join(body('Run_audit')?['result']?['reconciliation'], '<br>')`, then `join(body('Run_audit')?['result']?['qualityChecks'], '<br>')`.
   - Attachment: the result content.
3. **Delete file** in `Runs/`.

### Step 11: Error path

Add a **Send an email (V2)** after **Try**, with *has failed / has timed out*. Send the error, `coalesce(outputs('Run_audit')?['body']?['error']?['message'], 'see the flow run history')`, to the requester.

Common errors:

- "No VCAMS asset export found": no Asset no. / serial column.
- "The current export has no Asset Type column": add the column, or a Mappings `Column` row.

## 5. Second flow: Get latest audit

The agent calls this when the user asks "show my audit" or "what did the audit find?".

| Part | Content |
|---|---|
| Inputs | RequesterEmail (`System.User.Email`), Period (optional), Market (optional) |
| Steps | **SharePoint, Get files (properties only)** in `Results/`, including nested folders. Filter `substringof('_@{toLower(first(split(<RequesterEmail>,'@')))}_', FileLeafRef) and substringof('_brief.md', FileLeafRef)`. Order By `Modified desc`, Top Count 1. Then **Get file content**, then **Respond to the agent** with `agentInput` = the file text, plus `workbookLink`. |

If nothing is found yet, respond with "Still running – try again in a few minutes".

## 6. Limits and risks

| Item | Limit or risk | Mitigation |
|---|---|---|
| Run time | ~120 s per Run script action | About 11 s of script time locally for 107k + 100k assets. Excel on the web adds read and write time. Mitigations: Exceptions capped at 20k rows; All Assets not repeated above 20k assets; wide exports read column by column. **Measure on the real workbook first.** If it is too slow, run per market or by hand in Excel. |
| Agent wait | ~100 s | The flow responds first and emails later; the "Get latest audit" flow returns the brief. |
| Parameters / result | ~5 MB per request | `agentInput` is about 30 KB. The snapshot is ~30 characters per asset (~3 MB at 100k). The users / devices JSON is ~80 characters per row. Keep the total under ~4 MB, or run per market. |
| Two runs at once | Both write AuditHistory.xlsx | Rows are only added; the retry policy handles the lock. |
| Personal data | Names and emails stay in results (user decision) | Access is the requester + service team only. Set `shareSafe = Yes` for anything shared wider. |
| Retention | Uploads, results and snapshots hold asset and personal data | Restrict the folders and agree retention with IT. |
