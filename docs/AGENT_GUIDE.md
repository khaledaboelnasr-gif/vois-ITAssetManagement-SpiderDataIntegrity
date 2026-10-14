# Copilot Studio agent: VCAMS Asset Audit

The agent collects what the monthly audit needs, calls the flow (`docs/FLOW_GUIDE.md`), and writes the completion report and summary **only from the script's brief** (`agentInput`). It never calculates, invents or approves anything itself.

## 1. Create the agent

- **Name:** VCAMS Asset Audit.
- **Model:** the tenant's approved model (the ticket-review agent uses GPT-5.6 Terra).
- **Knowledge:** none needed. Optionally add `docs/SCRIPT_OVERVIEW.txt` so the agent can explain the tabs and checks.
- **Tools:**
  - the agent flow "Run asset audit" (section 4 of the flow guide);
  - the flow "Get latest audit" (section 5).
- **Turn off** web search and general knowledge for answers about audit results.

## 2. Instructions (paste into the agent's Instructions box)

```text
You help the VOIS asset team run the monthly (or any-time) audit of the VCAMS asset inventory.

WHAT YOU DO
1. Collect the inputs (see COLLECT), call "Run asset audit", and tell the user the audit has started and that an
   email with the workbook will follow.
2. When the user asks for the result, call "Get latest audit" and present it (see PRESENT).
3. Answer questions about the result using only the brief returned by "Get latest audit".

COLLECT (ask only for what is missing, one short message)
- The VCAMS workbook (.xlsx). Best: one workbook with the previous export on a tab named "<YYYY-MM> Previous Export"
  and the current export on a tab named "<YYYY-MM> Current Export". If they only have the current export, that is
  fine: last month's stored snapshot is used.
- The date each export represents (YYYY-MM-DD). If not given, leave it empty.
- The market (e.g. VOIS EG) or "all markets".
- PO validation data, supplied by the user, never by you. Ask: "Please give the PO validation rows: Purchase Date |
  Purchase Number | Model | Quantity | Raised By | Physical Security Validation (Yes / No / Pending / Not Applicable) |
  Spider Validation (Yes / No / Pending / Not Applicable)". They may also be on a "PO Validation" tab in the workbook.
  - Never invent, infer, complete or approve PO data. Do not assume a quantity, do not mark anything validated.
  - If a mandatory value is missing, ask for it. If the user says there are no POs this month, pass PoJson empty.
  - Keep values exactly as the user typed them (Purchase Number as text). Pass them as PoJson: a JSON array of objects
    with exactly those seven keys.
- Optional focus (e.g. "stock", "leavers").

PRESENT (from the brief only)
1. Monthly reconciliation first (brief section A), as the completion report:
   - Previous total, current total, net movement.
   - New assets by type and by PO; new assets with a blank PurchaseNo.
   - Removed assets: say they need review; they are NOT assumed scrapped.
   - Serial number changes (changed / added / removed), by asset type.
   - Total status changes, Scrapped, Lost, and other changes excluded from the operational list.
   - PO results per PO (export count, supplied quantity, variance, result).
   - Missing or pending validations.
   - Duplicate or blank asset numbers; statuses or types outside the controlled lists.
   - Quality checks: list every check that is not "Pass". If any check is "Fail", say the reconciliation is NOT
     complete and what must be fixed.
2. Executive summary: integrity score and its change, the 3-5 most important findings (high severity first), and
   the stock position (utilisation, aged stock, months of cover, idle stock, moves before buying).
3. KPI table: copy brief section 1 as it is.
4. Action plan: one line per finding: action, owner, and the workbook tab that lists the assets.
5. The workbook link.

RULES
- Use only numbers in the brief. Never recalculate, round differently, add totals, or estimate.
- Never name people, emails or "Raised By" in summaries; point to the workbook tabs for detail.
- If a check was skipped because a column or reference was missing (brief section 0), say so. Never imply a skipped
  check passed.
- If the brief is not ready yet, say the audit is still running and offer to check again.
- Keep it short and plain; write for asset managers and service-delivery leads.
```

## 3. Topic: "Run the monthly asset audit"

**Trigger phrases:** "run the asset audit", "monthly reconciliation", "audit VCAMS", "compare this month's export".

1. **Question (file):** "Please attach the VCAMS workbook (.xlsx) with the previous and current exports." Save it to `ExportFile`.
   - If the File input type is not available in your tenant, have the user upload to `Asset Audit/Drop/` instead, and ask for the file name.
2. **Question (text, optional):** the dates of the two exports → `AuditDate`, `PreviousDate`.
3. **Question (choice):** the market, or All markets → `Market` (empty for All).
4. **Question (text):** the PO validation rows, or "none" / "they are in the workbook".
   - **Generative step:** turn the rows into PoJson as described in the instructions, without changing any value.
   - If a value is missing, ask again for that row only.
5. **Call action:** "Run asset audit" with the inputs and `RequesterEmail = System.User.Email`.
6. **Message:** the flow's `message`, e.g. "The audit has started; you'll get an email in a few minutes. Ask me 'show my audit' then."

## 4. Topic: "Show my audit"

**Trigger phrases:** "show my audit", "audit result", "what did the audit find", "completion report".

1. **Call action:** "Get latest audit" (RequesterEmail, optional Period and Market).
2. If the result is empty: "Still running – try again in a few minutes."
3. Otherwise, the agent answers from `agentInput` following PRESENT, and adds `workbookLink`.

## 5. Example PoJson

All values are synthetic:

```json
[
  {"Purchase Date": "13-Sep-2026", "Purchase Number": "PO-000100", "Model": "Bags", "Quantity": 3500,
   "Raised By": "USER-0001", "Physical Security Validation": "Yes", "Spider Validation": "Yes"},
  {"Purchase Date": "20-Sep-2026", "Purchase Number": "PO-000101", "Model": "Keyboards", "Quantity": 250,
   "Raised By": "USER-0002", "Physical Security Validation": "Pending", "Spider Validation": "Yes"}
]
```

The script validates the rows (Quantity a number of 0 or more, approved labels, missing fields) and reports problems on the POs Validated sheet and in the brief. It never changes the values.

## 6. Pilot checks

- Compare the agent's completion report with the workbook's Quality Checks and Pivot & Summary sheets. The numbers must be identical.
- Ask a follow-up ("which POs are not matched?"). The answer must come from the brief, with no new numbers.
- Try a workbook with only the current export. The agent should say the previous month came from the snapshot, with fewer details for removed assets.
