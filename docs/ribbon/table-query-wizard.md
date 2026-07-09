# Table Query Wizard

**Ribbon:** `btnTableWizard` — “Table Query Wizard”  
**Legacy:** `ForceConnector.QueryTableWizard()` → `TableWizard.QueryWizard()` → `Operation.QueryData()`  
**Login required:** Yes (`CheckLoginAndAct`)  
**COM / VBA:** None

See also: [common-performance-requirements.md](./common-performance-requirements.md), [query-table-data.md](./query-table-data.md)

---

## Summary

Guided flow: Excel **InputBox** for the sheet anchor → WPF steps (object → fields → WHERE) → lay out the **ForceConnector table** → automatically run **Query Table Data**. The destination cell address is shown at the top of every WPF step.

---

## Functional requirements

### FR-TQW-1 Session

- If not authenticated, prompt login via `SessionGate` before the anchor prompt.
- On login failure, abort with clear message (no partial table).

### FR-TQW-2 Anchor cell (Excel InputBox, before WPF)

- Prompt user for top-left cell of the new table (legacy Excel `InputBox` type 8 range picker).
- Default suggestion: current selection’s top-left.
- Cancel → exit wizard with no sheet changes.
- There is **no** separate WPF “confirm anchor” screen — the chosen cell is shown as **Destination cell: {A1}** on subsequent steps.

### FR-TQW-3 Step 1 — Object selection

- Load global object list (`GET .../sobjects/` or equivalent).
- User filters/searches standard, custom, and system objects.
- Show destination cell at the top of the step.
- On confirm:
  - Write **object label** to anchor cell (row 1, col 1).
  - Store **API name** in cell comment on A1 (e.g. `Account`).
- **Back** disabled on this first WPF step; **Cancel** → exit.

### FR-TQW-4 Step 2 — Field selection

- `DescribeSObject` for chosen object.
- Multi-select fields; **Record Id** (Salesforce primary key) is **required**, **non-deselectable** in the UI, and always included in the result set.
- Show destination cell at the top.
- **Back** → object step; **Cancel** → exit.

### FR-TQW-5 Step 3 — Query clauses

- UI to build zero or more WHERE triplets: `[field] | [operator] | [value]`.
- Operators and value rules match [query-table-data.md](./query-table-data.md) (FR-QTD-3).
- If user adds no clauses, legacy default: `RECORD ID | not equals | (empty)` — equivalent to “all records” semantics via empty-id filter (preserve behavior or document intentional SOQL change).
- Show destination cell at the top.
- **Run Query** writes triplets into **row 1** starting at column B (field label in cell, API name in comment where needed).

### FR-TQW-6 Header layout (row 2)

Write column headers starting at the **same column as the row 1 object cell**. **Id is always column 0 of the field header block** (aligned with the object name on row 1). Order:

1. Id (always first; same column as row 1 object cell)  
2. Required fields (`nillable == false`, createable)  
3. Name fields  
4. Other standard fields  
5. Custom fields (`__c`)  
6. Read-only / non-updateable fields  

Per header cell:

- **Display:** field label  
- **Comment:** `API Name: {name}`, type, read-only/required hints, picklist values for picklists  

Field label → API name mapping for later query/update uses describe metadata — see [query-table-data.md](./query-table-data.md) FR-QTD-1.

### FR-TQW-7 Auto-query

- On successful criteria step, invoke the same pipeline as **Query Table Data** without requiring a second ribbon click.
- Row 3+ cleared/populated per query spec.

---

## Non-functional requirements

### NFR-TQW-1 Bulk sheet writes ([common](./common-performance-requirements.md))

**Legacy anti-pattern:** `drawField` sets `Value`, `WrapText`, and `AddComment` per column.

**Target:**

- Build header label array `object[1, fieldCount]`; one write to row 2.
- Comments/metadata: second pass acceptable (errors/metadata only), or store API names in row-2 comments in batch if Excel API allows — minimize COM crossings vs legacy.
- Picklist comment text: cap height / truncate very long value lists in UI with “see Salesforce” fallback.

### NFR-TQW-2 Describe caching

- Use the shared metadata cache ([NFR-META-1](./common-performance-requirements.md)) so `DescribeSObject` / object-list downloads are not repeated across wizard steps and the subsequent query.

### NFR-TQW-3 Threading

- Wizard is WPF on the UI thread; describe/query HTTP async; final sheet layout via `QueueAsMacro` in ExcelDna.

### NFR-TQW-4 Testability

- Core: field ordering comparator, clause DTO → row-1 cell model (no Excel).
- ExcelDna: integration smoke on Windows.

---

## Salesforce API

| Step | Call |
|------|------|
| 1 | List sobjects |
| 2–3 | Describe sobject |
| 7 | SOQL query (see [query-table-data.md](./query-table-data.md)) |

---

## Legacy reference

- `ForceConnector/TableWizard.cs`
- `frmWizardStep2.cs`, `frmWizardStep3.cs`, `frmWizardStep4.cs`
- `drawField` / `drawWizard`

## Improvements (allowed)

- Batch reference-join queries (inherits from query-table spec).
- Single WPF wizard window instead of legacy multi-step WinForms — see [ui-platform.md](./ui-platform.md).
