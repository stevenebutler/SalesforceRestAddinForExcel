# Table Query Wizard

**Ribbon:** `btnTableWizard` — “Table Query Wizard”  
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

- Prompt user for top-left cell of the new table (Excel `InputBox` type 8 range picker).
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
- The field picker is a sortable grid with columns `Category`, `Label`, and `API Name`.
- `Category` is shown as a one-based visual label such as `1 Id`, `2 Name`, etc. The underlying bucket order remains zero-based in Core.
- Category cells keep their bucket fill and black text even when the row is selected.
- Field list display remains `{Label} ({ApiName})` in the row data, but the grid exposes the label and API name separately for sorting and filtering.
- The old side legend was removed; bucket information is encoded directly in the category column.
- Show destination cell at the top.
- **Back** → object step; **Cancel** → exit.

Grid sorting follows the shared rule used by other WPF pickers:

- clicking a non-first column promotes it to the front and sorts ascending the first time
- clicking the current first column toggles that column between ascending and descending
- clicking another column preserves the remaining secondary sort keys

### FR-TQW-5 Step 3 — Query clauses

- UI to build zero or more WHERE triplets: `[field] | [operator] | [value]`.
- Operators are `equals`, `not equals`, `like`, `starts with`, `ends with`, `less than`, `greater than`, `includes`, `excludes`, and `regexp`; value rules match [query-table-data.md](./query-table-data.md) (FR-QTD-3). `in` is hidden ForceConnector compatibility syntax, not a wizard choice; `on` is rejected with guidance to use `in`.
- If user adds no clauses, default: `RECORD ID | not equals | (empty)` — equivalent to “all records” semantics via empty-id filter.
- Show destination cell at the top.
- **Run Query** writes triplets into **row 1** starting at column B (field label in cell, API name in comment where needed).

### FR-TQW-6 Header layout (row 2)

Write column headers starting at the **same column as the row 1 object cell**. Wizard layout places **Id first** (aligned with the object name on row 1); later operations locate Id by header/API name, so sheets with Id elsewhere still work if Id is present. Order when creating a table (sort by bucket, then stable describe field index):

| Sort bucket | Kind | Fill |
|-------------|------|------|
| 0 | Id | light blue |
| 1 | Name (`nameField` or API name `Name`; classified before Required) | soft gold |
| 2 | Required (standard) | peach |
| 3 | Required (custom) | darker peach |
| 4 | Standard | mint |
| 5 | Custom | lavender |
| 6 | Read-only (standard) | gray |
| 7 | Read-only (custom) | darker gray |

Required on create = `nillable == false` and createable. The same eight labels and fills are used for sheet headers and the step-2 category column.

Per header cell:

- **Display:** field label in **bold**  
- **Fill:** pastel background by bucket (table above)  

The wizard AutoFits the new header block immediately only when column sizing is set to **Fit columns to headers only**. First-page and all-data policies defer column sizing to the result writer.
- **Comment:** `API Name: {name}`, type, read-only/required hints, picklist values for picklists  

Field label → API name mapping for later query/update uses describe metadata — see [query-table-data.md](./query-table-data.md) FR-QTD-1.

### FR-TQW-7 Auto-query

- On successful criteria step, invoke the same pipeline as **Query Table Data** without requiring a second ribbon click.
- Row 3+ cleared/populated per query spec.

---

## Non-functional requirements

### NFR-TQW-1 Bulk sheet writes ([common](./common-performance-requirements.md))

- Build header label array `object[1, fieldCount]`; one write to row 2; bold the header range once.
- Apply header fills in contiguous bucket runs (O(buckets), not O(columns) when ordered).
- Comments/metadata: second pass acceptable (errors/metadata only), or store API names in row-2 comments in batch if Excel API allows — minimize COM crossings.
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

## Notes

- Single WPF wizard window — see [ui-platform.md](./ui-platform.md).
- Hidden legacy reference-list batching inherits from [query-table-data.md](./query-table-data.md).
