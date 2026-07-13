# Query Table Data

**Ribbon:** `btnQueryTable` — “Query Table Data”  
**Login required:** Yes  
**COM / VBA:** `QueryTableDataApi()` — same behavior

See also: [common-performance-requirements.md](./common-performance-requirements.md), [table-query-wizard.md](./table-query-wizard.md)

---

## Summary

Parse WHERE criteria from **row 1**, run SOQL against the table’s object, paginate results, and bulk-write rows starting at row 3. Clears existing body data before query.

---

## Functional requirements

### FR-QTD-1 Preconditions

1. Authenticated session.
2. If Excel has no usable worksheet context, abort early with a prompt to open a workbook containing a Salesforce Connector table.
3. Valid Salesforce Connector table under the selection/active cell:
   - **Multi-table sheets:** tables are separated by **at least one blank column** in the header row. From the selected cell, scan the header row **left** to the start of the contiguous non-blank header block (or column A). The object name is in the cell **above that first column**. If that cell has no valid entity name, **abort** with an error citing the cell address.
   - Object name in the row 1 anchor cell (API name in comment, or value when no comment).
   - Headers in row 2 of the same column span.
   - **Record Id** (Salesforce primary key from describe metadata) is **required** somewhere in the row-2 header block. The wizard places it first (aligned with the row 1 object cell); other column positions are allowed if Id is present and resolvable.
   - Field headers resolve to API names via **describe metadata** (`FieldCatalog`): prefer row-2 comment `API Name: …`, then field label match, then API-name fallback.
4. Query results, clears, Id write-back, and error notes/colour apply only within that table’s column range (never from sheet column A unless the table starts there).

### FR-QTD-2 Clear body

- Before write: clear **all** cell content (values, formats, comments — Excel `Range.Clear`, not values-only) of body rows within the **table’s column span** (entity only; do not clear neighbouring tables on the same sheet).
- Clear from the first body row through the **last existing body row** captured before the query, or through the last new result row if that is larger — so a smaller result set does not leave stale rows.
- Do not delete Excel rows or remove row 1–2 structure.

### FR-QTD-3 WHERE clause builder (row 1)

Starting column B, read repeating triplets: `[field label or API] | [operator] | [value]`.

**Field resolution:**

- Match row-1 field label to describe `fieldLabelMap`, else cell comment, else treat as API name.

**Operator aliases (case-insensitive):**

| UI operator | SOQL |
|-------------|------|
| equals | `=` |
| contains | `like` (wrap `%value%`) |
| not equals | `!=` |
| less than / greater than | `<` / `>` |
| begins with / starts with | `LIKE` with `value%` |
| ends with | `LIKE` with `%value` |
| regexp | `like` (user supplies wildcards) |
| includes / excludes | `includes` / `excludes` |

**Value rules:**

- Comma-separated values → OR group, wrapped in parentheses.
- Empty value → `= ''` or `= null` for date types.
- `like` on picklist → validation error with cell address.
- Multipicklist: `includes` / `excludes` semantics for negated operators.
- Reference values: `NameToId` when `UseReference`.
- Escape SOQL special characters.
- **Date / datetime criteria:**
  - **Accept:** Excel native date/datetime cells (OADate / `DateTime` via `Value2`); ISO strings (`yyyy-MM-dd` or `yyyy-MM-dd[T| ]HH:mm:ss…`); relative SOQL literals (`TODAY`, `LAST_N_DAYS`, …).
  - **Reject:** locale-formatted text (e.g. `26/07/2021`, `7/9/2026`) — abort with a validation error citing the criteria cell. Do not call `Evaluate("DATEVALUE")` or guess culture.

**Validation errors** must cite the offending cell address.

### FR-QTD-4 Hidden legacy reference-list compatibility (`in` on `Id` / reference fields)

- `in` is not a visible Table Query Wizard operator. It remains accepted for Salesforce Connector-style sheets.
- Its value cell must name an Excel range or named range containing Salesforce record Ids. Empty cells and non-Id values in that range are ignored; no valid Ids is a validation error citing the criteria cell.
- Comma-separated Id text is not a supported reference-list input.
- `in` executes as a normal matching query over the supplied Id set.
- `on` is rejected with: `ON is not supported - use IN with a range reference to select multiple items.`

**Required approach:** batch with `WHERE ref IN (...)` within Salesforce limits (not one HTTP query per Id).

### FR-QTD-5 SELECT list

- Fields from row 2 header mapping (same order as columns).

### FR-QTD-6 Pre-query confirmation (optional)

When `NoConfirmQueryDownload` is **false** in Options (default — show confirmation):

1. `SELECT COUNT(Id) ...` with same WHERE.
2. Reject if count > `excelLimit` (1,048,570).
3. Ask user to confirm download count (any result size).

When `NoConfirmQueryDownload` is true → skip COUNT + confirm and proceed.

Independent of `NoWarning` (which gates insert / include-hidden update confirms).

### FR-QTD-7 Query execution

- `GET .../query?q=SELECT {fields} FROM {object} {where}`
- Paginate with `nextRecordsUrl` until `done`.
- Report progress: records downloaded / `totalSize`.

### FR-QTD-8 Write results

- Bulk write via the shared projection pattern ([common](./common-performance-requirements.md)).
- **Zero rows:** write `#N/F` in first body cell of Id column.
- **Error mid-run:** `#Err` in current row Id cell; surface exception message.

### FR-QTD-9 Cancel

- Stop fetching next pages; leave partial results already written (do not roll back prior pages).

---

## Non-functional requirements

### NFR-QTD-1 Bulk write

- One `object[,]` per page (or accumulate pages up to memory policy) → minimal `Range.Value` calls.
- Column formats applied at column scope after write.
- Apply the selected column and row sizing policies after write — see [common-performance-requirements.md](./common-performance-requirements.md). First-page and all-data column policies do not AutoFit during wizard header creation.

### NFR-QTD-2 SOQL generation in Core

- Pure function: table model + describe → SOQL string + metadata.
- Unit test golden cases for operators, multipicklist, empty values, OR groups.

### NFR-QTD-3 Reference-list batching

| Approach | API cost |
|----------|----------|
| One query per reference Id | M HTTP round-trips — **do not use** |
| Batched `IN` clauses | ⌈M / limit⌉ round-trips — **required** |

### NFR-QTD-4 Memory (**deprecated**)

~~For very large `totalSize`, stream pages to sheet without holding full result set in memory unless necessary.~~

**Deprecated:** streaming each page to Excel mid-query fights the bulk-COM goal (many clear/write cycles, harder cancel/clear-body semantics) and complicates the apply path. Prefer accumulate pages in Core, then one clear + one bulk write (NFR-QTD-1). Revisit only if memory pressure becomes a real product issue.

### NFR-QTD-5 Status bar

- During query HTTP: StatusBar / progress may show download progress (e.g. records downloaded / `totalSize`, or *Select Data From {Object}*) — ≤128 chars.
- After HTTP completes, Excel apply uses **broad phase** labels only (e.g. *Clearing cells*, *Updating cells*, *Formatting cells*) — see [common-performance-requirements.md](./common-performance-requirements.md).

---

## Salesforce API

| Call | Purpose |
|------|---------|
| Describe sobject | Headers / types |
| `GET .../query` | Data + pagination |
| `GET .../query` (COUNT) | Optional confirmation |

---

## Wizard integration

[Table Query Wizard](./table-query-wizard.md) ends by calling this pipeline (FR-TQW-7).
