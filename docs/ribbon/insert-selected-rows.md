# Insert Selected Rows

**Ribbon:** `btnInsertRows` — “Insert Selected Rows”  
**Legacy:** `ForceConnector.InsertSelectedRows()` → `Operation.InsertRows()`  
**Login required:** Yes  
**COM / VBA:** `InsertSelectedRowsApi()` — same behavior

See also: [common-performance-requirements.md](./common-performance-requirements.md), [options.md](./options.md)

---

## Summary

POST new Salesforce records for selected table rows whose Id cell contains **`New`** (case-insensitive). On success, writes returned Salesforce Id into the Id cell.

---

## Functional requirements

### FR-ISR-1 Preconditions

1. Authenticated session.
2. Valid ForceConnector table with Id column.
3. Selection is within table body.
4. Selection row count ≤ **3,500** (legacy has **no** `NoQueryLimit` bypass for insert).

### FR-ISR-2 Confirmation

- Confirm before processing **unless** `NoWarning` is set (Options: *Do not show warning dialogs before commencing operations.*):
  - *“You try to insert N records. Are you sure?”* (N = selected row count, not “New” row count).
- When `NoWarning` is true → skip the dialog and proceed.
- Cancel → no API calls.

### FR-ISR-3 Row eligibility

- Process row only if Id cell matches pattern `[nN][eE][wW]*` (e.g. `New`, `new`).
- Rows without `New` are skipped silently within the chunk.

### FR-ISR-4 Record construction

Per eligible row:

- Include `attributes.type` = object API name.
- **Do not** send `Id` in body.
- For each column except Id:
  - Skip empty cells.
  - Include field only if `createable == true`.
  - Coerce types; reference fields → `NameToId` when `UseReference`.
- If no createable field values → error *“No insertable columns selected”* for that batch context.
- If no eligible rows in chunk → message *“No records to Insert… enter the string 'New' on one or more rows”*.

### FR-ISR-5 Salesforce create

- `POST /services/data/vXX/composite/sobjects`
- Batched per [common-performance-requirements.md](./common-performance-requirements.md).
- `Sforce-Auto-Assign: FALSE` header when `AutoAssignRule` is false.

### FR-ISR-6 Progress and cancel

- Chunk loop with progress during Salesforce create (*“Create the record block”* / *“Insert the record block”*).
- Cancel between chunks.
- After HTTP completes, Excel StatusBar uses **broad phase** labels only (e.g. *Writing results*) — see [common-performance-requirements.md](./common-performance-requirements.md).

### FR-ISR-7 Result feedback

| Outcome | Sheet effect |
|---------|----------------|
| Success | Id cell ← returned 18-char Id; clear row highlight and stale comments |
| Failure | Row highlight warning color; Id cell comment with SF errors |
| Summary | Optional dialog if any failures |

---

## Non-functional requirements

### NFR-ISR-1 Bulk read

**Legacy anti-pattern:** `g_table.Cells[row, col].value` per field per row.

**Target:**

- For each chunk, read a single `object[,]` covering selected rows × header columns.
- Build create payloads in Core from array + `headerFields` metadata.

### NFR-ISR-2 Bulk write (success)

- Collect `{rowIndex → newId}` in memory; write Id column cells in one or few range writes per chunk.

### NFR-ISR-3 Error path

- Per failed row: comment on Id cell (legacy). O(failures) COM acceptable.

### NFR-ISR-4 Core tests

- Filter `New` rows from id column array.
- Payload omits non-createable and empty fields.
- Mock POST composite responses.

---

## Salesforce API

| Call | Purpose |
|------|---------|
| Describe sobject | Table setup |
| `POST .../composite/sobjects` | Create records |

---

## Legacy reference

- `ForceConnector/processDatabaseInsertRows.cs`
- `Operation.insertSelectedRange`, `Operation.insertResultHandler`

## Screen tip note

Legacy screentip says “one row”; implementation supports **multiple** selected rows in batches.
