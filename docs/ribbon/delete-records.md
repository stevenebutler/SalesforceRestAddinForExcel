# Delete Records

**Ribbon:** `btnDeleteRecords` — “Delete Records”  
**Legacy:** `ForceConnector.DeleteSelectedRecords()` → `Operation.DeleteRecords()`  
**Login required:** Yes  
**COM / VBA:** `DeleteSelectedRecordsApi()` — same behavior

See also: [common-performance-requirements.md](./common-performance-requirements.md), [options.md](./options.md)

---

## Summary

DELETE Salesforce records for selected table rows using Ids from the Id column. Successful rows show **`deleted`** in the Id cell; failures get a row comment.

---

## Functional requirements

### FR-DR-1 Preconditions

1. Authenticated session.
2. Valid ForceConnector table with Id column.
3. Selection within body rows.
4. Unless `NoQueryLimit`: selected rows ≤ **3,500**.

### FR-DR-2 Confirmation

- **Always** confirm before delete (legacy dialog text incorrectly says “download” — use clear delete wording in new UI).
- Cancel → no API calls.

### FR-DR-3 Id collection

- For each selected row, read Id from Id column intersection (not necessarily column A).
- Normalize 15-char Ids to 18-char where applicable.
- Skip empty / `New` / invalid Ids **silently**. If no deletable Ids remain after filtering → error *“No Salesforce Ids found to delete.”* (no per-row warning for skipped cells).

### FR-DR-4 Salesforce delete

- `DELETE /services/data/vXX/composite/sobjects?ids={comma-separated}&allOrNone=false` — prefer per-id outcomes over failing the whole batch.
- Batch per [common-performance-requirements.md](./common-performance-requirements.md).

### FR-DR-5 Result mapping

| Outcome | Sheet effect |
|---------|----------------|
| Success | Id cell value ← literal `"deleted"` |
| Failure | Comment on row’s first cell: *“Delete Row Failed”* + SF error detail when available |

**Legacy quirk:** failure branch checks `rw.Cells[1, 1]` for some paths — new implementation must key off **Id column**, not column A.

### FR-DR-6 Progress and cancel

- Chunked processing with progress bar; cancel between batches.

### FR-DR-7 Partial failure

- With `allOrNone=false`, composite returns per-id success/failure; map each result to the corresponding sheet row and continue other rows in the batch.

---

## Non-functional requirements

### NFR-DR-1 Bulk Id read

- Read Ids for chunk via single `Intersect(g_ids, todo)` value array (legacy builds `string[]` from rows — optimize to array from one range read).

### NFR-DR-2 Bulk status write

- Success: write `"deleted"` to Id column cells in one column range write per chunk where possible.

### NFR-DR-3 Error annotations

- O(failures) comments acceptable.

### NFR-DR-4 Safety

- Destructive operation — confirmation cannot be disabled by `NoWarning` in legacy; keep hard confirm unless product explicitly changes policy.

### NFR-DR-5 Core tests

- Id array → DELETE URL / composite request.
- Map `DeleteResult[]` → per-row status DTO.

---

## Salesforce API

| Call | Purpose |
|------|---------|
| Describe sobject | Table setup |
| `DELETE .../composite/sobjects` | Delete by ids |

---

## Legacy reference

- `ForceConnector/processDatabaseDeleteRows.cs`
- `Operation.deleteSelectedRange`

## Data integrity

- Does not remove Excel row; only changes Id cell text. User may refresh or clear sheet separately.
