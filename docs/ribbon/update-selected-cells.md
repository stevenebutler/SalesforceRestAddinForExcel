# Update Selected Cells

**Ribbon:** `btnUpdateCells` — “Update Selected Cells”  
**Login required:** Yes  
**COM / VBA:** `UpdateSelectedCellsApi()` — same behavior

See also: [common-performance-requirements.md](./common-performance-requirements.md), [options.md](./options.md)

---

## Summary

PATCH selected data cells in a ForceConnector table to Salesforce. Reads Ids from the table’s Id column, sends updateable fields only, reports per-row failures on the sheet.

---

## Functional requirements

### FR-USC-1 Preconditions

1. Authenticated session.
2. Active cell lies within a valid ForceConnector table.
3. Table has mappable **Id** column in row 2.
4. User selection intersects table **body** (row 3+).
5. Selection may be **multi-area** (Ctrl+select) **within a single ForceConnector table**. Each body row’s PATCH includes only columns selected on that row (union if overlapping areas hit the same row). If any selected cell falls outside the captured table (including another entity’s columns), **abort** before HTTP with a clear error. Multi-entity updates in one operation are not supported.
6. Unless `NoQueryLimit`: selection ≤ **3,500 rows** and ≤ **20 distinct columns** (not the bounding-box width of sparse multi-area selections).

### FR-USC-2 Path selection (`IncludeHiddenCells` option)

| Option | Path |
|--------|------|
| `IncludeHiddenCells == false` (**default**) | **Hidden-aware update** — omit AutoFilter/manually hidden rows and columns from the PATCH payload |
| `IncludeHiddenCells == true` | **Include hidden** — every selected cell is sent, including hidden ones |

Both paths must produce identical Salesforce payloads for the same visible data when no rows/columns are hidden.

### FR-USC-3 Confirmation

- Confirm every update unless `NoWarning` is true.
- **Default (skip hidden):** confirm the selected record count and object.
- **Include-hidden path:** use confirmation wording that warns hidden rows/columns in the selection will be updated.

### FR-USC-4 Record construction

For each selected row in chunk:

- **Id:** from Id column (normalize 15 → 18 char).
- **Fields:** only columns where describe says `updateable == true`, and only columns **selected on that row** (per-row field sets for multi-area).
- **Skip** Id field in body; include `attributes.type`.
- Coerce cell values via field type: dates, references (`NameToId` when `UseReference`), picklists, etc.
- **Empty / cleared cells** → JSON `null` for all field types. Included in the PATCH body so Salesforce clears the field. May revisit string `""` vs `null` after sandbox testing.
- Cell values come from the **bulk-captured** table `Body` array (one Excel read); selection supplies indices only.
- If no updateable columns in selection → cancel with *“No updatable columns selected”*.

### FR-USC-5 Salesforce update

- `PATCH /services/data/vXX/composite/sobjects` with batched records.
- Chunk size per [common-performance-requirements.md](./common-performance-requirements.md) (target 200).
- Optional header: `Sforce-Auto-Assign: FALSE` when `AutoAssignRule` is false.

### FR-USC-6 Progress and cancel

- Show progress across chunks (rows processed / total) during Salesforce PATCH.
- Cooperative cancel between chunks; do not start a new HTTP batch after cancel.
- Do **not** highlight the working chunk on the sheet during processing.

### FR-USC-7 Result feedback

| Outcome | Sheet effect |
|---------|----------------|
| Success | Clear error styling on row |
| Row failure | Row interior warning color; Id cell comment with SF `errors[]` |
| Partial / all failed | Continue other rows; `ErrorDialogWindow` with per-row SF errors (same text as comments) |

### FR-USC-8 Status bar

- During API work: StatusBar may reflect batch/progress messages (≤128 chars).
- After HTTP completes, Excel apply uses **broad phase** labels only (e.g. *Updating cells*, *Formatting cells*) — not per-row thrashing. See [common-performance-requirements.md](./common-performance-requirements.md).

---

## Non-functional requirements

### NFR-USC-1 Bulk read (standard path)

- **Must** read selection/chunk values as `object[,]` once per chunk.
- Map array indices to field metadata in managed code.

### NFR-USC-2 Bulk read (hidden-row filtering)

Hidden state is captured once with the table snapshot (`HiddenRowIndices` / `HiddenColumnIndices`). Default update filters those indices in memory after the bulk `Body` read — no per-cell Excel COM for visibility.

### NFR-USC-3 Error write path

- Failed rows only: apply interior + comment. Accept O(failures) COM calls, not O(all cells).

### NFR-USC-4 Core testability

- Unit tests: given `object[,]` + field map → JSON records; mock PATCH responses → per-row result DTO.

### NFR-USC-5 Idempotency

- Re-running update on unchanged data should be safe (Salesforce validates).

---

## Salesforce API

| Call | Purpose |
|------|---------|
| `GET .../sobjects/{type}/describe` | Field metadata (on table setup) |
| `PATCH .../composite/sobjects` | Update records |
