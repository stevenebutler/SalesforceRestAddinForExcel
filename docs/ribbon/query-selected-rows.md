# Query Selected Rows

**Ribbon:** `btnQueryRows` — “Query Selected Rows”  
**Login required:** Yes  
**COM / VBA:** `QuerySelectedRowsApi()` — same behavior

See also: [common-performance-requirements.md](./common-performance-requirements.md), [options.md](./options.md)

---

## Summary

Retrieve current field values from Salesforce for **selected table rows** by Id, writing results into the row-aligned body columns (Id column left unchanged).

---

## Functional requirements

### FR-QSR-1 Preconditions

1. Authenticated session.
2. Valid ForceConnector table.
3. Selection intersects body rows with valid Ids in Id column (15/18-char; not `New`).
4. Unless `NoQueryLimit`: selected rows ≤ **3,500**.

### FR-QSR-2 Field set

- Query **all** non-Id columns mapped in row 2 (`headerFields`).
- Write path skips overwriting Id column in sheet. Ids from the selection are used only for retrieve requests and row alignment.

### FR-QSR-3 Confirmation

- No count dialog before retrieve.

### FR-QSR-4 Retrieve

- Collect Ids from selection ∩ Id column for chunk.
- `POST .../composite/sobjects/{sObject}` with `{ ids, fields }`.
- Batch per [common-performance-requirements.md](./common-performance-requirements.md).

### FR-QSR-5 Write results

- Map records to row order of selection.
- Missing record / bad Id: rely on retrieve response alignment — preserve “no silent wrong row” behavior.
- Apply reference display names when `UseReference` (`IdToName`).

### FR-QSR-6 Progress and cancel

- Progress during Salesforce retrieve (e.g. download / batch status).
- Cancel between chunks.
- Do **not** highlight the working chunk on the sheet during processing.
- After HTTP completes, Excel StatusBar uses **broad phase** labels only (e.g. *Updating cells*, *Formatting cells*) — see [common-performance-requirements.md](./common-performance-requirements.md).

### FR-QSR-7 Refresh Table Data (COM API)

`RefreshTableDataApi()` uses the same worker with `RefreshAll = true`:

- Auto-expand selection to **contiguous** Id rows in the Id column (from first data row through last non-empty Id).
- No separate requirement doc; same bulk write rules apply.

---

## Non-functional requirements

### NFR-QSR-1 Bulk write (reference implementation)

1. Build `object[recordCount, fieldCount]` in memory.
2. Flatten compound types (address, location) to strings in Core.
3. **Single** `Range.Value` assign to body rectangle.

**This is the pattern all query/write features should follow.**

### NFR-QSR-2 Bulk Id read

- Read Id cells for chunk via one range intersection → `Value` array.

### NFR-QSR-3 Formatting

- Column-level number formats for types (date, datetime, text).
- Avoid per-cell format loops where a column shares one format.
- Column AutoFit and capped row AutoFit via shared `SheetProjectionWriter` (same as query table).

### NFR-QSR-4 Core tests

- Id list extraction from column array.
- Field list → REST retrieve request body.
- Response rows → `object[,]` including compound field flattening.

---

## Salesforce API

| Call | Purpose |
|------|---------|
| Describe sobject | Table metadata |
| `POST .../composite/sobjects/{type}` | Retrieve by ids |
