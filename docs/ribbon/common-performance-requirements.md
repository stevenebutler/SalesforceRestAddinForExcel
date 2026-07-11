# Common performance and platform requirements

Applies to **all** ribbon data operations and the six VBA COM methods. Individual button specs reference this document instead of repeating the same rules.

## Architecture split

| Layer | Responsibility |
|-------|----------------|
| `SalesforceRestAddin.Core` | Session gate, SOQL building, REST/composite calls, field typing, validation — **no Excel or UI** |
| `SalesforceRestAddin.Windows.Ui` | **WPF** dialogs (options, wizard, progress, About); **WebView2** for OAuth redirect only |
| `SalesforceRestAddin.ExcelDna` | Ribbon/COM entry, **bulk** range read/write, `QueueAsMacro` scheduling |
| `SalesforceRestAddin.Tests` | Core logic with mocked HTTP; no Excel or WPF in CI |

## Excel interop — mandatory patterns

Prefer bulk range I/O. Do **not** read or write worksheet data cell-by-cell in hot paths.

### Read path

1. Resolve the target **worksheet** and **index-based** range (row/column indices), not `Selection` when the operation has an explicit table context.
2. Read contiguous data with **one** `Range.Value` / `Value2` into `object[,]` per logical block (header row, body block, selection chunk).
3. Parse the array in managed code; do not read cell-by-cell for data values.

### Write path

1. Build `object[,]` (or column-scoped arrays) in memory from Salesforce responses.
2. Assign with **one** `Range.Value` per contiguous rectangle via `SheetProjectionWriter` (not per-row cell writes).
3. Apply **number formats** at **column** or **range** scope before or after the bulk value assign.
4. After write: **AutoFit columns** for the written block (include the header row above the body when present so widths fit labels). **AutoFit rows**, then cap any row taller than **3×** `StandardHeight` (long-text guard) — O(rows) clamp only; no per-cell height loops.
5. Do not set comments, interior color, or validation **per cell** in the hot path. Batch error annotation where possible (see Error feedback).

### Formatting and derived display

| Field type | Write strategy |
|------------|----------------|
| `reference` | Optional name resolution via `UseReference` option; resolve in Core, write final display values in bulk |
| `date` / `datetime` | Preserve Excel date types; apply column `NumberFormat` (`yyyy-MM-dd` / `yyyy-MM-dd HH:mm:ss`) |
| `address` / `location` | Flatten compound objects to strings in Core before array write |
| `picklist` / text | Respect text formats (`@`) to preserve leading zeros |

Row height: AutoFit the written rows, then cap auto-expansion for long text at **3×** standard height without per-cell loops in the success path.

### COM boundary crossings

Treat each property get/set on `Excel.Range` / `Excel.Worksheet` as expensive.

| Operation | Target |
|-----------|--------|
| Query results | Single `object[,]` → one `Range.Value` |
| Update (default, skip hidden) | `todo.Value` bulk read + filter hidden indices in memory |
| Update (include hidden) | Same bulk read; no hidden filter |
| Insert | Read selected row slice as `object[,]` once per chunk |
| Wizard draw | Batch header labels + defer comments or use a single metadata pass |
| Describe output | Bulk values only (comments out of scope) |
| Error feedback | Acceptable for **errors only** (small N); avoid on success path |

**Goal:** O(1) range reads and O(1) range writes per chunk per data plane, plus O(columns) format passes — not O(rows × columns) COM calls.

## Salesforce API — batching and limits

| API | Endpoint (REST) | Salesforce limit | **Target chunk** |
|-----|-----------------|------------------|------------------|
| Query | `GET /services/data/vXX/query` | 2,000 rows/page | Paginate via `nextRecordsUrl`; use Core pagination |
| Retrieve | `POST .../composite/sobjects/{sObject}` | 2,000 ids | **Up to 200** (configurable, ≤ platform max) |
| Create | `POST .../composite/sobjects` | 200 records | **Up to 200** |
| Update | `PATCH .../composite/sobjects` | 200 records | **Up to 200** |
| Delete | `DELETE .../composite/sobjects?ids=` | 200 ids | **Up to 200** |

- Batch size should be **configurable** (`ConnectorOptions` or dedicated setting), defaulting to 200 where REST allows.
- Send `Sforce-Auto-Assign: FALSE` on create/update when `AutoAssignRule` is false (product default). See [options.md](./options.md).

## HTTP compression (NFR-HTTP-1)

Salesforce REST supports gzip/deflate ([Compression Headers](https://developer.salesforce.com/docs/atlas.en-us.api_rest.meta/api_rest/intro_rest_compression.htm)).

| Direction | Mechanism |
|-----------|-----------|
| **Responses** | Shared `HttpClient` uses `AutomaticDecompression = GZip \| Deflate` (`SalesforceHttpClientFactory`) so `Accept-Encoding` is sent and bodies are unwrapped. |
| **Requests** | Composite create/update JSON bodies ≥ 512 UTF-8 bytes are gzip-compressed with `Content-Encoding: gzip` (`GzipJsonContent`). Smaller bodies stay plain JSON. OAuth token posts are not compressed. |

## Selection and table limits

Defaults (overridable via `NoQueryLimit` where noted):

| Limit | Value | Applies to |
|-------|-------|------------|
| `maxRows` | 3,500 | Update, query-rows, insert, delete selection row count |
| `maxCols` | 20 | Update selection width |
| `excelLimit` | 1,048,570 | Query-table result row cap vs Excel grid |
| `excelColLimit` | 16,384 | Wide result guard |

**Multi-area selection:** Update allows Ctrl+multi-range **within one ForceConnector table** (per-row field sets; values from the bulk-captured table body). Abort if any selected cell is outside that table. Query / Insert / Delete still reject multi-area (`Areas.Count > 1`).

## Metadata cache (NFR-META-1)

Any time the connector downloads Salesforce **metadata**, cache it and consult the cache before re-downloading:

| Metadata | REST source | Cache key |
|----------|-------------|-----------|
| Global object list | `GET …/sobjects/` | **Salesforce instance URL** + API version |
| Per-object describe | `GET …/sobjects/{name}/describe` | **Salesforce instance URL** + API version + object API name |

Rules:

- **Per-instance isolation (hard requirement)** — never mix metadata between Salesforce instances or sandboxes. Cache paths are partitioned by the session `instance_url` (e.g. `https://kjr--kjr2026.sandbox.my.salesforce.com` vs production). Production, sandboxes, and scratch orgs each keep separate trees. A missing/blank instance URL must **not** read or write the cache (no shared fallback bucket).
- **Two-tier cache** — in-process typed L1 (`ParsedMetadataCache`: object lists, `SObjectDescribe`, `FieldCatalog`) in front of durable JSON L2 (`FileMetadataCache`). Lookup order: memory → disk → HTTP. Writes update L1 (parsed) and L2 (JSON) together.
- **Write-through on miss** — after a successful download, store the raw JSON under `%LOCALAPPDATA%\ForceConnector\metadata-cache\{instance}\{apiVersion}\` (see `SalesforceRestAddinDataPaths`).
- **Inspect before HTTP** — `ListObjectsAsync` / `DescribeAsync` check the cache first.
- **No TTL in v1** — entries remain until the user clears the cache for the **current login instance** (Options → Clear metadata cache for this org). Clearing does **not** wipe other orgs/sandboxes.
- Does **not** cache SOQL query results or composite CRUD payloads.

See [options.md](./options.md) FR-OPT-6 for the clear-cache UI.

## Session and threading

| Requirement | Detail |
|-------------|--------|
| Login gate | Data buttons require authenticated session (`SessionGate`); About and Options do not |
| Excel thread | All worksheet mutations on Excel main thread (`ExcelAsyncUtil.QueueAsMacro`) |
| Async I/O | HTTP to Salesforce off UI thread; marshal results back to macro queue for sheet writes |
| Cancellation | Cooperative cancel between batches; no `Thread.Sleep` polling in tests |

## Excel STA async (NFR-STA-1)

Excel-DNA rule: **never** block the Excel main thread with bare `GetAwaiter().GetResult()` / `.Result` / `.Wait()`, and **never** touch Excel COM off the main STA. See [Performing Asynchronous Work](https://excel-dna.github.io/docs/guides-advanced/performing-asynchronous-work/).

| Rule | Detail |
|------|--------|
| Entry | Ribbon/COM always schedules via `ExcelAsyncUtil.QueueAsMacro` (no `async` delegates inside the macro) |
| Sole wait/pump | [`ExcelStaAsyncHost`](../../src/SalesforceRestAddin.ExcelDna/ExcelStaAsyncHost.cs) owns modal WPF `ShowDialog` (nested message loop), progress/Cancel, and StatusBar during background work |
| COM marshal | From workers: `ExcelComThread` / `QueueAsMacro` — thin primitive, **not** a second wait/pump |
| Snapshot before async | Capture Excel state (`object[,]`, ranges, selection) on STA before starting HTTP; sheet writes after await only on STA |
| Forbidden | Bare `GetResult` / `.Result` / `.Wait` on the Excel STA; `Dispatcher.PushFrame`; a second wait/pump class parallel to `ExcelStaAsyncHost` |

**Allowed blocking wait:** only `ExcelStaAsyncHost.Run` (modal WPF `ShowDialog`). Post-dialog `GetResult()` is fine because the task is already complete.

**Zombie Excel:** COM RCWs created off-STA (e.g. StatusBar after `ConfigureAwait(false)`) can leave `EXCEL.EXE` running after close, which then steals the XLL and breaks ribbon/COM. Kill leftover Excel processes before blaming packing. See [bugs.md](../bugs.md) #6/#7 and [AGENTS.md](../../AGENTS.md).

## Progress and status

- Driven by `ExcelStaAsyncHost` during async ops (WPF progress window + Excel StatusBar ≤128 chars).
- **During Salesforce API work:** StatusBar may update with progress (download counts, batch status, etc.).
- **After the progress UI closes (Excel apply):** StatusBar uses **broad phase** labels only — e.g. *Clearing cells*, *Updating cells*, *Formatting cells*, *Writing results*. Do not thrash StatusBar per row/cell after HTTP completes.
- Cancel when the op supports it (query download); status-only waits (login, ListObjects) hide Cancel.
- Disable re-entry while a long operation runs.

## Error feedback (shared)

| Signal | When |
|--------|------|
| Modal / non-blocking summary | Unhandled exception, batch completed with failures |
| Row interior highlight | Failed update/insert/delete row — **table columns only** (not entire worksheet row) |
| Cell comment on Id column of failed row | Salesforce `errors[]` statusCode + message |
| `#N/F` | Query returned zero rows (first body Id cell) |
| `#Err` | Query failed mid-write |
| `"deleted"` | Successful delete (Id cell text) |
| Returned Id | Successful insert (replaces `"New"` in Id cell) |

Prefer recording errors in a Core result DTO; ExcelDna applies visuals in one pass over failed row indices.

## Options cross-reference

`ConnectorOptions` ([`src/SalesforceRestAddin.Core/Session/ConnectorOptions.cs`](../../src/SalesforceRestAddin.Core/Session/ConnectorOptions.cs)) persists to JSON (`JsonConnectorOptionsStore`). Flags affecting data plane:

| Option | Effect |
|--------|--------|
| `UseReference` | Name ↔ Id for reference fields |
| `NoWarning` | Skip insert and include-hidden update confirmation dialogs |
| `NoQueryLimit` | Skip maxRows/maxCols checks |
| `AutoAssignRule` | When false, suppress auto-assignment header |
| `IncludeHiddenCells` | When false (default), update omits hidden rows/columns; when true, include them |

## Testability

- Core: SOQL builder, record payload builder, batch splitter, type coercion — unit tested with mocks.
- ExcelDna: thin integration; prefer testing Core + a small “range adapter” if needed.
- Tests must be **deterministic** (no `Thread.Sleep`); see [AGENTS.md](../../AGENTS.md).
- Architecture and test catalog: [design.md](./design.md), [test-cases.md](./test-cases.md).

## Intentionally out of scope / not used

| Topic | Product rule |
|-------|--------------|
| SOAP describe / logout | REST describe; OAuth session clear |
| METAAPI / translation columns in Describe | Dropped — see [AGENTS.md](../../AGENTS.md) |
| Composite batch size under REST limits | Configurable up to 200 (default 200) |
| Per-reference-id SOQL in join mode | Batch `IN` queries — see [query-table-data.md](./query-table-data.md) |
