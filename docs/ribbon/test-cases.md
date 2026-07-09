# Test cases

Test plan for the [data plane design](./design.md) (module layout and shared types), traced to [ribbon requirements](./README.md). Implement in `SalesforceRestAddin.Tests` (TUnit, `net10.0`) unless noted.

**Rules:** deterministic orchestration only — see [AGENTS.md](../../AGENTS.md#unit-tests). Prefer testing Core without Excel; use `SequentialMockHttpHandler` or inject a test double for `SalesforceDataClient`.

---

## Legend

| Tag | Meaning |
|-----|---------|
| **Unit** | Pure Core, no HTTP |
| **HTTP** | Core + `SequentialMockHttpHandler` |
| **Excel** | Windows manual / future adapter test — not in Linux CI |

Status column left blank for implementation tracking.

---

## Shared: table parsing (`ForceTableParser`)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-TBL-01 | FR-QTD-1 | Valid snapshot: A1 comment `Account`, row-2 headers with Id, body 2×N → binding succeeds, `IdColumnIndex` correct | Unit |
| T-TBL-02 | FR-QTD-1 | A1 value `Account` when no comment → object name resolved | Unit |
| T-TBL-03 | FR-QTD-1 | A1 empty / contains space, no comment → validation error cites A1 | Unit |
| T-TBL-04 | FR-QTD-1 | Row 2 has unknown label without API comment → binding stops at gap; error identifies column | Unit |
| T-TBL-05 | FR-QTD-1 | Row-2 comment `API Name: Custom__c` maps to field | Unit |
| T-TBL-06 | FR-QTD-1 | No Id column in headers → validation error | Unit |
| T-TBL-07 | common | `HiddenRowIndices` excluded from update payload by default (`IncludeHiddenCells` false) | Unit |
| T-TBL-08 | common | `HiddenColumnIndices` excluded from update payload | Unit |
| T-TBL-09 | FR-QTD-1 | Id present but not first field column → binding succeeds with that `IdColumnIndex` | Unit |
| T-TBL-10 | FR-QTD-1 | Header left-scan finds table start after blank gap | Unit |
| T-TBL-11 | FR-QTD-1 | Active cell in left block resolves left table bounds | Unit |
| T-TBL-12 | FR-QTD-1 | No header at selection → resolve error | Unit |
| T-TBL-13 | FR-QTD-1 | Bind preserves `StartRow` / `StartColumn` | Unit |
| T-TBL-14 | FR-QTD-1 | Missing object cites anchor cell address | Unit |

---

## Shared: metadata cache (`IMetadataCache` / `SalesforceDataClient`)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-META-01 | NFR-META-1 | Describe miss → HTTP once; second call hits cache (no second HTTP) | HTTP |
| T-META-02 | NFR-META-1 | `ClearInstance` for current org forces re-download on next describe | HTTP |
| T-META-03 | NFR-META-1 | Separate cache entries per API version / instance key | Unit |
| T-META-09 | NFR-META-1 | Blank/invalid instance URL neither reads nor writes cache | Unit |
| T-META-10 | NFR-META-1 | Sandbox vs production instance URLs do not share cache entries | Unit |
| T-META-11 | NFR-META-1 | Typed L1 serves second describe without HTTP or disk re-parse | HTTP |
| T-META-12 | NFR-META-1 | `ClearInstance(sandbox)` leaves production L1+L2 intact | Unit |

---

## Shared: selection limits (`SelectionLimits`)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-LIM-01 | FR-USC-1 | 3501 rows, default options → rejected | Unit |
| T-LIM-02 | FR-USC-1 | 21 columns on update → rejected | Unit |
| T-LIM-03 | options | `NoQueryLimit=true` → 3501 rows allowed | Unit |
| T-LIM-04 | common | Multi-area selection flag → rejected by default (Query/Insert/Delete) | Unit |
| T-LIM-05 | FR-USC-1 | Multi-area allowed for Update (`allowMultiArea: true`) when in-bounds | Unit |
| T-LIM-06 | FR-USC-1 | Distinct column count for sparse multi-area (B+D → 2, not span) | Unit |

---

## Shared: batching (`RecordBatchSplitter`)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-BAT-01 | common | 450 ids, batch 200 → chunks [200,200,50] | Unit |
| T-BAT-02 | common | 0 ids → empty chunk list | Unit |
| T-BAT-03 | common | Custom `CompositeBatchSize=50` honored | Unit |

---

## Shared: field values (`FieldValueConverter`)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-VAL-01 | common | Empty cell → JSON `null` on update/create (create omits nulls) | Unit |
| T-VAL-02 | common | Date cell → ISO date in Salesforce payload | Unit |
| T-VAL-03 | common | Date/datetime display → preserves Excel-friendly value (not trimmed to broken text) | Unit |
| T-VAL-04 | options | Reference with `UseReference`: name → Id via resolver | Unit |
| T-VAL-05 | options | Reference display: Id → name via resolver | Unit |
| T-VAL-06 | common | Picklist value passed through | Unit |
| T-VAL-07 | common | Multipicklist semicolon join | Unit |
| T-VAL-08 | common | Address compound → single display string | Unit |
| T-VAL-09 | common | Leading-zero text field → text format hint in projection | Unit |
| T-VAL-10 | FR-USC-4 | Non-updateable field omitted from PATCH body | Unit |
| T-VAL-11 | FR-ISR-4 | Non-createable field omitted from POST body | Unit |

---

## SOQL builder (`SoqlCriteriaParser`, `SoqlQueryBuilder`)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-SOQL-01 | FR-QTD-3 | `equals` → `Field = 'x'` | Unit |
| T-SOQL-02 | FR-QTD-3 | `contains` → `LIKE '%x%'` | Unit |
| T-SOQL-03 | FR-QTD-3 | `not equals` → `!=` | Unit |
| T-SOQL-04 | FR-QTD-3 | `begins with` → `LIKE 'x%'` | Unit |
| T-SOQL-05 | FR-QTD-3 | `ends with` → `LIKE '%x'` | Unit |
| T-SOQL-06 | FR-QTD-3 | Comma-separated values → `(a OR b)` | Unit |
| T-SOQL-07 | FR-QTD-3 | Empty value on text → `= ''` | Unit |
| T-SOQL-08 | FR-QTD-3 | Empty value on date → `= null` | Unit |
| T-SOQL-09 | FR-QTD-3 | `like` on picklist → validation error with cell ref | Unit |
| T-SOQL-10 | FR-QTD-3 | Multipicklist `includes` / `excludes` for negated ops | Unit |
| T-SOQL-11 | FR-QTD-3 | SOQL special chars escaped in literal | Unit |
| T-SOQL-12 | options | Reference value with `UseReference` → Id in WHERE | Unit |
| T-SOQL-13 | FR-QTD-5 | SELECT field order matches header binding order | Unit |
| T-SOQL-14 | FR-QTD-4 | Reference `in` with id list → `IN (...)` | Unit |
| T-SOQL-15 | FR-QTD-4 | Reference `on` with 250 ids → batched IN (2 queries), not 250 | Unit |
| T-SOQL-16 | FR-QTD-6 | COUNT query built with same WHERE as data query | Unit |

---

## Query table (`QueryTable`)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-QTD-01 | FR-QTD-7 | Single-page query → projector rows match mock records | HTTP |
| T-QTD-02 | FR-QTD-7 | `nextRecordsUrl` pagination → two HTTP calls, merged projection | HTTP |
| T-QTD-03 | FR-QTD-8 | Zero rows → projection marks `#N/F` at first body Id cell | Unit |
| T-QTD-04 | FR-QTD-8 | HTTP 400 mid-run → `#Err` marker + exception message in result | HTTP |
| T-QTD-05 | FR-QTD-6 | Count > excelLimit → rejected before data query | Unit |
| T-QTD-06 | FR-QTD-9 | Cancel after page 1 → partial `RecordsWritten`, `WasCancelled` | HTTP |
| T-QTD-07 | NFR-QTD-1 | Projector output is single rectangular `object[,]` | Unit |
| T-QTD-08 | FR-QTD-2 | Plan includes clear-body through full previous body when result is smaller | Unit |
| T-QTD-08b | FR-QTD-2 | Zero rows clears full existing body then writes `#N/F` | Unit |
| T-QTD-08c | FR-QTD-2 | Larger result clear EndRow covers new projection | Unit |

---

## Query selected rows (`QueryRows`)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-QSR-01 | FR-QSR-4 | 3 ids → one composite retrieve request with all fields | HTTP |
| T-QSR-02 | FR-QSR-5 | Id column not present in value projection columns | Unit |
| T-QSR-03 | FR-QSR-4 | 250 rows → 2 HTTP batches (200+50) | HTTP |
| T-QSR-04 | FR-QSR-1 | Row with invalid Id → row outcome or empty slot per design | Unit |
| T-QSR-05 | Refresh | Refresh mode: contiguous Id rows selected from snapshot metadata | Unit |
| T-QSR-06 | T-LIM-01 | 3501 selected rows rejected unless `NoQueryLimit` | Unit |

---

## Update cells (`UpdateCells`)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-USC-01 | FR-USC-4 | PATCH body contains Id + updateable fields only | Unit |
| T-USC-02 | FR-USC-4 | No updateable columns → error, no HTTP | Unit |
| T-USC-03 | FR-USC-5 | `AutoAssignRule=false` → `Sforce-Auto-Assign: FALSE` header | HTTP |
| T-USC-04 | FR-USC-5 | `AutoAssignRule=true` → header absent | HTTP |
| T-USC-05 | FR-USC-7 | Partial failure → `RowOutcome` + `ErrorSummary` for dialog | HTTP |
| T-USC-06 | FR-USC-2 | Default: hidden row in selection omitted from PATCH | Unit |
| T-USC-07 | FR-USC-2 | Default: hidden column omitted from PATCH | Unit |
| T-USC-08 | NFR-USC-1 | Standard path builds records from 2D array without per-cell API | Unit |
| T-USC-09 | FR-USC-6 | Cancel between batches → second batch not sent | HTTP |
| T-USC-10 | FR-USC-1/4 | Multi-area same table, different columns per row → per-row fields only | Unit |
| T-USC-11 | FR-USC-1/4 | Overlapping areas on same row → union of columns | Unit |
| T-USC-12 | FR-USC-1 | Selected column outside table width → error, no records | Unit |

---

## Insert rows (`InsertRows`)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-ISR-01 | FR-ISR-3 | Rows `New`, `new`, `NEW` included; `Acme` skipped | Unit |
| T-ISR-02 | FR-ISR-4 | Empty createable set for row → batch error message | Unit |
| T-ISR-03 | FR-ISR-5 | Success → `IdWriteback` on row outcome | HTTP |
| T-ISR-03b | FR-ISR-7 | Partial/create failure → `RowOutcomes` + `ErrorSummary` for dialog | HTTP |
| T-ISR-04 | FR-ISR-5 | No eligible rows → error, no HTTP | Unit |
| T-ISR-05 | FR-ISR-1 | 3501 rows selected → rejected (no NoQueryLimit) | Unit |
| T-ISR-06 | NFR-ISR-1 | Records built from body array slice in one pass | Unit |

---

## Delete records (`DeleteRecords`)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-DR-01 | FR-DR-4 | DELETE URL contains all ids in batch | HTTP |
| T-DR-02 | FR-DR-5 | Success → `IdDisplayOverride == "deleted"` | Unit |
| T-DR-03 | FR-DR-5 | Failure → row outcome + `ErrorSummary` for dialog | HTTP |
| T-DR-04 | FR-DR-3 | Id from Id column index, not hardcoded column A | Unit |
| T-DR-05 | FR-DR-6 | Cancel between batches → remaining batches skipped | HTTP |

---

## Describe object (`DescribeObject`)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-DSO-01 | FR-DSO-3 | REST describe JSON → field grid rows | HTTP |
| T-DSO-02 | FR-DSO-4 | Field ordering: named standard → standard → custom | Unit |
| T-DSO-03 | FR-DSO-4 | Picklist values column populated | Unit |
| T-DSO-04 | FR-DSO-7 | One object fails describe → others still processed | HTTP |

---

## Metadata client (`SalesforceDataClient`)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-API-01 | — | Describe parses into `SObjectDescribe` | HTTP |
| T-API-02 | — | Query encodes SOQL in URL correctly | HTTP |
| T-API-03 | — | List objects parses summaries | HTTP |
| T-API-04 | — | 401 → `SalesforceAuthenticatedClient` recovery (existing tests extended) | HTTP |
| T-API-05 | NFR-HTTP-1 | Large composite create body → `Content-Encoding: gzip` | HTTP |
| T-API-06 | NFR-HTTP-1 | Small composite create body → no content encoding | HTTP |
| T-HTTP-01 | NFR-HTTP-1 | `GzipJsonContent` threshold + factory decompression flags | Unit |

---

## Session gate integration (ribbon entry)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-GATE-01 | common | Data operation without session → login flow invoked (existing `SessionGateTests`) | Unit |
| T-GATE-02 | FR-LO | Logout clears session/target, keeps refresh tokens; subsequent gated op shows login options | Unit |
| T-GATE-03 | FR-OPT-1 | Options always enabled; clear-cache disabled when not signed in | Excel |

---

## Options (`ConnectorOptions` / store)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-OPT-01 | FR-OPT-4 | Round-trip JSON store (existing `JsonConnectorOptionsStoreTests`) | Unit |
| T-OPT-02 | FR-OPT-2 | Defaults match `ConnectorOptions` product defaults when store has no keys | Unit |
| T-OPT-03 | design | `CompositeBatchSize` persisted and used by `RecordBatchSplitter` | Unit |

---

## Sheet projection (Core projector + Excel applier)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-PRJ-01 | NFR-QSR-1 | 100×10 result → one `Values` array 100×10 | Unit |
| T-PRJ-02 | common | Column format list length matches column count | Unit |
| T-PRJ-03 | common | Row outcomes only for failed rows | Unit |
| T-PRJ-04 | common | Applier applies one Value assign per projection (Excel spy) | Excel |

---

## Wizard (`WizardTableLayoutBuilder`)

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-WIZ-01 | FR-TQW-6 | Field order: Id, Name, required (std/custom), standard, custom, read-only (std/custom); describe index within bucket | Unit |
| T-WIZ-02 | FR-TQW-5 | Default clause when none → equivalent WHERE plan | Unit |
| T-WIZ-03 | FR-TQW-7 | Wizard completion calls same `QueryTable.Run` as ribbon | Unit |
| T-WIZ-04 | FR-TQW-6 | Eight distinct header fills; Name before Required; `nameField` parse; display text without flags | Unit |

---

## COM API parity

| ID | Req | Test | Tag |
|----|-----|------|-----|
| T-COM-01 | VBA | Each of six COM methods calls the same `DataPlane` method as ribbon | Unit |
| T-COM-02 | VBA | `RefreshTableDataApi` passes refresh flag to `QueryRows` | Unit |

---

## Suggested test file layout

```
SalesforceRestAddin.Tests/
├── Tables/ForceTableParserTests.cs          # T-TBL-*
├── Tables/SelectionLimitsTests.cs           # T-LIM-*
├── Soql/SoqlQueryBuilderTests.cs            # T-SOQL-*
├── Values/FieldValueConverterTests.cs       # T-VAL-*
├── DataPlane/RecordBatchSplitterTests.cs    # T-BAT-*
├── DataPlane/QueryTableTests.cs           # T-QTD-*
├── DataPlane/QueryRowsTests.cs            # T-QSR-*
├── DataPlane/UpdateCellsTests.cs          # T-USC-*
├── DataPlane/InsertRowsTests.cs           # T-ISR-*
├── DataPlane/DeleteRecordsTests.cs        # T-DR-*
├── DataPlane/DescribeObjectTests.cs       # T-DSO-*
└── Rest/SalesforceDataClientTests.cs        # T-API-*
```

---

## Fixtures

| Fixture | Purpose |
|---------|---------|
| `AccountDescribe.json` | Minimal describe with Id, Name, custom field, picklist, reference |
| `AccountQueryPage1.json` / `Page2.json` | Paginated query |
| `CompositeSaveSuccess.json` / `CompositeSavePartial.json` | Update/create results |
| `ForceTableSnapshots.cs` | Factory methods for common table arrays |
| `SalesforceMockResponses.cs` | Extend existing helpers for composite + query |

---

## Out of scope (no tests)

- Translation Helper ribbon (8 buttons)
- SOAP describe/logout
- Per-cell Excel COM loops in hot paths (ensure **absent** via review, not runtime test)

---

## CI command

```bash
./dev.sh test -- --filter DataPlane
./dev.sh test -- --filter Soql
./dev.sh test -- --filter ForceTable
```

Filter names should match test class `Category` or namespace once implemented.
