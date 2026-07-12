# Ribbon button requirements

Behavioral specifications for the **Salesforce Rest** ribbon (Excel-DNA + WPF). Each button has its own file; shared rules live in [common-performance-requirements.md](./common-performance-requirements.md) and UI policy in [ui-platform.md](./ui-platform.md).

For worksheet criteria examples, see [Query Criteria Operators](../query-criteria-operators.md).

| Doc | Purpose |
|-----|---------|
| [design.md](./design.md) | Module layout: cohesive folders, per-button Core files, Excel reader/writer |
| [test-cases.md](./test-cases.md) | Test plan with requirement IDs — implement test-first |
| [ui-platform.md](./ui-platform.md) | WPF code-only (no XAML); WebView2 for OAuth; no WinForms |
| [common-performance-requirements.md](./common-performance-requirements.md) | Bulk Excel I/O, batching, **NFR-STA-1** (sole `ExcelStaAsyncHost` wait/pump) |

## Scope

| Button | Spec | Phase | COM / VBA surface |
|--------|------|-------|-------------------|
| [About](./about.md) | Info dialog | 3 | — |
| [Table Query Wizard](./table-query-wizard.md) | 4-step table builder + auto-query | 2–3 | — |
| [Update Selected Cells](./update-selected-cells.md) | PATCH composite update | 2 | `UpdateSelectedCellsApi` |
| [Insert Selected Rows](./insert-selected-rows.md) | POST composite create | 2 | `InsertSelectedRowsApi` |
| [Query Selected Rows](./query-selected-rows.md) | Composite retrieve by Id | 2 | `QuerySelectedRowsApi` |
| [Query Table Data](./query-table-data.md) | SOQL query → sheet | 2 | `QueryTableDataApi` |
| [Delete Records](./delete-records.md) | Composite delete | 2 | `DeleteSelectedRecordsApi` |
| [Describe Sforce Object](./describe-sforce-object.md) | Metadata worksheet | 2–3 | — |
| [Options](./options.md) | Connector toggles | 1–3 | — |
| [Logout](./logout.md) | End session | 1 | — |

**Out of scope:** Translation Helper / i18n metadata tooling (`METAAPI.*`). See [AGENTS.md](../../AGENTS.md).

**Related but not a ribbon button:** `RefreshTableDataApi` reuses the query-selected-rows flow with `RefreshAll=true` (auto-select contiguous Id rows). Spec: [query-selected-rows.md](./query-selected-rows.md#refresh-table-data-com-api).

## Shared table layout

Most data-plane buttons assume a **ForceConnector table** on the active worksheet:

```
Row 1:  [Object label] | [field] | [operator] | [value] | …   (query criteria; optional after wizard)
Row 2:  [Column labels] — cell comments hold API field names
Row 3+: [Data rows]     — must include an Id column (15/18-char Salesforce Id, or "New" for insert)
```

Object API name: comment on **A1**, or A1 value if no comment / no spaces.

Discovery: scan the header row left from the selection to the start of the contiguous header block; the object name is above that first column. Multi-table sheets separate tables with at least one blank column — see [query-table-data.md](./query-table-data.md) (FR-QTD-1).

## How to use these specs

1. Read [common-performance-requirements.md](./common-performance-requirements.md) and [ui-platform.md](./ui-platform.md).
2. Read [design.md](./design.md) for where code belongs (`Tables/`, `Soql/`, `DataPlane/QueryTable.cs`, etc.).
3. Add tests from [test-cases.md](./test-cases.md) before or alongside Core implementation.
4. Implement **Core** logic without Excel or UI references.
5. Implement **WPF** dialogs in `SalesforceRestAddin.Windows.Ui` — **C# only, no XAML** (see [ui-platform.md](./ui-platform.md)); WebView2 for OAuth only.
6. Implement **ExcelDna** as a thin layer: bulk range read/write, `QueueAsMacro`.
7. Wire ribbon and COM to the same `DataPlane` methods in `AddInHost`.
