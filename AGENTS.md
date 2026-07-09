# Salesforce REST Add-in for Excel — agent guide

Excel-DNA host with a testable **Core** library for Salesforce REST + OAuth, while keeping existing **VBA macro automation** working unchanged (`ForceConnector.NextGen`).

## Non‑negotiables

### VBA compatibility is a hard contract

Existing workbooks use `ConnectorAdaptor` unchanged:

```vb
Set addIn = Application.COMAddIns("ForceConnector.NextGen")
Set GetForceAutomationObject = addIn.Object
' .QueryTableDataApi, .UpdateSelectedCellsApi, etc.
```

Preserve:

- **ProgId:** `ForceConnector.NextGen`
- **COM API class GUIDs** in `ComApiIdentity` (must match legacy `SalesForceAddInApi`)
- **Six COM methods:** `QuerySelectedRowsApi`, `QueryTableDataApi`, `UpdateSelectedCellsApi`, `InsertSelectedRowsApi`, `DeleteSelectedRecordsApi`, `RefreshTableDataApi`

### Worksheet table invariants

Every connector table on a worksheet must satisfy:

1. **Row 1 anchor cell** — object **label** (display) with the object **API name** in the cell comment.
2. **Row 2 same column** — the Salesforce **Record Id** field (primary key from describe metadata). Id is **always present** and is **always the first field column** (aligned with the row 1 object cell).
3. **Remaining row 2 cells** — other selected fields (wizard order: Id → required → Name → standard → custom → read-only).
4. **Field mapping** — header labels and criteria fields resolve to API names via **describe metadata** (`FieldCatalog`), not ad-hoc string guessing. Prefer row-2 comment `API Name: …`, then label match, then API-name fallback.
5. **Multi-table sheets** — tables are separated by **at least one blank column** in the header row. Discovery scans left from the selection to the start of the contiguous header block; the object name is above that first column. Writes, Id write-back, and error styling stay within that column span.

See [query-table-data.md](docs/ribbon/query-table-data.md) (FR-QTD-1) and [table-query-wizard.md](docs/ribbon/table-query-wizard.md) (FR-TQW-4/6). Metadata downloads are cached **per Salesforce instance URL** (sandbox/production must never share a cache bucket) — see [common-performance-requirements.md](docs/ribbon/common-performance-requirements.md) (NFR-META-1).

## Architecture

```
SalesforceRestAddinForExcel.sln
├── src/SalesforceRestAddin.Core/          net48 + net10.0 — Salesforce + session (no Excel/UI)
├── src/SalesforceRestAddin.ExcelDna/      net48 — Excel-DNA host, COM shell, ribbon, range I/O
├── src/SalesforceRestAddin.Windows.Ui/    net48 — WPF dialogs (code-only, no XAML); WebView2 for OAuth
└── tests/SalesforceRestAddin.Tests/       net10.0 — TUnit (Linux CI/dev)
```

**Core** holds session state (`SessionContext`), OAuth/REST, and pure logic. Dual TFM: **net48** for the Excel add-in; **net10.0** so tests can run on Linux.

**ExcelDna** is a thin host: COM registration, `SalesForceAddInApi`, ribbon, bulk worksheet read/write. Delegate to Core; avoid business logic here.

**Windows.Ui** is **WPF in C# only** (no XAML — see [ui-platform.md](docs/ribbon/ui-platform.md)) for dialogs. **WebView2** (`OAuthWebViewWindow`) is used only for the OAuth authorization redirect. **No WinForms** in `src/`.

**Tests** run on `net10.0` via `./dev.sh test`. Prefer test-first for Core behavior. See **Unit tests** below for determinism rules.

## Intentionally out of scope

- **Translation Helper / i18n metadata** (CustomLabel, Object Translation, METAAPI)
- **SOAP username/password login**
- **SOAP/WCF Connected Services**
- **VSTO installer / vdproj** — single packed XLL deployment instead

REST **describe** for standard field labels is fine; full metadata translation tooling is out of scope.

## Development workflow

### Bug reports (`Bug:`)

When the user starts a conversation with **`Bug:`** (or clearly opens with a new bug report):

1. **Append** the issue to [docs/bugs.md](docs/bugs.md) as the next numbered item in **open** state.
2. Capture the user-visible behaviour in plain language. Ask if reproduction steps are unclear.
3. Do **not** start implementing a fix unless the user asks.
4. When a fix lands, update the same item in `docs/bugs.md` (strike through + brief **Fixed:** note).

### Linux (primary)

**Every project in `src/` must compile on Linux** via `./dev.sh build` (container + `Microsoft.NETFramework.ReferenceAssemblies.net48`). Packed XLLs are produced in CI/dev containers. **Runtime** is Windows-only (Excel, WebView2, COM).

```bash
./dev.sh build   # net48: Core + Windows.Ui + ExcelDna + pack → ~/Downloads/SalesforceRestAddin
./dev.sh test    # net10.0: Core + TUnit
./dev.sh shell   # interactive container
```

Uses `docker compose` + `mcr.microsoft.com/dotnet/sdk:10.0`.

`FC_SKIP_EXCELDNA_BUILD=1` is only for faster Core/tests iteration.

Packed outputs under `~/Downloads/SalesforceRestAddin/`:

- `SalesforceRestAddin64-packed.xll` — **64‑bit Excel**
- `SalesforceRestAddin-packed.xll` — 32‑bit Excel

### WPF — compile on Linux (no XAML)

WPF UI is **code-only** (no `.xaml`, `<UseWPF>false</UseWPF>`). See [ui-platform.md](docs/ribbon/ui-platform.md).

| Project | WPF? | Required references / packages |
|---------|------|--------------------------------|
| `SalesforceRestAddin.Windows.Ui` | Defines dialogs | `<UseWPF>false</UseWPF>`; refs: `PresentationCore`, `PresentationFramework`, `WindowsBase`, `System.Xaml`; `Microsoft.Web.WebView2` for OAuth |
| `SalesforceRestAddin.ExcelDna` | Uses `Windows.Ui` | **ProjectReference** + same WPF framework refs on the host |
| `SalesforceRestAddin.Core` | **No UI** | Do not add WPF or WinForms refs |

### Windows VM (Excel smoke tests)

1. Overwrite the XLL at a **fixed path** (e.g. `Downloads\SalesforceRestAddin\SalesforceRestAddin64-packed.xll`).
2. Excel remembers the path in **File → Options → Add-ins → Excel Add-ins**.
3. **No regsvr32** — loading the XLL runs `AutoOpen` → COM registration.
4. Verify: `? Application.COMAddIns("ForceConnector.NextGen").Object Is Nothing` → `False`

Restart Excel after replacing the XLL at the same path.

### Excel STA async (NFR-STA-1)

After `await`, continuations are on the thread pool. Touching Excel COM there pins `EXCEL.EXE`. Use `ExcelComThread` — never COM off the Excel STA. Wait only via `ExcelStaAsyncHost` (modal WPF). See [common-performance-requirements.md](docs/ribbon/common-performance-requirements.md).

### Excel-DNA packing rule

In `SalesforceRestAddin.ExcelDna-AddIn.dna`:

- **`ExternalLibrary`** — only the main add-in assembly (`SalesforceRestAddin.dll`) with `ComServer="true"`
- **`Reference`** — dependency DLLs with `Pack="true"` (Core, Windows.Ui, System.Text.Json satellites, WebView2 managed)
- **Do not pack:** `Microsoft.Office.Interop.Excel`
- **WebView2:** pack managed DLLs; embed native `WebView2Loader.dll` in Windows.Ui (extracted to `%LOCALAPPDATA%\SalesforceRestAddin\native\`)
- **TLS on net48:** `TlsProtocolBootstrap.EnsureEnabled()` in `AutoOpen`

## Ribbon and data-plane specs

| Doc | Purpose |
|-----|---------|
| [docs/ribbon/README.md](docs/ribbon/README.md) | Index of per-button requirement specs |
| [docs/ribbon/common-performance-requirements.md](docs/ribbon/common-performance-requirements.md) | Bulk Excel I/O, batching, NFR-STA-1 |
| [docs/ribbon/design.md](docs/ribbon/design.md) | Module layout |
| [docs/ribbon/test-cases.md](docs/ribbon/test-cases.md) | Test plan |
| [docs/ribbon/ui-platform.md](docs/ribbon/ui-platform.md) | WPF code-only; WebView2 for OAuth |

**Implementation contract:** Salesforce logic lives in Core (`Tables/`, `Soql/`, `Values/`, `DataPlane/`); Excel types stay in ExcelDna. Core accepts `object[,]` snapshots and returns `SheetProjection`.

## Code quality

- **Small, named types**; explicit data flow; modern C# (nullable, async)
- **Zero warnings** — `TreatWarningsAsErrors` is on
- **Tests must be deterministic** — no `Thread.Sleep` / timing coordination; use orchestration (`SequentialMockHttpHandler`, fakes, `await`)

## Quick commands

```bash
./dev.sh build
./dev.sh test
./dev.sh test -- --filter SessionContext
FC_SKIP_EXCELDNA_BUILD=1 ./dev.sh build
```
