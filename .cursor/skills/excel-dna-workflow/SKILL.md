---
name: excel-dna-workflow
description: >-
  Builds, packs, and smoke-tests the Salesforce REST Add-in for Excel on Linux and
  Windows. Use when editing .dna files, SalesforceRestAddin.ExcelDna, XLL deployment,
  COM registration, or ConnectorAdaptor VBA testing.
---

# Excel-DNA workflow

## Solution layout

- `src/SalesforceRestAddin.ExcelDna/` — net48 Excel-DNA host
- `SalesforceRestAddin.ExcelDna-AddIn.dna` — pack + COM server config
- Output: `bin/{Debug|Release}/net48/publish/*-packed.xll`

## Build (Linux)

**Required:** `./dev.sh build` must succeed on Linux (compile + XLL pack). Runtime smoke test is Windows-only.

```bash
./dev.sh build
```

Copies to:

- `~/Downloads/SalesforceRestAddin/` — **VM deploy bundle** (packed XLLs only)

Native `WebView2Loader.dll` is **not** copied beside the XLL — it is embedded in `SalesforceRestAddin.Windows.Ui` and extracted to `%LOCALAPPDATA%\SalesforceRestAddin\native\` on first use.

Optional — skip XLL when iterating on Core/tests only:

```bash
FC_SKIP_EXCELDNA_BUILD=1 ./dev.sh build
```

## DNA packing — critical rules

**Wrong** — Core as `ExternalLibrary` → init error `No objects loaded from SalesforceRestAddin.Core`:

```xml
<ExternalLibrary Path="SalesforceRestAddin.Core.dll" ... />
```

**Correct** — main add-in vs dependencies:

```xml
<ExternalLibrary Path="SalesforceRestAddin.dll" ComServer="true" ExplicitExports="false"
                 LoadFromBytes="true" Pack="true" IncludePdb="false" />
<Reference Path="SalesforceRestAddin.Core.dll" LoadFromBytes="true" Pack="true" IncludePdb="false" />
<Reference Path="SalesforceRestAddin.Windows.Ui.dll" LoadFromBytes="true" Pack="true" IncludePdb="false" />
<!-- Plus every copy-local NuGet DLL Core needs on net48 — see SalesforceRestAddin.ExcelDna-AddIn.dna -->
```

Verify pack log includes at least:

```
ASSEMBLY_LZMA, Name: SALESFORCERESTADDIN
ASSEMBLY_LZMA, Name: SALESFORCERESTADDIN.CORE
ASSEMBLY_LZMA, Name: SALESFORCERESTADDIN.WINDOWS.UI
ASSEMBLY_LZMA, Name: SYSTEM.TEXT.JSON
ASSEMBLY_LZMA, Name: MICROSOFT.WEB.WEBVIEW2.CORE
ASSEMBLY_LZMA, Name: MICROSOFT.WEB.WEBVIEW2.WPF
```

## COM registration

- **No regsvr32** for normal use
- `ExcelAddIn.AutoOpen` → `ExcelComAddInHelper.LoadComAddIn` + `ComServer.DllRegisterServer()`
- ProgId: `ForceConnector.NextGen` (`ComApiIdentity.ComAddInProgId`) — unchanged for VBA

## Windows VM smoke test

1. Use **64-bit Excel** → `SalesforceRestAddin64-packed.xll`
2. Load once via **File → Options → Add-ins → Excel Add-ins → Browse**
3. Keep a **fixed path**; overwrite file after `./dev.sh build`; restart Excel
4. Immediate window:

```vb
? Application.COMAddIns("ForceConnector.NextGen").Object Is Nothing
```

Expect `False`.
