# Salesforce REST Add-in for Excel

Bring Salesforce into Excel with REST and native OAuth Login support.

Excel-DNA add-in for querying, updating, inserting, and deleting Salesforce records from worksheets. OAuth login (PKCE public client), REST describe/query/composite CRUD, and VBA automation via ProgId `ForceConnector.NextGen` for existing `ConnectorAdaptor` macros.

## Layout

```
SalesforceRestAddinForExcel.sln
├── src/
│   ├── SalesforceRestAddin.Core/          # net48 + net10.0 — Salesforce + session (no Excel/UI)
│   ├── SalesforceRestAddin.ExcelDna/      # net48 — Excel-DNA host, COM, ribbon
│   └── SalesforceRestAddin.Windows.Ui/    # net48 — WPF dialogs (code-only); WebView2 for OAuth
├── tests/
│   └── SalesforceRestAddin.Tests/         # net10.0 — TUnit
└── docs/ribbon/                           # Behavioral specs
```

## Build / test (Linux)

```bash
./dev.sh build   # Core + Windows.Ui + ExcelDna pack → ~/Downloads/SalesforceRestAddin/
./dev.sh test    # Core net10.0 + TUnit
```

Packed outputs:

- `SalesforceRestAddin64-packed.xll` — 64-bit Excel
- `SalesforceRestAddin-packed.xll` — 32-bit Excel

## Install (Windows)

Each successful `develop` build publishes a rolling [**latest** release](https://github.com/stevenebutler/SalesforceRestAddinForExcel/releases/latest) with both XLLs and `Install-SalesforceRestAddin.ps1`.

In PowerShell:

```powershell
irm https://github.com/stevenebutler/SalesforceRestAddinForExcel/releases/latest/download/Install-SalesforceRestAddin.ps1 -OutFile Install-SalesforceRestAddin.ps1
Unblock-File .\Install-SalesforceRestAddin.ps1
.\Install-SalesforceRestAddin.ps1
```

Or download and run in one step:

```powershell
irm https://github.com/stevenebutler/SalesforceRestAddinForExcel/releases/latest/download/Install-SalesforceRestAddin.ps1 | iex
```

The script detects Excel bitness, downloads the matching packed XLL, removes Mark of the Web, and registers the add-in for the current user. Restart Excel after install. Uninstall with `.\Install-SalesforceRestAddin.ps1 -Uninstall`.

## VBA

```vb
Set addIn = Application.COMAddIns("ForceConnector.NextGen")
Set GetForceAutomationObject = addIn.Object
```

The ProgId `ForceConnector.NextGen` preserves VBA compatibility with enhancements the author of this project added in a fork of Force.com Connector Next Generation (see Credits).

## Docs

See [AGENTS.md](AGENTS.md) and [docs/ribbon/README.md](docs/ribbon/README.md).

## License

[MIT](LICENSE) — Copyright © 2026 Steven Butler.

## Credits

Inspired by the [Force.com Connector Next Generation](https://github.com/good-ghost/ForceConnector) Excel VSTO add-in (`good-ghost/ForceConnector`, MIT).

This REST + Excel-DNA project reimplements that worksheet/ribbon workflow with:
* Completely redesigned Excel interoperability for maximum performance.
* UI improvements for varying DPI settings.
* More integrated, modern OAuth authentication flows.

Reference fields show Ids rather than names to simplify VBA scripting of cross-table interactions. Other behaviour differences may exist; please raise them as issues for consideration.

