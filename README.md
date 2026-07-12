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

Each successful `develop` build publishes a rolling [**latest** release](https://github.com/stevenebutler/SalesforceRestAddinForExcel/releases/latest) with both XLLs, `Install-SalesforceRestAddin.exe`, and `Install-SalesforceRestAddin.ps1`.

### Installer application for simplified installation

Recommended for most users with internet access:

1. Download `Install-SalesforceRestAddin.exe`.
2. In File Explorer, right-click the downloaded EXE and select **Properties**.
3. On the **General** tab, if an **Unblock** checkbox appears near the bottom of the window, select it. Then select **Apply** and **OK**.
4. Run `Install-SalesforceRestAddin.exe` and choose **Install / Update**.

The installer selects the correct package for supported Intel/AMD Office editions: 64-bit Office and older 32-bit Office. ARM-based Office is not currently supported. The installer downloads a package only when the installed XLL is not already the latest build, and registers the add-in for the current user without requiring administrator access. Fully close and reopen Excel after installation.

Before installation, the Installer application detects whether legacy ForceConnector is enabled in Excel. If it is, continuing will disable it for the current user to avoid COM conflicts; this does not require administrator access or uninstall the Windows application. The optional **Uninstall ForceConnector** button can still launch the registered Windows uninstaller.

### Offline install

Only use this when the target machine cannot access the internet. Download `Install-SalesforceRestAddin.exe` and both packed XLL files into the same folder on an internet-connected machine, transfer that folder to the target machine, unblock the EXE using the steps above, then run it and choose **Install / Update**. The installer selects the package for the target machine’s supported Intel/AMD Office edition without downloading anything.

- `SalesforceRestAddin64-packed.xll` — 64-bit Office on Intel/AMD Windows PCs
- `SalesforceRestAddin-packed.xll` — older 32-bit Office on Intel/AMD Windows PCs

### Advanced PowerShell fallback

The Installer application is simpler and is the recommended option. Use PowerShell only if you cannot use the installer application.

This command downloads and runs the installer script in memory, so it works when PowerShell policy blocks execution of `.ps1` files:

```powershell
irm https://github.com/stevenebutler/SalesforceRestAddinForExcel/releases/latest/download/Install-SalesforceRestAddin.ps1 | iex
```

To uninstall the Salesforce REST add-in with the same policy-independent approach:

```powershell
& ([scriptblock]::Create((irm https://github.com/stevenebutler/SalesforceRestAddinForExcel/releases/latest/download/Install-SalesforceRestAddin.ps1))) -Uninstall
```

If your PowerShell execution policy permits scripts, you can instead download and run the script locally:

```powershell
irm https://github.com/stevenebutler/SalesforceRestAddinForExcel/releases/latest/download/Install-SalesforceRestAddin.ps1 -OutFile Install-SalesforceRestAddin.ps1
Unblock-File .\Install-SalesforceRestAddin.ps1
.\Install-SalesforceRestAddin.ps1
```

Running `.\Install-SalesforceRestAddin.ps1` or `.\Install-SalesforceRestAddin.ps1 -Uninstall` requires a PowerShell execution policy that allows script execution. To remove a Windows-installed ForceConnector VSTO add-in through PowerShell, use `.\Install-SalesforceRestAddin.ps1 -UninstallForceConnector`. The Installer application can also uninstall the Salesforce REST add-in. Uninstalling removes only the Salesforce REST XLL files and Excel registrations; login and other JSON settings are kept.

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
