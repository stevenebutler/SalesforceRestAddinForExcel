Rolling build from `develop` (`{{SHORT_SHA}}`), built `{{BUILD_UTC}}`.

## Installer application (Windows)

Recommended for most users with / without internet access on the target machine:

## Online Install

1. Download `Install-SalesforceRestAddin.exe`.
2. In File Explorer, right-click the downloaded EXE and select **Properties**.
3. On the **General** tab, if an **Unblock** checkbox appears near the bottom of the window, select it. Then select **Apply** and **OK**.
4. Run `Install-SalesforceRestAddin.exe` and choose **Install / Update**.

The installer selects the correct package for supported Intel/AMD Office editions: 64-bit Office and older 32-bit Office. ARM-based Office is not currently supported. The installer downloads a package only when the installed XLL is not already the latest build, and registers the add-in for the current user without requiring administrator access. Fully close and reopen Excel after installation.

Before installation, the Installer application detects a Windows-installed ForceConnector add-in and recommends removing it first to avoid COM conflicts. It can launch the registered ForceConnector uninstaller for you.

## Offline install

Only use this when the target machine cannot access the internet. Download `Install-SalesforceRestAddin.exe` and both packed XLL files into the same folder on an internet-connected machine, transfer that folder to the target machine, unblock the EXE using the steps above, then run it and choose **Install / Update**. The installer selects the package for the target machine’s supported Intel/AMD Office edition without downloading anything.

- `SalesforceRestAddin64-packed.xll` — 64-bit Office on Intel/AMD Windows PCs
- `SalesforceRestAddin-packed.xll` — older 32-bit Office on Intel/AMD Windows PCs

## Advanced PowerShell fallback

The Installer application is simpler and is the recommended option. Use PowerShell only if you cannot use the installer application.

The following command downloads and runs the installer script in memory, so it works when PowerShell policy blocks execution of `.ps1` files:

```powershell
irm https://github.com/{{REPOSITORY}}/releases/latest/download/Install-SalesforceRestAddin.ps1 | iex
```

To uninstall the Salesforce REST add-in with the same policy-independent approach:

```powershell
& ([scriptblock]::Create((irm https://github.com/{{REPOSITORY}}/releases/latest/download/Install-SalesforceRestAddin.ps1))) -Uninstall
```

If your PowerShell execution policy permits scripts, you can instead download and run the script locally:

```powershell
irm https://github.com/{{REPOSITORY}}/releases/latest/download/Install-SalesforceRestAddin.ps1 -OutFile Install-SalesforceRestAddin.ps1
Unblock-File .\Install-SalesforceRestAddin.ps1
.\Install-SalesforceRestAddin.ps1
```

Running `.\Install-SalesforceRestAddin.ps1` or `.\Install-SalesforceRestAddin.ps1 -Uninstall` requires a PowerShell execution policy that allows script execution. To remove a Windows-installed ForceConnector VSTO add-in through PowerShell, use `.\Install-SalesforceRestAddin.ps1 -UninstallForceConnector`. The Installer application can also uninstall the Salesforce REST add-in.

Each successful `develop` build replaces this release.
