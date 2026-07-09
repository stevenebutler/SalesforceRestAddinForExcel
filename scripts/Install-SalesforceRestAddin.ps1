<#
.SYNOPSIS
    Installs the Salesforce REST Excel add-in (SalesforceRestAddin) for the current user.

.DESCRIPTION
    - Detects whether the installed Excel is 32-bit or 64-bit.
    - Downloads the matching packed .xll from the GitHub "latest" release.
    - Removes the "Mark of the Web" (Unblock-File) so Excel/Windows don't quarantine it.
    - Registers the add-in with Excel via the per-user OPEN/OPENx registry keys, so it
      loads automatically every time Excel starts — no manual Add-ins dialog needed.

.PARAMETER GitHubRepo
    "owner/repo" for the GitHub repository hosting the release.

.PARAMETER InstallDir
    Where the .xll is copied to locally. Defaults to a per-user AppData folder.

.EXAMPLE
    ./Install-SalesforceRestAddin.ps1
    ./Install-SalesforceRestAddin.ps1 -Uninstall
#>

[CmdletBinding()]
param(
    [string]$GitHubRepo = "stevenebutler/SalesforceRestAddinForExcel",
    [string]$ReleaseTag = "latest",
    [string]$InstallDir = "$env:LOCALAPPDATA\SalesforceRestAddin",
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"

# Office version key varies slightly across releases (16.0 covers Office 2016 through
# current Microsoft 365 builds, which is the vast majority of installs today).
$OfficeVersionKey = "16.0"
$ExcelOptionsKey  = "HKCU:\Software\Microsoft\Office\$OfficeVersionKey\Excel\Options"

function Get-ExcelBitness {
    # 1) Fast path: Click-to-Run installs (most Microsoft 365 setups today)
    $c2r = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" `
            -Name Platform -ErrorAction SilentlyContinue).Platform
    if ($c2r) { return $c2r }   # "x64" or "x86"

    # 2) Fallback: locate EXCEL.EXE via its registered App Path
    $exePath = (Get-ItemProperty `
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE" `
        -Name '(default)' -ErrorAction SilentlyContinue).'(default)'

    if (-not $exePath -or -not (Test-Path $exePath)) {
        throw "Could not locate Excel on this machine. Is Excel installed?"
    }

    # 3) Ground truth: read the PE header's Machine field directly from the binary
    $bytes    = [System.IO.File]::ReadAllBytes($exePath)
    $peOffset = [BitConverter]::ToInt32($bytes, 0x3C)
    $machine  = [BitConverter]::ToUInt16($bytes, $peOffset + 4)

    switch ($machine) {
        0x8664  { return "x64" }
        0x014c  { return "x86" }
        default { throw "Unrecognized Excel binary architecture (0x$($machine.ToString('X'))). Please install manually." }
    }
}

function Get-ReleaseAsset {
    param([string]$Repo, [string]$Tag, [string]$AssetName)

    $uri = "https://api.github.com/repos/$Repo/releases/tags/$Tag"
    Write-Verbose "Querying $uri"

    $headers = @{ "User-Agent" = "SalesforceRestAddin-Installer" }
    $release = Invoke-RestMethod -Uri $uri -Headers $headers

    $asset = $release.assets | Where-Object { $_.name -eq $AssetName }
    if (-not $asset) {
        $available = ($release.assets | Select-Object -ExpandProperty name) -join ", "
        throw "Asset '$AssetName' not found in release '$Tag'. Available assets: $available"
    }
    return $asset.browser_download_url
}

function Register-ExcelAddin {
    param([string]$XllPath)

    if (-not (Test-Path $ExcelOptionsKey)) {
        New-Item -Path $ExcelOptionsKey -Force | Out-Null
    }

    # Check whether this path is already registered under any OPEN/OPENn value.
    $existing = Get-Item $ExcelOptionsKey
    foreach ($name in $existing.Property) {
        if ($name -match '^OPEN\d*$') {
            $val = (Get-ItemProperty $ExcelOptionsKey -Name $name).$name
            if ($val -eq "`"$XllPath`"" -or $val -eq $XllPath) {
                Write-Host "Add-in already registered under '$name'. Nothing to do."
                return
            }
        }
    }

    # Find the next free OPEN / OPEN1 / OPEN2 ... slot.
    if (-not (Get-ItemProperty $ExcelOptionsKey -Name "OPEN" -ErrorAction SilentlyContinue)) {
        $slot = "OPEN"
    } else {
        $n = 1
        while (Get-ItemProperty $ExcelOptionsKey -Name "OPEN$n" -ErrorAction SilentlyContinue) { $n++ }
        $slot = "OPEN$n"
    }

    New-ItemProperty -Path $ExcelOptionsKey -Name $slot -Value "`"$XllPath`"" `
        -PropertyType String -Force | Out-Null

    Write-Host "Registered add-in under Excel registry value '$slot'."
}

function Unregister-ExcelAddin {
    param([string]$XllPath)

    if (-not (Test-Path $ExcelOptionsKey)) { return }

    $existing = Get-Item $ExcelOptionsKey
    foreach ($name in $existing.Property) {
        if ($name -match '^OPEN\d*$') {
            $val = (Get-ItemProperty $ExcelOptionsKey -Name $name).$name
            if ($val -eq "`"$XllPath`"" -or $val -eq $XllPath) {
                Remove-ItemProperty -Path $ExcelOptionsKey -Name $name
                Write-Host "Removed registry value '$name'."
            }
        }
    }
}

# ---------------------------------------------------------------------------

if ($Uninstall) {
    Write-Host "Uninstalling SalesforceRestAddin..."
    $bitness = Get-ExcelBitness
    $assetName = if ($bitness -eq "x64") { "SalesforceRestAddin64-packed.xll" } else { "SalesforceRestAddin-packed.xll" }
    $xllPath = Join-Path $InstallDir $assetName

    Unregister-ExcelAddin -XllPath $xllPath

    if (Test-Path $InstallDir) {
        Remove-Item $InstallDir -Recurse -Force
        Write-Host "Removed $InstallDir."
    }

    Write-Host "Done. Restart Excel for the change to take effect."
    return
}

Write-Host "Detecting Excel bitness..."
$bitness = Get-ExcelBitness
Write-Host "Detected: $bitness"

$assetName = if ($bitness -eq "x64") { "SalesforceRestAddin64-packed.xll" } else { "SalesforceRestAddin-packed.xll" }
Write-Host "Target asset: $assetName"

Write-Host "Looking up release '$ReleaseTag' in $GitHubRepo..."
$downloadUrl = Get-ReleaseAsset -Repo $GitHubRepo -Tag $ReleaseTag -AssetName $assetName

New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
$xllPath = Join-Path $InstallDir $assetName

Write-Host "Downloading $assetName..."
Invoke-WebRequest -Uri $downloadUrl -OutFile $xllPath -Headers @{ "User-Agent" = "SalesforceRestAddin-Installer" }

Write-Host "Unblocking file (removing Mark of the Web)..."
Unblock-File -Path $xllPath

Write-Host "Registering with Excel..."
Register-ExcelAddin -XllPath $xllPath

Write-Host ""
Write-Host "Install complete." -ForegroundColor Green
Write-Host "Installed to: $xllPath"
Write-Host "Restart Excel (fully close and reopen) for the add-in to load."
Write-Host ""
Write-Host "To uninstall later, run: .\Install-SalesforceRestAddin.ps1 -Uninstall"