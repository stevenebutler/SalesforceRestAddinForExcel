<#
.SYNOPSIS
    Installs the Salesforce REST Excel add-in for the current user.

.DESCRIPTION
    Uses a matching packed XLL beside this script when one is available, enabling
    offline installs. Otherwise it checks the GitHub latest release.

.EXAMPLE
    ./Install-SalesforceRestAddin.ps1
    ./Install-SalesforceRestAddin.ps1 -Uninstall
    ./Install-SalesforceRestAddin.ps1 -UninstallForceConnector
#>

[CmdletBinding()]
param(
    [string]$GitHubRepo = "stevenebutler/SalesforceRestAddinForExcel",
    [string]$ReleaseTag = "latest",
    [string]$InstallDir = "$env:LOCALAPPDATA\SalesforceRestAddin",
    [switch]$Uninstall,
    [switch]$UninstallForceConnector
)

$ErrorActionPreference = "Stop"

function Get-RegistryViews {
    return @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)
}

function Get-ExcelBitness {
    foreach ($view in Get-RegistryViews) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
        try {
            $key = $base.OpenSubKey("SOFTWARE\Microsoft\Office\ClickToRun\Configuration")
            try {
                if ($key) {
                    $platform = $key.GetValue("Platform")
                    if ($platform) { return $platform }
                }
            } finally { if ($key) { $key.Dispose() } }
        } finally { $base.Dispose() }
    }

    foreach ($view in Get-RegistryViews) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
        try {
            $key = $base.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE")
            try {
                $exePath = if ($key) { $key.GetValue("") } else { $null }
                if ($exePath -and (Test-Path $exePath)) {
                    $bytes = [System.IO.File]::ReadAllBytes($exePath)
                    $machine = [BitConverter]::ToUInt16($bytes, [BitConverter]::ToInt32($bytes, 0x3C) + 4)
                    if ($machine -eq 0x8664) { return "x64" }
                    if ($machine -eq 0x014c) { return "x86" }
                    throw "Unrecognized Excel binary architecture (0x$($machine.ToString('X')))."
                }
            } finally { if ($key) { $key.Dispose() } }
        } finally { $base.Dispose() }
    }

    throw "Could not locate Excel on this machine. Is Excel installed?"
}

function Get-Release {
    param([string]$Repo, [string]$Tag)
    $headers = @{ "User-Agent" = "SalesforceRestAddin-Installer" }
    return Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases/tags/$Tag" -Headers $headers
}

function Get-InstallManifest {
    $manifestPath = Join-Path $InstallDir "installer-manifest.json"
    if (-not (Test-Path $manifestPath)) { return $null }
    try { return Get-Content -Raw $manifestPath | ConvertFrom-Json } catch { return $null }
}

function Write-InstallManifest {
    param([string]$XllPath, [string]$AssetName, [string]$ReleaseTagValue, [string]$TargetCommitish)
    [ordered]@{
        repository = $GitHubRepo
        releaseTag = $ReleaseTagValue
        targetCommitish = $TargetCommitish
        assetName = $AssetName
        installDirectory = $InstallDir
        xllPath = $XllPath
        installedAtUtc = [DateTimeOffset]::UtcNow.ToString("o")
    } | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $InstallDir "installer-manifest.json") -Encoding UTF8
}

function Get-OfficeOptionsPaths {
    param([Microsoft.Win32.RegistryKey]$BaseKey)
    $office = $BaseKey.OpenSubKey("SOFTWARE\Microsoft\Office")
    if (-not $office) { return @() }
    try {
        return @($office.GetSubKeyNames() |
            Where-Object { $_ -match '^\d+(\.\d+)+$' } |
            ForEach-Object { "SOFTWARE\Microsoft\Office\$($_)\Excel\Options" })
    } finally { $office.Dispose() }
}

function Test-SalesforceRestAddinXll {
    param([string]$Value)
    $searchStart = 0
    while ($searchStart -lt $Value.Length) {
        $extensionIndex = $Value.IndexOf('.xll', $searchStart, [StringComparison]::OrdinalIgnoreCase)
        if ($extensionIndex -lt 0) { return $false }
        $separatorIndex = [Math]::Max($Value.LastIndexOf('\', $extensionIndex), $Value.LastIndexOf('/', $extensionIndex))
        $fileName = $Value.Substring($separatorIndex + 1, $extensionIndex - $separatorIndex - 1)
        if ($fileName.StartsWith('SalesforceRestAddin', [StringComparison]::OrdinalIgnoreCase)) { return $true }
        $searchStart = $extensionIndex + 4
    }
    return $false
}

function Get-XllVersion {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $offset = -[Math]::Min($stream.Length, 4096)
        [void]$stream.Seek($offset, [System.IO.SeekOrigin]::End)
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $false)
        try {
            $tail = $reader.ReadToEnd()
            $matches = [regex]::Matches($tail, 'SalesforceRestAddin-Version:\s*([^\r\n\0]+)')
            if ($matches.Count -gt 0) { return $matches[$matches.Count - 1].Groups[1].Value.Trim() }
        } finally { $reader.Dispose() }
    } finally { $stream.Dispose() }
    return $null
}

function Remove-SalesforceRestAddinRegistrations {
    foreach ($view in Get-RegistryViews) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            foreach ($path in Get-OfficeOptionsPaths $base) {
                $key = $base.OpenSubKey($path, $true)
                if (-not $key) { continue }
                try {
                    foreach ($name in @($key.GetValueNames())) {
                        if ($name -notmatch '^OPEN\d*$') { continue }
                        $value = $key.GetValue($name)
                        if ($value -is [string] -and (Test-SalesforceRestAddinXll $value)) {
                            $key.DeleteValue($name, $false)
                            Write-Host "Removed $view startup registration '$name' ($value)."
                        }
                    }
                } finally { $key.Dispose() }
            }
        } finally { $base.Dispose() }
    }
}

function Register-ExcelAddin {
    param([string]$XllPath)
    Remove-SalesforceRestAddinRegistrations
    foreach ($view in Get-RegistryViews) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            $key = $base.CreateSubKey("SOFTWARE\Microsoft\Office\16.0\Excel\Options")
            try {
                $slot = "OPEN"
                if ($null -ne $key.GetValue($slot)) {
                    for ($number = 1; $number -lt 100; $number++) {
                        $candidate = "OPEN$number"
                        if ($null -eq $key.GetValue($candidate)) { $slot = $candidate; break }
                    }
                }
                if ($null -ne $key.GetValue($slot)) { throw "Excel options registry is full." }
                $key.SetValue($slot, "`"$XllPath`"", [Microsoft.Win32.RegistryValueKind]::String)
                Write-Host "Registered $view add-in under '$slot'."
            } finally { $key.Dispose() }
        } finally { $base.Dispose() }
    }
}

function Get-ForceConnectorUninstallEntries {
    $results = @()
    foreach ($hive in @([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryHive]::CurrentUser)) {
        foreach ($view in Get-RegistryViews) {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
            try {
                $uninstall = $base.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
                if (-not $uninstall) { continue }
                try {
                    foreach ($name in $uninstall.GetSubKeyNames()) {
                        $key = $uninstall.OpenSubKey($name)
                        try {
                            $displayName = $key.GetValue("DisplayName")
                            $command = $key.GetValue("UninstallString")
                            if ($displayName -and $command -and ($displayName -match 'ForceConnector|Force\.com Connector')) {
                                $results += [pscustomobject]@{ DisplayName = $displayName; Command = $command; Scope = $hive; View = $view }
                            }
                        } finally { if ($key) { $key.Dispose() } }
                    }
                } finally { if ($uninstall) { $uninstall.Dispose() } }
            } finally { $base.Dispose() }
        }
    }
    return @($results | Sort-Object DisplayName, Command -Unique)
}

function Get-EnabledForceConnectorExcelRegistrations {
    $results = @()
    foreach ($view in Get-RegistryViews) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            $addins = $base.OpenSubKey("SOFTWARE\Microsoft\Office\Excel\Addins")
            if (-not $addins) { continue }
            try {
                foreach ($name in $addins.GetSubKeyNames() | Where-Object { $_ -ieq "ForceConnector" }) {
                    $key = $addins.OpenSubKey($name)
                    try {
                        $loadBehavior = $key.GetValue("LoadBehavior")
                        if ($null -ne $loadBehavior -and [int]$loadBehavior -ne 0) {
                            $results += [pscustomobject]@{
                                View = $view
                                Path = "SOFTWARE\Microsoft\Office\Excel\Addins\$name"
                                LoadBehavior = [int]$loadBehavior
                            }
                        }
                    } finally { if ($key) { $key.Dispose() } }
                }
            } finally { $addins.Dispose() }
        } finally { $base.Dispose() }
    }
    return @($results)
}

function Disable-ForceConnectorExcelRegistrations {
    $registrations = @(Get-EnabledForceConnectorExcelRegistrations)
    if ($registrations.Count -eq 0) { return }

    Write-Warning "ForceConnector is enabled in Excel and may conflict with the Salesforce REST Add-in. Continuing will disable it for the current user."
    foreach ($registration in $registrations) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $registration.View)
        try {
            $key = $base.OpenSubKey($registration.Path, $true)
            try {
                if (-not $key) { throw "ForceConnector registration '$($registration.Path)' could not be opened for the current user." }
                $key.SetValue("LoadBehavior", 0, [Microsoft.Win32.RegistryValueKind]::DWord)
                Write-Host "Disabled ForceConnector Excel registration for $($registration.View)."
            } finally { if ($key) { $key.Dispose() } }
        } finally { $base.Dispose() }
    }
}

if ($UninstallForceConnector) {
    $entries = @(Get-ForceConnectorUninstallEntries)
    if ($entries.Count -eq 0) { throw "No Windows-installed ForceConnector product was found." }
    if ($entries.Count -gt 1) {
        Write-Host "More than one ForceConnector uninstaller was found:" -ForegroundColor Yellow
        for ($index = 0; $index -lt $entries.Count; $index++) { Write-Host "[$index] $($entries[$index].DisplayName) ($($entries[$index].Scope), $($entries[$index].View))" }
        $selection = [int](Read-Host "Enter the number to uninstall")
        if ($selection -lt 0 -or $selection -ge $entries.Count) { throw "Invalid selection." }
        $entry = $entries[$selection]
    } else { $entry = $entries[0] }
    Write-Host "Starting the registered Windows uninstaller for $($entry.DisplayName)..."
    Start-Process -FilePath $env:ComSpec -ArgumentList "/c $($entry.Command)"
    Write-Host "Complete the Windows uninstaller, then restart Excel."
    return
}

if ($Uninstall) {
    Write-Host "Removing Salesforce REST Add-in startup registrations..."
    Remove-SalesforceRestAddinRegistrations
    if (Test-Path $InstallDir) {
        Get-ChildItem -Path $InstallDir -Filter 'SalesforceRestAddin*.xll' -File -ErrorAction SilentlyContinue |
            Remove-Item -Force
        Write-Host "Removed installed XLL files. JSON settings were kept."
    }
    Write-Host "Done. Restart Excel for the change to take effect."
    return
}

Write-Host "Detecting Excel bitness..."
$bitness = Get-ExcelBitness
$assetName = if ($bitness -eq "x64") { "SalesforceRestAddin64-packed.xll" } else { "SalesforceRestAddin-packed.xll" }
$xllPath = Join-Path $InstallDir $assetName
$bundledPath = Join-Path $PSScriptRoot $assetName
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
Disable-ForceConnectorExcelRegistrations

if (Test-Path $bundledPath) {
    $bundledVersion = Get-XllVersion $bundledPath
    $installedVersion = Get-XllVersion $xllPath
    if ($bundledVersion -and $bundledVersion -eq $installedVersion) {
        Write-Host "Latest bundled version $bundledVersion is already installed; repairing the installation."
    } else {
        $displayBundledVersion = if ($bundledVersion) { $bundledVersion } else { "unknown" }
        Write-Host "Using bundled local XLL version $displayBundledVersion (treated as latest): $bundledPath"
    }
    Copy-Item -LiteralPath $bundledPath -Destination $xllPath -Force
    $releaseTagValue = "bundled-local"
    $targetCommitish = if ($bundledVersion) { $bundledVersion } else { "unknown" }
} else {
    Write-Host "No matching XLL beside this script; checking GitHub release '$ReleaseTag'..."
    $release = Get-Release -Repo $GitHubRepo -Tag $ReleaseTag
    $asset = $release.assets | Where-Object { $_.name -eq $assetName } | Select-Object -First 1
    if (-not $asset) { throw "Asset '$assetName' was not found in release '$ReleaseTag'." }
    $installedVersion = Get-XllVersion $xllPath
    $alreadyCurrent = $installedVersion -and $installedVersion -eq $release.target_commitish
    if ($alreadyCurrent) {
        Write-Host "Latest GitHub version $installedVersion is already installed; repairing the registration."
    } else {
        Write-Host "Downloading $assetName..."
        Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $xllPath -Headers @{ "User-Agent" = "SalesforceRestAddin-Installer" }
        Unblock-File -Path $xllPath
    }
    $releaseTagValue = $release.tag_name
    $targetCommitish = $release.target_commitish
}

Write-Host "Registering with Excel..."
Register-ExcelAddin -XllPath $xllPath
Write-InstallManifest -XllPath $xllPath -AssetName $assetName -ReleaseTagValue $releaseTagValue -TargetCommitish $targetCommitish
Write-Host "Install complete. Restart Excel (fully close and reopen) for the add-in to load." -ForegroundColor Green
