#Requires -Version 5.1
<#
.SYNOPSIS
  Install the Excel-DNA ForceConnector add-in on Windows (Phase 0 spike).

.DESCRIPTION
  Copies packed XLL outputs and registers the COM add-in as ForceConnector.NextGen.
  Run from repo root after: dotnet build src/ForceConnector.ExcelDna -c Release -f net48

  Windows VM acceptance test:
    1. Uninstall/disable the legacy VSTO ForceConnector add-in.
    2. Run this script (optionally -InstallDir).
    3. Start Excel; confirm Application.COMAddIns("ForceConnector.NextGen").Object is not Nothing.
    4. Run unchanged ConnectorAdaptor VBA (e.g. sfQuery stub).
#>
param(
    [string]$InstallDir = "$env:LOCALAPPDATA\ForceConnector",
    [ValidateSet("x64", "x86")]
    [string]$Bitness = "x64"
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$PublishDir = Join-Path $Root "src\ForceConnector.ExcelDna\bin\Release\net48\publish"
if (-not (Test-Path $PublishDir)) {
    $PublishDir = Join-Path $Root "src\ForceConnector.ExcelDna\bin\Debug\net48\publish"
}

$packedName = if ($Bitness -eq "x64") { "ForceConnector.ExcelDna-AddIn64-packed.xll" } else { "ForceConnector.ExcelDna-AddIn-packed.xll" }
$SourceXll = Join-Path $PublishDir $packedName
if (-not (Test-Path $SourceXll)) {
    throw "Packed XLL not found: $SourceXll. Build ForceConnector.ExcelDna first."
}

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
$TargetXll = Join-Path $InstallDir "ForceConnector.xll"
Copy-Item -Force $SourceXll $TargetXll

$BinDir = Split-Path -Parent $PublishDir
$RuntimesSource = Join-Path $BinDir "runtimes"
if (Test-Path $RuntimesSource) {
    $RuntimesTarget = Join-Path $InstallDir "runtimes"
    if (Test-Path $RuntimesTarget) {
        Remove-Item -Recurse -Force $RuntimesTarget
    }
    Copy-Item -Recurse -Force $RuntimesSource $RuntimesTarget
    Write-Host "Installed WebView2 native loaders to $RuntimesTarget"
}

Write-Host "Installed to $TargetXll"
Write-Host "Load in Excel: File -> Options -> Add-ins -> Excel Add-ins -> Browse -> select the .xll"
Write-Host "No regsvr32 required; COM registration runs when Excel loads the add-in (AutoOpen)."
Write-Host "Restart Excel, then test: Application.COMAddIns(""ForceConnector.NextGen"").Object"
