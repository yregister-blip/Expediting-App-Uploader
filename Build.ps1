<#
================================================================================
 Build.ps1  -  Compiles Create-ExpeditingRelease.ps1 into a single .exe (PS2EXE).
================================================================================
 WHY
   The .ps1 is the human-readable, MSP-reviewable source of truth. The .exe is
   the deployed artifact: it runs regardless of PowerShell execution policy and
   is not casually editable by end users.

 PREREQUISITES (build machine only - NOT end-user machines)
   - Windows PowerShell 5.1 or PowerShell 7+
   - The 'ps2exe' module (installed below for the current user if missing;
     needs internet to the PowerShell Gallery, build machine only).

 USAGE
   powershell -ExecutionPolicy Bypass -File .\Build.ps1

 OUTPUT
   Create-ExpeditingRelease.exe  (no console window)

 NOTE ON SIGNING / SMARTSCREEN
   The produced .exe is UNSIGNED, so Windows SmartScreen / Defender will warn on
   first run. The fix is an OPS step, not a code change: have the MSP either
   code-sign the .exe or whitelist it. (Bitdefender's MSILHeracles quarantine of
   PS2EXE output is a known false positive - see the Playbook.)
================================================================================
#>

$ErrorActionPreference = 'Stop'

$here    = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcPath = Join-Path $here 'Create-ExpeditingRelease.ps1'
$exePath = Join-Path $here 'Create-ExpeditingRelease.exe'

if (-not (Test-Path $srcPath)) { throw "Source not found: $srcPath" }

if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "ps2exe module not found - installing for current user..."
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
    Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber
}
Import-Module ps2exe

Write-Host "Compiling '$srcPath' -> '$exePath' ..."
# -STA: the GUI uses the Windows clipboard ("Copy Log"), which requires a
# Single-Threaded Apartment. The compiled exe must be told explicitly.
Invoke-ps2exe `
    -inputFile  $srcPath `
    -outputFile $exePath `
    -noConsole `
    -STA `
    -title       "Create Expediting Release" `
    -description "Tekla -> Smartsheet expediting release sheet creator" `
    -company     "Qualico Steel"

if (Test-Path $exePath) {
    Write-Host ""
    Write-Host "BUILD OK -> $exePath"
    Write-Host "Reminder: the .exe is unsigned; have the MSP sign or hash-whitelist it before wide deployment."
} else {
    throw "Build did not produce $exePath"
}
