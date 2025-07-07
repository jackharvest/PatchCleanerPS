# PatchCleanerPS — Interactive & Auto Modes
# Version : 1.0.0
# Author  : jackharvest
# License : MIT
# Last Updated : 2025‑07‑07
# =============================================================================
# PowerShell successor to PatchCleaner. Scans **C:\Windows\Installer** for
# orphaned .MSI / .MSP files, then deletes or quarantines them. Empty
# sub-folders are pruned afterwards.
#
# ---------- Quick switches ----------
# -Auto         live run, default vendor filter "Acrobat"
# -AutoAll      live run, **no** vendor filter (aggressive)
# -AutoDry      dry-run, default vendor filter "Acrobat"
# -AutoDryAll   dry-run, no vendor filter
# (no switch)   interactive prompts + progress bar
#
# ---------- Safeties ----------
# • Restore point required for live runs
# • Elevation check for live runs
# • Vendor filter list (filename OR signer cert subject)
# • Quarantine option (Move-Run) instead of delete
# • Progress bar + color summary footer
# =============================================================================
param(
    [string[]] $ExcludeVendors = @(),
    [switch]   $Auto        = $false,
    [switch]   $AutoAll     = $false,
    [switch]   $AutoDry     = $false,
    [switch]   $AutoDryAll  = $false
)

# ---------------------------------------------------------------------------
# Determine run mode
# ---------------------------------------------------------------------------
$autoFlags = @($Auto,$AutoAll,$AutoDry,$AutoDryAll) | Where-Object { $_ }
if($autoFlags.Count -gt 1){ throw 'Use only ONE of -Auto, -AutoAll, -AutoDry, -AutoDryAll.' }
$Interactive = ($autoFlags.Count -eq 0)
$MoveMode    = $false
$WhatIf      = $true   # dry-run by default

switch ($true) {
    $Auto        { $WhatIf=$false; if(-not $ExcludeVendors){ $ExcludeVendors=@('Acrobat') } }
    $AutoAll     { $WhatIf=$false }
    $AutoDry     { $WhatIf=$true ; if(-not $ExcludeVendors){ $ExcludeVendors=@('Acrobat') } }
    $AutoDryAll  { $WhatIf=$true }
}

# ---------------------------------------------------------------------------
# Helper: human-readable sizes
# ---------------------------------------------------------------------------
function Format-Size([long]$bytes){
    if($bytes -lt 1MB){ return "{0:N2} KB" -f ($bytes/1KB) }
    elseif($bytes -lt 1GB){ return "{0:N2} MB" -f ($bytes/1MB) }
    else{ return "{0:N2} GB" -f ($bytes/1GB) }
}

# ---------------------------------------------------------------------------
# Interactive banner + prompts
# ---------------------------------------------------------------------------
if($Interactive){
    Clear-Host
    $banner=@'
 _____      _       _      _____ _                            _____   _____ 
|  __ \    | |     | |    / ____| |                          |  __ \ / ____|
| |__) |_ _| |_ ___| |__ | |    | | ___  __ _ _ __   ___ _ __| |__) | (___  
|  ___/ _` | __/ __| '_ \| |    | |/ _ \/ _` | '_ \ / _ \ '__|  ___/ \___ \ 
| |  | (_| | || (__| | | | |____| |  __/ (_| | | | |  __/ |  | |     ____) |
|_|   \__,_|\__\___|_| |_|\_____|_|\___|\__,_|_| |_|\___|_|  |_|    |_____/  
'@
    Write-Host $banner -ForegroundColor DarkYellow
    # Inform user about headless switches
    Write-Host 'Headless usage examples:' -ForegroundColor Cyan
    Write-Host '  powershell -File PatchCleanerScript.ps1 -Auto        # Live run with default Acrobat filter' -ForegroundColor White
    Write-Host '  powershell -File PatchCleanerScript.ps1 -AutoAll     # Live run with NO vendor filter (aggressive)' -ForegroundColor White
    Write-Host '  powershell -File PatchCleanerScript.ps1 -AutoDry     # Dry-run with default Acrobat filter' -ForegroundColor White
    Write-Host '  powershell -File PatchCleanerScript.ps1 -AutoDryAll  # Dry-run with NO vendor filter' -ForegroundColor White
    Write-Host ''
        Write-Host 'Select run mode:' -ForegroundColor Cyan
    Write-Host '  1) Dry-Run   (no changes)'
    Write-Host '  2) Move-Run  (quarantine)'
    Write-Host '  3) Delete-Run'
    $sel = Read-Host 'Enter 1, 2, or 3 [default 1]'; if(-not $sel){ $sel='1' }
    switch($sel){
        '2' { $WhatIf=$false; $MoveMode=$true; $QuarDir='C:\InstallerQuarantine'; if(-not (Test-Path $QuarDir)){ New-Item $QuarDir -ItemType Directory -Force | Out-Null } }
        '3' { $WhatIf=$false }
        default { $WhatIf=$true }
    }

    # elevation check for live runs
    if(-not $WhatIf){
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if(-not $isAdmin){ Write-Host 'ERROR: run PowerShell **as Administrator** for live actions.' -ForegroundColor Red; exit }
    }

    # vendor list prompt
    if(-not $ExcludeVendors){ $ExcludeVendors=@('Acrobat') }
    Write-Host "Current vendor filter list: $($ExcludeVendors -join ', ')" -ForegroundColor Cyan
    $add = Read-Host 'Add vendors (comma-separated) or press Enter';
    if($add){ $ExcludeVendors += ($add -split ',').Trim(); $ExcludeVendors = $ExcludeVendors | Select-Object -Unique }
}

# ---------------------------------------------------------------------------
# Paths & logger
# ---------------------------------------------------------------------------
$logDir = 'C:\Temp'
if(-not (Test-Path $logDir)){ New-Item $logDir -ItemType Directory -Force | Out-Null }
$logFile = Join-Path $logDir 'PatchCleanerPS.log'
function Log($msg){ (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')+"`t$msg" | Out-File -FilePath $logFile -Append -Encoding utf8 }
$insDir = 'C:\Windows\Installer'

# ---------------------------------------------------------------------------
# Restore point for live runs
# ---------------------------------------------------------------------------
if(-not $WhatIf){
    try{
        Checkpoint-Computer -Description 'PatchCleanerPS' -RestorePointType MODIFY_SETTINGS -ErrorAction Stop | Out-Null
        Write-Host 'System restore point created.' -ForegroundColor Green
        Log 'Restore point created.'
    }catch{
        Write-Host 'ERROR: failed to create restore point, aborting.' -ForegroundColor Red; exit 1
    }
}

# ---------------------------------------------------------------------------
# Build KEEP set (installed products + patches)
# ---------------------------------------------------------------------------
function Build-KeepSet{
    $keep=@{}; function AddK($n){ $keep[$n.ToLower()] = $true }
    $wi = New-Object -ComObject WindowsInstaller.Installer
    foreach($prod in $wi.Products()){ if($wi.ProductState($prod) -eq 5){ AddK (Split-Path ($wi.ProductInfo($prod,'LocalPackage')) -Leaf) } }
    $ud='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData'
    if(Test-Path $ud){ Get-ChildItem $ud | ForEach-Object{
        $patchKey = Join-Path $_.PsPath 'Patches'; if(Test-Path $patchKey){ Get-ChildItem $patchKey | ForEach-Object{ $p=Get-ItemProperty $_.PsPath -EA SilentlyContinue; if($p.LocalPackage){ AddK (Split-Path $p.LocalPackage -Leaf) } }}
        $prodKey  = Join-Path $_.PsPath 'Products'; if(Test-Path $prodKey){ Get-ChildItem $prodKey | ForEach-Object{ $ip=Join-Path $_.PsPath 'InstallProperties'; if(Test-Path $ip){ $iprop=Get-ItemProperty $ip -EA SilentlyContinue; if($iprop.LocalPackage){ AddK (Split-Path $iprop.LocalPackage -Leaf) } } }}
    }}
    return $keep
}

# check vendor exclusions ------------------------------------------------
function Is-Excluded($file){
    if(-not $ExcludeVendors){ return $false }
    foreach($v in $ExcludeVendors){
        if($file.Name -match $v){ return $true }
        try{ $sig=Get-AuthenticodeSignature $file.FullName; if($sig.SignerCertificate -and $sig.SignerCertificate.Subject -match $v){ return $true } }catch{}
    }
    return $false
}

# ---------------------------------------------------------------------------
# Main scan & action loop
# ---------------------------------------------------------------------------
$keepSet = Build-KeepSet
$files   = Get-ChildItem $insDir -Recurse -File -Force | Where-Object { $_.Extension -match '\.(msi|msp)$' }
$total   = $files.Count

# counters
[long]$usedBytes=0; $usedCnt=0
[long]$excBytes =0; $excCnt=0
[long]$orpBytes =0; $orpCnt=0

$i=0
foreach($f in $files){
    $i++
    if($Interactive){
        Write-Progress -Activity 'Scanning installer files' -Status "$i / $total" -PercentComplete ([math]::Floor($i/$total*100))
    }

    $size = $f.Length
    $key  = $f.Name.ToLower()

    if($keepSet[$key]){ $usedCnt++; $usedBytes+=$size; continue }
    if(Is-Excluded $f){ $excCnt++; $excBytes+=$size; continue }

    # orphaned
    $orpCnt++; $orpBytes+=$size
    if(-not $WhatIf){
        try{
            if($MoveMode){
                Move-Item $f.FullName -Destination $QuarDir -Force -Confirm:$false
            } else {
                Remove-Item $f.FullName -Force -Confirm:$false
            }
        }catch{}
    }
}
if($Interactive){ Write-Progress -Activity 'Scanning installer files' -Completed }

# ---------------------------------------------------------------------------
# Remove empty directories after cleanup
# ---------------------------------------------------------------------------
Get-ChildItem $insDir -Recurse -Directory -Force -ErrorAction SilentlyContinue |
    Sort-Object FullName -Descending | Where-Object {
        (Get-ChildItem $_ -Force -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0
    } | ForEach-Object {
        if(-not $WhatIf){ try{ Remove-Item $_.FullName -Recurse -Force -Confirm:$false }catch{} }
    }

# ---------------------------------------------------------------------------
# Summary banner
# ---------------------------------------------------------------------------
$usedStr = Format-Size $usedBytes
$excStr  = Format-Size $excBytes
$orpStr  = Format-Size $orpBytes
$mode    = if($WhatIf){ 'Completed (WhatIf)' } else { 'Completed' }

Write-Host '=======================' -ForegroundColor Cyan -BackgroundColor DarkBlue
Write-Host "=== $mode ===" -ForegroundColor Cyan -BackgroundColor DarkBlue
Write-Host '=======================' -ForegroundColor Cyan -BackgroundColor DarkBlue
Write-Host "=== $usedCnt files still used,  $usedStr ===" -ForegroundColor Red
Write-Host "=== $excCnt files excluded by filters,  $excStr ===" -ForegroundColor Yellow
if($WhatIf){
    Write-Host "=== $orpCnt files are orphaned,  $orpStr ===" -ForegroundColor Green
}elseif($MoveMode){
    Write-Host "=== $orpCnt orphaned files moved to quarantine,  $orpStr ===" -ForegroundColor Green
}else{
    Write-Host "=== $orpCnt orphaned files deleted,  $orpStr ===" -ForegroundColor Green
}
Write-Host "Log: $logFile"
