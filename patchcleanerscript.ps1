# PatchCleanerPS — Interactive & Auto Modes
# Version       : 1.0.0
# Author        : jackharvest
# License       : MIT
# Last Updated  : 2025-07-07
# =============================================================================
# PowerShell successor to PatchCleaner. Scans **C:\Windows\Installer** for
# orphaned .MSI / .MSP files, then deletes or quarantines them. Empty
# sub-folders are pruned afterwards.

param(
    [string[]] $ExcludeVendors = @(),
    [switch]   $Auto        = $false,
    [switch]   $AutoAll     = $false,
    [switch]   $AutoDry     = $false,
    [switch]   $AutoDryAll  = $false
)

# ---------------------------------------------------------------------------
# Update-check: Ensure this exe is the latest release (skip when running .ps1)
# ---------------------------------------------------------------------------
function Check-LatestVersion {
    param([string]$CurrentVersion)
    try {
        $apiUrl = 'https://api.github.com/repos/jackharvest/PatchCleanerPS/releases/latest'
        $resp   = Invoke-RestMethod -Uri $apiUrl -Headers @{ 'User-Agent' = 'PatchCleanerPS' } -ErrorAction Stop
        $latest = $resp.tag_name.TrimStart('v','V')
        if ([version]$CurrentVersion -lt [version]$latest) {
            Write-Warning "PatchCleanerPS v$latest is available (you have v$CurrentVersion)."
            Write-Warning "Download it here: https://github.com/jackharvest/PatchCleanerPS/releases"
            exit 1
        }
    } catch {
        Write-Verbose "Version check failed: $_"
    }
}

# Only run update-check when packaged as .exe, dynamically get version from exe
if ($MyInvocation.MyCommand.Path -match '\.exe$') {
    $scriptPath     = $MyInvocation.MyCommand.Path
    $currentVersion = (Get-Item $scriptPath).VersionInfo.FileVersion
    Check-LatestVersion -CurrentVersion $currentVersion
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Format-Size([long]$bytes) {
    if ($bytes -lt 1MB)   { "{0:N2} KB" -f ($bytes/1KB) }
    elseif ($bytes -lt 1GB){ "{0:N2} MB" -f ($bytes/1MB) }
    else                   { "{0:N2} GB" -f ($bytes/1GB) }
}

function Build-KeepSet {
    $keep = @{}; function AddK($n){ $keep[$n.ToLower()] = $true }
    $wi   = New-Object -ComObject WindowsInstaller.Installer
    foreach ($prod in $wi.Products()) {
        if ($wi.ProductState($prod) -eq 5) {
            AddK (Split-Path ($wi.ProductInfo($prod,'LocalPackage')) -Leaf)
        }
    }
    $ud = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData'
    if (Test-Path $ud) {
        Get-ChildItem $ud | ForEach-Object {
            $patchKey = Join-Path $_.PsPath 'Patches'
            if (Test-Path $patchKey) {
                Get-ChildItem $patchKey | ForEach-Object {
                    $p = Get-ItemProperty $_.PsPath -EA SilentlyContinue
                    if ($p.LocalPackage) { AddK (Split-Path $p.LocalPackage -Leaf) }
                }
            }
            $prodKey = Join-Path $_.PsPath 'Products'
            if (Test-Path $prodKey) {
                Get-ChildItem $prodKey | ForEach-Object {
                    $ip = Join-Path $_.PsPath 'InstallProperties'
                    if (Test-Path $ip) {
                        $iprop = Get-ItemProperty $ip -EA SilentlyContinue
                        if ($iprop.LocalPackage) { AddK (Split-Path $iprop.LocalPackage -Leaf) }
                    }
                }
            }
        }
    }
    return $keep
}

function Is-Excluded($file) {
    if (-not $ExcludeVendors) { return $false }
    foreach ($v in $ExcludeVendors) {
        if ($file.Name -match $v) { return $true }
        try {
            $sig = Get-AuthenticodeSignature $file.FullName
            if ($sig.SignerCertificate -and $sig.SignerCertificate.Subject -match $v) {
                return $true
            }
        } catch {}
    }
    return $false
}

# ---------------------------------------------------------------------------
# Core scan + action logic
# ---------------------------------------------------------------------------
function Run-Cleanup {
    param([bool]$Interactive)

    # Determine run mode flags
    $autoFlags = @($Auto, $AutoAll, $AutoDry, $AutoDryAll) | Where-Object { $_ }
    if ($autoFlags.Count -gt 1) {
        throw 'Use only ONE of -Auto, -AutoAll, -AutoDry, -AutoDryAll.'
    }
    $MoveMode = $false
    $WhatIf   = $true
    switch ($true) {
        $Auto       { $WhatIf = $false;  if (-not $ExcludeVendors) { $ExcludeVendors = @('Acrobat') } }
        $AutoAll    { $WhatIf = $false }
        $AutoDry    { $WhatIf = $true;   if (-not $ExcludeVendors) { $ExcludeVendors = @('Acrobat') } }
        $AutoDryAll { $WhatIf = $true }
    }

    if ($Interactive) {
        Clear-Host
        # Warn about admin rights only if not elevated
        $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Write-Host 'NOTE: Admin rights required for Move/Delete modes.' -ForegroundColor Yellow
        }

        # Banner & mode selection
        $banner = @'
 _____      _       _      _____ _                            _____   _____ 
|  __ \    | |     | |    / ____| |                          |  __ \ / ____|
| |__) |_ _| |_ ___| |__ | |    | | ___  __ _ _ __   ___ _ __| |__) | (___  
|  ___/ _` | __/ __| '_ \| |    | |/ _ \/ _` | '_ \ / _ \ '__|  ___/ \___ \ 
| |  | (_| | || (__| | | | |____| |  __/ (_| | | | |  __/ |  | |     ____) |
|_|   \__,_|\__\___|_| |_|\_____|_|\___|\__,_|_| |_|\___|_|  |_|    |_____/ 
'@
        Write-Host $banner -ForegroundColor DarkYellow
        Write-Host ' 1) Dry-Run (no changes)'
        Write-Host ' 2) Move-Run (quarantine)'
        Write-Host ' 3) Delete-Run'
        $sel = Read-Host 'Enter 1, 2, or 3 [default 1]'
        if (-not $sel) { $sel = '1' }
        switch ($sel) {
            '2' {
                $WhatIf = $false
                $MoveMode = $true
                $QuarDir = 'C:\InstallerQuarantine'
                if (-not (Test-Path $QuarDir)) {
                    New-Item $QuarDir -ItemType Directory -Force | Out-Null
                }
            }
            '3' { $WhatIf = $false }
            default { $WhatIf = $true }
        }

        # Elevation check for live runs
        if (-not $WhatIf) {
            $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
            if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                Write-Host 'ERROR: run PowerShell **as Administrator** for live actions.' -ForegroundColor Red
                exit 1
            }
        }

        # Vendor filter prompt
        if (-not $ExcludeVendors) { $ExcludeVendors = @('Acrobat') }
        Write-Host "Current vendor filter list: $($ExcludeVendors -join ', ')" -ForegroundColor Cyan
        $add = Read-Host 'Add vendors (comma-separated) or press Enter'
        if ($add) {
            $ExcludeVendors += ($add -split ',').Trim()
            $ExcludeVendors = $ExcludeVendors | Select-Object -Unique
        }
    }

    # Paths & logger
    $logDir  = 'C:\Temp'
    if (-not (Test-Path $logDir)) {
        New-Item $logDir -ItemType Directory -Force | Out-Null
    }
    $logFile = Join-Path $logDir 'PatchCleanerPS.log'
    function Log($msg) {
        (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + "`t$msg" |
            Out-File -FilePath $logFile -Append -Encoding utf8
    }
    $insDir = 'C:\Windows\Installer'

    # Create restore point for live runs
    if (-not $WhatIf) {
        try {
            Checkpoint-Computer -Description 'PatchCleanerPS' -RestorePointType MODIFY_SETTINGS -ErrorAction Stop | Out-Null
            Write-Host 'System restore point created.' -ForegroundColor Green
            Log 'Restore point created.'
        } catch {
            Write-Host 'ERROR: failed to create restore point, aborting.' -ForegroundColor Red
            exit 1
        }
    }

    # Build keep set & scan files
    $keepSet = Build-KeepSet
    $files   = Get-ChildItem $insDir -Recurse -File -Force | Where-Object { $_.Extension -match '\.(msi|msp)$' }
    $total   = $files.Count

    [long]$usedBytes = 0; $usedCnt = 0
    [long]$excBytes  = 0; $excCnt  = 0
    [long]$orpBytes  = 0; $orpCnt  = 0

    $i = 0
    foreach ($f in $files) {
        $i++
        if ($Interactive) {
            Write-Progress -Activity 'Scanning installer files' -Status "$i / $total" -PercentComplete ([math]::Floor($i/$total*100))
        }
        $size = $f.Length; $key = $f.Name.ToLower()
        if ($keepSet[$key]) {
            $usedCnt++; $usedBytes += $size
            continue
        }
        if (Is-Excluded $f) {
            $excCnt++; $excBytes += $size
            continue
        }
        # orphaned
        $orpCnt++; $orpBytes += $size
        if (-not $WhatIf) {
            if ($MoveMode) {
                Move-Item $f.FullName -Destination $QuarDir -Force -Confirm:$false
            } else {
                Remove-Item $f.FullName -Force -Confirm:$false
            }
        }
    }
    if ($Interactive) { Write-Progress -Activity 'Scanning installer files' -Completed }

    # Remove empty directories
    Get-ChildItem $insDir -Recurse -Directory -Force -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending |
        Where-Object { (Get-ChildItem $_ -Force -ErrorAction SilentlyContinue).Count -eq 0 } |
        ForEach-Object {
            if (-not $WhatIf) {
                try { Remove-Item $_.FullName -Recurse -Force -Confirm:$false } catch {}
            }
        }

    # Summary banner
    $usedStr = Format-Size $usedBytes
    $excStr  = Format-Size $excBytes
    $orpStr  = Format-Size $orpBytes
    $mode    = if ($WhatIf) { 'Completed (WhatIf)' } else { 'Completed' }

    Write-Host '=======================' -ForegroundColor Cyan -BackgroundColor DarkBlue
    Write-Host "=== $mode ===" -ForegroundColor Cyan -BackgroundColor DarkBlue
    Write-Host '=======================' -ForegroundColor Cyan -BackgroundColor DarkBlue
    Write-Host "=== $usedCnt files still used, $usedStr ===" -ForegroundColor Red
    Write-Host "=== $excCnt files excluded, $excStr ===" -ForegroundColor Yellow
    if ($WhatIf) {
        Write-Host "=== $orpCnt files orphaned, $orpStr ===" -ForegroundColor Green
    } elseif ($MoveMode) {
        Write-Host "=== $orpCnt files moved, $orpStr ===" -ForegroundColor Green
    } else {
        Write-Host "=== $orpCnt files deleted, $orpStr ===" -ForegroundColor Green
    }
    Write-Host "Log: $logFile"
}

# ---------------------------------------------------------------------------
# Entry point: headless vs interactive
# ---------------------------------------------------------------------------
if (@($Auto, $AutoAll, $AutoDry, $AutoDryAll) -contains $true) {
    # headless: single run, no prompts
    Run-Cleanup -Interactive:$false
} else {
    # interactive: allow repeating
    do {
        Run-Cleanup -Interactive:$true
        $resp = Read-Host 'Start over? [y/N]'
    } while ($resp.Trim().ToLower() -eq 'y')
}
