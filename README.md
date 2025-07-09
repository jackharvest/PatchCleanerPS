![PatchCleanerPS](./icons/patchcleanerPSlogo.jpg)

# PatchCleanerPS
The spiritual successor to John Crawford's PatchCleaner at homedev.com.au, but written completely in powershell.

![Version](https://img.shields.io/badge/version-1.0.4-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Author](https://img.shields.io/badge/author-jackharvest-orange)

---

## ✨ $\textcolor{yellow}{\text{Features}}$

- Identifies **orphaned `.MSI` and `.MSP` files**
- Supports **interactive** and **automatic** modes
- Optional **dry-run** functionality to preview changes
- **Vendor exclusion list** support
- **Cleans up empty folders** in `C:\Windows\Installer`

---

## 🧰 $\textcolor{yellow}{\text{Requirements}}$

- Windows 10 or later
- PowerShell 5.1 or PowerShell Core
- Administrator privileges

---

## 🚀 $\textcolor{yellow}{\text{Getting Started}}$

### 🔹 $\textcolor{cyan}{\text{Download}}$

Download a release on the right, or download the script from this repository.

```powershell
Invoke-WebRequest -Uri "https://github.com/jackharvest/PatchCleanerPS/raw/main/patchcleanerscript.ps1" -OutFile "PatchCleanerPS.ps1"
```

## $\textcolor{yellow}{\text{Running It}}$
### $\textcolor{cyan}{\text{Interactive Mode}}$
.\PatchCleanerPS.ps1

## $\textcolor{yellow}{\text{Automatic mode (skips prompts)}}$
.\PatchCleanerPS.ps1 -Auto

## Supported Parameters
| Parameter         | Description                                                      |
| ----------------- | ---------------------------------------------------------------- |
| `-Auto`           | Automatically delete orphaned files (prompts for each one)       |
| `-AutoAll`        | Automatically delete **all** orphaned files without prompting    |
| `-AutoDry`        | Perform a dry-run (preview deletions), with confirmation prompts |
| `-AutoDryAll`     | Perform a dry-run (preview deletions), no prompts                |
| `-ExcludeVendors` | Exclude files from specified vendors (e.g. Microsoft, Adobe)     |

## $\textcolor{yellow}{\text{Example}}$
### Preview what would be deleted, no prompts, excluding Microsoft and Adobe installers
.\PatchCleanerPS.ps1 -AutoDryAll -ExcludeVendors 'Microsoft','Adobe'
or
PatchCleanerPS.exe -AutoDryAll -ExcludeVendors 'Microsoft','Adobe'

## $\textcolor{yellow}{\text{Safety First}}$
-Files are cross-referenced against installed applications and patches

-Move move supports just moving the files instead of deletion

-Dry-run mode is available to preview what would be affected without making changes

-You can exclude vendor-related files with -ExcludeVendors for added safety

## $\textcolor{yellow}{\text{What it does}}$
-Scans C:\Windows\Installer for .msi and .msp files

-Cross-checks each file against installed applications and updates in the system registry

-Flags unreferenced/orphaned files for deletion

-Optionally deletes empty subfolders afterward for a clean finish

## $\textcolor{yellow}{\text{Deployment Options}}$
-Compiled .exe builds are available in the Releases section:

-Great for deploying via SCCM, Intune, PDQ Deploy, or other RMM tools

-Mirrors the same logic as the .ps1 script

-Supports silent parameters for automated workflows

## $\textcolor{yellow}{\text{Credits}}$
Inspired by the retired PatchCleaner by John Crawford (homedev.com.au)
Rewritten entirely from scratch in PowerShell by jackharvest

## $\textcolor{yellow}{\text{Contributions}}$
-Pull requests are welcome!

If you have:

-Feature suggestions

-Bug reports

-Enterprise integration tips (SCCM/Intune/etc)

-Please open an issue or PR. Just make sure you test changes responsibly — deleting system installers is serious business!

## $\textcolor{yellow}{\text{License}}$

This project is licensed under the MIT License.
See the LICENSE file for details.

## $\textcolor{yellow}{\text{Changelog}}$
Changes can be found in the releases section on the right.

# $\textcolor{Cyan}{\text{Advanced Headless Usage Via Remote Powershell Script}}$
Let's say you wanted to utilize the "Scripts" section of SCCM for headless deployment. Below, you'll find an example of a highly dynamic script that auto downloads the latest release from github, stores it in a temporary location, performs a harsh -AutoDryAll, and returns the output.

```
# PowerShell script to fetch and run the latest PatchCleanerPS.exe in dry-run mode
# Safe for use in SCCM/Intune "Scripts" section — output is returned via stdout

$repoOwner = "jackharvest"
$repoName = "PatchCleanerPS"
$exeName = "PatchCleanerPS.exe"
$tempDir = "$env:ProgramData\PatchCleanerPS"
$exePath = Join-Path $tempDir $exeName

# Ensure temp directory exists
if (-not (Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
}

try {
    Write-Output "Checking GitHub API for latest release..."
    $releaseInfo = Invoke-RestMethod -Uri "https://api.github.com/repos/$repoOwner/$repoName/releases/latest" -Headers @{ "User-Agent" = "PatchCleanerPS-Agent" }

    $asset = $releaseInfo.assets | Where-Object { $_.name -eq $exeName }
    if (-not $asset) {
        throw "Could not find $exeName in latest GitHub release."
    }

    $downloadUrl = $asset.browser_download_url
    Write-Output "Downloading latest $exeName from: $downloadUrl"

    Invoke-WebRequest -Uri $downloadUrl -OutFile $exePath -UseBasicParsing

    if (-not (Test-Path $exePath)) {
        throw "Download failed or $exePath not found."
    }

    Write-Output "Running PatchCleanerPS in dry-run mode..."
    
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = $exePath
    $processInfo.Arguments = "-AutoDryAll"
    $processInfo.RedirectStandardOutput = $true
    $processInfo.UseShellExecute = $false
    $processInfo.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processInfo
    $process.Start() | Out-Null

    $stdout = $process.StandardOutput.ReadToEnd()
    $process.WaitForExit()

    Write-Output "---- PatchCleanerPS Output ----"
    Write-Output $stdout.Trim()
    Write-Output "--------------------------------"

} catch {
    Write-Error "ERROR: $($_.Exception.Message)"
    exit 1
}
```

In this scenario, running it against a single machine in SCCM would yield something like this:
![image](https://github.com/user-attachments/assets/98d985ea-c2bb-43f5-9cb6-bdf55eb031c0)

