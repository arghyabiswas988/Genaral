<#
.SYNOPSIS
  Installs Node.js and Microsoft 365 CLI (cli-microsoft365) fully offline (machine-wide)

.DESCRIPTION
  Run this script as SYSTEM (Intune, SCCM, or task scheduler).
  Installs Node.js into Program Files and CLI globally, so all users can run `m365`.
#>

param (
    [string]$SourceFolder = "C:\M365CLI_Offline"
)

$nodeInstaller = Join-Path $SourceFolder "node-v20.11.0-x64.msi"
$npmPackage    = Join-Path $SourceFolder "npm_package"
$npmCache      = Join-Path $SourceFolder "npm-cache"

if (!(Test-Path $nodeInstaller)) {
    Write-Error "Node.js installer not found at $nodeInstaller"
    exit 1
}

Write-Host ">>> Installing Node.js system-wide..."
Start-Process msiexec.exe -ArgumentList "/i `"$nodeInstaller`" /qn /norestart" -Wait

# Make sure Program Files path is in PATH
[Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\Program Files\nodejs", [System.EnvironmentVariableTarget]::Machine)

# Update current session PATH
$env:Path += ";C:\Program Files\nodejs"

Write-Host ">>> Installing cli-microsoft365 globally..."
Set-Location $npmPackage
npm install -g . --offline --cache $npmCache

Write-Host ">>> Verifying system-wide install..."
$exe = "C:\Program Files\nodejs\m365.cmd"
if (Test-Path $exe) {
    Write-Host "Microsoft 365 CLI installed successfully at $exe"
} else {
    Write-Warning "CLI not found in Program Files path, check logs."
}

Write-Host ">>> Done. All users on this machine can run 'm365'."
