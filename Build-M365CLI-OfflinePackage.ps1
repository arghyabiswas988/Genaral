<#
.SYNOPSIS
  Builds an offline package for Microsoft 365 CLI (cli-microsoft365) using 7-Zip for fast multi-threaded compression.

.DESCRIPTION
  Downloads Node.js LTS installer and cli-microsoft365 (with all dependencies).
  Prepares npm cache for offline use.
  Creates a ZIP using 7-Zip for faster compression.
#>

param (
    [string]$OutputFolder = "C:\M365CLI_Offline",
    [string]$SevenZipExe = "C:\Temp\7zr.exe"
)

# Validate 7-Zip executable
if (!(Test-Path $SevenZipExe)) {
    Write-Error "7-Zip executable not found at $SevenZipExe"
    exit 1
}

# --- 1. Create working directories ---
if (Test-Path $OutputFolder) { Remove-Item $OutputFolder -Recurse -Force }
New-Item -ItemType Directory -Path $OutputFolder | Out-Null

$NodeFolder = Join-Path $OutputFolder "NodeJS"
$PackageFolder = Join-Path $OutputFolder "npm_package"
$CacheFolder = Join-Path $OutputFolder "npm-cache"

New-Item -ItemType Directory -Path $NodeFolder | Out-Null
New-Item -ItemType Directory -Path $PackageFolder | Out-Null
New-Item -ItemType Directory -Path $CacheFolder | Out-Null

# --- 2. Download Node.js MSI ---
Write-Host ">>> Downloading Node.js LTS..."
$nodeUrl = "https://github.com/nodejs/node/releases/download/v20.11.0/node-v20.11.0-x64.msi"
$nodeInstaller = Join-Path $NodeFolder "node-v20.11.0-x64.msi"

if (!(Test-Path $nodeInstaller)) {
    Invoke-WebRequest -Uri $nodeUrl -OutFile $nodeInstaller
    Write-Host "Node.js downloaded successfully."
} else {
    Write-Host "Node.js installer already exists. Skipping download."
}

# --- 3. Install CLI locally ---
Write-Host ">>> Installing cli-microsoft365 and dependencies locally..."
Set-Location $PackageFolder
npm init -y | Out-Null
npm install @pnp/cli-microsoft365 --ignore-scripts

# --- 4. Prepare offline npm cache ---
Write-Host ">>> Preparing npm cache for offline use..."
npm config set cache $CacheFolder --global
npm install --cache $CacheFolder --prefer-offline

# --- 5. Create ZIP using 7-Zip ---
Write-Host ">>> Creating offline ZIP package using 7-Zip..."
$zipFile = Join-Path $OutputFolder "M365CLI_OfflinePackage.zip"

# Remove old zip if exists
if (Test-Path $zipFile) { Remove-Item $zipFile -Force }

# Change to OutputFolder parent directory to avoid including full path
Push-Location (Split-Path $OutputFolder)

# 7zr command:
# a = add to archive
# -tzip = zip format
# -mx=9 = maximum compression
# -mmt = multi-threaded
& $SevenZipExe a -tzip -mx=9 -mmt "$zipFile" (Split-Path $OutputFolder -Leaf)

Pop-Location

Write-Host ">>> Offline package created successfully!"
Write-Host "ZIP location: $zipFile"
