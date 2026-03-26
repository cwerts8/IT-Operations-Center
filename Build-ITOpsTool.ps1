# Build-ExchangeAdminTool.ps1
# Converts IT-Operations-Center.ps1 to EXE with full metadata

#Requires -Version 5.1

$Version = "4.7.0"

param(
    [string]$SourcePath,
    [string]$OutputPath,
    [string]$IconPath,
    [switch]$NoAdmin
)

# Set defaults if not provided
if (-not $SourcePath) {
    $SourcePath = "$PSScriptRoot\IT-Operations-Center.ps1"
}
if (-not $OutputPath) {
    $OutputPath = "$PSScriptRoot\IT-Operations-Center_$Version.exe"
}
if (-not $IconPath) {
    $IconPath = "$PSScriptRoot\GellerIcon.ico"
}

Write-Host "======================================" -ForegroundColor Cyan
Write-Host "IT Ops Tool - Build Script" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

# Check if PS2EXE is installed
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "PS2EXE module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name ps2exe -Scope CurrentUser -Force -ErrorAction Stop
        Write-Host "PS2EXE installed successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to install PS2EXE: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# Import PS2EXE
Import-Module ps2exe -Force

# Verify source file exists
if (-not (Test-Path $SourcePath)) {
    Write-Host "Source file not found: $SourcePath" -ForegroundColor Red
    exit 1
}

# Check for logo file
if (-not (Test-Path "$PSScriptRoot\FullColorLogo.png")) {
    Write-Host "Warning: FullColorLogo.png not found in current directory" -ForegroundColor Yellow
    Write-Host "The EXE will need this file to display the logo" -ForegroundColor Yellow
    Write-Host ""
}

# Build parameters
$ps2exeParams = @{
    inputFile   = $SourcePath
    outputFile  = $OutputPath
    noConsole   = $true
    STA         = $true
    title       = "IT Operations Center Toolkit v$Version"
    description = "GUI-based Administration tool for Geller & Co."
    company     = "Geller & Co."
    product     = "IT Operations Center"
    version     = $Version
    copyright   = "Copyright 2026 Geller & Co. All rights reserved."
    verbose     = $true
}

# Add admin requirement unless -NoAdmin specified
if (-not $NoAdmin) {
    $ps2exeParams.requireAdmin = $true
}

# Add icon if provided
if ($IconPath -and (Test-Path $IconPath)) {
    $ps2exeParams.iconFile = $IconPath
    Write-Host "Using custom icon: $IconPath" -ForegroundColor Green
}

# Convert to EXE
Write-Host ""
Write-Host "Building EXE..." -ForegroundColor Cyan
Write-Host ""

try {
    Invoke-ps2exe @ps2exeParams -ErrorAction Stop
    
    Write-Host ""
    Write-Host "======================================" -ForegroundColor Green
    Write-Host "BUILD SUCCESSFUL!" -ForegroundColor Green
    Write-Host "======================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Output: $OutputPath" -ForegroundColor Cyan
    
    if (Test-Path $OutputPath) {
        $fileSize = [math]::Round((Get-Item $OutputPath).Length / 1MB, 2)
        Write-Host "Size: $fileSize MB" -ForegroundColor Cyan
    }
    
    Write-Host ""
    Write-Host "DEPLOYMENT CHECKLIST:" -ForegroundColor Yellow
    Write-Host "  [ ] FullColorLogo.png in same folder as EXE" -ForegroundColor White
    Write-Host "  [ ] ExchangeOnlineManagement module installed on target" -ForegroundColor White
    Write-Host "  [ ] ActiveDirectory module (RSAT) installed on target" -ForegroundColor White
    Write-Host "  [ ] User has appropriate Exchange admin permissions" -ForegroundColor White
    Write-Host ""
    
    # Offer to test
    $test = Read-Host "Would you like to test the EXE now? (Y/N)"
    if ($test -eq 'Y' -or $test -eq 'y') {
        Write-Host "Launching..." -ForegroundColor Cyan
        Start-Process $OutputPath
    }
}
catch {
    Write-Host ""
    Write-Host "======================================" -ForegroundColor Red
    Write-Host "BUILD FAILED!" -ForegroundColor Red
    Write-Host "======================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    exit 1
}

# Create deployment package
Write-Host ""
$package = Read-Host "Create deployment package? (Y/N)"
if ($package -eq 'Y' -or $package -eq 'y') {
    $packageName = "IT-Operations-Center_v${Version}_Deploy"
    $packagePath = Join-Path (Get-Location) $packageName
    
    if (Test-Path $packagePath) {
        Remove-Item $packagePath -Recurse -Force
    }
    
    New-Item -ItemType Directory -Path $packagePath | Out-Null
    
    # Copy files
    Copy-Item $OutputPath -Destination $packagePath
    if (Test-Path "$PSScriptRoot\FullColorLogo.png") {
        Copy-Item "$PSScriptRoot\FullColorLogo.png" -Destination $packagePath
    }
    
    # Create README
    $readme = @"
IT Operations Center Tool v$Version
========================================

PREREQUISITES:
- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1 or higher
- Exchange Online Management module
- Active Directory RSAT module (for Group Members feature)
- Appropriate Exchange Online admin permissions

INSTALLATION:
1. Copy this folder to a convenient location
2. Ensure FullColorLogo.png is in the same folder as the EXE
3. Install prerequisites:
   - Install-Module ExchangeOnlineManagement -Scope CurrentUser
   - Install RSAT from: Settings > Apps > Optional Features
4. Run IT-Operations-Center_$Version.exe

FEATURES:
- Mailbox Permissions (Full Access & Send As)
- Calendar Permissions Management
- Automatic Replies (Out of Office)
- AD Group Members Viewer
- Excel Export for all modules

SUPPORT:
For assistance, contact IT Administration - Desktop Engineering
Craig Werts - Geller & Co.

"@
    $readme | Out-File -FilePath (Join-Path $packagePath "README.txt") -Encoding UTF8
    
    Write-Host "Deployment package created: $packagePath" -ForegroundColor Green
    Write-Host ""
    Write-Host "You can now distribute this folder to users." -ForegroundColor Cyan
}

Write-Host ""
Write-Host "Build script complete!" -ForegroundColor Green