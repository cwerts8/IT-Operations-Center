# Convert Image to Base64 Helper Script
# This script converts an image file to Base64 string for embedding in PowerShell scripts

param(
    [Parameter(Mandatory=$true)]
    [string]$ImagePath
)

if (-not (Test-Path $ImagePath)) {
    Write-Host "ERROR: Image file not found: $ImagePath" -ForegroundColor Red
    exit
}

Write-Host "Converting image to Base64..." -ForegroundColor Cyan
Write-Host "Image: $ImagePath" -ForegroundColor Yellow

# Read the image file as bytes
$imageBytes = [System.IO.File]::ReadAllBytes($ImagePath)

# Convert to Base64
$base64String = [Convert]::ToBase64String($imageBytes)

# Output to console (can be copied)
Write-Host "`nBase64 String (copy this):" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Gray
Write-Host $base64String -ForegroundColor White
Write-Host "========================================" -ForegroundColor Gray

# Also save to a text file for easy copying
$outputFile = "$PSScriptRoot\Base64-Logo.txt"
$base64String | Out-File -FilePath $outputFile -Encoding UTF8
Write-Host "`nBase64 string saved to: $outputFile" -ForegroundColor Green

# Show file size info
$originalSize = (Get-Item $ImagePath).Length
Write-Host "`nOriginal file size: $([math]::Round($originalSize/1KB, 2)) KB" -ForegroundColor Cyan
Write-Host "Base64 string length: $($base64String.Length) characters" -ForegroundColor Cyan

Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "1. Copy the Base64 string above (or from Base64-Logo.txt)" -ForegroundColor White
Write-Host "2. Paste it into your PowerShell script as the `$logoBase64 variable" -ForegroundColor White
Write-Host "3. The script will convert it back to an image at runtime" -ForegroundColor White