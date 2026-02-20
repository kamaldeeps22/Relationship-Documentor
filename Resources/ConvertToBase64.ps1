<#
.SYNOPSIS
    Converts a PNG image to Base64 string for XrmToolBox plugin icons.

.DESCRIPTION
    This script reads a PNG file and converts it to a Base64-encoded string
    that can be used in the ExportMetadata attributes of XrmToolBox plugins.
    
    The Base64 string is displayed in the console and copied to clipboard.

.PARAMETER ImagePath
    Path to the PNG image file to convert.

.EXAMPLE
    .\ConvertToBase64.ps1 -ImagePath "Logo_32x32.png"
    
.EXAMPLE
    .\ConvertToBase64.ps1 "Logo_80x80.png"

.NOTES
    Author: Kamaldeep Singh
    Website: https://kamaldeepsingh.com
    
    Use this script when updating plugin icons:
    1. Edit your PNG files
    2. Run this script
    3. Copy output to RelationshipDocumentorPlugin.cs
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$ImagePath
)

# Check if file exists
if (-not (Test-Path $ImagePath)) {
    Write-Error "File not found: $ImagePath"
    exit 1
}

# Check if it's a PNG file
$extension = [System.IO.Path]::GetExtension($ImagePath)
if ($extension -ne ".png") {
    Write-Warning "Warning: File is not a PNG. Expected .png, got $extension"
}

try {
    # Read file bytes
    $bytes = [System.IO.File]::ReadAllBytes($ImagePath)
    
    # Convert to Base64
    $base64 = [System.Convert]::ToBase64String($bytes)
    
    # Get file info
    $fileInfo = Get-Item $ImagePath
    $fileSizeKB = [math]::Round($fileInfo.Length / 1KB, 2)
    $base64Length = $base64.Length
    
    # Display results
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Base64 Conversion Complete" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "File:           $ImagePath"
    Write-Host "File Size:      $fileSizeKB KB"
    Write-Host "Base64 Length:  $base64Length characters"
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Base64 String:" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host $base64
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    
    # Copy to clipboard if available
    if (Get-Command Set-Clipboard -ErrorAction SilentlyContinue) {
        $base64 | Set-Clipboard
        Write-Host "✓ Copied to clipboard!" -ForegroundColor Green
    } else {
        Write-Host "Note: Clipboard not available. Copy manually." -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "Usage in code:" -ForegroundColor Cyan
    Write-Host '  ExportMetadata("SmallImageBase64", "' -NoNewline
    Write-Host $base64.Substring(0, [Math]::Min(40, $base64.Length)) -NoNewline -ForegroundColor Gray
    Write-Host '...")' 
    Write-Host ""
    
} catch {
    Write-Error "Error converting file: $_"
    exit 1
}