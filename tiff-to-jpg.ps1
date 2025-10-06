#powershell -NoProfile -ExecutionPolicy Bypass -File .\tiff-to-jpg.ps1
# Convert all .tif and .tiff files to .jpg and save them in a subfolder "JPG"

Add-Type -AssemblyName System.Drawing

$sourceFolder = Get-Location
$destFolder = Join-Path $sourceFolder "JPG"
if (!(Test-Path $destFolder)) {
    New-Item -ItemType Directory -Path $destFolder | Out-Null
}

# JPEG quality (0–100)
$jpegQuality = 90

# Prepare JPEG encoder
$encoder = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
    Where-Object { $_.MimeType -eq 'image/jpeg' }
$encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
$encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter(
    [System.Drawing.Imaging.Encoder]::Quality, [int]$jpegQuality
)

# Convert all .tif and .tiff files
Get-ChildItem -Path $sourceFolder -File | Where-Object { $_.Extension -match 'tif{1,2}$' } | ForEach-Object {
    try {
        $img = [System.Drawing.Image]::FromFile($_.FullName)
        $destPath = Join-Path $destFolder ($_.BaseName + ".jpg")
        $img.Save($destPath, $encoder, $encoderParams)
        $img.Dispose()
        Write-Host "Saved: $($_.BaseName).jpg"
    } catch {
        Write-Host "Error processing file $($_.Name): $_"
    }
}

Write-Host "Conversion finished. Output folder: $destFolder"
