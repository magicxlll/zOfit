param(
    [string]$SrcDir = "D:\03_APP\3. System\GITHUB\zOfit\Hinhdangthep",
    [string]$OutDir = "D:\03_APP\3. System\GITHUB\zOfit\src\data\shapes_bmp"
)

Add-Type -AssemblyName System.Drawing
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
Get-ChildItem $OutDir -Filter "*.bmp" | Remove-Item -Force -ErrorAction SilentlyContinue

$W = 240; $H = 84

for ($i = 1; $i -le 41; $i++) {
    $srcPath = Join-Path $SrcDir "dang ($i).png"
    if (-not (Test-Path $srcPath)) {
        throw "Missing source image: $srcPath"
    }
    $code = "D" + $i.ToString("00")
    $srcImg = [System.Drawing.Image]::FromFile($srcPath)

    $bmp = New-Object System.Drawing.Bitmap($W, $H)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.Clear([System.Drawing.Color]::White)

    # Fit image preserving aspect ratio inside 234x78 margin
    $ratioX = ($W - 6) / [double]$srcImg.Width
    $ratioY = ($H - 6) / [double]$srcImg.Height
    $ratio = [Math]::Min($ratioX, $ratioY)
    $drawW = [int][Math]::Round($srcImg.Width * $ratio)
    $drawH = [int][Math]::Round($srcImg.Height * $ratio)
    $offX = [int](($W - $drawW) / 2)
    $offY = [int](($H - $drawH) / 2)

    $g.DrawImage($srcImg, $offX, $offY, $drawW, $drawH)

    $outPath = Join-Path $OutDir "$code.bmp"
    $bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Bmp)
    $g.Dispose()
    $bmp.Dispose()
    $srcImg.Dispose()
}

Write-Host "Prepared 41 crisp BMPs (240x84) from Hinhdangthep in $OutDir"
