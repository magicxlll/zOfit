# Generates 26 crisp CAD-style BMP icons (72x44 px) for frmBarShape UserForm
param(
    [string]$OutDir = "D:\03_APP\3. System\GITHUB\zOfit\src\data\shapes_bmp"
)

Add-Type -AssemblyName System.Drawing
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

$W = 76; $H = 44
$bgColor = [System.Drawing.Color]::White
$barColor = [System.Drawing.Color]::FromArgb(0, 70, 160)
$accentColor = [System.Drawing.Color]::FromArgb(192, 0, 0)
$textColor = [System.Drawing.Color]::FromArgb(60, 60, 60)

$penBar = New-Object System.Drawing.Pen($barColor, 2.4)
$penBar.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
$penBar.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$penBar.EndCap = [System.Drawing.Drawing2D.LineCap]::Round

$penHook = New-Object System.Drawing.Pen($accentColor, 2.0)
$penHook.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$penHook.EndCap = [System.Drawing.Drawing2D.LineCap]::Round

$font = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
$brush = New-Object System.Drawing.SolidBrush($textColor)

function Draw-Pts($g, $pts, $pen = $penBar) {
    $arr = New-Object 'System.Drawing.PointF[]' $pts.Count
    for ($i = 0; $i -lt $pts.Count; $i++) {
        $arr[$i] = New-Object System.Drawing.PointF(([float]($pts[$i][0] * ($W - 16) + 8)), ([float]($pts[$i][1] * ($H - 16) + 10)))
    }
    $g.DrawLines($pen, $arr)
}

$shapes = @{
    "S01" = @(@(0.05, 0.55), @(0.95, 0.55))
    "S02" = @(@(0.08, 0.25), @(0.08, 0.78), @(0.92, 0.78), @(0.92, 0.25))
    "S03" = @(@(0.08, 0.25), @(0.08, 0.78), @(0.92, 0.78))
    "S04" = @(@(0.08, 0.78), @(0.92, 0.78), @(0.92, 0.25))
    "S05" = @(@(0.08, 0.35), @(0.08, 0.78), @(0.92, 0.78), @(0.92, 0.15))
    "S06" = @(@(0.08, 0.20), @(0.08, 0.58), @(0.92, 0.58), @(0.92, 0.92))
    "S07" = @(@(0.22, 0.32), @(0.08, 0.32), @(0.08, 0.72), @(0.92, 0.72), @(0.92, 0.32), @(0.78, 0.32))
    "S08" = @(@(0.22, 0.32), @(0.08, 0.32), @(0.08, 0.72), @(0.92, 0.72))
    "S09" = @(@(0.35, 0.32), @(0.22, 0.16), @(0.82, 0.16), @(0.82, 0.88), @(0.22, 0.88), @(0.22, 0.16), @(0.38, 0.16))
    "S10" = @(@(0.20, 0.22), @(0.10, 0.42), @(0.10, 0.68), @(0.90, 0.68), @(0.90, 0.42), @(0.80, 0.22))
    "S11" = @(@(0.38, 0.38), @(0.22, 0.16), @(0.82, 0.16), @(0.82, 0.88), @(0.22, 0.88), @(0.22, 0.16), @(0.38, 0.38))
    "S12" = @(@(0.55, 0.28), @(0.50, 0.15), @(0.15, 0.85), @(0.85, 0.85), @(0.50, 0.15), @(0.42, 0.28))
    "S13" = "CIRCLE"
    "S14" = "HELIX"
    "S15" = @(@(0.05, 0.85), @(0.22, 0.85), @(0.22, 0.22), @(0.78, 0.22), @(0.78, 0.85), @(0.95, 0.85))
    "S16" = @(@(0.10, 0.82), @(0.10, 0.25), @(0.90, 0.25), @(0.90, 0.82))
    "S17" = @(@(0.05, 0.75), @(0.38, 0.75), @(0.62, 0.28), @(0.95, 0.28))
    "S18" = @(@(0.05, 0.78), @(0.22, 0.78), @(0.32, 0.28), @(0.68, 0.28), @(0.78, 0.78), @(0.95, 0.78))
    "S19" = @(@(0.05, 0.25), @(0.22, 0.25), @(0.36, 0.80), @(0.64, 0.80), @(0.78, 0.25), @(0.95, 0.25))
    "S20" = @(@(0.75, 0.20), @(0.15, 0.20), @(0.15, 0.85), @(0.75, 0.85))
    "S21" = @(@(0.05, 0.82), @(0.28, 0.82), @(0.68, 0.24), @(0.95, 0.24), @(0.95, 0.55))
    "S22" = @(@(0.08, 0.82), @(0.25, 0.82), @(0.25, 0.22), @(0.75, 0.22), @(0.75, 0.82), @(0.92, 0.82))
    "S23" = @(@(0.08, 0.22), @(0.08, 0.78), @(0.90, 0.78), @(0.90, 0.55))
    "S24" = @(@(0.15, 0.22), @(0.15, 0.82), @(0.85, 0.82), @(0.65, 0.55))
    "S25" = @(@(0.10, 0.25), @(0.10, 0.78), @(0.90, 0.78), @(0.90, 0.42))
    "S26" = @(@(0.32, 0.35), @(0.20, 0.20), @(0.20, 0.85), @(0.80, 0.85), @(0.80, 0.20), @(0.68, 0.35))
}

foreach ($code in ($shapes.Keys | Sort-Object)) {
    $bmp = New-Object System.Drawing.Bitmap($W, $H)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear($bgColor)
    $g.DrawString($code, $font, $brush, 2, 1)

    $spec = $shapes[$code]
    if ($spec -eq "CIRCLE") {
        $g.DrawEllipse($penBar, 24, 9, 28, 28)
        $g.DrawLine($penHook, 44, 11, 52, 19)
    } elseif ($spec -eq "HELIX") {
        for ($k = 0; $k -lt 4; $k++) {
            $x = 12 + $k * 13
            $g.DrawArc($penBar, $x, 14, 16, 22, 0, 360)
        }
    } else {
        Draw-Pts $g $spec $penBar
    }

    $outPath = Join-Path $OutDir "$code.bmp"
    $bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Bmp)
    $g.Dispose()
    $bmp.Dispose()
}

Write-Host "Generated 26 shape BMPs in $OutDir"
