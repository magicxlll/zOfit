@echo off
chcp 65001 >nul
echo ============================================================
echo   zOfit QS Pro - Cai dat tu dong Excel Add-in (.xlam)
echo ============================================================
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$src = 'D:\03_APP\3. System\GITHUB\zOfit\dist\zOfit_QS_Pro.xlam';" ^
  "$addinsDir = Join-Path $env:APPDATA 'Microsoft\AddIns';" ^
  "if (-not (Test-Path $addinsDir)) { New-Item -ItemType Directory -Path $addinsDir | Out-Null };" ^
  "$dst = Join-Path $addinsDir 'zOfit_QS_Pro.xlam';" ^
  "Copy-Item $src $dst -Force;" ^
  "$xl = New-Object -ComObject Excel.Application;" ^
  "$xl.Visible = $false;" ^
  "$xl.Workbooks.Add() | Out-Null;" ^
  "$addin = $xl.AddIns.Add($dst, $true);" ^
  "$addin.Installed = $true;" ^
  "$xl.Quit();" ^
  "[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null;" ^
  "Write-Host 'Da cai dat va kich hoat thanh cong:' $dst"
pause
