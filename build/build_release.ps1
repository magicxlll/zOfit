$ErrorActionPreference = "Stop"
$Root = "D:\03_APP\3. System\GITHUB\zOfit"

Write-Host "====================================================="
Write-Host " zOfit QS Pro v2.1 - Full Build, Test & EXE Packager "
Write-Host "====================================================="

Write-Host "`n[1/5] Running .NET 8 xUnit tests..."
dotnet test (Join-Path $Root "engine\zOfit.Engine.Tests\zOfit.Engine.Tests.csproj") --verbosity minimal
if ($LASTEXITCODE -ne 0) { throw "dotnet test failed!" }

Write-Host "`n[2/5] Publishing single-file zofit-engine.exe..."
dotnet publish (Join-Path $Root "engine\zOfit.Engine\zOfit.Engine.csproj") `
    -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true `
    -o (Join-Path $Root "dist\engine")
if ($LASTEXITCODE -ne 0) { throw "dotnet publish engine failed!" }

Write-Host "`n[3/5] Building zOfit_QS_Pro.xlam + CustomUI14 Ribbon..."
powershell -ExecutionPolicy Bypass -File (Join-Path $Root "build\build_xlam.ps1")
if ($LASTEXITCODE -ne 0) { throw "build_xlam.ps1 failed!" }

Write-Host "`n[4/5] Running v1.0 Baseline & v2.1 E2E Excel COM Verification..."
powershell -ExecutionPolicy Bypass -File (Join-Path $Root "build\test_xlam.ps1")
if ($LASTEXITCODE -ne 0) { throw "test_xlam.ps1 failed!" }
powershell -ExecutionPolicy Bypass -File (Join-Path $Root "build\test_v2_e2e.ps1")
if ($LASTEXITCODE -ne 0) { throw "test_v2_e2e.ps1 failed!" }

Write-Host "`n[5/5] Packaging 1-Click Installer zOfit_QS_Pro_Setup.exe..."
dotnet publish (Join-Path $Root "installer\zOfit.Installer\zOfit.Installer.csproj") `
    -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true `
    -o (Join-Path $Root "dist\setup")
if ($LASTEXITCODE -ne 0) { throw "dotnet publish installer failed!" }

$setupExe = Join-Path $Root "dist\setup\zOfit_QS_Pro_Setup.exe"
$sizeMb = [math]::Round((Get-Item $setupExe).Length / 1MB, 2)
Write-Host "`nRELEASE COMPLETE! Installer: $setupExe ($sizeMb MB)"
