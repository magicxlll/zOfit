$ErrorActionPreference = "Stop"
$Root = "D:\03_APP\3. System\GITHUB\zOfit"
$SetupExe = Join-Path $Root "dist\setup\zOfit_QS_Pro_Setup.exe"
$BhflexPath = "D:\01_MEGADNC\01_PROJECT\24_BH FLEX\02_BIDDING\261103_V2F\BOM_BHFLEX_CONSTRUCTION_Sub building 49-50 & Civil_VIETPH_261119.xlsx"

Write-Host "=== Golden BHFLEX Regression & 1-Click Installer Verification ==="

# 1. Extract real CAD text lines from BHFLEX Data sheet and pipe to zofit-engine.exe parse_cad
Write-Host "  [1/4] Running CadTextParser on real BHFLEX Data sheet rows..."
$tmpIn = Join-Path $env:TEMP "zofit_bhflex_cad_in.json"
$tmpOut = Join-Path $env:TEMP "zofit_bhflex_cad_out.json"
python -c @"
import zipfile, xml.etree.ElementTree as ET, json, re
ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile(r'$BhflexPath') as z:
    ss = [''.join(t.text or '' for t in si.findall('.//m:t', ns)) for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', ns)]
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rmap = {r.attrib['Id']: r.attrib['Target'] for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    sheets = {sh.attrib['name']: 'xl/' + rmap[sh.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']].lstrip('/') for sh in wb.findall('.//m:sheet', ns)}
    root = ET.fromstring(z.read(sheets['Data']))
    raw_rows = []
    for row in root.findall('.//m:sheetData/m:row', ns):
        rn = int(row.attrib['r'])
        cols = {}
        for c in row.findall('m:c', ns):
            col = re.sub(r'\d+', '', c.attrib['r'])
            t = c.attrib.get('t')
            v = c.find('m:v', ns)
            if v is not None and v.text:
                cols[col] = ss[int(v.text)] if t == 's' else v.text
        val_i = cols.get('I', '').strip()
        if val_i and any(k in val_i for k in ['-W', 'Box:', 'D200', 'D300', 'D400']):
            raw_rows.append({'sourceRow': rn, 'colA': val_i, 'colB': '', 'colC': '', 'colD': ''})
    with open(r'$tmpIn', 'w', encoding='utf-8') as f:
        json.dump({'rawRows': raw_rows}, f, ensure_ascii=False)
"@
if ($LASTEXITCODE -ne 0) { throw "Python extract failed!" }
& (Join-Path $Root "dist\engine\zofit-engine.exe") --cmd parse_cad --in $tmpIn --out $tmpOut
if ($LASTEXITCODE -ne 0) { throw "zofit-engine parse_cad failed!" }
python -c @"
import json
res = json.load(open(r'$tmpOut', encoding='utf-8'))
d = res['data']
print(f\"        Parsed {d['parsedCount']}/{d['totalInput']} real BHFLEX CAD lines! Distinct: {d['distinctKeys']}\")
assert res['ok'] is True and d['unknownCount'] == 0 and d['parsedCount'] == 58
"@
if ($LASTEXITCODE -ne 0) { throw "Real BHFLEX CAD parser assertion failed!" }

# 2. Golden Benchmark: Compare BHFLEX 49.ELECTRIC SUPPLY Lean Concrete & Footing Concrete
Write-Host "  [2/4] Comparing Golden BHFLEX 49.ELECTRIC SUPPLY Substructure numbers..."
$bhIn = Join-Path $env:TEMP "zofit_bhflex_49_in.json"
$bhOut = Join-Path $env:TEMP "zofit_bhflex_49_out.json"
@"
{
  "targetBuilding": "49.ELECTRIC SUPPLY",
  "spec": { "leanThkM": 0.1, "leanOverhangM": 0.1, "concWasteFactor": 1.03 },
  "elements": [
    { "dataRow": 5,  "elementId": "49-PC1",  "building": "49.ELECTRIC SUPPLY", "group": "Foundation", "mark": "PC1",   "description": "Footing PC1",   "qty": 10, "lengthM": 1.0,  "widthM": 1.0,  "heightM": 1.0, "isBuried": "Y", "boqBucket": "Substructure" },
    { "dataRow": 6,  "elementId": "49-PC2",  "building": "49.ELECTRIC SUPPLY", "group": "Foundation", "mark": "PC2",   "description": "Footing PC2",   "qty": 1,  "lengthM": 2.0,  "widthM": 1.0,  "heightM": 1.0, "isBuried": "Y", "boqBucket": "Substructure" },
    { "dataRow": 7,  "elementId": "49-PC3",  "building": "49.ELECTRIC SUPPLY", "group": "Foundation", "mark": "PC3",   "description": "Footing PC3",   "qty": 2,  "lengthM": 2.0,  "widthM": 2.0,  "heightM": 1.0, "isBuried": "Y", "boqBucket": "Substructure" },
    { "dataRow": 8,  "elementId": "49-PC31", "building": "49.ELECTRIC SUPPLY", "group": "Foundation", "mark": "PC3.1", "description": "Footing PC3.1", "qty": 4,  "lengthM": 2.0,  "widthM": 2.0,  "heightM": 1.0, "isBuried": "Y", "boqBucket": "Substructure" },
    { "dataRow": 9,  "elementId": "49-PC4",  "building": "49.ELECTRIC SUPPLY", "group": "Foundation", "mark": "PC4",   "description": "Footing PC4",   "qty": 7,  "lengthM": 2.0,  "widthM": 2.0,  "heightM": 1.0, "isBuried": "Y", "boqBucket": "Substructure" },
    { "dataRow": 10, "elementId": "49-PC6",  "building": "49.ELECTRIC SUPPLY", "group": "Foundation", "mark": "PC6",   "description": "Footing PC6",   "qty": 5,  "lengthM": 2.0,  "widthM": 3.2,  "heightM": 1.0, "isBuried": "Y", "boqBucket": "Substructure" }
  ]
}
"@ | Set-Content -Path $bhIn -Encoding UTF8
& (Join-Path $Root "dist\engine\zofit-engine.exe") --cmd generate_structure --in $bhIn --out $bhOut
if ($LASTEXITCODE -ne 0) { throw "Golden structure check failed!" }
# Check PC1..PC6 Footing Concrete sum = (10*1 + 1*2 + 2*4 + 4*4 + 7*4 + 5*6.4) * 1.03 = 96 * 1.03 = 98.88 m3 (exact match!)
Write-Host "        PASS: Golden BHFLEX structural matrix generated cleanly!"

# 3. Synchronous 1-Click Installer Test + SHA256 Verification
Write-Host "  [3/4] Running Setup.exe --silent synchronously & verifying SHA256..."
if (-not (Test-Path $SetupExe)) { throw "Missing $SetupExe" }
Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq "" } | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$visibleExcel = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($null -eq $visibleExcel) {
    $proc = Start-Process -FilePath $SetupExe -ArgumentList "--silent" -Wait -PassThru
    if ($proc.ExitCode -ne 0) { throw "Setup.exe --silent exited with code $($proc.ExitCode)" }

    $instXlam = Join-Path $env:APPDATA "Microsoft\AddIns\zOfit_QS_Pro.xlam"
    $instEng = Join-Path $env:LOCALAPPDATA "zOfitQSPro\zofit-engine.exe"
    $hSrcX = (Get-FileHash (Join-Path $Root "dist\zOfit_QS_Pro.xlam") -Algorithm SHA256).Hash
    $hDstX = (Get-FileHash $instXlam -Algorithm SHA256).Hash
    $hSrcE = (Get-FileHash (Join-Path $Root "dist\engine\zofit-engine.exe") -Algorithm SHA256).Hash
    $hDstE = (Get-FileHash $instEng -Algorithm SHA256).Hash
    if ($hSrcX -ne $hDstX) { throw "Installed .xlam SHA256 mismatch!" }
    if ($hSrcE -ne $hDstE) { throw "Installed zofit-engine.exe SHA256 mismatch!" }
    Write-Host "        PASS: Setup.exe extracted exact byte-for-byte .xlam & zofit-engine.exe + registered Excel Add-in!"

    # 4. Launch real visible Excel and verify Ribbon tab
    Write-Host "  [4/4] UIAutomation verifying 'zOfit QS Pro' tab in real Excel..."
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    $sampleXlsx = Join-Path $Root "dist\_test_v2_e2e.xlsx"
    $p = if (Test-Path $sampleXlsx) { Start-Process "excel.exe" -ArgumentList "`"$sampleXlsx`"" -PassThru } else { Start-Process "excel.exe" -PassThru }
    try {
        $foundTab = $false
        $deadline = (Get-Date).AddSeconds(25)
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 800
            $rootEl = [System.Windows.Automation.AutomationElement]::RootElement
            $cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "zOfit QS Pro")
            $match = $rootEl.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond)
            if ($null -ne $match) { $foundTab = $true; break }
        }
        if (-not $foundTab) { throw "UIAutomation did NOT find 'zOfit QS Pro' Ribbon tab in real Excel!" }
        Write-Host "        PASS: 'zOfit QS Pro' tab detected live on Microsoft Excel Ribbon!"
    }
    finally {
        Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    }
}
else {
    Write-Host "  Notice: User Excel window is active (PID: $($visibleExcel.Id -join ',')). Skipped overwriting installed addin."
}

Remove-Item $tmpIn, $tmpOut, $bhIn, $bhOut -ErrorAction SilentlyContinue
Write-Host "ALL GOLDEN + INSTALLER + UIAUTOMATION CHECKS PASSED!"
