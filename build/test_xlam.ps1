$ErrorActionPreference = "Stop"
$Root = "D:\03_APP\3. System\GITHUB\zOfit"
$XlamPath = Join-Path $Root "dist\zOfit_QS_Pro.xlam"
$TestOut = Join-Path $Root "dist\_test_output_p2.xlsx"
$TestCleanBoq = Join-Path $Root "dist\_test_cleaned_boq.xlsx"

if (Test-Path $TestOut) { Remove-Item $TestOut -Force }
if (Test-Path $TestCleanBoq) { Remove-Item $TestCleanBoq -Force }

# 1. Verify OOXML structure & Ribbon XML inside .xlam
Write-Host "=== Step 1: Verify OOXML inside .xlam ==="
python -c @"
import zipfile, xml.etree.ElementTree as ET
with zipfile.ZipFile(r'$XlamPath') as z:
    names = z.namelist()
    assert 'customUI/customUI14.xml' in names, 'Missing customUI14.xml'
    assert 'xl/vbaProject.bin' in names, 'Missing vbaProject.bin'
    ET.fromstring(z.read('customUI/customUI14.xml'))
    rels = z.read('_rels/.rels').decode('utf-8')
    assert 'customUI/customUI14.xml' in rels, 'Missing rels entry for customUI14.xml'
    print('OOXML check PASSED! Entries:', len(names))
"@
if ($LASTEXITCODE -ne 0) { throw "OOXML verification failed" }

# 2. Functional COM testing (Phase 1 + Phase 2)
Write-Host "=== Step 2: Functional COM testing (Phase 1 & 2) ==="
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1 # msoAutomationSecurityLow

try {
    for ($wi = $xl.Workbooks.Count; $wi -ge 1; $wi--) {
        $w = $xl.Workbooks.Item($wi)
        if ($w.Name -like "*zOfit*") { try { $w.Close($false) } catch {} }
    }
    $addinWb = $xl.Workbooks.Open($XlamPath)
    $addinProj = $addinWb.Name

    $wb = $xl.Workbooks.Add()
    $wb.Activate()

    # Phase 1 tests
    Write-Host "  [P1] CreateOrShowSettingsSheet & ExportCatalog & MasterBOQ..."
    $xl.Run("'$addinProj'!modSettings.CreateOrShowSettingsSheet", $false)
    $xl.Run("'$addinProj'!modCatalog.ExportCatalogToActiveBook", $false)
    $xl.Run("'$addinProj'!modTools.CreateMasterBOQSheet", $false)

    # Phase 2 tests: TKT sheets (41 shapes D01..D41)
    Write-Host "  [P2] Creating TKT_Mong & TKT_Dam..."
    $xl.Run("'$addinProj'!modRebarTKT.CreateTKTSheet", "TKT_Mong", "Mong / Dai coc", $false)
    $wsMong = $wb.Worksheets.Item("TKT_Mong")
    $wsMong.Activate()

    # Insert various shapes + test Coupler (D25, L=15000mm) and Overlap (D16, L=15000mm)
    $xl.Run("'$addinProj'!modRebarTKT.InsertBarRowFromUI", "D39", $false) # Closed stirrup
    $xl.Run("'$addinProj'!modRebarTKT.InsertBarRowFromUI", "D01", $false) # Straight
    $wsMong.Cells(13, 6).Value = 15000
    $wsMong.Cells(13, 15).Value = 25 # Col O = Dia >= CFG_CouplerMinD (22) -> Couplers
    $xl.Run("'$addinProj'!modRebarTKT.InsertBarRowFromUI", "D07", $false) # U-bar (L1=250, L2=2400, L3=250)
    $wsMong.Cells(14, 7).Value = 15000
    $wsMong.Cells(14, 15).Value = 16 # < 22 -> adds 40D overlap splice (640mm)

    $xl.Run("'$addinProj'!modRebarTKT.CreateTKTSheet", "TKT_Dam", "Dam", $false)
    $wsDam = $wb.Worksheets.Item("TKT_Dam")
    $wsDam.Activate()
    $xl.Run("'$addinProj'!modRebarTKT.InsertBarRowFromUI", "D41", $false) # Circle hoop
    $xl.Run("'$addinProj'!modRebarTKT.InsertBarRowFromUI", "D17", $false) # Bent-up 4 seg

    # Test row duplicate, move up, delete, clear shapes (Light mode), redraw shapes
    $wsDam.Cells(12, 6).Select() | Out-Null
    $xl.Run("'$addinProj'!modRebarTKT.DuplicateBarRow")
    $xl.Run("'$addinProj'!modRebarTKT.MoveBarRowUp")
    $xl.Run("'$addinProj'!modRebarTKT.DeleteBarRow")
    $xl.Run("'$addinProj'!modRebarTKT.ClearActiveSheetShapes")
    if ($wsDam.Shapes.Count -ne 0) { throw "ClearActiveSheetShapes did not remove all shapes!" }
    $xl.Run("'$addinProj'!modRebarTKT.RedrawActiveSheetShapes", $false)
    if ($wsDam.Shapes.Count -lt 3) { throw "RedrawActiveSheetShapes failed!" }

    $xl.Calculate()

    $mongShapes = $wsMong.Shapes.Count
    $mongKg = $wsMong.Range("T7").Value2
    $mongCouplers = $wsMong.Range("R7").Value2
    $d16Len = $wsMong.Cells(14, 19).Value2
    Write-Host "  [P2] TKT_Mong: shapes=$mongShapes | totalKg=$mongKg | couplers=$mongCouplers | D16(15m+2*0.25+lap)=$d16Len m"
    if ($mongShapes -lt 4) { throw "Expected >= 4 shapes in TKT_Mong, got $mongShapes" }
    if ($mongCouplers -le 0) { throw "Expected Couplers > 0 for D25 L=15000mm" }
    if ($mongKg -le 0) { throw "Expected TKT_Mong total kg > 0" }

    # Phase 2 tests: Steel Shape & Plate
    Write-Host "  [P2] Creating TK_ThepHinh & TK_ThepTam..."
    $xl.Run("'$addinProj'!modSteelShape.CreateSteelShapeSheet", $false)
    $xl.Run("'$addinProj'!modSteelShape.CreateSteelPlateSheet", $false)
    $xl.Calculate()
    $shapeKg = $wb.Worksheets.Item("TK_ThepHinh").Range("L2").Value2
    $plateKg = $wb.Worksheets.Item("TK_ThepTam").Range("M2").Value2
    Write-Host "  [P2] TK_ThepHinh total=$shapeKg kg | TK_ThepTam total=$plateKg kg"
    if ($shapeKg -le 0 -or $plateKg -le 0) { throw "Steel shape or plate total <= 0" }

    # Phase 2 tests: Cross-sheet Summary & 1D Cutting Optimizer
    Write-Host "  [P2] Running SummarizeAllSteel & OptimizeBarCutting1D..."
    $xl.Run("'$addinProj'!modRebarSummary.SummarizeAllSteel", $false)
    $xl.Run("'$addinProj'!modRebarSummary.OptimizeBarCutting1D", $false)
    $xl.Calculate()

    $sumWs = $wb.Worksheets.Item("TH_Thep_ToanCT")
    $cutWs = $wb.Worksheets.Item("TKT_CatToiUu")
    Write-Host "  [P2] TH_Thep_ToanCT & TKT_CatToiUu created! Cut optimizer row4: $($cutWs.Cells(4,2).Text) -> $($cutWs.Cells(4,4).Text) bars, waste=$($cutWs.Cells(4,7).Text)"

    # Formula Audit across ALL sheets
    $xl.Run("'$addinProj'!modTools.AuditWorkbookFormulas", $false)
    if ($wb.Worksheets.Item(1).Name -eq "QS_Audit_Report") {
        throw "Formula errors detected in Phase 2 output workbook! Check QS_Audit_Report."
    }
    Write-Host "  [Audit] Zero formula errors across all sheets!"

    $wb.SaveAs($TestOut, 51)
    $wb.Close($false)
    $addinWb.Close($false)
    Write-Host "ALL PHASE 1 & PHASE 2 TESTS PASSED!"
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
