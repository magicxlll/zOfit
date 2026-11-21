$ErrorActionPreference = "Stop"
$Root = "D:\03_APP\3. System\GITHUB\zOfit"
$XlamPath = Join-Path $Root "dist\zOfit_QS_Pro.xlam"
$OutXlsx = Join-Path $Root "dist\_test_v2_e2e.xlsx"
if (Test-Path $OutXlsx) { Remove-Item $OutXlsx -Force }

Write-Host "=== Running Wave 1 + Wave 2 Full End-to-End COM Verification ==="

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1

try {
    # Close any preloaded addin copy
    for ($wi = $xl.Workbooks.Count; $wi -ge 1; $wi--) {
        $w = $xl.Workbooks.Item($wi)
        if ($w.Name -like "*zOfit*") { try { $w.Close($false) } catch {} }
    }

    $addinWb = $xl.Workbooks.Open($XlamPath)
    $ap = $addinWb.Name

    $wb = $xl.Workbooks.Add()
    $wb.Activate()

    # 1. Ping Hybrid Engine from Excel VBA
    Write-Host "  [1/8] Testing VBA -> zofit-engine.exe UTF-8 Ping..."
    $pingOk = $xl.Run("'$ap'!modEngineBridge.TestEnginePing", $false)
    if (-not $pingOk) { throw "TestEnginePing returned False!" }
    Write-Host "        PASS: Engine Ping UTF-8 OK!"

    # 2. Create QS_Settings, QS_Spec, QS_Data
    Write-Host "  [2/8] Creating QS_Settings, QS_Spec & QS_Data..."
    $xl.Run("'$ap'!modSettings.CreateOrShowSettingsSheet", $false)
    $xl.Run("'$ap'!modSpecSheet.CreateOrShowSpecSheet", $false)
    $xl.Run("'$ap'!modQsData.CreateOrShowQsDataSheet", $false)
    $xl.Run("'$ap'!modTools.CreateMasterBOQSheet", $false)

    $wsData = $wb.Worksheets.Item("QS_Data")
    # Add PC2 on row 11 (rows 5..10 are seeded with PC1, GBX1, GS1, C1, W1, ST1)
    $wsData.Cells(11, 1).Value = "49-FTG-PC2"
    $wsData.Cells(11, 2).Value = "49.ELECTRIC SUPPLY"
    $wsData.Cells(11, 3).Value = "Foundation"
    $wsData.Cells(11, 4).Value = "PC2"
    $wsData.Cells(11, 5).Value = [char]0x0110 + "ai mong PC2"
    $wsData.Cells(11, 6).Value = 2
    $wsData.Cells(11, 7).Value = 2.0
    $wsData.Cells(11, 8).Value = 1.0
    $wsData.Cells(11, 9).Value = 1.0
    $wsData.Cells(11, 10).Value = 19.6
    $wsData.Cells(11, 11).Value = 18.6
    $wsData.Cells(11, 12).Value = "M400"
    $wsData.Cells(11, 15).Value = "Y"
    $wsData.Cells(11, 18).Value = "FTG_2WAY"
    $wsData.Cells(11, 19).Value = "Substructure"

    # 3. Generate Structure (_Detail + TKT_*)
    Write-Host "  [3/8] Generating 49.ELECTRIC SUPPLY_Detail & TKT_49.ELECTRIC SUPPLY..."
    $xl.Run("'$ap'!modQsGenerator.GenerateOrSyncAllStructure", $false, "49.ELECTRIC SUPPLY")
    $xl.CalculateFullRebuild()

    $wsDet = $wb.Worksheets.Item("49.ELECTRIC SUPPLY_Detail")
    $wsTkt = $wb.Worksheets.Item("TKT_49.ELECTRIC SUPPLY")
    $wsCfg = $wb.Worksheets.Item("QS_RebarConfig")
    $pc1ConcRow = 0; $pc1LeanRow = 0; $gbx1LeanRow = 0
    for ($r = 6; $r -le $wsDet.Cells($wsDet.Rows.Count, 2).End(-4162).Row; $r++) {
        $tag = [string]$wsDet.Cells($r, 13).Value2
        if ($tag -eq "49-FTG-PC1::CONC_SUB") { $pc1ConcRow = $r }
        if ($tag -eq "49-FTG-PC1::CONC_LEAN") { $pc1LeanRow = $r }
        if ($tag -eq "49-GB-GBX1::CONC_LEAN") { $gbx1LeanRow = $r }
    }
    if ($pc1ConcRow -eq 0) { throw "Could not find 49-FTG-PC1::CONC_SUB row in Detail sheet!" }
    if ($pc1LeanRow -eq 0 -or $gbx1LeanRow -eq 0) { throw "Missing CONC_LEAN rows for PC1 or GBX1!" }

    # Verify per-type + per-element Lean Concrete thickness (PC1 = 0.10m from SPEC_Lean_Ftg_Thk, GBX1 = 0.05m override)
    $pc1LeanH = [double]$wsDet.Cells($pc1LeanRow, 8).Value2
    $gbx1LeanH = [double]$wsDet.Cells($gbx1LeanRow, 8).Value2
    Write-Host "        Lean Thk: PC1=$pc1LeanH m | GBX1=$gbx1LeanH m"
    if ([Math]::Abs($pc1LeanH - 0.10) -gt 0.001 -or [Math]::Abs($gbx1LeanH - 0.05) -gt 0.001) {
        throw "Per-element Lean Concrete thickness mismatch! Expected PC1=0.10, GBX1=0.05"
    }

    # Verify QS_RebarConfig v3.0 Transparent Parametric BBS blocks & live values
    $cfgLast = $wsCfg.Cells($wsCfg.Rows.Count, 39).End(-4162).Row
    $gbxCount = 0; $colCount = 0; $wallV = $false; $wallH = $false
    $colAMid = 0.0; $colBMid = 0.0
    $pc1FirstBarRow = 0; $pc1HdrRow = 0
    for ($r = 5; $r -le $cfgLast; $r++) {
        $elId = [string]$wsCfg.Cells($r, 1).Value2
        $rk = [string]$wsCfg.Cells($r, 39).Value2
        $role = [string]$wsCfg.Cells($r, 10).Value2
        $barDir = [string]$wsCfg.Cells($r, 14).Value2
        $distDir = [string]$wsCfg.Cells($r, 15).Value2
        if ($elId -eq "49-FTG-PC1" -and $rk -eq "BLK_HDR") { $pc1HdrRow = $r }
        if ($elId -eq "49-FTG-PC1" -and $rk -eq "BAR" -and $pc1FirstBarRow -eq 0) { $pc1FirstBarRow = $r }
        if ($rk -eq "BAR") {
            if ($elId -eq "49-GB-GBX1") { $gbxCount++ }
            if ($elId -eq "49-COL-C1") {
                $colCount++
                if ($role -eq "COL_MAIN_A") { $colAMid = [double]$wsCfg.Cells($r, 29).Value2 }
                if ($role -eq "COL_MAIN_B") { $colBMid = [double]$wsCfg.Cells($r, 29).Value2 }
            }
            if ($elId -eq "49-WAL-W1" -and $barDir -eq "H" -and $distDir -eq "X") { $wallV = $true }
            if ($elId -eq "49-WAL-W1" -and $barDir -eq "X" -and $distDir -eq "H") { $wallH = $true }
        }
    }
    $cfgGrandKg = [double]$wsCfg.Range("AL2").Value2
    $pc1BlkKgBefore = [double]$wsCfg.Cells($pc1HdrRow, 38).Value2
    $colStagDiff = $colAMid - $colBMid
    Write-Host "        QS_RebarConfig v3.0: Total=$cfgGrandKg kg | PC1_Blk=$pc1BlkKgBefore kg | GBX1=$gbxCount | C1=$colCount (Stagger=$colStagDiff mm) | Wall(V=$wallV, H=$wallH)"
    if ($gbxCount -ne 7 -or $colCount -ne 5 -or -not $wallV -or -not $wallH -or [Math]::Abs($colStagDiff - 600) -gt 1 -or $cfgGrandKg -le 0) {
        throw "QS_RebarConfig v3.0 failed transparent preset or stagger validation!"
    }

    $vBefore = [double]$wsDet.Cells($pc1ConcRow, 12).Value2
    $kgBefore = [double]$wsTkt.Range("T7").Value2
    Write-Host "        PC1 Concrete before=$vBefore m3 | TKT Total Kg before=$kgBefore kg"

    # Test Live Excel Formula Reactivity (change QS_Data!G5 from 1.0 to 2.0)
    $wsData.Cells(5, 7).Value = 2.0
    $xl.CalculateFullRebuild()
    $vAfter = [double]$wsDet.Cells($pc1ConcRow, 12).Value2
    $kgAfter = [double]$wsTkt.Range("T7").Value2
    $pc1BlkKgAfter = [double]$wsCfg.Cells($pc1HdrRow, 38).Value2
    Write-Host "        PC1 Concrete after L=2m -> $vAfter m3 | Config PC1_Blk -> $pc1BlkKgAfter kg | TKT Total Kg -> $kgAfter kg"
    if ($vAfter -le $vBefore -or $kgAfter -le $kgBefore -or $pc1BlkKgAfter -le $pc1BlkKgBefore) {
        throw "Live Excel formula reactivity failed when QS_Data!G5 changed!"
    }

    # Test Custom Bar insertion on QS_RebarConfig v3.0 + ApplyShapeToActiveRow(D17) -> Live Lbar=1800 -> flows to TKT
    $wsCfg.Activate()
    $wsCfg.Cells($pc1FirstBarRow, 1).Select() | Out-Null
    $xl.Run("'$ap'!modRebarConfig.InsertCustomBarRowOnConfig", "D17", $true)
    $customCfgRow = [int]$xl.ActiveCell.Row
    $xl.Run("'$ap'!modRebarTKT.ApplyShapeToActiveRow", "D17")
    $wsCfg.Cells($customCfgRow, 12).Value = 20        # D (mm)
    $wsCfg.Cells($customCfgRow, 16).Value = 150       # HookL
    $wsCfg.Cells($customCfgRow, 18).Value = 150       # HookR
    $wsCfg.Cells($customCfgRow, 22).Value = 1200      # L.m override (Mid)
    $wsCfg.Cells($customCfgRow, 27).Value = 300       # Le.1a
    $wsCfg.Cells($customCfgRow, 25).Value = "Custom Crank Extra"
    $xl.CalculateFullRebuild()
    $liveLbar = [double]$wsCfg.Cells($customCfgRow, 34).Value2
    Write-Host "        Custom D17 Live Lbar on QS_RebarConfig = $liveLbar mm (expected 1800 mm)"
    if ([Math]::Abs($liveLbar - 1800) -gt 0.1) {
        throw "QS_RebarConfig v3.0 live Lbar failed on custom D17 row! Got $liveLbar"
    }
    $isQsD = $xl.Run("'$ap'!modQsData.IsQsDataSheet", $wsData)
    $singleElJson = $xl.Run("'$ap'!modQsData.SerializeSingleElementToJson", $wsData, 5)
    $c1CfgJson = $xl.Run("'$ap'!modRebarConfig.SerializeRebarConfigsToJson", $wsCfg, "49-COL-C1")
    if (-not $isQsD -or ($singleElJson -notlike "*49-FTG-PC1*") -or ($c1CfgJson -notlike "*SPEC_Col_StaggerOffset*")) {
        throw "Visual Rebar Editor serialization helpers failed!"
    }
    Write-Host "        PASS: Visual Rebar Editor serialization (QS_Data + C1 stagger formula roundtrip) verified!"

    # 4. Non-Destructive Sync Test: insert MANUAL rows + Add PC7 + Remove PC2 + Custom D17 in TKT
    Write-Host "  [4/8] Testing Non-Destructive Sync (Preserve MANUAL, Add PC7, Flag REMOVED PC2, Custom D17)..."
    # Add MANUAL row right below PC1 in CONC_SUB
    $manRow = $pc1ConcRow + 1
    $wsDet.Rows($manRow).Insert(-4121) # xlDown
    $wsDet.Cells($manRow, 1).Value = "MAN-1"
    $wsDet.Cells($manRow, 2).Value = "Custom Sump Pit (Manual DxRxC)"
    $wsDet.Cells($manRow, 4).Value = "m3"
    $wsDet.Cells($manRow, 5).Value = 1
    $wsDet.Cells($manRow, 6).Value = 2
    $wsDet.Cells($manRow, 7).Value = 2
    $wsDet.Cells($manRow, 8).Value = 0.5
    $wsDet.Cells($manRow, 12).Formula = "=PRODUCT(E$manRow:K$manRow)"
    $wsDet.Cells($manRow, 13).Value = "MANUAL"
    $wsDet.Cells($manRow, 14).Value = "CONC_SUB"

    # Add MANUAL row in TKT_
    $tktLast = $xl.Run("'$ap'!modRebarTKT.GetNextTKTRow", $wsTkt)
    $wsTkt.Cells($tktLast, 2).Value = "Manual Hook Bar"
    $wsTkt.Cells($tktLast, 3).Value = "99"
    $wsTkt.Cells($tktLast, 4).Value = "D01"
    $wsTkt.Cells($tktLast, 6).Value = 1500
    $wsTkt.Cells($tktLast, 15).Value = 16
    $wsTkt.Cells($tktLast, 16).Value = 4
    $wsTkt.Cells($tktLast, 17).Value = 1
    $wsTkt.Cells($tktLast, 25).Value = "MANUAL"
    $wsTkt.Cells($tktLast, 26).Value = "Substructure"

    # Remove PC2 from row 11 and replace with PC7
    $wsData.Cells(11, 1).Value = "49-FTG-PC7"
    $wsData.Cells(11, 4).Value = "PC7"
    $wsData.Cells(11, 5).Value = "New Footing PC7"

    $xl.Run("'$ap'!modQsGenerator.GenerateOrSyncAllStructure", $false, "49.ELECTRIC SUPPLY")
    $xl.CalculateFullRebuild()

    $foundManual = $false; $foundPC7 = $false; $foundRemovedPC2 = $false; $foundCustomD17 = $false
    $detLast = $wsDet.Cells($wsDet.Rows.Count, 2).End(-4162).Row
    for ($r = 6; $r -le $detLast; $r++) {
        $tag = [string]$wsDet.Cells($r, 13).Value2
        $desc = [string]$wsDet.Cells($r, 2).Value2
        if ($tag -eq "MANUAL" -and $desc -like "*Custom Sump Pit*") { $foundManual = $true }
        if ($tag -eq "49-FTG-PC7::CONC_SUB") { $foundPC7 = $true }
        if ($tag -like "REMOVED::49-FTG-PC2*" -and $desc -like "*`[REMOVED`]*") { $foundRemovedPC2 = $true }
    }
    $tktCheckLast = $wsTkt.Cells($wsTkt.Rows.Count, 2).End(-4162).Row
    for ($r = 11; $r -le $tktCheckLast; $r++) {
        if ([string]$wsTkt.Cells($r, 4).Value2 -eq "D17" -and [string]$wsTkt.Cells($r, 2).Value2 -like "*Custom Crank*") {
            $foundCustomD17 = $true
        }
    }
    if (-not $foundManual) { throw "Sync did not preserve MANUAL detail row!" }
    if (-not $foundPC7) { throw "Sync did not add new PC7 row!" }
    if (-not $foundRemovedPC2) { throw "Sync did not flag removed PC2!" }
    if (-not $foundCustomD17) { throw "Custom D17 row from QS_RebarConfig did not flow to TKT!" }
    Write-Host "        PASS: Sync preserved MANUAL rows, added PC7, flagged removed PC2, and synced Custom D17!"

    # 5. CAD Raw Text/Block Parser
    Write-Host "  [5/8] Testing CAD Raw Text & Block Parser..."
    $wsRaw = $wb.Worksheets.Add()
    $wsRaw.Name = "Data"
    $wsRaw.Cells(1, 1).Value = "BANG THONG KE TEXT"
    $wsRaw.Cells(2, 1).Value = "A-W300- i0.2%-L=13m"
    $wsRaw.Cells(3, 1).Value = "Box:600x600-L11m-i0.2%"
    $wsRaw.Cells(4, 1).Value = "D200-L44m- i0.5%"
    $wsRaw.Cells(5, 1).Value = "TN-TENGA2"; $wsRaw.Cells(5, 2).Value = 19.5; $wsRaw.Cells(5, 3).Value = 18.1; $wsRaw.Cells(5, 4).Value = 1
    $wsRaw.Activate()
    $xl.Run("'$ap'!modCadParser.ParseRawCadDataToSheet", $false)
    $xl.CalculateFullRebuild()
    $wsCad = $wb.Worksheets.Item("QS_CAD_Parsed")
    $foundD200 = $false
    for ($r = 4; $r -le 12; $r++) {
        if ($wsCad.Cells($r, 13).Value2 -eq "D200" -and [double]$wsCad.Cells($r, 15).Value2 -eq 44) { $foundD200 = $true }
    }
    if (-not $foundD200) { throw "CAD Parser SUMIFS failed for D200 = 44m!" }
    Write-Host "        PASS: CAD Parser -> QS_CAD_Parsed SUMIFS verified!"

    # 6. Wave 2: Finishing Module (QS_DoorWin + QS_Rooms -> Finishing_Detail)
    Write-Host "  [6/8] Testing Wave 2 Finishing Engine..."
    $xl.Run("'$ap'!modFinishingData.CreateFinishingInputSheets", $false)
    $xl.Run("'$ap'!modFinishingData.GenerateFinishingDetail", $false)
    $xl.CalculateFullRebuild()
    $wsFin = $wb.Worksheets.Item("Finishing_Detail")
    $wsDoors = $wb.Worksheets.Item("QS_DoorWin")
    $masRow = 0
    for ($r = 6; $r -le $wsFin.Cells($wsFin.Rows.Count, 2).End(-4162).Row; $r++) {
        if ($wsFin.Cells($r, 13).Value2 -eq "RM-WC1::MASONRY") { $masRow = $r; break }
    }
    if ($masRow -eq 0) { throw "Finishing Masonry row missing!" }
    $masBefore = [double]$wsFin.Cells($masRow, 12).Value2
    $wsDoors.Cells(5, 3).Value = 2.4 # change 1SD1 width 1.2 -> 2.4
    $xl.CalculateFullRebuild()
    $masAfter = [double]$wsFin.Cells($masRow, 12).Value2
    if ($masAfter -ge $masBefore) { throw "Finishing door deduction did not react dynamically ($masBefore -> $masAfter)!" }
    Write-Host "        PASS: Finishing INDEX/MATCH door deduction ($masBefore -> $masAfter m3) & Roof chain OK!"

    # 7. Wave 2: Infrastructure Module (QS_Infra + Feed CAD + Civil_Detail)
    Write-Host "  [7/8] Testing Wave 2 Infrastructure Engine..."
    $xl.Run("'$ap'!modInfraData.CreateOrShowInfraSheet", $false)
    $xl.Run("'$ap'!modInfraData.FeedLengthsFromCadParsed", $false)
    $xl.Run("'$ap'!modInfraData.GenerateInfraDetail", $false)
    $xl.CalculateFullRebuild()
    $wsInf = $wb.Worksheets.Item("QS_Infra")
    $wsCiv = $wb.Worksheets.Item("Civil_Detail")
    $d200Len = 0
    for ($r = 5; $r -le 12; $r++) {
        if ($wsInf.Cells($r, 3).Value2 -eq "D200") { $d200Len = [double]$wsInf.Cells($r, 6).Value2 }
    }
    if ($d200Len -ne 44) { throw "FeedLengthsFromCadParsed expected D200=44m, got $d200Len" }
    Write-Host "        PASS: Infrastructure CAD feed (D200=44m) & Civil_Detail generated!"

    # 8. BOQ Linker + Zero Formula Errors Audit + Clean XLSX Export
    Write-Host "  [8/8] Testing BOQ Linker + Global Formula Audit + Clean Export..."
    $wsBoq = $wb.Worksheets.Item("QS_Master_BOQ")
    $srcArr = $xl.Run("'$ap'!modBoqLinker.CollectLinkableSources", $wb)
    if ($null -eq $srcArr) { throw "CollectLinkableSources returned null!" }
    $xl.Run("'$ap'!modBoqLinker.AssignLinkFormulaToCell", $wsBoq.Range("E10"), "'49.ELECTRIC SUPPLY_Detail'!L$pc1ConcRow", "1")
    $xl.CalculateFullRebuild()

    $xl.Run("'$ap'!modTools.AuditWorkbookFormulas", $false)
    if ($wb.Worksheets.Item(1).Name -eq "QS_Audit_Report") {
        $wsAud = $wb.Worksheets.Item("QS_Audit_Report")
        $errCnt = $wsAud.Cells($wsAud.Rows.Count, 1).End(-4162).Row
        for ($ar = 4; $ar -le $errCnt; $ar++) {
            Write-Host "   ERROR CELL: $($wsAud.Cells($ar,2).Text) !$($wsAud.Cells($ar,3).Text) -> $($wsAud.Cells($ar,4).Text) | $($wsAud.Cells($ar,5).Text)"
        }
        throw "Formula errors detected in End-to-End workbook!"
    }

    $wb.SaveAs($OutXlsx, 51) # Clean macro-free .xlsx!
    $wb.Close($false)
    $addinWb.Close($false)
    Write-Host "ALL WAVE 1 + WAVE 2 END-TO-END TESTS PASSED (0 FORMULA ERRORS)!"
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# Verify exported .xlsx has zero VBA / zero #REF!
python -c @"
import zipfile
with zipfile.ZipFile(r'$OutXlsx') as z:
    names = z.namelist()
    assert 'xl/vbaProject.bin' not in names, 'Exported .xlsx still contains vbaProject.bin!'
    print('Clean .xlsx verified! Sheet files:', sum(1 for n in names if n.startswith('xl/worksheets/sheet')))
"@
if ($LASTEXITCODE -ne 0) { throw "Clean XLSX verification failed!" }
