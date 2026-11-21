param(
    [string]$OutXlam = "D:\03_APP\3. System\GITHUB\zOfit\dist\zOfit_QS_Pro.xlam"
)

$ErrorActionPreference = "Stop"
$Root = "D:\03_APP\3. System\GITHUB\zOfit"
$DistDir = Split-Path $OutXlam -Parent
if (-not (Test-Path $DistDir)) { New-Item -ItemType Directory -Path $DistDir | Out-Null }
if (Test-Path $OutXlam) { Remove-Item $OutXlam -Force }
$TmpXlam = Join-Path $DistDir ("_build_" + [System.Guid]::NewGuid().ToString("N").Substring(0, 8) + ".xlam")

function Load-CsvToSheet {
    param(
        [object]$Worksheet,
        [string]$CsvPath
    )
    $rows = @(Import-Csv -Path $CsvPath -Encoding UTF8)
    if ($rows.Count -eq 0) { return }
    $headers = @($rows[0].PSObject.Properties.Name)
    $nRows = $rows.Count + 1
    $nCols = $headers.Count

    $arr = [object[,]]::new($nRows, $nCols)
    for ($c = 0; $c -lt $nCols; $c++) {
        $arr[0, $c] = [string]$headers[$c]
    }
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $rowObj = $rows[$r]
        $rr = $r + 1
        for ($c = 0; $c -lt $nCols; $c++) {
            $val = $rowObj.($headers[$c])
            if ($null -eq $val -or $val -eq "") {
                $arr[$rr, $c] = ""
            } elseif ($val -match '^-?\d+(\.\d+)?$' -and $headers[$c] -ne "Key") {
                $arr[$rr, $c] = [double]$val
            } else {
                $arr[$rr, $c] = [string]$val
            }
        }
    }
    $startCell = $Worksheet.Cells(1, 1)
    $endCell = $Worksheet.Cells($nRows, $nCols)
    $Worksheet.Range($startCell, $endCell).Value2 = $arr
}

Write-Host "Starting Excel COM to build $OutXlam..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

try {
    # Close any pre-loaded addin with same name to avoid filename collision
    for ($wi = $xl.Workbooks.Count; $wi -ge 1; $wi--) {
        $w = $xl.Workbooks.Item($wi)
        if ($w.Name -like "*zOfit*") { try { $w.Close($false) } catch {} }
    }
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) {
        $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
    }

    # 1. Create hidden data sheets
    $wsI18n = $wb.Worksheets.Item(1)
    $wsI18n.Name = "DB_I18N"
    Load-CsvToSheet -Worksheet $wsI18n -CsvPath (Join-Path $Root "src\data\i18n.csv")

    $wsRebar = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wsI18n)
    $wsRebar.Name = "DB_Rebar"
    Load-CsvToSheet -Worksheet $wsRebar -CsvPath (Join-Path $Root "src\data\catalog_rebar.csv")

    $wsMat = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wsRebar)
    $wsMat.Name = "DB_Materials"
    Load-CsvToSheet -Worksheet $wsMat -CsvPath (Join-Path $Root "src\data\catalog_materials.csv")

    $wsShapes = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wsMat)
    $wsShapes.Name = "DB_BarShapes"
    Load-CsvToSheet -Worksheet $wsShapes -CsvPath (Join-Path $Root "src\data\catalog_shapes.csv")

    $wsRebar.Visible = 2
    $wsMat.Visible = 2
    $wsShapes.Visible = 2

    # 2. Validate Pure-ASCII VBA files & Import VBA modules & classes
    $vbaFiles = Get-ChildItem (Join-Path $Root "src\vba") -File
    foreach ($vf in $vbaFiles) {
        $bytes = [System.IO.File]::ReadAllBytes($vf.FullName)
        foreach ($b in $bytes) {
            if ($b -gt 127) { throw "Non-ASCII byte ($b) detected in $($vf.Name)! All VBA files must be Pure ASCII." }
        }
    }

    $vbProj = $wb.VBProject
    $basFiles = @(
        "modUnicode.bas",
        "modRibbon.bas",
        "modSettings.bas",
        "modCatalog.bas",
        "modTools.bas",
        "modShapeDraw.bas",
        "modRebarTKT.bas",
        "modSteelShape.bas",
        "modRebarSummary.bas",
        "modEngineBridge.bas",
        "modSpecSheet.bas",
        "modQsData.bas",
        "modRebarConfig.bas",
        "modQsGenerator.bas",
        "modCadParser.bas",
        "modBoqLinker.bas",
        "modFinishingData.bas",
        "modInfraData.bas"
    )
    foreach ($f in $basFiles) {
        $fullPath = Join-Path $Root "src\vba\$f"
        Write-Host "  Importing $f..."
        [void]$vbProj.VBComponents.Import($fullPath)
    }

    # Attach ThisWorkbook event code (Context Menu for TKT_*)
    $twCode = Get-Content -Path (Join-Path $Root "src\vba\ThisWorkbook_code.txt") -Raw -Encoding UTF8
    $twComp = $vbProj.VBComponents.Item("ThisWorkbook")
    if ($twComp.CodeModule.CountOfLines -gt 0) { $twComp.CodeModule.DeleteLines(1, $twComp.CodeModule.CountOfLines) }
    $twComp.CodeModule.AddFromString($twCode)

    function Add-Ctrl($des, $type, $name, [int]$left, [int]$top, [int]$width, [int]$height, $caption = "") {
        $c = $des.Controls.Add($type, $name, $true)
        $c.Left = $left; $c.Top = $top; $c.Width = $width; $c.Height = $height
        if ($caption -ne "") { $c.Caption = $caption }
        return $c
    }

    # 3a. Build frmSteelLookup UserForm
    Write-Host "  Building frmSteelLookup UserForm..."
    $frmComp = $vbProj.VBComponents.Add(3)
    $frmComp.Name = "frmSteelLookup"
    $frmComp.Properties.Item("Caption").Value = "zOfit QS Pro - Material Lookup"
    $frmComp.Properties.Item("Width").Value = 560
    $frmComp.Properties.Item("Height").Value = 395
    $des = $frmComp.Designer

    [void](Add-Ctrl $des "Forms.Label.1" "lblCategory" 12 10 110 16 "Category:")
    [void](Add-Ctrl $des "Forms.ComboBox.1" "cboCategory" 12 26 165 20)
    [void](Add-Ctrl $des "Forms.Label.1" "lblSearch" 190 10 340 16 "Quick Search:")
    [void](Add-Ctrl $des "Forms.TextBox.1" "txtSearch" 190 26 350 20)
    [void](Add-Ctrl $des "Forms.ListBox.1" "lstResults" 12 54 528 215)
    [void](Add-Ctrl $des "Forms.Label.1" "lblStatus" 12 273 300 16 "0 items")
    [void](Add-Ctrl $des "Forms.Label.1" "lblCalc" 12 292 140 16 "Quick Calc:")
    [void](Add-Ctrl $des "Forms.Label.1" "lblLen" 12 310 120 16 "Len (m) / Area (m2):")
    [void](Add-Ctrl $des "Forms.TextBox.1" "txtLen" 135 308 60 18)
    [void](Add-Ctrl $des "Forms.Label.1" "lblQty" 208 310 60 16 "Qty:")
    [void](Add-Ctrl $des "Forms.TextBox.1" "txtQty" 270 308 50 18)
    [void](Add-Ctrl $des "Forms.Label.1" "lblResult" 332 310 208 18 "Total Weight: 0.00 kg")
    [void](Add-Ctrl $des "Forms.CommandButton.1" "btnInsert" 12 336 165 26 "Insert to Active Row")
    [void](Add-Ctrl $des "Forms.CommandButton.1" "btnExport" 188 336 175 26 "Export Catalog Sheet")
    [void](Add-Ctrl $des "Forms.CommandButton.1" "btnClose" 445 336 95 26 "Close")

    $codeText = Get-Content -Path (Join-Path $Root "src\vba\frmSteelLookup_code.txt") -Raw -Encoding UTF8
    if ($frmComp.CodeModule.CountOfLines -gt 0) { $frmComp.CodeModule.DeleteLines(1, $frmComp.CodeModule.CountOfLines) }
    $frmComp.CodeModule.AddFromString($codeText)
    try { $frmComp.DesignerWindow().Close() } catch {}

    # 3b. Build frmBarShape UserForm (41 BMP tiles 7x6 + Large Live Preview Panel)
    Write-Host "  Adding clsShapeBtn class..."
    $clsComp = $vbProj.VBComponents.Add(2)
    $clsComp.Name = "clsShapeBtn"
    $clsRaw = Get-Content -Path (Join-Path $Root "src\vba\clsShapeBtn.cls") -Raw -Encoding UTF8
    $clsIdx = $clsRaw.IndexOf("Option Explicit")
    $clsBody = if ($clsIdx -ge 0) { $clsRaw.Substring($clsIdx) } else { $clsRaw }
    $clsBody = $clsBody -replace "(?m)^Attribute\s+.*\r?\n", ""
    $clsComp.CodeModule.AddFromString($clsBody)
    Write-Host "  Building frmBarShape UserForm (41 rebar shapes + preview)..."
    $frmBar = $vbProj.VBComponents.Add(3)
    $frmBar.Name = "frmBarShape"
    $frmBar.Properties.Item("Caption").Value = "zOfit QS Pro - Rebar Shapes"
    $frmBar.Properties.Item("Width").Value = 762
    $frmBar.Properties.Item("Height").Value = 318
    $desBar = $frmBar.Designer

    [void](Add-Ctrl $desBar "Forms.Label.1" "lblGridTitle" 10 8 475 16 "1. Select Shape:")
    $bmpDir = Join-Path $Root "src\data\shapes_bmp"
    $cols = 7; $tileW = 64; $tileH = 34; $gapX = 4; $gapY = 4; $startX = 10; $startY = 26
    for ($i = 1; $i -le 41; $i++) {
        $idx = $i - 1
        $rIdx = [math]::Floor($idx / $cols)
        $cIdx = $idx % $cols
        $left = $startX + $cIdx * ($tileW + $gapX)
        $top = $startY + $rIdx * ($tileH + $gapY)
        $code = "D" + $i.ToString("00")
        $im = Add-Ctrl $desBar "Forms.Image.1" ("img" + $code) $left $top $tileW $tileH
        $im.PictureSizeMode = 3 # fmPictureSizeModeZoom
        $im.BackColor = 16777215
    }

    # Right Live Preview & Details Panel
    [void](Add-Ctrl $desBar "Forms.Label.1" "lblPrevTitle" 496 8 248 16 "2. Segment Preview:")
    $imPrev = Add-Ctrl $desBar "Forms.Image.1" "imgPreview" 496 26 248 92
    $imPrev.PictureSizeMode = 3 # fmPictureSizeModeZoom
    $imPrev.BorderStyle = 1
    $imPrev.BackColor = 16777215

    [void](Add-Ctrl $desBar "Forms.Label.1" "lblSelected" 496 124 248 34 "D01")
    [void](Add-Ctrl $desBar "Forms.Label.1" "lblParams" 496 162 248 18 "Params: L1")
    [void](Add-Ctrl $desBar "Forms.Label.1" "lblFormula" 496 182 248 30 "Formula: L1")
    [void](Add-Ctrl $desBar "Forms.CommandButton.1" "btnAddEnd" 496 218 120 26 "Add to End")
    [void](Add-Ctrl $desBar "Forms.CommandButton.1" "btnInsertHere" 624 218 120 26 "Insert at Cursor")
    [void](Add-Ctrl $desBar "Forms.CommandButton.1" "btnApplyRow" 496 250 156 24 "Apply to Active Row")
    [void](Add-Ctrl $desBar "Forms.CommandButton.1" "btnClose" 660 250 84 24 "Close")

    $barCode = Get-Content -Path (Join-Path $Root "src\vba\frmBarShape_code.txt") -Raw -Encoding UTF8
    if ($frmBar.CodeModule.CountOfLines -gt 0) { $frmBar.CodeModule.DeleteLines(1, $frmBar.CodeModule.CountOfLines) }
    $frmBar.CodeModule.AddFromString($barCode)
    try { $frmBar.DesignerWindow().Close() } catch {}

    # 3c. Build frmBoqLinker UserForm
    Write-Host "  Building frmBoqLinker UserForm..."
    $frmB = $vbProj.VBComponents.Add(3)
    $frmB.Name = "frmBoqLinker"
    $frmB.Properties.Item("Caption").Value = "zOfit QS Pro - BOQ Linker"
    $frmB.Properties.Item("Width").Value = 630
    $frmB.Properties.Item("Height").Value = 370
    $desB = $frmB.Designer

    [void](Add-Ctrl $desB "Forms.Label.1" "lblTarget" 12 10 170 16 "Target Cell:")
    [void](Add-Ctrl $desB "Forms.TextBox.1" "txtTarget" 185 8 280 20)
    [void](Add-Ctrl $desB "Forms.CommandButton.1" "btnRefresh" 475 8 135 20 "Get ActiveCell")
    [void](Add-Ctrl $desB "Forms.Label.1" "lblFilter" 12 36 110 16 "Filter:")
    [void](Add-Ctrl $desB "Forms.TextBox.1" "txtFilter" 125 34 340 20)
    [void](Add-Ctrl $desB "Forms.Label.1" "lblCount" 475 36 135 16 "0 items")
    [void](Add-Ctrl $desB "Forms.ListBox.1" "lstSources" 12 62 598 220)
    [void](Add-Ctrl $desB "Forms.Label.1" "lblMult" 12 292 170 16 "Multiplier:")
    [void](Add-Ctrl $desB "Forms.ComboBox.1" "cboMult" 185 290 120 20)
    [void](Add-Ctrl $desB "Forms.CommandButton.1" "btnLink" 330 288 175 26 "Assign Link")
    [void](Add-Ctrl $desB "Forms.CommandButton.1" "btnClose" 515 288 95 26 "Close")

    $boqLinkCode = Get-Content -Path (Join-Path $Root "src\vba\frmBoqLinker_code.txt") -Raw -Encoding UTF8
    if ($frmB.CodeModule.CountOfLines -gt 0) { $frmB.CodeModule.DeleteLines(1, $frmB.CodeModule.CountOfLines) }
    $frmB.CodeModule.AddFromString($boqLinkCode)
    try { $frmB.DesignerWindow().Close() } catch {}

    # Use a temporary VBA build helper to embed the 41 BMPs via native LoadPicture()
    Write-Host "  Embedding 41 shape BMPs via VBA LoadPicture..."
    $tmpMod = $vbProj.VBComponents.Add(1)
    $tmpMod.Name = "modBuildHelper"
    $helperCode = @'
Public Function EmbedShapePictures(ByVal bmpFolder As String) As String
    On Error GoTo ErrH
    Dim i As Long, code As String
    Dim frm As Object
    Set frm = ActiveWorkbook.VBProject.VBComponents("frmBarShape").Designer
    For i = 1 To 41
        code = "D" & Format$(i, "00")
        Set frm.Controls("img" & code).Picture = LoadPicture(bmpFolder & "\" & code & ".bmp")
    Next i
    Set frm.Controls("imgPreview").Picture = LoadPicture(bmpFolder & "\D01.bmp")
    EmbedShapePictures = "OK"
    Exit Function
ErrH:
    EmbedShapePictures = "ERR: " & Err.Number & " - " & Err.Description
End Function
'@
    $tmpMod.CodeModule.AddFromString($helperCode)
    $wb.Activate()
    $res = $xl.Run("'$($wb.Name)'!modBuildHelper.EmbedShapePictures", $bmpDir)
    Write-Host "  EmbedShapePictures result: $res"
    $vbProj.VBComponents.Remove($tmpMod)
    if ($res -ne "OK") { throw "EmbedShapePictures failed: $res" }

    # Convert workbook to Add-in and save as .xlam (xlOpenXMLAddIn = 55)
    $wb.IsAddin = $true
    $wb.SaveAs($TmpXlam, 55)
    $wb.Close($false)
    Write-Host "Saved raw XLAM temp: $TmpXlam"
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Move-Item -Path $TmpXlam -Destination $OutXlam -Force

# 4. Inject customUI14.xml Ribbon into .xlam
Write-Host "Injecting customUI14.xml Ribbon..."
python (Join-Path $Root "build\inject_ribbon.py") $OutXlam (Join-Path $Root "src\ribbon\customUI14.xml")
if ($LASTEXITCODE -ne 0) { throw "Ribbon injection failed with exit code $LASTEXITCODE" }
Write-Host "BUILD COMPLETE: $OutXlam"
