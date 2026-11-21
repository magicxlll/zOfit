Attribute VB_Name = "modSettings"
Option Explicit

' ==============================================================================
' modSettings - Parametric QS Measurement Rules, Named Ranges & Presets
' ==============================================================================

Private Const SETTINGS_SHEET As String = "QS_Settings"

Public Sub InitQSWorkspace()
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    CreateOrShowSettingsSheet False
    modSpecSheet.CreateOrShowSpecSheet False
    modQsData.CreateOrShowQsDataSheet False
    modCatalog.ExportCatalogToActiveBook False
    modTools.CreateMasterBOQSheet False

    ActiveWorkbook.Worksheets(modQsData.QSDATA_SHEET).Activate
    Application.ScreenUpdating = True
    modUnicode.MsgBoxU modUnicode.T("MSG_INIT_DONE", "VI"), vbInformation
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    modUnicode.MsgBoxU "InitQSWorkspace Error: " & Err.Description, vbCritical
End Sub

Public Sub CreateOrShowSettingsSheet(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim ws As Worksheet
    Dim isNew As Boolean

    On Error Resume Next
    Set ws = wb.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        ws.Name = SETTINGS_SHEET
        isNew = True
    End If

    If isNew Then
        BuildSettingsLayout ws
    Else
        RefreshSettingsHeaders ws
    End If

    ws.Activate
    If showMsg Then
        modUnicode.MsgBoxU modUnicode.T("MSG_SETTINGS_CREATED", "VI"), vbInformation
    End If
End Sub

Private Sub BuildSettingsLayout(ByVal ws As Worksheet)
    Application.ScreenUpdating = False
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10

    ' Title Banner
    With ws.Range("B2:F2")
        .Merge
        .Value = modUnicode.T("SET_TITLE")
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(22, 54, 92)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .RowHeight = 32
    End With

    With ws.Range("B3:F3")
        .Merge
        .Value = modUnicode.T("SET_SUBTITLE")
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = RGB(64, 64, 64)
        .Interior.Color = RGB(237, 242, 248)
        .IndentLevel = 1
        .RowHeight = 22
    End With

    ' Table Header
    Dim r As Long
    r = 5
    ws.Cells(r, 2).Value = modUnicode.T("SET_COL_PARAM")
    ws.Cells(r, 3).Value = modUnicode.T("SET_COL_CODE")
    ws.Cells(r, 4).Value = modUnicode.T("SET_COL_VAL")
    ws.Cells(r, 5).Value = modUnicode.T("SET_COL_UNIT")
    ws.Cells(r, 6).Value = modUnicode.T("SET_COL_NOTE")
    With ws.Range("B5:F5")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 85, 151)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 26
    End With

    r = 6
    ' Section A
    r = AddSectionBanner(ws, r, modUnicode.T("SET_SEC_A"))
    r = AddSettingRow(ws, r, modUnicode.U("T\u00ean d\u1ef1 \u00e1n / Project Name"), "CFG_ProjectName", modUnicode.U("D\u1ef1 \u00e1n M\u1eabu 2026"), "Text", modUnicode.U("Hi\u1ec3n th\u1ecb tr\u00ean ti\u00eau \u0111\u1ec1 t\u1ea5t c\u1ea3 b\u1ea3ng BOQ v\u00e0 TKT"))
    r = AddSettingRow(ws, r, modUnicode.U("Chu\u1ea9n \u0111o b\u00f3c / Standard Preset"), "CFG_Preset", "TT13/2021/TT-BXD (VN)", "TT13/2021/TT-BXD (VN),POMI / SMM7,Custom", modUnicode.U("B\u1ed9 quy t\u1eafc ti\u00eau chu\u1ea9n \u00e1p d\u1ee5ng cho d\u1ef1 \u00e1n"))
    r = AddSettingRow(ws, r, modUnicode.U("Ng\u00f4n ng\u1eef BOQ / Language Mode"), "CFG_Lang", "VI", "VI,EN,BI", modUnicode.U("VI = Ti\u1ebfng Vi\u1ec7t | EN = English | BI = Song ng\u1eef"))

    ' Section B
    r = AddSectionBanner(ws, r, modUnicode.T("SET_SEC_B"))
    r = AddSettingRow(ws, r, modUnicode.U("Quy t\u1eafc giao D\u1ea7m - S\u00e0n (B\u00ea t\u00f4ng)"), "CFG_BeamSlabRule", "1 - FloorFull_BeamMinusSlab", "1 - FloorFull_BeamMinusSlab,2 - BeamFull_FloorMinusBeam", modUnicode.U("1: T\u00ednh s\u00e0n su\u1ed1t, d\u1ea7m tr\u1eeb chi\u1ec1u d\u00e0y s\u00e0n (chu\u1ea9n VN) | 2: D\u1ea7m su\u1ed1t, s\u00e0n tr\u1eeb d\u1ea7m"))
    r = AddSettingRow(ws, r, modUnicode.U("Quy t\u1eafc giao C\u1ed9t - D\u1ea7m/S\u00e0n"), "CFG_ColBeamRule", "1 - ColThrough_BeamMinusCol", "1 - ColThrough_BeamMinusCol,2 - SlabBeamThrough_ColMinus", modUnicode.U("1: C\u1ed9t t\u00ednh su\u1ed1t chi\u1ec1u cao t\u1ea7ng, d\u1ea7m tr\u1eeb c\u1ed9t | 2: D\u1ea7m/s\u00e0n su\u1ed1t, c\u1ed9t tr\u1eeb d\u1ea7m"))
    r = AddSettingRow(ws, r, modUnicode.U("Tr\u1eeb v\u00e1n khu\u00f4n \u0111\u1ea7u d\u1ea7m giao c\u1ed9t/v\u00e1ch"), "CFG_FormDeduction", "Yes", "Yes,No", modUnicode.U("Yes: Tr\u1eeb ph\u1ea7n giao c\u1eaft | No: Kh\u00f4ng tr\u1eeb giao theo th\u1ecfa thu\u1eadn thi c\u00f4ng"))
    r = AddSettingRow(ws, r, modUnicode.U("Ng\u01b0\u1ee1ng l\u1ed7 m\u1edf kh\u00f4ng tr\u1eeb VK/Ho\u00e0n thi\u1ec7n"), "CFG_OpenThreshold", 0.5, "m2", modUnicode.U("L\u1ed7 m\u1edf nh\u1ecf h\u01a1n ng\u01b0\u1ee1ng n\u00e0y kh\u00f4ng tr\u1eeb kh\u1ed1i l\u01b0\u1ee3ng (TT13=0.5m2, POMI=1.0m2)"))

    ' Section C
    r = AddSectionBanner(ws, r, modUnicode.T("SET_SEC_C"))
    r = AddSettingRow(ws, r, modUnicode.U("Chi\u1ec1u d\u00e0i c\u00e2y th\u00e9p ti\u00eau chu\u1ea9n"), "CFG_BarStdLen", 11700, "11700,12000", modUnicode.U("Chi\u1ec1u d\u00e0i th\u01b0\u01a1ng m\u1ea1i 1 c\u00e2y th\u00e9p thanh (mm). Th\u00e9p <= D8 t\u00ednh cu\u1ed9n"))
    r = AddSettingRow(ws, r, modUnicode.U("H\u1ec7 s\u1ed1 n\u1ed1i ch\u1ed3ng m\u1eb7c \u0111\u1ecbnh (xD)"), "CFG_LapRatio", 40, "30,35,40,45", modUnicode.U("Chi\u1ec1u d\u00e0i \u0111o\u1ea1n n\u1ed1i ch\u1ed3ng = xD * \u0110\u01b0\u1eddng k\u00ednh thanh th\u00e9p"))
    r = AddSettingRow(ws, r, modUnicode.U("H\u1ec7 s\u1ed1 gi\u1ea3m u\u1ed1n m\u1ed7i g\u00f3c 90 \u0111\u1ed9 (xD)"), "CFG_BendDeduct", 1, "0,1,2", modUnicode.U("0D: T\u00ednh theo k\u00edch th\u01b0\u1edbc ph\u1ee7 b\u00ec | 1D/2D: Gi\u1ea3m tr\u1eeb \u0111o\u1ea1n d\u00e3n d\u00e0i khi u\u1ed1n"))
    r = AddSettingRow(ws, r, modUnicode.U("\u0110\u01b0\u1eddng k\u00ednh b\u1eaft \u0111\u1ea7u d\u00f9ng Coupler"), "CFG_CouplerMinD", 22, "20,22,25,28,32,99", modUnicode.U("T\u1eeb D n\u00e0y tr\u1edf l\u00ean t\u1ef1 chuy\u1ec3n n\u1ed1i ch\u1ed3ng sang Coupler (99 = kh\u00f4ng d\u00f9ng)"))
    r = AddSettingRow(ws, r, modUnicode.U("L\u1edbp b\u00ea t\u00f4ng b\u1ea3o v\u1ec7 M\u00f3ng / \u0110\u00e0i c\u1ecdc"), "CFG_CoverFdn", 50, "mm", modUnicode.U("L\u1edbp b\u1ea3o v\u1ec7 c\u1ed1t th\u00e9p c\u1ea5u ki\u1ec7n ti\u1ebfp x\u00fac \u0111\u1ea5t (mm)"))
    r = AddSettingRow(ws, r, modUnicode.U("L\u1edbp b\u00ea t\u00f4ng b\u1ea3o v\u1ec7 C\u1ed9t / D\u1ea7m"), "CFG_CoverColBeam", 30, "mm", modUnicode.U("L\u1edbp b\u00ea t\u00f4ng b\u1ea3o v\u1ec7 c\u1ed1t, d\u1ea7m m\u1eb7c \u0111\u1ecbnh (mm)"))
    r = AddSettingRow(ws, r, modUnicode.U("L\u1edbp b\u00ea t\u00f4ng b\u1ea3o v\u1ec7 S\u00e0n / V\u00e1ch"), "CFG_CoverSlabWall", 20, "mm", modUnicode.U("L\u1edbp b\u00ea t\u00f4ng b\u1ea3o v\u1ec7 s\u00e0n, v\u00e1ch, c\u1ea7u thang (mm)"))
    r = AddSettingRow(ws, r, modUnicode.U("H\u1ec7 s\u1ed1 hao h\u1ee5t c\u1ed1t th\u00e9p d\u1ef1 to\u00e1n"), "CFG_RebarWastePct", 2, "%", modUnicode.U("Ph\u1ea7n tr\u0103m hao h\u1ee5t c\u1eaft n\u1ed1i tham kh\u1ea3o (VD: 2%)"))

    ' Section D
    r = AddSectionBanner(ws, r, modUnicode.T("SET_SEC_D"))
    r = AddSettingRow(ws, r, modUnicode.U("H\u1ec7 s\u1ed1 m\u00e1i d\u1ed1c taluy h\u1ed1 \u0111\u00e0o (m)"), "CFG_SoilSlopeM", 0.5, "0,0.25,0.5,0.67,0.75,1", modUnicode.U("0 = \u0111\u00e0o th\u1eb3ng \u0111\u1ee9ng (c\u00f3 c\u1eeb); >0 = m\u00e1i d\u1ed1c taluy m theo TCVN 4447"))
    r = AddSettingRow(ws, r, modUnicode.U("H\u1ec7 s\u1ed1 t\u01a1i x\u1ed1p \u0111\u1ea5t v\u1eadn chuy\u1ec3n (K_toi)"), "CFG_SoilBulkingK", 1.15, "1.00 - 1.30", modUnicode.U("H\u1ec7 s\u1ed1 quy \u0111\u1ed5i \u0111\u1ea5t nguy\u00ean th\u1ed5 sang \u0111\u1ea5t r\u1eddi v\u1eadn chuy\u1ec3n \u0111\u1ed5 \u0111i"))
    r = AddSettingRow(ws, r, modUnicode.U("H\u1ec7 s\u1ed1 \u0111\u1ea7m ch\u1eb7t \u0111\u1ea5t \u0111\u1eafp (K_dam)"), "CFG_SoilCompactK", 1.1, "1.00 - 1.25", modUnicode.U("H\u1ec7 s\u1ed1 \u0111\u1ea5t \u0111\u1eafp \u0111\u1ea7m K95/K98 so v\u1edbi th\u1ec3 t\u00edch h\u00ecnh h\u1ecdc"))
    r = AddSettingRow(ws, r, modUnicode.U("Chi\u1ec1u d\u00e0y b\u00ea t\u00f4ng l\u00f3t m\u1eb7c \u0111\u1ecbnh"), "CFG_LeanThick", 100, "mm", modUnicode.U("B\u1ec1 d\u00e0y b\u00ea t\u00f4ng l\u00f3t \u0111\u00e1y m\u00f3ng / gi\u1eb1ng m\u00f3ng (mm)"))
    r = AddSettingRow(ws, r, modUnicode.U("\u0110\u1ed9 v\u01b0\u01a1n m\u00e9p b\u00ea t\u00f4ng l\u00f3t m\u1ed7i b\u00ean"), "CFG_LeanFlange", 100, "mm", modUnicode.U("Kho\u1ea3ng m\u1edf r\u1ed9ng b\u00ea t\u00f4ng l\u00f3t ra m\u1ed7i ph\u00eda c\u1ea5u ki\u1ec7n (mm)"))
    r = AddSettingRow(ws, r, modUnicode.U("Kho\u1ea3ng m\u1edf r\u1ed9ng thi c\u00f4ng h\u1ed1 \u0111\u00e0o (Offset)"), "CFG_ExcavOffset", 300, "mm", modUnicode.U("Kho\u1ea3ng thao t\u00e1c v\u00e1n khu\u00f4n t\u00ednh t\u1eeb m\u00e9p ngo\u00e0i b\u00ea t\u00f4ng l\u00f3t (mm)"))

    ' Column widths & borders
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 42
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 26
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 54

    With ws.Range("B5:F" & (r - 1)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191)
    End With
    ws.Range("B5:F" & (r - 1)).VerticalAlignment = xlCenter
    On Error Resume Next
    If Not Application.ActiveWindow Is Nothing Then Application.ActiveWindow.DisplayGridlines = True
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

Private Function AddSectionBanner(ByVal ws As Worksheet, ByVal rowIdx As Long, ByVal titleText As String) As Long
    With ws.Range(ws.Cells(rowIdx, 2), ws.Cells(rowIdx, 6))
        .Merge
        .Value = titleText
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 120)
        .Interior.Color = RGB(217, 225, 242)
        .IndentLevel = 1
        .RowHeight = 22
    End With
    AddSectionBanner = rowIdx + 1
End Function

Private Function AddSettingRow(ByVal ws As Worksheet, ByVal rowIdx As Long, _
                               ByVal paramLabel As String, ByVal codeName As String, _
                               ByVal defaultVal As Variant, ByVal unitOrList As String, _
                               ByVal noteText As String) As Long
    ws.Cells(rowIdx, 2).Value = paramLabel
    ws.Cells(rowIdx, 3).Value = codeName
    ws.Cells(rowIdx, 3).Font.Name = "Consolas"
    ws.Cells(rowIdx, 3).Font.Size = 9
    ws.Cells(rowIdx, 3).Font.Color = RGB(89, 89, 89)

    Dim valCell As Range
    Set valCell = ws.Cells(rowIdx, 4)
    valCell.Value = defaultVal
    valCell.Font.Bold = True
    valCell.Font.Color = RGB(0, 51, 153)
    valCell.Interior.Color = RGB(255, 250, 220)
    valCell.HorizontalAlignment = xlCenter

    ' Add dropdown validation if unitOrList contains comma-separated options
    If InStr(1, unitOrList, ",") > 0 Then
        With valCell.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=unitOrList
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        ws.Cells(rowIdx, 5).Value = "(List)"
    Else
        ws.Cells(rowIdx, 5).Value = unitOrList
    End If
    ws.Cells(rowIdx, 5).HorizontalAlignment = xlCenter

    ws.Cells(rowIdx, 6).Value = noteText
    ws.Cells(rowIdx, 6).WrapText = True
    ws.Rows(rowIdx).RowHeight = 22

    ' Register clean workbook-level Named Range pointing to this parameter cell
    RegisterWorkbookName ws.Parent, codeName, valCell

    AddSettingRow = rowIdx + 1
End Function

Public Sub RegisterWorkbookName(ByVal wb As Workbook, ByVal nmName As String, ByVal targetCell As Range)
    On Error Resume Next
    wb.Names(nmName).Delete
    On Error GoTo 0
    wb.Names.Add Name:=nmName, RefersTo:="='" & targetCell.Worksheet.Name & "'!" & targetCell.Address(True, True)
End Sub

Private Sub RefreshSettingsHeaders(ByVal ws As Worksheet)
    ws.Range("B2").Value = modUnicode.T("SET_TITLE")
    ws.Range("B3").Value = modUnicode.T("SET_SUBTITLE")
    ws.Range("B5").Value = modUnicode.T("SET_COL_PARAM")
    ws.Range("C5").Value = modUnicode.T("SET_COL_CODE")
    ws.Range("D5").Value = modUnicode.T("SET_COL_VAL")
    ws.Range("E5").Value = modUnicode.T("SET_COL_UNIT")
    ws.Range("F5").Value = modUnicode.T("SET_COL_NOTE")
End Sub

Public Sub ApplyMeasurementPreset(ByVal presetName As String)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    CreateOrShowSettingsSheet False

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    SetCfgValue wb, "CFG_Preset", presetName

    Select Case presetName
        Case "TT13/2021/TT-BXD (VN)"
            SetCfgValue wb, "CFG_BeamSlabRule", "1 - FloorFull_BeamMinusSlab"
            SetCfgValue wb, "CFG_ColBeamRule", "1 - ColThrough_BeamMinusCol"
            SetCfgValue wb, "CFG_FormDeduction", "Yes"
            SetCfgValue wb, "CFG_OpenThreshold", 0.5
            SetCfgValue wb, "CFG_BarStdLen", 11700
            SetCfgValue wb, "CFG_LapRatio", 40
            SetCfgValue wb, "CFG_BendDeduct", 1
            SetCfgValue wb, "CFG_SoilSlopeM", 0.5
        Case "POMI / SMM7"
            SetCfgValue wb, "CFG_BeamSlabRule", "1 - FloorFull_BeamMinusSlab"
            SetCfgValue wb, "CFG_ColBeamRule", "1 - ColThrough_BeamMinusCol"
            SetCfgValue wb, "CFG_FormDeduction", "Yes"
            SetCfgValue wb, "CFG_OpenThreshold", 1#
            SetCfgValue wb, "CFG_BarStdLen", 12000
            SetCfgValue wb, "CFG_LapRatio", 40
            SetCfgValue wb, "CFG_BendDeduct", 0
            SetCfgValue wb, "CFG_SoilSlopeM", 0
        Case Else
            ' Keep user values for Custom
    End Select

    modUnicode.MsgBoxU modUnicode.T("MSG_PRESET_APPLIED", "VI") & " " & presetName, vbInformation
End Sub

Public Sub SetBookLanguage(ByVal langCode As String)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    CreateOrShowSettingsSheet False

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    SetCfgValue wb, "CFG_Lang", UCase$(langCode)

    On Error Resume Next
    RefreshSettingsHeaders wb.Worksheets(SETTINGS_SHEET)
    modTools.RefreshMasterBOQHeaders wb
    modCatalog.RefreshCatalogHeaders wb
    On Error GoTo 0

    modUnicode.MsgBoxU modUnicode.T("MSG_LANG_CHANGED", "VI") & " " & UCase$(langCode), vbInformation
End Sub

Private Sub SetCfgValue(ByVal wb As Workbook, ByVal cfgName As String, ByVal newVal As Variant)
    On Error Resume Next
    wb.Names(cfgName).RefersToRange.Value = newVal
    On Error GoTo 0
End Sub
