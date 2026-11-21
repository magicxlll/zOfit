Attribute VB_Name = "modRebarConfig"
Option Explicit

' ==============================================================================
' modRebarConfig - v3.0 Transparent Parametric Rebar Input Sheet (QS_RebarConfig)
' Combines Single-Entry link from QS_Data + Live AAA_BOQ_CAL segment/qty control
' Columns A..H (Master Link) | I..Y (Rebar Inputs) | Z..AL (Live Results) | AM (RowKind)
' ==============================================================================

Public Const REBAR_CFG_SHEET As String = "QS_RebarConfig"
Public Const REBAR_CFG_FIRST_ROW As Long = 5
Public Const REBAR_CFG_SHAPE_COL As Long = 11   ' Col K = ShapeCode
Public Const REBAR_CFG_LAST_COL As Long = 39    ' Col AM = RowKind
Private Const REBAR_CFG_MARKER As String = "ZOFIT_RCFG_V3"

Public Function IsRebarConfigSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    If CStr(ws.Range("A1").Value) = REBAR_CFG_MARKER Or StrComp(ws.Name, REBAR_CFG_SHEET, vbTextCompare) = 0 Then
        If Trim$(CStr(ws.Cells(4, 1).Value)) = "ElementID" And Trim$(CStr(ws.Cells(4, REBAR_CFG_SHAPE_COL).Value)) = "Shape" Then
            IsRebarConfigSheet = True
        End If
    End If
    On Error GoTo 0
End Function

Public Sub CreateOrShowRebarConfigSheet(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook

    Dim hasData As Boolean
    On Error Resume Next
    hasData = Not (wb.Worksheets(modQsData.QSDATA_SHEET) Is Nothing)
    On Error GoTo 0
    If Not hasData Then modQsData.CreateOrShowQsDataSheet False

    EnsureRebarConfigSheet wb, True
    wb.Worksheets(REBAR_CFG_SHEET).Activate
    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 m\u1edf b\u1ea3ng Th\u00e9p T\u01b0\u1eddng Minh QS_RebarConfig v3.0! M\u1ed7i d\u00f2ng th\u00e9p hi\u1ec3n th\u1ecb tr\u1ef1c ti\u1ebfp Left..Mid..Right, SL thanh, Lbar, Lap & kg ngay khi nh\u1eadp."), vbInformation
    End If
End Sub

Public Sub EnsureRebarConfigSheet(ByVal wb As Workbook, Optional ByVal expandIfEmpty As Boolean = True)
    Dim ws As Worksheet, isNew As Boolean
    On Error Resume Next
    Set ws = wb.Worksheets(REBAR_CFG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = REBAR_CFG_SHEET
        BuildRebarConfigLayout ws
        isNew = True
    ElseIf CStr(ws.Range("A1").Value) <> REBAR_CFG_MARKER Then
        ' Upgrade legacy v2 layout cleanly when upgrading workbook
        BuildRebarConfigLayout ws
        isNew = True
    End If

    If isNew Or (expandIfEmpty And GetLastConfigRow(ws) < REBAR_CFG_FIRST_ROW) Then
        ExpandRebarPresetsFromQsData False
    End If
End Sub

Private Sub BuildRebarConfigLayout(ByVal ws As Worksheet)
    Application.ScreenUpdating = False
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 9.5

    ws.Range("A1").Value = REBAR_CFG_MARKER
    ws.Range("A1").Font.Color = RGB(22, 54, 92)
    ws.Range("A1").Interior.Color = RGB(22, 54, 92)

    With ws.Range("B1:AL1")
        .Merge
        .Value = modUnicode.U("B\u1ea2NG NH\u1eacP & KI\u1ec2M SO\u00c1T TH\u00c9P T\u01af\u1edcNG MINH (QS_REBARCONFIG v3.0 - TRANSPARENT PARAMETRIC BBS)")
        .Font.Size = 12.5: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(22, 54, 92): .RowHeight = 28: .IndentLevel = 1: .VerticalAlignment = xlCenter
    End With
    ws.Range("AM1").Interior.Color = RGB(22, 54, 92)

    With ws.Range("A2:AJ2")
        .Merge
        .Value = modUnicode.U("Quy \u01b0\u1edbc M\u1ecf (HookL/R): g\u00f5 <50 t\u1ef1 nh\u00e2n xD (VD: 20 = 20D), g\u00f5 >=50 l\u1ea5y th\u1eb3ng mm (VD: 150). \u0110\u1ec3 tr\u1ed1ng SL.m / L.m = t\u1ef1 t\u00ednh; g\u00f5 s\u1ed1 = ghi \u0111\u00e8 tr\u1ef1c ti\u1ebfp!")
        .Font.Italic = True: .Font.Size = 9: .Font.Color = RGB(40, 40, 40)
        .Interior.Color = RGB(237, 242, 248): .RowHeight = 20: .IndentLevel = 1: .VerticalAlignment = xlCenter
    End With
    ws.Range("AK2").Value = "TOTAL (kg):"
    ws.Range("AK2").Font.Bold = True: ws.Range("AK2").HorizontalAlignment = xlRight
    ws.Range("AK2").Interior.Color = RGB(226, 239, 218)
    ws.Range("AL2").Formula = "=SUBTOTAL(9,AL5:AL5000)"
    ws.Range("AL2").Font.Bold = True: ws.Range("AL2").Font.Size = 11
    ws.Range("AL2").NumberFormat = "#,##0.00"
    ws.Range("AL2").Interior.Color = RGB(226, 239, 218)

    ' Row 3: 3-Zone Group Header Banner
    With ws.Range("A3:H3")
        .Merge
        .Value = modUnicode.U("A. LI\u00caN K\u1ebeT C\u1ea4U KI\u1ec6N (LINK T\u1ef0 \u0110\u1ed8NG T\u1eea QS_Data)")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(58, 90, 128): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    With ws.Range("I3:Y3")
        .Merge
        .Value = modUnicode.U("B. THAM S\u1ed0 NH\u1eacP TH\u00c9P (D\u1ec4 NH\u1eacP - \u0110\u1ea6Y \u0110\u1ee6 OVERRIDE SL.m / L.m / +-Pb / +-Lth)")
        .Font.Bold = True: .Font.Color = RGB(40, 30, 0)
        .Interior.Color = RGB(255, 217, 102): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    With ws.Range("Z3:AL3")
        .Merge
        .Value = modUnicode.U("C. K\u1ebeT QU\u1ea2 T\u01af\u1edcNG MINH LIVE (100% C\u00d4NG TH\u1ee8C EXCEL 2007-365)")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(56, 87, 35): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Range("A3:AL3").RowHeight = 20

    ' Row 4: 39 Columns Header
    Dim hdrs As Variant
    hdrs = Array( _
        "ElementID", "Building", "Group", "Mark", "QT'Y Ele", "X / L (mm)", "Y / W (mm)", "H (mm)", _
        "BarID", "Role / Vi tri", "Shape", "D (mm)", "Part", "BarDir", "DistDir", _
        "HookL (xD|mm)", "Bv (mm)", "HookR (xD|mm)", "@1 / @X", "@2 / @Y", "SL.m (Ovr)", "L.m (Ovr)", "+-Pb (mm)", "+-Lth (mm)", "Ghi chu", _
        "Left (L1)", "Le.1a (L2)", "Le.1b (L3)", "Mid (L4)", "Ri.2b (L5)", "Ri.2a (L6)", "Right (L7)", _
        "QT'Y bar", "Lbar (mm)", "Lap (mm)", "Sum Len (m)", "Density", "Rebar (kg)", "RowKind" _
    )
    Dim c As Long
    For c = 0 To UBound(hdrs): ws.Cells(4, c + 1).Value = hdrs(c): Next c
    With ws.Range("A4:AM4")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 85, 151): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True: .RowHeight = 28
    End With
    ws.Range("AM1:AM2000").Font.Color = RGB(160, 160, 160)
    ws.Columns(39).ColumnWidth = 8

    ' Dropdown Validations (Rows 5..2000)
    AddValList ws.Range("A5:A2000"), "=QS_Data!$A$5:$A$2000"
    AddValList ws.Range("J5:J2000"), "BOT.X,BOT.Y,TOP.X,TOP.Y,TOP,BOT,SIDE,TOP-Goi,BOT-Nhip,MAIN.A,MAIN.B,MAIN.V,MAIN.H,STIR.R,STIR.IN,STIR.C,TIE.C,HAT.X,HAT.Y,CHAIR,ST-01,ST-03,ST-04,CUSTOM"
    AddValList ws.Range("K5:K2000"), "D01,D06,D07,D08,D15,D17,D24,D39,D41"
    AddValList ws.Range("L5:L2000"), "6,8,10,12,14,16,18,20,22,25,28,32,36,40"
    AddValList ws.Range("N5:N2000"), "X,Y,H,X*0.25,Y*0.25,X*0.7,PERIM_YH,PERIM_XY,PERIM_XY_INNER,TIE_Y,TIE_X,SLOPED_XH,CHAIR_H,CUSTOM"
    AddValList ws.Range("O5:O2000"), "X,Y,H,ZONE2_X,ZONE2_H,XY_GRID,XH_GRID,FIXED"

    Dim widths As Variant
    widths = Array( _
        14, 15, 11, 8, 7.5, 9, 9, 8.5, _
        10, 13, 7.5, 7, 6, 13, 10.5, _
        9.5, 7.5, 9.5, 8.5, 8.5, 9.5, 9.5, 8.5, 10, 26, _
        8.5, 8, 8, 9.5, 8, 8, 8.5, _
        9, 9.5, 8.5, 10.5, 8, 11.5, 8 _
    )
    For c = 0 To UBound(widths): ws.Columns(c + 1).ColumnWidth = widths(c): Next c

    On Error Resume Next
    If Not Application.ActiveWindow Is Nothing Then
        Application.ActiveWindow.SplitRow = 4
        Application.ActiveWindow.SplitColumn = 4
        Application.ActiveWindow.FreezePanes = True
        Application.ActiveWindow.DisplayGridlines = True
    End If
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

Private Sub AddValList(ByVal rng As Range, ByVal formulaStr As String)
    On Error Resume Next
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=formulaStr
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0
End Sub

Public Function GetLastConfigRow(ByVal ws As Worksheet) As Long
    Dim r1 As Long, r2 As Long, r3 As Long
    r1 = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    r2 = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    r3 = ws.Cells(ws.Rows.Count, 39).End(xlUp).Row
    GetLastConfigRow = r1
    If r2 > GetLastConfigRow Then GetLastConfigRow = r2
    If r3 > GetLastConfigRow Then GetLastConfigRow = r3
End Function

Public Sub ExpandRebarPresetsFromQsData(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook

    Dim wsData As Worksheet
    On Error Resume Next
    Set wsData = wb.Worksheets(modQsData.QSDATA_SHEET)
    On Error GoTo 0
    If wsData Is Nothing Then
        If showMsg Then modUnicode.MsgBoxU modUnicode.U("Ch\u01b0a c\u00f3 sheet QS_Data!"), vbExclamation
        Exit Sub
    End If

    Dim errMsg As String
    If Not modQsData.ValidateQsData(wsData, errMsg) Then
        If showMsg Then modUnicode.MsgBoxU errMsg, vbExclamation
        Exit Sub
    End If

    Dim wsCfg As Worksheet
    On Error Resume Next
    Set wsCfg = wb.Worksheets(REBAR_CFG_SHEET)
    On Error GoTo 0
    If wsCfg Is Nothing Then
        Set wsCfg = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsCfg.Name = REBAR_CFG_SHEET
        BuildRebarConfigLayout wsCfg
    ElseIf CStr(wsCfg.Range("A1").Value) <> REBAR_CFG_MARKER Then
        BuildRebarConfigLayout wsCfg
    End If

    Dim reqJson As String
    reqJson = "{""spec"":" & modSpecSheet.SerializeSpecToJson(wb) & _
              ",""elements"":" & modQsData.SerializeQsDataToJson(wsData) & _
              ",""existingConfigs"":" & SerializeRebarConfigsToJson(wsCfg) & "}"

    Dim resJson As String
    If Not modEngineBridge.RunEngineCommand("expand_rebar_presets", reqJson, resJson, errMsg) Then
        If showMsg Then modUnicode.MsgBoxU errMsg, vbCritical
        Exit Sub
    End If

    Dim root As Object: Set root = modEngineBridge.JsonParseObject(resJson)
    Dim dataObj As Object: Set dataObj = root("data")
    Dim rowsCol As Collection: Set rowsCol = dataObj("rows")
    Dim addedCnt As Long: addedCnt = rowsCol.Count

    If addedCnt > 0 Then
        Application.ScreenUpdating = False
        ' Group returned rows by ElementID preserving order
        Dim elOrder As Collection: Set elOrder = New Collection
        Dim elGroups As Object: Set elGroups = CreateObject("Scripting.Dictionary"): elGroups.CompareMode = vbTextCompare
        Dim i As Long, it As Object, eid As String, grpCol As Collection
        For i = 1 To addedCnt
            Set it = rowsCol(i)
            eid = CStr(it("elementId"))
            If Not elGroups.Exists(eid) Then
                Set grpCol = New Collection
                elGroups.Add eid, grpCol
                elOrder.Add eid
            End If
            elGroups(eid).Add it
        Next i

        Dim curR As Long: curR = GetLastConfigRow(wsCfg) + 1
        If curR < REBAR_CFG_FIRST_ROW Then curR = REBAR_CFG_FIRST_ROW

        Dim k As Long, bIdx As Long, bars As Collection
        For k = 1 To elOrder.Count
            eid = elOrder(k)
            Set bars = elGroups(eid)
            Dim hdrR As Long: hdrR = curR
            Dim firstBarR As Long: firstBarR = hdrR + 1
            Dim endR As Long: endR = hdrR + bars.Count + 1

            WriteBlockHeaderRow wsCfg, hdrR, eid, firstBarR, endR
            For bIdx = 1 To bars.Count
                WriteBarRowFromDto wsCfg, hdrR + bIdx, bars(bIdx)
            Next bIdx
            WriteBlockEndRow wsCfg, endR, eid

            curR = endR + 1
        Next k
        Application.ScreenUpdating = True
    End If

    If showMsg Then
        wsCfg.Activate
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 n\u1ea1p th\u00eam ") & addedCnt & modUnicode.U(" thanh th\u00e9p v\u00e0o c\u00e1c Block theo c\u1ea5u ki\u1ec7n t\u1eeb QS_Data (c\u00e1c Block \u0111\u00e3 c\u00f3 \u0111\u01b0\u1ee3c gi\u1eef nguy\u00ean)."), vbInformation
    End If
End Sub

Public Sub AppendOrUpdateRebarBlock(ByVal wsCfg As Worksheet, ByVal eid As String, ByVal bars As Collection)
    If wsCfg Is Nothing Or bars Is Nothing Then Exit Sub
    If bars.Count = 0 Then Exit Sub

    Dim lastR As Long: lastR = GetLastConfigRow(wsCfg)
    Dim hdrR As Long: hdrR = 0
    Dim endR As Long: endR = 0
    Dim r As Long

    For r = REBAR_CFG_FIRST_ROW To lastR
        If StrComp(Trim$(CStr(wsCfg.Cells(r, 1).Value)), eid, vbTextCompare) = 0 Then
            If StrComp(Trim$(CStr(wsCfg.Cells(r, 9).Value)), "[BLOCK]", vbTextCompare) = 0 Then
                hdrR = r
            ElseIf StrComp(Trim$(CStr(wsCfg.Cells(r, 9).Value)), "--- End element ---", vbTextCompare) = 0 Then
                endR = r
                Exit For
            End If
        End If
    Next r

    If hdrR > 0 And endR >= hdrR Then
        wsCfg.Rows(hdrR & ":" & endR).Delete
    End If

    Dim curR As Long: curR = GetLastConfigRow(wsCfg) + 1
    If curR < REBAR_CFG_FIRST_ROW Then curR = REBAR_CFG_FIRST_ROW

    hdrR = curR
    Dim firstBarR As Long: firstBarR = hdrR + 1
    endR = hdrR + bars.Count + 1

    WriteBlockHeaderRow wsCfg, hdrR, eid, firstBarR, endR
    Dim bIdx As Long
    For bIdx = 1 To bars.Count
        WriteBarRowFromDto wsCfg, hdrR + bIdx, bars(bIdx)
    Next bIdx
    WriteBlockEndRow wsCfg, endR, eid
End Sub

Private Sub WriteBlockHeaderRow(ByVal ws As Worksheet, ByVal r As Long, ByVal eid As String, ByVal firstBarR As Long, ByVal endR As Long)
    ws.Cells(r, 1).Value = eid
    FillZoneAMasterFormulas ws, r
    ws.Cells(r, 9).Value = "[BLOCK]"
    ws.Cells(r, 10).Formula = "=""=== ""&D" & r & "&"" (""&E" & r & "&"" | ""&F" & r & "&""x""&G" & r & "&""x""&H" & r & "&"") ==="""
    ws.Cells(r, 38).Formula = "=SUBTOTAL(9,AL" & firstBarR & ":AL" & endR & ")"
    ws.Cells(r, 39).Value = "BLK_HDR"

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 38))
        .Interior.Color = RGB(255, 242, 153)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin: .Borders.Color = RGB(180, 180, 180)
    End With
    ws.Cells(r, 38).NumberFormat = "#,##0.00"
    ws.Rows(r).RowHeight = 20
End Sub

Private Sub WriteBlockEndRow(ByVal ws As Worksheet, ByVal r As Long, ByVal eid As String)
    ws.Cells(r, 1).Value = eid
    ws.Cells(r, 9).Value = "--- End element ---"
    ws.Cells(r, 39).Value = "BLK_END"
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 38))
        .Interior.Color = RGB(242, 242, 242)
        .Font.Italic = True: .Font.Size = 8.5: .Font.Color = RGB(140, 140, 140)
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin: .Borders.Color = RGB(220, 220, 220)
    End With
    ws.Rows(r).RowHeight = 15
End Sub

Private Sub FillZoneAMasterFormulas(ByVal ws As Worksheet, ByVal r As Long)
    ws.Cells(r, 2).Formula = "=IFERROR(INDEX(QS_Data!$B$5:$B$2000,MATCH(A" & r & ",QS_Data!$A$5:$A$2000,0)),"""")"
    ws.Cells(r, 3).Formula = "=IFERROR(INDEX(QS_Data!$C$5:$C$2000,MATCH(A" & r & ",QS_Data!$A$5:$A$2000,0)),"""")"
    ws.Cells(r, 4).Formula = "=IFERROR(INDEX(QS_Data!$D$5:$D$2000,MATCH(A" & r & ",QS_Data!$A$5:$A$2000,0)),"""")"
    ws.Cells(r, 5).Formula = "=IFERROR(INDEX(QS_Data!$F$5:$F$2000,MATCH(A" & r & ",QS_Data!$A$5:$A$2000,0)),1)"
    ws.Cells(r, 6).Formula = "=IFERROR(INDEX(QS_Data!$G$5:$G$2000,MATCH(A" & r & ",QS_Data!$A$5:$A$2000,0))*1000,0)"
    ws.Cells(r, 7).Formula = "=IFERROR(INDEX(QS_Data!$H$5:$H$2000,MATCH(A" & r & ",QS_Data!$A$5:$A$2000,0))*1000,0)"
    ws.Cells(r, 8).Formula = "=IFERROR(INDEX(QS_Data!$I$5:$I$2000,MATCH(A" & r & ",QS_Data!$A$5:$A$2000,0))*1000,0)"
    ws.Range(ws.Cells(r, 5), ws.Cells(r, 8)).NumberFormat = "#,##0"
End Sub

Private Sub FillZoneCLiveFormulas(ByVal ws As Worksheet, ByVal r As Long)
    ws.Cells(r, 26).Formula = "=IF(P" & r & "="""",0,IF(P" & r & "<50,L" & r & "*P" & r & ",P" & r & "))"
    ws.Cells(r, 29).Formula = BuildMidFormula(r)
    ws.Cells(r, 32).Formula = "=IF(R" & r & "="""",0,IF(R" & r & "<50,L" & r & "*R" & r & ",R" & r & "))"
    ws.Cells(r, 33).Formula = BuildQtyFormula(r)
    ws.Cells(r, 34).Formula = "=SUM(Z" & r & ":AF" & r & ")"
    ws.Cells(r, 35).Formula = BuildLapFormula(r)
    ws.Cells(r, 36).Formula = "=ROUND(E" & r & "*M" & r & "*AG" & r & "*(AH" & r & "+AI" & r & ")/1000,2)"
    ws.Cells(r, 37).Formula = "=ROUND(PI()*(L" & r & "/1000)^2/4*7850,3)"
    ws.Cells(r, 38).Formula = "=ROUND(AJ" & r & "*AK" & r & ",2)"

    ws.Range(ws.Cells(r, 26), ws.Cells(r, 35)).NumberFormat = "#,##0"
    ws.Cells(r, 36).NumberFormat = "#,##0.00"
    ws.Cells(r, 37).NumberFormat = "0.000"
    ws.Cells(r, 38).NumberFormat = "#,##0.00"
End Sub

Private Function BuildMidFormula(ByVal r As Long) As String
    Dim dlt As String: dlt = "IF(X" & r & "="""",0,X" & r & ")"
    Dim s As String
    s = "=MAX(0,IF(V" & r & "<>"""",V" & r & ","
    s = s & "IF(N" & r & "=""X"",F" & r & "-2*Q" & r & ","
    s = s & "IF(N" & r & "=""Y"",G" & r & "-2*Q" & r & ","
    s = s & "IF(N" & r & "=""H"",H" & r & "-2*Q" & r & ","
    s = s & "IF(N" & r & "=""X*0.25"",ROUND(F" & r & "*0.25,0),"
    s = s & "IF(N" & r & "=""Y*0.25"",ROUND(G" & r & "*0.25,0),"
    s = s & "IF(N" & r & "=""X*0.7"",ROUND(F" & r & "*0.7,0),"
    s = s & "IF(N" & r & "=""PERIM_YH"",2*(G" & r & "+H" & r & "-4*Q" & r & "),"
    s = s & "IF(N" & r & "=""PERIM_XY"",2*(F" & r & "+G" & r & "-4*Q" & r & "),"
    s = s & "IF(N" & r & "=""PERIM_XY_INNER"",2*(ROUND(F" & r & "*0.67,0)+G" & r & "-4*Q" & r & "),"
    s = s & "IF(N" & r & "=""TIE_Y"",G" & r & "-2*Q" & r & ","
    s = s & "IF(N" & r & "=""TIE_X"",F" & r & "-2*Q" & r & ","
    s = s & "IF(N" & r & "=""SLOPED_XH"",ROUND(SQRT(F" & r & "^2+H" & r & "^2),0),"
    s = s & "IF(N" & r & "=""CHAIR_H"",MAX(120,(H" & r & "-2*Q" & r & ")*2+150),0"
    s = s & ")))))))))))))+" & dlt & "))"
    BuildMidFormula = s
End Function

Private Function BuildQtyFormula(ByVal r As Long) As String
    Dim dp As String: dp = "IF(W" & r & "="""",0,W" & r & ")"
    Dim sp1 As String: sp1 = "MAX(1,S" & r & ")"
    Dim sp2 As String: sp2 = "MAX(1,IF(T" & r & "="""",S" & r & ",T" & r & "))"
    Dim s As String
    s = "=IFERROR(IF(U" & r & "<>"""",U" & r & ","
    s = s & "IF(O" & r & "=""X"",ROUND((F" & r & "+" & dp & ")/" & sp1 & ",0)+1,"
    s = s & "IF(O" & r & "=""Y"",ROUND((G" & r & "+" & dp & ")/" & sp1 & ",0)+1,"
    s = s & "IF(O" & r & "=""H"",ROUND((H" & r & "+" & dp & ")/" & sp1 & ",0)+1,"
    s = s & "IF(O" & r & "=""ZONE2_X"",ROUND(((F" & r & "+" & dp & ")*0.5)/" & sp1 & ",0)+ROUND(((F" & r & "+" & dp & ")*0.5)/" & sp2 & ",0)+2,"
    s = s & "IF(O" & r & "=""ZONE2_H"",ROUND(((H" & r & "+" & dp & ")*0.5)/" & sp1 & ",0)+ROUND(((H" & r & "+" & dp & ")*0.5)/" & sp2 & ",0)+1,"
    s = s & "IF(O" & r & "=""XY_GRID"",(ROUND((F" & r & "+" & dp & ")/" & sp1 & ",0)+1)*(ROUND((G" & r & "+" & dp & ")/" & sp2 & ",0)+1),"
    s = s & "IF(O" & r & "=""XH_GRID"",(ROUND((F" & r & "+" & dp & ")/" & sp1 & ",0)+1)*(ROUND((H" & r & "+" & dp & ")/" & sp2 & ",0)+1),1"
    s = s & ")))))))),1)"
    BuildQtyFormula = s
End Function

Private Function BuildLapFormula(ByVal r As Long) As String
    Dim s As String
    s = "=IF(L" & r & "<10,0,IF(L" & r & ">=IFERROR(SPEC_CouplerMinD,99),0,"
    s = s & "MAX(0,ROUNDUP(AH" & r & "/IFERROR(SPEC_StdBarLen,11700),0)-1)*L" & r & "*IFERROR(SPEC_Llap_Mult,45)))"
    BuildLapFormula = s
End Function

Private Sub WriteBarRowFromDto(ByVal ws As Worksheet, ByVal r As Long, ByVal it As Object)
    ws.Cells(r, 1).Value = CStr(it("elementId"))
    FillZoneAMasterFormulas ws, r

    ws.Cells(r, 9).Value = CStr(it("barId"))
    ws.Cells(r, 10).Value = CStr(it("role"))
    ws.Cells(r, 11).Value = CStr(it("shapeCode"))
    ws.Cells(r, 12).Value = CDbl(it("dia"))
    ws.Cells(r, 13).Value = CDbl(it("layerMult"))
    ws.Cells(r, 14).Value = CStr(it("barDir"))
    ws.Cells(r, 15).Value = CStr(it("distDir"))
    AssignExprOrNum ws.Cells(r, 16), CStr(it("hookLExpr")), r
    AssignExprOrNum ws.Cells(r, 17), CStr(it("coverExpr")), r
    AssignExprOrNum ws.Cells(r, 18), CStr(it("hookRExpr")), r
    If Not IsNullOrZero(it("space1Mm")) Then ws.Cells(r, 19).Value = CDbl(it("space1Mm"))
    If Not IsNullOrZero(it("space2Mm")) Then ws.Cells(r, 20).Value = CDbl(it("space2Mm"))
    AssignExprOrNum ws.Cells(r, 21), CStr(it("overrideQtyExpr")), r
    AssignExprOrNum ws.Cells(r, 22), CStr(it("overrideLenExpr")), r
    AssignExprOrNum ws.Cells(r, 23), CStr(it("deltaDistExpr")), r
    AssignExprOrNum ws.Cells(r, 24), CStr(it("deltaLenExpr")), r
    ws.Cells(r, 25).Value = CStr(it("note"))

    If Not IsNullOrZero(it("le1aMm")) Then ws.Cells(r, 27).Value = CDbl(it("le1aMm"))
    If Not IsNullOrZero(it("le1bMm")) Then ws.Cells(r, 28).Value = CDbl(it("le1bMm"))
    If Not IsNullOrZero(it("ri2bMm")) Then ws.Cells(r, 30).Value = CDbl(it("ri2bMm"))
    If Not IsNullOrZero(it("ri2aMm")) Then ws.Cells(r, 31).Value = CDbl(it("ri2aMm"))

    FillZoneCLiveFormulas ws, r
    ws.Cells(r, 39).Value = "BAR"
    StyleBarRow ws, r
End Sub

Private Sub StyleBarRow(ByVal ws As Worksheet, ByVal r As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 38))
        .Font.Bold = False: .Font.Italic = False: .Font.Color = RGB(0, 0, 0): .Font.Size = 9.5
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin: .Borders.Color = RGB(210, 210, 210)
    End With
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Interior.Color = RGB(242, 245, 250)
    ws.Range(ws.Cells(r, 9), ws.Cells(r, 25)).Interior.Color = RGB(255, 252, 232)
    ws.Cells(r, 12).Interior.Color = RGB(255, 242, 153): ws.Cells(r, 12).Font.Bold = True
    ws.Range(ws.Cells(r, 26), ws.Cells(r, 37)).Interior.Color = RGB(255, 255, 255)
    ws.Cells(r, 38).Interior.Color = RGB(238, 247, 235): ws.Cells(r, 38).Font.Bold = True
    ws.Range(ws.Cells(r, 11), ws.Cells(r, 15)).HorizontalAlignment = xlCenter
    ws.Cells(r, 38).HorizontalAlignment = xlRight
    ws.Rows(r).RowHeight = 19
End Sub

Private Sub AssignExprOrNum(ByVal rng As Range, ByVal rawVal As String, ByVal r As Long)
    Dim v As String: v = Trim$(rawVal)
    If Len(v) = 0 Then Exit Sub
    v = Replace$(v, "{r}", CStr(r))
    If Left$(v, 1) = "=" Then
        rng.Formula = v
    ElseIf IsNumeric(v) Then
        rng.Value = Val(v)
    Else
        rng.Value = v
    End If
End Sub

Private Function IsNullOrZero(ByVal v As Variant) As Boolean
    If IsNull(v) Or IsEmpty(v) Then IsNullOrZero = True: Exit Function
    If Val(CStr(v)) <= 0 Then IsNullOrZero = True
End Function

Public Sub InsertBarRowInBlock(Optional ByVal presetKind As String = "MAIN", Optional ByVal shapeCode As String = "D07", Optional ByVal insertAtCursor As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook
    EnsureRebarConfigSheet wb, False
    Dim ws As Worksheet: Set ws = wb.Worksheets(REBAR_CFG_SHEET)
    ws.Activate

    Dim rCur As Long: rCur = ActiveCell.Row
    Dim rkCur As String: rkCur = ""
    Dim eid As String: eid = ""
    Dim targetRow As Long

    If rCur >= REBAR_CFG_FIRST_ROW Then
        eid = Trim$(CStr(ws.Cells(rCur, 1).Value))
        rkCur = UCase$(Trim$(CStr(ws.Cells(rCur, 39).Value)))
    End If

    If insertAtCursor And rCur >= REBAR_CFG_FIRST_ROW Then
        If rkCur = "BLK_END" Then
            targetRow = rCur
        Else
            targetRow = rCur + 1
        End If
        ws.Rows(targetRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Else
        targetRow = GetLastConfigRow(ws) + 1
        If targetRow < REBAR_CFG_FIRST_ROW Then targetRow = REBAR_CFG_FIRST_ROW
        If targetRow > REBAR_CFG_FIRST_ROW Then eid = Trim$(CStr(ws.Cells(targetRow - 1, 1).Value))
    End If

    If Len(eid) = 0 Then
        Dim rScan As Long
        For rScan = targetRow - 1 To REBAR_CFG_FIRST_ROW Step -1
            If Len(Trim$(CStr(ws.Cells(rScan, 1).Value))) > 0 Then
                eid = Trim$(CStr(ws.Cells(rScan, 1).Value)): Exit For
            End If
        Next rScan
    End If
    If Len(eid) = 0 Then eid = "EL-CUSTOM"

    ws.Cells(targetRow, 1).Value = eid
    FillZoneAMasterFormulas ws, targetRow

    Dim baseBar As String: baseBar = "ADD-" & (targetRow - REBAR_CFG_FIRST_ROW + 1)
    ws.Cells(targetRow, 9).Value = baseBar
    ws.Cells(targetRow, 13).Value = 1
    ws.Cells(targetRow, 17).Formula = "=IFERROR(SPEC_Cover_Beam,30)"

    Select Case UCase$(presetKind)
        Case "STIR"
            ws.Cells(targetRow, 10).Value = "STIR.R"
            ws.Cells(targetRow, 11).Value = "D39"
            ws.Cells(targetRow, 12).Value = 8
            ws.Cells(targetRow, 14).Value = "PERIM_YH"
            ws.Cells(targetRow, 15).Value = "ZONE2_X"
            ws.Cells(targetRow, 16).Value = 10
            ws.Cells(targetRow, 18).Value = 10
            ws.Cells(targetRow, 19).Value = 150
            ws.Cells(targetRow, 20).Value = 200
            ws.Cells(targetRow, 25).Value = modUnicode.U("Th\u00eam \u0111ai t\u00f9y ch\u1ec9nh")
        Case "CUSTOM"
            ws.Cells(targetRow, 10).Value = "CUSTOM"
            ws.Cells(targetRow, 11).Value = UCase$(shapeCode)
            ws.Cells(targetRow, 12).Value = 14
            ws.Cells(targetRow, 14).Value = "X"
            ws.Cells(targetRow, 15).Value = "FIXED"
            ws.Cells(targetRow, 16).Value = 0
            ws.Cells(targetRow, 18).Value = 0
            ws.Cells(targetRow, 21).Value = 2
            ws.Cells(targetRow, 25).Value = modUnicode.U("Thanh t\u00f9y bi\u1ebfn (" & UCase$(shapeCode) & ")")
        Case Else ' "MAIN"
            ws.Cells(targetRow, 10).Value = "BOT.X"
            ws.Cells(targetRow, 11).Value = UCase$(shapeCode)
            ws.Cells(targetRow, 12).Value = 16
            ws.Cells(targetRow, 14).Value = "X"
            ws.Cells(targetRow, 15).Value = "Y"
            ws.Cells(targetRow, 16).Value = 150
            ws.Cells(targetRow, 18).Value = 150
            ws.Cells(targetRow, 19).Value = 150
            ws.Cells(targetRow, 25).Value = modUnicode.U("Th\u00eam thanh ch\u00ednh / gia c\u01b0\u1eddng")
    End Select

    FillZoneCLiveFormulas ws, targetRow
    ws.Cells(targetRow, 39).Value = "BAR"
    StyleBarRow ws, targetRow
    RefreshEnclosingBlockTotal ws, targetRow
    ws.Cells(targetRow, 11).Select
End Sub

Private Sub RefreshEnclosingBlockTotal(ByVal ws As Worksheet, ByVal rowInside As Long)
    Dim lr As Long: lr = GetLastConfigRow(ws)
    Dim hdrR As Long, endR As Long, rr As Long, rk As String
    For rr = rowInside To REBAR_CFG_FIRST_ROW Step -1
        rk = UCase$(Trim$(CStr(ws.Cells(rr, 39).Value)))
        If rk = "BLK_HDR" Then hdrR = rr: Exit For
        If rk = "BLK_END" And rr < rowInside Then Exit For
    Next rr
    If hdrR = 0 Then Exit Sub

    endR = hdrR
    For rr = hdrR + 1 To lr
        rk = UCase$(Trim$(CStr(ws.Cells(rr, 39).Value)))
        If rk = "BLK_END" Then endR = rr: Exit For
        If rk = "BLK_HDR" Then Exit For
        endR = rr
    Next rr
    ws.Cells(hdrR, 38).Formula = "=SUBTOTAL(9,AL" & (hdrR + 1) & ":AL" & endR & ")"
End Sub

Public Sub InsertCustomBarRowOnConfig(Optional ByVal shapeCode As String = "D01", Optional ByVal insertAtCursor As Boolean = True)
    InsertBarRowInBlock "CUSTOM", shapeCode, insertAtCursor
End Sub

Public Sub InsertMainBarFromMenu()
    InsertBarRowInBlock "MAIN", "D07", True
End Sub

Public Sub InsertStirrupBarFromMenu()
    InsertBarRowInBlock "STIR", "D39", True
End Sub

Public Sub InsertCustomBarFromMenu()
    InsertBarRowInBlock "CUSTOM", "D01", True
End Sub

Public Sub ExpandPresetsFromMenu()
    ExpandRebarPresetsFromQsData True
End Sub

Public Sub ApplyShapeToConfigRow(ByVal shapeCode As String)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not IsRebarConfigSheet(ws) Then Exit Sub
    Dim r As Long: r = ActiveCell.Row
    If r < REBAR_CFG_FIRST_ROW Then Exit Sub
    If UCase$(Trim$(CStr(ws.Cells(r, 39).Value))) = "BLK_HDR" Or UCase$(Trim$(CStr(ws.Cells(r, 39).Value))) = "BLK_END" Then Exit Sub

    Dim sc As String: sc = UCase$(Trim$(shapeCode))
    ws.Cells(r, REBAR_CFG_SHAPE_COL).Value = sc
    ws.Cells(r, REBAR_CFG_SHAPE_COL).Font.Bold = True

    ' Highlight extra bend segment cells AA, AB, AD, AE when shape uses >= 4 segments
    Dim segColor As Long
    Select Case sc
        Case "D08", "D13", "D17", "D18", "D19", "D20", "D21", "D22", "D23", "D28", "D29", "D30", "D31", "D32", "D33", "D35", "D36", "D37", "D38", "D40"
            segColor = RGB(255, 242, 204)
        Case Else
            segColor = RGB(255, 255, 255)
    End Select
    ws.Range(ws.Cells(r, 27), ws.Cells(r, 28)).Interior.Color = segColor
    ws.Range(ws.Cells(r, 30), ws.Cells(r, 31)).Interior.Color = segColor
End Sub

Public Sub OpenVisualRebarEditorFromMenu()
    OpenVisualRebarEditor True
End Sub

Public Sub OpenVisualRebarEditor(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim wsAct As Worksheet: Set wsAct = ActiveSheet

    Dim wsData As Worksheet
    On Error Resume Next
    Set wsData = wb.Worksheets(modQsData.QSDATA_SHEET)
    On Error GoTo 0
    If wsData Is Nothing Then
        If showMsg Then modUnicode.MsgBoxU modUnicode.U("Ch\u01b0a c\u00f3 sheet QS_Data!"), vbExclamation
        Exit Sub
    End If

    Dim hasSpec As Boolean
    On Error Resume Next
    hasSpec = Not (wb.Worksheets(modSpecSheet.SPEC_SHEET) Is Nothing)
    On Error GoTo 0
    If Not hasSpec Then modSpecSheet.CreateOrShowSpecSheet False

    EnsureRebarConfigSheet wb, False
    Dim wsCfg As Worksheet: Set wsCfg = wb.Worksheets(REBAR_CFG_SHEET)

    Dim eid As String, dataRow As Long, rCur As Long
    rCur = ActiveCell.Row

    If modQsData.IsQsDataSheet(wsAct) And rCur >= modQsData.QSDATA_FIRST_ROW Then
        modQsData.AutoGenerateElementIds wsData
        eid = Trim$(CStr(wsData.Cells(rCur, 1).Value))
        dataRow = rCur
    ElseIf IsRebarConfigSheet(wsAct) And rCur >= REBAR_CFG_FIRST_ROW Then
        eid = Trim$(CStr(wsCfg.Cells(rCur, 1).Value))
        dataRow = FindElementDataRow(wsData, eid)
    Else
        If showMsg Then modUnicode.MsgBoxU modUnicode.U("Vui l\u00f2ng ch\u1ecdn d\u00f2ng c\u1ea5u ki\u1ec7n tr\u00ean QS_Data ho\u1eb7c QS_RebarConfig r\u1ed3i m\u1edf B\u1ea3ng V\u1ebd Nh\u1eadp Th\u00e9p!"), vbExclamation
        Exit Sub
    End If

    If Len(eid) = 0 Or dataRow < modQsData.QSDATA_FIRST_ROW Then
        If showMsg Then modUnicode.MsgBoxU modUnicode.U("Kh\u00f4ng t\u00ecm th\u1ea5y th\u00f4ng tin c\u1ea5u ki\u1ec7n h\u1ee3p l\u1ec7 tr\u00ean QS_Data!"), vbExclamation
        Exit Sub
    End If

    Dim reqJson As String, resJson As String, errMsg As String
    reqJson = "{""parentHwnd"":" & modEngineBridge.JsonNum(CDbl(Application.Hwnd)) & ","
    reqJson = reqJson & """spec"":" & modSpecSheet.SerializeSpecToJson(wb) & ","
    reqJson = reqJson & """element"":" & modQsData.SerializeSingleElementToJson(wsData, dataRow) & ","
    reqJson = reqJson & """existingRows"":" & SerializeRebarConfigsToJson(wsCfg, eid) & "}"

    If Not modEngineBridge.RunEngineCommand("open_rebar_editor", reqJson, resJson, errMsg, 0) Then
        If showMsg Then modUnicode.MsgBoxU errMsg, vbCritical
        Exit Sub
    End If

    Dim root As Object: Set root = modEngineBridge.JsonParseObject(resJson)
    Dim dataObj As Object: Set dataObj = root("data")
    If Not CBool(dataObj("accepted")) Then Exit Sub

    Dim rowsCol As Collection: Set rowsCol = dataObj("rows")
    ReplaceOrAppendElementBlock wsCfg, eid, rowsCol
    wsCfg.Activate
End Sub

Private Function FindElementDataRow(ByVal wsData As Worksheet, ByVal eid As String) As Long
    If Len(eid) = 0 Then Exit Function
    Dim lr As Long: lr = modQsData.GetLastQsDataRow(wsData)
    Dim r As Long
    For r = modQsData.QSDATA_FIRST_ROW To lr
        If StrComp(Trim$(CStr(wsData.Cells(r, 1).Value)), eid, vbTextCompare) = 0 Then
            FindElementDataRow = r: Exit Function
        End If
    Next r
End Function

Public Sub ReplaceOrAppendElementBlock(ByVal ws As Worksheet, ByVal eid As String, ByVal rowsCol As Collection)
    Application.ScreenUpdating = False
    Dim lr As Long: lr = GetLastConfigRow(ws)
    Dim firstR As Long, lastR As Long, r As Long
    firstR = 0: lastR = 0

    If lr >= REBAR_CFG_FIRST_ROW Then
        For r = REBAR_CFG_FIRST_ROW To lr
            If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), eid, vbTextCompare) = 0 Then
                If firstR = 0 Then firstR = r
                lastR = r
            End If
        Next r
    End If

    Dim insertAt As Long
    If firstR > 0 Then
        ws.Rows(firstR & ":" & lastR).Delete Shift:=xlUp
        insertAt = firstR
    Else
        insertAt = GetLastConfigRow(ws) + 1
        If insertAt < REBAR_CFG_FIRST_ROW Then insertAt = REBAR_CFG_FIRST_ROW
    End If

    If rowsCol.Count > 0 Then
        Dim totalBlockRows As Long: totalBlockRows = rowsCol.Count + 2
        If firstR > 0 And insertAt <= GetLastConfigRow(ws) Then
            ws.Rows(insertAt & ":" & (insertAt + totalBlockRows - 1)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
        Dim endR As Long: endR = insertAt + totalBlockRows - 1
        WriteBlockHeaderRow ws, insertAt, eid, insertAt + 1, endR
        Dim i As Long
        For i = 1 To rowsCol.Count
            WriteBarRowFromDto ws, insertAt + i, rowsCol(i)
        Next i
        WriteBlockEndRow ws, endR, eid
        ws.Cells(insertAt, 1).Select
    End If
    Application.ScreenUpdating = True
End Sub

Public Function SerializeRebarConfigsToJson(ByVal ws As Worksheet, Optional ByVal filterElementId As String = "") As String
    Dim lr As Long: lr = GetLastConfigRow(ws)
    If lr < REBAR_CFG_FIRST_ROW Then SerializeRebarConfigsToJson = "[]": Exit Function

    Dim arr As Variant: arr = ws.Range(ws.Cells(REBAR_CFG_FIRST_ROW, 1), ws.Cells(lr, REBAR_CFG_LAST_COL)).Value2
    Dim fArr As Variant: fArr = ws.Range(ws.Cells(REBAR_CFG_FIRST_ROW, 1), ws.Cells(lr, 25)).Formula
    Dim items() As String, cnt As Long, cap As Long
    cap = 64: cnt = 0: ReDim items(0 To cap - 1)

    Dim i As Long, rr As Long, eid As String, rk As String
    For i = 1 To UBound(arr, 1)
        rr = REBAR_CFG_FIRST_ROW + i - 1
        eid = Trim$(CStr(arr(i, 1)))
        rk = UCase$(Trim$(CStr(arr(i, 39))))
        If Len(eid) > 0 And rk <> "BLK_HDR" And rk <> "BLK_END" And Val(CStr(arr(i, 12))) > 0 Then
            If Len(filterElementId) = 0 Or StrComp(eid, filterElementId, vbTextCompare) = 0 Then
                Dim covS As String, dLthS As String, lOvrS As String, dPbS As String, qOvrS As String
                covS = PreserveRowPlaceholder(CStr(fArr(i, 17)), rr)
                qOvrS = PreserveRowPlaceholder(CStr(fArr(i, 21)), rr)
                lOvrS = PreserveRowPlaceholder(CStr(fArr(i, 22)), rr)
                dPbS = PreserveRowPlaceholder(CStr(fArr(i, 23)), rr)
                dLthS = PreserveRowPlaceholder(CStr(fArr(i, 24)), rr)

                Dim s As String
                s = "{""configRow"":" & rr & ",""rowKind"":""BAR"","
                s = s & """elementId"":" & modEngineBridge.JsonStr(eid) & ","
                s = s & """building"":" & modEngineBridge.JsonStr(arr(i, 2)) & ","
                s = s & """group"":" & modEngineBridge.JsonStr(arr(i, 3)) & ","
                s = s & """barId"":" & modEngineBridge.JsonStr(arr(i, 9)) & ","
                s = s & """role"":" & modEngineBridge.JsonStr(arr(i, 10)) & ","
                s = s & """shapeCode"":" & modEngineBridge.JsonStr(arr(i, 11)) & ","
                s = s & """dia"":" & modEngineBridge.JsonNum(arr(i, 12)) & ","
                s = s & """layerMult"":" & modEngineBridge.JsonNum(IIf(Val(CStr(arr(i, 13))) > 0, arr(i, 13), 1)) & ","
                s = s & """barDir"":" & modEngineBridge.JsonStr(arr(i, 14)) & ","
                s = s & """distDir"":" & modEngineBridge.JsonStr(arr(i, 15)) & ","
                s = s & """hookLExpr"":" & modEngineBridge.JsonStr(arr(i, 16)) & ","
                s = s & """coverExpr"":" & modEngineBridge.JsonStr(covS) & ","
                s = s & """hookRExpr"":" & modEngineBridge.JsonStr(arr(i, 18)) & ","
                If IsNumeric(arr(i, 19)) And Not IsEmpty(arr(i, 19)) Then s = s & """space1Mm"":" & modEngineBridge.JsonNum(arr(i, 19)) & ","
                If IsNumeric(arr(i, 20)) And Not IsEmpty(arr(i, 20)) Then s = s & """space2Mm"":" & modEngineBridge.JsonNum(arr(i, 20)) & ","
                s = s & """overrideQtyExpr"":" & modEngineBridge.JsonStr(qOvrS) & ","
                s = s & """overrideLenExpr"":" & modEngineBridge.JsonStr(lOvrS) & ","
                s = s & """deltaDistExpr"":" & modEngineBridge.JsonStr(dPbS) & ","
                s = s & """deltaLenExpr"":" & modEngineBridge.JsonStr(dLthS) & ","
                If IsNumeric(arr(i, 27)) And Not IsEmpty(arr(i, 27)) Then s = s & """le1aMm"":" & modEngineBridge.JsonNum(arr(i, 27)) & ","
                If IsNumeric(arr(i, 28)) And Not IsEmpty(arr(i, 28)) Then s = s & """le1bMm"":" & modEngineBridge.JsonNum(arr(i, 28)) & ","
                If IsNumeric(arr(i, 30)) And Not IsEmpty(arr(i, 30)) Then s = s & """ri2bMm"":" & modEngineBridge.JsonNum(arr(i, 30)) & ","
                If IsNumeric(arr(i, 31)) And Not IsEmpty(arr(i, 31)) Then s = s & """ri2aMm"":" & modEngineBridge.JsonNum(arr(i, 31)) & ","
                s = s & """note"":" & modEngineBridge.JsonStr(arr(i, 25)) & "}"
                If cnt >= cap Then cap = cap * 2: ReDim Preserve items(0 To cap - 1)
                items(cnt) = s: cnt = cnt + 1
            End If
        End If
    Next i

    If cnt = 0 Then
        SerializeRebarConfigsToJson = "[]"
    Else
        ReDim Preserve items(0 To cnt - 1)
        SerializeRebarConfigsToJson = "[" & Join(items, ",") & "]"
    End If
End Function

Private Function PreserveRowPlaceholder(ByVal rawStr As String, ByVal r As Long) As String
    If Left$(Trim$(rawStr), 1) = "=" Then
        PreserveRowPlaceholder = Replace$(rawStr, "L" & CStr(r), "L{r}")
    Else
        PreserveRowPlaceholder = rawStr
    End If
End Function
