Attribute VB_Name = "modRebarTKT"
Option Explicit

' ==============================================================================
' modRebarTKT - Rebar Bending Schedule (TKT Pro) Engine (41 Standard Shapes)
' Columns A..X: L1..L7 (F..L), M1/M (M), M2 (N), Dia (O), Qty/El (P), ElQty (Q),
'               TotBars (R), BarLen (S), Coupler (T), TotLen (U), kg/m (V), Weight (W), Note (X)
' ==============================================================================

Public Const TKT_FIRST_DATA_ROW As Long = 11
Public Const TKT_MAX_DATA_ROW As Long = 5000
Public Const TKT_AUTO_SHAPE_LIMIT As Long = 250
Private Const TKT_MARKER_CELL As String = "A1"
Private Const TKT_MARKER_TAG As String = "ZOFIT_TKT_V2"

Public Function IsTKTSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    If CStr(ws.Range(TKT_MARKER_CELL).Value) = TKT_MARKER_TAG Or Left$(ws.Name, 4) = "TKT_" Then
        If Trim$(CStr(ws.Cells(10, 4).Value)) <> "" And Trim$(CStr(ws.Cells(10, 15).Value)) <> "" Then
            IsTKTSheet = True
        End If
    End If
    On Error GoTo 0
End Function

Public Sub CreateTKTSheet(ByVal baseName As String, ByVal titleVI As String, Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Dim hasSet As Boolean
    On Error Resume Next
    hasSet = Not (wb.Worksheets("QS_Settings") Is Nothing)
    On Error GoTo 0
    If Not hasSet Then
        modSettings.CreateOrShowSettingsSheet False
    End If

    Dim sName As String
    sName = GetUniqueSheetName(wb, baseName)

    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sName

    BuildTKTLayout ws, titleVI
    ws.Activate

    Application.ScreenUpdating = True
    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("T\u1ea1o b\u1ea3ng th\u1ed1ng k\u00ea th\u00e9p th\u00e0nh c\u00f4ng: ") & sName & vbCrLf & _
            modUnicode.U("M\u1eb9o: Click 'Ch\u1ecdn D\u00e1ng Th\u00e9p' \u0111\u1ec3 xem minh h\u1ecda L1..L7, M1, M2 v\u00e0 th\u00eam thanh th\u00e9p!"), vbInformation
    End If
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    modUnicode.MsgBoxU "CreateTKTSheet Error: " & Err.Description, vbCritical
End Sub

Private Function GetUniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim cleanBase As String
    cleanBase = Left$(baseName, 27)
    If Not SheetExists(wb, cleanBase) Then
        GetUniqueSheetName = cleanBase
        Exit Function
    End If
    Dim idx As Long
    idx = 2
    Do While SheetExists(wb, cleanBase & "_" & idx)
        idx = idx + 1
    Loop
    GetUniqueSheetName = cleanBase & "_" & idx
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal sName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sName)
    SheetExists = Not (ws Is Nothing)
    On Error GoTo 0
End Function

Private Sub BuildTKTLayout(ByVal ws As Worksheet, ByVal titleVI As String)
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 9.5

    ws.Range(TKT_MARKER_CELL).Value = TKT_MARKER_TAG
    ws.Range(TKT_MARKER_CELL).Font.Color = RGB(22, 54, 92)

    With ws.Range("A1:X1")
        .Interior.Color = RGB(22, 54, 92)
        .RowHeight = 28
    End With
    With ws.Range("B1:X1")
        .Merge
        .Value = modUnicode.U("B\u1ea2NG TH\u1ed0NG K\u00ca C\u1ed0T TH\u00c9P CHI TI\u1ebeT - ") & UCase$(titleVI)
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ws.Range("B2").Value = modUnicode.U("D\u1ef1 \u00e1n:")
    ws.Range("B2").Font.Bold = True
    ws.Range("C2:H2").Merge
    ws.Range("C2").Formula = "=IFERROR(CFG_ProjectName,"""")"
    ws.Range("C2").Font.Bold = True

    ws.Range("I2").Value = modUnicode.U("Quy t\u1eafc:")
    ws.Range("I2").Font.Bold = True
    ws.Range("J2:X2").Merge
    ws.Range("J2").Formula = _
        "=" & """" & modUnicode.U("N\u1ed1i=") & """" & "&IFERROR(CFG_LapRatio,40)&""D | ""&" & _
        """" & modUnicode.U("Gi\u1ea3m u\u1ed1n=") & """" & "&IFERROR(CFG_BendDeduct,1)&""D | ""&" & _
        """" & modUnicode.U("C\u00e2y=") & """" & "&IFERROR(CFG_BarStdLen,11700)&""mm | Coupler>=""&CHAR(216)&IFERROR(CFG_CouplerMinD,22)"
    ws.Range("J2").Font.Color = RGB(31, 78, 120)

    ' Top Summary Strip (Rows 4..8, Columns B..X)
    With ws.Range("B4:X4")
        .Merge
        .Value = modUnicode.U("B\u1ea2NG T\u1ed4NG H\u1ee2P NHANH THEO \u0110\u01af\u1edcNG K\u00cdNH (T\u1ef0 \u0110\u1ed8NG C\u1eacP NH\u1eacT)")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 85, 151)
        .HorizontalAlignment = xlCenter
        .RowHeight = 20
    End With

    ws.Range("B5:C5").Merge: ws.Range("B5").Value = modUnicode.U("\u0110\u01b0\u1eddng k\u00ednh (mm)")
    ws.Range("B6:C6").Merge: ws.Range("B6").Value = modUnicode.U("T\u1ed5ng chi\u1ec1u d\u00e0i (m)")
    ws.Range("B7:C7").Merge: ws.Range("B7").Value = modUnicode.U("T\u1ed5ng kh\u1ed1i l\u01b0\u1ee3ng (kg)")
    ws.Range("B8:C8").Merge: ws.Range("B8").Value = modUnicode.U("Ph\u00e2n nh\u00f3m d\u1ef1 to\u00e1n (kg)")
    ws.Range("B5:C8").Font.Bold = True
    ws.Range("B5:C8").Interior.Color = RGB(237, 242, 248)

    Dim dias As Variant
    dias = Array(6, 8, 10, 12, 14, 16, 18, 20, 22, 25, 28, 32, 36, 40)
    Dim i As Long, colIdx As Long, colLet As String
    For i = LBound(dias) To UBound(dias)
        colIdx = 4 + i ' D..Q (14 diameters)
        colLet = Split(ws.Cells(1, colIdx).Address(True, False), "$")(0)
        ws.Cells(5, colIdx).Value = dias(i)
        ws.Cells(5, colIdx).Font.Bold = True
        ws.Cells(5, colIdx).HorizontalAlignment = xlCenter
        ws.Cells(5, colIdx).NumberFormat = """" & ChrW$(216) & """0"
        ws.Cells(5, colIdx).Interior.Color = RGB(217, 225, 242)

        ws.Cells(6, colIdx).Formula = "=SUMIF($O$" & TKT_FIRST_DATA_ROW & ":$O$" & TKT_MAX_DATA_ROW & "," & colLet & "5,$U$" & TKT_FIRST_DATA_ROW & ":$U$" & TKT_MAX_DATA_ROW & ")"
        ws.Cells(6, colIdx).NumberFormat = "#,##0.0"

        ws.Cells(7, colIdx).Formula = "=SUMIF($O$" & TKT_FIRST_DATA_ROW & ":$O$" & TKT_MAX_DATA_ROW & "," & colLet & "5,$W$" & TKT_FIRST_DATA_ROW & ":$W$" & TKT_MAX_DATA_ROW & ")"
        ws.Cells(7, colIdx).NumberFormat = "#,##0.00"
        ws.Cells(7, colIdx).Font.Bold = True
    Next i

    ' Columns R..X in summary block
    ws.Range("R5:S5").Merge: ws.Range("R5").Value = "Coupler"
    ws.Range("R6:S6").Merge: ws.Range("R6").Value = modUnicode.U("(c\u00e1i)")
    ws.Range("R7:S7").Merge: ws.Range("R7").Formula = "=SUM($T$" & TKT_FIRST_DATA_ROW & ":$T$" & TKT_MAX_DATA_ROW & ")"
    ws.Range("R7").NumberFormat = "#,##0": ws.Range("R7").Font.Bold = True

    ws.Range("T5:V5").Merge: ws.Range("T5").Value = modUnicode.U("T\u1ed4NG C\u1ed8NG (kg)")
    ws.Range("T6:V6").Merge: ws.Range("T6").Value = modUnicode.U("L\u00fd thuy\u1ebft")
    ws.Range("T7:V7").Merge: ws.Range("T7").Formula = "=SUM(D7:Q7)": ws.Range("T7").NumberFormat = "#,##0.00"
    ws.Range("T5:V7").Font.Bold = True
    ws.Range("T7:V7").Interior.Color = RGB(255, 242, 204)

    ws.Range("W5:X5").Merge: ws.Range("W5").Value = modUnicode.U("T\u1ed4NG (T\u1ea5n)")
    ws.Range("W6:X6").Merge: ws.Range("W6").Value = modUnicode.U("L\u00fd thuy\u1ebft")
    ws.Range("W7:X7").Merge: ws.Range("W7").Formula = "=T7/1000": ws.Range("W7").NumberFormat = "#,##0.000"
    ws.Range("W5:X7").Font.Bold = True

    ws.Range("D8:F8").Merge: ws.Range("D8").Formula = "=" & """<= " & ChrW$(216) & "10: """ & "&TEXT(SUM(D7:F7),""#,##0.0"")&"" kg"""
    ws.Range("G8:J8").Merge: ws.Range("G8").Formula = "=""" & ChrW$(216) & "12-" & ChrW$(216) & "18: """ & "&TEXT(SUM(G7:J7),""#,##0.0"")&"" kg"""
    ws.Range("K8:Q8").Merge: ws.Range("K8").Formula = "=" & """> " & ChrW$(216) & "18: """ & "&TEXT(SUM(K7:Q7),""#,##0.0"")&"" kg"""
    ws.Range("R8:X8").Merge: ws.Range("R8").Formula = "=""" & modUnicode.U("C\u00f3 hao h\u1ee5t: ") & """&TEXT(T7*(1+IFERROR(CFG_RebarWastePct,0)/100),""#,##0.0"")&"" kg"""
    ws.Range("D8:X8").Font.Bold = True
    ws.Range("D8:X8").HorizontalAlignment = xlCenter
    ws.Range("D8:X8").Interior.Color = RGB(242, 242, 242)

    With ws.Range("B4:X8").Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(180, 180, 180)
    End With

    ' Detail table header (Row 10, Columns A..X)
    Dim headers As Variant
    headers = Array( _
        "STT", _
        modUnicode.U("T\u00ean c\u1ea5u ki\u1ec7n"), _
        modUnicode.U("S\u1ed1 hi\u1ec7u"), _
        modUnicode.U("M\u00e3 h\u00ecnh"), _
        modUnicode.U("H\u00ecnh d\u1ea1ng thanh th\u00e9p"), _
        "L1 / R", "L2", "L3", "L4", "L5", "L6", "L7", "M1 / M", "M2", _
        ChrW$(216) & " (mm)", _
        modUnicode.U("Thanh/CK"), _
        modUnicode.U("S\u1ed1 CK"), _
        modUnicode.U("T\u1ed5ng thanh"), _
        modUnicode.U("D\u00e0i 1 thanh (m)"), _
        "Coupler", _
        modUnicode.U("T\u1ed5ng d\u00e0i (m)"), _
        "kg/m", _
        modUnicode.U("Kh\u1ed1i l\u01b0\u1ee3ng (kg)"), _
        modUnicode.U("Ghi ch\u00fa") _
    )
    For i = LBound(headers) To UBound(headers)
        ws.Cells(10, i + 1).Value = headers(i)
    Next i
    With ws.Range("A10:X10")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(22, 54, 92)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True: .RowHeight = 30
    End With

    ' Diameter dropdown validation
    With ws.Range("O" & TKT_FIRST_DATA_ROW & ":O" & TKT_MAX_DATA_ROW).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="6,8,10,12,14,16,18,20,22,25,28,32,36,40"
        .InCellDropdown = True
    End With

    ' Column widths
    Dim widths As Variant
    widths = Array(5.5, 18, 8, 7.5, 28, 7.5, 7.5, 7.5, 7.5, 7.5, 7.5, 7.5, 8, 7.5, 7.5, 9, 7.5, 9.5, 12, 8, 11, 8, 13, 18)
    For i = LBound(widths) To UBound(widths)
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i

    ' Seed 1 clean sample row (Row 11: D07 U-bar)
    InsertBarRowInternal ws, TKT_FIRST_DATA_ROW, "D07", titleVI & " - M1", "1", 16, 250, 2400, 250, 0, 0, 0, 0, 0, 0, 6, 2

    ws.Activate
    On Error Resume Next
    If Not Application.ActiveWindow Is Nothing Then
        Application.ActiveWindow.SplitRow = 10
        Application.ActiveWindow.FreezePanes = True
    End If
    On Error GoTo 0
End Sub

Public Sub ShowBarShapePicker()
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    If Not IsTKTSheet(ActiveSheet) And Not modRebarConfig.IsRebarConfigSheet(ActiveSheet) Then
        modUnicode.MsgBoxU modUnicode.U("Vui l\u00f2ng \u0111\u1ee9ng trong sheet QS_RebarConfig ho\u1eb7c sheet TKT_* \u0111\u1ec3 ch\u1ecdn h\u00ecnh d\u1ea1ng th\u00e9p!"), vbExclamation
        Exit Sub
    End If
    frmBarShape.Show vbModeless
    SyncPickerWithActiveRow
End Sub

Public Sub SyncPickerWithActiveRow()
    On Error Resume Next
    Dim isOpen As Boolean, uf As Object
    For Each uf In VBA.UserForms
        If uf.Name = "frmBarShape" Then isOpen = True: Exit For
    Next uf
    If Not isOpen Then Exit Sub
    If ActiveSheet Is Nothing Then Exit Sub
    Dim shapeCol As Long, firstR As Long
    If IsTKTSheet(ActiveSheet) Then
        shapeCol = 4: firstR = TKT_FIRST_DATA_ROW
    ElseIf modRebarConfig.IsRebarConfigSheet(ActiveSheet) Then
        shapeCol = modRebarConfig.REBAR_CFG_SHAPE_COL: firstR = modRebarConfig.REBAR_CFG_FIRST_ROW
    Else
        Exit Sub
    End If
    Dim r As Long: r = ActiveCell.Row
    If r < firstR Then Exit Sub
    Dim sc As String
    sc = UCase$(Trim$(CStr(ActiveSheet.Cells(r, shapeCol).Value)))
    If Left$(sc, 1) = "D" And Len(sc) >= 2 Then
        frmBarShape.OnShapeClicked sc
    End If
    On Error GoTo 0
End Sub

Public Sub InsertBarRowFromUI(ByVal shapeCode As String, ByVal insertAtCursor As Boolean)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If modRebarConfig.IsRebarConfigSheet(ws) Then
        modRebarConfig.InsertCustomBarRowOnConfig shapeCode, insertAtCursor
        Exit Sub
    End If
    If Not IsTKTSheet(ws) Then Exit Sub

    Dim spec As Variant
    spec = GetShapeSpec(shapeCode)
    If Not IsArray(spec) Then Exit Sub

    Dim targetRow As Long
    If insertAtCursor And ActiveCell.Row >= TKT_FIRST_DATA_ROW Then
        targetRow = ActiveCell.Row
        ws.Rows(targetRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Else
        targetRow = GetNextTKTRow(ws)
    End If

    Dim prevComp As String, prevMark As String
    If targetRow > TKT_FIRST_DATA_ROW Then
        prevComp = CStr(ws.Cells(targetRow - 1, 2).Value)
        prevMark = CStr(ws.Cells(targetRow - 1, 3).Value)
    End If
    If Len(prevComp) = 0 Then prevComp = "CK1"
    If IsNumeric(prevMark) Then
        prevMark = CStr(CLng(prevMark) + 1)
    Else
        prevMark = "1"
    End If

    Application.ScreenUpdating = False
    InsertBarRowInternal ws, targetRow, shapeCode, prevComp, prevMark, 14, _
        Val(spec(4)), Val(spec(5)), Val(spec(6)), Val(spec(7)), Val(spec(8)), _
        Val(spec(9)), Val(spec(10)), Val(spec(11)), Val(spec(12)), 4, 1
    ws.Cells(targetRow, 6).Select
    Application.ScreenUpdating = True
End Sub

Public Sub ApplyShapeToActiveRow(ByVal shapeCode As String)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    If modRebarConfig.IsRebarConfigSheet(ws) Then
        modRebarConfig.ApplyShapeToConfigRow shapeCode
        Exit Sub
    End If
    If Not IsTKTSheet(ws) Then Exit Sub
    Dim r As Long: r = ActiveCell.Row
    If r < TKT_FIRST_DATA_ROW Then Exit Sub

    Dim spec As Variant
    spec = GetShapeSpec(shapeCode)
    If Not IsArray(spec) Then Exit Sub

    Dim compName As String, barMark As String
    Dim dia As Double, qEl As Double, qCk As Double
    compName = CStr(ws.Cells(r, 2).Value)
    barMark = CStr(ws.Cells(r, 3).Value)
    dia = Val(ws.Cells(r, 15).Value): If dia <= 0 Then dia = 14
    qEl = Val(ws.Cells(r, 16).Value): If qEl <= 0 Then qEl = 4
    qCk = Val(ws.Cells(r, 17).Value): If qCk <= 0 Then qCk = 1

    Application.ScreenUpdating = False
    InsertBarRowInternal ws, r, shapeCode, compName, barMark, dia, _
        Val(spec(4)), Val(spec(5)), Val(spec(6)), Val(spec(7)), Val(spec(8)), _
        Val(spec(9)), Val(spec(10)), Val(spec(11)), Val(spec(12)), qEl, qCk
    Application.ScreenUpdating = True
End Sub

Public Sub InsertBarRowInternal(ByVal ws As Worksheet, ByVal r As Long, ByVal shapeCode As String, _
                                ByVal compName As String, ByVal barMark As String, ByVal dia As Double, _
                                ByVal l1 As Double, ByVal l2 As Double, ByVal l3 As Double, _
                                ByVal l4 As Double, ByVal l5 As Double, ByVal l6 As Double, _
                                ByVal l7 As Double, ByVal m1 As Double, ByVal m2 As Double, _
                                ByVal qtyPerEl As Double, ByVal qtyEl As Double)
    Dim spec As Variant
    spec = GetShapeSpec(shapeCode)
    If Not IsArray(spec) Then Exit Sub

    Dim bendCnt As Long, hasLap As Long
    Dim lenExpr As String, activeCols As String
    bendCnt = CLng(spec(13))
    hasLap = CLng(spec(14))
    lenExpr = Replace(CStr(spec(15)), "{r}", CStr(r))
    activeCols = "|" & UCase$(CStr(spec(16))) & "|"

    ws.Cells(r, 2).Value = compName
    ws.Cells(r, 3).Value = barMark
    ws.Cells(r, 4).Value = UCase$(shapeCode)
    ws.Cells(r, 4).HorizontalAlignment = xlCenter
    ws.Cells(r, 4).Font.Bold = True

    ' Descriptive text cue in Column E (always visible in Light Mode 0-Shape)
    ws.Cells(r, 5).Value = UCase$(shapeCode) & " [" & CStr(spec(3)) & "]"
    ws.Cells(r, 5).Font.Size = 8.5
    ws.Cells(r, 5).Font.Color = RGB(90, 90, 90)
    ws.Cells(r, 5).HorizontalAlignment = xlCenter

    Dim vals As Variant
    vals = Array(l1, l2, l3, l4, l5, l6, l7, m1, m2)
    Dim cIdx As Long, colLet As String
    For cIdx = 6 To 14
        colLet = Split(ws.Cells(1, cIdx).Address(True, False), "$")(0)
        If InStr(1, activeCols, "|" & colLet & "|", vbTextCompare) > 0 Then
            ws.Cells(r, cIdx).Value = vals(cIdx - 6)
            ws.Cells(r, cIdx).Interior.Color = RGB(255, 250, 220)
        Else
            ws.Cells(r, cIdx).Value = ""
            ws.Cells(r, cIdx).Interior.Color = RGB(242, 242, 242)
        End If
    Next cIdx

    ws.Cells(r, 15).Value = dia
    ws.Cells(r, 16).Value = qtyPerEl
    ws.Cells(r, 17).Value = qtyEl
    ws.Range(ws.Cells(r, 15), ws.Cells(r, 17)).Interior.Color = RGB(255, 250, 220)

    ' Formulas (100% Excel 2007-365 compatible)
    ws.Cells(r, 18).Formula = "=IF(O" & r & "="""","""",P" & r & "*Q" & r & ")"

    Dim geomLen As String
    If bendCnt > 0 Then
        geomLen = "((" & lenExpr & ")-" & bendCnt & "*IFERROR(CFG_BendDeduct,1)*O" & r & ")"
    Else
        geomLen = "(" & lenExpr & ")"
    End If

    If hasLap = 1 Then
        ws.Cells(r, 19).Formula = "=IF(O" & r & "="""","""",ROUND((" & geomLen & _
            "+IF(AND(O" & r & ">=10,O" & r & "<IFERROR(CFG_CouplerMinD,22))," & _
            "MAX(0,CEILING(" & geomLen & "/IFERROR(CFG_BarStdLen,11700),1)-1)*IFERROR(CFG_LapRatio,40)*O" & r & ",0))/1000,3))"
        ws.Cells(r, 20).Formula = "=IF(OR(O" & r & "="""",S" & r & "=""""),0," & _
            "IF(O" & r & ">=IFERROR(CFG_CouplerMinD,22),R" & r & "*MAX(0,CEILING(" & geomLen & "/IFERROR(CFG_BarStdLen,11700),1)-1),0))"
    Else
        ws.Cells(r, 19).Formula = "=IF(O" & r & "="""","""",ROUND(" & geomLen & "/1000,3))"
        ws.Cells(r, 20).Value = 0
    End If

    ws.Cells(r, 21).Formula = "=IF(O" & r & "="""","""",ROUND(R" & r & "*S" & r & ",2))"
    ws.Cells(r, 22).Formula = "=IF(O" & r & "="""","""",ROUND(PI()*(O" & r & "/1000)^2/4*7850,3))"
    ws.Cells(r, 23).Formula = "=IF(O" & r & "="""","""",ROUND(U" & r & "*V" & r & ",2))"
    ws.Cells(r, 24).Value = CStr(spec(1))

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 24))
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(191, 191, 191)
    End With
    ws.Range(ws.Cells(r, 6), ws.Cells(r, 18)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r, 6), ws.Cells(r, 18)).NumberFormat = "#,##0"
    ws.Cells(r, 19).NumberFormat = "#,##0.000"
    ws.Cells(r, 20).NumberFormat = "#,##0"
    ws.Cells(r, 21).NumberFormat = "#,##0.00"
    ws.Cells(r, 22).NumberFormat = "0.000"
    ws.Cells(r, 23).NumberFormat = "#,##0.00"
    ws.Cells(r, 23).Font.Bold = True

    ' Smart Performance Rule: Auto-draw labeled image only when sheet has <= 250 rows
    Dim dataCount As Long
    dataCount = GetLastTKTRow(ws) - TKT_FIRST_DATA_ROW + 1
    If dataCount <= TKT_AUTO_SHAPE_LIMIT Then
        ws.Rows(r).RowHeight = 42
        modShapeDraw.DrawBarShapeAtRow ws, r, shapeCode
    Else
        ws.Rows(r).RowHeight = 22
    End If
    RenumberTKT ws
End Sub

Public Sub DuplicateBarRow()
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not IsTKTSheet(ws) Then Exit Sub
    Dim r As Long: r = ActiveCell.Row
    If r < TKT_FIRST_DATA_ROW Then Exit Sub

    Application.ScreenUpdating = False
    ws.Rows(r).Copy
    ws.Rows(r + 1).Insert Shift:=xlDown
    Application.CutCopyMode = False
    If (GetLastTKTRow(ws) - TKT_FIRST_DATA_ROW + 1) <= TKT_AUTO_SHAPE_LIMIT Then
        modShapeDraw.DrawBarShapeAtRow ws, r + 1, CStr(ws.Cells(r + 1, 4).Value)
    End If
    RenumberTKT ws
    ws.Cells(r + 1, 6).Select
    Application.ScreenUpdating = True
End Sub

Public Sub MoveBarRowUp()
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not IsTKTSheet(ws) Then Exit Sub
    Dim r As Long: r = ActiveCell.Row
    If r <= TKT_FIRST_DATA_ROW Then Exit Sub
    SwapTKTRows ws, r, r - 1
    ws.Cells(r - 1, ActiveCell.Column).Select
End Sub

Public Sub MoveBarRowDown()
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not IsTKTSheet(ws) Then Exit Sub
    Dim r As Long: r = ActiveCell.Row
    Dim lr As Long: lr = GetLastTKTRow(ws)
    If r < TKT_FIRST_DATA_ROW Or r >= lr Then Exit Sub
    SwapTKTRows ws, r, r + 1
    ws.Cells(r + 1, ActiveCell.Column).Select
End Sub

Private Sub SwapTKTRows(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long)
    Application.ScreenUpdating = False
    Dim sc1 As String, sc2 As String
    sc1 = CStr(ws.Cells(r1, 4).Value)
    sc2 = CStr(ws.Cells(r2, 4).Value)

    Dim tmpRow As Long
    tmpRow = ws.Rows.Count
    ws.Rows(r1).Copy ws.Rows(tmpRow)
    ws.Rows(r2).Copy ws.Rows(r1)
    ws.Rows(tmpRow).Copy ws.Rows(r2)
    ws.Rows(tmpRow).Clear
    Application.CutCopyMode = False

    If ws.Shapes.Count > 0 Then
        modShapeDraw.DrawBarShapeAtRow ws, r1, sc2
        modShapeDraw.DrawBarShapeAtRow ws, r2, sc1
    End If
    RenumberTKT ws
    Application.ScreenUpdating = True
End Sub

Public Sub DeleteBarRow()
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not IsTKTSheet(ws) Then Exit Sub
    Dim r As Long: r = ActiveCell.Row
    If r < TKT_FIRST_DATA_ROW Or r > GetLastTKTRow(ws) Then Exit Sub

    Application.ScreenUpdating = False
    modShapeDraw.DeleteBarShapeAtRow ws, r
    ws.Rows(r).Delete
    RenumberTKT ws
    Application.ScreenUpdating = True
End Sub

Public Sub RedrawActiveSheetShapes(Optional ByVal askIfLarge As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not IsTKTSheet(ws) Then Exit Sub
    Dim rowCnt As Long
    rowCnt = GetLastTKTRow(ws) - TKT_FIRST_DATA_ROW + 1
    If askIfLarge And rowCnt > 800 Then
        Dim ans As Long
        ans = modUnicode.MsgBoxU(modUnicode.U("B\u1ea3ng hi\u1ec7n c\u00f3 ") & rowCnt & modUnicode.U(" d\u00f2ng. V\u1ebd h\u00ecnh l\u00ean qu\u00e1 nhi\u1ec1u d\u00f2ng s\u1ebd l\u00e0m Excel n\u1eb7ng h\u01a1n khi cu\u1ed9n. B\u1ea1n c\u00f3 mu\u1ed1n v\u1ebd ti\u1ebfp kh\u00f4ng?"), vbYesNo + vbQuestion)
        If ans = vbNo Then Exit Sub
    End If
    Application.ScreenUpdating = False
    modShapeDraw.RedrawAllBarShapes ws, TKT_FIRST_DATA_ROW
    RenumberTKT ws
    Application.ScreenUpdating = True
End Sub

Public Sub ClearActiveSheetShapes()
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not IsTKTSheet(ws) Then Exit Sub
    Application.ScreenUpdating = False
    modShapeDraw.ClearAllBarShapes ws
    Dim lr As Long: lr = GetLastTKTRow(ws)
    If lr >= TKT_FIRST_DATA_ROW Then
        ws.Rows(TKT_FIRST_DATA_ROW & ":" & lr).RowHeight = 22
    End If
    Application.ScreenUpdating = True
End Sub

Public Sub RenumberTKT(ByVal ws As Worksheet)
    Dim lr As Long: lr = GetLastTKTRow(ws)
    Dim r As Long, n As Long
    n = 0
    For r = TKT_FIRST_DATA_ROW To lr
        If Trim$(CStr(ws.Cells(r, 4).Value)) <> "" Or Trim$(CStr(ws.Cells(r, 15).Value)) <> "" Then
            n = n + 1
            ws.Cells(r, 1).Value = n
            ws.Cells(r, 1).HorizontalAlignment = xlCenter
        End If
    Next r
End Sub

Public Function GetLastTKTRow(ByVal ws As Worksheet) As Long
    Dim r1 As Long, r2 As Long
    r1 = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    r2 = ws.Cells(ws.Rows.Count, 15).End(xlUp).Row
    GetLastTKTRow = IIf(r1 > r2, r1, r2)
    If GetLastTKTRow < TKT_FIRST_DATA_ROW - 1 Then GetLastTKTRow = TKT_FIRST_DATA_ROW - 1
End Function

Public Function GetNextTKTRow(ByVal ws As Worksheet) As Long
    GetNextTKTRow = GetLastTKTRow(ws) + 1
    If GetNextTKTRow < TKT_FIRST_DATA_ROW Then GetNextTKTRow = TKT_FIRST_DATA_ROW
End Function

Public Function GetAllShapesTable() As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("DB_BarShapes")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < 2 Then Exit Function
    GetAllShapesTable = ws.Range(ws.Cells(2, 1), ws.Cells(lr, 17)).Value
End Function

Public Function GetShapeSpec(ByVal shapeCode As String) As Variant
    Dim tbl As Variant
    tbl = GetAllShapesTable()
    If Not IsArray(tbl) Then Exit Function
    Dim r As Long, c As Long
    For r = 1 To UBound(tbl, 1)
        If StrComp(Trim$(CStr(tbl(r, 1))), Trim$(shapeCode), vbTextCompare) = 0 Then
            Dim res(0 To 16) As Variant
            For c = 1 To 17
                res(c - 1) = tbl(r, c)
            Next c
            GetShapeSpec = res
            Exit Function
        End If
    Next r
End Function
