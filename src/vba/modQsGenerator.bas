Attribute VB_Name = "modQsGenerator"
Option Explicit

' ==============================================================================
' modQsGenerator - Orchestrates 1-Click Generation & Non-Destructive Sync
' Writes Detail Sheets (*_Detail, Finishing_Detail, Civil_Detail) & TKT_*
' Uses Bulk 2D Variant Array + 3-Level Outline Grouping (< 2 sec for thousands)
' ==============================================================================

Public LastGenError As String
Public LastGenStep As String

Public Function GetLastGenError() As String
    GetLastGenError = LastGenError
End Function

Public Sub GenerateOrSyncAllStructure(Optional ByVal showMsg As Boolean = True, Optional ByVal targetBuilding As String = "")
    LastGenError = ""
    LastGenStep = "Start"
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook

    Dim wsData As Worksheet
    On Error Resume Next
    Set wsData = wb.Worksheets(modQsData.QSDATA_SHEET)
    On Error GoTo 0
    If wsData Is Nothing Then
        modQsData.CreateOrShowQsDataSheet False
        Set wsData = wb.Worksheets(modQsData.QSDATA_SHEET)
    End If

    Dim errMsg As String
    If Not modQsData.ValidateQsData(wsData, errMsg) Then
        If showMsg Then modUnicode.MsgBoxU errMsg, vbExclamation
        Exit Sub
    End If

    Dim hasSpec As Boolean
    On Error Resume Next
    hasSpec = Not (wb.Worksheets(modSpecSheet.SPEC_SHEET) Is Nothing)
    On Error GoTo 0
    If Not hasSpec Then modSpecSheet.CreateOrShowSpecSheet False

    ' Ensure QS_RebarConfig exists and auto-expand any newly added elements from QS_Data
    modRebarConfig.EnsureRebarConfigSheet wb, True
    modRebarConfig.ExpandRebarPresetsFromQsData False
    Dim wsCfg As Worksheet
    On Error Resume Next
    Set wsCfg = wb.Worksheets(modRebarConfig.REBAR_CFG_SHEET)
    On Error GoTo 0

    ' Collect existing Detail & TKT manual/element rows if targetBuilding specified or across workbook
    Dim exDetailJson As String, exTktJson As String, rebarCfgJson As String
    exDetailJson = CollectExistingDetailManualRowsJson(wb, targetBuilding)
    exTktJson = CollectExistingTktManualRowsJson(wb, targetBuilding)
    If Not wsCfg Is Nothing Then
        rebarCfgJson = modRebarConfig.SerializeRebarConfigsToJson(wsCfg)
    Else
        rebarCfgJson = "[]"
    End If

    Dim reqJson As String
    reqJson = "{" & _
        """targetBuilding"":" & modEngineBridge.JsonStr(targetBuilding) & "," & _
        """spec"":" & modSpecSheet.SerializeSpecToJson(wb) & "," & _
        """elements"":" & modQsData.SerializeQsDataToJson(wsData, targetBuilding) & "," & _
        """rebarConfigs"":" & rebarCfgJson & "," & _
        """existingDetailRows"":" & exDetailJson & "," & _
        """existingTktRows"":" & exTktJson & _
        "}"

    Dim resJson As String
    If Not modEngineBridge.RunEngineCommand("generate_structure", reqJson, resJson, errMsg) Then
        If showMsg Then modUnicode.MsgBoxU errMsg, vbCritical
        Exit Sub
    End If

    Dim root As Object
    Set root = modEngineBridge.JsonParseObject(resJson)
    Dim dataObj As Object
    Set dataObj = root("data")

    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Dim detCols As Collection, tktCols As Collection
    Set detCols = dataObj("detailSheets")
    Set tktCols = dataObj("tktSheets")

    Dim totalAdded As Long, totalUpdated As Long, totalManual As Long, totalRemoved As Long
    Dim i As Long
    LastGenStep = "DetailLoop"
    For i = 1 To detCols.Count
        LastGenStep = "DetailItem_" & i
        Dim dsObj As Object: Set dsObj = detCols(i)
        WriteDetailSheetFromDto wb, dsObj
        totalAdded = totalAdded + CLng(dsObj("addedCount"))
        totalUpdated = totalUpdated + CLng(dsObj("updatedCount"))
        totalManual = totalManual + CLng(dsObj("preservedManualCount"))
        totalRemoved = totalRemoved + CLng(dsObj("removedFlaggedCount"))
    Next i

    LastGenStep = "TktLoop"
    For i = 1 To tktCols.Count
        LastGenStep = "TktItem_" & i
        Dim tsObj As Object: Set tsObj = tktCols(i)
        WriteTktSheetFromDto wb, tsObj
    Next i

    LastGenStep = "Summary"
    modRebarSummary.SummarizeAllSteel False
    wb.Activate

    Application.ScreenUpdating = True
    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 sinh/\u0111\u1ed3ng b\u1ed9 th\u00e0nh c\u00f4ng ") & detCols.Count & modUnicode.U(" sheet Detail & ") & tktCols.Count & modUnicode.U(" sheet TKT_*!") & vbCrLf & _
            "- Added: " & totalAdded & " | Updated: " & totalUpdated & " | Preserved Manual: " & totalManual & " | Removed Flagged: " & totalRemoved, vbInformation
    End If
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    LastGenError = "GenerateOrSyncAllStructure Error [" & LastGenStep & "]: " & Err.Description
    If showMsg Then
        modUnicode.MsgBoxU LastGenError, vbCritical
    End If
End Sub

Public Sub WriteDetailSheetFromDto(ByVal wb As Workbook, ByVal sheetDto As Object)
    Dim sName As String: sName = CStr(sheetDto("sheetName"))
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sName
    Else
        ws.Cells.Clear
    End If

    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 9.5

    LastGenStep = sName & ":ReadRows"
    Dim rowsCol As Collection
    Set rowsCol = sheetDto("rows")
    Dim nRows As Long: nRows = rowsCol.Count
    If nRows = 0 Then Exit Sub

    Dim nCols As Long: nCols = 14
    Dim arr() As Variant
    ReDim arr(1 To nRows, 1 To nCols)

    Dim rIdx As Long, cIdx As Long
    Dim rObj As Object, vals As Collection, v As Variant
    For rIdx = 1 To nRows
        LastGenStep = sName & ":RowObj_" & rIdx
        Set rObj = rowsCol(rIdx)
        Set vals = rObj("values")
        For cIdx = 1 To nCols
            If cIdx <= vals.Count Then
                v = vals(cIdx)
                If IsEmpty(v) Or IsNull(v) Then
                    arr(rIdx, cIdx) = Empty
                ElseIf VarType(v) = vbString And Len(CStr(v)) = 0 Then
                    arr(rIdx, cIdx) = ""
                Else
                    arr(rIdx, cIdx) = v
                End If
            End If
        Next cIdx
    Next rIdx

    LastGenStep = sName & ":WriteFormulaArr"
    ws.Range(ws.Cells(1, 1), ws.Cells(nRows, nCols)).Formula = arr

    ' Apply Styling, NumberFormat & 3-Level Outline Grouping
    On Error Resume Next
    ws.Outline.SummaryRow = xlAbove
    On Error GoTo 0

    Dim kind As String, ol As Long
    For rIdx = 1 To nRows
        Set rObj = rowsCol(rIdx)
        kind = CStr(rObj("rowKind"))
        ol = CLng(rObj("outlineLevel"))
        Select Case kind
            Case "TITLE"
                With ws.Range("B1:L1")
                    .Merge
                    .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(22, 54, 92): .RowHeight = 28: .VerticalAlignment = xlCenter
                End With
                ws.Cells(1, 1).Font.Color = RGB(22, 54, 92)
                ws.Cells(1, 1).Interior.Color = RGB(22, 54, 92)
            Case "BANNER"
                ws.Range(ws.Cells(rIdx, 2), ws.Cells(rIdx, 12)).Interior.Color = RGB(237, 242, 248)
                ws.Cells(rIdx, 3).Font.Bold = True
            Case "WORK_HDR"
                With ws.Range(ws.Cells(rIdx, 1), ws.Cells(rIdx, 12))
                    .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(47, 85, 151): .RowHeight = 24
                End With
                ws.Range(ws.Cells(rIdx, 13), ws.Cells(rIdx, 14)).Interior.Color = RGB(47, 85, 151)
            Case "ITEM_HDR"
                With ws.Range(ws.Cells(rIdx, 1), ws.Cells(rIdx, 12))
                    .Font.Bold = True: .Interior.Color = RGB(217, 225, 242): .RowHeight = 21
                End With
            Case "SUB_HDR"
                With ws.Range(ws.Cells(rIdx, 1), ws.Cells(rIdx, 12))
                    .Font.Bold = True: .Font.Italic = True: .Interior.Color = RGB(242, 242, 242)
                End With
            Case "MANUAL"
                ws.Range(ws.Cells(rIdx, 1), ws.Cells(rIdx, 12)).Interior.Color = RGB(235, 245, 255)
            Case "REMOVED"
                ws.Range(ws.Cells(rIdx, 1), ws.Cells(rIdx, 12)).Interior.Color = RGB(255, 242, 204)
                ws.Range(ws.Cells(rIdx, 1), ws.Cells(rIdx, 12)).Font.Color = RGB(192, 0, 0)
        End Select

        If kind = "ITEM_HDR" Or kind = "SUB_HDR" Or kind = "DATA" Or kind = "MANUAL" Then
            ws.Range(ws.Cells(rIdx, 5), ws.Cells(rIdx, 12)).NumberFormat = "#,##0.00"
            ws.Cells(rIdx, 5).NumberFormat = "#,##0.##"
            ws.Cells(rIdx, 12).Font.Bold = True
        End If
        If ol > 0 And rIdx >= 5 Then
            On Error Resume Next
            ws.Rows(rIdx).OutlineLevel = ol + 1
            On Error GoTo 0
        End If
    Next rIdx

    With ws.Range(ws.Cells(4, 1), ws.Cells(nRows, 12)).Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(208, 208, 208)
    End With
    ' Subtle Tag metadata columns M..N
    ws.Range(ws.Cells(1, 13), ws.Cells(nRows, 14)).Font.Color = RGB(140, 140, 140)
    ws.Range(ws.Cells(1, 13), ws.Cells(nRows, 14)).Font.Size = 8

    Dim widths As Variant
    widths = Array(10, 46, 22, 7, 8, 9, 9, 9, 8, 8, 8, 13, 18, 12)
    Dim c As Long
    For c = 0 To UBound(widths)
        ws.Columns(c + 1).ColumnWidth = widths(c)
    Next c

    ws.Activate
    On Error Resume Next
    If Not Application.ActiveWindow Is Nothing Then
        Application.ActiveWindow.SplitRow = 4
        Application.ActiveWindow.FreezePanes = True
        Application.ActiveWindow.DisplayGridlines = True
    End If
    On Error GoTo 0
End Sub

Public Sub WriteTktSheetFromDto(ByVal wb As Workbook, ByVal tktDto As Object)
    Dim sName As String: sName = CStr(tktDto("sheetName"))
    Dim bld As String: bld = CStr(tktDto("building"))
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sName)
    On Error GoTo 0

    If ws Is Nothing Then
        modRebarTKT.CreateTKTSheet sName, bld, False
        ' Note: CreateTKTSheet handles unique naming; grab active or exact
        On Error Resume Next
        Set ws = wb.Worksheets(sName)
        On Error GoTo 0
        If ws Is Nothing Then Set ws = wb.ActiveSheet
    End If

    ' Clear old shapes and data rows 11..5000
    modShapeDraw.ClearAllBarShapes ws
    Dim oldLast As Long: oldLast = modRebarTKT.GetLastTKTRow(ws)
    If oldLast >= modRebarTKT.TKT_FIRST_DATA_ROW Then
        ws.Range("A" & modRebarTKT.TKT_FIRST_DATA_ROW & ":Z" & oldLast).Clear
    End If

    ws.Cells(10, 25).Value = "TagID"
    ws.Cells(10, 26).Value = "BOQ_Bucket"
    ws.Range("Y10:Z10").Font.Bold = True
    ws.Range("Y10:Z10").Font.Color = RGB(200, 200, 200)
    ws.Range("Y10:Z10").Interior.Color = RGB(22, 54, 92)

    Dim rowsCol As Collection
    Set rowsCol = tktDto("rows")
    Dim nRows As Long: nRows = rowsCol.Count
    If nRows = 0 Then Exit Sub

    Dim nCols As Long: nCols = 26
    Dim arr() As Variant
    ReDim arr(1 To nRows, 1 To nCols)

    Dim rIdx As Long, cIdx As Long
    Dim rObj As Object, vals As Collection, v As Variant
    For rIdx = 1 To nRows
        Set rObj = rowsCol(rIdx)
        Set vals = rObj("values")
        For cIdx = 1 To nCols
            If cIdx <= vals.Count Then
                v = vals(cIdx)
                If IsEmpty(v) Or IsNull(v) Then
                    arr(rIdx, cIdx) = Empty
                Else
                    arr(rIdx, cIdx) = v
                End If
            End If
        Next cIdx
    Next rIdx

    Dim startR As Long: startR = modRebarTKT.TKT_FIRST_DATA_ROW
    Dim endR As Long: endR = startR + nRows - 1
    ws.Range(ws.Cells(startR, 1), ws.Cells(endR, nCols)).Formula = arr

    ' Format TKT data block cleanly
    With ws.Range(ws.Cells(startR, 1), ws.Cells(endR, 24))
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(191, 191, 191)
    End With
    ws.Range(ws.Cells(startR, 1), ws.Cells(endR, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(startR, 4), ws.Cells(endR, 4)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(startR, 4), ws.Cells(endR, 4)).Font.Bold = True
    ws.Range(ws.Cells(startR, 6), ws.Cells(endR, 18)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(startR, 6), ws.Cells(endR, 18)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(startR, 19), ws.Cells(endR, 19)).NumberFormat = "#,##0.000"
    ws.Range(ws.Cells(startR, 20), ws.Cells(endR, 20)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(startR, 21), ws.Cells(endR, 21)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(startR, 22), ws.Cells(endR, 22)).NumberFormat = "0.000"
    ws.Range(ws.Cells(startR, 23), ws.Cells(endR, 23)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(startR, 23), ws.Cells(endR, 23)).Font.Bold = True
    ws.Range(ws.Cells(startR, 25), ws.Cells(endR, 26)).Font.Color = RGB(150, 150, 150)
    ws.Range(ws.Cells(startR, 25), ws.Cells(endR, 26)).Font.Size = 8
    ws.Columns("Y:Z").ColumnWidth = 14

    For rIdx = 1 To nRows
        Set rObj = rowsCol(rIdx)
        If CStr(rObj("rowKind")) = "MANUAL" Then
            ws.Range(ws.Cells(startR + rIdx - 1, 1), ws.Cells(startR + rIdx - 1, 24)).Interior.Color = RGB(235, 245, 255)
        ElseIf CStr(rObj("rowKind")) = "REMOVED" Then
            ws.Range(ws.Cells(startR + rIdx - 1, 1), ws.Cells(startR + rIdx - 1, 24)).Interior.Color = RGB(255, 242, 204)
        End If
    Next rIdx

    ' Draw shapes only when within smart limit
    If nRows <= modRebarTKT.TKT_AUTO_SHAPE_LIMIT Then
        modShapeDraw.RedrawAllBarShapes ws, startR
    Else
        ws.Rows(startR & ":" & endR).RowHeight = 22
    End If
End Sub

Public Function CollectExistingDetailManualRowsJson(ByVal wb As Workbook, ByVal targetBuilding As String, Optional ByVal exactSheetName As String = "") As String
    Dim items() As String, cnt As Long, cap As Long
    cap = 32: cnt = 0
    ReDim items(0 To cap - 1)

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim matchWs As Boolean: matchWs = False
        If Len(exactSheetName) > 0 Then
            matchWs = (StrComp(ws.Name, exactSheetName, vbTextCompare) = 0)
        Else
            matchWs = (Right$(ws.Name, 7) = "_Detail") And (StrComp(ws.Name, "Finishing_Detail", vbTextCompare) <> 0) And (StrComp(ws.Name, "Civil_Detail", vbTextCompare) <> 0)
            If matchWs And Len(targetBuilding) > 0 Then
                matchWs = (InStr(1, ws.Name, Left$(targetBuilding, 20), vbTextCompare) > 0)
            End If
        End If
        If matchWs Then
            Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            If lr >= 6 Then
                Dim arr As Variant: arr = ws.Range(ws.Cells(1, 1), ws.Cells(lr, 14)).Formula
                Dim r As Long, tagVal As String, secVal As String
                For r = 6 To lr
                    tagVal = Trim$(CStr(arr(r, 13)))
                    secVal = Trim$(CStr(arr(r, 14)))
                    If StrComp(tagVal, "MANUAL", vbTextCompare) = 0 Or InStr(1, tagVal, "::", vbTextCompare) > 0 Then
                        Dim s As String
                        s = "{""rowIndex"":" & r & ","
                        s = s & """code"":" & modEngineBridge.JsonStr(arr(r, 1)) & ","
                        s = s & """description"":" & modEngineBridge.JsonStr(arr(r, 2)) & ","
                        s = s & """specText"":" & modEngineBridge.JsonStr(arr(r, 3)) & ","
                        s = s & """unit"":" & modEngineBridge.JsonStr(arr(r, 4)) & ","
                        s = s & """noExpr"":" & modEngineBridge.JsonStr(arr(r, 5)) & ","
                        s = s & """lExpr"":" & modEngineBridge.JsonStr(arr(r, 6)) & ","
                        s = s & """wExpr"":" & modEngineBridge.JsonStr(arr(r, 7)) & ","
                        s = s & """hExpr"":" & modEngineBridge.JsonStr(arr(r, 8)) & ","
                        s = s & """aExpr"":" & modEngineBridge.JsonStr(arr(r, 9)) & ","
                        s = s & """bExpr"":" & modEngineBridge.JsonStr(arr(r, 10)) & ","
                        s = s & """cExpr"":" & modEngineBridge.JsonStr(arr(r, 11)) & ","
                        s = s & """qtyExpr"":" & modEngineBridge.JsonStr(arr(r, 12)) & ","
                        s = s & """tagId"":" & modEngineBridge.JsonStr(tagVal) & ","
                        s = s & """sectionKey"":" & modEngineBridge.JsonStr(secVal) & "}"
                        If cnt >= cap Then cap = cap * 2: ReDim Preserve items(0 To cap - 1)
                        items(cnt) = s: cnt = cnt + 1
                    End If
                Next r
            End If
        End If
    Next ws

    If cnt = 0 Then
        CollectExistingDetailManualRowsJson = "[]"
    Else
        ReDim Preserve items(0 To cnt - 1)
        CollectExistingDetailManualRowsJson = "[" & Join(items, ",") & "]"
    End If
End Function

Private Function CollectExistingTktManualRowsJson(ByVal wb As Workbook, ByVal targetBuilding As String) As String
    Dim items() As String, cnt As Long, cap As Long
    cap = 32: cnt = 0
    ReDim items(0 To cap - 1)

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If modRebarTKT.IsTKTSheet(ws) Then
            Dim lr As Long: lr = modRebarTKT.GetLastTKTRow(ws)
            If lr >= modRebarTKT.TKT_FIRST_DATA_ROW Then
                Dim arr As Variant: arr = ws.Range(ws.Cells(modRebarTKT.TKT_FIRST_DATA_ROW, 1), ws.Cells(lr, 26)).Formula
                Dim i As Long, r As Long, tagVal As String
                For i = 1 To UBound(arr, 1)
                    r = modRebarTKT.TKT_FIRST_DATA_ROW + i - 1
                    tagVal = Trim$(CStr(arr(i, 25)))
                    If StrComp(tagVal, "MANUAL", vbTextCompare) = 0 Or InStr(1, tagVal, "::", vbTextCompare) > 0 Then
                        Dim s As String
                        s = "{""rowIndex"":" & r & ","
                        s = s & """compName"":" & modEngineBridge.JsonStr(arr(i, 2)) & ","
                        s = s & """barMark"":" & modEngineBridge.JsonStr(arr(i, 3)) & ","
                        s = s & """shapeCode"":" & modEngineBridge.JsonStr(arr(i, 4)) & ","
                        s = s & """l1Expr"":" & modEngineBridge.JsonStr(arr(i, 6)) & ","
                        s = s & """l2Expr"":" & modEngineBridge.JsonStr(arr(i, 7)) & ","
                        s = s & """l3Expr"":" & modEngineBridge.JsonStr(arr(i, 8)) & ","
                        s = s & """l4Expr"":" & modEngineBridge.JsonStr(arr(i, 9)) & ","
                        s = s & """l5Expr"":" & modEngineBridge.JsonStr(arr(i, 10)) & ","
                        s = s & """l6Expr"":" & modEngineBridge.JsonStr(arr(i, 11)) & ","
                        s = s & """l7Expr"":" & modEngineBridge.JsonStr(arr(i, 12)) & ","
                        s = s & """m1Expr"":" & modEngineBridge.JsonStr(arr(i, 13)) & ","
                        s = s & """m2Expr"":" & modEngineBridge.JsonStr(arr(i, 14)) & ","
                        s = s & """dia"":" & modEngineBridge.JsonNum(Val(CStr(arr(i, 15)))) & ","
                        s = s & """qtyPerElExpr"":" & modEngineBridge.JsonStr(arr(i, 16)) & ","
                        s = s & """elQtyExpr"":" & modEngineBridge.JsonStr(arr(i, 17)) & ","
                        s = s & """note"":" & modEngineBridge.JsonStr(arr(i, 24)) & ","
                        s = s & """tagId"":" & modEngineBridge.JsonStr(tagVal) & ","
                        s = s & """boqBucket"":" & modEngineBridge.JsonStr(arr(i, 26)) & "}"
                        If cnt >= cap Then cap = cap * 2: ReDim Preserve items(0 To cap - 1)
                        items(cnt) = s: cnt = cnt + 1
                    End If
                Next i
            End If
        End If
    Next ws

    If cnt = 0 Then
        CollectExistingTktManualRowsJson = "[]"
    Else
        ReDim Preserve items(0 To cnt - 1)
        CollectExistingTktManualRowsJson = "[" & Join(items, ",") & "]"
    End If
End Function

Public Sub MarkSelectedRowAsManual()
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long: r = ActiveCell.Row
    If modRebarTKT.IsTKTSheet(ws) Then
        ws.Cells(r, 25).Value = "MANUAL"
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 24)).Interior.Color = RGB(235, 245, 255)
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 \u0111\u00e1nh d\u1ea5u d\u00f2ng ") & r & " = MANUAL (lu\u00f4n gi\u1eef nguy\u00ean khi Sync).", vbInformation
    ElseIf Right$(ws.Name, 7) = "_Detail" Then
        ws.Cells(r, 13).Value = "MANUAL"
        If Trim$(CStr(ws.Cells(r, 14).Value)) = "" And r > 6 Then
            ws.Cells(r, 14).Value = ws.Cells(r - 1, 14).Value
        End If
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)).Interior.Color = RGB(235, 245, 255)
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 \u0111\u00e1nh d\u1ea5u d\u00f2ng ") & r & " = MANUAL (s\u1ebd gi\u1eef nguy\u00ean khi Sync).", vbInformation
    Else
        modUnicode.MsgBoxU modUnicode.U("Vui l\u00f2ng \u0111\u1ee9ng trong sheet *_Detail ho\u1eb7c TKT_* \u0111\u1ec3 \u0111\u00e1nh d\u1ea5u MANUAL."), vbExclamation
    End If
End Sub
