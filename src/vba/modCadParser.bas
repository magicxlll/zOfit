Attribute VB_Name = "modCadParser"
Option Explicit

' ==============================================================================
' modCadParser - Parses Raw CAD Text/Block dumps (Data sheet or active selection)
' Creates QS_CAD_Parsed with pure Excel SUMIFS/AVERAGEIFS summary table
' ==============================================================================

Public Const CAD_PARSED_SHEET As String = "QS_CAD_Parsed"

Public Sub ParseRawCadDataToSheet(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook

    Dim srcWs As Worksheet, srcRng As Range
    Set srcWs = ActiveSheet

    ' If user selected >= 3 rows on active sheet, parse Selection; else look for sheet "Data"
    If TypeName(Selection) = "Range" Then
        If Selection.Rows.Count >= 3 And StrComp(srcWs.Name, CAD_PARSED_SHEET, vbTextCompare) <> 0 Then
            Set srcRng = Selection
        End If
    End If

    If srcRng Is Nothing Then
        On Error Resume Next
        Set srcWs = wb.Worksheets("Data")
        On Error GoTo 0
        If srcWs Is Nothing Then Set srcWs = ActiveSheet
        Dim lr As Long: lr = srcWs.UsedRange.Row + srcWs.UsedRange.Rows.Count - 1
        Dim lc As Long: lc = srcWs.UsedRange.Column + srcWs.UsedRange.Columns.Count - 1
        If lr < 2 Then
            If showMsg Then modUnicode.MsgBoxU modUnicode.U("Kh\u00f4ng t\u00ecm th\u1ea5y d\u1eef li\u1ec7u CAD th\u00f4 (h\u00e3y b\u00f4i \u0111en b\u1ea3ng text/block CAD ho\u1eb7c \u0111\u1eb7t v\u00e0o sheet Data)."), vbExclamation
            Exit Sub
        End If
        Set srcRng = srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lr, IIf(lc > 20, 20, lc)))
    End If

    Dim jsonReq As String
    jsonReq = BuildCadRequestJson(srcRng)

    Dim resJson As String, errMsg As String
    If Not modEngineBridge.RunEngineCommand("parse_cad", jsonReq, resJson, errMsg) Then
        If showMsg Then modUnicode.MsgBoxU errMsg, vbCritical
        Exit Sub
    End If

    Dim root As Object: Set root = modEngineBridge.JsonParseObject(resJson)
    Dim dataObj As Object: Set dataObj = root("data")

    WriteCadParsedSheet wb, dataObj

    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 t\u00e1ch xong d\u1eef li\u1ec7u CAD th\u00f4: ") & CLng(dataObj("parsedCount")) & "/" & CLng(dataObj("totalInput")) & modUnicode.U(" d\u00f2ng h\u1ee3p l\u1ec7 (c\u00f2n ") & CLng(dataObj("unknownCount")) & modUnicode.U(" d\u00f2ng c\u1ea7n ki\u1ec3m tra)."), vbInformation
    End If
End Sub

Private Function BuildCadRequestJson(ByVal rng As Range) As String
    Dim arr As Variant
    arr = rng.Value
    Dim items() As String, cnt As Long, cap As Long
    cap = 64: cnt = 0
    ReDim items(0 To cap - 1)

    Dim nR As Long: nR = UBound(arr, 1)
    Dim nC As Long: nC = UBound(arr, 2)
    Dim startRow As Long: startRow = rng.Row

    Dim r As Long, c As Long, txtVal As String
    For r = 1 To nR
        ' Scan columns in each row for non-empty CAD text/block pattern
        Dim aVal As String, bVal As String, cVal As String, dVal As String
        If nC <= 4 Then
            aVal = Trim$(CStr(arr(r, 1)))
            If nC >= 2 Then bVal = Trim$(CStr(arr(r, 2))) Else bVal = ""
            If nC >= 3 Then cVal = Trim$(CStr(arr(r, 3))) Else cVal = ""
            If nC >= 4 Then dVal = Trim$(CStr(arr(r, 4))) Else dVal = ""
            If Len(aVal) > 0 And Not IsNumeric(aVal) Then
                AddCadItem items, cnt, cap, startRow + r - 1, aVal, bVal, cVal, dVal
            End If
        Else
            ' Wide sheet like BHFLEX Data (multiple side-by-side tables in cols A..P)
            For c = 1 To nC
                txtVal = Trim$(CStr(arr(r, c)))
                If Len(txtVal) >= 4 And Not IsNumeric(txtVal) Then
                    If InStr(1, txtVal, "-L", vbTextCompare) > 0 Or InStr(1, txtVal, "L=", vbTextCompare) > 0 Or InStr(1, txtVal, "Box:", vbTextCompare) > 0 Or InStr(1, txtVal, "TN-", vbTextCompare) > 0 Then
                        bVal = IIf(c + 1 <= nC, Trim$(CStr(arr(r, c + 1))), "")
                        cVal = IIf(c + 2 <= nC, Trim$(CStr(arr(r, c + 2))), "")
                        dVal = IIf(c + 3 <= nC, Trim$(CStr(arr(r, c + 3))), "")
                        AddCadItem items, cnt, cap, startRow + r - 1, txtVal, bVal, cVal, dVal
                    End If
                End If
            Next c
        End If
    Next r

    If cnt = 0 Then
        BuildCadRequestJson = "{""rawRows"":[]}"
    Else
        ReDim Preserve items(0 To cnt - 1)
        BuildCadRequestJson = "{""rawRows"":[" & Join(items, ",") & "]}"
    End If
End Function

Private Sub AddCadItem(ByRef items() As String, ByRef cnt As Long, ByRef cap As Long, _
                       ByVal srcRow As Long, ByVal a As String, ByVal b As String, ByVal c As String, ByVal d As String)
    If cnt >= cap Then
        cap = cap * 2
        ReDim Preserve items(0 To cap - 1)
    End If
    items(cnt) = "{""sourceRow"":" & srcRow & ",""colA"":" & modEngineBridge.JsonStr(a) & ",""colB"":" & modEngineBridge.JsonStr(b) & ",""colC"":" & modEngineBridge.JsonStr(c) & ",""colD"":" & modEngineBridge.JsonStr(d) & "}"
    cnt = cnt + 1
End Sub

Private Sub WriteCadParsedSheet(ByVal wb As Workbook, ByVal dataObj As Object)
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(CAD_PARSED_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = CAD_PARSED_SHEET
    Else
        ws.Cells.Clear
    End If

    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 9.5
    With ws.Range("A1:K1")
        .Merge
        .Value = modUnicode.U("B\u1ea2NG T\u00c1CH & T\u1ed4NG H\u1ee2P D\u1eee LI\u1ec6U CAD TH\u00d4 (QS CAD PARSER)")
        .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(22, 54, 92): .RowHeight = 28: .IndentLevel = 1
    End With

    ' Left Table (Cols A..I): Parsed Rows (Row 3 header, rows 4..n)
    Dim itemsCol As Collection: Set itemsCol = dataObj("items")
    Dim keysCol As Collection: Set keysCol = dataObj("distinctKeys")

    Dim hdrs As Variant
    hdrs = Array("SourceRow", "RawText", "ItemType", "CategoryKey", "W/Dia (mm)", "H (mm)", "Slope (%)", "Length (m)", "Depth (m)", "Qty", "Check")
    Dim c As Long
    For c = 0 To UBound(hdrs)
        ws.Cells(3, c + 1).Value = hdrs(c)
    Next c
    With ws.Range("A3:K3")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 85, 151): .HorizontalAlignment = xlCenter: .RowHeight = 24
    End With

    Dim nItems As Long: nItems = itemsCol.Count
    If nItems > 0 Then
        Dim arr() As Variant
        ReDim arr(1 To nItems, 1 To 11)
        Dim i As Long, it As Object
        For i = 1 To nItems
            Set it = itemsCol(i)
            arr(i, 1) = CLng(it("sourceRow"))
            arr(i, 2) = CStr(it("rawText"))
            arr(i, 3) = CStr(it("itemType"))
            arr(i, 4) = CStr(it("categoryKey"))
            arr(i, 5) = IIf(CDbl(it("diameterMm")) > 0, CDbl(it("diameterMm")), CDbl(it("widthMm")))
            arr(i, 6) = CDbl(it("heightMm"))
            arr(i, 7) = CDbl(it("slopePct"))
            arr(i, 8) = CDbl(it("lengthM"))
            arr(i, 9) = CDbl(it("depthM"))
            arr(i, 10) = CDbl(it("qty"))
            arr(i, 11) = IIf(CBool(it("needsManualCheck")), "MANUAL_CHECK", "OK")
        Next i
        ws.Range(ws.Cells(4, 1), ws.Cells(3 + nItems, 11)).Value2 = arr
    End If

    ' Right Table (Cols M..Q): Live SUMIFS / AVERAGEIFS Summary by CategoryKey
    ws.Range("M2:Q2").Merge
    ws.Range("M2").Value = modUnicode.U("T\u1ed4NG H\u1ee2P THEO LO\u1ea0I (C\u00d4NG TH\u1ee8C SUMIFS T\u1ef0 \u0110\u1ed8NG)")
    ws.Range("M2").Font.Bold = True: ws.Range("M2").Interior.Color = RGB(217, 225, 242)

    ws.Cells(3, 13).Value = "CategoryKey"
    ws.Cells(3, 14).Value = "Count"
    ws.Cells(3, 15).Value = "Total L (m)"
    ws.Cells(3, 16).Value = "Avg Depth (m)"
    ws.Cells(3, 17).Value = "Total Qty"
    With ws.Range("M3:Q3")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(56, 87, 35): .HorizontalAlignment = xlCenter: .RowHeight = 24
    End With

    Dim lastRow As Long: lastRow = 3 + IIf(nItems > 0, nItems, 1)
    Dim k As Long
    For k = 1 To keysCol.Count
        Dim kr As Long: kr = 3 + k
        ws.Cells(kr, 13).Value = CStr(keysCol(k))
        ws.Cells(kr, 14).Formula = "=COUNTIF($D$4:$D$" & lastRow & ",M" & kr & ")"
        ws.Cells(kr, 15).Formula = "=SUMIF($D$4:$D$" & lastRow & ",M" & kr & ",$H$4:$H$" & lastRow & ")"
        ws.Cells(kr, 16).Formula = "=IFERROR(AVERAGEIF($D$4:$D$" & lastRow & ",M" & kr & ",$I$4:$I$" & lastRow & "),0)"
        ws.Cells(kr, 17).Formula = "=SUMIF($D$4:$D$" & lastRow & ",M" & kr & ",$J$4:$J$" & lastRow & ")"
        ws.Range("O" & kr & ":Q" & kr).NumberFormat = "#,##0.00"
    Next k

    ws.Columns("A:Q").AutoFit
    ws.Activate
    Application.ScreenUpdating = True
End Sub

Public Sub ScanCadBeamSectionToSheet(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook

    ' Ensure QS_Data and QS_RebarConfig exist
    modQsData.CreateOrShowQsDataSheet False
    Dim wsData As Worksheet: Set wsData = wb.Worksheets(modQsData.QSDATA_SHEET)

    modRebarConfig.EnsureRebarConfigSheet wb, False
    Dim wsCfg As Worksheet: Set wsCfg = wb.Worksheets(modRebarConfig.REBAR_CFG_SHEET)

    ' Call zofit-engine via JSON-RPC to AutoCAD MCP
    Dim jsonReq As String
    jsonReq = "{""host"":""127.0.0.1"",""port"":9877,""building"":""MAIN"",""floor"":""TANG1"",""defaultLengthM"":6.0,""autoOpenPreview"":true}"

    Dim resJson As String, errMsg As String
    If Not modEngineBridge.RunEngineCommand("scan_cad_beam", jsonReq, resJson, errMsg) Then
        If showMsg Then modUnicode.MsgBoxU errMsg, vbCritical
        Exit Sub
    End If

    Dim root As Object: Set root = modEngineBridge.JsonParseObject(resJson)
    If root Is Nothing Then Exit Sub
    If Not root.Exists("data") Then Exit Sub

    Dim dataObj As Object: Set dataObj = root("data")
    Dim accepted As Boolean: accepted = CBool(dataObj("accepted"))
    If Not accepted Then
        If showMsg Then modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 h\u1ee7y qu\u00e9t d\u1ea7m CAD ho\u1eb7c ch\u01b0a ch\u1ecdn 'S\u1eed d\u1ee5ng & Ghi v\u00e0o b\u1ea3ng'."), vbInformation
        Exit Sub
    End If

    Dim beamsCol As Collection: Set beamsCol = dataObj("beams")
    If beamsCol Is Nothing Or beamsCol.Count = 0 Then
        If showMsg Then modUnicode.MsgBoxU modUnicode.U("Kh\u00f4ng t\u00ecm th\u1ea5y b\u1ea3ng m\u1eb7t c\u1eaft d\u1ea7m n\u00e0o trong v\u00f9ng ch\u1ecdn CAD!"), vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim bIdx As Long, beamCount As Long: beamCount = beamsCol.Count
    Dim beamNames As String: beamNames = ""
    Dim totalBarsAdded As Long: totalBarsAdded = 0

    For bIdx = 1 To beamCount
        Dim bObj As Object: Set bObj = beamsCol(bIdx)
        Dim elObj As Object: Set elObj = bObj("element")
        Dim barsCol As Collection: Set barsCol = bObj("rebarRows")

        Dim eid As String: eid = Trim$(CStr(elObj("elementId")))
        Dim bld As String: bld = Trim$(CStr(elObj("building")))
        Dim grp As String: grp = Trim$(CStr(elObj("group")))
        Dim mk As String: mk = Trim$(CStr(elObj("mark")))
        Dim desc As String: desc = Trim$(CStr(elObj("description")))
        Dim lenM As Double: lenM = CDbl(elObj("lengthM"))
        Dim widM As Double: widM = CDbl(elObj("widthM"))
        Dim hgtM As Double: hgtM = CDbl(elObj("heightM"))
        Dim qtyVal As Double: qtyVal = CDbl(elObj("qty"))
        Dim grade As String: grade = Trim$(CStr(elObj("concGrade")))
        Dim preset As String: preset = Trim$(CStr(elObj("rebarPreset")))
        Dim bucket As String: bucket = Trim$(CStr(elObj("boqBucket")))

        If Len(bld) = 0 Then bld = "MAIN"
        If Len(grp) = 0 Then grp = "Beam"
        If Len(grade) = 0 Then grade = "M350"
        If Len(preset) = 0 Then preset = "BEAM_FULL"
        If Len(bucket) = 0 Then bucket = "Superstructure"
        If qtyVal <= 0 Then qtyVal = 1
        If lenM <= 0 Then lenM = 6#

        ' 1. Write or update row in QS_Data
        Dim rData As Long: rData = 0
        Dim lastDataR As Long: lastDataR = modQsData.GetLastQsDataRow(wsData)
        Dim r As Long
        For r = modQsData.QSDATA_FIRST_ROW To lastDataR
            If StrComp(Trim$(CStr(wsData.Cells(r, 1).Value)), eid, vbTextCompare) = 0 Then
                rData = r: Exit For
            End If
        Next r
        If rData = 0 Then
            rData = lastDataR + 1
            If rData < modQsData.QSDATA_FIRST_ROW Then rData = modQsData.QSDATA_FIRST_ROW
        End If

        wsData.Cells(rData, 1).Value = eid
        wsData.Cells(rData, 2).Value = bld
        wsData.Cells(rData, 3).Value = grp
        wsData.Cells(rData, 4).Value = mk
        wsData.Cells(rData, 5).Value = desc
        wsData.Cells(rData, 6).Value = qtyVal
        wsData.Cells(rData, 7).Value = lenM
        wsData.Cells(rData, 8).Value = widM
        wsData.Cells(rData, 9).Value = hgtM
        wsData.Cells(rData, 12).Value = grade
        wsData.Cells(rData, 18).Value = preset
        wsData.Cells(rData, 19).Value = bucket

        ' 2. Append or update Rebar Block in QS_RebarConfig with live formulas
        modRebarConfig.AppendOrUpdateRebarBlock wsCfg, eid, barsCol

        If Len(beamNames) > 0 Then beamNames = beamNames & ", "
        beamNames = beamNames & mk
        totalBarsAdded = totalBarsAdded + barsCol.Count
    Next bIdx

    Application.ScreenUpdating = True
    wsCfg.Activate

    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 qu\u00e9t v\u00e0 n\u1ea1p th\u00e0nh c\u00f4ng ") & beamCount & _
            modUnicode.U(" d\u1ea7m CAD (") & beamNames & modUnicode.U(") v\u1edbi ") & totalBarsAdded & _
            modUnicode.U(" d\u00f2ng th\u00e9p c\u00f4ng th\u1ee9c live sang QS_Data & QS_RebarConfig!"), vbInformation
    End If
End Sub
