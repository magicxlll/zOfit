Attribute VB_Name = "modInfraData"
Option Explicit

' ==============================================================================
' modInfraData - Wave 2 Infrastructure Input Sheet (QS_Infra)
' Auto-feeds pipe/trench lengths from QS_CAD_Parsed + generates Civil_Detail
' ==============================================================================

Public Const INFRA_SHEET As String = "QS_Infra"
Public Const INFRA_FIRST_ROW As Long = 5

Public Sub CreateOrShowInfraSheet(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook

    Dim hasSpec As Boolean
    On Error Resume Next
    hasSpec = Not (wb.Worksheets(modSpecSheet.SPEC_SHEET) Is Nothing)
    On Error GoTo 0
    If Not hasSpec Then modSpecSheet.CreateOrShowSpecSheet False

    Dim ws As Worksheet, isNew As Boolean
    On Error Resume Next
    Set ws = wb.Worksheets(INFRA_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = INFRA_SHEET
        isNew = True
    End If

    If isNew Then BuildInfraLayout ws
    ws.Activate
    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 t\u1ea1o/m\u1edf b\u1ea3ng H\u1ea1 t\u1ea7ng QS_Infra! C\u00f3 th\u1ec3 n\u1ea1p tr\u1ef1c ti\u1ebfp chi\u1ec1u d\u00e0i t\u1eeb QS_CAD_Parsed ho\u1eb7c ch\u1ea5m Linkcad."), vbInformation
    End If
End Sub

Private Sub BuildInfraLayout(ByVal ws As Worksheet)
    Application.ScreenUpdating = False
    ws.Cells.Clear: ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 9.5

    With ws.Range("A1:P1")
        .Merge
        .Value = modUnicode.U("B\u1ea2NG NH\u1eacP H\u1ea0 T\u1ea6NG K\u1ef8 THU\u1eacT (C\u1ed0NG, \u0110\u01af\u1edcNG N\u1ed8I B\u1ed8, T\u01af\u1edcNG R\u00c0O) - QS INFRA")
        .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(22, 54, 92): .RowHeight = 28: .IndentLevel = 1
    End With

    Dim hdrs As Variant
    hdrs = Array("InfraID", "Category", "Code", "Description", "Qty", "L (m)", "W (m)", "Depth/Thk (m)", "Dia (mm)", "CPDD1 (m)", "CPDD2 (m)", "Bedding (m)", "Sidewalk W (m)", "Curb L (m)", "ParamA (Joint/BrickH)", "ParamB (ColSpace)")
    Dim c As Long
    For c = 0 To UBound(hdrs): ws.Cells(4, c + 1).Value = hdrs(c): Next c
    With ws.Range("A4:P4")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(47, 85, 151): .HorizontalAlignment = xlCenter: .WrapText = True: .RowHeight = 28
    End With

    With ws.Range("B5:B500").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="PIPE,CULVERT,ROAD_BTXM,ROAD_ASPHALT,FENCE"
        .InCellDropdown = True
    End With

    SeedInfraRow ws, 5, "INF-P1", "PIPE", "D200", "Storm Drain HDPE D200", 1, 120, 0, 1.5, 200, 0, 0, 0.15, 0, 0, 0, 0
    SeedInfraRow ws, 6, "INF-R1", "ROAD_BTXM", "RD-A", "Internal Road BTXM M300", 1, 100, 7#, 0.22, 0, 0.15, 0.15, 0, 1.5, 200, 5#, 0
    SeedInfraRow ws, 7, "INF-F1", "FENCE", "F-1", "Perimeter Fence 2.4m", 1, 250, 0.2, 2.4, 0, 0, 0, 0, 0, 0, 0.8, 3#

    ws.Range("E5:P200").Interior.Color = RGB(255, 250, 220)
    ws.Columns("A:P").AutoFit
    Application.ScreenUpdating = True
End Sub

Private Sub SeedInfraRow(ByVal ws As Worksheet, ByVal r As Long, ParamArray vals() As Variant)
    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        ws.Cells(r, i + 1).Value = vals(i)
    Next i
End Sub

Public Sub FeedLengthsFromCadParsed(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook

    Dim wsCad As Worksheet, wsInf As Worksheet
    On Error Resume Next
    Set wsCad = wb.Worksheets(modCadParser.CAD_PARSED_SHEET)
    Set wsInf = wb.Worksheets(INFRA_SHEET)
    On Error GoTo 0

    If wsCad Is Nothing Then
        If showMsg Then modUnicode.MsgBoxU modUnicode.U("Ch\u01b0a c\u00f3 sheet QS_CAD_Parsed! H\u00e3y ch\u1ea1y 'T\u00e1ch CAD Data' tr\u01b0\u1edbc."), vbExclamation
        Exit Sub
    End If
    If wsInf Is Nothing Then
        CreateOrShowInfraSheet False
        Set wsInf = wb.Worksheets(INFRA_SHEET)
    End If

    Dim lrCad As Long: lrCad = wsCad.Cells(wsCad.Rows.Count, 13).End(xlUp).Row
    If lrCad < 4 Then Exit Sub

    Dim rNext As Long: rNext = wsInf.Cells(wsInf.Rows.Count, 1).End(xlUp).Row + 1
    If rNext < INFRA_FIRST_ROW Then rNext = INFRA_FIRST_ROW

    Dim r As Long, keyStr As String, addedCnt As Long
    addedCnt = 0
    For r = 4 To lrCad
        keyStr = Trim$(CStr(wsCad.Cells(r, 13).Value))
        If Len(keyStr) > 0 Then
            ' Check if Code already exists in QS_Infra Col C -> update L formula instead of duplicating
            Dim foundRow As Long: foundRow = 0
            Dim ir As Long, lrInf As Long: lrInf = wsInf.Cells(wsInf.Rows.Count, 1).End(xlUp).Row
            For ir = INFRA_FIRST_ROW To lrInf
                If StrComp(Trim$(CStr(wsInf.Cells(ir, 3).Value)), keyStr, vbTextCompare) = 0 Then
                    foundRow = ir: Exit For
                End If
            Next ir
            Dim tgtR As Long: tgtR = IIf(foundRow > 0, foundRow, rNext)
            If foundRow = 0 Then
                Dim cat As String: cat = IIf(Left$(keyStr, 1) = "D" Or InStr(1, keyStr, "HDPE", vbTextCompare) > 0, "PIPE", "CULVERT")
                wsInf.Cells(tgtR, 1).Value = "INF-CAD-" & (tgtR - INFRA_FIRST_ROW + 1)
                wsInf.Cells(tgtR, 2).Value = cat
                wsInf.Cells(tgtR, 3).Value = keyStr
                wsInf.Cells(tgtR, 4).Value = "Auto CAD: " & keyStr
                wsInf.Cells(tgtR, 5).Value = 1
                wsInf.Cells(tgtR, 8).Value = 1.5
                wsInf.Cells(tgtR, 9).Value = ExtractNumber(keyStr)
                wsInf.Cells(tgtR, 12).Value = 0.15
                rNext = rNext + 1
            End If
            ' Link L (Col F) directly to QS_CAD_Parsed Summary!
            wsInf.Cells(tgtR, 6).Formula = "=SUMIF(QS_CAD_Parsed!$M$4:$M$500,C" & tgtR & ",QS_CAD_Parsed!$O$4:$O$500)"
            addedCnt = addedCnt + 1
        End If
    Next r

    wsInf.Activate
    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 li\u00ean k\u1ebft chi\u1ec1u d\u00e0i ") & addedCnt & modUnicode.U(" tuy\u1ebfn c\u1ed1ng/m\u01b0\u01a1ng t\u1eeb QS_CAD_Parsed sang QS_Infra!"), vbInformation
    End If
End Sub

Private Function ExtractNumber(ByVal s As String) As Double
    Dim i As Long, ch As String, numStr As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            numStr = numStr & ch
        ElseIf Len(numStr) > 0 Then
            Exit For
        End If
    Next i
    ExtractNumber = Val(numStr)
End Function

Public Sub GenerateInfraDetail(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook

    Dim wsInf As Worksheet
    On Error Resume Next
    Set wsInf = wb.Worksheets(INFRA_SHEET)
    On Error GoTo 0
    If wsInf Is Nothing Then
        CreateOrShowInfraSheet False
        Set wsInf = wb.Worksheets(INFRA_SHEET)
    End If

    Dim reqJson As String
    reqJson = "{" & _
        """spec"":" & modSpecSheet.SerializeSpecToJson(wb) & "," & _
        """items"":" & SerializeInfraJson(wsInf) & "," & _
        """existingRows"":" & modQsGenerator.CollectExistingDetailManualRowsJson(wb, "", "Civil_Detail") & _
        "}"

    Dim resJson As String, errMsg As String
    If Not modEngineBridge.RunEngineCommand("generate_infra", reqJson, resJson, errMsg) Then
        If showMsg Then modUnicode.MsgBoxU errMsg, vbCritical
        Exit Sub
    End If

    Dim root As Object: Set root = modEngineBridge.JsonParseObject(resJson)
    Dim shObj As Object: Set shObj = root("data")("sheet")

    Application.ScreenUpdating = False
    modQsGenerator.WriteDetailSheetFromDto wb, shObj
    Application.ScreenUpdating = True

    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 sinh b\u1ea3ng b\u00f3c t\u00e1ch H\u1ea1 t\u1ea7ng Civil_Detail (T\u1ef1 tr\u1eeb chi\u1ebfm ch\u1ed7 \u1ed1ng, taluy, \u0111\u01b0\u1eddng & t\u01b0\u1eddng r\u00e0o)."), vbInformation
    End If
End Sub

Private Function SerializeInfraJson(ByVal ws As Worksheet) As String
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < INFRA_FIRST_ROW Then SerializeInfraJson = "[]": Exit Function
    Dim arr As Variant: arr = ws.Range(ws.Cells(INFRA_FIRST_ROW, 1), ws.Cells(lr, 16)).Value2
    Dim items() As String, cnt As Long, i As Long
    cnt = 0: ReDim items(0 To UBound(arr, 1) - 1)
    For i = 1 To UBound(arr, 1)
        If Len(Trim$(CStr(arr(i, 1)))) > 0 Then
            Dim s As String
            s = "{""dataRow"":" & (INFRA_FIRST_ROW + i - 1) & ","
            s = s & """infraId"":" & modEngineBridge.JsonStr(arr(i, 1)) & ","
            s = s & """category"":" & modEngineBridge.JsonStr(arr(i, 2)) & ","
            s = s & """code"":" & modEngineBridge.JsonStr(arr(i, 3)) & ","
            s = s & """description"":" & modEngineBridge.JsonStr(arr(i, 4)) & ","
            s = s & """qty"":" & modEngineBridge.JsonNum(arr(i, 5)) & ","
            s = s & """lengthM"":" & modEngineBridge.JsonNum(arr(i, 6)) & ","
            s = s & """widthM"":" & modEngineBridge.JsonNum(arr(i, 7)) & ","
            s = s & """depthOrThkM"":" & modEngineBridge.JsonNum(arr(i, 8)) & ","
            s = s & """diaMm"":" & modEngineBridge.JsonNum(arr(i, 9)) & ","
            s = s & """subBaseT1M"":" & modEngineBridge.JsonNum(arr(i, 10)) & ","
            s = s & """subBaseT2M"":" & modEngineBridge.JsonNum(arr(i, 11)) & ","
            s = s & """cushionThkM"":" & modEngineBridge.JsonNum(arr(i, 12)) & ","
            s = s & """extraSidewalkWM"":" & modEngineBridge.JsonNum(arr(i, 13)) & ","
            s = s & """curbLenM"":" & modEngineBridge.JsonNum(arr(i, 14)) & ","
            s = s & """paramA"":" & modEngineBridge.JsonNum(arr(i, 15)) & ","
            s = s & """paramB"":" & modEngineBridge.JsonNum(arr(i, 16)) & "}"
            items(cnt) = s
            cnt = cnt + 1
        End If
    Next i
    If cnt = 0 Then SerializeInfraJson = "[]" Else ReDim Preserve items(0 To cnt - 1): SerializeInfraJson = "[" & Join(items, ",") & "]"
End Function
