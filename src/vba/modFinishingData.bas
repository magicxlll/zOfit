Attribute VB_Name = "modFinishingData"
Option Explicit

' ==============================================================================
' modFinishingData - Wave 2 Finishing Input Sheets (QS_DoorWin & QS_Rooms)
' Connects to zofit-engine.exe --cmd generate_finishing -> Finishing_Detail
' ==============================================================================

Public Const DOORWIN_SHEET As String = "QS_DoorWin"
Public Const ROOMS_SHEET As String = "QS_Rooms"
Public Const FIN_FIRST_ROW As Long = 5

Public Sub CreateFinishingInputSheets(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook

    Dim hasSpec As Boolean
    On Error Resume Next
    hasSpec = Not (wb.Worksheets(modSpecSheet.SPEC_SHEET) Is Nothing)
    On Error GoTo 0
    If Not hasSpec Then modSpecSheet.CreateOrShowSpecSheet False

    Dim wsDoor As Worksheet, wsRoom As Worksheet
    On Error Resume Next
    Set wsDoor = wb.Worksheets(DOORWIN_SHEET)
    Set wsRoom = wb.Worksheets(ROOMS_SHEET)
    On Error GoTo 0

    If wsDoor Is Nothing Then
        Set wsDoor = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsDoor.Name = DOORWIN_SHEET
        BuildDoorWinLayout wsDoor
    End If
    If wsRoom Is Nothing Then
        Set wsRoom = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsRoom.Name = ROOMS_SHEET
        BuildRoomsLayout wsRoom
    End If

    wsRoom.Activate
    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 t\u1ea1o 2 b\u1ea3ng nh\u1eadp Ho\u00e0n thi\u1ec7n: QS_DoorWin (Tra c\u1eeda) & QS_Rooms (Ph\u00f2ng/T\u01b0\u1eddng/M\u00e1i)."), vbInformation
    End If
End Sub

Private Sub BuildDoorWinLayout(ByVal ws As Worksheet)
    Application.ScreenUpdating = False
    ws.Cells.Clear: ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 9.5
    With ws.Range("A1:G1")
        .Merge
        .Value = modUnicode.U("B\u1ea2NG TRA K\u00cdCH TH\u01af\u1edaC C\u1eecA \u0110I, C\u1eecA S\u1ed4 & V\u00c1CH K\u00cdNH (QS DOOR/WINDOW SCHEDULE)")
        .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(22, 54, 92): .RowHeight = 28: .IndentLevel = 1
    End With
    With ws.Range("A2:G2")
        .Merge
        .Value = modUnicode.U("T\u1ea5t c\u1ea3 c\u00f4ng th\u1ee9c tr\u1eeb c\u1eeda x\u00e2y/tr\u00e1t t\u1ea1i Finishing_Detail d\u00f9ng INDEX/MATCH tr\u1ecf t\u1edbi b\u1ea3ng n\u00e0y.")
        .Font.Italic = True: .Interior.Color = RGB(237, 242, 248)
    End With

    Dim hdrs As Variant
    hdrs = Array("DoorCode", "Type", "W (m)", "H (m)", "Area (m2)", "Lintel_L (m)", "Note")
    Dim c As Long
    For c = 0 To UBound(hdrs): ws.Cells(4, c + 1).Value = hdrs(c): Next c
    With ws.Range("A4:G4")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(47, 85, 151): .HorizontalAlignment = xlCenter
    End With

    SeedDoorRow ws, 5, "1SD1", "DOOR", 1.2, 2.2, "Single Steel Door"
    SeedDoorRow ws, 6, "2SD1", "DOOR", 2.0, 2.2, "Double Steel Door"
    SeedDoorRow ws, 7, "WD1", "WINDOW", 1.8, 1.5, "Alu Sliding Window"

    ws.Range("C5:D100").Interior.Color = RGB(255, 250, 220)
    ws.Columns("A:G").AutoFit
    Application.ScreenUpdating = True
End Sub

Private Sub SeedDoorRow(ByVal ws As Worksheet, ByVal r As Long, ByVal code As String, ByVal tp As String, ByVal w As Double, ByVal h As Double, ByVal noteStr As String)
    ws.Cells(r, 1).Value = code
    ws.Cells(r, 2).Value = tp
    ws.Cells(r, 3).Value = w
    ws.Cells(r, 4).Value = h
    ws.Cells(r, 5).Formula = "=C" & r & "*D" & r
    ws.Cells(r, 6).Formula = "=C" & r & "+0.4"
    ws.Cells(r, 7).Value = noteStr
End Sub

Private Sub BuildRoomsLayout(ByVal ws As Worksheet)
    Application.ScreenUpdating = False
    ws.Cells.Clear: ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 9.5
    With ws.Range("A1:P1")
        .Merge
        .Value = modUnicode.U("B\u1ea2NG NH\u1eacP PH\u00d2NG, T\u01af\u1edcNG & M\u00c1I HO\u00c0N THI\u1ec6N (SINGLE-ENTRY QS ROOMS)")
        .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(22, 54, 92): .RowHeight = 28: .IndentLevel = 1
    End With

    Dim hdrs As Variant
    hdrs = Array("RoomID", "Building", "Floor", "RoomName", "Area (m2)", "Perim (m)", "Height (m)", "WallType", "WallLen (m)", "DoorCodes", "FloorFinish", "WallTileH (m)", "Ceiling", "IsWetArea", "RoofArea (m2)", "ParapetLen (m)")
    Dim c As Long
    For c = 0 To UBound(hdrs): ws.Cells(4, c + 1).Value = hdrs(c): Next c
    With ws.Range("A4:P4")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(47, 85, 151): .HorizontalAlignment = xlCenter: .WrapText = True: .RowHeight = 26
    End With

    SeedRoomRow ws, 5, "RM-ROOF", "Factory", "Roof", "Main Roof R1", 0, 0, 0, "NONE", 0, "", "NONE", 0, "NONE", "N", 864#, 124.7
    SeedRoomRow ws, 6, "RM-WC1", "Office", "1F", "Toilet 1F", 25#, 20#, 3.2, "B200", 20#, "1SD1*1,WD1*2", "TILE", 2.4, "GYPSUM", "Y", 0, 0

    ws.Range("E5:G200").Interior.Color = RGB(255, 250, 220)
    ws.Range("I5:J200").Interior.Color = RGB(255, 250, 220)
    ws.Range("O5:P200").Interior.Color = RGB(255, 250, 220)
    ws.Columns("A:P").AutoFit
    Application.ScreenUpdating = True
End Sub

Private Sub SeedRoomRow(ByVal ws As Worksheet, ByVal r As Long, ParamArray vals() As Variant)
    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        ws.Cells(r, i + 1).Value = vals(i)
    Next i
End Sub

Public Sub GenerateFinishingDetail(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook

    Dim wsDoor As Worksheet, wsRoom As Worksheet
    On Error Resume Next
    Set wsDoor = wb.Worksheets(DOORWIN_SHEET)
    Set wsRoom = wb.Worksheets(ROOMS_SHEET)
    On Error GoTo 0
    If wsDoor Is Nothing Or wsRoom Is Nothing Then
        CreateFinishingInputSheets False
        Set wsDoor = wb.Worksheets(DOORWIN_SHEET)
        Set wsRoom = wb.Worksheets(ROOMS_SHEET)
    End If

    Dim reqJson As String
    reqJson = "{" & _
        """spec"":" & modSpecSheet.SerializeSpecToJson(wb) & "," & _
        """doors"":" & SerializeDoorsJson(wsDoor) & "," & _
        """rooms"":" & SerializeRoomsJson(wsRoom) & "," & _
        """existingRows"":" & modQsGenerator.CollectExistingDetailManualRowsJson(wb, "", "Finishing_Detail") & _
        "}"

    Dim resJson As String, errMsg As String
    If Not modEngineBridge.RunEngineCommand("generate_finishing", reqJson, resJson, errMsg) Then
        If showMsg Then modUnicode.MsgBoxU errMsg, vbCritical
        Exit Sub
    End If

    Dim root As Object: Set root = modEngineBridge.JsonParseObject(resJson)
    Dim shObj As Object: Set shObj = root("data")("sheet")

    Application.ScreenUpdating = False
    modQsGenerator.WriteDetailSheetFromDto wb, shObj
    Application.ScreenUpdating = True

    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 sinh b\u1ea3ng b\u00f3c t\u00e1ch Ho\u00e0n thi\u1ec7n Finishing_Detail! (T\u1ef1 tr\u1eeb c\u1eeda qua INDEX/MATCH & li\u00ean k\u1ebft m\u00e1i R1)."), vbInformation
    End If
End Sub

Private Function SerializeDoorsJson(ByVal ws As Worksheet) As String
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < FIN_FIRST_ROW Then SerializeDoorsJson = "[]": Exit Function
    Dim arr As Variant: arr = ws.Range(ws.Cells(FIN_FIRST_ROW, 1), ws.Cells(lr, 7)).Value2
    Dim items() As String, cnt As Long, i As Long
    cnt = 0: ReDim items(0 To UBound(arr, 1) - 1)
    For i = 1 To UBound(arr, 1)
        If Len(Trim$(CStr(arr(i, 1)))) > 0 Then
            items(cnt) = "{""dataRow"":" & (FIN_FIRST_ROW + i - 1) & ",""code"":" & modEngineBridge.JsonStr(arr(i, 1)) & ",""doorType"":" & modEngineBridge.JsonStr(arr(i, 2)) & ",""widthM"":" & modEngineBridge.JsonNum(arr(i, 3)) & ",""heightM"":" & modEngineBridge.JsonNum(arr(i, 4)) & "}"
            cnt = cnt + 1
        End If
    Next i
    If cnt = 0 Then SerializeDoorsJson = "[]" Else ReDim Preserve items(0 To cnt - 1): SerializeDoorsJson = "[" & Join(items, ",") & "]"
End Function

Private Function SerializeRoomsJson(ByVal ws As Worksheet) As String
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < FIN_FIRST_ROW Then SerializeRoomsJson = "[]": Exit Function
    Dim arr As Variant: arr = ws.Range(ws.Cells(FIN_FIRST_ROW, 1), ws.Cells(lr, 16)).Value2
    Dim items() As String, cnt As Long, i As Long
    cnt = 0: ReDim items(0 To UBound(arr, 1) - 1)
    For i = 1 To UBound(arr, 1)
        If Len(Trim$(CStr(arr(i, 1)))) > 0 Then
            Dim s As String
            s = "{""dataRow"":" & (FIN_FIRST_ROW + i - 1) & ","
            s = s & """roomId"":" & modEngineBridge.JsonStr(arr(i, 1)) & ","
            s = s & """building"":" & modEngineBridge.JsonStr(arr(i, 2)) & ","
            s = s & """floor"":" & modEngineBridge.JsonStr(arr(i, 3)) & ","
            s = s & """roomName"":" & modEngineBridge.JsonStr(arr(i, 4)) & ","
            s = s & """areaM2"":" & modEngineBridge.JsonNum(arr(i, 5)) & ","
            s = s & """perimM"":" & modEngineBridge.JsonNum(arr(i, 6)) & ","
            s = s & """heightM"":" & modEngineBridge.JsonNum(arr(i, 7)) & ","
            s = s & """wallType"":" & modEngineBridge.JsonStr(arr(i, 8)) & ","
            s = s & """wallLenM"":" & modEngineBridge.JsonNum(arr(i, 9)) & ","
            s = s & """doorCodes"":" & modEngineBridge.JsonStr(arr(i, 10)) & ","
            s = s & """floorFinish"":" & modEngineBridge.JsonStr(arr(i, 11)) & ","
            s = s & """wallTileHeightM"":" & modEngineBridge.JsonNum(arr(i, 12)) & ","
            s = s & """ceilingType"":" & modEngineBridge.JsonStr(arr(i, 13)) & ","
            s = s & """isWetArea"":" & modEngineBridge.JsonStr(arr(i, 14)) & ","
            s = s & """roofAreaM2"":" & modEngineBridge.JsonNum(arr(i, 15)) & ","
            s = s & """parapetLenM"":" & modEngineBridge.JsonNum(arr(i, 16)) & "}"
            items(cnt) = s
            cnt = cnt + 1
        End If
    Next i
    If cnt = 0 Then SerializeRoomsJson = "[]" Else ReDim Preserve items(0 To cnt - 1): SerializeRoomsJson = "[" & Join(items, ",") & "]"
End Function
