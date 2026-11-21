Attribute VB_Name = "modQsData"
Option Explicit

' ==============================================================================
' modQsData - Master Single-Entry Geometry & Lean Override Sheet (QS_Data)
' Clean 19-Column Schema (A..S):
'   A..E: ElementID, Building, Group, Mark, Description
'   F..O: Qty, L, W, H/t, TopElev, BotElev, ConcGrade, Deduct_V, Deduct_FW, IsBuried (Open for Linkcad!)
'   P..Q: LeanThk_Ovr, LeanExt_Ovr (Blank = use default SPEC_Lean_<Type>_*)
'   R..S: RebarPreset, BOQ_Bucket
' ==============================================================================

Public Const QSDATA_SHEET As String = "QS_Data"
Public Const QSDATA_FIRST_ROW As Long = 5

Public Function IsQsDataSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    If CStr(ws.Range("A1").Value) = "ZOFIT_DATA_V2" Or StrComp(ws.Name, QSDATA_SHEET, vbTextCompare) = 0 Then
        If Trim$(CStr(ws.Cells(4, 1).Value)) = "ElementID" Then IsQsDataSheet = True
    End If
    On Error GoTo 0
End Function

Public Sub CreateOrShowQsDataSheet(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub

    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim hasSpec As Boolean
    On Error Resume Next
    hasSpec = Not (wb.Worksheets(modSpecSheet.SPEC_SHEET) Is Nothing)
    On Error GoTo 0
    If Not hasSpec Then modSpecSheet.CreateOrShowSpecSheet False

    Dim ws As Worksheet, isNew As Boolean
    On Error Resume Next
    Set ws = wb.Worksheets(QSDATA_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = QSDATA_SHEET
        isNew = True
    End If

    If isNew Then BuildQsDataLayout ws

    ' Also ensure QS_RebarConfig exists
    modRebarConfig.EnsureRebarConfigSheet wb, False

    ws.Activate
    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 t\u1ea1o/m\u1edf QS_Data (H\u00ecnh h\u1ecdc & L\u00f3t t\u00f9y bi\u1ebfn) c\u00f9ng QS_RebarConfig (C\u1ea5u h\u00ecnh th\u00e9p chi ti\u1ebft \u0111a n\u0103ng)."), vbInformation
    End If
End Sub

Private Sub BuildQsDataLayout(ByVal ws As Worksheet)
    Application.ScreenUpdating = False
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 9.5

    ws.Range("A1").Value = "ZOFIT_DATA_V2"
    ws.Range("A1").Font.Color = RGB(22, 54, 92)

    With ws.Range("B1:S1")
        .Merge
        .Value = modUnicode.U("B\u1ea2NG NH\u1eacP C\u1ea4U KI\u1ec6N CHU\u1ea8N (QS_DATA) - H\u00ccNH H\u1eccC, B\u00ca T\u00d4NG L\u00d3T & T\u01af\u01a0NG TH\u00cdCH LINKCAD")
        .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(22, 54, 92): .RowHeight = 28: .IndentLevel = 1: .VerticalAlignment = xlCenter
    End With
    ws.Range("A1").Interior.Color = RGB(22, 54, 92)

    With ws.Range("A2:S2")
        .Merge
        .Value = modUnicode.U("M\u1eb9o: \u00d4 v\u00e0ng (F..O) m\u1edf cho Linkcad. C\u1ed9t P..Q \u0111\u1ec3 tr\u1ed1ng s\u1ebd l\u1ea5y theo lo\u1ea1i \u1edf QS_Spec (M\u00f3ng 100, D\u1ea7m 50, S\u00e0n 50), ho\u1eb7c nh\u1eadp s\u1ed1 ri\u00eang (0 = kh\u00f4ng l\u00f3t). Th\u00e9p chi ti\u1ebft nh\u1eadp/t\u00f9y bi\u1ebfn t\u1ea1i sheet QS_RebarConfig!")
        .Font.Italic = True: .Font.Size = 9: .Font.Color = RGB(64, 64, 64)
        .Interior.Color = RGB(237, 242, 248): .RowHeight = 20: .IndentLevel = 1
    End With

    With ws.Range("A3:E3"): .Merge: .Value = modUnicode.U("1. \u0110\u1ecaNH DANH & PH\u00c2N LO\u1ea0I"): .Interior.Color = RGB(47, 85, 151): End With
    With ws.Range("F3:O3"): .Merge: .Value = modUnicode.U("2. H\u00ccNH H\u1eccC, CAO \u0110\u1ed8 & TR\u1eea GIAO (M\u1ede CHO LINKCAD)"): .Interior.Color = RGB(31, 78, 120): End With
    With ws.Range("P3:Q3"): .Merge: .Value = modUnicode.U("3. L\u00d3T RI\u00caNG"): .Interior.Color = RGB(131, 60, 12): End With
    With ws.Range("R3:S3"): .Merge: .Value = modUnicode.U("4. TH\u00c9P M\u1eaaU & BOQ"): .Interior.Color = RGB(56, 87, 35): End With
    With ws.Range("A3:S3")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .HorizontalAlignment = xlCenter: .RowHeight = 20
    End With

    Dim hdrs As Variant
    hdrs = Array( _
        "ElementID", "Building / Zone", "Group", "Mark", modUnicode.U("T\u00ean c\u1ea5u ki\u1ec7n"), _
        "Qty", "L (m)", "W (m)", "H / t (m)", "TopElev (m)", "BotElev (m)", "ConcGrade", _
        "Deduct_V (m3)", "Deduct_FW (m2)", "IsBuried", _
        "LeanThk_Ovr (m)", "LeanExt_Ovr (m)", "RebarPreset", "BOQ_Bucket" _
    )
    Dim c As Long
    For c = 0 To UBound(hdrs): ws.Cells(4, c + 1).Value = hdrs(c): Next c
    With ws.Range("A4:S4")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(22, 54, 92): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True: .RowHeight = 28
    End With

    AddListVal ws.Range("C5:C2000"), "Foundation,GroundBeam,SlabOnGrade,Column,Wall,Beam,Slab,Stair,Lintel,Pile,Manhole,Tank"
    AddListVal ws.Range("O5:O2000"), "Y,N"
    AddListVal ws.Range("R5:R2000"), "FTG_2WAY,BEAM_FULL,BEAM_STD,COL_STAGGER,COL_STD,WALL_2FACE,SLAB_1WAY,SLAB_2WAY_1LAYER,SLAB_2WAY_2LAYER,SLAB_WITH_HATS,STAIR_6BAR,MANHOLE_STD,PILE_HEAD,NONE"
    AddListVal ws.Range("S5:S2000"), "Substructure,Superstructure,Misc/Stair/Lintel,Civil"

    ws.Range("F5:O200").Interior.Color = RGB(255, 250, 220)
    ws.Range("P5:Q200").Interior.Color = RGB(252, 240, 225)
    ws.Range("R5:S200").Interior.Color = RGB(242, 248, 238)

    ' Seed 6 rich starter elements showing Footing(100mm lean), GroundBeam(50mm lean, BEAM_FULL), Slab, Col(stagger), Wall(vertical+horz), Stair
    SeedSampleRow ws, 5, "49-FTG-PC1", "49.ELECTRIC SUPPLY", "Foundation", "PC1", "Footing PC1", 10, 1#, 1#, 1#, 19.6, 18.6, "M400", 0, 0, "Y", "", "", "FTG_2WAY", "Substructure"
    SeedSampleRow ws, 6, "49-GB-GBX1", "49.ELECTRIC SUPPLY", "GroundBeam", "GBX1", "Ground Beam GBX1", 2, 48.6, 0.6, 1#, 19.6, 18.6, "M400", 0, 0, "Y", 0.05, 0.1, "BEAM_FULL", "Substructure"
    SeedSampleRow ws, 7, "49-SLB-GS1", "49.ELECTRIC SUPPLY", "SlabOnGrade", "GS1", "Ground Slab GS1 (250thk)", 1, 40#, 21.1, 0.25, 19.7, 19.45, "M400", 0, 0, "N", "", "", "SLAB_2WAY_2LAYER", "Substructure"
    SeedSampleRow ws, 8, "49-COL-C1", "49.ELECTRIC SUPPLY", "Column", "C1", "Main Column C1", 12, 0.6, 0.6, 4.5, 24.5, 20#, "M400", 0, 1.44, "N", "", "", "COL_STAGGER", "Superstructure"
    SeedSampleRow ws, 9, "49-WAL-W1", "49.ELECTRIC SUPPLY", "Wall", "W1", "Shear Wall W1", 2, 6#, 0.3, 4#, 24#, 20#, "M400", 0, 0, "N", "", "", "WALL_2FACE", "Superstructure"
    SeedSampleRow ws, 10, "49-ST-ST1", "49.ELECTRIC SUPPLY", "Stair", "ST1", "Staircase ST1", 1, 3.2, 1.5, 2.8, 22.8, 20#, "M350", 0, 0, "N", "", "", "STAIR_6BAR", "Misc/Stair/Lintel"

    Dim widths As Variant
    widths = Array(15, 20, 14, 10, 24, 7, 8.5, 8.5, 9, 10, 10, 10, 11, 11, 8, 13, 13, 20, 16)
    For c = 0 To UBound(widths): ws.Columns(c + 1).ColumnWidth = widths(c): Next c

    With ws.Range("A4:S200").Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(217, 217, 217)
    End With

    On Error Resume Next
    If Not Application.ActiveWindow Is Nothing Then
        Application.ActiveWindow.SplitRow = 4
        Application.ActiveWindow.FreezePanes = True
        Application.ActiveWindow.DisplayGridlines = True
    End If
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

Private Sub AddListVal(ByVal rng As Range, ByVal items As String)
    With rng.Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=items
        .IgnoreBlank = True: .InCellDropdown = True
    End With
End Sub

Private Sub SeedSampleRow(ByVal ws As Worksheet, ByVal r As Long, ParamArray vals() As Variant)
    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        If CStr(vals(i)) <> "" Or IsNumeric(vals(i)) Then
            ws.Cells(r, i + 1).Value = vals(i)
        End If
    Next i
End Sub

Public Sub AutoGenerateElementIds(ByVal ws As Worksheet)
    Dim lr As Long: lr = GetLastQsDataRow(ws)
    Dim r As Long, idVal As String, bld As String, grp As String, mk As String
    Dim used As Object: Set used = CreateObject("Scripting.Dictionary"): used.CompareMode = vbTextCompare

    For r = QSDATA_FIRST_ROW To lr
        idVal = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(idVal) > 0 Then used(idVal) = True
    Next r

    For r = QSDATA_FIRST_ROW To lr
        mk = Trim$(CStr(ws.Cells(r, 4).Value))
        idVal = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(mk) > 0 And Len(idVal) = 0 Then
            bld = Trim$(CStr(ws.Cells(r, 2).Value))
            If Len(bld) = 0 Then bld = "BLD"
            Dim pfx As String: pfx = Left$(Replace(Split(bld, ".")(0), " ", ""), 6)
            grp = UCase$(Left$(Trim$(CStr(ws.Cells(r, 3).Value)), 3))
            If Len(grp) = 0 Then grp = "ELM"
            Dim cand As String: cand = pfx & "-" & grp & "-" & Replace(mk, " ", "")
            Dim seq As Long: seq = 1
            Dim finalId As String: finalId = cand
            Do While used.Exists(finalId)
                seq = seq + 1: finalId = cand & "-" & seq
            Loop
            used(finalId) = True
            ws.Cells(r, 1).Value = finalId
        End If
    Next r
End Sub

Public Function ValidateQsData(ByVal ws As Worksheet, ByRef errMsg As String) As Boolean
    AutoGenerateElementIds ws
    Dim lr As Long: lr = GetLastQsDataRow(ws)
    If lr < QSDATA_FIRST_ROW Then
        errMsg = modUnicode.U("B\u1ea3ng QS_Data ch\u01b0a c\u00f3 d\u1eef li\u1ec7u c\u1ea5u ki\u1ec7n n\u00e0o!")
        ValidateQsData = False: Exit Function
    End If

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    Dim r As Long, idVal As String, bld As String, grp As String, mk As String
    Dim validCount As Long: validCount = 0

    For r = QSDATA_FIRST_ROW To lr
        idVal = Trim$(CStr(ws.Cells(r, 1).Value))
        bld = Trim$(CStr(ws.Cells(r, 2).Value))
        grp = Trim$(CStr(ws.Cells(r, 3).Value))
        mk = Trim$(CStr(ws.Cells(r, 4).Value))
        If Len(idVal) = 0 And Len(mk) = 0 Then GoTo NextRow
        If Len(bld) = 0 Or Len(grp) = 0 Or Len(mk) = 0 Then
            errMsg = modUnicode.U("D\u00f2ng ") & r & modUnicode.U(" tr\u00ean QS_Data thi\u1ebfu Building / Group / Mark!")
            ValidateQsData = False: Exit Function
        End If
        If seen.Exists(idVal) Then
            errMsg = modUnicode.U("Tr\u00f9ng ElementID '") & idVal & modUnicode.U("' t\u1ea1i d\u00f2ng ") & r & "!"
            ValidateQsData = False: Exit Function
        End If
        seen(idVal) = r
        validCount = validCount + 1
NextRow:
    Next r

    If validCount = 0 Then
        errMsg = modUnicode.U("Kh\u00f4ng c\u00f3 d\u00f2ng c\u1ea5u ki\u1ec7n h\u1ee3p l\u1ec7 tr\u00ean QS_Data!")
        ValidateQsData = False: Exit Function
    End If
    ValidateQsData = True
End Function

Public Function GetLastQsDataRow(ByVal ws As Worksheet) As Long
    Dim r1 As Long, r2 As Long
    r1 = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    r2 = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    GetLastQsDataRow = IIf(r1 > r2, r1, r2)
End Function

Public Function SerializeQsDataToJson(ByVal ws As Worksheet, Optional ByVal filterBuilding As String = "") As String
    Dim lr As Long: lr = GetLastQsDataRow(ws)
    If lr < QSDATA_FIRST_ROW Then SerializeQsDataToJson = "[]": Exit Function

    Dim arr As Variant
    arr = ws.Range(ws.Cells(QSDATA_FIRST_ROW, 1), ws.Cells(lr, 19)).Value2

    Dim items() As String, cnt As Long, cap As Long
    cap = 64: cnt = 0: ReDim items(0 To cap - 1)

    Dim i As Long, bld As String, eid As String, rowNum As Long
    For i = 1 To UBound(arr, 1)
        rowNum = QSDATA_FIRST_ROW + i - 1
        eid = Trim$(CStr(arr(i, 1)))
        bld = Trim$(CStr(arr(i, 2)))
        If Len(eid) > 0 And Len(bld) > 0 Then
            If Len(filterBuilding) = 0 Or StrComp(bld, filterBuilding, vbTextCompare) = 0 Then
                Dim s As String
                s = SerializeElementRowFromArr(arr, i, rowNum)
                If cnt >= cap Then cap = cap * 2: ReDim Preserve items(0 To cap - 1)
                items(cnt) = s: cnt = cnt + 1
            End If
        End If
    Next i

    If cnt = 0 Then
        SerializeQsDataToJson = "[]"
    Else
        ReDim Preserve items(0 To cnt - 1)
        SerializeQsDataToJson = "[" & Join(items, ",") & "]"
    End If
End Function

Public Function SerializeSingleElementToJson(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    If rowNum < QSDATA_FIRST_ROW Then SerializeSingleElementToJson = "{}": Exit Function
    Dim arr As Variant: arr = ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 19)).Value2
    SerializeSingleElementToJson = SerializeElementRowFromArr(arr, 1, rowNum)
End Function

Private Function SerializeElementRowFromArr(ByVal arr As Variant, ByVal i As Long, ByVal rowNum As Long) As String
    Dim eid As String: eid = Trim$(CStr(arr(i, 1)))
    Dim bld As String: bld = Trim$(CStr(arr(i, 2)))
    Dim s As String
    s = "{"
    s = s & """dataRow"":" & rowNum & ","
    s = s & """elementId"":" & modEngineBridge.JsonStr(eid) & ","
    s = s & """building"":" & modEngineBridge.JsonStr(bld) & ","
    s = s & """group"":" & modEngineBridge.JsonStr(arr(i, 3)) & ","
    s = s & """mark"":" & modEngineBridge.JsonStr(arr(i, 4)) & ","
    s = s & """description"":" & modEngineBridge.JsonStr(arr(i, 5)) & ","
    s = s & """qty"":" & modEngineBridge.JsonNum(arr(i, 6)) & ","
    s = s & """lengthM"":" & modEngineBridge.JsonNum(arr(i, 7)) & ","
    s = s & """widthM"":" & modEngineBridge.JsonNum(arr(i, 8)) & ","
    s = s & """heightM"":" & modEngineBridge.JsonNum(arr(i, 9)) & ","
    s = s & """topElevM"":" & modEngineBridge.JsonNum(arr(i, 10)) & ","
    s = s & """botElevM"":" & modEngineBridge.JsonNum(arr(i, 11)) & ","
    s = s & """concGrade"":" & modEngineBridge.JsonStr(arr(i, 12)) & ","
    s = s & """clashDeductVolM3"":" & modEngineBridge.JsonNum(arr(i, 13)) & ","
    s = s & """clashDeductFwM2"":" & modEngineBridge.JsonNum(arr(i, 14)) & ","
    s = s & """isBuried"":" & modEngineBridge.JsonStr(arr(i, 15)) & ","
    If IsNumeric(arr(i, 16)) And Not IsEmpty(arr(i, 16)) Then
        s = s & """leanThkOvrM"":" & modEngineBridge.JsonNum(arr(i, 16)) & ","
    End If
    If IsNumeric(arr(i, 17)) And Not IsEmpty(arr(i, 17)) Then
        s = s & """leanExtOvrM"":" & modEngineBridge.JsonNum(arr(i, 17)) & ","
    End If
    s = s & """rebarTemplate"":" & modEngineBridge.JsonStr(arr(i, 18)) & ","
    s = s & """boqBucket"":" & modEngineBridge.JsonStr(arr(i, 19))
    s = s & "}"
    SerializeElementRowFromArr = s
End Function
