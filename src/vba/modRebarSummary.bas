Attribute VB_Name = "modRebarSummary"
Option Explicit

' ==============================================================================
' modRebarSummary - Cross-Sheet Steel Summary & 1D Bar Cutting Optimizer (FFD)
' ==============================================================================

Public Sub SummarizeAllSteel(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Dim tktSheets As New Collection
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If modRebarTKT.IsTKTSheet(ws) Then
            tktSheets.Add ws
        End If
    Next ws

    Dim hasShape As Boolean, hasPlate As Boolean
    On Error Resume Next
    hasShape = Not (wb.Worksheets("TK_ThepHinh") Is Nothing)
    hasPlate = Not (wb.Worksheets("TK_ThepTam") Is Nothing)
    On Error GoTo 0

    If tktSheets.Count = 0 And Not hasShape And Not hasPlate Then
        Application.ScreenUpdating = True
        If showMsg Then
            modUnicode.MsgBoxU modUnicode.U("Kh\u00f4ng t\u00ecm th\u1ea5y sheet TKT_*, TK_ThepHinh ho\u1eb7c TK_ThepTam n\u00e0o trong file hi\u1ec7n t\u1ea1i!"), vbExclamation
        End If
        Exit Sub
    End If

    Dim sumWs As Worksheet
    On Error Resume Next
    Set sumWs = wb.Worksheets("TH_Thep_ToanCT")
    On Error GoTo 0
    If sumWs Is Nothing Then
        Set sumWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        sumWs.Name = "TH_Thep_ToanCT"
    Else
        sumWs.Cells.Clear
    End If

    sumWs.Cells.Font.Name = "Segoe UI"
    sumWs.Cells.Font.Size = 9.5

    With sumWs.Range("B1:S1")
        .Merge
        .Value = modUnicode.U("B\u1ea2NG T\u1ed4NG H\u1ee2P KH\u1ed0I L\u01af\u1ee2NG TH\u00c9P TO\u00c0N C\u00d4NG TR\u00ccNH")
        .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(22, 54, 92): .RowHeight = 30
        .IndentLevel = 1: .VerticalAlignment = xlCenter
    End With

    sumWs.Range("B2").Formula = "=IFERROR(CFG_ProjectName,"""")"
    sumWs.Range("B2").Font.Bold = True

    Dim r As Long
    r = 4

    ' --- PART 1: Cross-sheet Rebar Summary by Component Sheet & Diameter ---
    Dim dias As Variant
    dias = Array(6, 8, 10, 12, 14, 16, 18, 20, 22, 25, 28, 32, 36, 40)

    With sumWs.Range("B" & r & ":S" & r)
        .Merge
        .Value = modUnicode.U("I. T\u1ed4NG H\u1ee2P C\u1ed0T TH\u00c9P XUY\u00caN SHEET THEO \u0110\u01af\u1edcNG K\u00cdNH (KG)")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 85, 151): .RowHeight = 22
    End With
    r = r + 1

    sumWs.Cells(r, 2).Value = "STT"
    sumWs.Cells(r, 3).Value = modUnicode.U("H\u1ea1ng m\u1ee5c / Sheet TKT")
    Dim i As Long, cIdx As Long
    For i = LBound(dias) To UBound(dias)
        cIdx = 4 + i
        sumWs.Cells(r, cIdx).Value = dias(i)
        sumWs.Cells(r, cIdx).NumberFormat = """" & ChrW$(216) & """0"
    Next i
    sumWs.Cells(r, 18).Value = "Coupler"
    sumWs.Cells(r, 19).Value = modUnicode.U("T\u1ed5ng (kg)")
    With sumWs.Range("B" & r & ":S" & r)
        .Font.Bold = True: .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(217, 225, 242): .RowHeight = 24
    End With

    Dim dataStart As Long: dataStart = r + 1
    Dim idx As Long
    For idx = 1 To tktSheets.Count
        Dim sh As Worksheet
        Set sh = tktSheets(idx)
        r = dataStart + idx - 1
        sumWs.Cells(r, 2).Value = idx
        sumWs.Cells(r, 2).HorizontalAlignment = xlCenter
        sumWs.Cells(r, 3).Value = sh.Name
        sumWs.Hyperlinks.Add Anchor:=sumWs.Cells(r, 3), Address:="", SubAddress:="'" & sh.Name & "'!A1", TextToDisplay:=sh.Name

        For i = LBound(dias) To UBound(dias)
            cIdx = 4 + i
            Dim colL As String
            colL = Split(sumWs.Cells(1, cIdx).Address(True, False), "$")(0)
            sumWs.Cells(r, cIdx).Formula = "=SUMIF('" & sh.Name & "'!$O$11:$O$5000," & colL & "$" & (dataStart - 1) & ",'" & sh.Name & "'!$W$11:$W$5000)"
            sumWs.Cells(r, cIdx).NumberFormat = "#,##0.00"
        Next i
        sumWs.Cells(r, 18).Formula = "=SUM('" & sh.Name & "'!$T$11:$T$5000)"
        sumWs.Cells(r, 18).NumberFormat = "#,##0"
        sumWs.Cells(r, 19).Formula = "=SUM(D" & r & ":Q" & r & ")"
        sumWs.Cells(r, 19).NumberFormat = "#,##0.00": sumWs.Cells(r, 19).Font.Bold = True
    Next idx

    Dim dataEnd As Long
    dataEnd = dataStart + IIf(tktSheets.Count > 0, tktSheets.Count - 1, 0)
    r = dataEnd + 1

    ' Totals row
    sumWs.Cells(r, 3).Value = modUnicode.U("T\u1ed4NG C\u1ed8NG (kg)")
    For cIdx = 4 To 19
        Dim cl As String: cl = Split(sumWs.Cells(1, cIdx).Address(True, False), "$")(0)
        If tktSheets.Count > 0 Then
            sumWs.Cells(r, cIdx).Formula = "=SUM(" & cl & dataStart & ":" & cl & dataEnd & ")"
        Else
            sumWs.Cells(r, cIdx).Value = 0
        End If
        sumWs.Cells(r, cIdx).NumberFormat = IIf(cIdx = 18, "#,##0", "#,##0.00")
    Next cIdx
    With sumWs.Range("B" & r & ":S" & r)
        .Font.Bold = True: .Interior.Color = RGB(255, 242, 204)
    End With
    Dim totRebarRow As Long: totRebarRow = r

    ' Row: Equivalent standard bars (11.7m / 12m) for D >= 10
    r = r + 1
    sumWs.Cells(r, 3).Value = modUnicode.U("S\u1ed1 c\u00e2y quy \u0111\u1ed5i (11.7/12m)")
    For i = LBound(dias) To UBound(dias)
        cIdx = 4 + i
        cl = Split(sumWs.Cells(1, cIdx).Address(True, False), "$")(0)
        If dias(i) <= 8 Then
            sumWs.Cells(r, cIdx).Value = modUnicode.U("Cu\u1ed9n")
            sumWs.Cells(r, cIdx).HorizontalAlignment = xlCenter
        Else
            sumWs.Cells(r, cIdx).Formula = "=ROUNDUP(" & cl & totRebarRow & "/(PI()*(" & cl & "$" & (dataStart - 1) & "/1000)^2/4*7850)/(IFERROR(CFG_BarStdLen,11700)/1000),0)"
            sumWs.Cells(r, cIdx).NumberFormat = "#,##0"
        End If
    Next i
    sumWs.Range("B" & r & ":S" & r).Font.Italic = True

    With sumWs.Range("B" & (dataStart - 1) & ":S" & r).Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(180, 180, 180)
    End With

    ' --- PART 2: Forecast Group Breakdown (<=10, <=18, >18) ---
    r = r + 2
    With sumWs.Range("B" & r & ":H" & r)
        .Merge
        .Value = modUnicode.U("II. PH\u00c2N NH\u00d3M TH\u00c9P THEO D\u1ef0 TO\u00c1N (<= \u00d810 | \u00d812-\u00d818 | > \u00d818)")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 85, 151): .RowHeight = 22
    End With
    r = r + 1
    sumWs.Cells(r, 2).Value = "STT"
    sumWs.Cells(r, 3).Value = modUnicode.U("Nh\u00f3m \u0111\u01b0\u1eddng k\u00ednh")
    sumWs.Cells(r, 4).Value = modUnicode.U("L\u00fd thuy\u1ebft (kg)")
    sumWs.Cells(r, 5).Value = modUnicode.U("L\u00fd thuy\u1ebft (T\u1ea5n)")
    sumWs.Cells(r, 6).Value = modUnicode.U("C\u00f3 hao h\u1ee5t (T\u1ea5n)")
    sumWs.Cells(r, 7).Value = modUnicode.U("T\u1ef7 tr\u1ecdng (%)")
    sumWs.Cells(r, 8).Value = modUnicode.U("Ghi ch\u00fa")
    With sumWs.Range("B" & r & ":H" & r)
        .Font.Bold = True: .HorizontalAlignment = xlCenter: .Interior.Color = RGB(217, 225, 242)
    End With

    Dim grpStart As Long: grpStart = r + 1
    AddGroupRow sumWs, grpStart, 1, modUnicode.U("C\u1ed1t th\u00e9p <= \u00d810"), "=SUM(D" & totRebarRow & ":F" & totRebarRow & ")", grpStart, modUnicode.U("Cu\u1ed9n \u00d86-\u00d88 + C\u00e2y \u00d810")
    AddGroupRow sumWs, grpStart + 1, 2, modUnicode.U("C\u1ed1t th\u00e9p \u00d812 - \u00d818"), "=SUM(G" & totRebarRow & ":J" & totRebarRow & ")", grpStart, modUnicode.U("Th\u00e9p thanh v\u1eb1n trung b\u00ecnh")
    AddGroupRow sumWs, grpStart + 2, 3, modUnicode.U("C\u1ed1t th\u00e9p > \u00d818"), "=SUM(K" & totRebarRow & ":Q" & totRebarRow & ")", grpStart, modUnicode.U("Th\u00e9p ch\u1ecbu l\u1ef1c ch\u00ednh")

    r = grpStart + 3
    sumWs.Cells(r, 3).Value = modUnicode.U("T\u1ed4NG C\u1ed0T TH\u00c9P")
    sumWs.Cells(r, 4).Formula = "=SUM(D" & grpStart & ":D" & (grpStart + 2) & ")"
    sumWs.Cells(r, 5).Formula = "=SUM(E" & grpStart & ":E" & (grpStart + 2) & ")"
    sumWs.Cells(r, 6).Formula = "=SUM(F" & grpStart & ":F" & (grpStart + 2) & ")"
    sumWs.Cells(r, 7).Formula = "=SUM(G" & grpStart & ":G" & (grpStart + 2) & ")"
    sumWs.Cells(r, 4).NumberFormat = "#,##0.00"
    sumWs.Cells(r, 5).NumberFormat = "#,##0.000"
    sumWs.Cells(r, 6).NumberFormat = "#,##0.000"
    sumWs.Cells(r, 7).NumberFormat = "0.0%"
    With sumWs.Range("B" & r & ":H" & r)
        .Font.Bold = True: .Interior.Color = RGB(255, 242, 204)
    End With
    With sumWs.Range("B" & (grpStart - 1) & ":H" & r).Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(180, 180, 180)
    End With

    ' --- PART 3: Structural Steel, Inox, Mesh & Plates Summary ---
    r = r + 2
    With sumWs.Range("B" & r & ":H" & r)
        .Merge
        .Value = modUnicode.U("III. T\u1ed4NG H\u1ee2P TH\u00c9P H\u00ccNH, INOX, L\u01af\u1edaI H\u00c0N & B\u1ea2N M\u00c3")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 85, 151): .RowHeight = 22
    End With
    r = r + 1
    sumWs.Cells(r, 2).Value = "STT"
    sumWs.Cells(r, 3).Value = modUnicode.U("Lo\u1ea1i v\u1eadt li\u1ec7u")
    sumWs.Cells(r, 4).Value = modUnicode.U("Kh\u1ed1i l\u01b0\u1ee3ng (kg)")
    sumWs.Cells(r, 5).Value = modUnicode.U("Kh\u1ed1i l\u01b0\u1ee3ng (T\u1ea5n)")
    sumWs.Cells(r, 6).Value = modUnicode.U("Ngu\u1ed3n sheet")
    sumWs.Range("B" & r & ":H" & r).Font.Bold = True
    sumWs.Range("B" & r & ":H" & r).Interior.Color = RGB(217, 225, 242)

    r = r + 1
    sumWs.Cells(r, 2).Value = 1: sumWs.Cells(r, 3).Value = modUnicode.U("Th\u00e9p h\u00ecnh, x\u00e0 g\u1ed3, inox, l\u01b0\u1edbi h\u00e0n")
    If hasShape Then
        sumWs.Cells(r, 4).Formula = "=IFERROR(TK_ThepHinh!L2,0)"
    Else
        sumWs.Cells(r, 4).Value = 0
    End If
    sumWs.Cells(r, 5).Formula = "=D" & r & "/1000"
    sumWs.Cells(r, 6).Value = "TK_ThepHinh"
    sumWs.Cells(r, 4).NumberFormat = "#,##0.00": sumWs.Cells(r, 5).NumberFormat = "#,##0.000"

    r = r + 1
    sumWs.Cells(r, 2).Value = 2: sumWs.Cells(r, 3).Value = modUnicode.U("Th\u00e9p t\u1ea5m & B\u1ea3n m\u00e3")
    If hasPlate Then
        sumWs.Cells(r, 4).Formula = "=IFERROR(TK_ThepTam!M2,0)"
    Else
        sumWs.Cells(r, 4).Value = 0
    End If
    sumWs.Cells(r, 5).Formula = "=D" & r & "/1000"
    sumWs.Cells(r, 6).Value = "TK_ThepTam"
    sumWs.Cells(r, 4).NumberFormat = "#,##0.00": sumWs.Cells(r, 5).NumberFormat = "#,##0.000"

    ' --- PART 4: Summary by BOQ_Bucket (Col Z on TKT_*) for Direct BOQ Linking ---
    r = r + 2
    With sumWs.Range("B" & r & ":H" & r)
        .Merge
        .Value = modUnicode.U("IV. PH\u00c2N B\u1ed4 KH\u1ed0I L\u01af\u1ee2NG TH\u00c9P THEO H\u1ea0NG M\u1ee4C BOQ (SUBSTRUCTURE / SUPERSTRUCTURE / STAIR / CIVIL)")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 85, 151): .RowHeight = 22
    End With
    r = r + 1
    sumWs.Cells(r, 2).Value = "STT"
    sumWs.Cells(r, 3).Value = "BOQ_Bucket"
    sumWs.Cells(r, 4).Value = modUnicode.U("<= \u00d810 (kg)")
    sumWs.Cells(r, 5).Value = modUnicode.U("\u00d812 - \u00d818 (kg)")
    sumWs.Cells(r, 6).Value = modUnicode.U("> \u00d818 (kg)")
    sumWs.Cells(r, 7).Value = modUnicode.U("T\u1ed5ng l\u00fd thuy\u1ebft (T\u1ea5n)")
    sumWs.Cells(r, 8).Value = modUnicode.U("C\u00f3 hao h\u1ee5t Spec (T\u1ea5n)")
    With sumWs.Range("B" & r & ":H" & r)
        .Font.Bold = True: .HorizontalAlignment = xlCenter: .Interior.Color = RGB(217, 225, 242)
    End With

    Dim buckets As Variant
    buckets = Array("Substructure", "Superstructure", "Misc/Stair/Lintel", "Civil")
    Dim bkIdx As Long
    For bkIdx = LBound(buckets) To UBound(buckets)
        r = r + 1
        sumWs.Cells(r, 2).Value = bkIdx + 1
        sumWs.Cells(r, 2).HorizontalAlignment = xlCenter
        sumWs.Cells(r, 3).Value = buckets(bkIdx)
        If tktSheets.Count > 0 Then
            Dim sSmall As String, sMed As String, sLg As String
            sSmall = "": sMed = "": sLg = ""
            For idx = 1 To tktSheets.Count
                Set sh = tktSheets(idx)
                If idx > 1 Then sSmall = sSmall & "+": sMed = sMed & "+": sLg = sLg & "+"
                sSmall = sSmall & "SUMIFS('" & sh.Name & "'!$W$11:$W$5000,'" & sh.Name & "'!$Z$11:$Z$5000,$C" & r & ",'" & sh.Name & "'!$O$11:$O$5000,""<=""&10)"
                sMed = sMed & "SUMIFS('" & sh.Name & "'!$W$11:$W$5000,'" & sh.Name & "'!$Z$11:$Z$5000,$C" & r & ",'" & sh.Name & "'!$O$11:$O$5000,"">10"",'" & sh.Name & "'!$O$11:$O$5000,""<=""&18)"
                sLg = sLg & "SUMIFS('" & sh.Name & "'!$W$11:$W$5000,'" & sh.Name & "'!$Z$11:$Z$5000,$C" & r & ",'" & sh.Name & "'!$O$11:$O$5000,"">18"")"
            Next idx
            sumWs.Cells(r, 4).Formula = "=" & sSmall
            sumWs.Cells(r, 5).Formula = "=" & sMed
            sumWs.Cells(r, 6).Formula = "=" & sLg
        Else
            sumWs.Cells(r, 4).Value = 0: sumWs.Cells(r, 5).Value = 0: sumWs.Cells(r, 6).Value = 0
        End If
        sumWs.Cells(r, 7).Formula = "=SUM(D" & r & ":F" & r & ")/1000"
        sumWs.Cells(r, 8).Formula = "=(D" & r & "*(1+IFERROR(SPEC_Scrap_Small,3)/100)+E" & r & "*(1+IFERROR(SPEC_Scrap_Med,2.5)/100)+F" & r & "*(1+IFERROR(SPEC_Scrap_Large,2)/100))/1000"
        sumWs.Cells(r, 4).NumberFormat = "#,##0.00"
        sumWs.Cells(r, 5).NumberFormat = "#,##0.00"
        sumWs.Cells(r, 6).NumberFormat = "#,##0.00"
        sumWs.Cells(r, 7).NumberFormat = "#,##0.000"
        sumWs.Cells(r, 8).NumberFormat = "#,##0.000"
        sumWs.Cells(r, 8).Font.Bold = True
    Next bkIdx

    sumWs.Columns("A").ColumnWidth = 3
    sumWs.Columns("B").ColumnWidth = 6
    sumWs.Columns("C").ColumnWidth = 26
    sumWs.Columns("D:S").ColumnWidth = 11.5
    sumWs.Columns("S").ColumnWidth = 14

    sumWs.Activate
    Application.ScreenUpdating = True
    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 t\u1ed5ng h\u1ee3p th\u00e9p to\u00e0n c\u00f4ng tr\u00ecnh t\u1ea1i sheet TH_Thep_ToanCT (li\u00ean k\u1ebft c\u00f4ng th\u1ee9c tr\u1ef1c ti\u1ebfp t\u1edbi ") & tktSheets.Count & modUnicode.U(" sheet TKT)!"), vbInformation
    End If
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    modUnicode.MsgBoxU "SummarizeAllSteel Error: " & Err.Description, vbCritical
End Sub

Private Sub AddGroupRow(ByVal ws As Worksheet, ByVal r As Long, ByVal stt As Long, _
                        ByVal grpLabel As String, ByVal kgFormula As String, _
                        ByVal grpStart As Long, ByVal noteStr As String)
    ws.Cells(r, 2).Value = stt: ws.Cells(r, 2).HorizontalAlignment = xlCenter
    ws.Cells(r, 3).Value = grpLabel
    ws.Cells(r, 4).Formula = kgFormula
    ws.Cells(r, 4).NumberFormat = "#,##0.00"
    ws.Cells(r, 5).Formula = "=D" & r & "/1000"
    ws.Cells(r, 5).NumberFormat = "#,##0.000"
    ws.Cells(r, 6).Formula = "=E" & r & "*(1+IFERROR(CFG_RebarWastePct,0)/100)"
    ws.Cells(r, 6).NumberFormat = "#,##0.000"
    ws.Cells(r, 7).Formula = "=IFERROR(D" & r & "/D$" & (grpStart + 3) & ",0)"
    ws.Cells(r, 7).NumberFormat = "0.0%"
    ws.Cells(r, 8).Value = noteStr
End Sub

' ==============================================================================
' 1D Bar Cutting Optimization (First-Fit Decreasing with Pattern Grouping)
' ==============================================================================
Public Sub OptimizeBarCutting1D(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Dim stdLen As Long
    stdLen = 11700
    On Error Resume Next
    If CLng(wb.Names("CFG_BarStdLen").RefersToRange.Value) > 1000 Then
        stdLen = CLng(wb.Names("CFG_BarStdLen").RefersToRange.Value)
    End If
    On Error GoTo 0

    ' Collect cut pieces per diameter (D >= 10) across all TKT_* sheets
    Dim dictDia As Object
    Set dictDia = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet, r As Long, lr As Long
    Dim d As Long, nBars As Long, barLenMm As Long
    Dim fullCount As Long, remMm As Long, i As Long

    For Each ws In wb.Worksheets
        If modRebarTKT.IsTKTSheet(ws) Then
            lr = modRebarTKT.GetLastTKTRow(ws)
            For r = modRebarTKT.TKT_FIRST_DATA_ROW To lr
                d = CLng(Val(ws.Cells(r, 15).Value))
                nBars = CLng(Val(ws.Cells(r, 18).Value))
                barLenMm = CLng(Round(Val(ws.Cells(r, 19).Value) * 1000#, 0))
                If d >= 10 And nBars > 0 And barLenMm > 0 Then
                    If Not dictDia.Exists(d) Then
                        dictDia.Add d, New Collection
                    End If
                    Dim col As Collection
                    Set col = dictDia(d)
                    ' Split bars > stdLen into full stdLen bars + remainder piece
                    If barLenMm > stdLen Then
                        fullCount = barLenMm \ stdLen
                        remMm = barLenMm Mod stdLen
                        For i = 1 To nBars * fullCount
                            col.Add stdLen
                        Next i
                        If remMm > 0 Then
                            For i = 1 To nBars
                                col.Add remMm
                            Next i
                        End If
                    Else
                        For i = 1 To nBars
                            col.Add barLenMm
                        Next i
                    End If
                End If
            Next r
        End If
    Next ws

    If dictDia.Count = 0 Then
        Application.ScreenUpdating = True
        If showMsg Then
            modUnicode.MsgBoxU modUnicode.U("Kh\u00f4ng c\u00f3 d\u1eef li\u1ec7u th\u00e9p c\u00e2y (>= \u00d810) trong c\u00e1c sheet TKT_* \u0111\u1ec3 t\u1ed1i \u01b0u c\u1eaft!"), vbExclamation
        End If
        Exit Sub
    End If

    Dim cutWs As Worksheet
    On Error Resume Next
    Set cutWs = wb.Worksheets("TKT_CatToiUu")
    On Error GoTo 0
    If cutWs Is Nothing Then
        Set cutWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        cutWs.Name = "TKT_CatToiUu"
    Else
        cutWs.Cells.Clear
    End If

    cutWs.Cells.Font.Name = "Segoe UI": cutWs.Cells.Font.Size = 9.5
    With cutWs.Range("B1:H1")
        .Merge
        .Value = modUnicode.U("S\u01a0 \u0110\u1ed2 C\u1eaeT TH\u00c9P T\u1ed0I \u01afU 1D (C\u00c2Y CHU\u1ea8N ") & stdLen & " mm)"
        .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(22, 54, 92): .RowHeight = 28: .IndentLevel = 1
    End With

    ' Summary header
    cutWs.Range("B3").Value = modUnicode.U("\u0110\u01b0\u1eddng k\u00ednh")
    cutWs.Range("C3").Value = modUnicode.U("T\u1ed5ng d\u00e0i TK (m)")
    cutWs.Range("D3").Value = modUnicode.U("S\u1ed1 c\u00e2y chu\u1ea9n")
    cutWs.Range("E3").Value = modUnicode.U("T\u1ed5ng d\u00e0i nh\u1eadp (m)")
    cutWs.Range("F3").Value = modUnicode.U("\u0110\u1ec1-x\u00ea th\u1eeba (m)")
    cutWs.Range("G3").Value = modUnicode.U("T\u1ef7 l\u1ec7 hao h\u1ee5t (%)")
    cutWs.Range("H3").Value = modUnicode.U("\u0110\u1ec1-x\u00ea t\u1eadn d\u1ee5ng >=1m (m)")
    With cutWs.Range("B3:H3")
        .Font.Bold = True: .HorizontalAlignment = xlCenter: .Interior.Color = RGB(217, 225, 242)
    End With

    Dim diaKeys As Variant: diaKeys = dictDia.Keys
    SortLongArray diaKeys

    Dim sumRow As Long: sumRow = 4
    Dim patRows As New Collection

    Dim k As Long
    For k = LBound(diaKeys) To UBound(diaKeys)
        d = CLng(diaKeys(k))
        Dim piecesCol As Collection: Set piecesCol = dictDia(d)
        Dim arr() As Long
        ReDim arr(0 To piecesCol.Count - 1)
        Dim totalDesignMm As Double: totalDesignMm = 0
        For i = 1 To piecesCol.Count
            arr(i - 1) = CLng(piecesCol(i))
            totalDesignMm = totalDesignMm + arr(i - 1)
        Next i
        ' Sort descending for FFD
        SortLongDesc arr

        ' Run First-Fit Decreasing (FFD) bin packing
        Dim binRem() As Long, binPat() As String
        Dim binCap As Long: binCap = 64
        ReDim binRem(0 To binCap - 1)
        ReDim binPat(0 To binCap - 1)
        Dim binCount As Long: binCount = 0

        Dim pIdx As Long, bIdx As Long, placed As Boolean, L As Long
        For pIdx = LBound(arr) To UBound(arr)
            L = arr(pIdx)
            placed = False
            For bIdx = 0 To binCount - 1
                If binRem(bIdx) >= L Then
                    binRem(bIdx) = binRem(bIdx) - L
                    binPat(bIdx) = binPat(bIdx) & "+" & L
                    placed = True
                    Exit For
                End If
            Next bIdx
            If Not placed Then
                If binCount >= binCap Then
                    binCap = binCap * 2
                    ReDim Preserve binRem(0 To binCap - 1)
                    ReDim Preserve binPat(0 To binCap - 1)
                End If
                binRem(binCount) = stdLen - L
                binPat(binCount) = CStr(L)
                binCount = binCount + 1
            End If
        Next pIdx

        ' Aggregate identical cutting patterns
        Dim patDict As Object: Set patDict = CreateObject("Scripting.Dictionary")
        Dim patRem As Object: Set patRem = CreateObject("Scripting.Dictionary")
        Dim reusableMm As Double: reusableMm = 0
        For bIdx = 0 To binCount - 1
            If binRem(bIdx) >= 1000 Then reusableMm = reusableMm + binRem(bIdx)
            If patDict.Exists(binPat(bIdx)) Then
                patDict(binPat(bIdx)) = patDict(binPat(bIdx)) + 1
            Else
                patDict.Add binPat(bIdx), 1
                patRem.Add binPat(bIdx), binRem(bIdx)
            End If
        Next bIdx

        Dim totalStockMm As Double: totalStockMm = CDbl(binCount) * CDbl(stdLen)
        Dim wasteMm As Double: wasteMm = totalStockMm - totalDesignMm

        cutWs.Cells(sumRow, 2).Value = ChrW$(216) & d
        cutWs.Cells(sumRow, 3).Value = Round(totalDesignMm / 1000#, 2)
        cutWs.Cells(sumRow, 4).Value = binCount
        cutWs.Cells(sumRow, 5).Value = Round(totalStockMm / 1000#, 2)
        cutWs.Cells(sumRow, 6).Value = Round(wasteMm / 1000#, 2)
        cutWs.Cells(sumRow, 7).Value = IIf(totalStockMm > 0, wasteMm / totalStockMm, 0)
        cutWs.Cells(sumRow, 8).Value = Round(reusableMm / 1000#, 2)
        cutWs.Cells(sumRow, 3).NumberFormat = "#,##0.00"
        cutWs.Cells(sumRow, 4).NumberFormat = "#,##0"
        cutWs.Cells(sumRow, 5).NumberFormat = "#,##0.00"
        cutWs.Cells(sumRow, 6).NumberFormat = "#,##0.00"
        cutWs.Cells(sumRow, 7).NumberFormat = "0.00%"
        cutWs.Cells(sumRow, 8).NumberFormat = "#,##0.00"
        sumRow = sumRow + 1

        Dim pKey As Variant
        For Each pKey In patDict.Keys
            patRows.Add Array(ChrW$(216) & d, patDict(pKey), CStr(pKey), CLng(patRem(pKey)))
        Next pKey
    Next k

    ' Cutting patterns detail table
    Dim pStart As Long: pStart = sumRow + 2
    cutWs.Cells(pStart - 1, 2).Value = modUnicode.U("CHI TI\u1ebeT PH\u01af\u01a0NG \u00c1N C\u1eaeT TH\u00c9P T\u1ea0I C\u00d4NG TR\u01af\u1edcNG")
    cutWs.Cells(pStart - 1, 2).Font.Bold = True: cutWs.Cells(pStart - 1, 2).Font.Color = RGB(31, 78, 120)

    cutWs.Cells(pStart, 2).Value = modUnicode.U("\u0110\u01b0\u1eddng k\u00ednh")
    cutWs.Cells(pStart, 3).Value = modUnicode.U("S\u1ed1 c\u00e2y c\u1eaft")
    cutWs.Cells(pStart, 4).Value = modUnicode.U("S\u01a1 \u0111\u1ed3 c\u1eaft tr\u00ean 1 c\u00e2y (mm)")
    cutWs.Cells(pStart, 5).Value = modUnicode.U("D\u01b0 / c\u00e2y (mm)")
    With cutWs.Range("B" & pStart & ":E" & pStart)
        .Font.Bold = True: .Interior.Color = RGB(47, 85, 151): .Font.Color = RGB(255, 255, 255)
    End With

    Dim pr As Long, item As Variant
    For pr = 1 To patRows.Count
        item = patRows(pr)
        cutWs.Cells(pStart + pr, 2).Value = item(0)
        cutWs.Cells(pStart + pr, 3).Value = item(1)
        cutWs.Cells(pStart + pr, 4).Value = item(2)
        cutWs.Cells(pStart + pr, 5).Value = item(3)
    Next pr

    cutWs.Columns("B:H").AutoFit
    cutWs.Columns("D").ColumnWidth = 48
    cutWs.Activate
    Application.ScreenUpdating = True
    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 ch\u1ea1y t\u1ed1i \u01b0u c\u1eaft th\u00e9p 1D xong! Xem chi ti\u1ebft s\u01a1 \u0111\u1ed3 c\u1eaft & t\u1ef7 l\u1ec7 hao h\u1ee5t t\u1ea1i sheet TKT_CatToiUu."), vbInformation
    End If
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    modUnicode.MsgBoxU "OptimizeBarCutting1D Error: " & Err.Description, vbCritical
End Sub

Private Sub SortLongArray(ByRef arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CLng(arr(i)) > CLng(arr(j)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Sub SortLongDesc(ByRef arr() As Long)
    QuickSortDesc arr, LBound(arr), UBound(arr)
End Sub

Private Sub QuickSortDesc(ByRef arr() As Long, ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub
    Dim pivot As Long: pivot = arr((lo + hi) \ 2)
    Dim i As Long: i = lo
    Dim j As Long: j = hi
    Dim tmp As Long
    Do While i <= j
        Do While arr(i) > pivot: i = i + 1: Loop
        Do While arr(j) < pivot: j = j - 1: Loop
        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortDesc arr, lo, j
    If i < hi Then QuickSortDesc arr, i, hi
End Sub
