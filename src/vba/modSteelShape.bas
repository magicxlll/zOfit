Attribute VB_Name = "modSteelShape"
Option Explicit

' ==============================================================================
' modSteelShape - Structural Steel Shapes, Inox, Wire Mesh & Gusset Plates
' ==============================================================================

Public Sub CreateSteelShapeSheet(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    ' Ensure DB_VatLieu exists in active workbook so INDEX/MATCH formulas are standalone
    Dim hasDB As Boolean
    On Error Resume Next
    hasDB = Not (wb.Worksheets("DB_VatLieu") Is Nothing)
    On Error GoTo 0
    If Not hasDB Then
        modCatalog.ExportCatalogToActiveBook False
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("TK_ThepHinh")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "TK_ThepHinh"
        BuildSteelShapeLayout ws
    End If

    ws.Activate
    Application.ScreenUpdating = True
    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 t\u1ea1o b\u1ea3ng Th\u1ed1ng k\u00ea Th\u00e9p H\u00ecnh / Inox / L\u01b0\u1edbi h\u00e0n: TK_ThepHinh.") & vbCrLf & _
            modUnicode.U("M\u1eb9o: Ch\u1ecdn \u00f4 c\u1ed9t 'Quy c\u00e1ch v\u1eadt li\u1ec7u' (c\u1ed9t D) r\u1ed3i m\u1edf 'Tra c\u1ee9u Th\u00e9p & Inox' \u0111\u1ec3 ch\u00e8n nhanh quy c\u00e1ch chu\u1ea9n!"), vbInformation
    End If
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    modUnicode.MsgBoxU "CreateSteelShapeSheet Error: " & Err.Description, vbCritical
End Sub

Private Sub BuildSteelShapeLayout(ByVal ws As Worksheet)
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 9.5

    With ws.Range("A1:N1")
        .Interior.Color = RGB(22, 54, 92)
        .RowHeight = 28
    End With
    With ws.Range("B1:N1")
        .Merge
        .Value = modUnicode.U("B\u1ea2NG TH\u1ed0NG K\u00ca TH\u00c9P H\u00ccNH, X\u00c0 G\u1ed2, INOX & L\u01af\u1edaI TH\u00c9P H\u00c0N")
        .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
    End With

    ws.Range("B2").Value = modUnicode.U("D\u1ef1 \u00e1n:")
    ws.Range("B2").Font.Bold = True
    ws.Range("C2:G2").Merge
    ws.Range("C2").Formula = "=IFERROR(CFG_ProjectName,"""")"

    ws.Range("K2").Value = modUnicode.U("T\u1ed4NG KL (kg):")
    ws.Range("K2").Font.Bold = True
    ws.Range("L2").Formula = "=SUM(K6:K2000)"
    ws.Range("L2").NumberFormat = "#,##0.00"
    ws.Range("L2").Font.Bold = True
    ws.Range("L2").Interior.Color = RGB(255, 242, 204)

    ws.Range("M2").Value = modUnicode.U("T\u1ea5n:")
    ws.Range("M2").Font.Bold = True
    ws.Range("N2").Formula = "=L2/1000"
    ws.Range("N2").NumberFormat = "#,##0.000"
    ws.Range("N2").Font.Bold = True

    Dim hdrs As Variant
    hdrs = Array( _
        "STT", _
        modUnicode.U("H\u1ea1ng m\u1ee5c / C\u1ea5u ki\u1ec7n"), _
        modUnicode.U("S\u1ed1 hi\u1ec7u"), _
        modUnicode.U("Quy c\u00e1ch v\u1eadt li\u1ec7u"), _
        modUnicode.U("\u0110VT"), _
        modUnicode.U("D\u00e0i (m) / DT (m2)"), _
        modUnicode.U("SL / CK"), _
        modUnicode.U("S\u1ed1 CK"), _
        modUnicode.U("T\u1ed5ng d\u00e0i / DT"), _
        modUnicode.U("KL \u0111\u01a1n v\u1ecb (kg)"), _
        modUnicode.U("T\u1ed5ng KL (kg)"), _
        modUnicode.U("\u0110\u01a1n gi\u00e1 (\u0111/kg)"), _
        modUnicode.U("Th\u00e0nh ti\u1ec1n (\u0111)"), _
        modUnicode.U("Nh\u00f3m / Ghi ch\u00fa") _
    )
    Dim i As Long
    For i = LBound(hdrs) To UBound(hdrs)
        ws.Cells(5, i + 1).Value = hdrs(i)
    Next i
    With ws.Range("A5:N5")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 85, 151)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True: .RowHeight = 28
    End With

    ' Seed 4 practical sample rows linked dynamically via INDEX/MATCH to DB_VatLieu
    SeedShapeRow ws, 6, 1, modUnicode.U("C\u1ed9t k\u00e8o th\u00e9p"), "CK1", "I 200x100", 6, 1, 4
    SeedShapeRow ws, 7, 2, modUnicode.U("X\u00e0 g\u1ed3 m\u00e1i"), "XG1", "C150x50x20x2.0", 6, 10, 2
    SeedShapeRow ws, 8, 3, modUnicode.U("Lan can Inox"), "LC1", "Inox tr\u00f2n \u00d838.1 x 1.2mm", 12, 1, 1
    SeedShapeRow ws, 9, 4, modUnicode.U("L\u01b0\u1edbi th\u00e9p n\u1ec1n"), "LT1", "D8-A200 (200\u00d7200mm) - L\u01b0\u1edbi h\u00e0n \u00f4 vu\u00f4ng D8 A200", 24, 1, 1

    ' Pre-fill formula rows down to row 100 for instant manual typing
    Dim r As Long
    For r = 10 To 100
        ApplySteelShapeRowFormulas ws, r
    Next r

    Dim widths As Variant
    widths = Array(6, 20, 9, 34, 7, 13, 8, 8, 12, 13, 13, 13, 15, 22)
    For i = LBound(widths) To UBound(widths)
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i
End Sub

Private Sub SeedShapeRow(ByVal ws As Worksheet, ByVal r As Long, ByVal stt As Long, _
                         ByVal comp As String, ByVal mark As String, ByVal specEscaped As String, _
                         ByVal dimVal As Double, ByVal q1 As Double, ByVal q2 As Double)
    ws.Cells(r, 1).Value = stt
    ws.Cells(r, 2).Value = comp
    ws.Cells(r, 3).Value = mark
    ws.Cells(r, 4).Value = modUnicode.U(specEscaped)
    ws.Cells(r, 6).Value = dimVal
    ws.Cells(r, 7).Value = q1
    ws.Cells(r, 8).Value = q2
    ApplySteelShapeRowFormulas ws, r
End Sub

Private Sub ApplySteelShapeRowFormulas(ByVal ws As Worksheet, ByVal r As Long)
    ws.Cells(r, 1).HorizontalAlignment = xlCenter
    ws.Cells(r, 5).Formula = "=IF(D" & r & "="""","""",IFERROR(INDEX(DB_VatLieu!$D:$D,MATCH(D" & r & ",DB_VatLieu!$C:$C,0)),""m""))"
    ws.Cells(r, 5).HorizontalAlignment = xlCenter
    ws.Cells(r, 9).Formula = "=IF(D" & r & "="""","""",ROUND(F" & r & "*G" & r & "*H" & r & ",2))"
    ws.Cells(r, 9).NumberFormat = "#,##0.00"
    ws.Cells(r, 10).Formula = "=IF(D" & r & "="""","""",IFERROR(INDEX(DB_VatLieu!$F:$F,MATCH(D" & r & ",DB_VatLieu!$C:$C,0)),0))"
    ws.Cells(r, 10).NumberFormat = "0.000"
    ws.Cells(r, 11).Formula = "=IF(D" & r & "="""","""",ROUND(I" & r & "*J" & r & ",2))"
    ws.Cells(r, 11).NumberFormat = "#,##0.00"
    ws.Cells(r, 11).Font.Bold = True
    ws.Cells(r, 12).Formula = "=IF(D" & r & "="""","""",IFERROR(INDEX(DB_VatLieu!$H:$H,MATCH(D" & r & ",DB_VatLieu!$C:$C,0)),0))"
    ws.Cells(r, 12).NumberFormat = "#,##0"
    ws.Cells(r, 13).Formula = "=IF(D" & r & "="""","""",ROUND(K" & r & "*L" & r & ",0))"
    ws.Cells(r, 13).NumberFormat = "#,##0"
    ws.Cells(r, 14).Formula = "=IF(D" & r & "="""","""",IFERROR(INDEX(DB_VatLieu!$A:$A,MATCH(D" & r & ",DB_VatLieu!$C:$C,0)),""""))"

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 14)).Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(200, 200, 200)
    End With
End Sub

Public Sub CreateSteelPlateSheet(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("TK_ThepTam")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "TK_ThepTam"
        BuildSteelPlateLayout ws
    End If

    ws.Activate
    Application.ScreenUpdating = True
    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 t\u1ea1o b\u1ea3ng Th\u1ed1ng k\u00ea Th\u00e9p T\u1ea5m / B\u1ea3n M\u00e3 \u0111a h\u00ecnh d\u1ea1ng (c\u00f3 tr\u1eeb l\u1ed7 bu-l\u00f4ng): TK_ThepTam."), vbInformation
    End If
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    modUnicode.MsgBoxU "CreateSteelPlateSheet Error: " & Err.Description, vbCritical
End Sub

Private Sub BuildSteelPlateLayout(ByVal ws As Worksheet)
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 9.5

    With ws.Range("A1:P1")
        .Interior.Color = RGB(22, 54, 92)
        .RowHeight = 28
    End With
    With ws.Range("B1:P1")
        .Merge
        .Value = modUnicode.U("B\u1ea2NG TH\u1ed0NG K\u00ca TH\u00c9P T\u1ea4M & B\u1ea2N M\u00c3 \u0110A H\u00ccNH D\u1ea0NG (7850 kg/m3)")
        .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
    End With

    ws.Range("L2").Value = modUnicode.U("T\u1ed4NG KL (kg):")
    ws.Range("L2").Font.Bold = True
    ws.Range("M2").Formula = "=SUM(P6:P1000)"
    ws.Range("M2").NumberFormat = "#,##0.00": ws.Range("M2").Font.Bold = True
    ws.Range("M2").Interior.Color = RGB(255, 242, 204)

    ws.Range("N2").Value = modUnicode.U("T\u1ea5n:")
    ws.Range("O2").Formula = "=M2/1000": ws.Range("O2").NumberFormat = "#,##0.000": ws.Range("O2").Font.Bold = True

    Dim hdrs As Variant
    hdrs = Array( _
        "STT", _
        modUnicode.U("C\u1ea5u ki\u1ec7n"), _
        modUnicode.U("S\u1ed1 hi\u1ec7u"), _
        modUnicode.U("H\u00ecnh d\u1ea1ng b\u1ea3n m\u00e3"), _
        "L1 (mm)", "L2 (mm)", "L3 (mm)", "L4 (mm)", _
        modUnicode.U("D\u00e0y t (mm)"), _
        modUnicode.U("S\u1ed1 l\u1ed7 BL"), _
        modUnicode.U("\u00d8 l\u1ed7 (mm)"), _
        modUnicode.U("SL/CK"), _
        modUnicode.U("S\u1ed1 CK"), _
        modUnicode.U("DT t\u1ecbnh 1 b\u1ea3n (m2)"), _
        modUnicode.U("Th\u1ec3 t\u00edch (m3)"), _
        modUnicode.U("T\u1ed5ng KL (kg)") _
    )
    Dim i As Long
    For i = LBound(hdrs) To UBound(hdrs)
        ws.Cells(5, i + 1).Value = hdrs(i)
    Next i
    With ws.Range("A5:P5")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 85, 151)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True: .RowHeight = 28
    End With

    Dim shapeList As String
    shapeList = "1-Chu nhat (L1xL2),2-Hinh thang ((L1+L2)/2*L3),3-Tam giac (0.5*L1*L2),4-Vai cot vat (L1*L2-0.5*(L1-L3)*(L2-L4)),5-Hinh tron (PI*L1^2/4)"
    With ws.Range("D6:D100").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=shapeList
        .InCellDropdown = True
    End With

    ' Seed 2 sample plates
    ws.Cells(6, 1).Value = 1: ws.Cells(6, 2).Value = modUnicode.U("B\u1ea3n m\u00e3 ch\u00e2n c\u1ed9t"): ws.Cells(6, 3).Value = "BM1"
    ws.Cells(6, 4).Value = "1-Chu nhat (L1xL2)": ws.Cells(6, 5).Value = 400: ws.Cells(6, 6).Value = 400: ws.Cells(6, 9).Value = 16: ws.Cells(6, 10).Value = 4: ws.Cells(6, 11).Value = 24: ws.Cells(6, 12).Value = 1: ws.Cells(6, 13).Value = 8

    ws.Cells(7, 1).Value = 2: ws.Cells(7, 2).Value = modUnicode.U("Vai c\u1ed9t v\u00e1t"): ws.Cells(7, 3).Value = "VC1"
    ws.Cells(7, 4).Value = "4-Vai cot vat (L1*L2-0.5*(L1-L3)*(L2-L4))": ws.Cells(7, 5).Value = 1000: ws.Cells(7, 6).Value = 800: ws.Cells(7, 7).Value = 300: ws.Cells(7, 8).Value = 300: ws.Cells(7, 9).Value = 12: ws.Cells(7, 10).Value = 0: ws.Cells(7, 11).Value = 0: ws.Cells(7, 12).Value = 2: ws.Cells(7, 13).Value = 4

    Dim r As Long
    For r = 6 To 100
        ws.Cells(r, 14).Formula = "=IF(D" & r & "="""","""",ROUND((IF(LEFT(D" & r & ",1)=""1"",E" & r & "*F" & r & "," & _
            "IF(LEFT(D" & r & ",1)=""2"",(E" & r & "+F" & r & ")/2*G" & r & "," & _
            "IF(LEFT(D" & r & ",1)=""3"",0.5*E" & r & "*F" & r & "," & _
            "IF(LEFT(D" & r & ",1)=""4"",E" & r & "*F" & r & "-0.5*(E" & r & "-G" & r & ")*(F" & r & "-H" & r & ")," & _
            "IF(LEFT(D" & r & ",1)=""5"",PI()*(E" & r & "^2)/4,0))))) - N(J" & r & ")*PI()*(N(K" & r & ")^2)/4)*10^-6,4))"
        ws.Cells(r, 14).NumberFormat = "0.0000"
        ws.Cells(r, 15).Formula = "=IF(D" & r & "="""","""",ROUND(N" & r & "*(I" & r & "/1000)*L" & r & "*M" & r & ",4))"
        ws.Cells(r, 15).NumberFormat = "0.0000"
        ws.Cells(r, 16).Formula = "=IF(D" & r & "="""","""",ROUND(O" & r & "*7850,2))"
        ws.Cells(r, 16).NumberFormat = "#,##0.00": ws.Cells(r, 16).Font.Bold = True
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, 16)).Borders
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(200, 200, 200)
        End With
    Next r

    Dim widths As Variant
    widths = Array(6, 20, 8, 32, 9, 9, 9, 9, 10, 8, 9, 7, 7, 14, 12, 13)
    For i = LBound(widths) To UBound(widths)
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i
End Sub
