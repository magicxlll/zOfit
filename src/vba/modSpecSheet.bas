Attribute VB_Name = "modSpecSheet"
Option Explicit

' ==============================================================================
' modSpecSheet - Standardized Project Spec Input Sheet (QS_Spec)
' Supports per-element-type Lean Concrete thickness/overhang + Rebar stagger
' ==============================================================================

Public Const SPEC_SHEET As String = "QS_Spec"

Public Sub CreateOrShowSpecSheet(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub

    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet
    Dim isNew As Boolean

    On Error Resume Next
    Set ws = wb.Worksheets(SPEC_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        ws.Name = SPEC_SHEET
        isNew = True
    End If

    If isNew Then BuildSpecLayout ws

    ws.Activate
    If showMsg Then
        modUnicode.MsgBoxU modUnicode.U("\u0110\u00e3 t\u1ea1o/m\u1edf b\u1ea3ng Ti\u00eau chu\u1ea9n K\u1ef9 thu\u1eadt D\u1ef1 \u00e1n (QS_Spec) v\u1edbi \u0111\u1ea7y \u0111\u1ee7 Named Range SPEC_*."), vbInformation
    End If
End Sub

Private Sub BuildSpecLayout(ByVal ws As Worksheet)
    Application.ScreenUpdating = False
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 9.5

    ws.Range("A1").Value = "ZOFIT_SPEC_V2"
    ws.Range("A1").Font.Color = RGB(22, 54, 92)

    With ws.Range("B2:F2")
        .Merge
        .Value = modUnicode.U("TI\u00caU CHU\u1ea8N K\u1ef8 THU\u1eacT D\u1ef0 \u00c1N & QUY C\u00c1CH \u0110O B\u00d3C (QS SPEC)")
        .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(22, 54, 92): .RowHeight = 30: .IndentLevel = 1: .VerticalAlignment = xlCenter
    End With

    With ws.Range("B3:F3")
        .Merge
        .Value = modUnicode.U("T\u00f9y ch\u1ec9nh chi\u1ec1u d\u00e0y/m\u1edf r\u1ed9ng b\u00ea t\u00f4ng l\u00f3t theo t\u1eebng lo\u1ea1i c\u1ea5u ki\u1ec7n, l\u1edbp b\u1ea3o v\u1ec7, neo/n\u1ed1i so le v\u00e0 Coupler.")
        .Font.Italic = True: .Font.Size = 9: .Font.Color = RGB(64, 64, 64)
        .Interior.Color = RGB(237, 242, 248): .RowHeight = 20: .IndentLevel = 1
    End With

    Dim r As Long: r = 5
    ws.Cells(r, 2).Value = modUnicode.U("Th\u00f4ng s\u1ed1 k\u1ef9 thu\u1eadt")
    ws.Cells(r, 3).Value = "Named Range"
    ws.Cells(r, 4).Value = modUnicode.U("Gi\u00e1 tr\u1ecb \u00e1p d\u1ee5ng")
    ws.Cells(r, 5).Value = modUnicode.U("\u0110VT / L\u1ef1a ch\u1ecdn")
    ws.Cells(r, 6).Value = modUnicode.U("Di\u1ec5n gi\u1ea3i theo Spec")
    With ws.Range("B5:F5")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(47, 85, 151): .HorizontalAlignment = xlCenter: .RowHeight = 24
    End With

    r = 6
    r = AddBanner(ws, r, modUnicode.U("A. CAO \u0110\u1ed8 & C\u00d4NG T\u00c1C \u0110\u1ea4T (EARTHWORK & ELEVATIONS)"))
    r = AddRow(ws, r, modUnicode.U("Cao \u0111\u1ed9 \u0111\u1ea5t t\u1ef1 nhi\u00ean hi\u1ec7n tr\u1ea1ng (m)"), "SPEC_Elev_Natural", 19.7, "m", modUnicode.U("C\u1ed1t m\u1eb7t \u0111\u1ea5t t\u1ef1 nhi\u00ean kh\u1ea3o s\u00e1t (VD: +19.70m)"))
    r = AddRow(ws, r, modUnicode.U("Cao \u0111\u1ed9 n\u1ec1n ho\u00e0n thi\u1ec7n FFL (m)"), "SPEC_Elev_FFL", 20#, "m", modUnicode.U("C\u1ed1t +/-0.000 ho\u1eb7c c\u1ed1t tuy\u1ec7t \u0111\u1ed1i n\u1ec1n t\u1ea7ng 1"))
    r = AddRow(ws, r, modUnicode.U("Chi\u1ec1u d\u00e0y b\u00f3c \u0111\u1ea5t h\u1eefu c\u01a1 (m)"), "SPEC_Topsoil_Thk", 0.1, "m", modUnicode.U("B\u00f3c l\u1edbp \u0111\u1ea5t kh\u00f4ng ph\u00f9 h\u1ee3p"))
    r = AddRow(ws, r, modUnicode.U("M\u1edf r\u1ed9ng \u0111\u00e0o m\u00f3ng m\u1ed7i b\u00ean (m)"), "SPEC_Exc_WorkSpace_Ftg", 0.7, "m", modUnicode.U("Kho\u1ea3ng thao t\u00e1c v\u00e1n khu\u00f4n m\u00f3ng / \u0111\u00e0i c\u1ecdc"))
    r = AddRow(ws, r, modUnicode.U("M\u1edf r\u1ed9ng \u0111\u00e0o gi\u1eb1ng m\u00f3ng m\u1ed7i b\u00ean (m)"), "SPEC_Exc_WorkSpace_Beam", 0.5, "m", modUnicode.U("Kho\u1ea3ng thao t\u00e1c r\u00e3nh \u0111\u00e0o gi\u1eb1ng m\u00f3ng / \u0111\u00e0 ki\u1ec1ng"))
    r = AddRow(ws, r, modUnicode.U("H\u1ec7 s\u1ed1 m\u1edf taluy h\u1ed1 \u0111\u00e0o (K_taluy)"), "SPEC_Exc_SlopeK", 1.3, "Factor", modUnicode.U("H\u1ec7 s\u1ed1 quy \u0111\u1ed5i m\u00e1i d\u1ed1c h\u1ed1 \u0111\u00e0o"))
    r = AddRow(ws, r, modUnicode.U("T\u1ef7 l\u1ec7 \u0111\u00e0o \u0111\u1ea5t b\u1eb1ng m\u00e1y"), "SPEC_Exc_MachinePct", 0.9, "Ratio", modUnicode.U("0.90 = 90%"))
    r = AddRow(ws, r, modUnicode.U("T\u1ef7 l\u1ec7 \u0111\u00e0o \u0111\u1ea5t th\u1ee7 c\u00f4ng"), "SPEC_Exc_LaborPct", 0.1, "Ratio", modUnicode.U("0.10 = 10%"))
    r = AddRow(ws, r, modUnicode.U("H\u1ec7 s\u1ed1 \u0111\u1ea7m n\u00e9n l\u1ea5p K95"), "SPEC_Backfill_CompK", 1.3, "Factor", modUnicode.U("H\u1ec7 s\u1ed1 n\u00e9n ch\u1eb7t khi l\u1ea5p c\u00e1t / \u0111\u1ea5t"))

    r = AddBanner(ws, r, modUnicode.U("B. B\u00ca T\u00d4NG & B\u00ca T\u00d4NG L\u00d3T T\u00d9Y BI\u1ebeN THEO LO\u1ea0I C\u1ea4U KI\u1ec6N"))
    r = AddRow(ws, r, modUnicode.U("L\u00f3t M\u00f3ng / \u0110\u00e0i c\u1ecdc - Chi\u1ec1u d\u00e0y (m)"), "SPEC_Lean_Ftg_Thk", 0.1, "m", modUnicode.U("M\u1eb7c \u0111\u1ecbnh l\u00f3t m\u00f3ng (100mm = 0.10m)"))
    r = AddRow(ws, r, modUnicode.U("L\u00f3t M\u00f3ng / \u0110\u00e0i c\u1ecdc - M\u1edf r\u1ed9ng m\u1ed7i b\u00ean (m)"), "SPEC_Lean_Ftg_Ext", 0.1, "m", modUnicode.U("M\u1edf r\u1ed9ng m\u1ed7i m\u00e9p m\u00f3ng (100mm = 0.10m)"))
    r = AddRow(ws, r, modUnicode.U("L\u00f3t Gi\u1eb1ng m\u00f3ng / D\u1ea7m m\u00f3ng - Chi\u1ec1u d\u00e0y (m)"), "SPEC_Lean_Beam_Thk", 0.05, "m", modUnicode.U("M\u1eb7c \u0111\u1ecbnh l\u00f3t gi\u1eb1ng/d\u1ea7m m\u00f3ng (50mm = 0.05m)"))
    r = AddRow(ws, r, modUnicode.U("L\u00f3t Gi\u1eb1ng m\u00f3ng / D\u1ea7m m\u00f3ng - M\u1edf r\u1ed9ng (m)"), "SPEC_Lean_Beam_Ext", 0.1, "m", modUnicode.U("M\u1edf r\u1ed9ng m\u1ed7i m\u00e9p d\u1ea7m m\u00f3ng (100mm = 0.10m)"))
    r = AddRow(ws, r, modUnicode.U("L\u00f3t S\u00e0n tr\u1ec7t / Ram d\u1ed1c - Chi\u1ec1u d\u00e0y (m)"), "SPEC_Lean_Slab_Thk", 0.05, "m", modUnicode.U("M\u1eb7c \u0111\u1ecbnh l\u00f3t s\u00e0n n\u1ec1n (50mm = 0.05m)"))
    r = AddRow(ws, r, modUnicode.U("L\u00f3t S\u00e0n tr\u1ec7t / Ram d\u1ed1c - M\u1edf r\u1ed9ng (m)"), "SPEC_Lean_Slab_Ext", 0#, "m", modUnicode.U("M\u1edf r\u1ed9ng m\u00e9p l\u00f3t s\u00e0n (th\u01b0\u1eddng 0.00m)"))
    r = AddRow(ws, r, modUnicode.U("L\u00f3t B\u1ec3 ng\u1ea7m / H\u1ed1 ga - Chi\u1ec1u d\u00e0y (m)"), "SPEC_Lean_Tank_Thk", 0.1, "m", modUnicode.U("M\u1eb7c \u0111\u1ecbnh l\u00f3t \u0111\u00e1y b\u1ec3 ng\u1ea7m/h\u1ed1 ga (0.10m)"))
    r = AddRow(ws, r, modUnicode.U("L\u00f3t B\u1ec3 ng\u1ea7m / H\u1ed1 ga - M\u1edf r\u1ed9ng m\u1ed7i b\u00ean (m)"), "SPEC_Lean_Tank_Ext", 0.15, "m", modUnicode.U("M\u1edf r\u1ed9ng \u0111\u00e1y l\u00f3t b\u1ec3 ng\u1ea7m (0.15 - 0.20m)"))
    r = AddRow(ws, r, modUnicode.U("H\u1ec7 s\u1ed1 hao h\u1ee5t / b\u00f9 b\u00ea t\u00f4ng"), "SPEC_Conc_Waste", 1.03, "Factor", modUnicode.U("1.03 = +3% theo form BHFLEX; 1.00 = l\u00fd thuy\u1ebft"))
    r = AddRow(ws, r, modUnicode.U("H\u1ec7 s\u1ed1 ch\u1ed3ng m\u00e9p m\u00e0ng PE l\u00f3t s\u00e0n"), "SPEC_PE_Overlap", 1.01, "Factor", modUnicode.U("1.01 = +1% ch\u1ed3ng m\u00e9p"))
    r = AddRow(ws, r, modUnicode.U("Ng\u01b0\u1ee1ng l\u1ed7 m\u1edf kh\u00f4ng tr\u1eeb VK/BT (m2)"), "SPEC_Opening_DeductMin", 0.5, "m2", modUnicode.U("TT13=0.5m2 | POMI=1.0m2"))

    ' Legacy aliases pointing to Footing defaults
    modSettings.RegisterWorkbookName ws.Parent, "SPEC_Lean_Thk", ws.Parent.Names("SPEC_Lean_Ftg_Thk").RefersToRange
    modSettings.RegisterWorkbookName ws.Parent, "SPEC_Lean_Overhang", ws.Parent.Names("SPEC_Lean_Ftg_Ext").RefersToRange

    r = AddBanner(ws, r, modUnicode.U("C. QUY CHU\u1ea8N C\u1ed0T TH\u00c9P (B\u1ea2O V\u1ec6, NEO, N\u1ed0I SO LE, COUPLER)"))
    r = AddRow(ws, r, modUnicode.U("B\u00ea t\u00f4ng b\u1ea3o v\u1ec7 M\u00f3ng / \u0110\u00e0i c\u1ecdc (mm)"), "SPEC_Cover_Ftg", 50, "mm", "c_nom = MAX(c_base, d)")
    r = AddRow(ws, r, modUnicode.U("B\u00ea t\u00f4ng b\u1ea3o v\u1ec7 C\u1ed9t (mm)"), "SPEC_Cover_Col", 35, "mm", modUnicode.U("C\u1ed9t khung ch\u1ecbu l\u1ef1c"))
    r = AddRow(ws, r, modUnicode.U("B\u00ea t\u00f4ng b\u1ea3o v\u1ec7 D\u1ea7m (mm)"), "SPEC_Cover_Beam", 30, "mm", modUnicode.U("D\u1ea7m ch\u00ednh / D\u1ea7m ph\u1ee5"))
    r = AddRow(ws, r, modUnicode.U("B\u00ea t\u00f4ng b\u1ea3o v\u1ec7 S\u00e0n / C\u1ea7u thang (mm)"), "SPEC_Cover_Slab", 20, "mm", modUnicode.U("B\u1ea3n s\u00e0n / V\u1ebf thang"))
    r = AddRow(ws, r, modUnicode.U("B\u00ea t\u00f4ng b\u1ea3o v\u1ec7 V\u00e1ch / B\u1ec3 ng\u1ea7m (mm)"), "SPEC_Cover_Wall", 30, "mm", modUnicode.U("V\u00e1ch ch\u1ecbu l\u1ef1c / Th\u00e0nh b\u1ec3"))
    r = AddRow(ws, r, modUnicode.U("H\u1ec7 s\u1ed1 chi\u1ec1u d\u00e0i neo k\u00e9o Lan (xD)"), "SPEC_Lan_Tension", 40, "xD", modUnicode.U("V\u00f9ng k\u00e9o = 40D"))
    r = AddRow(ws, r, modUnicode.U("H\u1ec7 s\u1ed1 chi\u1ec1u d\u00e0i neo n\u00e9n (xD)"), "SPEC_Lan_Comp", 30, "xD", modUnicode.U("V\u00f9ng n\u00e9n = 30D"))
    r = AddRow(ws, r, modUnicode.U("H\u1ec7 s\u1ed1 n\u1ed1i ch\u1ed3ng L_lap (xD)"), "SPEC_Llap_Mult", 45, "xD", modUnicode.U("N\u1ed1i ch\u1ed3ng = 45D"))
    r = AddRow(ws, r, modUnicode.U("Kho\u1ea3ng l\u1ec7ch n\u1ed1i so le c\u1ed9t (mm)"), "SPEC_Col_StaggerOffset", 600, "mm", modUnicode.U("\u0110o\u1ea1n l\u1ec7ch nh\u00f3m n\u1ed1i so le c\u1ed9t (>= 600mm)"))
    r = AddRow(ws, r, modUnicode.U("\u0110\u01b0\u1eddng k\u00ednh b\u1eaft \u0111\u1ea7u d\u00f9ng Coupler"), "SPEC_CouplerMinD", 20, "20,22,25,28,32,99", modUnicode.U(">= D n\u00e0y t\u1ef1 \u0111\u1ed5i sang Coupler"))
    r = AddRow(ws, r, modUnicode.U("Chi\u1ec1u d\u00e0i c\u00e2y th\u00e9p chu\u1ea9n (mm)"), "SPEC_StdBarLen", 11700, "11700,12000", modUnicode.U("Chi\u1ec1u d\u00e0i 1 c\u00e2y th\u00e9p th\u01b0\u01a1ng m\u1ea1i"))
    r = AddRow(ws, r, modUnicode.U("Gi\u1ea3m u\u1ed1n m\u1ed7i g\u00f3c 90 \u0111\u1ed9 (xD)"), "SPEC_BendDeduct90", 1, "0,1,2", modUnicode.U("Tr\u1eeb d\u00e3n ph\u00f4i khi u\u1ed1n 90 \u0111\u1ed9"))
    r = AddRow(ws, r, modUnicode.U("Hao h\u1ee5t th\u00e9p <= D10 (%)"), "SPEC_Scrap_Small", 3#, "%", "<= 10mm = 3.0%")
    r = AddRow(ws, r, modUnicode.U("Hao h\u1ee5t th\u00e9p D12 - D18 (%)"), "SPEC_Scrap_Med", 2.5, "%", "12-18mm = 2.5%")
    r = AddRow(ws, r, modUnicode.U("Hao h\u1ee5t th\u00e9p > D18 (%)"), "SPEC_Scrap_Large", 2#, "%", "> 18mm = 2.0%")

    r = AddBanner(ws, r, modUnicode.U("D. C\u1eccC PHC & N\u1ec0N X\u01af\u1edeNG C\u00d4NG NGHI\u1ec6P"))
    r = AddRow(ws, r, modUnicode.U("Chi\u1ec1u d\u00e0i \u0111\u1eadp / c\u1eaft \u0111\u1ea7u c\u1ecdc (m)"), "SPEC_Pile_CutHead", 0.5, "m", modUnicode.U("C\u1eaft \u0111\u1ea7u c\u1ecdc m\u1eb7c \u0111\u1ecbnh"))
    r = AddRow(ws, r, modUnicode.U("Chi\u1ec1u d\u00e0i 1 \u0111o\u1ea1n c\u1ecdc PHC (m)"), "SPEC_Pile_SegLen", 11#, "m", modUnicode.U("T\u00ednh s\u1ed1 m\u1ed1i n\u1ed1i c\u1ecdc"))
    r = AddRow(ws, r, modUnicode.U("B\u01b0\u1edbc l\u01b0\u1edbi c\u1eaft khe co gi\u00e3n s\u00e0n (m)"), "SPEC_Joint_Grid", 6#, "m", modUnicode.U("\u00d4 c\u1edd c\u1eaft khe (6.0m x 6.0m)"))

    r = AddBanner(ws, r, modUnicode.U("E. HO\u00c0N THI\u1ec6N & H\u1ea0 T\u1ea6NG"))
    r = AddRow(ws, r, modUnicode.U("Ng\u01b0\u1ee1ng l\u1ed7 c\u1eeda kh\u00f4ng tr\u1eeb x\u00e2y/tr\u00e1t (m2)"), "SPEC_Wall_OpeningMin", 0.25, "m2", "TT13 > 0.25m2")
    r = AddRow(ws, r, modUnicode.U("B\u1ec1 r\u1ed9ng l\u01b0\u1edbi m\u1eaft c\u00e1o ch\u1ed1ng n\u1ee9t (m)"), "SPEC_Mesh_Width", 0.2, "m", "TCVN 9377-2 >= 0.20m")
    r = AddRow(ws, r, modUnicode.U("Chi\u1ec1u cao v\u00e9n ch\u00e2n t\u01b0\u1eddng WC (m)"), "SPEC_WP_CurbH", 0.3, "m", modUnicode.U("V\u00e9n ch\u00e2n t\u01b0\u1eddng ch\u1ed1ng th\u1ea5m 300mm"))
    r = AddRow(ws, r, modUnicode.U("Chi\u1ec1u d\u00e0y \u0111\u1ec7m c\u00e1t/\u0111\u00e1 \u0111\u00e1y c\u1ed1ng (m)"), "SPEC_Pipe_Bedding", 0.15, "m", "TCVN 7957 / 9113")
    r = AddRow(ws, r, modUnicode.U("M\u00e1i d\u1ed1c taluy \u0111\u00e0o r\u00e3nh c\u1ed1ng"), "SPEC_Pipe_Slope", 0.3, "Factor", "m = 0.30..0.50")
    r = AddRow(ws, r, modUnicode.U("Chi\u1ec1u d\u00e0i thanh Dowel (m)"), "SPEC_Road_DowelLen", 0.5, "m", "D20 L=0.50m")
    r = AddRow(ws, r, modUnicode.U("B\u01b0\u1edbc r\u1ea3i thanh Dowel (m)"), "SPEC_Road_DowelSpace", 0.3, "m", "@300mm")

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 44
    ws.Columns("C").ColumnWidth = 26
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 48

    With ws.Range("B5:F" & (r - 1)).Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(191, 191, 191)
    End With
    ws.Range("B5:F" & (r - 1)).VerticalAlignment = xlCenter
    On Error Resume Next
    If Not Application.ActiveWindow Is Nothing Then Application.ActiveWindow.DisplayGridlines = True
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

Private Function AddBanner(ByVal ws As Worksheet, ByVal rowIdx As Long, ByVal titleText As String) As Long
    With ws.Range(ws.Cells(rowIdx, 2), ws.Cells(rowIdx, 6))
        .Merge: .Value = titleText: .Font.Bold = True
        .Font.Color = RGB(31, 78, 120): .Interior.Color = RGB(217, 225, 242): .IndentLevel = 1: .RowHeight = 22
    End With
    AddBanner = rowIdx + 1
End Function

Private Function AddRow(ByVal ws As Worksheet, ByVal rowIdx As Long, _
                        ByVal label As String, ByVal nm As String, _
                        ByVal val As Variant, ByVal unitOrList As String, ByVal noteStr As String) As Long
    ws.Cells(rowIdx, 2).Value = label
    ws.Cells(rowIdx, 3).Value = nm
    ws.Cells(rowIdx, 3).Font.Name = "Consolas": ws.Cells(rowIdx, 3).Font.Size = 9: ws.Cells(rowIdx, 3).Font.Color = RGB(89, 89, 89)

    Dim cVal As Range: Set cVal = ws.Cells(rowIdx, 4)
    cVal.Value = val: cVal.Font.Bold = True: cVal.Font.Color = RGB(0, 51, 153)
    cVal.Interior.Color = RGB(255, 250, 220): cVal.HorizontalAlignment = xlCenter

    If InStr(1, unitOrList, ",") > 0 Then
        With cVal.Validation
            .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=unitOrList: .InCellDropdown = True
        End With
        ws.Cells(rowIdx, 5).Value = "(List)"
    Else
        ws.Cells(rowIdx, 5).Value = unitOrList
    End If
    ws.Cells(rowIdx, 5).HorizontalAlignment = xlCenter
    ws.Cells(rowIdx, 6).Value = noteStr: ws.Cells(rowIdx, 6).WrapText = True: ws.Rows(rowIdx).RowHeight = 21

    modSettings.RegisterWorkbookName ws.Parent, nm, cVal
    AddRow = rowIdx + 1
End Function

Public Function GetSpecNum(ByVal wb As Workbook, ByVal nameA As String, Optional ByVal fallbackCfg As String = "", Optional ByVal defVal As Double = 0#) As Double
    On Error Resume Next
    Dim r As Range: Set r = wb.Names(nameA).RefersToRange
    If Not r Is Nothing Then If IsNumeric(r.Value2) Then GetSpecNum = CDbl(r.Value2): Exit Function
    If Len(fallbackCfg) > 0 Then
        Set r = Nothing: Set r = wb.Names(fallbackCfg).RefersToRange
        If Not r Is Nothing Then If IsNumeric(r.Value2) Then GetSpecNum = CDbl(r.Value2): Exit Function
    End If
    On Error GoTo 0
    GetSpecNum = defVal
End Function

Public Function GetSpecStr(ByVal wb As Workbook, ByVal nameA As String, Optional ByVal defVal As String = "") As String
    On Error Resume Next
    Dim r As Range: Set r = wb.Names(nameA).RefersToRange
    If Not r Is Nothing Then GetSpecStr = CStr(r.Value): Exit Function
    On Error GoTo 0
    GetSpecStr = defVal
End Function

Public Function SerializeSpecToJson(ByVal wb As Workbook) As String
    Dim sb As String
    sb = "{"
    sb = sb & """projectName"":" & modEngineBridge.JsonStr(GetSpecStr(wb, "CFG_ProjectName", "zOfit Project")) & ","
    sb = sb & """langMode"":" & modEngineBridge.JsonStr(GetSpecStr(wb, "CFG_Lang", "VI")) & ","
    sb = sb & """elevNatural"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Elev_Natural", "", 19.7)) & ","
    sb = sb & """elevFfl"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Elev_FFL", "", 20#)) & ","
    sb = sb & """topsoilThkM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Topsoil_Thk", "", 0.1)) & ","
    sb = sb & """excWorkSpaceFtgM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Exc_WorkSpace_Ftg", "", 0.7)) & ","
    sb = sb & """excWorkSpaceBeamM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Exc_WorkSpace_Beam", "", 0.5)) & ","
    sb = sb & """excSlopeK"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Exc_SlopeK", "", 1.3)) & ","
    sb = sb & """excMachinePct"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Exc_MachinePct", "", 0.9)) & ","
    sb = sb & """excLaborPct"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Exc_LaborPct", "", 0.1)) & ","
    sb = sb & """backfillCompK"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Backfill_CompK", "", 1.3)) & ","
    sb = sb & """leanFtgThkM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Lean_Ftg_Thk", "", 0.1)) & ","
    sb = sb & """leanFtgExtM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Lean_Ftg_Ext", "", 0.1)) & ","
    sb = sb & """leanBeamThkM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Lean_Beam_Thk", "", 0.05)) & ","
    sb = sb & """leanBeamExtM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Lean_Beam_Ext", "", 0.1)) & ","
    sb = sb & """leanSlabThkM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Lean_Slab_Thk", "", 0.05)) & ","
    sb = sb & """leanSlabExtM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Lean_Slab_Ext", "", 0#)) & ","
    sb = sb & """leanTankThkM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Lean_Tank_Thk", "", 0.1)) & ","
    sb = sb & """leanTankExtM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Lean_Tank_Ext", "", 0.15)) & ","
    sb = sb & """concWasteFactor"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Conc_Waste", "", 1.03)) & ","
    sb = sb & """peOverlapFactor"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_PE_Overlap", "", 1.01)) & ","
    sb = sb & """openingDeductMinM2"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Opening_DeductMin", "CFG_OpenThreshold", 0.5)) & ","
    sb = sb & """coverFtgMm"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Cover_Ftg", "CFG_CoverFdn", 50)) & ","
    sb = sb & """coverColMm"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Cover_Col", "CFG_CoverColBeam", 35)) & ","
    sb = sb & """coverBeamMm"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Cover_Beam", "CFG_CoverColBeam", 30)) & ","
    sb = sb & """coverSlabMm"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Cover_Slab", "CFG_CoverSlabWall", 20)) & ","
    sb = sb & """coverWallMm"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Cover_Wall", "CFG_CoverSlabWall", 30)) & ","
    sb = sb & """lanTensionMult"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Lan_Tension", "", 40)) & ","
    sb = sb & """lanCompMult"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Lan_Comp", "", 30)) & ","
    sb = sb & """llapMult"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Llap_Mult", "CFG_LapRatio", 45)) & ","
    sb = sb & """colStaggerOffsetMm"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Col_StaggerOffset", "", 600)) & ","
    sb = sb & """couplerMinD"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_CouplerMinD", "CFG_CouplerMinD", 20)) & ","
    sb = sb & """stdBarLenMm"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_StdBarLen", "CFG_BarStdLen", 11700)) & ","
    sb = sb & """bendDeduct90Mult"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_BendDeduct90", "CFG_BendDeduct", 1)) & ","
    sb = sb & """scrapSmallPct"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Scrap_Small", "", 3#)) & ","
    sb = sb & """scrapMedPct"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Scrap_Med", "", 2.5)) & ","
    sb = sb & """scrapLargePct"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Scrap_Large", "", 2#)) & ","
    sb = sb & """pileCutHeadM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Pile_CutHead", "", 0.5)) & ","
    sb = sb & """pileSegmentLenM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Pile_SegLen", "", 11#)) & ","
    sb = sb & """jointGridSpacingM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Joint_Grid", "", 6#)) & ","
    sb = sb & """wallOpeningMinM2"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Wall_OpeningMin", "", 0.25)) & ","
    sb = sb & """crackMeshWidthM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Mesh_Width", "", 0.2)) & ","
    sb = sb & """waterproofCurbHM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_WP_CurbH", "", 0.3)) & ","
    sb = sb & """pipeBeddingM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Pipe_Bedding", "", 0.15)) & ","
    sb = sb & """pipeSlopeM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Pipe_Slope", "", 0.3)) & ","
    sb = sb & """roadDowelLenM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Road_DowelLen", "", 0.5)) & ","
    sb = sb & """roadDowelSpacingM"":" & modEngineBridge.JsonNum(GetSpecNum(wb, "SPEC_Road_DowelSpace", "", 0.3))
    sb = sb & "}"
    SerializeSpecToJson = sb
End Function
