Attribute VB_Name = "modCatalog"
Option Explicit

' ==============================================================================
' modCatalog - Material Catalog Lookup Engine (Rebar, Steel Shapes, Inox, Mesh)
' ==============================================================================

Private mCatalogData As Variant
Private mCatalogCount As Long

Public Sub ShowSteelLookupForm()
    EnsureCatalogLoaded
    frmSteelLookup.Show vbModeless
End Sub

Public Sub EnsureCatalogLoaded()
    If mCatalogCount > 0 Then Exit Sub

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("DB_Materials")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    mCatalogData = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 9)).Value
    mCatalogCount = UBound(mCatalogData, 1)
End Sub

Public Function GetCatalogCategories() As Variant
    EnsureCatalogLoaded
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim i As Long, cat As String
    For i = 1 To mCatalogCount
        cat = CStr(mCatalogData(i, 1))
        If Len(cat) > 0 And Not dict.Exists(cat) Then
            dict.Add cat, True
        End If
    Next i
    GetCatalogCategories = dict.Keys
End Function

' Filter materials in memory; returns 2D array (1..matchCount, 1..9) or Empty
Public Function SearchCatalog(ByVal categoryFilter As String, ByVal keyword As String, Optional ByVal maxRows As Long = 350) As Variant
    EnsureCatalogLoaded
    If mCatalogCount = 0 Then Exit Function

    Dim allCats As String
    allCats = modUnicode.T("FRM_ALL_CATS")

    Dim tokens() As String
    Dim kw As String
    kw = Trim$(keyword)
    Dim hasKw As Boolean
    If Len(kw) > 0 Then
        tokens = Split(kw, " ")
        hasKw = True
    End If

    Dim idxList() As Long
    ReDim idxList(1 To maxRows)
    Dim cnt As Long
    cnt = 0

    Dim i As Long, k As Long, matched As Boolean
    Dim rowText As String, catVal As String
    Dim useCat As Boolean
    useCat = (Len(categoryFilter) > 0 And categoryFilter <> allCats)

    For i = 1 To mCatalogCount
        catVal = CStr(mCatalogData(i, 1))
        If useCat Then
            If StrComp(catVal, categoryFilter, vbTextCompare) <> 0 Then GoTo NextRow
        End If

        If hasKw Then
            rowText = catVal & " " & CStr(mCatalogData(i, 2)) & " " & CStr(mCatalogData(i, 3))
            matched = True
            For k = LBound(tokens) To UBound(tokens)
                If Len(tokens(k)) > 0 Then
                    If InStr(1, rowText, tokens(k), vbTextCompare) = 0 Then
                        matched = False
                        Exit For
                    End If
                End If
            Next k
            If Not matched Then GoTo NextRow
        End If

        cnt = cnt + 1
        idxList(cnt) = i
        If cnt >= maxRows Then Exit For
NextRow:
    Next i

    If cnt = 0 Then Exit Function

    Dim res() As Variant
    ReDim res(1 To cnt, 1 To 9)
    Dim c As Long
    For i = 1 To cnt
        For c = 1 To 9
            res(i, c) = mCatalogData(idxList(i), c)
        Next c
    Next i
    SearchCatalog = res
End Function

Public Sub InsertMaterialToActiveRow(ByVal cat As String, ByVal subCat As String, _
                                     ByVal spec As String, ByVal unitStr As String, _
                                     ByVal stdLen As Variant, ByVal wtPerUnit As Double, _
                                     ByVal calcQty As Double, ByVal calcLenOrArea As Double, _
                                     ByVal totalWeight As Double)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    If ActiveCell Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim r As Long, c As Long
    r = ActiveCell.Row
    c = ActiveCell.Column

    ' Smart column mapping: if user is in a table with >= 2 columns, write structured columns starting at ActiveCell
    ws.Cells(r, c).Value = spec
    ws.Cells(r, c + 1).Value = unitStr
    ws.Cells(r, c + 2).Value = calcLenOrArea
    ws.Cells(r, c + 3).Value = calcQty
    ws.Cells(r, c + 4).Value = wtPerUnit
    ws.Cells(r, c + 5).FormulaR1C1 = "=ROUND(RC[-3]*RC[-2]*RC[-1],2)"
    ws.Cells(r, c + 6).Value = cat & " - " & subCat
End Sub

Public Sub ExportCatalogToActiveBook(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    ' 1. Export DB_VatLieu (1,559 items)
    Dim srcMat As Worksheet, dstMat As Worksheet
    Set srcMat = ThisWorkbook.Worksheets("DB_Materials")
    Set dstMat = EnsureOrResetSheet(wb, "DB_VatLieu")

    Dim lastRow As Long
    lastRow = srcMat.Cells(srcMat.Rows.Count, 1).End(xlUp).Row
    dstMat.Range("A1:I" & lastRow).Value = srcMat.Range("A1:I" & lastRow).Value
    RefreshCatalogHeaders wb
    FormatCatalogSheet dstMat, lastRow, 9
    modSettings.RegisterWorkbookName wb, "TBL_VATLIEU", dstMat.Range("A2:I" & lastRow)

    ' 2. Export DB_ThepThanh (15 rebar diameters)
    Dim srcReb As Worksheet, dstReb As Worksheet
    Set srcReb = ThisWorkbook.Worksheets("DB_Rebar")
    Set dstReb = EnsureOrResetSheet(wb, "DB_ThepThanh")

    Dim lastReb As Long
    lastReb = srcReb.Cells(srcReb.Rows.Count, 1).End(xlUp).Row
    dstReb.Range("A1:G" & lastReb).Value = srcReb.Range("A1:G" & lastReb).Value
    FormatCatalogSheet dstReb, lastReb, 7
    modSettings.RegisterWorkbookName wb, "TBL_THEPTHANH", dstReb.Range("A2:G" & lastReb)

    Application.ScreenUpdating = True
    If showMsg Then
        dstMat.Activate
        modUnicode.MsgBoxU modUnicode.T("MSG_CATALOG_EXPORTED", "VI"), vbInformation
    End If
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    modUnicode.MsgBoxU "ExportCatalog Error: " & Err.Description, vbCritical
End Sub

Public Sub RefreshCatalogHeaders(ByVal wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("DB_VatLieu")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ws.Cells(1, 1).Value = modUnicode.T("CAT_HDR_CAT")
    ws.Cells(1, 2).Value = modUnicode.T("CAT_HDR_SUB")
    ws.Cells(1, 3).Value = modUnicode.T("CAT_HDR_SPEC")
    ws.Cells(1, 4).Value = modUnicode.T("CAT_HDR_UNIT")
    ws.Cells(1, 5).Value = modUnicode.T("CAT_HDR_LEN")
    ws.Cells(1, 6).Value = modUnicode.T("CAT_HDR_WT")
    ws.Cells(1, 7).Value = modUnicode.T("CAT_HDR_WTBAR")
    ws.Cells(1, 8).Value = modUnicode.T("CAT_HDR_PRICE")
    ws.Cells(1, 9).Value = modUnicode.T("CAT_HDR_SRC")
End Sub

Private Function EnsureOrResetSheet(ByVal wb As Workbook, ByVal sName As String) As Worksheet
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
    Set EnsureOrResetSheet = ws
End Function

Private Sub FormatCatalogSheet(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal colCount As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 28
    End With
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, colCount))
        .Font.Name = "Segoe UI"
        .Font.Size = 9.5
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(217, 217, 217)
    End With
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, colCount)).AutoFilter
    ws.Columns.AutoFit
    ws.Activate
    On Error Resume Next
    If Not Application.ActiveWindow Is Nothing Then
        Application.ActiveWindow.SplitRow = 1
        Application.ActiveWindow.FreezePanes = True
    End If
    On Error GoTo 0
End Sub
