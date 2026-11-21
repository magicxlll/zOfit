Attribute VB_Name = "modTools"
Option Explicit

' ==============================================================================
' modTools - QS Audit, Garbage Cleaner, Clean .xlsx Export & Master BOQ Stub
' ==============================================================================

Public Sub CleanWorkbookGarbage(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim totalCount As Long
    totalCount = wb.Names.Count
    If totalCount = 0 Then
        modUnicode.MsgBoxU modUnicode.T("MSG_CLEAN_NONE", "VI"), vbInformation
        Exit Sub
    End If

    Dim deletedCount As Long
    Dim i As Long
    Dim nm As Name
    Dim nmName As String, refStr As String
    Dim isTrash As Boolean

    Application.ScreenUpdating = False
    For i = totalCount To 1 Step -1
        Set nm = wb.Names(i)
        nmName = nm.Name
        refStr = ""
        On Error Resume Next
        refStr = nm.RefersTo
        On Error GoTo 0

        isTrash = False
        ' Keep built-in Excel print/filter names unless broken (#REF!)
        If InStr(1, refStr, "#REF!", vbTextCompare) > 0 Then
            isTrash = True
        ElseIf Not nm.Visible Then
            ' Hidden names that are not standard _xlnm built-ins or our CFG_/TBL_ ranges
            If Left$(nmName, 6) <> "_xlnm." And Left$(nmName, 4) <> "CFG_" And Left$(nmName, 4) <> "TBL_" Then
                If InStr(1, nmName, "\") > 0 Or InStr(1, nmName, "___") > 0 Or InStr(1, refStr, "#N/A") > 0 Or Len(nmName) > 40 Then
                    isTrash = True
                ElseIf Left$(refStr, 2) = "={" Then
                    isTrash = True
                End If
            End If
        End If

        If isTrash Then
            On Error Resume Next
            nm.Delete
            If Err.Number = 0 Then deletedCount = deletedCount + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next i
    Application.ScreenUpdating = True

    If showMsg Then
        If deletedCount = 0 Then
            modUnicode.MsgBoxU modUnicode.T("MSG_CLEAN_NONE", "VI"), vbInformation
        Else
            modUnicode.MsgBoxU modUnicode.T("MSG_CLEAN_DONE", "VI") & vbCrLf & _
                "(" & Format$(deletedCount, "#,##0") & " / " & Format$(totalCount, "#,##0") & " Defined Names)", vbInformation
        End If
    End If
End Sub

Public Sub AuditWorkbookFormulas(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim errList As Collection
    Set errList = New Collection

    Dim ws As Worksheet, cell As Range, fCells As Range
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible And ws.Name <> "QS_Audit_Report" Then
            Set fCells = Nothing
            On Error Resume Next
            Set fCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
            On Error GoTo 0
            If Not fCells Is Nothing Then
                For Each cell In fCells
                    errList.Add Array(ws.Name, cell.Address(False, False), cell.Text, CStr(cell.Formula))
                Next cell
            End If
        End If
    Next ws

    If errList.Count = 0 Then
        If showMsg Then modUnicode.MsgBoxU modUnicode.T("MSG_AUDIT_CLEAN", "VI"), vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Dim rpt As Worksheet
    On Error Resume Next
    Set rpt = wb.Worksheets("QS_Audit_Report")
    On Error GoTo 0
    If rpt Is Nothing Then
        Set rpt = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        rpt.Name = "QS_Audit_Report"
    Else
        rpt.Cells.Clear
    End If

    rpt.Range("A1:D1").Merge
    rpt.Range("A1").Value = modUnicode.T("AUDIT_TITLE")
    rpt.Range("A1").Font.Name = "Segoe UI"
    rpt.Range("A1").Font.Size = 13
    rpt.Range("A1").Font.Bold = True
    rpt.Range("A1").Font.Color = RGB(255, 255, 255)
    rpt.Range("A1:D1").Interior.Color = RGB(192, 0, 0)
    rpt.Rows(1).RowHeight = 28

    rpt.Cells(3, 1).Value = "Sheet"
    rpt.Cells(3, 2).Value = "Cell (Click)"
    rpt.Cells(3, 3).Value = "Error"
    rpt.Cells(3, 4).Value = "Formula"
    With rpt.Range("A3:D3")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
    End With

    Dim idx As Long, item As Variant
    For idx = 1 To errList.Count
        item = errList(idx)
        rpt.Cells(idx + 3, 1).Value = item(0)
        rpt.Cells(idx + 3, 2).Value = item(1)
        rpt.Hyperlinks.Add Anchor:=rpt.Cells(idx + 3, 2), Address:="", _
            SubAddress:="'" & item(0) & "'!" & item(1), TextToDisplay:=CStr(item(1))
        rpt.Cells(idx + 3, 3).Value = item(2)
        rpt.Cells(idx + 3, 4).Value = "'" & item(3)
    Next idx

    rpt.Columns("A:D").AutoFit
    rpt.Activate
    Application.ScreenUpdating = True
    If showMsg Then modUnicode.MsgBoxU modUnicode.T("MSG_AUDIT_FOUND", "VI") & " (" & errList.Count & " cells)", vbExclamation
End Sub

Public Sub ExportCleanXlsx()
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim defaultPath As String
    Dim baseName As String
    baseName = wb.Name
    If InStrRev(baseName, ".") > 0 Then
        baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    End If
    If Len(wb.Path) > 0 Then
        defaultPath = wb.Path & "\" & baseName & "_Clean.xlsx"
    Else
        defaultPath = baseName & "_Clean.xlsx"
    End If

    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename(InitialFileName:=defaultPath, _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
        Title:="Export Clean Macro-Free .xlsx")
    If VarType(savePath) = vbBoolean Then Exit Sub

    Application.DisplayAlerts = False
    On Error GoTo ErrHandler
    wb.SaveCopyAs Replace(CStr(savePath), ".xlsx", ".xlsm_tmp_copy")
    ' If current workbook is already .xlsx, SaveAs or Workbook copy ensures macro-free output
    Dim tmpWb As Workbook
    Set tmpWb = Workbooks.Open(Filename:=Replace(CStr(savePath), ".xlsx", ".xlsm_tmp_copy"), ReadOnly:=True)
    tmpWb.SaveAs Filename:=CStr(savePath), FileFormat:=xlOpenXMLWorkbook
    tmpWb.Close SaveChanges:=False
    On Error Resume Next
    Kill Replace(CStr(savePath), ".xlsx", ".xlsm_tmp_copy")
    On Error GoTo 0
    Application.DisplayAlerts = True

    modUnicode.MsgBoxU modUnicode.T("MSG_EXPORT_OK", "VI") & vbCrLf & CStr(savePath), vbInformation
    Exit Sub
ErrHandler:
    Application.DisplayAlerts = True
    On Error Resume Next
    Kill Replace(CStr(savePath), ".xlsx", ".xlsm_tmp_copy")
    On Error GoTo 0
    modUnicode.MsgBoxU "ExportCleanXlsx Error: " & Err.Description, vbCritical
End Sub

Public Sub CreateMasterBOQSheet(Optional ByVal showMsg As Boolean = True)
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("QS_Master_BOQ")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(1))
        ws.Name = "QS_Master_BOQ"
        ws.Cells.Font.Name = "Segoe UI"
        ws.Cells.Font.Size = 10

        With ws.Range("B2:H2")
            .Merge
            .Value = modUnicode.T("BOQ_TITLE")
            .Font.Size = 13
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(22, 54, 92)
            .IndentLevel = 1
            .RowHeight = 30
        End With

        ws.Range("B3").Formula = "=IFERROR(CFG_ProjectName,"""")"
        ws.Range("B3").Font.Bold = True

        RefreshMasterBOQHeaders wb

        With ws.Range("B5:H5")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(47, 85, 151)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 25
        End With

        ' Seed standard WBS sections for Phases 2 & 3
        Dim wbs As Variant
        wbs = Array( _
            Array("I", "01", modUnicode.U("C\u00d4NG T\u00c1C \u0110\u1ea4T & B\u00ca T\u00d4NG L\u00d3T / EARTHWORK & LEAN CONC.")), _
            Array("II", "02", modUnicode.U("K\u1ebeT C\u1ea4U B\u00ca T\u00d4NG C\u1ed0T TH\u00c9P / RC STRUCTURE")), _
            Array("III", "03", modUnicode.U("C\u1ed0T TH\u00c9P & TH\u00c9P H\u00ccNH / REBAR & STRUCTURAL STEEL")), _
            Array("IV", "04", modUnicode.U("C\u00d4NG T\u00c1C HO\u00c0N THI\u1ec6N / FINISHING WORKS")) _
        )
        Dim i As Long, r As Long
        For i = LBound(wbs) To UBound(wbs)
            r = 6 + i
            ws.Cells(r, 2).Value = wbs(i)(0)
            ws.Cells(r, 3).Value = wbs(i)(1)
            ws.Cells(r, 4).Value = wbs(i)(2)
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Font.Bold = True
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Interior.Color = RGB(217, 225, 242)
        Next i

        ws.Columns("A").ColumnWidth = 3
        ws.Columns("B").ColumnWidth = 8
        ws.Columns("C").ColumnWidth = 14
        ws.Columns("D").ColumnWidth = 48
        ws.Columns("E").ColumnWidth = 10
        ws.Columns("F").ColumnWidth = 16
        ws.Columns("G").ColumnWidth = 22
        ws.Columns("H").ColumnWidth = 30
    Else
        RefreshMasterBOQHeaders wb
    End If
End Sub

Public Sub RefreshMasterBOQHeaders(ByVal wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("QS_Master_BOQ")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ws.Range("B2").Value = modUnicode.T("BOQ_TITLE")
    ws.Cells(5, 2).Value = modUnicode.T("BOQ_STT")
    ws.Cells(5, 3).Value = modUnicode.T("BOQ_CODE")
    ws.Cells(5, 4).Value = modUnicode.T("BOQ_DESC")
    ws.Cells(5, 5).Value = modUnicode.T("BOQ_UNIT")
    ws.Cells(5, 6).Value = modUnicode.T("BOQ_QTY")
    ws.Cells(5, 7).Value = modUnicode.T("BOQ_REF")
    ws.Cells(5, 8).Value = modUnicode.T("BOQ_NOTE")
End Sub
