Attribute VB_Name = "modBoqLinker"
Option Explicit

' ==============================================================================
' modBoqLinker - Safe Map-Link Engine from *_Detail & TH_Thep_ToanCT to Client BOQ
' Never alters Client BOQ layout/formatting or GAP (A-B) columns
' ==============================================================================

Public Sub ShowBoqLinker()
    If Not modUnicode.EnsureActiveWorkbook() Then Exit Sub
    frmBoqLinker.Show vbModeless
    frmBoqLinker.RefreshTargetCell
End Sub

Public Function CollectLinkableSources(ByVal wb As Workbook) As Variant
    Dim rowsList As New Collection
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If Right$(ws.Name, 7) = "_Detail" Then
            Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            Dim r As Long
            For r = 5 To lr
                Dim tag As String: tag = Trim$(CStr(ws.Cells(r, 13).Value))
                Dim qtyF As String: qtyF = Trim$(CStr(ws.Cells(r, 12).Formula))
                Dim desc As String: desc = Trim$(CStr(ws.Cells(r, 2).Value))
                Dim unit As String: unit = Trim$(CStr(ws.Cells(r, 4).Value))
                If (Left$(tag, 4) = "PILE" Or Left$(tag, 4) = "SOIL" Or Left$(tag, 2) = "FW" Or Left$(tag, 4) = "CONC" Or Left$(tag, 3) = "FIN" Or Left$(tag, 3) = "INF") And Len(desc) > 0 And Len(unit) > 0 Then
                    Dim valNum As Double: valNum = 0
                    On Error Resume Next
                    valNum = CDbl(ws.Cells(r, 12).Value2)
                    On Error GoTo 0
                    Dim refAddr As String: refAddr = "'" & ws.Name & "'!L" & r
                    rowsList.Add Array(ws.Name, desc, unit, Format$(valNum, "#,##0.00"), refAddr)
                End If
            Next r
        ElseIf StrComp(ws.Name, "TH_Thep_ToanCT", vbTextCompare) = 0 Then
            Dim lr2 As Long: lr2 = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
            Dim r2 As Long
            For r2 = 5 To lr2
                Dim lbl As String: lbl = Trim$(CStr(ws.Cells(r2, 3).Value))
                If lbl = "Substructure" Or lbl = "Superstructure" Or lbl = "Misc/Stair/Lintel" Or lbl = "Civil" Then
                    Dim tonVal As Double: tonVal = 0
                    On Error Resume Next
                    tonVal = CDbl(ws.Cells(r2, 7).Value2)
                    On Error GoTo 0
                    rowsList.Add Array("TH_Thep_ToanCT", "Rebar Net [" & lbl & "]", "tons", Format$(tonVal, "#,##0.000"), "'TH_Thep_ToanCT'!G" & r2)
                    rowsList.Add Array("TH_Thep_ToanCT", "Rebar Gross + Scrap [" & lbl & "]", "tons", Format$(CDbl(ws.Cells(r2, 8).Value2), "#,##0.000"), "'TH_Thep_ToanCT'!H" & r2)
                End If
            Next r2
        End If
    Next ws

    If rowsList.Count = 0 Then Exit Function
    Dim res() As Variant
    ReDim res(0 To rowsList.Count - 1, 0 To 4)
    Dim i As Long, item As Variant
    For i = 1 To rowsList.Count
        item = rowsList(i)
        res(i - 1, 0) = item(0)
        res(i - 1, 1) = item(1)
        res(i - 1, 2) = item(2)
        res(i - 1, 3) = item(3)
        res(i - 1, 4) = item(4)
    Next i
    CollectLinkableSources = res
End Function

Public Function AssignLinkFormulaToCell(ByVal targetCell As Range, ByVal refAddresses As Variant, Optional ByVal factorExpr As String = "") As String
    If targetCell Is Nothing Then Exit Function
    Dim joined As String
    If IsArray(refAddresses) Then
        joined = Join(refAddresses, "+")
    Else
        joined = CStr(refAddresses)
    End If
    If Len(Trim$(joined)) = 0 Then Exit Function

    Dim formulaStr As String
    Dim fac As String: fac = Trim$(factorExpr)
    If Len(fac) > 0 And fac <> "1" Then
        formulaStr = "=(" & joined & ")" & fac
    Else
        formulaStr = "=" & joined
    End If
    targetCell.Formula = formulaStr
    AssignLinkFormulaToCell = formulaStr
End Function
