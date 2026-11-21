Attribute VB_Name = "modShapeDraw"
Option Explicit

' ==============================================================================
' modShapeDraw - Rebar Diagram Renderer & Light/Print Performance Manager
' Uses embedded Hinhdangthep BMPs (with L1..L7, M1, M2, R labels) cached in
' %TEMP%\zofit_shapes\ so AddPicture never touches the user's Clipboard.
' Supports Light Mode (0-Shape for 1000-5000 rows) & On-Demand Render.
' ==============================================================================

Private Const SHAPE_PREFIX As String = "ZOFIT_BAR_"
Private mCacheDir As String

Private Function EnsureShapeBmpCached(ByVal shapeCode As String) As String
    On Error GoTo ErrHandler
    If Len(mCacheDir) = 0 Then
        mCacheDir = Environ$("TEMP") & "\zofit_shapes"
        If Dir(mCacheDir, vbDirectory) = "" Then MkDir mCacheDir
    End If

    Dim code As String
    code = UCase$(Trim$(shapeCode))
    Dim bmpPath As String
    bmpPath = mCacheDir & "\" & code & ".bmp"

    If Dir(bmpPath) = "" Then
        Dim wasOpen As Boolean, uf As Object
        For Each uf In VBA.UserForms
            If uf.Name = "frmBarShape" Then wasOpen = True: Exit For
        Next uf
        ' Extract embedded IPictureDisp from frmBarShape without showing the form
        Dim picObj As Object
        Set picObj = frmBarShape.Controls("img" & code).Picture
        If Not picObj Is Nothing Then
            SavePicture picObj, bmpPath
        End If
        If Not wasOpen Then Unload frmBarShape
    End If

    If Dir(bmpPath) <> "" Then EnsureShapeBmpCached = bmpPath
    Exit Function
ErrHandler:
    EnsureShapeBmpCached = ""
End Function

Public Sub DeleteBarShapeAtRow(ByVal ws As Worksheet, ByVal rowIdx As Long)
    Dim targetCell As Range
    Set targetCell = ws.Cells(rowIdx, 5) ' Column E

    Dim i As Long
    Dim shp As Shape
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Left$(shp.Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then
            On Error Resume Next
            If Not Application.Intersect(shp.TopLeftCell, targetCell) Is Nothing Then
                shp.Delete
            ElseIf shp.Top >= targetCell.Top - 2 And shp.Top < targetCell.Top + targetCell.Height - 2 Then
                If shp.Left >= targetCell.Left - 5 And shp.Left < targetCell.Left + targetCell.Width Then
                    shp.Delete
                End If
            End If
            On Error GoTo 0
        End If
    Next i
End Sub

Public Sub DrawBarShapeAtRow(ByVal ws As Worksheet, ByVal rowIdx As Long, ByVal shapeCode As String)
    On Error GoTo ErrHandler
    DeleteBarShapeAtRow ws, rowIdx

    Dim code As String
    code = UCase$(Trim$(shapeCode))
    If Left$(code, 1) <> "D" Then Exit Sub

    Dim cell As Range
    Set cell = ws.Cells(rowIdx, 5) ' Column E
    If ws.Rows(rowIdx).RowHeight < 40 Then ws.Rows(rowIdx).RowHeight = 42

    Dim bmpPath As String
    bmpPath = EnsureShapeBmpCached(code)
    If Len(bmpPath) = 0 Then Exit Sub

    Dim padX As Double, padY As Double
    padX = 3: padY = 2
    Dim maxW As Double, maxH As Double
    maxW = cell.Width - 2 * padX
    maxH = cell.Height - 2 * padY
    If maxW < 30 Then maxW = 30
    If maxH < 20 Then maxH = 20

    ' Source BMP aspect ratio is 240:84 (~2.857:1)
    Dim drawW As Double, drawH As Double
    drawH = maxH
    drawW = drawH * (240# / 84#)
    If drawW > maxW Then
        drawW = maxW
        drawH = drawW * (84# / 240#)
    End If

    Dim leftPos As Double, topPos As Double
    leftPos = cell.Left + (cell.Width - drawW) / 2#
    topPos = cell.Top + (cell.Height - drawH) / 2#

    Dim shp As Shape
    Set shp = ws.Shapes.AddPicture(bmpPath, msoFalse, msoTrue, CSng(leftPos), CSng(topPos), CSng(drawW), CSng(drawH))
    If Not shp Is Nothing Then
        shp.Name = SHAPE_PREFIX & code & "_" & rowIdx
        shp.Placement = xlMoveAndSize
    End If
    Exit Sub
ErrHandler:
    ' Fail safely
End Sub

Public Sub RedrawAllBarShapes(ByVal ws As Worksheet, Optional ByVal startRow As Long = 11)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    If lastRow < startRow Then Exit Sub

    Dim r As Long, sc As String
    For r = startRow To lastRow
        sc = UCase$(Trim$(CStr(ws.Cells(r, 4).Value)))
        If Left$(sc, 1) = "D" And Len(sc) >= 2 Then
            DrawBarShapeAtRow ws, r, sc
        Else
            DeleteBarShapeAtRow ws, r
        End If
    Next r
End Sub

Public Sub ClearAllBarShapes(ByVal ws As Worksheet)
    Dim i As Long, shp As Shape
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Left$(shp.Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then
            shp.Delete
        End If
    Next i
End Sub
