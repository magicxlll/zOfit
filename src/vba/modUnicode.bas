Attribute VB_Name = "modUnicode"
Option Explicit

' ==============================================================================
' modUnicode - Cross-version (32/64-bit) Unicode & Bilingual I18N Engine
' Pure ASCII source code -> 100% immune to Windows ANSI codepage corruption.
' ==============================================================================

#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpText As LongPtr, _
        ByVal lpCaption As LongPtr, _
        ByVal wType As Long) As Long
    Private Declare PtrSafe Function FindWindowW Lib "user32" ( _
        ByVal lpClassName As LongPtr, _
        ByVal lpWindowName As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetWindowTextW Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpString As LongPtr) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal lpText As Long, _
        ByVal lpCaption As Long, _
        ByVal wType As Long) As Long
    Private Declare Function FindWindowW Lib "user32" ( _
        ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long
    Private Declare Function SetWindowTextW Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal lpString As Long) As Long
#End If

Private mI18nCache As Object

' Convert \uXXXX escape sequence in ASCII string to true VBA UTF-16 Unicode String
Public Function U(ByVal escapedText As String) As String
    Dim i As Long, n As Long
    Dim res As String
    Dim hexCode As String
    n = Len(escapedText)
    i = 1
    Do While i <= n
        If i + 5 <= n And Mid$(escapedText, i, 2) = "\u" Then
            hexCode = Mid$(escapedText, i + 2, 4)
            res = res & ChrW$(CLng("&H" & hexCode))
            i = i + 6
        Else
            res = res & Mid$(escapedText, i, 1)
            i = i + 1
        End If
    Loop
    U = res
End Function

' Native Unicode MessageBox (works flawlessly on any Windows locale)
Public Function MsgBoxU(ByVal promptText As String, _
                        Optional ByVal buttons As Long = vbOKOnly + vbInformation, _
                        Optional ByVal titleText As String = "") As Long
    Dim ttl As String
    If Len(titleText) = 0 Then
        ttl = T("APP_TITLE", "VI")
    Else
        ttl = titleText
    End If
    #If VBA7 Then
        Dim hApp As LongPtr
        hApp = Application.hwnd
    #Else
        Dim hApp As Long
        hApp = Application.hwnd
    #End If
    MsgBoxU = MessageBoxW(hApp, StrPtr(promptText), StrPtr(ttl), buttons)
End Function

' Set true UTF-16 Unicode title bar text on any MSForms UserForm (ThunderDFrame)
Public Sub SetFormCaptionU(ByVal frm As Object, ByVal titleUnicode As String)
    On Error Resume Next
    Dim tagCap As String
    tagCap = "ZOFIT_FRM_" & CStr(ObjPtr(frm))
    frm.Caption = tagCap
    #If VBA7 Then
        Dim hFrm As LongPtr
    #Else
        Dim hFrm As Long
    #End If
    hFrm = FindWindowW(StrPtr("ThunderDFrame"), StrPtr(tagCap))
    If hFrm <> 0 Then
        SetWindowTextW hFrm, StrPtr(titleUnicode)
    Else
        frm.Caption = titleUnicode
    End If
    On Error GoTo 0
End Sub

' Load DB_I18N dictionary from hidden sheet inside ThisWorkbook (cached)
Private Sub EnsureI18nLoaded()
    If Not mI18nCache Is Nothing Then Exit Sub
    Set mI18nCache = CreateObject("Scripting.Dictionary")
    mI18nCache.CompareMode = vbTextCompare

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("DB_I18N")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim arr As Variant
    arr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 3)).Value
    Dim r As Long, k As String
    For r = 1 To UBound(arr, 1)
        k = Trim$(CStr(arr(r, 1)))
        If Len(k) > 0 Then
            mI18nCache(k & "|VI") = CStr(arr(r, 2))
            mI18nCache(k & "|EN") = CStr(arr(r, 3))
        End If
    Next r
End Sub

Public Sub ClearI18nCache()
    Set mI18nCache = Nothing
End Sub

' Get active workbook language mode: "VI", "EN", or "BI" (Bilingual VI / EN)
Public Function GetActiveLangMode() As String
    Dim modeVal As String
    modeVal = "VI"
    On Error Resume Next
    If Not ActiveWorkbook Is Nothing Then
        Dim nm As Name
        Set nm = ActiveWorkbook.Names("CFG_Lang")
        If Not nm Is Nothing Then
            Dim v As String
            v = UCase$(Trim$(CStr(nm.RefersToRange.Value)))
            If InStr(1, v, "BI", vbTextCompare) > 0 Or InStr(1, v, "SONG", vbTextCompare) > 0 Then
                modeVal = "BI"
            ElseIf InStr(1, v, "EN", vbTextCompare) > 0 Then
                modeVal = "EN"
            Else
                modeVal = "VI"
            End If
        End If
    End If
    On Error GoTo 0
    GetActiveLangMode = modeVal
End Function

' Translate key from DB_I18N according to active mode or explicit lang ("VI", "EN", "BI")
Public Function T(ByVal key As String, Optional ByVal langMode As String = "") As String
    EnsureI18nLoaded
    Dim mode As String
    If Len(langMode) = 0 Then
        mode = GetActiveLangMode()
    Else
        mode = UCase$(langMode)
    End If

    Dim sVI As String, sEN As String
    If mI18nCache.Exists(key & "|VI") Then sVI = mI18nCache(key & "|VI") Else sVI = key
    If mI18nCache.Exists(key & "|EN") Then sEN = mI18nCache(key & "|EN") Else sEN = sVI

    Select Case mode
        Case "EN"
            T = sEN
        Case "BI"
            If StrComp(sVI, sEN, vbTextCompare) = 0 Then
                T = sVI
            Else
                T = sVI & vbLf & sEN
            End If
        Case Else
            T = sVI
    End Select
End Function

' Helper: verify an active user workbook exists (not the hidden Add-in itself)
Public Function EnsureActiveWorkbook() As Boolean
    If ActiveWorkbook Is Nothing Then
        MsgBoxU T("MSG_NO_WB", "VI"), vbExclamation
        EnsureActiveWorkbook = False
        Exit Function
    End If
    If ActiveWorkbook Is ThisWorkbook Then
        MsgBoxU T("MSG_NO_WB", "VI"), vbExclamation
        EnsureActiveWorkbook = False
        Exit Function
    End If
    EnsureActiveWorkbook = True
End Function
