Attribute VB_Name = "modEngineBridge"
Option Explicit

' ==============================================================================
' modEngineBridge - Hybrid IPC Bridge between Excel VBA and zofit-engine.exe
' UTF-8 JSON file exchange via ADODB.Stream + Pure-VBA JSON Parser (32/64-bit)
' ==============================================================================

Private mJsonIdx As Long
Private mJsonLen As Long
Private mJsonText As String

Public Function FindEngineExecutable() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim p1 As String, p2 As String, p3 As String, p4 As String
    ' 1. Alongside ThisWorkbook (.xlam) - guarantees matching engine in dist/ or portable bundle
    On Error Resume Next
    If Len(ThisWorkbook.Path) > 0 Then
        p2 = ThisWorkbook.Path & "\zofit-engine.exe"
        If fso.FileExists(p2) Then FindEngineExecutable = p2: Exit Function
        p3 = ThisWorkbook.Path & "\engine\zofit-engine.exe"
        If fso.FileExists(p3) Then FindEngineExecutable = p3: Exit Function
    End If
    On Error GoTo 0

    ' 2. Installed user path (%LOCALAPPDATA%\zOfitQSPro\zofit-engine.exe)
    p1 = Environ$("LOCALAPPDATA") & "\zOfitQSPro\zofit-engine.exe"
    If Len(Environ$("LOCALAPPDATA")) > 0 And fso.FileExists(p1) Then
        FindEngineExecutable = p1
        Exit Function
    End If

    ' 3. Development repo fallback
    p4 = "D:\03_APP\3. System\GITHUB\zOfit\dist\engine\zofit-engine.exe"
    If fso.FileExists(p4) Then
        FindEngineExecutable = p4
        Exit Function
    End If

    FindEngineExecutable = ""
End Function

Public Function RunEngineCommand(ByVal cmdName As String, ByVal reqJson As String, _
                                 ByRef resJson As String, ByRef errMsg As String, _
                                 Optional ByVal winStyle As Long = 0) As Boolean
    On Error GoTo ErrHandler
    resJson = ""
    errMsg = ""

    Dim exePath As String
    exePath = FindEngineExecutable()
    If Len(exePath) = 0 Then
        errMsg = modUnicode.U("Kh\u00f4ng t\u00ecm th\u1ea5y zofit-engine.exe! Vui l\u00f2ng ch\u1ea1y l\u1ea1i b\u1ed9 c\u00e0i zOfit_QS_Pro_Setup.exe.")
        RunEngineCommand = False
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim ipcDir As String
    ipcDir = Environ$("TEMP") & "\zofit_ipc"
    If Not fso.FolderExists(ipcDir) Then fso.CreateFolder ipcDir

    Dim stamp As String
    Randomize
    stamp = Format$(Now, "yyyymmdd_hhnnss") & "_" & Right$("0000" & CStr(Int(Rnd() * 10000)), 4)
    Dim inFile As String, outFile As String
    inFile = ipcDir & "\req_" & stamp & ".json"
    outFile = ipcDir & "\res_" & stamp & ".json"

    WriteUtf8File inFile, reqJson

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim cmdLine As String
    cmdLine = """" & exePath & """ --cmd " & cmdName & " --in """ & inFile & """ --out """ & outFile & """"

    Dim exitCode As Long
    exitCode = sh.Run(cmdLine, winStyle, True)

    If fso.FileExists(outFile) Then
        resJson = ReadUtf8File(outFile)
    End If

    ' Clean temp files
    On Error Resume Next
    If fso.FileExists(inFile) Then fso.DeleteFile inFile, True
    If fso.FileExists(outFile) Then fso.DeleteFile outFile, True
    On Error GoTo 0

    If Len(resJson) = 0 Then
        errMsg = "zofit-engine.exe exited with code " & exitCode & " and empty response."
        RunEngineCommand = False
        Exit Function
    End If

    Dim parsed As Variant
    Set parsed = JsonParseObject(resJson)
    If IsObject(parsed) Then
        If parsed.Exists("ok") Then
            If Not CBool(parsed("ok")) Then
                errMsg = CStr(parsed("message"))
                RunEngineCommand = False
                Exit Function
            End If
        End If
    End If

    RunEngineCommand = (exitCode = 0)
    Exit Function
ErrHandler:
    errMsg = "EngineBridge Error: " & Err.Description
    RunEngineCommand = False
End Function

Public Function TestEnginePing(Optional ByVal showMsg As Boolean = True) As Boolean
    Dim echoStr As String
    echoStr = modUnicode.U("\u0110\u00e0i c\u1ecdc PC1 - B\u00ea t\u00f4ng M400 \u00d825")
    Dim req As String
    req = "{""echo"":""" & JsonEscape(echoStr) & """}"
    Dim res As String, errMsg As String
    If Not RunEngineCommand("ping", req, res, errMsg) Then
        If showMsg Then modUnicode.MsgBoxU errMsg, vbCritical
        TestEnginePing = False
        Exit Function
    End If
    Dim root As Object
    Set root = JsonParseObject(res)
    Dim msg As String
    msg = CStr(root("message"))
    Dim ver As String
    ver = CStr(root("data")("version"))
    If showMsg Then
        modUnicode.MsgBoxU msg & " (v" & ver & ")" & vbCrLf & "Echo: " & CStr(root("data")("echo")), vbInformation
    End If
    TestEnginePing = (CStr(root("data")("echo")) = echoStr)
End Function

' --- UTF-8 File I/O (ADODB.Stream) ---
Public Sub WriteUtf8File(ByVal filePath As String, ByVal text As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText text
    stm.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stm.Close
End Sub

Public Function ReadUtf8File(ByVal filePath As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadUtf8File = stm.ReadText(-1)
    stm.Close
End Function

' --- JSON Serialization Helpers ---
Public Function JsonEscape(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, sb As String
    sb = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch) And &HFFFF&
        Select Case ch
            Case """": sb = sb & "\"""
            Case "\": sb = sb & "\\"
            Case vbCr: sb = sb & "\r"
            Case vbLf: sb = sb & "\n"
            Case vbTab: sb = sb & "\t"
            Case Else
                If code < 32 Then
                    sb = sb & "\u" & Right$("0000" & Hex$(code), 4)
                Else
                    sb = sb & ch
                End If
        End Select
    Next i
    JsonEscape = sb
End Function

Public Function JsonStr(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        JsonStr = """"""
    Else
        JsonStr = """" & JsonEscape(CStr(v)) & """"
    End If
End Function

Public Function JsonNum(ByVal v As Variant) As String
    If IsNumeric(v) Then
        JsonNum = Replace(CStr(CDbl(v)), ",", ".")
    Else
        JsonNum = "0"
    End If
End Function

' --- Pure-VBA JSON Parser (32/64-bit Safe, ByRef Object/Scalar clean) ---
Public Function JsonParseObject(ByVal text As String) As Object
    Dim v As Variant
    ParseJsonRoot text, v
    If IsObject(v) Then
        Set JsonParseObject = v
    Else
        Set JsonParseObject = CreateObject("Scripting.Dictionary")
    End If
End Function

Public Sub ParseJsonRoot(ByVal text As String, ByRef outVal As Variant)
    ' Strip UTF-8 BOM if present
    If Len(text) > 0 Then
        If AscW(Left$(text, 1)) = &HFEFF Then text = Mid$(text, 2)
    End If
    mJsonText = text
    mJsonLen = Len(text)
    mJsonIdx = 1
    SkipWs
    ParseValueInto outVal
End Sub

Private Sub SkipWs()
    Dim c As String
    Do While mJsonIdx <= mJsonLen
        c = Mid$(mJsonText, mJsonIdx, 1)
        If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Then
            mJsonIdx = mJsonIdx + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub ParseValueInto(ByRef outVal As Variant)
    SkipWs
    If IsObject(outVal) Then Set outVal = Nothing
    outVal = Empty
    If mJsonIdx > mJsonLen Then outVal = Null: Exit Sub
    Dim ch As String
    ch = Mid$(mJsonText, mJsonIdx, 1)
    Select Case ch
        Case "{"
            Set outVal = ParseObject()
        Case "["
            Set outVal = ParseArray()
        Case """"
            outVal = ParseString()
        Case "t"
            mJsonIdx = mJsonIdx + 4: outVal = True
        Case "f"
            mJsonIdx = mJsonIdx + 5: outVal = False
        Case "n"
            mJsonIdx = mJsonIdx + 4: outVal = Empty
        Case Else
            outVal = ParseNumber()
    End Select
End Sub

Private Function ParseObject() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    mJsonIdx = mJsonIdx + 1 ' skip {
    SkipWs
    If mJsonIdx <= mJsonLen Then
        If Mid$(mJsonText, mJsonIdx, 1) = "}" Then
            mJsonIdx = mJsonIdx + 1
            Set ParseObject = dict
            Exit Function
        End If
    End If

    Dim k As String, v As Variant, guard As Long
    Do While mJsonIdx <= mJsonLen
        guard = mJsonIdx
        SkipWs
        k = ParseString()
        SkipWs
        If mJsonIdx <= mJsonLen Then
            If Mid$(mJsonText, mJsonIdx, 1) = ":" Then mJsonIdx = mJsonIdx + 1
        End If
        ParseValueInto v
        If IsObject(v) Then
            Set dict(k) = v
            Set v = Nothing
        Else
            dict(k) = v
        End If
        v = Empty
        SkipWs
        If mJsonIdx > mJsonLen Then Exit Do
        Dim ch As String: ch = Mid$(mJsonText, mJsonIdx, 1)
        If ch = "," Then
            mJsonIdx = mJsonIdx + 1
        ElseIf ch = "}" Then
            mJsonIdx = mJsonIdx + 1
            Exit Do
        Else
            Exit Do
        End If
        If mJsonIdx <= guard Then mJsonIdx = guard + 1
    Loop
    Set ParseObject = dict
End Function

Private Function ParseArray() As Collection
    Dim col As New Collection
    mJsonIdx = mJsonIdx + 1 ' skip [
    SkipWs
    If mJsonIdx <= mJsonLen Then
        If Mid$(mJsonText, mJsonIdx, 1) = "]" Then
            mJsonIdx = mJsonIdx + 1
            Set ParseArray = col
            Exit Function
        End If
    End If

    Dim v As Variant, guard As Long
    Do While mJsonIdx <= mJsonLen
        guard = mJsonIdx
        ParseValueInto v
        col.Add v
        v = Empty
        SkipWs
        If mJsonIdx > mJsonLen Then Exit Do
        Dim ch As String: ch = Mid$(mJsonText, mJsonIdx, 1)
        If ch = "," Then
            mJsonIdx = mJsonIdx + 1
        ElseIf ch = "]" Then
            mJsonIdx = mJsonIdx + 1
            Exit Do
        Else
            Exit Do
        End If
        If mJsonIdx <= guard Then mJsonIdx = guard + 1
    Loop
    Set ParseArray = col
End Function

Private Function ParseString() As String
    Dim parts() As String, cap As Long, cnt As Long
    cap = 16: cnt = 0
    ReDim parts(0 To cap - 1)
    mJsonIdx = mJsonIdx + 1 ' skip opening quote

    Dim startPos As Long: startPos = mJsonIdx
    Dim ch As String, esc As String, hex4 As String
    Do While mJsonIdx <= mJsonLen
        ch = Mid$(mJsonText, mJsonIdx, 1)
        If ch = """" Then
            If mJsonIdx > startPos Then
                AppendChunk parts, cnt, cap, Mid$(mJsonText, startPos, mJsonIdx - startPos)
            End If
            mJsonIdx = mJsonIdx + 1
            Exit Do
        ElseIf ch = "\" Then
            If mJsonIdx > startPos Then
                AppendChunk parts, cnt, cap, Mid$(mJsonText, startPos, mJsonIdx - startPos)
            End If
            mJsonIdx = mJsonIdx + 1
            If mJsonIdx > mJsonLen Then Exit Do
            esc = Mid$(mJsonText, mJsonIdx, 1)
            Select Case esc
                Case """", "\", "/": AppendChunk parts, cnt, cap, esc
                Case "b": AppendChunk parts, cnt, cap, ChrW$(8)
                Case "f": AppendChunk parts, cnt, cap, ChrW$(12)
                Case "n": AppendChunk parts, cnt, cap, vbLf
                Case "r": AppendChunk parts, cnt, cap, vbCr
                Case "t": AppendChunk parts, cnt, cap, vbTab
                Case "u"
                    hex4 = Mid$(mJsonText, mJsonIdx + 1, 4)
                    AppendChunk parts, cnt, cap, ChrW$(CLng("&H" & hex4))
                    mJsonIdx = mJsonIdx + 4
                Case Else
                    AppendChunk parts, cnt, cap, esc
            End Select
            mJsonIdx = mJsonIdx + 1
            startPos = mJsonIdx
        Else
            mJsonIdx = mJsonIdx + 1
        End If
    Loop

    If cnt = 0 Then
        ParseString = ""
    Else
        ReDim Preserve parts(0 To cnt - 1)
        ParseString = Join(parts, "")
    End If
End Function

Private Sub AppendChunk(ByRef parts() As String, ByRef cnt As Long, ByRef cap As Long, ByVal chunk As String)
    If cnt >= cap Then
        cap = cap * 2
        ReDim Preserve parts(0 To cap - 1)
    End If
    parts(cnt) = chunk
    cnt = cnt + 1
End Sub

Private Function ParseNumber() As Double
    Dim startPos As Long: startPos = mJsonIdx
    Dim ch As String
    Do While mJsonIdx <= mJsonLen
        ch = Mid$(mJsonText, mJsonIdx, 1)
        If (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "+" Or ch = "." Or ch = "e" Or ch = "E" Then
            mJsonIdx = mJsonIdx + 1
        Else
            Exit Do
        End If
    Loop
    Dim numStr As String
    numStr = Mid$(mJsonText, startPos, mJsonIdx - startPos)
    ParseNumber = Val(numStr)
End Function
