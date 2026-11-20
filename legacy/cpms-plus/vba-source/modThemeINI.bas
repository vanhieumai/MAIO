Attribute VB_Name = "modThemeINI"
Option Explicit

'===========================================================
' modINI — INI File Reader / Writer (UTF-8)
' (Gi? nguyên toàn b? logic cu — ch? THÊM INIDeleteKey và
'  INIWriteSection d? Theme Manager ghi hàng lo?t key m?t l?n,
'  ví d? colors.ini có 20+ key)
'===========================================================

' -- Ð?c 1 giá tr? t? INI ---------------------------------
Private m_FileCache As Object
Private Function GetFileCache() As Object
    If m_FileCache Is Nothing Then
        Set m_FileCache = CreateObject("Scripting.Dictionary")
        m_FileCache.CompareMode = 1
    End If
    Set GetFileCache = m_FileCache
End Function
Public Sub ClearIniFileCache(Optional ByVal filePath As String = "")
    Dim c As Object: Set c = GetFileCache()
    If Len(filePath) = 0 Then
        c.RemoveAll
    ElseIf c.Exists(LCase$(filePath)) Then
        c.Remove LCase$(filePath)
    End If
End Sub
Private Function ReadUTF8Cached(ByVal filePath As String) As String
    Dim c As Object: Set c = GetFileCache()
    Dim k As String: k = LCase$(filePath)
    If c.Exists(k) Then
        ReadUTF8Cached = c(k)
    Else
        ReadUTF8Cached = ReadUTF8(filePath)
        c.Add k, ReadUTF8Cached
    End If
End Function
Public Function INIRead(ByVal filePath As String, _
                        ByVal section As String, _
                        ByVal key As String) As String
    If Not ThemeFileExists(filePath) Then Exit Function

    Dim lines()   As String
    Dim line      As String
    Dim inSection As Boolean
    Dim eqPos     As Long
    Dim i         As Long

    lines = Split(ReadUTF8Cached(filePath), vbNewLine)

    For i = 0 To UBound(lines)
        line = Trim$(lines(i))
        If Len(line) = 0 Or Left$(line, 1) = ";" Or Left$(line, 1) = "#" Then
            ' skip
        ElseIf Left$(line, 1) = "[" Then
            inSection = (LCase$(line) = LCase$("[" & section & "]"))
        ElseIf inSection Then
            eqPos = InStr(line, "=")
            If eqPos > 0 Then
                If LCase$(Trim$(Left$(line, eqPos - 1))) = LCase$(key) Then
                    INIRead = Trim$(mID$(line, eqPos + 1))
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' -- Ð?c toàn b? section ? Dictionary {key: value} --------
Public Function INIReadSection(ByVal filePath As String, _
                               ByVal section As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1   ' case-insensitive

    If Not ThemeFileExists(filePath) Then Set INIReadSection = dict: Exit Function

    Dim lines()   As String
    Dim line      As String
    Dim inSection As Boolean
    Dim eqPos     As Long
    Dim i         As Long

    lines = Split(ReadUTF8Cached(filePath), vbNewLine)

    For i = 0 To UBound(lines)
        line = Trim$(lines(i))
        If Len(line) = 0 Or Left$(line, 1) = ";" Or Left$(line, 1) = "#" Then
            ' skip
        ElseIf Left$(line, 1) = "[" Then
            If inSection Then Exit For       ' dã qua section ? d?ng
            inSection = (LCase$(line) = LCase$("[" & section & "]"))
        ElseIf inSection Then
            eqPos = InStr(line, "=")
            If eqPos > 0 Then
                Dim k As String: k = Trim$(Left$(line, eqPos - 1))
                Dim v As String: v = Trim$(mID$(line, eqPos + 1))
                If Len(k) > 0 And Not dict.Exists(k) Then dict.Add k, v
            End If
        End If
    Next i

    Set INIReadSection = dict
End Function

' -- Ghi 1 giá tr? vào INI (t?o m?i n?u chua có) ---------
Public Sub INIWrite(ByVal filePath As String, _
                    ByVal section As String, _
                    ByVal key As String, _
                    ByVal value As String)
    Dim content As String
    If ThemeFileExists(filePath) Then content = ReadUTF8(filePath)

    Dim lines()   As String
    If Len(content) > 0 Then
        lines = Split(content, vbNewLine)
    Else
        ReDim lines(0): lines(0) = ""
    End If

    Dim secTarget   As String: secTarget = "[" & section & "]"
    Dim secLine     As Long: secLine = -1
    Dim keyLine     As Long: keyLine = -1
    Dim nextSec     As Long: nextSec = -1
    Dim i           As Long
    Dim ln          As String

    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If LCase$(ln) = LCase$(secTarget) Then
            secLine = i
        ElseIf secLine >= 0 And Left$(ln, 1) = "[" And i > secLine And nextSec = -1 Then
            nextSec = i
        End If
        If secLine >= 0 And keyLine = -1 And nextSec = -1 Then
            Dim eq As Long: eq = InStr(ln, "=")
            If eq > 0 Then
                If LCase$(Trim$(Left$(ln, eq - 1))) = LCase$(key) Then keyLine = i
            End If
        End If
    Next i

    Dim result As String

    If secLine = -1 Then
        ' Section chua có ? thêm cu?i file
        result = content & vbCrLf & secTarget & vbCrLf & key & "=" & value & vbCrLf

    ElseIf keyLine >= 0 Then
        ' Key dã có ? c?p nh?t
        lines(keyLine) = key & "=" & value
        result = Join(lines, vbNewLine)

    ElseIf nextSec = -1 Then
        ' Section có, key chua có, không có section ti?p ? thêm cu?i
        result = Join(lines, vbNewLine) & vbCrLf & key & "=" & value

    Else
        ' Chèn tru?c section ti?p theo
        Dim newLines() As String
        ReDim newLines(UBound(lines) + 1)
        For i = 0 To nextSec - 1: newLines(i) = lines(i): Next i
        newLines(nextSec) = key & "=" & value
        For i = nextSec To UBound(lines): newLines(i + 1) = lines(i): Next i
        result = Join(newLines, vbNewLine)
    End If

    WriteUTF8 filePath, result
    ClearIniFileCache filePath
End Sub

' -- Ghi c? m?t Dictionary key=value vào 1 section --------
' Dùng khi Theme Manager luu 15-20 màu cùng lúc (tab Colors)
' overwriteSection = True ? xóa h?t key cu c?a section r?i ghi l?i t? d?u
'                   = False ? ch? INIWrite t?ng key (gi? key khác không li?t kê)
Public Sub INIWriteSection(ByVal filePath As String, _
                           ByVal section As String, _
                           ByVal data As Object, _
                           Optional ByVal overwriteSection As Boolean = True)
    If overwriteSection Then
        RemoveSectionBlock filePath, section
    End If
    Dim k As Variant
    For Each k In data.keys
        INIWrite filePath, section, CStr(k), CStr(data(k))
    Next k
End Sub

' -- Xóa 1 key kh?i section (th?t, không ph?i ghi "0") ----
Public Sub INIDeleteKey(ByVal filePath As String, _
                        ByVal section As String, _
                        ByVal key As String)
    If Not ThemeFileExists(filePath) Then Exit Sub

    Dim lines()   As String: lines = Split(ReadUTF8(filePath), vbNewLine)
    Dim outLines() As String: ReDim outLines(UBound(lines))
    Dim outCount  As Long
    Dim inSection As Boolean
    Dim ln As String, eqPos As Long
    Dim i As Long

    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If Left$(ln, 1) = "[" Then
            inSection = (LCase$(ln) = LCase$("[" & section & "]"))
            outLines(outCount) = lines(i): outCount = outCount + 1
        ElseIf inSection Then
            eqPos = InStr(ln, "=")
            Dim isKeyMatch As Boolean: isKeyMatch = False
            ' FIX: VBA "And" KHONG short-circuit -- neu viet chung
            ' "eqPos > 0 And Left$(ln, eqPos - 1) = ..." tren 1 dong,
            ' Left$ van bi tinh ngay ca khi eqPos = 0 (=> Left$(ln, -1)
            ' => Run-time error 5). Phai tach rieng nhu duoi day.
            If eqPos > 0 Then
                isKeyMatch = (LCase$(Trim$(Left$(ln, eqPos - 1))) = LCase$(key))
            End If
            If isKeyMatch Then
                ' skip dòng này = xóa key
            Else
                outLines(outCount) = lines(i): outCount = outCount + 1
            End If
        Else
            outLines(outCount) = lines(i): outCount = outCount + 1
        End If
    Next i

    If outCount = 0 Then
        WriteUTF8 filePath, ""
    Else
        ReDim Preserve outLines(outCount - 1)
        WriteUTF8 filePath, Join(outLines, vbNewLine)
    End If
    ClearIniFileCache filePath
End Sub

' -- Xóa toàn b? n?i dung 1 section (private helper) ------
Private Sub RemoveSectionBlock(ByVal filePath As String, ByVal section As String)
    If Not ThemeFileExists(filePath) Then Exit Sub

    Dim lines()   As String: lines = Split(ReadUTF8(filePath), vbNewLine)
    Dim outLines() As String: ReDim outLines(UBound(lines))
    Dim outCount  As Long
    Dim skipping  As Boolean
    Dim ln As String
    Dim i As Long

    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If Left$(ln, 1) = "[" Then
            skipping = (LCase$(ln) = LCase$("[" & section & "]"))
            If Not skipping Then
                outLines(outCount) = lines(i): outCount = outCount + 1
            End If
        ElseIf Not skipping Then
            outLines(outCount) = lines(i): outCount = outCount + 1
        End If
    Next i

    If outCount = 0 Then
        WriteUTF8 filePath, ""
        ClearIniFileCache filePath
    Else
        ReDim Preserve outLines(outCount - 1)
        WriteUTF8 filePath, Join(outLines, vbNewLine)
        ClearIniFileCache filePath
    End If
End Sub

' -- I/O UTF-8 ---------------------------------------------
Public Function ReadUTF8(ByVal filePath As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    With st
        .Type = 2: .Charset = "UTF-8": .Open
        .LoadFromFile filePath
        Dim raw As String: raw = .ReadText
        .Close
    End With
    ' Normalize line endings
    raw = Replace(raw, vbCrLf, vbLf)
    raw = Replace(raw, vbCr, vbLf)
    ReadUTF8 = Replace(raw, vbLf, vbNewLine)
End Function

Public Sub WriteUTF8(ByVal filePath As String, ByVal content As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    With st
        .Type = 2: .Charset = "UTF-8": .Open
        .WriteText content
        .SaveToFile filePath, 2   ' adSaveCreateOverWrite
        .Close
    End With
End Sub




