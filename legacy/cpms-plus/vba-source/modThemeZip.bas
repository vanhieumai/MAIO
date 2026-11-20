Attribute VB_Name = "modThemeZip"
Option Explicit

'===========================================================
' modZip — Install / Remove / Package .cpmstheme packages
' (Giu nguyen InstallThemeFromFile/RemoveTheme cu — THEM
'  PackageTheme, InstallThemeFromFileEx voi tuy chon
'  Overwrite/Rename/Cancel theo muc 9 cua Upgrade Plan)
'
' CAP NHAT (gop resources.ini + ComponentManifest.ini):
' ComponentManifest.ini gio la file THEO TUNG THEME (nam ngay
' trong theme folder, cung cap voi theme.ini/colors.ini/resources.ini
' cu), KHONG con la 1 file GLOBAL rieng trong Themes\ nua. Vi vay:
'   - PackageTheme: doc va dong goi ComponentManifest.ini truc tiep
'     tu themeFolder (nhu 1 file .ini binh thuong trong officialFiles),
'     resources.ini bi loai khoi danh sach nay (khong con can thiet,
'     du lieu RelFile da nam trong ComponentManifest.ini).
'   - InstallThemeFromFileEx: ComponentManifest.ini gio chi don gian
'     "chay theo" CopyFolderContents cung srcRoot->destFolder giong
'     theme.ini/colors.ini, KHONG can buoc merge/xoa rieng nao nua —
'     MergeComponentManifestFile (danh cho file global truoc day) da
'     bi XOA hoan toan.
'===========================================================
Private Const ENGINE_VERSION As String = "1.0"
'--------------------------------------------------------------
' Windows Color Picker (ChooseColor API)
' Dat trong standard module (modFile hoac modColorPicker)
'--------------------------------------------------------------




' -- Cai dat Theme tu file .cpmstheme (giu lai ban cu, luon Overwrite) ---
Public Function InstallThemeFromFile(ByVal cpmsThemePath As String) As String
    InstallThemeFromFile = InstallThemeFromFileEx(cpmsThemePath, "ask")
End Function

' -- Phien ban mo rong: conflictMode = "ask" | "overwrite" | "rename" | "cancel"
Public Function InstallThemeFromFileEx(ByVal cpmsThemePath As String, _
                                       ByVal conflictMode As String) As String
    If Not ThemeFileExists(cpmsThemePath) Then
        MsgBox "File not found:" & vbCrLf & cpmsThemePath, vbCritical
        Exit Function
    End If

    Dim tempZip     As String
    Dim tempExtract As String

    On Error GoTo EH

    tempZip = Environ("TEMP") & "\cpms_th_" & Format$(Now, "HHmmss") & ".zip"
    tempExtract = Environ("TEMP") & "\cpms_th_ex_" & Format$(Now, "HHmmss") & "\"

    FileCopy cpmsThemePath, tempZip
    EnsureFolder tempExtract
    UnzipTo tempZip, tempExtract

    WaitForFiles tempExtract, 10

    Dim themeIni As String
    themeIni = FindThemeIni(tempExtract)
    If Len(themeIni) = 0 Then
        MsgBox "Invalid package: theme.ini not found.", vbCritical
        GoTo Cleanup
    End If

    Dim themeName As String
    themeName = INIRead(themeIni, "Theme", "Name")
    If Len(Trim$(themeName)) = 0 Then
        MsgBox "Invalid theme.ini: Name is missing.", vbCritical
        GoTo Cleanup
    End If
    themeName = SafeName(themeName)

    Dim minVer As String
    minVer = INIRead(themeIni, "Theme", "MinimumEngineVersion")
    If Len(minVer) > 0 And Not IsCompatible(minVer) Then
        MsgBox "Theme requires Engine v" & minVer & "+." & vbCrLf & _
               "Current Engine: v" & ENGINE_VERSION, vbCritical
        GoTo Cleanup
    End If

    Dim destFolder As String
    destFolder = GetThemeFolder(themeName)

    If ThemeFolderExists(destFolder) Then
        If LCase$(themeName) = "default" Then
            MsgBox "Cannot overwrite the Default theme.", vbExclamation
            GoTo Cleanup
        End If

        Dim mode As String: mode = LCase$(conflictMode)
        If mode = "ask" Then
            Dim ans As Long
            ans = MsgBox("Theme """ & themeName & """ is already installed." & vbCrLf & _
                         "Yes = Overwrite   No = Rename (auto-append a suffix)   Cancel = Cancel", _
                         vbYesNoCancel + vbQuestion, "Install Theme")
            Select Case ans
                Case vbYes:    mode = "overwrite"
                Case vbNo:     mode = "rename"
                Case Else:     mode = "cancel"
            End Select
        End If

        Select Case mode
            Case "cancel"
                GoTo Cleanup

            Case "rename"
                Dim suffix As Long: suffix = 2
                Do While ThemeFolderExists(GetThemeFolder(themeName & "_" & suffix))
                    suffix = suffix + 1
                Loop
                themeName = themeName & "_" & suffix
                destFolder = GetThemeFolder(themeName)

            Case "overwrite"
                Dim fsoD As Object
                Set fsoD = CreateObject("Scripting.FileSystemObject")
                fsoD.DeleteFolder destFolder, True
                Set fsoD = Nothing
        End Select
    End If

    EnsureFolder GetThemesFolder()
    EnsureFolder destFolder

    Dim srcRoot As String
    srcRoot = GetParentFolder(themeIni)
    CopyFolderContents srcRoot, destFolder

    ' Neu da doi ten do conflict "rename" hoac da Overwrite, dam bao
    ' theme.ini ben trong khop Name kh?p ten thu muc thuc su
    INIWrite destFolder & "theme.ini", "Theme", "Name", themeName

    InstallThemeFromFileEx = themeName
    GoTo Cleanup

EH:
    MsgBox "InstallTheme error: " & Err.Description, vbCritical

Cleanup:
    On Error Resume Next
    Kill tempZip
    Dim fso2 As Object
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    If fso2.FolderExists(tempExtract) Then fso2.DeleteFolder tempExtract, True
    Set fso2 = Nothing
    On Error GoTo 0
End Function

' -- Xoa Theme (khong duoc xoa Default) -------------------
Public Function RemoveTheme(ByVal themeName As String) As Boolean
    If LCase$(themeName) = LCase$(GetDefaultThemeName()) Then
        MsgBox "Cannot remove the theme currently set as Default." & vbCrLf & _
               "Please set another theme as Default before removing this one.", vbExclamation
        Exit Function
    End If
    If LCase$(themeName) = LCase$(GetCurrentTheme()) Then
        MsgBox "Cannot remove the theme currently in use (Current theme)." & vbCrLf & _
               "Please apply another theme before removing this one.", vbExclamation
        Exit Function
    End If
    Dim themeFolder As String
    themeFolder = GetThemeFolder(themeName)
    If Not ThemeFolderExists(themeFolder) Then
        MsgBox "Theme not found: " & themeName, vbExclamation
        Exit Function
    End If
    Dim ans As Long
    ans = MsgBox("Remove theme """ & themeName & """?", vbYesNo + vbQuestion, "Remove Theme")
    If ans = vbNo Then Exit Function
    On Error GoTo EH_Remove
    DeleteFolderRecursive themeFolder
    RemoveTheme = True
    Exit Function
EH_Remove:
    MsgBox "Cannot remove theme: " & Err.Description, vbCritical
End Function

' Xoa de quy bang Kill/RmDir — on dinh hon FileSystemObject.DeleteFolder voi path dai
Private Sub DeleteFolderRecursive(ByVal folderPath As String)
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Dim f As String
    f = Dir(folderPath & "*.*", vbNormal)
    Do While Len(f) > 0
        If f <> "." And f <> ".." Then
            On Error Resume Next
            SetAttr folderPath & f, vbNormal   ' bo ReadOnly/Hidden neu co, tranh Permission denied
            Kill folderPath & f
            On Error GoTo 0
        End If
        f = Dir()
    Loop

    Dim sub_ As String
    sub_ = Dir(folderPath, vbDirectory)
    ' Dir() khong liet ke tat subfolder cung luc voi file; quet lai rieng cho thu muc con
    Dim subs As Collection: Set subs = New Collection
    Dim d As String: d = Dir(folderPath, vbDirectory)
    Do While Len(d) > 0
        If d <> "." And d <> ".." Then
            If (GetAttr(folderPath & d) And vbDirectory) = vbDirectory Then subs.Add d
        End If
        d = Dir()
    Loop

    Dim s As Variant
    For Each s In subs
        DeleteFolderRecursive folderPath & s & "\"
    Next s

    RmDir folderPath
End Sub

'----------------------------------------------------------
' PackageTheme — Nen thu muc Theme thanh file .cpmstheme
' (Muc 8 Upgrade Plan). Khong dong goi file tam/cache: chi
' nen dung 4 thu muc con (Icons/Backgrounds/Ribbon/Preview)
' + theme.ini/colors.ini/ComponentManifest.ini (day gio deu la
' file RIENG CUA THEME nay, nam ngay trong themeFolder).
'
' targetFolder: thu muc nguoi dung chon de luu file .cpmstheme
' Tra ve full path file .cpmstheme neu thanh cong, "" neu that bai
'----------------------------------------------------------
Public Function PackageTheme(ByVal themeName As String, _
                             ByVal targetFolder As String) As String
    Dim themeFolder As String: themeFolder = GetThemeFolder(themeName)
    If Not ThemeFolderExists(themeFolder) Then
        MsgBox "Theme not found: " & themeName, vbExclamation
        Exit Function
    End If

    Dim stageFolder As String
    stageFolder = Environ("TEMP") & "\cpms_pkg_" & Format$(Now, "HHmmss") & "\"
    EnsureFolder stageFolder

    On Error GoTo EH

    ' Copy cac file chinh thuc vao stage. ComponentManifest.ini gio la file
    ' RIENG CUA THEME (nam ngay trong themeFolder, nhu theme.ini/colors.ini),
    ' KHONG con doc tu vi tri global nao khac — resources.ini khong con
    ' duoc dong goi (du lieu RelFile da nam trong ComponentManifest.ini).
    Dim officialFiles As Variant
    officialFiles = Array("theme.ini", "colors.ini", "ComponentManifest.ini")
    Dim f As Variant
    For Each f In officialFiles
        If ThemeFileExists(themeFolder & CStr(f)) Then
            FileCopy themeFolder & CStr(f), stageFolder & CStr(f)
        End If
    Next f

    Dim officialFolders As Variant
    officialFolders = Array("Icons", "Backgrounds", "Preview")
    Dim sub_ As Variant
    For Each sub_ In officialFolders
        If ThemeFolderExists(themeFolder & CStr(sub_)) Then
            CopyFolderTree themeFolder & CStr(sub_), stageFolder & CStr(sub_)
        End If
    Next sub_

    EnsureFolder targetFolder
    Dim zipPath As String
    zipPath = targetFolder & SafeName(themeName) & ".zip"
    DeleteFileSafe zipPath

    CreateEmptyZip zipPath

    ' Kiem tra zip duoc tao thanh cong truoc khi dung Shell
    If Not ThemeFileExists(zipPath) Then
        Err.Raise vbObjectError, "PackageTheme", "Cannot create zip file: " & zipPath
    End If

    Dim sh  As Object: Set sh = CreateObject("Shell.Application")
    Dim stageFolderNoSlash As String
    stageFolderNoSlash = Left$(stageFolder, Len(stageFolder) - 1)
    ' BUG FIX: Shell.Application.Namespace() yeu cau Variant, khong phai String
    ' Truyen String se tra Nothing tren nhieu Windows version
    '----

    ' -- Gan Variant (fix Shell.Namespace tra Nothing) --------
    Dim vZipPath     As Variant: vZipPath = zipPath
    Dim vStageFolder As Variant: vStageFolder = stageFolderNoSlash

    Dim dst As Object: Set dst = sh.Namespace(vZipPath)
    Dim src As Object: Set src = sh.Namespace(vStageFolder)

    If dst Is Nothing Then Err.Raise vbObjectError, "PackageTheme", _
        "Cannot open zip namespace: " & zipPath
    If src Is Nothing Then Err.Raise vbObjectError, "PackageTheme", _
        "Cannot open stage namespace: " & stageFolderNoSlash

    ' FIX: dat placeholder vao empty folders truoc khi CopyHere
    ' Shell.Application khong the zip empty folders — se hien warning dialog
    Dim fsoPlh As Object: Set fsoPlh = CreateObject("Scripting.FileSystemObject")
    PlaceholderEmptyFolders fsoPlh, stageFolderNoSlash
    Set fsoPlh = Nothing

    dst.CopyHere src.items(), 4 + 16 + 512 + 1024
    WaitForZipCount zipPath, CountTopLevelItems(stageFolder), 15

    '----

    Dim finalPath As String
    finalPath = targetFolder & SafeName(themeName) & ".cpmstheme"
    DeleteFileSafe finalPath

    ' Kiem tra zip ton tai truoc khi rename
    If Not ThemeFileExists(zipPath) Then
        Err.Raise vbObjectError, "PackageTheme", "Zip file missing after CopyHere: " & zipPath
    End If
    Name zipPath As finalPath

    ' FIX: luu duong dan package vao theme.ini de GetThemeStats dung
    Dim srcThemeIni As String
    srcThemeIni = GetThemeFolder(themeName) & "theme.ini"
    INIWrite srcThemeIni, "Theme", "LastPackagePath", finalPath

    PackageTheme = finalPath
    GoTo Cleanup

EH:
    MsgBox "PackageTheme error: " & Err.Description, vbCritical

Cleanup:
    On Error Resume Next
    ' BUG FIX: guard empty string — fso.DeleteFolder("") van throw du co On Error Resume Next
    If Len(stageFolder) > 0 Then
        Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(stageFolder) Then fso.DeleteFolder stageFolder, True
        Set fso = Nothing
    End If
    On Error GoTo 0
End Function
'--------------------------------------------------------------
' PlaceholderEmptyFolders
' Shell.Application khong the zip empty folders — se hien warning dialog.
' Fix: de quy tim folder rong va dat 1 file placeholder vao.
' Placeholder se nam trong zip — acceptable trade-off.
'--------------------------------------------------------------
Private Sub PlaceholderEmptyFolders(fso As Object, ByVal folderPath As String)
    Dim f       As Object
    Dim subf    As Object
    Dim ph      As String
    Dim ff      As Integer

    Set f = fso.GetFolder(folderPath)

    ' De quy truoc (depth-first), sau do check lai folder nay
    For Each subf In f.SubFolders
        PlaceholderEmptyFolders fso, subf.path
    Next subf

    ' Sau khi de quy: neu folder van khong co files thi tao placeholder
    If f.Files.count = 0 Then
        ph = folderPath & "\.gitkeep"
        ff = FreeFile
        Open ph For Output As #ff
        Close #ff
    End If
End Sub
' -- Private helpers ---------------------------------------
Private Sub UnzipTo(ByVal zipPath As String, ByVal destFolder As String)
    EnsureFolder destFolder
    ' Strip trailing backslash — Shell.Namespace doi khi tu choi path co dau \
    Dim cleanDest As String
    cleanDest = destFolder
    If Right$(cleanDest, 1) = "\" Then cleanDest = Left$(cleanDest, Len(cleanDest) - 1)
    Dim sh As Object: Set sh = CreateObject("Shell.Application")
    ' BUG FIX: phai dung Variant — String se sh.Namespace() tra Nothing silently
    Dim vZip  As Variant: vZip = zipPath
    Dim vDest As Variant: vDest = cleanDest
    Dim zip As Object: Set zip = sh.Namespace(vZip)
    Dim dst As Object: Set dst = sh.Namespace(vDest)
    If zip Is Nothing Then Err.Raise vbObjectError, "UnzipTo", "Cannot open zip: " & zipPath
    If dst Is Nothing Then Err.Raise vbObjectError, "UnzipTo", "Cannot open dest: " & cleanDest
    dst.CopyHere zip.items(), 4 + 16 + 512 + 1024
    Set sh = Nothing: Set zip = Nothing: Set dst = Nothing
End Sub

Private Sub WaitForFiles(ByVal folderPath As String, ByVal maxSec As Long)
    Dim t  As Date: t = Now
    Do
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
        If FindThemeIni(folderPath) <> "" Then Exit Do
    Loop While DateDiff("s", t, Now) < maxSec
End Sub

Private Sub CreateEmptyZip(ByVal zipPath As String)
    ' Dung Open Binary thay ADODB.Stream — khong can tham chieu external
    Dim fileNum As Integer: fileNum = FreeFile
    Dim emptyZip(21) As Byte
    ' PK end-of-central-directory record (22 bytes)
    emptyZip(0) = &H50: emptyZip(1) = &H4B
    emptyZip(2) = &H5: emptyZip(3) = &H6
    ' bytes 4-21 = 0 (da la 0 mac dinh khi ReDim)
    Open zipPath For Binary As #fileNum
    Put #fileNum, 1, emptyZip
    Close #fileNum
End Sub

Private Function ByteArrFromHex(ByVal hexStr As String) As Byte()
    Dim n As Long: n = Len(hexStr) \ 2
    Dim arr() As Byte: ReDim arr(n - 1)
    Dim i As Long
    For i = 0 To n - 1
        arr(i) = CByte("&H" & mID$(hexStr, i * 2 + 1, 2))
    Next i
    ByteArrFromHex = arr
End Function

Private Sub WaitForZipCount(ByVal zipPath As String, ByVal expectedCount As Long, ByVal maxSec As Long)
    Dim t  As Date: t = Now
    Dim sh As Object: Set sh = CreateObject("Shell.Application")
    Dim vZip As Variant: vZip = zipPath          ' FIX: Variant
    Do
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
        Dim ns As Object: Set ns = sh.Namespace(vZip)
        If Not ns Is Nothing Then
            If ns.items().count >= expectedCount Then Exit Do
        End If
    Loop While DateDiff("s", t, Now) < maxSec
    Set sh = Nothing
End Sub

Private Function CountTopLevelItems(ByVal folderPath As String) As Long
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim n As Long
    n = fso.GetFolder(folderPath).Files.count + fso.GetFolder(folderPath).SubFolders.count
    Set fso = Nothing
    CountTopLevelItems = n
End Function

Private Function FindThemeIni(ByVal folderPath As String, Optional ByVal depth As Long = 0) As String
    If ThemeFileExists(folderPath & "theme.ini") Then
        FindThemeIni = folderPath & "theme.ini"
        Exit Function
    End If

    If depth >= 3 Then Exit Function   ' chan de quy vo han / qua sau bat thuong

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Set fso = Nothing: Exit Function

    Dim subFolder As Object
    For Each subFolder In fso.GetFolder(folderPath).SubFolders
        Dim found As String
        found = FindThemeIni(subFolder.path & "\", depth + 1)
        If Len(found) > 0 Then
            FindThemeIni = found
            Set fso = Nothing
            Exit Function
        End If
    Next subFolder
    Set fso = Nothing
End Function

Private Sub CopyFolderContents(ByVal src As String, ByVal dst As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureFolder dst
    Dim f As Object
    For Each f In fso.GetFolder(src).Files
        On Error Resume Next
        fso.CopyFile f.path, dst & "\" & f.Name, True
        On Error GoTo 0
    Next f
    Dim subFolder As Object
    For Each subFolder In fso.GetFolder(src).SubFolders
        CopyFolderContents subFolder.path, dst & "\" & subFolder.Name
    Next subFolder
    Set fso = Nothing
End Sub

Private Function IsCompatible(ByVal required As String) As Boolean
    On Error Resume Next
    Dim reqMaj As Long: reqMaj = CLng(Split(required, ".")(0))
    Dim engMaj As Long: engMaj = CLng(Split(ENGINE_VERSION, ".")(0))
    On Error GoTo 0
    IsCompatible = (engMaj >= reqMaj)
End Function


