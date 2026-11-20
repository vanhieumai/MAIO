Attribute VB_Name = "modThemeFile"
Option Explicit

'===========================================================
' modFile — File/Folder utilities & Path helpers
' (Refactor: gi? nguyên toàn b? API cu, ch? THÊM helper m?i
'  ph?c v? Theme Manager. Không xóa/d?i tên hàm dã có d? tránh
'  v? các module dang g?i t?i modFile.*)
'===========================================================

Private Const CPMS_APPDATA_FOLDER As String = "CPMS_R\"
Private Const THEMES_FOLDER       As String = "Themes\"
Private Const SETTINGS_FILE       As String = "Settings.ini"
'--------------------------------------------------------------
' Windows Color Picker — thêm vào modFile
'--------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function ChooseColorAPI Lib "comdlg32.dll" _
        Alias "ChooseColorW" (pChoosecolor As CHOOSECOLOR) As Long
    Private Type CHOOSECOLOR
        lStructSize    As Long
        hwndOwner      As LongPtr
        hInstance      As LongPtr
        rgbResult      As Long
        lpCustColors   As LongPtr
        Flags          As Long
        lCustData      As LongPtr
        lpfnHook       As LongPtr
        lpTemplateName As LongPtr
    End Type
#End If

'--------------------------------------------------------------
' OLE picture loader (PNG/JPG/BMP/GIF on dinh hon LoadPicture thuan)
' -- gop tu modImageUtils, dung chung cho LoadImageFromPath ben duoi
'--------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function OleLoadPicturePath Lib "oleaut32.dll" ( _
        ByVal szURLorPath As LongPtr, _
        ByVal punkCaller As LongPtr, _
        ByVal dwReserved As Long, _
        ByVal clrReserved As Long, _
        riid As Any, _
        ByRef ppvRet As IPictureDisp) As Long
#End If

Public Function PickColorDialog(ByVal initialColor As Long) As Long
    Dim cc As CHOOSECOLOR
    Dim custColors(15) As Long
    cc.lStructSize = LenB(cc)
    cc.rgbResult = initialColor
    cc.lpCustColors = VarPtr(custColors(0))
    cc.Flags = &H1 Or &H2           ' CC_RGBINIT | CC_FULLOPEN
    If ChooseColorAPI(cc) <> 0 Then
        PickColorDialog = cc.rgbResult
    Else
        PickColorDialog = -1         ' Cancel
    End If
End Function
' -- Ðu?ng d?n h? th?ng ------------------------------------
Public Function GetCPMSAppFolder() As String
    GetCPMSAppFolder = Environ("LOCALAPPDATA") & "\" & CPMS_APPDATA_FOLDER
End Function

Public Function GetThemesFolder() As String
    GetThemesFolder = GetCPMSAppFolder() & THEMES_FOLDER
End Function

Public Function GetThemeFolder(ByVal themeName As String) As String
    GetThemeFolder = GetThemesFolder() & themeName & "\"
End Function

Public Function GetDefaultThemeFolder() As String
    GetDefaultThemeFolder = GetThemeFolder("Default")
End Function

Public Function GetSettingsFilePath() As String
    GetSettingsFilePath = GetCPMSAppFolder() & SETTINGS_FILE
End Function

' -- File & Folder checks ----------------------------------
Public Function ThemeFileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    ThemeFileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

' Alias gi? tên cu FileExists dã du?c các module khác g?i tr?c ti?p
'Public Function FileExists(ByVal filePath As String) As Boolean
'    FileExists = ThemeFileExists(filePath)
'End Function

Public Function ThemeFolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    ThemeFolderExists = (Len(Dir(folderPath, vbDirectory)) > 0)
    On Error GoTo 0
End Function

' T?o thu m?c (k? c? nested path) b?ng FileSystemObject
Public Sub EnsureFolder(ByVal folderPath As String)
    ' Strip trailing backslash d? x? lý nh?t quán
    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    
    ' Ðã t?n t?i ? không làm gì
    If ThemeFolderExists(folderPath) Then Exit Sub
    
    ' T?o parent tru?c (d? quy — gi?ng mkdir -p)
    Dim parent As String
    Dim pos As Long: pos = InStrRev(folderPath, "\")
    If pos > 1 Then
        parent = Left$(folderPath, pos - 1)
        EnsureFolder parent   ' ? d? quy t?o t?ng c?p m?t
    End If
    
    ' Bây gi? parent dã t?n t?i ? t?o folder này
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    fso.CreateFolder folderPath
    On Error GoTo 0
    Set fso = Nothing
End Sub

' T?o toàn b? c?u trúc thu m?c chu?n c?a 1 Theme
' (Icons, Backgrounds, Ribbon, Preview) — dùng cho Create/Clone/InitDefault
Public Sub EnsureThemeSkeleton(ByVal themeFolder As String)
    EnsureFolder themeFolder
    EnsureFolder themeFolder & "Icons\"
    EnsureFolder themeFolder & "Backgrounds\"
    EnsureFolder themeFolder & "Preview\"
End Sub

' -- Danh sách Theme dã cài --------------------------------
' Tr? v? m?ng tên thu m?c có ch?a theme.ini h?p l?
Public Function ListInstalledThemes() As Variant
    Dim fso       As Object
    Dim folder    As Object
    Dim subFolder As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not ThemeFolderExists(GetThemesFolder()) Then
        ListInstalledThemes = Array()
        Set fso = Nothing
        Exit Function
    End If

    Set folder = fso.GetFolder(GetThemesFolder())

    Dim names()  As String
    Dim count    As Long

    If folder.SubFolders.count = 0 Then
        ListInstalledThemes = Array()
        Set fso = Nothing
        Exit Function
    End If

    ReDim names(folder.SubFolders.count - 1)

    For Each subFolder In folder.SubFolders
        If ThemeFileExists(subFolder.path & "\theme.ini") Then
            names(count) = subFolder.Name
            count = count + 1
        End If
    Next subFolder

    Set fso = Nothing

    If count = 0 Then
        ListInstalledThemes = Array()
    Else
        ReDim Preserve names(count - 1)
        ListInstalledThemes = names
    End If
End Function

' -- Helpers -----------------------------------------------
' Lo?i ký t? không h?p l? trong tên file/folder
Public Function SafeName(ByVal s As String) As String
    Dim bad As Variant, ch As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    SafeName = s
    For Each ch In bad
        SafeName = Join(Split(SafeName, CStr(ch)), "_")
    Next ch
End Function

' Chuy?n màu hex #RRGGBB ? VBA Long (OLE Color)
Public Function HexToColor(ByVal hexStr As String) As Long
    Dim s As String: s = Replace(Trim$(hexStr), "#", "")
    If Len(s) <> 6 Then HexToColor = 0: Exit Function
    On Error Resume Next
    Dim r As Long: r = CLng("&H" & Left$(s, 2))
    Dim g As Long: g = CLng("&H" & mID$(s, 3, 2))
    Dim b As Long: b = CLng("&H" & Right$(s, 2))
    On Error GoTo 0
    HexToColor = RGB(r, g, b)
End Function

' Chi?u ngu?c l?i: VBA Long (RGB) ? "#RRGGBB" — dùng cho Colors tab
Public Function ColorToHex(ByVal oleColor As Long) As String
    Dim r As Long, g As Long, b As Long
    r = oleColor Mod 256
    g = (oleColor \ 256) Mod 256
    b = (oleColor \ 65536) Mod 256
    ColorToHex = "#" & Right$("0" & Hex$(r), 2) & _
                 Right$("0" & Hex$(g), 2) & _
                 Right$("0" & Hex$(b), 2)
End Function

' Load ?nh t? du?ng d?n ? Object (StdPicture / IPictureDisp)
'
' FIX (Reset/Change chap chon, phai bam nhieu lan / dong-mo lai form
' moi thay anh moi): OleLoadPicturePath co CACHE RIENG o cap OLE/OS,
' DOC LAP voi cache noi bo cua modCache (m_ImageCache) - cache nay
' GAN VOI DUONG DAN FILE, khong theo dung noi dung. Khi Reset/Change
' ghi de anh moi vao CUNG 1 filePath cu (vd Icons\Image1.png), OLE co
' the van tra ve anh CU da cache cho duong dan do trong 1 khoang thoi
' gian, du modCache da duoc invalidate dung. De tranh HOAN TOAN cache
' nay, ta luon copy file ra 1 TEN TAM DUY NHAT (khong bao gio trung
' lap) trong %TEMP% truoc khi goi OleLoadPicturePath, roi xoa file tam
' ngay sau khi da doc xong vao memory (Picture object khong con phu
' thuoc file tam nua sau khi nap).
Public Function LoadImageFromPath(ByVal filePath As String) As Object
    If Not ThemeFileExists(filePath) Then Exit Function

    Dim loadPath As String: loadPath = filePath

#If VBA7 Then
    Static uniqueCounter As Long
    uniqueCounter = uniqueCounter + 1

    Dim tempPath As String
    tempPath = Environ$("TEMP") & "\cpms_img_" & Format$(Now, "yyyymmddhhnnss") & _
               "_" & uniqueCounter & "_" & CLng(Timer * 1000) & "." & GetFileExt(filePath)

    On Error Resume Next
    FileCopy filePath, tempPath
    On Error GoTo 0

    If ThemeFileExists(tempPath) Then loadPath = tempPath

    Dim IID(0 To 3) As Long
    Dim pic As IPictureDisp

    IID(0) = &H7BF80981
    IID(1) = &H101ABF32
    IID(2) = &HAA00BB8B
    IID(3) = &HAB0C3000

    On Error Resume Next
    OleLoadPicturePath StrPtr(loadPath), 0, 0, 0, IID(0), pic
    On Error GoTo 0

    ' File tam da duoc doc het vao Picture object (neu load thanh cong)
    ' hoac khong con can nua (neu that bai) - xoa ngay, khong de rac lai.
    If loadPath <> filePath Then DeleteFileSafe loadPath

    If Not pic Is Nothing Then
        Set LoadImageFromPath = pic
        Exit Function
    End If
#End If

    On Error Resume Next
    Set LoadImageFromPath = LoadPicture(filePath)
    On Error GoTo 0
End Function

' L?y thu m?c ch?a file
Public Function GetParentFolder(ByVal filePath As String) As String
    Dim pos As Long: pos = InStrRev(filePath, "\")
    If pos > 0 Then GetParentFolder = Left$(filePath, pos - 1)
End Function

' L?y tên file không có du?ng d?n
Public Function GetFileName(ByVal filePath As String) As String
    Dim pos As Long: pos = InStrRev(filePath, "\")
    GetFileName = IIf(pos > 0, mID$(filePath, pos + 1), filePath)
End Function

' L?y ph?n m? r?ng (không d?u ch?m), vd "png"
Public Function GetFileExt(ByVal filePath As String) As String
    Dim pos As Long: pos = InStrRev(filePath, ".")
    GetFileExt = IIf(pos > 0, mID$(filePath, pos + 1), "")
End Function

' Copy toàn b? n?i dung 1 thu m?c (file + subfolder) sang thu m?c khác
' Dùng chung cho InstallTheme / CreateNewTheme / CloneTheme
Public Sub CopyFolderTree(ByVal src As String, ByVal dst As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not ThemeFolderExists(src) Then Set fso = Nothing: Exit Sub
    EnsureFolder dst

    Dim f As Object
    For Each f In fso.GetFolder(src).Files
        On Error Resume Next
        fso.CopyFile f.path, dst & "\" & f.Name, True
        On Error GoTo 0
    Next f

    Dim subFolder As Object
    For Each subFolder In fso.GetFolder(src).SubFolders
        CopyFolderTree subFolder.path, dst & "\" & subFolder.Name
    Next subFolder
    Set fso = Nothing
End Sub

' T?ng dung lu?ng 1 thu m?c (bytes), d? quy
Public Function GetFolderSizeBytes(ByVal folderPath As String) As Double
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If ThemeFolderExists(folderPath) Then
        GetFolderSizeBytes = fso.GetFolder(folderPath).Size
    End If
    Set fso = Nothing
    On Error GoTo 0
End Function

' Ð?m s? file có ph?n m? r?ng thu?c danh sách (vd "png|jpg|bmp|gif")
' trong 1 thu m?c (không d? quy subfolder)
Public Function CountFilesByExt(ByVal folderPath As String, ByVal extList As String) As Long
    Dim n As Long
    If Not ThemeFolderExists(folderPath) Then Exit Function
    Dim exts() As String: exts = Split(extList, "|")
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim f As Object
    For Each f In fso.GetFolder(folderPath).Files
        Dim e As Variant
        For Each e In exts
            If LCase$(GetFileExt(f.Name)) = LCase$(Replace(CStr(e), ".", "")) Then
                n = n + 1
                Exit For
            End If
        Next e
    Next f
    Set fso = Nothing
    CountFilesByExt = n
End Function

' Xóa 1 file an toàn (không báo l?i n?u không t?n t?i)
Public Sub DeleteFileSafe(ByVal filePath As String)
    On Error Resume Next
    If ThemeFileExists(filePath) Then Kill filePath
    On Error GoTo 0
End Sub

' Ð?i tên/di chuy?n file an toàn
Public Function RenameFileSafe(ByVal oldPath As String, ByVal newPath As String) As Boolean
    On Error Resume Next
    Name oldPath As newPath
    RenameFileSafe = (Err.Number = 0)
    On Error GoTo 0
End Function

' Ngày t?o / ngày s?a d?i c?a 1 file ho?c thu m?c ("" n?u không có)
Public Function GetPathCreatedDate(ByVal p As String) As String
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If ThemeFolderExists(p) Then
        GetPathCreatedDate = Format$(fso.GetFolder(p).DateCreated, "yyyy-mm-dd hh:nn")
    ElseIf fso.FileExists(p) Then
        GetPathCreatedDate = Format$(fso.GetFile(p).DateCreated, "yyyy-mm-dd hh:nn")
    End If
    Set fso = Nothing
    On Error GoTo 0
End Function

Public Function GetPathModifiedDate(ByVal p As String) As String
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If ThemeFolderExists(p) Then
        GetPathModifiedDate = Format$(fso.GetFolder(p).DateLastModified, "yyyy-mm-dd hh:nn")
    ElseIf fso.FileExists(p) Then
        GetPathModifiedDate = Format$(fso.GetFile(p).DateLastModified, "yyyy-mm-dd hh:nn")
    End If
    Set fso = Nothing
    On Error GoTo 0
End Function

Public Sub EnsureSettingsFile()

    Dim filePath As String
    filePath = GetSettingsFilePath()

    ' File already exists -> nothing to do
    If ThemeFileExists(filePath) Then Exit Sub

    ' Make sure the parent folder exists first
    EnsureFolder GetCPMSAppFolder()

    ' Write minimal default INI
    Dim ff As Integer
    ff = FreeFile()
    On Error GoTo EH_Write

    Open filePath For Output As #ff
    Print #ff, "[Theme]"
    Print #ff, "CurrentTheme=Default"
    Print #ff, "DefaultTheme=Default"
    Close #ff

    Exit Sub

EH_Write:
    On Error Resume Next
    Close #ff
    On Error GoTo 0
    Debug.Print "EnsureSettingsFile: could not create " & filePath & _
                " — " & Err.Description

End Sub


