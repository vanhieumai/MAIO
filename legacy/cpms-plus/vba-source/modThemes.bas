Attribute VB_Name = "modThemes"
Option Explicit

'===========================================================
' modTheme — Theme Engine chính
'
' S? d?ng trong m?i Form:
'     Private Sub UserForm_Initialize()
'         ApplyTheme Me           ' áp d?ng màu s?c & ?nh
'         RegisterFormButtons Me  ' dang ký icon button vuông
'     End Sub
'
' (Refactor cho Theme Manager: TOÀN B? logic cu du?c GI? NGUYÊN.
'  Ph?n THÊM M?I n?m t? "PH?N THEME MANAGER" tr? xu?ng: qu?n lý
'  Colors chi ti?t theo control, Theme Info read/write, Create/
'  Clone/Delete Theme, th?ng kê Theme, validation truoc Package.
'  Không t?o module song song — t?t c? d?u g?i l?i modINI/modFile/
'  modCache/modZip nhu cu.)
'===========================================================

Private Const ENGINE_VER    As String = "1.0"
Private Const DEFAULT_THEME As String = "Default"
Private Const SETT_SECTION  As String = "Theme"
Private Const SETT_KEY      As String = "CurrentTheme"
Private Const SETT_DEFAULT_KEY As String = "DefaultTheme"
Private Const IMG_EXTS      As String = ".png|.jpg|.bmp|.gif"

Private m_CurrentTheme As String
Private m_OpenForms As Collection


'===========================================================
' PUBLIC API (CU — GI? NGUYÊN)
'===========================================================

Public Sub InitThemeEngine()
    InitCache
    EnsureFolder GetCPMSAppFolder()
    EnsureFolder GetThemesFolder()
    'EnsureDefaultTheme
    m_CurrentTheme = GetCurrentTheme()
    LoadTheme m_CurrentTheme
End Sub

Public Sub LoadTheme(ByVal themeName As String)
    If Len(Trim$(themeName)) = 0 Then themeName = GetDefaultThemeName()
    If Not ThemeFolderExists(GetThemeFolder(themeName)) Then
        If LCase$(themeName) <> LCase$(GetDefaultThemeName()) Then LoadTheme GetDefaultThemeName()
        Exit Sub
    End If
    m_CurrentTheme = themeName
    InternalLoadThemeInfo themeName
    InternalLoadColors themeName
End Sub

Public Sub ApplyTheme(ByVal frm As Object)
    If Not IsCacheReady() Then InitThemeEngine
    ' Ðang ký form vào registry d? ReloadAllOpenForms bi?t
    If m_OpenForms Is Nothing Then Set m_OpenForms = New Collection
    On Error Resume Next
    m_OpenForms.Add frm, TypeName(frm)  ' key = tên class, t? d?ng b? qua n?u trùng
    On Error GoTo 0
    ApplyThemeInternal frm
End Sub

Public Sub ReloadAllOpenForms()
    If m_OpenForms Is Nothing Then Exit Sub
    Dim deadKeys As Collection: Set deadKeys = New Collection
    Dim f As Object
    For Each f In m_OpenForms
        On Error Resume Next
        Dim chk As Boolean: chk = f.Visible
        If Err.Number <> 0 Then
            deadKeys.Add TypeName(f)
            Err.Clear
        Else
            ApplyThemeInternal f
        End If
        On Error GoTo 0
    Next f
    Dim k As Variant
    For Each k In deadKeys
        On Error Resume Next
        m_OpenForms.Remove CStr(k)
        On Error GoTo 0
    Next k
End Sub
Public Sub ReloadAllOpenFormsold(Optional ByVal excludeFrm As Object = Nothing)
    If m_OpenForms Is Nothing Then Exit Sub
    Dim deadKeys As Collection: Set deadKeys = New Collection
    Dim f As Object

    For Each f In m_OpenForms
        On Error Resume Next
        Dim chk As Boolean: chk = f.Visible
        If Err.Number <> 0 Then
            deadKeys.Add TypeName(f)
            Err.Clear
        ElseIf Not excludeFrm Is Nothing And (f Is excludeFrm) Then
            ' FIX: b? Application.OnTime (macro name resolution không ?n d?nh
            ' gi?a workbook/add-in). Thay b?ng DoEvents d? "x?" h?t message
            ' queue (Click event dang x? lý d?) tru?c khi t? apply lên chính
            ' form dang g?i — tránh xung d?t message loop gây kh?t chu?t.
            DoEvents
            ApplyThemeInternal f
        Else
            ApplyThemeInternal f
        End If
        On Error GoTo 0
    Next f

    Dim k As Variant
    For Each k In deadKeys
        On Error Resume Next
        m_OpenForms.Remove CStr(k)
        On Error GoTo 0
    Next k
End Sub

' ---------------------------------------------------------------
' ApplyThemeInternal — FIX: thêm LoadPageBackgrounds frm
'
' TRUOC FIX:
'   ApplyColorsToControls  — màu s?c (dã x? lý pg.BackColor bên trong)
'   LoadBackground         — background c?a chính UserForm
'   LoadControlImages      — icon/?nh cho các control thông thu?ng
'   (THI?U: background cho Page bên trong MultiPage)
'
' SAU FIX:
'   Thêm LoadPageBackgrounds frm ? cu?i — x? lý pg.Picture cho m?i
'   Page trong MultiPage, dùng key "TypeName(frm)_pageName" kh?p
'   chính xác v?i resource name mà modResourceScanner ghi vào manifest.
' ---------------------------------------------------------------
Private Sub ApplyThemeInternal(ByVal frm As Object)
    On Error GoTo EH_Apply
    On Error Resume Next
    frm.BackColor = GetThemeColor("Background")
    On Error GoTo EH_Apply
    ApplyColorsToControls frm.Controls
    LoadBackground frm
    LoadControlImages frm.Controls
    LoadPageBackgrounds frm      ' FIX: ap background Picture cho cac Page trong MultiPage
    Exit Sub
EH_Apply:
    Debug.Print "ApplyTheme [" & TypeName(frm) & "]: " & Err.Description
End Sub

Public Sub UnregisterForm(ByVal frm As Object)
    If m_OpenForms Is Nothing Then Exit Sub
    On Error Resume Next
    m_OpenForms.Remove TypeName(frm)
    On Error GoTo 0
End Sub
'Public Sub RegisterFormButtons(ByVal frm As Object)
'    On Error Resume Next
'    modIconRegistry.RegisterFormButtons frm
'    On Error GoTo 0
'End Sub

'--------------------------------------------------------------
' LoadBackground — FIX: luôn clear tru?c, r?i load t? theme folder
' Convention: Backgrounds\{TypeName(frm)}.png|jpg|bmp|gif
'--------------------------------------------------------------
Public Sub LoadBackground(ByVal frm As Object)
    ' Bu?c 1: luôn clear picture cu (embedded design-time) tru?c
    ' ? d?m b?o workbook không gi? ?nh nhúng dù theme có ?nh hay không
    On Error Resume Next
    Dim imgCtrl As Object
    Set imgCtrl = frm.Controls("imgBackground")
    If Not imgCtrl Is Nothing Then
        Set imgCtrl.Picture = Nothing
    Else
        Set frm.Picture = Nothing
    End If
    On Error GoTo 0

    ' Bu?c 2: tìm ?nh trong theme folder theo tên form (TypeName)
    ' Convention: Themes\{CurrentTheme}\Backgrounds\frmMain.png
    Dim imgPath As String
    imgPath = ResolveImagePath(TypeName(frm), "Backgrounds")
    If Len(imgPath) = 0 Then Exit Sub   ' Không có background ? gi? tr?ng (dúng ý)

    Dim img As Object
    Set img = GetOrLoadImage(imgPath)
    If img Is Nothing Then Exit Sub

    ' Bu?c 3: gán ?nh t? theme
    On Error Resume Next
    If Not imgCtrl Is Nothing Then
        Set imgCtrl.Picture = img
        imgCtrl.PictureSizeMode = 1
    Else
        Set frm.Picture = img
        frm.PictureSizeMode = 1
        frm.PictureAlignment = 2
    End If
    On Error GoTo 0
End Sub

'--------------------------------------------------------------
' LoadPageBackgrounds — FIX: ap background Picture cho cac Page
' ben trong MultiPage.
'
' NGUYEN NHAN LOI:
'   frm.Controls la danh sach PHANG chi chua cac control truc tiep
'   tren form. Page (tab) ben trong MultiPage KHONG co trong danh
'   sach nay. LoadControlImages chi xu ly pg.Controls (icon/anh cua
'   cac control TREN page) nhung bo qua pg.Picture (background cua
'   chinh page do). LoadBackground chi xu ly frm.Picture.
'   => Ket qua: dau sau ReplaceResourceImage copy anh moi vao
'      thu muc theme thanh cong, pg.Picture khong bao gio duoc cap
'      nhat khi ApplyTheme chay.
'
' CACH SUA:
'   Duyet frm.Controls tim cac MultiPage, voi moi Page ben trong
'   xay dung resource key = TypeName(frm) & "_" & pg.Name
'   (khop chinh xac voi pattern modResourceScanner ghi vao manifest
'   va BackupOriginalPictures dung de backup anh goc).
'   Goi ResolveImagePath de tim file anh trong thu muc Backgrounds\,
'   roi gan vao pg.Picture nhu cach LoadBackground xu ly frm.Picture.
'--------------------------------------------------------------
Private Sub LoadPageBackgrounds(ByVal frm As Object)
    Dim ctrl      As Object
    Dim pg        As Object
    Dim pgResName As String
    Dim imgPath   As String
    Dim img       As Object

    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "MultiPage" Then
            For Each pg In ctrl.Pages
                ' Key: vi du "frmThemeManager_resources" — khop voi ten
                ' resource trong lstResources (Category = Background)
                pgResName = TypeName(frm) & "_" & pg.Name

                ' Buoc 1: luon clear Picture cu truoc (tuong tu LoadBackground)
                On Error Resume Next
                Set pg.Picture = Nothing
                On Error GoTo 0

                ' Buoc 2: tim file anh trong Backgrounds\ cua theme hien tai
                imgPath = ResolveImagePath(pgResName, "Backgrounds")
                If Len(imgPath) = 0 Then GoTo NextPage  ' khong co background -> giu trong

                ' Buoc 3: load anh qua cache (tuong tu LoadBackground)
                Set img = GetOrLoadImage(imgPath)
                If img Is Nothing Then GoTo NextPage

                ' Buoc 4: gan vao pg.Picture
                On Error Resume Next
                Set pg.Picture = img
                pg.PictureSizeMode = 1   ' fmPictureSizeModeStretch - tuy chinh theo y muon
                pg.PictureAlignment = 2  ' fmPictureAlignmentCenter
                On Error GoTo 0

NextPage:
            Next pg
        End If
    Next ctrl
End Sub

'--------------------------------------------------------------
' LoadControlImages — FIX: luôn clear tru?c, r?i load t? theme folder
' Convention: Icons\{ctrl.Name}.png|jpg|bmp|gif
' Không c?n resources.ini — fallback trong ResolveImagePath t? x? lý
'--------------------------------------------------------------
Public Sub LoadControlImages(ByVal Controls As Object)
    Dim ctrl    As Object
    Dim tName   As String

    For Each ctrl In Controls
        On Error Resume Next
        tName = TypeName(ctrl)
        On Error GoTo 0

        Select Case tName
            Case "Frame"
                LoadControlImages ctrl.Controls

            Case "MultiPage"
                Dim pg As Object
                For Each pg In ctrl.Pages
                    LoadControlImages pg.Controls
                Next pg

            Case "Image", "CommandButton", "ToggleButton", "CheckBox", "OptionButton"
                If ctrl.Name <> "imgBackground" Then
                    ' NANG CAP: logic clear/tim/gan anh cho 1 control duoc
                    ' tach ra RefreshControlImage (public, xem duoi) de cac
                    ' module khac (vd modImageUtils) co the goi lai sau khi
                    ' doi/reset/xoa 1 icon rieng le, khong can ApplyTheme
                    ' lai toan bo form.
                    RefreshControlImage ctrl, "Icons"
                End If
        End Select
NextCtrl:
    Next ctrl
End Sub

'--------------------------------------------------------------
' RefreshControlImage — nap (hoac xoa trong) anh theme cho 1 CONTROL
' DUY NHAT, khong can quet lai toan bo form. Convention giong
' LoadControlImages: Themes\{CurrentTheme}\{category}\{ctrl.Name}.*
'
' Dung khi mot module khac (vd modImageUtils) vua Replace/Restore/
' Clear resource anh cua 1 control qua modResourceManager va can ve
' lai ngay control do, khong can goi ApplyTheme cho ca form.
'--------------------------------------------------------------
Public Sub RefreshControlImage(ByVal ctrl As Object, Optional ByVal category As String = "Icons")
    On Error Resume Next
    Set ctrl.Picture = Nothing
    On Error GoTo 0

    Dim imgPath As String
    imgPath = ResolveImagePath(ctrl.Name, category)
    If Len(imgPath) = 0 Then Exit Sub   ' Không có ?nh ? gi? tr?ng (dúng ý)

    Dim img As Object
    Set img = GetOrLoadImage(imgPath)
    If img Is Nothing Then Exit Sub

    On Error Resume Next
    Set ctrl.Picture = img
    On Error GoTo 0
End Sub


Public Function GetThemeColor(ByVal colorKey As String) As Long
    Dim Colors As Object

    Set Colors = GetCachedColors(m_CurrentTheme)
    If Not Colors Is Nothing Then
        If Colors.Exists(LCase$(colorKey)) Then
            GetThemeColor = CLng(Colors(LCase$(colorKey)))
            Exit Function
        End If
    End If

    If LCase$(m_CurrentTheme) <> LCase$(GetDefaultThemeName()) Then
        Set Colors = GetCachedColors(GetDefaultThemeName())
        If Not Colors Is Nothing Then
            If Colors.Exists(LCase$(colorKey)) Then
                GetThemeColor = CLng(Colors(LCase$(colorKey)))
                Exit Function
            End If
        End If
    End If

    GetThemeColor = FallbackColor(colorKey)
End Function

Public Function GetThemeInfo(ByVal themeName As String, _
                             ByVal key As String) As String
    If Len(themeName) = 0 Then themeName = m_CurrentTheme
    Dim info As Object
    Set info = GetCachedThemeInfo(themeName)
    If info Is Nothing Then
        InternalLoadThemeInfo themeName
        Set info = GetCachedThemeInfo(themeName)
    End If
    If Not info Is Nothing Then
        If info.Exists(LCase$(key)) Then GetThemeInfo = CStr(info(LCase$(key)))
    End If
End Function

Public Function GetCurrentTheme() As String
    Dim v As String
    v = INIRead(GetSettingsFilePath(), SETT_SECTION, SETT_KEY)
    GetCurrentTheme = IIf(Len(Trim$(v)) > 0, Trim$(v), GetDefaultThemeName())
End Function

Public Sub SetCurrentTheme(ByVal themeName As String)
    EnsureFolder GetCPMSAppFolder()
    INIWrite GetSettingsFilePath(), SETT_SECTION, SETT_KEY, themeName
    m_CurrentTheme = themeName
    ClearThemeCache themeName
    LoadTheme themeName
    ReloadAllOpenForms
End Sub
Public Sub SetCurrentThemeold(ByVal themeName As String, Optional ByVal excludeFrm As Object = Nothing)
    EnsureFolder GetCPMSAppFolder()
    INIWrite GetSettingsFilePath(), SETT_SECTION, SETT_KEY, themeName
    m_CurrentTheme = themeName
    ClearThemeCache themeName
    LoadTheme themeName
    ReloadAllOpenForms excludeFrm
End Sub

Public Function InstallTheme(ByVal cpmsThemePath As String) As String
    InstallTheme = InstallThemeFromFile(cpmsThemePath)
    If Len(InstallTheme) > 0 Then
        ClearThemeCache InstallTheme
        MsgBox "Theme """ & InstallTheme & """ installed!", vbInformation, "Theme Engine"
    End If
End Function

Public Function ReadThemeInfo(ByVal themeName As String) As Object
    Dim themeIni As String
    themeIni = GetThemeFolder(themeName) & "theme.ini"
    If ThemeFileExists(themeIni) Then
        Set ReadThemeInfo = INIReadSection(themeIni, "Theme")
    Else
        Set ReadThemeInfo = CreateObject("Scripting.Dictionary")
    End If
End Function


'===========================================================
' PRIVATE — Internal loaders (CU — GI? NGUYÊN)
'===========================================================

Private Sub InternalLoadThemeInfo(ByVal themeName As String)
    SetCachedThemeInfo themeName, ReadThemeInfo(themeName)
End Sub

Private Sub InternalLoadColors(ByVal themeName As String)
    Dim colorsIni As String
    colorsIni = GetThemeFolder(themeName) & "colors.ini"

    Dim colorDict As Object
    Set colorDict = CreateObject("Scripting.Dictionary")
    colorDict.CompareMode = 1

    If Not ThemeFileExists(colorsIni) Then
        If LCase$(themeName) <> LCase$(DEFAULT_THEME) Then
            Dim defColors As Object
            Set defColors = GetCachedColors(DEFAULT_THEME)
            If defColors Is Nothing Then
                InternalLoadColors DEFAULT_THEME
                Set defColors = GetCachedColors(DEFAULT_THEME)
            End If
            If Not defColors Is Nothing Then Set colorDict = defColors
        End If
        SetCachedColors themeName, colorDict
        Exit Sub
    End If

    Dim raw As Object
    Set raw = INIReadSection(colorsIni, "Colors")
    Dim k As Variant
    For Each k In raw.keys
        colorDict.Add LCase$(CStr(k)), HexToColor(CStr(raw(k)))
    Next k

    SetCachedColors themeName, colorDict
End Sub


Private Sub ApplyColorsToControls(ByVal Controls As Object)
    Dim ctrl  As Object
    Dim tName As String

    For Each ctrl In Controls
        On Error Resume Next
        tName = TypeName(ctrl)

        Select Case tName
            Case "Label"
                ctrl.BackStyle = 0
                ctrl.ForeColor = GetThemeColor("Text")

            Case "TextBox"
                ctrl.BackColor = GetThemeColor("TextBox.BackColor")
                ctrl.ForeColor = GetThemeColor("TextBox.ForeColor")
                ctrl.BorderColor = GetThemeColor("TextBox.BorderColor")

            Case "ListBox"
                ctrl.BackColor = GetThemeColor("ListBox.BackColor")
                ctrl.ForeColor = GetThemeColor("ListBox.ForeColor")
                ctrl.BorderColor = GetThemeColor("ListBox.BorderColor")

            Case "ComboBox"
                ctrl.BackColor = GetThemeColor("ComboBox.BackColor")
                ctrl.ForeColor = GetThemeColor("ComboBox.ForeColor")
                ctrl.BorderColor = GetThemeColor("ComboBox.BorderColor")

            Case "Frame"
                ctrl.BackColor = GetThemeColor("Frame.BackColor")
                ctrl.BorderColor = GetThemeColor("Frame.BorderColor")
                ApplyColorsToControls ctrl.Controls

            Case "MultiPage"
                ctrl.BackColor = GetThemeColor("Panel")
                Dim pg As Object
                For Each pg In ctrl.Pages
                    ApplyColorsToControls pg.Controls
                Next pg

            Case "CommandButton"
                ctrl.BackColor = GetThemeColor("CommandButton.BackColor")
                ctrl.ForeColor = GetThemeColor("CommandButton.ForeColor")
                ' HoverColor du?c dùng ? MouseMove event c?a form ch?a nút (xem README)

            Case "CheckBox", "OptionButton"
                ctrl.BackStyle = 0
                ctrl.ForeColor = GetThemeColor("Text")
        End Select
        On Error GoTo 0
    Next ctrl
End Sub

Private Function FallbackColor(ByVal colorKey As String) As Long
    Select Case LCase$(colorKey)

        ' --- Global palette ---
        Case "background":              FallbackColor = RGB(255, 255, 255)  ' #FFFFFF
        Case "panel":                   FallbackColor = RGB(245, 245, 245)  ' #F5F5F5
        Case "text":                    FallbackColor = RGB(30, 30, 30)     ' #1E1E1E
        Case "textlight":               FallbackColor = RGB(120, 120, 120)  ' #787878
        Case "primary":                 FallbackColor = RGB(244, 115, 45)   ' #F4732D - CPMS orange
        Case "secondary":               FallbackColor = RGB(31, 56, 100)    ' #1F3864 - CPMS navy
        Case "accent":                  FallbackColor = RGB(244, 115, 45)   ' #F4732D
        Case "border":                  FallbackColor = RGB(200, 200, 200)  ' #C8C8C8
        Case "success":                 FallbackColor = RGB(22, 163, 74)    ' #16A34A
        Case "warning":                 FallbackColor = RGB(234, 179, 8)    ' #EAB308
        Case "danger":                  FallbackColor = RGB(220, 38, 38)    ' #DC2626

        ' --- ListBox ---
        Case "listbox.bordercolor":     FallbackColor = RGB(200, 200, 200)  ' #C8C8C8
        Case "listbox.backcolor":       FallbackColor = RGB(255, 255, 255)  ' #FFFFFF
        Case "listbox.forecolor":       FallbackColor = RGB(30, 30, 30)     ' #1E1E1E

        ' --- ComboBox ---
        Case "combobox.bordercolor":    FallbackColor = RGB(200, 200, 200)  ' #C8C8C8
        Case "combobox.backcolor":      FallbackColor = RGB(255, 255, 255)  ' #FFFFFF
        Case "combobox.forecolor":      FallbackColor = RGB(30, 30, 30)     ' #1E1E1E

        ' --- TextBox ---
        Case "textbox.bordercolor":     FallbackColor = RGB(200, 200, 200)  ' #C8C8C8
        Case "textbox.backcolor":       FallbackColor = RGB(255, 255, 255)  ' #FFFFFF
        Case "textbox.forecolor":       FallbackColor = RGB(30, 30, 30)     ' #1E1E1E

        ' --- Frame ---
        Case "frame.bordercolor":       FallbackColor = RGB(200, 200, 200)  ' #C8C8C8
        Case "frame.backcolor":         FallbackColor = RGB(245, 245, 245)  ' #F5F5F5

        ' --- CommandButton ---
        Case "commandbutton.backcolor":  FallbackColor = RGB(246, 246, 246) ' #F6F6F6
        Case "commandbutton.forecolor":  FallbackColor = RGB(0, 0, 0)
        Case "commandbutton.hovercolor": FallbackColor = RGB(210, 95, 30)   ' #D25F1E

        ' --- Label ---
        Case "label.forecolor":         FallbackColor = RGB(31, 56, 100)    ' #1F3864

        Case Else:                      FallbackColor = RGB(0, 0, 0)

    End Select
End Function

Public Function ResolveImagePath(ByVal resName As String, ByVal category As String) As String
    ' NANG CAP: resource co the bi "ngat lien ket hien thi" (Shift+Right-
    ' Click / modResourceManager.ClearResourceImage) ma KHONG bi xoa RelFile
    ' hay file anh — chi bi danh dau Hidden. Kiem tra truoc tien, tra ve ""
    ' (trong) ngay ca khi RelFile van con hop le, de ResolveImagePath phan
    ' anh dung y "an di, khong xoa" cua thao tac nay.
    If IsResourceHidden(m_CurrentTheme, category, resName) Then Exit Function

    ' MOI (gop resources.ini + ComponentManifest.ini): RelFile gio duoc doc
    ' tu ComponentManifest.ini CUA THEME (qua modResourceManager), khong con
    ' qua resources.ini/GetCachedResources nua.
    Dim relFile As String: relFile = GetResourceRelFile(m_CurrentTheme, resName, category)
    If Len(relFile) > 0 Then
        Dim relPath As String: relPath = GetThemeFolder(m_CurrentTheme) & relFile
        If ThemeFileExists(relPath) Then
            ResolveImagePath = relPath
            Exit Function
        End If
    End If

    ' Fallback: do thang theo quy uoc thu muc category + cac ext anh
    ' (dung khi chua co entry manifest nao, hoac RelFile da ghi khong con ton tai)
    Dim exts() As String: exts = Split(IMG_EXTS, "|")
    Dim i As Long, p As String
    For i = 0 To UBound(exts)
        p = GetThemeFolder(m_CurrentTheme) & category & "\" & resName & exts(i)
        If ThemeFileExists(p) Then ResolveImagePath = p: Exit Function
    Next i
End Function

Private Function GetOrLoadImage(ByVal imgPath As String) As Object
    Dim img As Object
    Set img = GetCachedImage(imgPath)
    If img Is Nothing Then
        Set img = LoadImageFromPath(imgPath)
        If Not img Is Nothing Then SetCachedImage imgPath, img
    End If
    Set GetOrLoadImage = img
End Function



'===========================================================
' PH?N THEME MANAGER — CH?C NANG M?I
' (T?t c? d?u tái s? d?ng modINI/modFile/modCache/modZip)
'===========================================================

' -- Danh sách ??Y Ð? các color key mà Theme Manager cho phép
'    ch?nh (tab Colors). Dùng d? build UI d?ng lu?i (name | hex box)
'    và d? SaveThemeColors ghi dúng th? t? mong mu?n.
Public Function GetColorKeyList() As Variant
    GetColorKeyList = Array( _
        "Primary", "Secondary", "Accent", "Background", "Panel", _
        "Border", "Text", "TextLight", "Success", "Warning", "Danger", _
        "ListBox.BorderColor", "ListBox.BackColor", "ListBox.ForeColor", _
        "ComboBox.BorderColor", "ComboBox.BackColor", "ComboBox.ForeColor", _
        "TextBox.BorderColor", "TextBox.BackColor", "TextBox.ForeColor", _
        "Frame.BorderColor", "Frame.BackColor", _
        "CommandButton.BackColor", "CommandButton.ForeColor", "CommandButton.HoverColor", _
        "Label.ForeColor")
End Function

' Giá tr? m?c d?nh dùng khi t?o Theme m?i t? d?u (không clone)
Public Function GetDefaultColorMap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d.Add "Primary", "#F4732D":            d.Add "Secondary", "#1F3864"
    d.Add "Accent", "#F4732D":             d.Add "Background", "#FFFFFF"
    d.Add "Panel", "#F5F5F5":              d.Add "Border", "#C8C8C8"
    d.Add "Text", "#1E1E1E":               d.Add "TextLight", "#787878"
    d.Add "Success", "#16A34A":            d.Add "Warning", "#EAB308"
    d.Add "Danger", "#DC2626"
    d.Add "ListBox.BorderColor", "#C8C8C8":   d.Add "ListBox.BackColor", "#FFFFFF":   d.Add "ListBox.ForeColor", "#1E1E1E"
    d.Add "ComboBox.BorderColor", "#C8C8C8":  d.Add "ComboBox.BackColor", "#FFFFFF":  d.Add "ComboBox.ForeColor", "#1E1E1E"
    d.Add "TextBox.BorderColor", "#C8C8C8":   d.Add "TextBox.BackColor", "#FFFFFF":   d.Add "TextBox.ForeColor", "#1E1E1E"
    d.Add "Frame.BorderColor", "#C8C8C8":     d.Add "Frame.BackColor", "#F5F5F5"
    d.Add "CommandButton.BackColor", "#F6F6F6": d.Add "CommandButton.ForeColor", "#000000": d.Add "CommandButton.HoverColor", "#D25F1E"
    d.Add "Label.ForeColor", "#1F3864"
    Set GetDefaultColorMap = d
End Function

' -- Ð?c toàn b? colors.ini c?a 1 theme ra Dict {key(hoa/thu?ng g?c) ? hex string}
'    Dùng d? bind lên tab Colors (khác GetCachedColors vì cache dùng key
'    thu?ng + giá tr? Long, còn hàm này gi? nguyên hex string d? user edit)
Public Function LoadThemeColorsForEdit(ByVal themeName As String) As Object
    Dim colorsIni As String: colorsIni = GetThemeFolder(themeName) & "colors.ini"
    Dim raw As Object: Set raw = INIReadSection(colorsIni, "Colors")

    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    Dim keys As Variant: keys = GetColorKeyList()
    Dim i As Long
    For i = 0 To UBound(keys)
        Dim kName As String: kName = CStr(keys(i))
        If raw.Exists(kName) Then
            result.Add kName, CStr(raw(kName))
        Else
            result.Add kName, CStr(GetDefaultColorMap()(kName))
        End If
    Next i
    Set LoadThemeColorsForEdit = result
End Function

' -- Ghi toàn b? colors.ini t? Dict {key ? hex string} + reload cache
Public Sub SaveThemeColors(ByVal themeName As String, ByVal colorMap As Object)
    Dim colorsIni As String: colorsIni = GetThemeFolder(themeName) & "colors.ini"
    INIWriteSection colorsIni, "Colors", colorMap, True
    ClearThemeCache themeName
    If LCase$(themeName) = LCase$(m_CurrentTheme) Then LoadTheme themeName
End Sub

' -- Ghi toàn b? theme.ini t? Dict {key ? value} + reload cache
' Dict nên có: Name, Version, Author, Description, Preview,
'              IsDefault (0/1), MinimumEngineVersion
Public Sub SaveThemeInfo(ByVal themeName As String, ByVal infoMap As Object)
    Dim themeIni As String: themeIni = GetThemeFolder(themeName) & "theme.ini"
    INIWriteSection themeIni, "Theme", infoMap, True
    ClearThemeCache themeName
    If LCase$(themeName) = LCase$(m_CurrentTheme) Then LoadTheme themeName
End Sub

' -- T?o Theme M?I t? m?t Theme ngu?n (Create New Theme)
'    Copy toàn b? c?u trúc + d?i Name/Version/Author/Description.
'    Tr? v? tên Theme m?i, "" n?u th?t b?i.
Public Function CreateNewThemeFrom(ByVal sourceThemeName As String, _
                                   ByVal newThemeName As String, _
                                   ByVal newVersion As String, _
                                   ByVal newAuthor As String, _
                                   ByVal newDescription As String) As String
    newThemeName = SafeName(Trim$(newThemeName))
    If Len(newThemeName) = 0 Then Exit Function

    Dim destFolder As String: destFolder = GetThemeFolder(newThemeName)
    If ThemeFolderExists(destFolder) Then
        MsgBox "Theme """ & newThemeName & """ already exists.", vbExclamation
        Exit Function
    End If

    Dim srcFolder As String: srcFolder = GetThemeFolder(sourceThemeName)
    If Not ThemeFolderExists(srcFolder) Then
        MsgBox "Source theme does not exist: " & sourceThemeName, vbExclamation
        Exit Function
    End If

    EnsureThemeSkeleton destFolder
    CopyFolderTree srcFolder, destFolder

    Dim info As Object: Set info = CreateObject("Scripting.Dictionary")
    info.CompareMode = 1
    info.Add "Name", newThemeName
    info.Add "Version", IIf(Len(newVersion) = 0, "1.0", newVersion)
    info.Add "Author", newAuthor
    info.Add "Description", newDescription
    info.Add "IsDefault", "0"
    info.Add "MinimumEngineVersion", ENGINE_VER
    SaveThemeInfo newThemeName, info

    CreateNewThemeFrom = newThemeName
End Function

' -- Clone Theme: gi? nguyên 100% metadata, ch? d?i tên thu m?c
'    + Name bên trong theme.ini
Public Function CloneTheme(ByVal sourceThemeName As String, _
                           ByVal newThemeName As String) As String
    newThemeName = SafeName(Trim$(newThemeName))
    If Len(newThemeName) = 0 Then Exit Function

    Dim destFolder As String: destFolder = GetThemeFolder(newThemeName)
    If ThemeFolderExists(destFolder) Then
        MsgBox "Theme """ & newThemeName & """ already exists.", vbExclamation
        Exit Function
    End If

    Dim srcFolder As String: srcFolder = GetThemeFolder(sourceThemeName)
    If Not ThemeFolderExists(srcFolder) Then Exit Function

    EnsureFolder destFolder
    CopyFolderTree srcFolder, destFolder

    INIWrite destFolder & "theme.ini", "Theme", "Name", newThemeName
    INIWrite destFolder & "theme.ini", "Theme", "IsDefault", "0"

    CloneTheme = newThemeName
End Function

' -- Xóa Theme (delegate modZip.RemoveTheme, ch?n Default & Current)
'Public Function DeleteTheme(ByVal themeName As String) As Boolean
'    If LCase$(themeName) = LCase$(DEFAULT_THEME) Then
'        MsgBox "Không th? xóa Theme Default.", vbExclamation
'        Exit Function
'    End If
'    If LCase$(themeName) = LCase$(GetCurrentTheme()) Then
'        MsgBox "Không th? xóa Theme dang du?c s? d?ng. Hãy Apply Theme khác truoc.", vbExclamation
 '       Exit Function
'    End If
'    DeleteTheme = RemoveTheme(themeName)
'    If DeleteTheme Then ClearThemeCache themeName
'End Function

' -- Th?ng kê nhanh cho khung Preview (s? icon/background/ribbon,
'    dung lu?ng, ngày t?o/s?a)
Public Function GetThemeStats(ByVal themeName As String) As Object
    Dim folder As String: folder = GetThemeFolder(themeName)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d.Add "IconCount", CountFilesByExt(folder & "Icons", IMG_EXTS)
    d.Add "BackgroundCount", CountFilesByExt(folder & "Backgrounds", IMG_EXTS)

    ' FIX: uu tiên size file .cpmstheme (compressed) n?u dã t?ng Package
    ' Fallback v? folder size (uncompressed) n?u chua package l?n nào
    Dim sizeBytes As Long
    Dim sizeLabel As String
    Dim pkgPath As String
    pkgPath = INIRead(folder & "theme.ini", "Theme", "LastPackagePath")

    If Len(pkgPath) > 0 And ThemeFileExists(pkgPath) Then
        ' Dùng size file .cpmstheme — dây là con s? user th?y trong Explorer
        sizeBytes = FileLen(pkgPath)
        sizeLabel = Format$(sizeBytes / 1024, "#,##0.000") & " KB"
    Else
        ' Chua package ho?c file dã b? xóa ? hi?n folder size, ghi rõ "Installed"
        sizeBytes = GetFolderSizeBytes(folder)
        sizeLabel = Format$(sizeBytes / 1024, "#,##0.000") & " KB (installed)"
    End If

    d.Add "SizeBytes", sizeBytes
    d.Add "SizeText", sizeLabel
    d.Add "CreatedDate", GetPathCreatedDate(folder)
    d.Add "ModifiedDate", GetPathModifiedDate(folder)
    Set GetThemeStats = d
End Function

' -- Validation truoc khi Package (m?c 14 trong Upgrade Plan)
'    Tr? v? "" n?u OK, ho?c chu?i li?t kê các l?i/warning (m?i dòng 1 l?i)
Public Function ValidateThemeForPackage(ByVal themeName As String) As String
    Dim folder As String: folder = GetThemeFolder(themeName)
    Dim issues As String

    If Not ThemeFileExists(folder & "theme.ini") Then issues = issues & "- Missing theme.ini" & vbCrLf
    If Not ThemeFileExists(folder & "colors.ini") Then issues = issues & "- Missing colors.ini" & vbCrLf
    If Not ThemeFileExists(manifestPath(folder)) Then issues = issues & "- Missing ComponentManifest.ini (will be created automatically when scanning resources)" & vbCrLf

    Dim previewRel As String: previewRel = GetThemeInfo(themeName, "Preview")
    If Len(previewRel) = 0 Then previewRel = "Preview\Preview.jpg"
    If Not ThemeFileExists(folder & Replace(previewRel, "/", "\")) Then
        issues = issues & "- Missing Preview image (" & previewRel & ")" & vbCrLf
    End If

    If CountFilesByExt(folder & "Icons", IMG_EXTS) = 0 Then
        issues = issues & "- Icons\ folder is empty" & vbCrLf
    End If
    If CountFilesByExt(folder & "Backgrounds", IMG_EXTS) = 0 Then
        issues = issues & "- Backgrounds\ folder is empty (warning, not required)" & vbCrLf
    End If

    ValidateThemeForPackage = issues
End Function

Public Function GetEngineVersion() As String
    GetEngineVersion = ENGINE_VER
End Function
' Dam bao IsDefault trong theme.ini cua TAT CA theme luon khop voi con tro
' Default trong Settings.ini (Settings.ini la nguon du lieu goc/authoritative).
' Goi ham nay o RefreshThemeList de moi lan list theme deu tu sua lech neu co
' (theme.ini bi sua tay, theme cai tu package mang san IsDefault=1 tu may khac, v.v.)
Public Sub SyncDefaultThemeFlag()
    Dim defName As String: defName = GetDefaultThemeName()
    Dim themes As Variant: themes = ListInstalledThemes()
    If Not IsArray(themes) Then Exit Sub

    Dim i As Long
    For i = 0 To UBound(themes)
        Dim tName As String: tName = CStr(themes(i))
        Dim iniPath As String: iniPath = GetThemeFolder(tName) & "theme.ini"
        If ThemeFileExists(iniPath) Then
            Dim curFlag As String: curFlag = Trim$(INIRead(iniPath, "Theme", "IsDefault"))
            Dim shouldBe As String: shouldBe = IIf(LCase$(tName) = LCase$(defName), "1", "0")
            If curFlag <> shouldBe Then
                INIWrite iniPath, "Theme", "IsDefault", shouldBe
                ClearThemeCache tName
            End If
        End If
    Next i
End Sub
Public Function GetDefaultThemeName() As String
    Dim v As String: v = INIRead(GetSettingsFilePath(), SETT_SECTION, SETT_DEFAULT_KEY)
    GetDefaultThemeName = IIf(Len(Trim$(v)) > 0, Trim$(v), DEFAULT_THEME)
End Function
Private Sub SetDefaultThemeNameInternal(ByVal themeName As String)
    EnsureFolder GetCPMSAppFolder()
    INIWrite GetSettingsFilePath(), SETT_SECTION, SETT_DEFAULT_KEY, Trim$(themeName)
End Sub

Public Function SetDefaultTheme(ByVal themeName As String) As Boolean
    themeName = Trim$(themeName)
    If Len(themeName) = 0 Then Exit Function
    If Not ThemeFolderExists(GetThemeFolder(themeName)) Then Exit Function
    Dim oldDefault As String: oldDefault = GetDefaultThemeName()
    If LCase$(oldDefault) <> LCase$(themeName) And ThemeFolderExists(GetThemeFolder(oldDefault)) Then
        INIWrite GetThemeFolder(oldDefault) & "theme.ini", "Theme", "IsDefault", "0"
        ClearThemeCache oldDefault
    End If
    SetDefaultThemeNameInternal themeName
    INIWrite GetThemeFolder(themeName) & "theme.ini", "Theme", "IsDefault", "1"
    ClearThemeCache themeName
    SetDefaultTheme = True
End Function
Public Sub RenameDefaultThemePointer(ByVal newThemeName As String)
    SetDefaultThemeNameInternal newThemeName
End Sub


