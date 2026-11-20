Attribute VB_Name = "frmThemeManager"
Attribute VB_Base = "0{95DAC8C4-7824-4E3C-A8EE-348B13001B4B}{18CB5418-5E6A-4BBF-A111-FA6974CBEE46}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'===========================================================
' frmThemeManager — Theme Manager (nang cap tu Icon Manager cu)
'
' === Controls phai them/giu trong Form Designer ===
' Xem chi tiet day du trong README_ThemeManager_Controls.txt
' (dinh kem cung bo file nay). Tom tat nhanh:
'
' mpMain (MultiPage) nay co 5 tab:
'   [0] Themes        — giu nguyen toan bo control cu
'   [1] Icons         — giu nguyen toan bo control cu
'   [2] Colors        — MOI: lstColorKeys, txtColorHex, imgColorSwatch,
'                        btnPickColor, btnSaveColors, btnResetColors
'   [3] Theme Info    — MOI: txtInfoName, txtInfoVersion, txtInfoAuthor,
'                        txtInfoDesc, chkIsDefault, txtInfoMinEngine,
'                        btnBrowsePreview, imgInfoPreview, btnSaveInfo
'   [4] Resources     — MOI: lstResources (3 cot: Name|Category|Status),
'                        imgResPreview, btnResReplace, btnResDelete,
'                        btnResRestoreDefault, btnResRename
'                        MOI (Resource Size): lblResSize — Label don, chi
'                        hien thi "Size: W x H pt" cua resource dang chon.
'                        cboResFilter: combo loc theo All / Icon / Background
'                        chkHasImage: CheckBox chi hien thi resource co anh
'                        KHONG them cot vao lstResources — chi 1 Label moi
'                        tren page Resources, dat gan imgResPreview (vi du
'                        ngay duoi anh preview hoac canh lblResTheme).
'
' Ngoai MultiPage (thanh cong cu chung, luon hien):
'   btnLoadDefaultResources CommandButton — Load Default Resources (muc 1)
'   btnLoadExternalResources CommandButton — MOI: Load External Resources
'                             (quet 1 file VBA KHAC, xuat resources ra
'                             ngoai dang goi .zip, khong dong bo vao theme
'                             cua project hien tai)
'   btnCreateNewTheme       CommandButton — Create New Theme (muc 5)
'   btnCloneTheme           CommandButton — Clone Theme (muc 6)
'   btnPackageTheme         CommandButton — Package Theme (muc 8)
'   btnDeleteTheme          CommandButton — Delete Theme (muc 10)
'   btnReloadCache          CommandButton — Reload Cache (muc 12)
'   btnClearCache           CommandButton — Clear Cache (muc 12)
'   btnRebuildCache         CommandButton — Rebuild Cache (muc 12)
'   lblStats                Label         — thong ke Theme (muc 13)
'   btnClose                CommandButton — dong form
'===========================================================

Private m_IconTheme As String
' Theme dang duoc chinh sua o cac tab Colors/Theme Info/Resources.
' Mac dinh = theme dang chon ben tab Themes.
Private m_EditTheme  As String
Private m_CurrentTheme As String
Private mOriginalPictures As Object ' Scripting.Dictionary: ten control -> Picture goc

' FIX: co bao ve vong lap tai nhap cho RefreshResourcesTab
' (vi cboResFilter_Change va chkHasImage_Click cung goi lai RefreshResourcesTab)
Private m_RefreshingResources As Boolean




Private Sub txtInfoVersion_Change()

End Sub

'===========================================================
' PHAN 1 — Khoi tao & Tab switching
'===========================================================
' Trong MOI UserForm can ho tro live reload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    UnregisterForm Me   ' tu dong xoa khoi registry khi dong
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "CPMS Theme Manager"
    
    BackupOriginalPictures
    ApplyTheme Me
    
    On Error Resume Next
    imgPreview.BorderStyle = 1
    imgPreview.BackColor = RGB(240, 240, 240)
    On Error GoTo 0

    On Error Resume Next
    lblCurrentTag.ForeColor = RGB(22, 163, 74)
    lblCurrentTag.FontBold = True
    lblCurrentTag.Visible = False
    On Error GoTo 0

    ' MOI: hien thi trang thai theme dang active ngay khi mo form, kem
    ' luon ten Default theme de nguoi dung biet theme nao la Default.
    On Error Resume Next
    lblApplyStatus.Caption = "Current theme: " & GetCurrentTheme() & "   |   Default theme: " & GetDefaultThemeName()
    lblApplyStatus.ForeColor = RGB(22, 163, 74)
    lblApplyStatus.Visible = True
    On Error GoTo 0

    On Error Resume Next
    ' lstThemes: cot 0 = thu tu, cot 1 = ten theme
    lstThemes.ColumnCount = 2
    lstThemes.ColumnWidths = "20pt;180pt"
    On Error GoTo 0

    On Error Resume Next
    lstColorKeys.ColumnCount = 2
    lstColorKeys.ColumnWidths = "150pt;70pt"
    On Error GoTo 0

    On Error Resume Next
    lstResources.ColumnCount = 3
    lstResources.ColumnWidths = "125pt;55pt;40pt"
    imgResPreview.BorderStyle = 1
    imgResPreview.PictureSizeMode = 3
    On Error GoTo 0

    ' MOI (Resource Size): khoi tao label size — an cho den khi user chon
    ' 1 resource trong lstResources.
    On Error Resume Next
    lblResSize.Caption = ""
    lblResSize.Visible = False
    On Error GoTo 0

    On Error Resume Next
    imgInfoPreview.BorderStyle = 1
    imgInfoPreview.PictureSizeMode = 3
    On Error GoTo 0

    ' FIX: ApplyTheme truoc — tranh ApplyColorsToControls reset ListIndex
    ' cua lstThemes ve -1 sau khi da chon


    ' Sau khi ApplyTheme xong moi populate + chon theme
    mpMain.value = 0
    RefreshThemeList
    If lstThemes.ListIndex >= 0 Then lstThemes_Click
End Sub
' Thêm ? ph?n khai báo d?u module (General Declarations)


Private Sub BackupOriginalPictures()
      Dim ctrl As MSForms.Control
      Dim pic As IPictureDisp

      Set mOriginalPictures = CreateObject("Scripting.Dictionary")
      mOriginalPictures.CompareMode = 1 ' vbTextCompare

      ' FIX 1: backup luon Picture cua chinh UserForm (truong hop dung Me.Picture
      ' lam Background thay vi 1 Image control imgBackground rieng). Key dung
      ' = Me.Name, khop voi resName="formName" ma modResourceScanner.ScanSelfForm
      ' dung khi doc lai qua GetOriginalPicture.
      Set pic = Nothing
      On Error Resume Next
      Set pic = Me.Picture
      On Error GoTo 0
      If Not pic Is Nothing Then
          If pic.Type <> 0 Then
              Set mOriginalPictures(Me.Name) = pic
          End If
      End If

      ' Me.Controls la danh sach PHANG, da bao gom ca control nam trong
      ' Frame/Page cua MultiPage nen 1 vong lap la du cho toan bo control.
      For Each ctrl In Me.Controls
          Set pic = Nothing
          On Error Resume Next
          Set pic = ctrl.Picture
          On Error GoTo 0

          If Not pic Is Nothing Then
              If pic.Type <> 0 Then ' 0 = placeholder rong, bo qua
                  Set mOriginalPictures(ctrl.Name) = pic
              End If
          End If

          ' FIX 2: rieng Page ben trong MultiPage co Picture (background) cua
          ' rieng no nhung Page KHONG nam trong Me.Controls (chi cac control DAT
          ' TREN page moi nam trong Controls) nen phai duyet rieng va backup
          ' voi dung KEY "<FormName>_<PageName>" -- khop voi pgResName ma
          ' modResourceScanner.ScanControlsForImages dung.
          If TypeName(ctrl) = "MultiPage" Then
              Dim pg As Object
              For Each pg In ctrl.Pages
                  Set pic = Nothing
                  On Error Resume Next
                  Set pic = pg.Picture
                  On Error GoTo 0
                  If Not pic Is Nothing Then
                      If pic.Type <> 0 Then
                          Set mOriginalPictures(Me.Name & "_" & pg.Name) = pic
                      End If
                  End If
              Next pg
          End If
      Next ctrl
    End Sub
    
' Ham public de modResourceScanner doc lai Picture goc
Public Function GetOriginalPicture(ctrlName As String) As IPictureDisp
    If Not mOriginalPictures Is Nothing Then
        If mOriginalPictures.Exists(ctrlName) Then
            Set GetOriginalPicture = mOriginalPictures(ctrlName)
            Exit Function
        End If
    End If
    Set GetOriginalPicture = Nothing
End Function
Private Sub lstThemes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstThemes.ListIndex < 0 Then Exit Sub

    ' Nho ten theme TRUOC khi Apply, vi SetCurrentTheme -> ReloadAllOpenForms
    ' se ap lai theme len chinh form nay (frmThemeManager da tu dang ky
    ' qua ApplyTheme Me trong UserForm_Initialize), khien ApplyColorsToControls
    ' set BackColor/ForeColor len lstThemes -> reset ListIndex ve -1.
    Dim themeName As String
    themeName = lstThemes.List(lstThemes.ListIndex, 1)

    btnApply_Click

    ' FIX: tim lai theo TEN thay vi tin ListIndex cu, vi ListIndex co the
    ' da bi reset trong luc btnApply_Click chay (do ReloadAllOpenForms
    ' ap lai theme len chinh form nay).
    Dim i As Long
    For i = 0 To lstThemes.ListCount - 1
        If LCase$(lstThemes.List(i, 1)) = LCase$(themeName) Then
            lstThemes.ListIndex = i
            Exit For
        End If
    Next i

    If lstThemes.ListIndex >= 0 Then lstThemes_Click
End Sub

Private Sub btnApply_Click()
    If lstThemes.ListIndex < 0 Then
        lblApplyStatus.Caption = "Please select a theme first."
        lblApplyStatus.ForeColor = RGB(220, 38, 38)
        lblApplyStatus.Visible = True
        Exit Sub
    End If

    Dim themeName As String
    themeName = lstThemes.List(lstThemes.ListIndex, 1)

    ' FIX: bo tinh nang disable btnApply (trong xau khi grey-out), chan
    ' bang MsgBox ngay khi bam nut thay vi khong cho bam.
    If LCase$(themeName) = LCase$(GetCurrentTheme()) Then
        MsgBox "Theme """ & themeName & """ is already active.", vbInformation
        Exit Sub
    End If

    SetCurrentTheme themeName

    lblCurrentTag.Visible = True

    ' FIX: bo MsgBox — MsgBox che len UserForm modal ngay sau khi vua
    ' reload Picture/mau hang loat control khien Windows khong tra lai
    ' dung focus/activation cho form ben duoi (phai click ra ngoai roi
    ' click lai moi thao tac duoc). Thay bang label khong-modal.
    lblApplyStatus.Caption = "Theme """ & themeName & """ is now active.   |   Default theme: " & GetDefaultThemeName()
    lblApplyStatus.ForeColor = RGB(22, 163, 74)
    lblApplyStatus.Visible = True
End Sub

Private Sub lstThemes_Click()
    If lstThemes.ListIndex < 0 Then
        ClearPreviewPanel
        Exit Sub
    End If

    Dim themeName As String
    themeName = lstThemes.List(lstThemes.ListIndex, 1)

    ShowThemeDetails themeName
    m_EditTheme = themeName
    UpdateStatsLabel

    Select Case mpMain.value
        Case 0: RefreshResourcesTab
        Case 1: RefreshColorsTab
        Case 2: RefreshInfoTab
    End Select
End Sub
Private Sub mpMain_Change()
    Select Case mpMain.value
        Case 0: RefreshResourcesTab
        Case 1: RefreshColorsTab
        Case 2: RefreshInfoTab
    End Select
End Sub

'===========================================================
' PHAN 2 — Tab THEMES (logic cu, giu nguyen)
'===========================================================

Public Sub RefreshThemeList()

    SyncDefaultThemeFlag
    lstThemes.Clear
   
    Dim themes As Variant
    themes = ListInstalledThemes()

    If Not IsArray(themes) Then Exit Sub
    If UBound(themes) < 0 Then
        ClearPreviewPanel
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To UBound(themes)
        lstThemes.AddItem CStr(i + 1)
        lstThemes.List(lstThemes.ListCount - 1, 1) = CStr(themes(i))
    Next i

    Dim current As String: current = GetCurrentTheme()
    For i = 0 To lstThemes.ListCount - 1
        If LCase$(lstThemes.List(i, 1)) = LCase$(current) Then
            lstThemes.ListIndex = i
            Exit For
        End If
    Next i

    If lstThemes.ListIndex < 0 And lstThemes.ListCount > 0 Then
        lstThemes.ListIndex = 0
    End If

    UpdateStatsLabel
End Sub



Private Sub ShowThemeDetails(ByVal themeName As String)
    Dim tName    As String: tName = GetThemeInfo(themeName, "Name")
    Dim tVersion As String: tVersion = GetThemeInfo(themeName, "Version")
    Dim tAuthor  As String: tAuthor = GetThemeInfo(themeName, "Author")
    Dim tDesc    As String: tDesc = GetThemeInfo(themeName, "Description")

    If Len(tName) = 0 Then tName = themeName

    lblVersion.Caption = "Version " & tVersion & "   |   By: " & tAuthor
    lblDesc.Caption = tDesc

    lblCurrentTag.Visible = (LCase$(themeName) = LCase$(GetCurrentTheme()))

    Dim previewRel As String
    previewRel = GetThemeInfo(themeName, "Preview")
    If Len(previewRel) = 0 Then previewRel = "Preview\Preview.jpg"

    Dim previewPath As String
    previewPath = GetThemeFolder(themeName) & Replace(previewRel, "/", "\")

    On Error Resume Next
    Set imgPreview.Picture = Nothing
    If ThemeFileExists(previewPath) Then
        Set imgPreview.Picture = LoadPicture(previewPath)
        imgPreview.PictureSizeMode = 3
    End If
    On Error GoTo 0

    'btnRemove.Enabled = (LCase$(themeName) <> "default")
End Sub

Private Sub ClearPreviewPanel()
    lblVersion.Caption = ""
    lblDesc.Caption = ""
    lblCurrentTag.Visible = False
    'btnRemove.Enabled = False
    On Error Resume Next
    Set imgPreview.Picture = Nothing
    On Error GoTo 0
End Sub

Private Sub btnInstall_Click()
    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
        "CPMS Theme (*.cpmstheme),*.cpmstheme", , "Install Theme Package")
    If VarType(filePath) = vbBoolean Then Exit Sub

    Dim installed As String
    installed = InstallThemeFromFileEx(CStr(filePath), "ask")
    If Len(installed) > 0 Then
        RefreshThemeList
        Dim i As Long
        For i = 0 To lstThemes.ListCount - 1
            If LCase$(lstThemes.List(i, 1)) = LCase$(installed) Then
                lstThemes.ListIndex = i: Exit For
            End If
        Next i
    End If
End Sub

Private Sub btnRemove_Click()
    If lstThemes.ListIndex < 0 Then Exit Sub

    Dim themeName As String
    themeName = lstThemes.List(lstThemes.ListIndex, 1)

    If LCase$(themeName) = LCase$(GetDefaultThemeName()) Then
        MsgBox "Theme """ & themeName & """ is currently set as Default and cannot be removed." & vbCrLf & _
               "Set another theme as Default first.", vbExclamation
        Exit Sub
    End If

    If LCase$(themeName) = LCase$(GetCurrentTheme()) Then
        MsgBox "Theme """ & themeName & """ is currently in use (Current theme) and cannot be removed." & vbCrLf & _
               "Apply a different theme first.", vbExclamation
        Exit Sub
    End If

    If RemoveTheme(themeName) Then
        ClearAllCache
        RefreshThemeList
    End If
End Sub

Private Sub btnOpenFolder_Click()
    Dim folderPath As String: folderPath = GetThemesFolder()
    EnsureFolder folderPath
    Shell "explorer.exe """ & folderPath & """", vbNormalFocus
End Sub

'===========================================================
' PHAN 4 — Tab COLORS (MOI, muc 2 Upgrade Plan)
'===========================================================

Public Sub RefreshColorsTab()
    If Len(m_EditTheme) = 0 Then m_EditTheme = GetCurrentTheme()

    Dim colorMap As Object: Set colorMap = LoadThemeColorsForEdit(m_EditTheme)
    Dim keys As Variant: keys = GetColorKeyList()

    On Error Resume Next
    lstColorKeys.Clear
    On Error GoTo 0

    Dim i As Long
    For i = 0 To UBound(keys)
        Dim kName As String: kName = CStr(keys(i))
        lstColorKeys.AddItem kName
        lstColorKeys.List(lstColorKeys.ListCount - 1, 1) = CStr(colorMap(kName))
    Next i

    On Error Resume Next
    lblColorTheme.Caption = "Theme: " & m_EditTheme
    On Error GoTo 0

    ' MOI: mac dinh chon luon dong dau tien trong lstColorKeys khi load lai
    ' page Colors, thay vi de trong (khong chon gi) nhu truoc.
    If lstColorKeys.ListCount > 0 Then
        lstColorKeys.ListIndex = 0
        lstColorKeys_Click
    End If
End Sub

Private Sub lstColorKeys_Click()
    If lstColorKeys.ListIndex < 0 Then Exit Sub
    Dim hexVal As String: hexVal = lstColorKeys.List(lstColorKeys.ListIndex, 1)

    On Error Resume Next
    txtColorHex.value = hexVal
    imgColorSwatch.BackColor = HexToColor(hexVal)
    On Error GoTo 0
End Sub

' Nguoi dung go truc tiep ma hex roi Enter/LostFocus -> cap nhat list + swatch
Private Sub txtColorHex_AfterUpdate()
    If lstColorKeys.ListIndex < 0 Then Exit Sub
    Dim hexVal As String: hexVal = Trim$(txtColorHex.value)
    If Left$(hexVal, 1) <> "#" Or Len(hexVal) <> 7 Then
        MsgBox "Color code must be in #RRGGBB format.", vbExclamation
        Exit Sub
    End If
    lstColorKeys.List(lstColorKeys.ListIndex, 1) = hexVal
    On Error Resume Next
    imgColorSwatch.BackColor = HexToColor(hexVal)
    On Error GoTo 0
End Sub

' Dung hop thoai chon mau chuan cua Office (Application.Dialogs)
Private Sub btnPickColor_Click()
    If lstColorKeys.ListIndex < 0 Then
        MsgBox "Please select a color from the list first", vbExclamation
        Exit Sub
    End If

    Dim currentHex As String
    currentHex = lstColorKeys.List(lstColorKeys.ListIndex, 1)

    Dim picked As Long: picked = PickColorDialog(HexToColor(currentHex))
    If picked = -1 Then Exit Sub

    Dim hexVal As String: hexVal = ColorToHex(picked)
    lstColorKeys.List(lstColorKeys.ListIndex, 1) = hexVal

    On Error Resume Next
    txtColorHex.value = hexVal
    imgColorSwatch.BackColor = picked
    On Error GoTo 0
End Sub

Private Sub btnSaveColors_Click()
    Dim colorMap As Object: Set colorMap = CreateObject("Scripting.Dictionary")
    colorMap.CompareMode = 1

    Dim i As Long
    For i = 0 To lstColorKeys.ListCount - 1
        colorMap.Add lstColorKeys.List(i, 0), lstColorKeys.List(i, 1)
    Next i

    SaveThemeColors m_EditTheme, colorMap
    MsgBox "Saved colors.ini for theme """ & m_EditTheme & """.", vbInformation
End Sub

Private Sub btnResetColors_Click()
    Dim ans As Long
    ans = MsgBox("Reset all colors to default?", vbYesNo + vbQuestion)
    If ans = vbNo Then Exit Sub

    Dim defMap As Object: Set defMap = GetDefaultColorMap()
    Dim i As Long
    For i = 0 To lstColorKeys.ListCount - 1
        lstColorKeys.List(i, 1) = CStr(defMap(lstColorKeys.List(i, 0)))
    Next i
End Sub

'===========================================================
' PHAN 5 — Tab THEME INFO (MOI, muc 3 Upgrade Plan)
'===========================================================

Public Sub RefreshInfoTab()
    If Len(m_EditTheme) = 0 Then m_EditTheme = GetCurrentTheme()

    txtInfoName.value = GetThemeInfo(m_EditTheme, "Name")
    txtInfoVersion.value = GetThemeInfo(m_EditTheme, "Version")
    txtInfoAuthor.value = GetThemeInfo(m_EditTheme, "Author")
    txtInfoDesc.value = GetThemeInfo(m_EditTheme, "Description")
    txtInfoMinEngine.value = GetThemeInfo(m_EditTheme, "MinimumEngineVersion")

    On Error Resume Next
    chkIsDefault.value = (GetThemeInfo(m_EditTheme, "IsDefault") = "1")
    On Error GoTo 0

    Dim previewRel As String: previewRel = GetThemeInfo(m_EditTheme, "Preview")
    If Len(previewRel) = 0 Then previewRel = "Preview\Preview.jpg"
    Dim previewPath As String: previewPath = GetThemeFolder(m_EditTheme) & Replace(previewRel, "/", "\")

    On Error Resume Next
    Set imgInfoPreview.Picture = Nothing
    If ThemeFileExists(previewPath) Then Set imgInfoPreview.Picture = LoadPicture(previewPath)
    On Error GoTo 0
End Sub

Private Sub btnBrowsePreview_Click()
    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename( _
        "Image Files (*.png;*.jpg;*.bmp;*.gif),*.png;*.jpg;*.bmp;*.gif", , "Select Preview Image")
    If VarType(srcPath) = vbBoolean Then Exit Sub

    If ReplaceResourceImage(m_EditTheme, "Preview", "Preview", CStr(srcPath)) Then
        INIWrite GetThemeFolder(m_EditTheme) & "theme.ini", "Theme", "Preview", _
                 "Preview\Preview" & "." & GetFileExt(CStr(srcPath))
        RefreshInfoTab
    End If
End Sub

Private Sub btnSaveInfo_Click()
    If Len(Trim$(txtInfoName.value)) = 0 Then
        MsgBox "Theme Name cannot be empty.", vbExclamation
        Exit Sub
    End If

    Dim oldName As String: oldName = m_EditTheme
    Dim newName As String: newName = SafeName(Trim$(txtInfoName.value))

    ' --- 1) Doi ten thu muc neu Name thay doi (duoc phep ca khi do la
    '        theme Default — chi cam khi do la Current theme) ---
    If LCase$(newName) <> LCase$(oldName) Then
        If LCase$(oldName) = LCase$(GetCurrentTheme()) Then
            MsgBox "Theme """ & oldName & """ is currently in use (Current theme)." & vbCrLf & _
                   "Please apply a different theme before renaming.", vbExclamation, "Cannot Rename"
            txtInfoName.value = oldName
            Exit Sub
        End If

        Dim destFolder As String: destFolder = GetThemeFolder(newName)
        If ThemeFolderExists(destFolder) Then
            MsgBox "A theme named """ & newName & """ already exists.", vbExclamation
            Exit Sub
        End If

        Dim wasDefault As Boolean
        wasDefault = (LCase$(oldName) = LCase$(GetDefaultThemeName()))

        Dim srcFolder As String: srcFolder = GetThemeFolder(oldName)
        On Error GoTo EH_Rename
        Name Left$(srcFolder, Len(srcFolder) - 1) As Left$(destFolder, Len(destFolder) - 1)
        On Error GoTo 0

        ClearThemeCache oldName
        m_EditTheme = newName

        ' Theme vua doi ten chinh la theme Default -> cap nhat lai con tro
        ' trong Settings.ini de fallback khong bi tro sai thu muc
        If wasDefault Then RenameDefaultThemePointer newName
    End If

    ' --- 2) Ghi theme.ini ---
    Dim info As Object: Set info = CreateObject("Scripting.Dictionary")
    info.CompareMode = 1
    info.Add "Name", Trim$(txtInfoName.value)
    info.Add "Version", Trim$(txtInfoVersion.value)
    info.Add "Author", Trim$(txtInfoAuthor.value)
    info.Add "Description", Trim$(txtInfoDesc.value)
    info.Add "IsDefault", IIf(chkIsDefault.value, "1", "0")
    info.Add "MinimumEngineVersion", IIf(Len(Trim$(txtInfoMinEngine.value)) = 0, _
                                          GetEngineVersion(), Trim$(txtInfoMinEngine.value))

    Dim previewRel As String: previewRel = GetThemeInfo(m_EditTheme, "Preview")
    If Len(previewRel) > 0 Then info.Add "Preview", previewRel

    SaveThemeInfo m_EditTheme, info

    ' --- 3) Neu tich "Set Default" -> chuyen con tro Default sang theme nay ---
    If chkIsDefault.value Then
        If LCase$(m_EditTheme) <> LCase$(GetDefaultThemeName()) Then
            SetDefaultTheme m_EditTheme
        End If
    End If
    On Error Resume Next
    lblApplyStatus.Caption = "Current theme: " & GetCurrentTheme() & "   |   Default theme: " & GetDefaultThemeName()
    lblApplyStatus.ForeColor = RGB(22, 163, 74)
    lblApplyStatus.Visible = True
    On Error GoTo 0
    RefreshThemeList
    If lstThemes.ListIndex >= 0 Then lstThemes_Click
    MsgBox "Theme saved successfully.", vbInformation
    Exit Sub

EH_Rename:
    MsgBox "Cannot rename theme folder: " & Err.Description, vbCritical
End Sub

'===========================================================
' PHAN 6 — Tab RESOURCES (MOI, muc 4 Upgrade Plan)
'
' CAP NHAT (Resource Size): lstResources_Click va ClearResourcePreview
' gio cung cap nhat lblResSize — Label MOI can them trong Form Designer
' tren page Resources (xem huong dan o dau file). Label nay hien thi
' "Size: <W> x <H> pt" cua component dang duoc chon trong lstResources,
' lay tu key "SizeText" do modResourceManager.ListThemeResources tra ve
' (chinh la du lieu do modResourceScanner ghi vao ComponentManifest.ini
' khi chay Load Default Resources). KHONG them cot size vao lstResources.
'
' CONTROLS MOI CAN THEM TRONG FORM DESIGNER (page Resources):
'   cboResFilter  — ComboBox: loc theo All / Icon / Background
'   chkHasImage   — CheckBox: "Has Image" — chi hien resource co anh
'===========================================================

' ---------------------------------------------------------------
' RefreshResourcesTab
'
' selectResName / selectCategory (optional):
'   Khi truyen vao, sau khi load lai danh sach se tu dong chon lai
'   dong co ResName + Category tuong ung — dung cho btnResReplace va
'   btnResRestoreDefault de giu nguyen vi tri chon sau khi thay the anh.
'   Neu khong truyen (mac dinh = ""), chon dong dau tien nhu cu.
' ---------------------------------------------------------------
Public Sub RefreshResourcesTab(Optional ByVal selectResName As String = "", _
                                Optional ByVal selectCategory As String = "")
    ' FIX: bao ve vong lap tai nhap (cboResFilter_Change / chkHasImage_Click
    ' deu goi lai sub nay, neu sub dang chay ma event len thi bo qua)
    If m_RefreshingResources Then Exit Sub
    m_RefreshingResources = True

    On Error GoTo Cleanup

    If Len(m_EditTheme) = 0 Then m_EditTheme = GetCurrentTheme()

    On Error Resume Next
    lstResources.Clear
    On Error GoTo 0

    ' -----------------------------------------------------------------
    ' FIX: Khoi tao / dong bo cboResFilter voi option moi:
    '      "All" | "Icon" | "Background"
    ' Neu ListCount khac 3 hoac item thu 2 khong phai "Icon" (cu la
    ' "Has Image") -> xoa va tao lai, giu nguyen ListIndex cu neu hop le.
    ' -----------------------------------------------------------------
    On Error Resume Next
    Dim needRebuildCombo As Boolean
    needRebuildCombo = (cboResFilter.ListCount <> 3)
    If Not needRebuildCombo Then
        needRebuildCombo = (cboResFilter.List(0) <> "All") Or _
                           (cboResFilter.List(1) <> "Icon") Or _
                           (cboResFilter.List(2) <> "Background")
    End If
    If needRebuildCombo Then
        Dim savedComboIdx As Long: savedComboIdx = cboResFilter.ListIndex
        cboResFilter.Clear
        cboResFilter.AddItem "All"
        cboResFilter.AddItem "Icon"
        cboResFilter.AddItem "Background"
        ' Giu index cu neu con hop le, nguoc lai ve 0
        If savedComboIdx >= 0 And savedComboIdx <= 2 Then
            cboResFilter.ListIndex = savedComboIdx
        Else
            cboResFilter.ListIndex = 0
        End If
    End If
    On Error GoTo 0

    ' Doc gia tri loc hien tai
    Dim filterMode As String
    On Error Resume Next
    filterMode = cboResFilter.List(cboResFilter.ListIndex)
    On Error GoTo 0
    If Len(filterMode) = 0 Then filterMode = "All"

    ' FIX: Doc trang thai checkbox chkHasImage (chi hien resource co anh)
    Dim filterHasImage As Boolean
    On Error Resume Next
    filterHasImage = (chkHasImage.value = True)
    On Error GoTo 0

    Dim items As Collection: Set items = ListThemeResources(m_EditTheme)
    Dim it As Variant
    Dim totalCount As Long: totalCount = items.count
    Dim missingCount As Long: missingCount = 0

    ' -----------------------------------------------------------------
    ' FIX: Thu thap tung muc vao mang tam de sap xep truoc khi dua vao list
    ' -----------------------------------------------------------------
    Dim arrResName() As String
    Dim arrCategory() As String
    Dim arrStatus() As String
    Dim arrCount As Long: arrCount = 0

    If totalCount > 0 Then
        ReDim arrResName(totalCount - 1)
        ReDim arrCategory(totalCount - 1)
        ReDim arrStatus(totalCount - 1)
    End If

    For Each it In items
        If Not it("Exists") Then missingCount = missingCount + 1

        ' --- Filter theo Category (All / Icon / Background) ---
        Dim passFilter As Boolean
        Select Case filterMode
            Case "Icon"
                ' So sanh khong phan biet hoa/thuong, Category co chua "icon"
                passFilter = (InStr(1, CStr(it("Category")), "icon", vbTextCompare) > 0)
            Case "Background"
                passFilter = (InStr(1, CStr(it("Category")), "background", vbTextCompare) > 0)
            Case Else ' "All"
                passFilter = True
        End Select

        ' --- FIX: Filter theo checkbox HasImage (chi hien resource co anh) ---
        If passFilter And filterHasImage Then
            passFilter = it("Exists")
        End If

        If passFilter Then
            arrResName(arrCount) = CStr(it("ResName"))
            arrCategory(arrCount) = CStr(it("Category"))
            arrStatus(arrCount) = IIf(it("Exists"), "OK", "Missing")
            arrCount = arrCount + 1
        End If
    Next it

    ' -----------------------------------------------------------------
    ' FIX: Sap xep mang theo ResName theo thu tu bang chu cai (A -> Z)
    ' Dung Bubble Sort — du dung cho so luong resources thong thuong
    ' -----------------------------------------------------------------
    If arrCount > 1 Then
        Dim i As Long, j As Long
        Dim tmpName As String, tmpCat As String, tmpStat As String
        For i = 0 To arrCount - 2
            For j = i + 1 To arrCount - 1
                If LCase$(arrResName(i)) > LCase$(arrResName(j)) Then
                    ' Hoan vi ResName
                    tmpName = arrResName(i)
                    arrResName(i) = arrResName(j)
                    arrResName(j) = tmpName
                    ' Hoan vi Category
                    tmpCat = arrCategory(i)
                    arrCategory(i) = arrCategory(j)
                    arrCategory(j) = tmpCat
                    ' Hoan vi Status
                    tmpStat = arrStatus(i)
                    arrStatus(i) = arrStatus(j)
                    arrStatus(j) = tmpStat
                End If
            Next j
        Next i
    End If

    ' -----------------------------------------------------------------
    ' Dua du lieu da sap xep vao lstResources, dong thoi tim index can chon lai
    ' -----------------------------------------------------------------
    Dim targetIdx As Long: targetIdx = -1
    Dim k As Long
    For k = 0 To arrCount - 1
        lstResources.AddItem arrResName(k)
        lstResources.List(lstResources.ListCount - 1, 1) = arrCategory(k)
        lstResources.List(lstResources.ListCount - 1, 2) = arrStatus(k)

        ' FIX: Neu truyen vao ten resource can chon lai, tim vi tri chinh xac
        If Len(selectResName) > 0 Then
            If LCase$(arrResName(k)) = LCase$(selectResName) And _
               LCase$(arrCategory(k)) = LCase$(selectCategory) Then
                targetIdx = k
            End If
        End If
    Next k

    On Error Resume Next
    lblResTheme.Caption = "Theme: " & m_EditTheme & "  (" & arrCount & "/" & totalCount & _
                           " shown, " & missingCount & " missing)"
    On Error GoTo 0

    ' -----------------------------------------------------------------
    ' FIX: Chon lai resource vua duoc Replace/Restore;
    '      neu khong co yeu cau cu the -> chon dong dau tien nhu cu.
    ' -----------------------------------------------------------------
    If lstResources.ListCount > 0 Then
        If targetIdx >= 0 Then
            lstResources.ListIndex = targetIdx
        Else
            lstResources.ListIndex = 0
        End If
        lstResources_Click
    Else
        ClearResourcePreview
    End If

Cleanup:
    m_RefreshingResources = False
End Sub

' ------------------------------------------------------------------
' cboResFilter_Change: nguoi dung thay doi category filter
' ------------------------------------------------------------------
Private Sub cboResFilter_Change()
    RefreshResourcesTab
End Sub

' ------------------------------------------------------------------
' chkHasImage_Click: nguoi dung bat/tat filter "Has Image"
' ------------------------------------------------------------------
Private Sub chkHasImage_Click()
    RefreshResourcesTab
End Sub

Private Sub lstResources_Click()
    If lstResources.ListIndex < 0 Then
        ClearResourcePreview
        Exit Sub
    End If

    Dim resName  As String: resName = lstResources.List(lstResources.ListIndex, 0)
    Dim category As String: category = lstResources.List(lstResources.ListIndex, 1)

    Dim items As Collection: Set items = ListThemeResources(m_EditTheme)
    Dim it As Variant
    Dim sizeText As String: sizeText = "N/A"

    For Each it In items
        If it("ResName") = resName And it("Category") = category Then
            On Error Resume Next
            Set imgResPreview.Picture = Nothing
            If it("Exists") Then Set imgResPreview.Picture = LoadPicture(it("FullPath"))
            On Error GoTo 0
            sizeText = it("SizeText")   ' MOI: size cua component dang chon
            Exit For
        End If
    Next it

    ' MOI (Resource Size): hien thi size len lblResSize tren page Resources
    On Error Resume Next
    lblResSize.Caption = "Size: " & sizeText
    lblResSize.Visible = True
    On Error GoTo 0

End Sub

Private Sub ClearResourcePreview()
    On Error Resume Next
    Set imgResPreview.Picture = Nothing
    ' MOI (Resource Size): an label size khi khong co resource nao duoc chon
    lblResSize.Caption = ""
    lblResSize.Visible = False
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------
' btnResReplace_Click
' FIX: Truyen resName + category vao RefreshResourcesTab de sau khi
'      replace, lstResources tu dong chon lai dung dong vua replace
'      (khong bi nhay ve dau danh sach nhu truoc).
' ------------------------------------------------------------------
Private Sub btnResReplace_Click()
    If lstResources.ListIndex < 0 Then
        MsgBox "Please select a resource from the list first.", vbExclamation
        Exit Sub
    End If
    Dim resName  As String: resName = lstResources.List(lstResources.ListIndex, 0)
    Dim category As String: category = lstResources.List(lstResources.ListIndex, 1)

    Dim srcPath As Variant
    srcPath = Application.GetOpenFilename( _
        "Image Files (*.png;*.jpg;*.bmp;*.gif),*.png;*.jpg;*.bmp;*.gif", , "Select Replacement Image")
    If VarType(srcPath) = vbBoolean Then Exit Sub

    If ReplaceResourceImage(m_EditTheme, category, resName, CStr(srcPath)) Then
        ' FIX: truyen resName + category de RefreshResourcesTab chon lai
        '      dung dong vua duoc replace thay vi nhay ve dau danh sach.
        RefreshResourcesTab resName, category
    End If
End Sub


Private Sub btnResRestoreDefault_Click()
    If lstResources.ListIndex < 0 Then
        MsgBox "Please select a resource from the list first.", vbExclamation
        Exit Sub
    End If

    ' FIX: bo tinh nang disable btnResRestoreDefault khi dang sua Default
    ' theme (trong xau khi grey-out), chan bang MsgBox ngay khi bam nut.
    If LCase$(m_EditTheme) = LCase$(GetDefaultThemeName()) Then
        MsgBox "This is already the Default theme; there is nothing to restore.", vbExclamation
        Exit Sub
    End If

    Dim resName  As String: resName = lstResources.List(lstResources.ListIndex, 0)
    Dim category As String: category = lstResources.List(lstResources.ListIndex, 1)

    If RestoreDefaultResourceImage(m_EditTheme, category, resName) Then
        ' FIX: tuong tu btnResReplace — giu nguyen vi tri chon sau khi restore
        RefreshResourcesTab resName, category
    End If
End Sub

'===========================================================
' PHAN 7 — Thanh cong cu chung (MOI)
'===========================================================

Private Sub btnLoadDefaultResources_Click()
    Dim choice As Long
    choice = MsgBox("Choose an action:" & vbCrLf & vbCrLf & _
                 "YES    = Load Default Resources for THIS project" & vbCrLf & _
                 "           (scan all UserForms here and fill in the Default Theme)." & vbCrLf & _
                 "NO     = Load External Resources from ANOTHER file" & vbCrLf & _
                 "           (scan a different .xlsm/.xlam/.xls and export a .zip package)." & vbCrLf & _
                 "CANCEL = Do nothing.", _
                 vbYesNoCancel + vbQuestion, "Load Resources")

    Select Case choice
        Case vbYes
            RunLoadDefaultResources
        Case vbNo
            RunLoadExternalResourcesToZip
        Case Else
            ' vbCancel - do nothing
    End Select
End Sub
Private Sub RunLoadDefaultResources()
    Dim ans As Long
    ans = MsgBox("Scan all UserForms in the VBA Project, record the names of every" & vbCrLf & _
                 "form/page/button (including components with no image yet), and copy" & vbCrLf & _
                 "any missing images into the Default Theme?" & vbCrLf & _
                 "The Width x Height of each component will also be recorded for" & vbCrLf & _
                 "display on the Resources page.", _
                 vbYesNo + vbQuestion, "Load Default Resources")
    If ans = vbNo Then Exit Sub

    Dim result As ScanResult
    result = LoadDefaultResourcesFromProject(Me)

    If Len(result.ErrorMessage) > 0 Then
        MsgBox result.ErrorMessage, vbCritical
        Exit Sub
    End If

    InitThemeEngine
    RefreshThemeList
    RefreshResourcesTab   ' MOI: cap nhat ngay lstResources + lblResSize voi component/size moi quet

    MsgBox "Scanned " & result.FormsScanned & " form(s)." & vbCrLf & _
           "Components discovered: " & result.ComponentsFound & vbCrLf & _
           "Sizes recorded: " & result.SizesRecorded & vbCrLf & _
           "Images found: " & result.ImagesFound & vbCrLf & _
           "Newly copied: " & result.ImagesCopied & vbCrLf & _
           "Already existed (skipped): " & result.ImagesSkipped, _
           vbInformation, "Load Default Resources"
End Sub
Private Sub RunLoadExternalResourcesToZip()
    Dim ans As Long
    ans = MsgBox("Select another Excel file (with a VBA project) to scan." & vbCrLf & _
                 "Every image/component found on its UserForms will be exported" & vbCrLf & _
                 "into a separate .zip package (Icons\ + Backgrounds\)." & vbCrLf & _
                 "This will NOT change any theme in the current project." & vbCrLf & vbCrLf & _
                 "Continue?", vbYesNo + vbQuestion, "Load External Resources")
    If ans = vbNo Then Exit Sub

    Dim srcFile As Variant
    srcFile = Application.GetOpenFilename( _
        "Excel Macro-Enabled Files (*.xlsm;*.xlam;*.xls),*.xlsm;*.xlam;*.xls", , _
        "Select the source file to scan")
    If VarType(srcFile) = vbBoolean Then Exit Sub

    If StrComp(CStr(srcFile), ThisWorkbook.fullName, vbTextCompare) = 0 Then
        MsgBox "Please choose a different file — this is the current project." & vbCrLf & _
               "Use ""Load Default Resources"" instead for this project.", vbExclamation
        Exit Sub
    End If

    Dim destZip As Variant
    destZip = Application.GetSaveAsFilename( _
        InitialFileName:=GetFileBaseName(CStr(srcFile)) & "_Resources", _
        FileFilter:="Zip Package (*.zip),*.zip")
    If VarType(destZip) = vbBoolean Then Exit Sub

    Dim result As ScanResult
    result = ScanExternalProjectToZip(CStr(srcFile), CStr(destZip))

    If Len(result.ErrorMessage) > 0 Then
        MsgBox result.ErrorMessage, vbCritical
        Exit Sub
    End If

    MsgBox "Scanned " & result.FormsScanned & " form(s) from the external file." & vbCrLf & _
           "Components discovered: " & result.ComponentsFound & vbCrLf & _
           "Images found: " & result.ImagesFound & vbCrLf & _
           "Images exported to the package: " & result.ImagesCopied & vbCrLf & vbCrLf & _
           "Package saved at:" & vbCrLf & CStr(destZip), _
           vbInformation, "Load External Resources"
End Sub

    
' Muc 6 — Clone Theme
Private Sub btnCloneTheme_Click()
    If lstThemes.ListIndex < 0 Then
        MsgBox "Please select a theme to clone.", vbExclamation
        Exit Sub
    End If
    Dim sourceTheme As String: sourceTheme = lstThemes.List(lstThemes.ListIndex, 1)

    Dim newName As String
    newName = InputBox("New theme name (cloned from """ & sourceTheme & """):", "Clone Theme")
    If Len(Trim$(newName)) = 0 Then Exit Sub

    Dim cloned As String: cloned = CloneTheme(sourceTheme, newName)
    If Len(cloned) > 0 Then
        RefreshThemeList
        MsgBox "Cloned as """ & cloned & """.", vbInformation
    End If
End Sub


' Muc 8 — Package Theme (kem Validation muc 14 truoc khi nen)
Private Sub btnPackageTheme_Click()
    If lstThemes.ListIndex < 0 Then
        MsgBox "Please select a theme to package.", vbExclamation
        Exit Sub
    End If
    Dim themeName As String: themeName = lstThemes.List(lstThemes.ListIndex, 1)

    Dim issues As String: issues = ValidateThemeForPackage(themeName)
    If Len(issues) > 0 Then
        Dim ans As Long
        ans = MsgBox("Theme has issues:" & vbCrLf & issues & vbCrLf & _
                     "Continue packaging anyway?", vbYesNo + vbExclamation, "Validation")
        If ans = vbNo Then Exit Sub
    End If

    Dim targetFolder As Variant
    targetFolder = Application.GetSaveAsFilename( _
        InitialFileName:=themeName, FileFilter:="CPMS Theme (*.cpmstheme),*.cpmstheme")
    If VarType(targetFolder) = vbBoolean Then Exit Sub

    ' Nguoi dung chon ten file dich, tach lai thanh folder + dung PackageTheme
    ' theo dung ten Theme (da yeu cau trong modZip.PackageTheme)
    Dim destFolder As String: destFolder = GetParentFolder(CStr(targetFolder)) & "\"
    Dim finalPath As String: finalPath = PackageTheme(themeName, destFolder)

    If Len(finalPath) > 0 Then
        MsgBox "Packaged:" & vbCrLf & finalPath, vbInformation, "Package Theme"
    End If
End Sub


'--------------------------------------------------------------
' btnRebuildCache — Refresh UI + Rebuild Cache (gop 3 nut cu)
' Xoa cache RAM -> doc lai tu disk -> cap nhat lstThemes + cac tab
'--------------------------------------------------------------
Private Sub btnRebuildCache_Click()
    ClearAllCache
    InitThemeEngine
    RefreshThemeList
    If lstThemes.ListIndex >= 0 Then
        RefreshResourcesTab
        RefreshColorsTab
        RefreshInfoTab
    End If
    ReloadAllOpenForms     ' -> reload tat ca form sau rebuild cache
    MsgBox "Cache cleared. Theme list refreshed.", vbInformation, "Refresh"
End Sub

' Muc 13 — Preview mo rong (so Icon/Background/Ribbon, dung luong, ngay)
Private Sub UpdateStatsLabel()
    On Error Resume Next
    If lstThemes.ListIndex < 0 Then
        lblStats.Caption = ""
        Exit Sub
    End If

    Dim themeName As String: themeName = lstThemes.List(lstThemes.ListIndex, 1)
    Dim stats As Object: Set stats = GetThemeStats(themeName)

    lblStats.Caption = _
        "Icons: " & stats("IconCount") & "    " & _
        "Backgrounds: " & stats("BackgroundCount") & "    " & _
        "Size: " & stats("SizeText") & "    " & _
        "Created: " & stats("CreatedDate") & "    " & _
        "Modified: " & stats("ModifiedDate")
        
    On Error GoTo 0
End Sub




