Attribute VB_Name = "modThemeResourceManager"
Option Explicit

'===========================================================
' modResourceManager — Resources Manager (muc 4 Upgrade Plan)
'
' Danh sach + thao tac tren tung Resource (Icon/Background/
' Ribbon/Preview) cua 1 Theme: Thay anh, Khoi phuc mac dinh
' (tu Default theme).
'
' CAP NHAT (gop resources.ini + ComponentManifest.ini): mapping
' resource-name -> RelFile (ten file anh tuong doi) KHONG con nam
' trong resources.ini nua. ComponentManifest.ini gio la file
' THEO TUNG THEME (nam ngay trong theme folder), luu ca RelFile va
' size cho moi component, va la NGUON DU LIEU DUY NHAT. Cache trong
' module nay (m_ManifestCache) key theo ten theme (khong con la 1
' cache GLOBAL duy nhat nhu ban truoc).
'
' Tuong thich nguoc: theme cu chi co resources.ini (chua tung duoc
' "nang cap") van hoat dong binh thuong — BuildManifestEntriesForTheme
' se TU DONG hap thu (absorb) cac entry con thieu tu resources.ini
' cua theme do vao ComponentManifest.ini ngay trong lan doc dau tien,
' giup theme "tot nghiep" ma khong can xoa hay migrate thu cong.
' resources.ini sau do tro thanh file "chet" (khong con duoc doc/ghi
' nua) — co the xoa thu cong nhung khong bat buoc.
'===========================================================

' Ghi chu ky thuat: VBA Collection KHONG the chua Public Type (UDT)
' — Collection.Add chi nhan Variant/Object. Vi vay moi "resource item"
' duoc dai dien bang 1 Scripting.Dictionary voi cac key: "ResName",
' "Category", "RelPath", "FullPath", "Exists", "SizeText" (thay cho
' Type ResourceItem).

' m_ManifestCache: themeNameLower -> Dictionary("category|key" -> Dictionary("RelFile","SizeText"))
Private m_ManifestCache As Object

'----------------------------------------------------------
' GetManifestEntries — Tra ve (va cache) toan bo entry manifest cua
' 1 theme, xay dung qua BuildManifestEntriesForTheme (bao gom logic
' hap thu resources.ini cu).
'----------------------------------------------------------
Public Function GetManifestEntries(ByVal themeName As String) As Object
    If m_ManifestCache Is Nothing Then
        Set m_ManifestCache = CreateObject("Scripting.Dictionary")
        m_ManifestCache.CompareMode = 1
    End If

    Dim k As String: k = LCase$(themeName)
    If Not m_ManifestCache.Exists(k) Then
        m_ManifestCache.Add k, BuildManifestEntriesForTheme(themeName)
    End If
    Set GetManifestEntries = m_ManifestCache(k)
End Function

'----------------------------------------------------------
' InvalidateManifestCache — Xoa cache manifest cua 1 theme cu the,
' hoac TOAN BO cache neu khong truyen themeName. Goi tu modCache
' (ClearThemeCache/ClearAllCache) va modResourceScanner (sau khi
' quet xong Load Default Resources).
'----------------------------------------------------------
Public Sub InvalidateManifestCache(Optional ByVal themeName As String = "")
    If m_ManifestCache Is Nothing Then Exit Sub
    If Len(themeName) = 0 Then
        Set m_ManifestCache = Nothing
    Else
        Dim k As String: k = LCase$(themeName)
        If m_ManifestCache.Exists(k) Then m_ManifestCache.Remove k
    End If
End Sub

'----------------------------------------------------------
' BuildManifestEntriesForTheme — Doc ComponentManifest.ini cua 1
' theme (qua ParseManifestEntry cua modResourceScanner), roi hap thu
' them cac entry con thieu tu resources.ini CU cua chinh theme do
' (neu file van con ton tai) — vua de tuong thich nguoc, vua GHI LAI
' ngay vao ComponentManifest.ini de lan doc sau khong can dung den
' resources.ini nua.
'----------------------------------------------------------
Private Function BuildManifestEntriesForTheme(ByVal themeName As String) As Object
    Dim entries As Object: Set entries = CreateObject("Scripting.Dictionary")
    entries.CompareMode = 1

    Dim themeFolder As String: themeFolder = GetThemeFolder(themeName)
    Dim mPath As String: mPath = manifestPath(themeFolder)

    Dim categories As Variant: categories = Array("Icons", "Backgrounds", "Preview")
    Dim c As Variant

    If ThemeFileExists(mPath) Then
        For Each c In categories
            Dim sec As Object: Set sec = INIReadSection(mPath, CStr(c))
            Dim k As Variant
            For Each k In sec.keys
                entries.Add LCase$(CStr(c)) & "|" & LCase$(CStr(k)), ParseManifestEntry(CStr(sec(k)))
            Next k
        Next c
    End If

    ' Tuong thich nguoc / tu-nang-cap: theme cu chi co resources.ini —
    ' hap thu RelFile cho cac key CON THIEU trong manifest, roi ghi
    ' lai ngay vao ComponentManifest.ini.
    Dim resIni As String: resIni = themeFolder & "resources.ini"
    If ThemeFileExists(resIni) Then
        For Each c In categories
            Dim rsec As Object: Set rsec = INIReadSection(resIni, CStr(c))
            Dim rk As Variant
            For Each rk In rsec.keys
                Dim ek As String: ek = LCase$(CStr(c)) & "|" & LCase$(CStr(rk))
                If Not entries.Exists(ek) Then
                    Dim relFile As String: relFile = Replace(CStr(rsec(rk)), "/", "\")
                    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
                    d.CompareMode = 1
                    d.Add "RelFile", relFile
                    d.Add "SizeText", "N/A"
                    entries.Add ek, d
                    ' "Tot nghiep": ghi luon vao ComponentManifest.ini de tu lan sau
                    ' khong con phai doc resources.ini nua.
                    INIWrite mPath, CStr(c), CStr(rk), "1|" & relFile
                End If
            Next rk
        Next c
    End If

    Set BuildManifestEntriesForTheme = entries
End Function

'----------------------------------------------------------
' GetResourceRelFile — Dung boi modTheme.ResolveImagePath de lay
' ten file anh (tuong doi) cua 1 resource, tra ve "" neu chua co.
'----------------------------------------------------------
Public Function GetResourceRelFile(ByVal themeName As String, ByVal resName As String, ByVal category As String) As String
    Dim entries As Object: Set entries = GetManifestEntries(themeName)
    Dim k As String: k = LCase$(category) & "|" & LCase$(resName)
    If entries.Exists(k) Then
        GetResourceRelFile = CStr(entries(k)("RelFile"))
    End If
End Function

'----------------------------------------------------------
' GetComponentSizeTextCached — Tra ve chuoi "W x H pt" (hoac "N/A")
' cho 1 component cua 1 theme cu the, dung cache theo-theme o tren.
'----------------------------------------------------------
Public Function GetComponentSizeTextCached(ByVal themeName As String, ByVal category As String, ByVal itemKey As String) As String
    Dim entries As Object: Set entries = GetManifestEntries(themeName)
    Dim k As String: k = LCase$(category) & "|" & LCase$(itemKey)
    If entries.Exists(k) Then
        GetComponentSizeTextCached = CStr(entries(k)("SizeText"))
    Else
        GetComponentSizeTextCached = "N/A"
    End If
End Function

Public Function ListThemeResources(ByVal themeName As String) As Collection
    Dim result As New Collection
    Dim themeFolder As String: themeFolder = GetThemeFolder(themeName)

    Dim categories As Variant
    categories = Array("Icons", "Backgrounds", "Preview")

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    Dim entries As Object: Set entries = GetManifestEntries(themeName)

    Dim c As Variant
    For Each c In categories
        ' Nguon DUY NHAT: ComponentManifest.ini cua theme (qua cache entries),
        ' hien ca nhung component CHUA co anh (RelFile rong).
        Dim prefix As String: prefix = LCase$(CStr(c)) & "|"
        Dim ek As Variant
        For Each ek In entries.keys
            If Left$(CStr(ek), Len(prefix)) = prefix Then
                Dim resName As String: resName = mID$(CStr(ek), Len(prefix) + 1)
                Dim relFile As String: relFile = CStr(entries(ek)("RelFile"))
                Dim fullPath As String
                Dim existsFlag As Boolean
                If Len(relFile) > 0 Then
                    fullPath = themeFolder & relFile
                    existsFlag = ThemeFileExists(fullPath)
                Else
                    relFile = CStr(c) & "\" & resName
                    fullPath = themeFolder & relFile
                    existsFlag = False
                End If
                result.Add MakeResourceDict(resName, CStr(c), relFile, fullPath, existsFlag, themeName)
                seen.Add LCase$(CStr(c)) & "|" & LCase$(resName), True
            End If
        Next ek

        ' Quet them file vat ly chua duoc dang ky trong manifest (anh "mo")
        If ThemeFolderExists(themeFolder & CStr(c)) Then
            Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
            Dim f As Object
            For Each f In fso.GetFolder(themeFolder & CStr(c)).Files
                Dim baseName As String: baseName = Left$(f.Name, Len(f.Name) - Len(GetFileExt(f.Name)) - 1)
                Dim keyCheck As String: keyCheck = LCase$(CStr(c)) & "|" & LCase$(baseName)
                If Not seen.Exists(keyCheck) Then
                    result.Add MakeResourceDict(baseName, CStr(c), CStr(c) & "\" & f.Name, f.path, True, themeName)
                    seen.Add keyCheck, True
                End If
            Next f
            Set fso = Nothing
        End If
    Next c

    Set ListThemeResources = result
End Function

' -- Private: tao 1 "resource item" dang Dictionary de nhoi vao Collection
Private Function MakeResourceDict(ByVal resName As String, ByVal category As String, _
                                  ByVal relPath As String, ByVal fullPath As String, _
                                  ByVal existsFlag As Boolean, ByVal themeName As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d.Add "ResName", resName
    d.Add "Category", category
    d.Add "RelPath", relPath
    d.Add "FullPath", fullPath
    d.Add "Exists", existsFlag
    ' Size thiet ke (Width x Height) cua component, doc tu ComponentManifest.ini
    ' cua theme nay (qua cache). Tra ve "N/A" neu component chua duoc quet
    ' hoac chua co size.
    d.Add "SizeText", GetComponentSizeTextCached(themeName, category, resName)
    Set MakeResourceDict = d
End Function

'----------------------------------------------------------
' ReplaceResourceImage — Thay anh cho 1 resource da co/chua co
' srcImagePath: file anh nguoi dung chon (bat ky dau tren may)
'
' NANG CAP (Ctrl+Click chi la thay TAM): truoc khi ghi de, neu day la
' lan DAU TIEN doi anh cho resource nay (chua co ban sao luu), anh
' HIEN TAI se duoc sao luu lai (xem BackupCurrentResourceImage) truoc
' khi bi ghi de. Anh goc cua theme vi vay KHONG bi mat — RestoreOriginal
' ResourceImage (Shift+Click) se dung ban sao luu nay de phuc hoi dung
' anh cua CHINH theme nay (khac RestoreDefaultResourceImage, von lay
' anh tu theme Default).
'----------------------------------------------------------
Public Function ReplaceResourceImage(ByVal themeName As String, _
                                     ByVal category As String, _
                                     ByVal resName As String, _
                                     ByVal srcImagePath As String) As Boolean
    If Not ThemeFileExists(srcImagePath) Then Exit Function

    Dim themeFolder As String: themeFolder = GetThemeFolder(themeName)
    Dim destFolder  As String: destFolder = themeFolder & category & "\"
    EnsureFolder destFolder

    ' Sao luu anh GOC cua theme (chi lam 1 lan duy nhat, lan doi dau tien)
    If Len(GetBackupRelFile(themeFolder, category, resName)) = 0 Then
        BackupCurrentResourceImage themeFolder, category, resName
    End If

    ' Xoa cac file cung ten nhung khac ext (tranh trung nhieu dinh dang)
    RemoveAllVariants destFolder, resName

    Dim ext As String: ext = "." & GetFileExt(srcImagePath)
    Dim destPath As String: destPath = destFolder & resName & ext

    On Error GoTo EH
    FileCopy srcImagePath, destPath

    ' FIX: modCache.GetCachedImage cache theo DUONG DAN (khong theo theme),
    ' nen neu destPath trung voi lan truoc (cung resName+ext) thi anh CU van
    ' con trong cache -> RefreshControlImage se ve lai anh cu du file da bi
    ' ghi de. Phai xoa cache anh cua duong dan nay NGAY SAU KHI ghi file moi.
    SetCachedImage destPath, Nothing

    ' Cap nhat RelFile trong ComponentManifest.ini cua theme (giu nguyen
    ' size da co, neu co)
    WriteManifestRelFile themeFolder, category, resName, category & "\" & resName & ext

    ' Anh moi luon duoc HIEN (bo co "an" neu truoc do da Clear/ngat lien ket)
    SetResourceHidden themeFolder, category, resName, False

    ClearThemeCache themeName
    If LCase$(GetCurrentTheme()) = LCase$(themeName) Then LoadTheme themeName

    ReplaceResourceImage = True
    Exit Function
EH:
    ReplaceResourceImage = False
End Function

'----------------------------------------------------------
' DeleteResourceImage — Xoa anh + entry (RelFile) khoi ComponentManifest.ini
' CHU Y (out of scope, phat hien nhung khong sua trong lan cap nhat nay):
' frmThemeManager.btnResDelete_Click dang GOI ham nay nhung toan bo than
' ham lai dang bi COMMENT (dead code) — nut Delete tren UI hien dang
' khong hoat dong. Day la 1 loi CO SAN tu truoc, khong lien quan den
' viec gop resources.ini/ComponentManifest.ini lan nay.
'----------------------------------------------------------
'Public Function DeleteResourceImage(ByVal themeName As String, _
'                                    ByVal category As String, _
'                                    ByVal resName As String) As Boolean
'    Dim themeFolder As String: themeFolder = GetThemeFolder(themeName)
'    RemoveAllVariants themeFolder & category & "\", resName
'    INIDeleteKey manifestPath(themeFolder), category, resName
'
'    ClearThemeCache themeName
'    If LCase$(GetCurrentTheme()) = LCase$(themeName) Then LoadTheme themeName
'    DeleteResourceImage = True
'End Function

'----------------------------------------------------------
' RestoreDefaultResourceImage — Khoi phuc anh tu theme Default
'----------------------------------------------------------
Public Function RestoreDefaultResourceImage(ByVal themeName As String, _
                                            ByVal category As String, _
                                            ByVal resName As String) As Boolean
    If LCase$(themeName) = LCase$(GetDefaultThemeName()) Then
        MsgBox "This is already the Default theme.", vbExclamation
        Exit Function
    End If

    Dim defFolder As String: defFolder = GetDefaultThemeFolder() & category & "\"
    Dim exts() As String: exts = Split(".png|.jpg|.jpeg|.bmp|.gif", "|")
    Dim i As Long, srcPath As String
    For i = 0 To UBound(exts)
        If ThemeFileExists(defFolder & resName & exts(i)) Then
            srcPath = defFolder & resName & exts(i)
            Exit For
        End If
    Next i

    If Len(srcPath) = 0 Then
        MsgBox "The Default theme has no default image for """ & resName & """.", vbExclamation
        Exit Function
    End If

    RestoreDefaultResourceImage = ReplaceResourceImage(themeName, category, resName, srcPath)
End Function

'----------------------------------------------------------
' ClearResourceImage — CHI ngat LIEN KET HIEN THI cua 1 resource TRONG
' THEME HIEN TAI (vd Shift+Right-Click tren Image1..6 cua frmLinkKL).
' KHONG xoa file anh, KHONG xoa RelFile da gan trong ComponentManifest.ini
' — anh (va ban sao luu, neu co, xem BackupCurrentResourceImage) van
' nguyen ven tren dia, chi bi danh dau "an" (Hidden) nen ResolveImagePath
' se tra "" (trong) cho control nay o theme nay. ReplaceResourceImage
' (Change) va RestoreOriginalResourceImage (Reset) deu tu dong go bo
' dau "an" nay khi duoc goi, vi luc do da co anh hop le de hien thi lai.
'----------------------------------------------------------
Public Function ClearResourceImage(ByVal themeName As String, _
                                    ByVal category As String, _
                                    ByVal resName As String) As Boolean
    Dim themeFolder As String: themeFolder = GetThemeFolder(themeName)
    SetResourceHidden themeFolder, category, resName, True

    ClearThemeCache themeName
    If LCase$(GetCurrentTheme()) = LCase$(themeName) Then LoadTheme themeName

    ClearResourceImage = True
End Function

'----------------------------------------------------------
' RestoreOriginalResourceImage — Khoi phuc anh GOC CUA CHINH THEME HIEN
' TAI (khac RestoreDefaultResourceImage, von lay anh tu theme Default).
' Dung ban sao luu duoc tao boi ReplaceResourceImage o lan Ctrl+Click
' dau tien (BackupCurrentResourceImage) de "hoan tac" thay doi tam thoi
' do. Neu resource nay CHUA TUNG bi doi (chua co ban sao luu) — nghia
' la RelFile hien tai DA LA anh goc — Reset chi can go dau "an" (neu
' truoc do bi Clear) de hien lai anh goc co san, khong can copy file
' nao ca. Chi tra False khi thuc su khong co gi de khoi phuc (chua doi
' va cung khong bi an).
'----------------------------------------------------------
Public Function RestoreOriginalResourceImage(ByVal themeName As String, _
                                             ByVal category As String, _
                                             ByVal resName As String) As Boolean
    Dim themeFolder As String: themeFolder = GetThemeFolder(themeName)

    Dim backupRel As String: backupRel = GetBackupRelFile(themeFolder, category, resName)
    If Len(backupRel) = 0 Then
        ' Chua tung Ctrl+Click doi anh -> RelFile hien tai chinh la anh
        ' goc. Neu dang bi "an" (Shift+Right-Click truoc do), go an de
        ' hien lai; neu khong bi an thi khong co gi de Reset.
        If IsResourceHidden(themeName, category, resName) Then
            SetResourceHidden themeFolder, category, resName, False
            ClearThemeCache themeName
            If LCase$(GetCurrentTheme()) = LCase$(themeName) Then LoadTheme themeName
            RestoreOriginalResourceImage = True
        End If
        Exit Function
    End If

    Dim backupPath As String: backupPath = themeFolder & backupRel
    If Not ThemeFileExists(backupPath) Then Exit Function

    Dim destFolder As String: destFolder = themeFolder & category & "\"
    EnsureFolder destFolder
    RemoveAllVariants destFolder, resName

    Dim ext As String: ext = "." & GetFileExt(backupPath)
    Dim destPath As String: destPath = destFolder & resName & ext

    On Error GoTo EH
    FileCopy backupPath, destPath

    ' FIX: xem chu thich trong ReplaceResourceImage -- phai xoa cache anh
    ' cua destPath ngay sau khi ghi de, neu khong control se tiep tuc hien
    ' anh CU (da doi tam) sau khi Reset vi GetOrLoadImage doc tu cache.
    SetCachedImage destPath, Nothing

    WriteManifestRelFile themeFolder, category, resName, category & "\" & resName & ext
    SetResourceHidden themeFolder, category, resName, False

    ClearThemeCache themeName
    If LCase$(GetCurrentTheme()) = LCase$(themeName) Then LoadTheme themeName

    RestoreOriginalResourceImage = True
    Exit Function
EH:
    RestoreOriginalResourceImage = False
End Function

'----------------------------------------------------------
' BackupCurrentResourceImage — Sao luu anh DANG GAN cho 1 resource (neu
' co) vao thu muc con "_Backup" trong chinh category do cua theme, roi
' ghi duong dan tuong doi vao section "<Category>_Backup" cua
' ComponentManifest.ini. Chi goi 1 LAN DUY NHAT boi ReplaceResourceImage,
' truoc khi ghi de anh moi lan dau — day chinh la "anh goc cua theme"
' ma RestoreOriginalResourceImage se phuc hoi lai sau nay.
'----------------------------------------------------------
Private Sub BackupCurrentResourceImage(ByVal themeFolder As String, ByVal category As String, ByVal resName As String)
    Dim curRel As String: curRel = GetResourceRelFileRaw(themeFolder, category, resName)
    If Len(curRel) = 0 Then Exit Sub   ' chua co anh nao -> khong co gi de sao luu

    Dim curPath As String: curPath = themeFolder & curRel
    If Not ThemeFileExists(curPath) Then Exit Sub

    Dim backupFolder As String: backupFolder = themeFolder & category & "\_Backup\"
    EnsureFolder backupFolder

    Dim ext As String: ext = "." & GetFileExt(curPath)
    Dim backupRel As String: backupRel = category & "\_Backup\" & resName & ext
    Dim backupPath As String: backupPath = themeFolder & backupRel

    On Error Resume Next
    FileCopy curPath, backupPath
    On Error GoTo 0

    If ThemeFileExists(backupPath) Then
        INIWrite manifestPath(themeFolder), BackupSectionName(category), resName, backupRel
    End If
End Sub

' -- Private: doc RelFile TRUC TIEP tu ComponentManifest.ini (khong qua
' cache m_ManifestCache) — dung khi can gia tri MOI NHAT ngay trong luc
' dang ghi/doi anh (vd BackupCurrentResourceImage).
Private Function GetResourceRelFileRaw(ByVal themeFolder As String, ByVal category As String, ByVal resName As String) As String
    Dim raw As String: raw = INIRead(manifestPath(themeFolder), category, resName)
    If Len(raw) = 0 Then Exit Function
    Dim parsed As Object: Set parsed = ParseManifestEntry(raw)
    GetResourceRelFileRaw = CStr(parsed("RelFile"))
End Function

Private Function BackupSectionName(ByVal category As String) As String
    BackupSectionName = category & "_Backup"
End Function

Private Function HiddenSectionName(ByVal category As String) As String
    HiddenSectionName = category & "_Hidden"
End Function

Private Function GetBackupRelFile(ByVal themeFolder As String, ByVal category As String, ByVal resName As String) As String
    GetBackupRelFile = INIRead(manifestPath(themeFolder), BackupSectionName(category), resName)
End Function

'----------------------------------------------------------
' IsResourceHidden — Public: dung boi modTheme.ResolveImagePath de biet
' 1 resource co dang bi "ngat lien ket hien thi" (Shift+Right-Click /
' ClearResourceImage) hay khong, du RelFile cua no van con nguyen trong
' manifest.
'----------------------------------------------------------
Public Function IsResourceHidden(ByVal themeName As String, ByVal category As String, ByVal resName As String) As Boolean
    Dim themeFolder As String: themeFolder = GetThemeFolder(themeName)
    IsResourceHidden = (INIRead(manifestPath(themeFolder), HiddenSectionName(category), resName) = "1")
End Function

Private Sub SetResourceHidden(ByVal themeFolder As String, ByVal category As String, ByVal resName As String, ByVal hidden As Boolean)
    If hidden Then
        INIWrite manifestPath(themeFolder), HiddenSectionName(category), resName, "1"
    Else
        INIDeleteKey manifestPath(themeFolder), HiddenSectionName(category), resName
    End If
End Sub

'----------------------------------------------------------
' RenameResource — Doi ten logic cua 1 resource (file + key manifest)
'----------------------------------------------------------
'Public Function RenameResource(ByVal themeName As String, _
'                               ByVal category As String, _
'                               ByVal oldName As String, _
'                               ByVal newName As String) As Boolean
'    newName = SafeName(Trim$(newName))
'    If Len(newName) = 0 Then Exit Function

'    Dim themeFolder As String: themeFolder = GetThemeFolder(themeName)
'    Dim folder As String: folder = themeFolder & category & "\"

'    Dim exts() As String: exts = Split(".png|.jpg|.jpeg|.bmp|.gif", "|")
'    Dim i As Long
'    Dim didRename As Boolean
'    For i = 0 To UBound(exts)
'        If FileExists(folder & oldName & exts(i)) Then
'            If RenameFileSafe(folder & oldName & exts(i), folder & newName & exts(i)) Then
'                INIDeleteKey manifestPath(themeFolder), category, oldName
'                WriteManifestRelFile themeFolder, category, newName, category & "\" & newName & exts(i)
'                didRename = True
'            End If
'            Exit For
'        End If
'    Next i
'
'    ClearThemeCache themeName
'    If LCase$(GetCurrentTheme()) = LCase$(themeName) Then LoadTheme themeName
'    RenameResource = didRename
'End Function

' -- Private: xoa moi file cung resName (khac ext) trong 1 thu muc
Private Sub RemoveAllVariants(ByVal folder As String, ByVal resName As String)
    Dim exts() As String: exts = Split(".png|.jpg|.jpeg|.bmp|.gif", "|")
    Dim i As Long
    For i = 0 To UBound(exts)
        DeleteFileSafe folder & resName & exts(i)
    Next i
End Sub

'----------------------------------------------------------
' WriteManifestRelFile — Ghi RelFile MOI cho 1 resource vao
' ComponentManifest.ini cua theme, GIU NGUYEN phan size (width/height
' hoac ERR) da co san trong entry cu (neu co) — dung boi
' ReplaceResourceImage/RestoreDefaultResourceImage khi nguoi dung
' thay anh, de khong lam mat du lieu size da quet duoc truoc do.
'----------------------------------------------------------
Private Sub WriteManifestRelFile(ByVal themeFolder As String, ByVal category As String, _
                                  ByVal resName As String, ByVal newRelFile As String)
    Dim mPath As String: mPath = manifestPath(themeFolder)
    Dim raw As String: raw = INIRead(mPath, category, resName)

    Dim sizeSuffix As String: sizeSuffix = ""
    If Len(raw) > 0 Then
        Dim parsed As Object: Set parsed = ParseManifestEntry(raw)
        Dim sizeText As String: sizeText = CStr(parsed("SizeText"))
        If sizeText <> "N/A" And Left$(sizeText, 5) <> "N/A (" Then
            ' sizeText dang la "W x H pt" -> tach lai thanh "W|H" de ghi
            Dim posX As Long: posX = InStr(1, sizeText, " x ")
            Dim posPt As Long: posPt = InStrRev(sizeText, " pt")
            If posX > 0 And posPt > posX Then
                Dim wStr As String: wStr = Left$(sizeText, posX - 1)
                Dim hStr As String: hStr = mID$(sizeText, posX + 3, posPt - (posX + 3))
                sizeSuffix = "|" & wStr & "|" & hStr
            End If
        ElseIf Left$(sizeText, 5) = "N/A (" Then
            Dim errText As String: errText = mID$(sizeText, 6, Len(sizeText) - 6)
            sizeSuffix = "|ERR|" & errText
        End If
    End If

    INIWrite mPath, category, resName, "1|" & newRelFile & sizeSuffix
End Sub


