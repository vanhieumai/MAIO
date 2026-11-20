Attribute VB_Name = "modThemeResourceScanner"
Option Explicit

'===========================================================
' modResourceScanner — "Load Default Resources" (muc 1 Upgrade Plan)
'
' Quet toan bo VBA Project, tim tat ca UserForm, lay moi Image
' control / button vuong / imgBackground / logo va COPY (khong
' move) anh Picture dang gan san trong Form Designer vao dung
' cau truc Theme Default:
'     Icons\<ControlName>.<ext>
'     Backgrounds\<FormName>.<ext>     (control imgBackground)
'
' YEU CAU: File > Options > Trust Center > Trust access to the
' VBA project object model phai duoc BAT (chi can 1 lan, do day
' la thao tac quan tri, khong phai user cuoi dung).
'
' Chi copy file CON THIEU — khong ghi de file da ton tai, nen
' goi lai nhieu lan an toan (idempotent), dung yeu cau "lan sau
' mo Theme Manager se khong can thuc hien lai".
'
' CAP NHAT (Resource Size): ngoai ten component, moi lan quet
' gio con ghi kem KICH THUOC thiet ke (Width x Height, don vi
' points) cua tung component vao ComponentManifest.ini, de trang
' Resources co the hien thi size ma khong can mo Form Designer.
' Vi kich thuoc component co the thay doi giua cac lan thiet ke,
' RegisterComponentName gio LUON GHI DE gia tri size moi nhat
' (khac voi truoc day chi ghi khi key chua ton tai).
'
' CAP NHAT (gop resources.ini + ComponentManifest.ini):
' ComponentManifest.ini gio la file THEO TUNG THEME (nam ngay
' trong theme folder, cung cap voi theme.ini/colors.ini), KHONG
' con la 1 file GLOBAL duy nhat trong Themes\ nua — manifestPath()
' nhan tham so themeFolder. Moi entry gio luu them RelFile (ten
' file anh tuong doi, vi du "Icons\btnApply.png") ben canh size,
' theo dinh dang "1|<RelFile>|<W>|<H>". Day la nguon du lieu DUY
' NHAT cho ten anh — resources.ini khong con duoc module nay dung
' hay ghi nua (xem modResourceManager cho logic migrate/absorb tu
' resources.ini cu). Tham so "manifestFileOverride" cua
' ScanFormComponent/ScanControlsForImages/RegisterComponentName da
' bi BO vi khong con can thiet: manifest luon nam ngay trong
' destFolder dang duoc quet toi (theme folder thuc, hoac thu muc
' tam cua External Export), nen chi can manifestPath(destFolder).
'===========================================================

Private Const IMG_EXTS_SCAN As String = ".png|.jpg|.jpeg|.bmp|.gif"

' Ket qua tra ve cho UI hien thi tom tat sau khi quet
Public Type ScanResult
    FormsScanned    As Long
    ImagesFound     As Long
    ImagesCopied    As Long
    ImagesSkipped   As Long
    ImagesFailed    As Long
    ComponentsFound As Long   ' Tong component quet duoc, ke ca chua co anh
    SizesRecorded   As Long   ' MOI: so component da ghi duoc kich thuoc (W x H)
    ErrorMessage    As String
End Type

'----------------------------------------------------------
' LoadDefaultResourcesFromProject — Entry point cho btnLoadDefaultResources
'----------------------------------------------------------
' modResourceScanner.bas
Public Function LoadDefaultResourcesFromProject(Optional ByVal callingForm As Object = Nothing) As ScanResult
    Dim result As ScanResult
    On Error GoTo EH

    Dim destFolder As String: destFolder = GetDefaultThemeFolder()
    EnsureThemeSkeleton destFolder
    EnsureDefaultThemeIni destFolder

    Dim vbProj As Object
      Set vbProj = ThisWorkbook.VBProject

      Dim comp As Object
      For Each comp In vbProj.VBComponents
          If comp.Type = 3 Then
              result.FormsScanned = result.FormsScanned + 1

              ' FIX: form dang duoc mo (chinh la callingForm) PHAI luon quet qua
              ' ScanSelfForm (dung Picture da backup boi BackupOriginalPictures),
              ' kiem tra dieu nay TRUOC ca khi xet Designer. Ly do: khi 1 UserForm
              ' dang duoc Load/Show, comp.Designer KHONG phai 1 ban sao "thiet ke"
              ' doc lap -- no tra ve CUNG MOT instance voi form dang chay. Vi
              ' UserForm_Initialize da goi ApplyTheme Me truoc do, ca Picture cua
              ' control dang chay LAN comp.Designer.Controls(...).Picture deu da
              ' bi ghi de bang anh cua theme dang duoc chon (co the la anh
              ' thieu/hong) -- nen quet qua Designer cho chinh form nay se doc
              ' nham anh theme thay vi anh goc trong Form Designer.
              Dim isSelf As Boolean
              isSelf = (Not callingForm Is Nothing) And _
                       (StrComp(comp.Name, TypeName(callingForm), vbTextCompare) = 0)

              If isSelf Then
                  ScanSelfForm callingForm, comp.Name, destFolder, result
              Else
                  Dim frmDesignerTest As Object
                  On Error Resume Next
                  Set frmDesignerTest = comp.Designer
                  On Error GoTo 0

                  If Not frmDesignerTest Is Nothing Then
                      ' Voi cac form KHAC (dang dong, khong phai callingForm),
                      ' Designer van la ban thiet ke doc lap, an toan de doc.
                      ScanFormComponent comp, destFolder, result
                  End If
              End If
          End If
      Next comp
    
    LoadDefaultResourcesFromProject = result
    Exit Function

EH:
    result.ErrorMessage = "Cannot access the VBA project (Err " & Err.Number & ": " & _
        Err.Description & ")." & vbCrLf & _
        "Please enable: File > Options > Trust Center > Trust Center Settings > " & _
        "Macro Settings > Trust access to the VBA project object model."
    InvalidateManifestCache
    
    LoadDefaultResourcesFromProject = result
    
End Function
' Tao theme.ini + colors.ini mac dinh cho Default NEU CHUA CO — chi goi tu
' Load Default Resources (khong con tu dong chay khi mo form nua).
Private Sub EnsureDefaultThemeIni(ByVal defFolder As String)
    If ThemeFileExists(defFolder & "theme.ini") Then Exit Sub

    INIWrite defFolder & "theme.ini", "Theme", "Name", "Default"
    INIWrite defFolder & "theme.ini", "Theme", "Version", "1.0"
    INIWrite defFolder & "theme.ini", "Theme", "Author", "CPMS"
    INIWrite defFolder & "theme.ini", "Theme", "Description", "Default CPMS theme."
    INIWrite defFolder & "theme.ini", "Theme", "MinimumEngineVersion", "1.0"

    Dim k As Variant
    For Each k In GetDefaultColorMap().keys
        INIWrite defFolder & "colors.ini", "Colors", CStr(k), CStr(GetDefaultColorMap()(k))
    Next k
End Sub
' Quet form dang chay chinh no, dung Me thay vi comp.Designer
Private Sub ScanSelfForm(ByVal frm As Object, ByVal formName As String, _
                            ByVal destFolder As String, ByRef result As ScanResult)
      ' MOI (Resource Size): doc Width/Height an toan, co fallback +
      ' chan doan loi that neu that bai (xem TryReadSize).
      Dim wPt As Double, hPt As Double, szErr As String
      szErr = TryReadSize(frm, wPt, hPt)

      ' FIX: doc Picture GOC tu backup (frmThemeManager.BackupOriginalPictures /
      ' GetOriginalPicture), KHONG doc truc tiep frm.Picture. Ham nay chay SAU
      ' khi UserForm_Initialize da goi ApplyTheme Me, nen luc nay frm.Picture la
      ' anh cua THEME dang duoc chon (co the thieu/hong), khong con la anh goc
      ' gan san trong Form Designer nua.
      Dim selfPic As Object
      On Error Resume Next
      Set selfPic = frm.GetOriginalPicture(formName)
      On Error GoTo 0
      If Not selfPic Is Nothing Then
          If selfPic.Type <> 0 Then
              result.ImagesFound = result.ImagesFound + 1
              Select Case SavePictureIfMissing(selfPic, destFolder & "Backgrounds\" & formName)
                  Case 1:  result.ImagesCopied = result.ImagesCopied + 1
                  Case 0:  result.ImagesSkipped = result.ImagesSkipped + 1
                  Case -1: result.ImagesFailed = result.ImagesFailed + 1
              End Select
          End If
      End If

      ' MOI (RelFile): sau khi da luu/xac nhan anh, tim dung phan mo rong
      ' thuc te dang co tren dia de ghi vao manifest — hoat dong ca voi anh
      ' vua duoc quet MOI lan voi anh nguoi dung da REPLACE thu cong truoc do
      ' (ReplaceResourceImage luon dat ten file la "<resName>.<ext>").
      Dim relFile As String: relFile = ""
      Dim foundExt As String: foundExt = FindExistingImageFile(destFolder & "Backgrounds\" & formName)
      If Len(foundExt) > 0 Then relFile = "Backgrounds\" & formName & foundExt

      RegisterComponentName "Backgrounds", formName, destFolder, relFile, wPt, hPt, szErr
      result.ComponentsFound = result.ComponentsFound + 1
      If Len(szErr) = 0 Then result.SizesRecorded = result.SizesRecorded + 1

      ' Truyen "frm" lam useOriginalFrom de ScanControlsForImages doc Picture tu
      ' backup thay vi tu control dang chay (da bi ApplyTheme ghi de).
      ScanControlsForImages frm.Controls, formName, destFolder, result, frm
    End Sub
    
'----------------------------------------------------------
' ScanFormComponent — Quet 1 UserForm (design-time, qua Designer)
'----------------------------------------------------------
Private Sub ScanFormComponent(ByVal comp As Object, ByVal destFolder As String, ByRef result As ScanResult)
    Dim frmDesigner As Object
    On Error Resume Next
    Set frmDesigner = comp.Designer
    On Error GoTo 0
    If frmDesigner Is Nothing Then Exit Sub

    ' MOI (Resource Size): doc Width/Height an toan, xem giai thich trong TryReadSize
    Dim wPt As Double, hPt As Double, szErr As String
    szErr = TryReadSize(frmDesigner, wPt, hPt)

    Dim frmPic As Object
    On Error Resume Next
    Set frmPic = frmDesigner.Picture
    On Error GoTo 0
    If Not frmPic Is Nothing Then
        If frmPic.Type <> 0 Then
            result.ImagesFound = result.ImagesFound + 1
            Select Case SavePictureIfMissing(frmPic, destFolder & "Backgrounds\" & comp.Name)
                Case 1:  result.ImagesCopied = result.ImagesCopied + 1
                Case 0:  result.ImagesSkipped = result.ImagesSkipped + 1
                Case -1: result.ImagesFailed = result.ImagesFailed + 1
            End Select
        End If
    End If

    Dim relFile As String: relFile = ""
    Dim foundExt As String: foundExt = FindExistingImageFile(destFolder & "Backgrounds\" & comp.Name)
    If Len(foundExt) > 0 Then relFile = "Backgrounds\" & comp.Name & foundExt

    RegisterComponentName "Backgrounds", comp.Name, destFolder, relFile, wPt, hPt, szErr
    result.ComponentsFound = result.ComponentsFound + 1
    If Len(szErr) = 0 Then result.SizesRecorded = result.SizesRecorded + 1

    ScanControlsForImages frmDesigner.Controls, comp.Name, destFolder, result
End Sub

'----------------------------------------------------------
' ScanControlsForImages — Quet de quy Controls cua 1 form (design-time)
'----------------------------------------------------------
Private Sub ScanControlsForImages(ByVal Controls As Object, ByVal formName As String, _
                                   ByVal destFolder As String, ByRef result As ScanResult, _
                                   Optional ByVal useOriginalFrom As Object = Nothing)
    ' FIX: tham so "useOriginalFrom". Khi KHAC Nothing (chi truyen tu
    ' ScanSelfForm cho chinh form dang mo), Picture cua tung control/Page duoc
    ' doc qua useOriginalFrom.GetOriginalPicture(...) (anh da backup TRUOC khi
    ' ApplyTheme chay) thay vi doc truc tiep ctrl.Picture/pg.Picture -- vi voi
    ' form dang mo, Picture "song" luc nay da bi ApplyTheme ghi de thanh anh
    ' cua theme dang chon. Voi cac form KHAC (useOriginalFrom = Nothing, goi tu
    ' ScanFormComponent qua Designer), hanh vi doc ctrl.Picture truc tiep giu
    ' nguyen nhu cu vi Designer cua form dong la ban thiet ke doc lap, an toan.
    Dim ctrl As Object
    For Each ctrl In Controls
        On Error Resume Next
        Dim tName As String: tName = TypeName(ctrl)
        On Error GoTo 0

        Select Case tName
            Case "Frame"
                ScanControlsForImages ctrl.Controls, formName, destFolder, result, useOriginalFrom

            Case "MultiPage"
                Dim pg As Object
                For Each pg In ctrl.Pages
                    Dim pgResName As String: pgResName = formName & "_" & pg.Name

                    ' MOI (Resource Size): doc Width/Height an toan, xem giai thich trong TryReadSize.
                    ' Page thuong KHONG co .Width/.Height rieng (kich thuoc luon bang client-area
                    ' cua MultiPage cha) nen se tu dong fallback sang .InsideWidth/.InsideHeight.
                    Dim pgW As Double, pgH As Double, pgErr As String
                    pgErr = TryReadSize(pg, pgW, pgH)

                    Dim pgPic As Object
                    On Error Resume Next
                    If useOriginalFrom Is Nothing Then
                        Set pgPic = pg.Picture
                    Else
                        Set pgPic = useOriginalFrom.GetOriginalPicture(pgResName)
                    End If
                    On Error GoTo 0
                    If Not pgPic Is Nothing Then
                        If pgPic.Type <> 0 Then
                            result.ImagesFound = result.ImagesFound + 1
                            Select Case SavePictureIfMissing(pgPic, destFolder & "Backgrounds\" & pgResName)
                                Case 1:  result.ImagesCopied = result.ImagesCopied + 1
                                Case 0:  result.ImagesSkipped = result.ImagesSkipped + 1
                                Case -1: result.ImagesFailed = result.ImagesFailed + 1
                            End Select
                        End If
                    End If

                    Dim pgRelFile As String: pgRelFile = ""
                    Dim pgFoundExt As String: pgFoundExt = FindExistingImageFile(destFolder & "Backgrounds\" & pgResName)
                    If Len(pgFoundExt) > 0 Then pgRelFile = "Backgrounds\" & pgResName & pgFoundExt

                    RegisterComponentName "Backgrounds", pgResName, destFolder, pgRelFile, pgW, pgH, pgErr
                    result.ComponentsFound = result.ComponentsFound + 1
                    If Len(pgErr) = 0 Then result.SizesRecorded = result.SizesRecorded + 1

                    ScanControlsForImages pg.Controls, formName, destFolder, result, useOriginalFrom
                Next pg

            Case "Image", "CommandButton", "ToggleButton", "CheckBox", "OptionButton"
                Dim category As String
                Dim resName  As String

                If LCase$(ctrl.Name) = "imgbackground" Then
                    category = "Backgrounds"
                    resName = formName
                ElseIf InStr(1, LCase$(ctrl.Name), "logo", vbTextCompare) > 0 Then
                    category = "Backgrounds"
                    resName = ctrl.Name
                Else
                    category = "Icons"
                    resName = ctrl.Name
                End If

                ' MOI (Resource Size): doc Width/Height an toan, xem giai thich trong TryReadSize
                Dim ctW As Double, ctH As Double, ctErr As String
                ctErr = TryReadSize(ctrl, ctW, ctH)

                Dim pic As Object
                On Error Resume Next
                If useOriginalFrom Is Nothing Then
                    Set pic = ctrl.Picture
                Else
                    Set pic = useOriginalFrom.GetOriginalPicture(ctrl.Name)
                End If
                On Error GoTo 0
                'If pic Is Nothing Then
                '    Debug.Print "    ctrl=" & ctrl.Name & " | pic Is Nothing=True"
                'Else
                '    Debug.Print "    ctrl=" & ctrl.Name & " | pic Is Nothing=False | Type=" & pic.Type
                'End If
                If Not pic Is Nothing Then
                    If pic.Type <> 0 Then
                        result.ImagesFound = result.ImagesFound + 1
                        Select Case SavePictureIfMissing(pic, destFolder & category & "\" & resName)
                            Case 1:  result.ImagesCopied = result.ImagesCopied + 1
                            Case 0:  result.ImagesSkipped = result.ImagesSkipped + 1
                            Case -1: result.ImagesFailed = result.ImagesFailed + 1
                        End Select
                    End If
                End If

                Dim ctRelFile As String: ctRelFile = ""
                Dim ctFoundExt As String: ctFoundExt = FindExistingImageFile(destFolder & category & "\" & resName)
                If Len(ctFoundExt) > 0 Then ctRelFile = category & "\" & resName & ctFoundExt

                RegisterComponentName category, resName, destFolder, ctRelFile, ctW, ctH, ctErr
                result.ComponentsFound = result.ComponentsFound + 1
                If Len(ctErr) = 0 Then result.SizesRecorded = result.SizesRecorded + 1
        End Select
    Next ctrl
End Sub
'----------------------------------------------------------
' SavePictureIfMissing — Ghi StdPicture ra file neu CHUA ton tai
' file nao cung ten (bat ky ext). Tra ve True neu da ghi moi.
'
' Ghi chu ky thuat: VBA khong co API xuat StdPicture ra .png/.jpg
' truc tiep. Cach on dinh nhat la dung SavePicture (xuat .bmp),
' sau do doi ten logic sang .bmp trong registry neu chua co
' dinh dang nen. Neu can .png that su, nguoi dung co the mo
' Icons\ va convert thu cong — day khong phai thao tac thuong
' xuyen vi Load Default Resources chi chay 1 lan luc khoi tao.
'----------------------------------------------------------
Private Function SavePictureIfMissing(ByVal pic As Object, ByVal destPathNoExt As String) As Long
    ' 0 = skipped (existed), 1 = copied, -1 = failed
    If Len(FindExistingImageFile(destPathNoExt)) > 0 Then
        SavePictureIfMissing = 0   ' existed/skipped
        Exit Function
    End If
    On Error GoTo EH
    EnsureFolder GetParentFolder(destPathNoExt & ".bmp")
    SavePicture pic, destPathNoExt & ".bmp"
    SavePictureIfMissing = 1   ' success
    Exit Function
EH:
    Debug.Print "SavePicture FAILED: " & destPathNoExt & " -- Err " & Err.Number & ": " & Err.Description
    SavePictureIfMissing = -1  ' failed
End Function

'----------------------------------------------------------
' FindExistingImageFile — MOI: kiem tra xem "<destPathNoExt>.<ext>" da
' ton tai voi ext nao (trong danh sach IMG_EXTS_SCAN), tra ve ext do
' (ke ca dau ".", vi du ".png") hoac "" neu khong tim thay file nao.
' Dung de xac dinh RelFile ghi vao manifest — hoat dong dung cho ca
' anh vua duoc SavePictureIfMissing tao MOI (luon la .bmp) VA anh da
' duoc nguoi dung REPLACE thu cong truoc do (co the la .png/.jpg/...,
' vi ReplaceResourceImage luon dat ten "<resName>.<ext>").
'----------------------------------------------------------
Public Function FindExistingImageFile(ByVal destPathNoExt As String) As String
    Dim exts() As String: exts = Split(IMG_EXTS_SCAN, "|")
    Dim i As Long
    For i = 0 To UBound(exts)
        If ThemeFileExists(destPathNoExt & exts(i)) Then
            FindExistingImageFile = exts(i)
            Exit Function
        End If
    Next i
End Function

'----------------------------------------------------------
' manifestPath — MOI: ComponentManifest.ini gio nam NGAY TRONG
' theme folder (hoac thu muc tam cua External Export) duoc truyen
' vao, khong con la 1 file GLOBAL duy nhat trong Themes\ nua.
'----------------------------------------------------------
Public Function manifestPath(ByVal themeFolder As String) As String
    manifestPath = themeFolder & "ComponentManifest.ini"
End Function

'----------------------------------------------------------
' TryReadSize — MOI: doc kich thuoc 1 doi tuong (UserForm/Designer/
' Page/Control) mot cach an toan va co fallback.
'
' Mot so doi tuong (dac biet la MSForms.Page trong MultiPage — vi kich
' thuoc cua 1 Page luon bang dung client-area cua MultiPage cha, KHONG
' the tu resize doc lap) khong ho tro .Width/.Height tren moi phien ban
' Office. Ham nay thu .Width/.Height truoc, neu that bai thu tiep
' .InsideWidth/.InsideHeight (Page chac chan co 2 property nay).
'
' Tra ve "" (chuoi rong) neu doc thanh cong (w/h da duoc gan gia tri
' >= 0). Neu that bai ca 2 cach, tra ve chuoi mo ta LOI THAT (Err.Number
' + Err.Description cua lan thu cuoi) de ghi vao manifest lam du lieu
' chan doan — thay vi chi im lang ghi "N/A" khong ro nguyen nhan.
'----------------------------------------------------------
Private Function TryReadSize(ByVal obj As Object, ByRef w As Double, ByRef h As Double) As String
    Dim lastErrNum As Long, lastErrDesc As String

    w = -1: h = -1
    On Error Resume Next
    Err.Clear
    w = obj.Width
    h = obj.Height
    If Err.Number = 0 And w >= 0 And h >= 0 Then
        On Error GoTo 0
        TryReadSize = ""
        Exit Function
    End If
    lastErrNum = Err.Number: lastErrDesc = Err.Description
    On Error GoTo 0

    w = -1: h = -1
    On Error Resume Next
    Err.Clear
    w = obj.InsideWidth
    h = obj.InsideHeight
    If Err.Number = 0 And w >= 0 And h >= 0 Then
        On Error GoTo 0
        TryReadSize = ""
        Exit Function
    End If
    lastErrNum = Err.Number: lastErrDesc = Err.Description
    On Error GoTo 0

    w = -1: h = -1
    TryReadSize = "Err " & lastErrNum & ": " & lastErrDesc
End Function

' Ghi ten component + RelFile + kich thuoc (Width|Height, points) vao
' ComponentManifest.ini CUA THEME dang duoc quet (destFolder). Gia tri
' duoc ghi dang:
'   "1"                              — chua biet RelFile, chua biet size
'   "1|<RelFile>"                    — biet RelFile, chua biet size
'   "1|<RelFile>|<width>|<height>"   — biet RelFile + size hop le
'   "1|<RelFile>|ERR|<mo ta loi>"    — biet RelFile, doc size that bai
'                                       kem ly do that (Err.Number/Description)
'                                       de nguoi dung nhin thay tren label
'                                       thay vi "N/A" mu mo.
' (Cac dinh dang cu khong co RelFile — "1", "1|<w>|<h>", "1|ERR|<desc>" —
'  van duoc ParseManifestEntry hieu dung, xem modResourceManager.)
' LUON GHI DE de cap nhat gia tri moi nhat moi lan quet (khac voi ban
' truoc chi ghi khi key chua ton tai).
Private Sub RegisterComponentName(ByVal category As String, ByVal resName As String, ByVal destFolder As String, _
                                   Optional ByVal relFile As String = "", _
                                   Optional ByVal widthPt As Double = -1, _
                                   Optional ByVal heightPt As Double = -1, _
                                   Optional ByVal sizeErr As String = "")
    Dim valueToWrite As String: valueToWrite = "1"
    If Len(relFile) > 0 Then valueToWrite = valueToWrite & "|" & relFile

    If widthPt >= 0 And heightPt >= 0 Then
        valueToWrite = valueToWrite & "|" & Format$(widthPt, "0.##") & "|" & Format$(heightPt, "0.##")
    ElseIf Len(sizeErr) > 0 Then
        valueToWrite = valueToWrite & "|ERR|" & Replace(sizeErr, "|", ";")
    End If

    INIWrite manifestPath(destFolder), category, resName, valueToWrite
End Sub

'----------------------------------------------------------
' ParseManifestEntry — MOI: parser TRUNG TAM cho 1 dong gia tri trong
' ComponentManifest.ini (xem dinh dang trong RegisterComponentName).
' Dung chung boi modResourceScanner va modResourceManager de tranh
' trung lap logic. Tra ve Dictionary voi 2 key:
'   "RelFile"  — ten file anh tuong doi (vd "Icons\btnApply.png"),
'                "" neu chua biet.
'   "SizeText" — chuoi hien thi "W x H pt", "N/A", hoac "N/A (loi...)"
' Tuong thich nguoc voi dinh dang CU (truoc khi co RelFile), khi
' parts(1) la mot so (width) hoac "ERR" thay vi ten file.
'----------------------------------------------------------
Public Function ParseManifestEntry(ByVal raw As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d.Add "RelFile", ""
    d.Add "SizeText", "N/A"

    If Len(raw) = 0 Then
        Set ParseManifestEntry = d
        Exit Function
    End If

    Dim parts() As String: parts = Split(raw, "|")
    If UBound(parts) < 1 Then
        Set ParseManifestEntry = d   ' chi co "1", chua co gi khac
        Exit Function
    End If

    Dim idx As Long: idx = 1   ' parts(0) luon la "1" (co flag)

    ' parts(1) la RelFile NEU no khong phai so va khong phai "ERR"
    ' (dinh dang cu, truoc khi co RelFile, luu truc tiep width hoac
    ' "ERR" ngay tai vi tri nay).
    If Not IsNumeric(parts(idx)) And UCase$(parts(idx)) <> "ERR" Then
        d("RelFile") = parts(idx)
        idx = idx + 1
    End If

    If idx > UBound(parts) Then
        Set ParseManifestEntry = d   ' chi co RelFile (hoac khong co gi), chua co size
        Exit Function
    End If

    If UCase$(parts(idx)) = "ERR" Then
        Dim i As Long, errText As String
        errText = parts(idx + 1)
        For i = idx + 2 To UBound(parts)   ' phong khi mo ta loi co chua ky tu "|"
            errText = errText & "|" & parts(i)
        Next i
        d("SizeText") = "N/A (" & errText & ")"
        Set ParseManifestEntry = d
        Exit Function
    End If

    If idx + 1 <= UBound(parts) Then
        On Error GoTo EH
        d("SizeText") = Format$(CDbl(parts(idx)), "0.##") & " x " & Format$(CDbl(parts(idx + 1)), "0.##") & " pt"
    End If
    Set ParseManifestEntry = d
    Exit Function
EH:
    d("SizeText") = "N/A"
    Set ParseManifestEntry = d
End Function



'===========================================================
' MOI (External Export) - Quet resources tu 1 file VBA KHAC (chua mo,
' do nguoi dung chon qua hop thoai) va xuat ra 1 goi .zip DOC LAP,
' KHONG dong bo/ghi de vao bat ky Theme nao cua project hien tai.
' Dung boi frmThemeManager.btnLoadExternalResources_Click.
'===========================================================

' GetFileBaseName - tra ve ten file khong co duong dan va khong co
' phan mo rong. Vi du: "C:\A\Report.xlsm" -> "Report"
Public Function GetFileBaseName(ByVal fullPath As String) As String
    Dim p As Long, d As Long, nameOnly As String
    p = InStrRev(fullPath, "\")
    nameOnly = mID$(fullPath, p + 1)
    d = InStrRev(nameOnly, ".")
    If d > 0 Then
        GetFileBaseName = Left$(nameOnly, d - 1)
    Else
        GetFileBaseName = nameOnly
    End If
End Function

' DeleteFolderRecursive - xoa 1 thu muc tam (kem toan bo noi dung ben
' trong) sau khi da dong goi xong. Im lang bo qua neu khong xoa duoc
' (vi day chi la don dep thu muc tam, khong phai loi nghiem trong).
Private Sub DeleteFolderRecursive(ByVal folderPath As String)
    On Error Resume Next
    If Len(folderPath) = 0 Then Exit Sub
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then fso.DeleteFolder folderPath, True
    On Error GoTo 0
End Sub

' CreateZipFromFolder - nen toan bo noi dung cua srcFolder vao 1 file
' .zip moi tai destZipPath, dung ky thuat pho bien trong VBA (tao file
' rong voi header PK chuan roi dung Shell.Application.NameSpace.CopyHere),
' khong can thu vien/COM reference ngoai nao khac.
Private Function CreateZipFromFolder(ByVal srcFolder As String, ByVal destZipPath As String) As Boolean
    On Error GoTo EH
    If Len(Dir$(destZipPath)) > 0 Then Kill destZipPath

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tS As Object: Set tS = fso.CreateTextFile(destZipPath, True)
    tS.Write Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String$(18, Chr$(0))
    tS.Close

    Dim shellApp As Object: Set shellApp = CreateObject("Shell.Application")
    Dim zipNs As Object: Set zipNs = shellApp.Namespace(CVar(destZipPath))
    Dim srcNs As Object: Set srcNs = shellApp.Namespace(CVar(srcFolder))
    If zipNs Is Nothing Or srcNs Is Nothing Then GoTo EH

    zipNs.CopyHere srcNs.items, 4 Or 16   ' 4 = an progress dialog, 16 = tu dong Yes to All

    ' CopyHere chay bat dong bo -- doi den khi so item trong zip khop voi
    ' thu muc nguon (toi da ~60 giay, tranh treo app neu Explorer gap loi).
    Dim waited As Long
    Do While zipNs.items.count < srcNs.items.count
        Application.Wait Now + TimeSerial(0, 0, 1)
        waited = waited + 1
        If waited > 60 Then Exit Do
    Loop

    CreateZipFromFolder = (zipNs.items.count >= srcNs.items.count) And (Len(Dir$(destZipPath)) > 0)
    Exit Function
EH:
    CreateZipFromFolder = False
End Function

'----------------------------------------------------------
' ScanExternalProjectToZip - Entry point cho btnLoadExternalResources.
' Mo READ-ONLY 1 file Excel co macro KHAC (nguoi dung chon), quet toan
' bo UserForm cua file do (giong ScanFormComponent/ScanControlsForImages
' o tren, nhung ghi vao manifest RIENG trong thu muc tam — vi destFolder
' truyen xuong la tempFolder nen manifestPath(tempFolder) tu dong dung
' vi tri, KHONG can tham so rieng nao nua), roi nen ket qua (Icons\,
' Backgrounds\, ComponentManifest.ini) thanh 1 file .zip tai destZipPath.
' File nguon LUON duoc dong lai (khong luu) sau khi quet, va thu muc tam
' LUON bi xoa sau khi nen xong -- KHONG dung lai bat ky du lieu nao
' trong Theme cua project hien tai.
'----------------------------------------------------------
Public Function ScanExternalProjectToZip(ByVal sourceFilePath As String, ByVal destZipPath As String) As ScanResult
    Dim result As ScanResult
    Dim tempFolder As String
    Dim extWb As Object
    On Error GoTo EH

    If Len(Dir$(sourceFilePath)) = 0 Then
        result.ErrorMessage = "Source file not found:" & vbCrLf & sourceFilePath
        ScanExternalProjectToZip = result
        Exit Function
    End If

    ' Neu file da dang mo san trong Excel, KHONG dung Workbooks.Open/Close
    ' voi no -- lam vay co the vo tinh dong mat workbook that cua nguoi
    ' dung (co the dang co thay doi chua luu). Yeu cau dong lai truoc.
    Dim wbCheck As Object
    For Each wbCheck In Application.Workbooks
        If StrComp(wbCheck.fullName, sourceFilePath, vbTextCompare) = 0 Then
            result.ErrorMessage = "This file is already open (""" & wbCheck.Name & """)." & vbCrLf & _
                "Please close it first, then try again."
            ScanExternalProjectToZip = result
            Exit Function
        End If
    Next wbCheck

    tempFolder = Environ$("TEMP") & "\CPMS_ExternalExport_" & Format$(Now, "yyyymmdd_hhnnss") & "\"
    EnsureFolder tempFolder & "Icons\"
    EnsureFolder tempFolder & "Backgrounds\"

    Set extWb = Application.Workbooks.Open(sourceFilePath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)

    Dim vbProj As Object
    On Error Resume Next
    Set vbProj = extWb.VBProject
    On Error GoTo 0
    If vbProj Is Nothing Then
        extWb.Close SaveChanges:=False
        Set extWb = Nothing
        DeleteFolderRecursive tempFolder
        result.ErrorMessage = "Cannot access the VBA project of the selected file (Err " & Err.Number & ")." & vbCrLf & _
            "Please enable: File > Options > Trust Center > Trust Center Settings > " & _
            "Macro Settings > Trust access to the VBA project object model."
        ScanExternalProjectToZip = result
        Exit Function
    End If

    Dim comp As Object
    For Each comp In vbProj.VBComponents
        If comp.Type = 3 Then   ' vbext_ct_MSForm
            result.FormsScanned = result.FormsScanned + 1
            Dim frmDesignerTest As Object
            On Error Resume Next
            Set frmDesignerTest = comp.Designer
            On Error GoTo 0
            If Not frmDesignerTest Is Nothing Then
                ScanFormComponent comp, tempFolder, result
            End If
        End If
    Next comp

    extWb.Close SaveChanges:=False
    Set extWb = Nothing

    If result.ComponentsFound = 0 Then
        result.ErrorMessage = "No UserForm/component found in the selected file."
        DeleteFolderRecursive tempFolder
        ScanExternalProjectToZip = result
        Exit Function
    End If

    If Not CreateZipFromFolder(tempFolder, destZipPath) Then
        DeleteFolderRecursive tempFolder
        result.ErrorMessage = "Failed to create the zip package at:" & vbCrLf & destZipPath
        ScanExternalProjectToZip = result
        Exit Function
    End If

    DeleteFolderRecursive tempFolder
    ScanExternalProjectToZip = result
    Exit Function

EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.Description
    On Error Resume Next
    If Not extWb Is Nothing Then extWb.Close SaveChanges:=False
    DeleteFolderRecursive tempFolder
    On Error GoTo 0
    result.ErrorMessage = "Error while scanning the external file (Err " & errNum & ": " & errDesc & ")."
    ScanExternalProjectToZip = result
End Function


