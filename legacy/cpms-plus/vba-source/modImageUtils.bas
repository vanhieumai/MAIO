Attribute VB_Name = "modImageUtils"
'modImageUtils
Option Explicit

' =========================================================
' QSM IMAGE MANAGER — THEME-AWARE VERSION
' PNG/JPG/BMP/GIF supported
'
' NANG CAP (nam cho cau truc Theme Engine):
' Truoc day module nay tu quan ly 1 kho anh RIENG, luu tai
' %LOCALAPPDATA%\CPMS_R\defaults|users\<TenControl>.bmp — hoan toan
' tach biet voi modTheme/modResourceManager. Vi vay cac icon do
' nguoi dung tuy chinh (Image1..Image6 cua frmLinkKL) KHONG theo
' theme: doi theme khong doi icon, icon khong duoc dong goi trong
' file .cpmstheme, va khong xuat hien trong Theme Manager.
'
' Gio module nay chi con la 1 lop ADAPTER MONG: giu nguyen API cong
' khai (ChangeImage/ResetImage/ClearImage/ExportImage) de KHONG phai
' sua cac diem goi trong frmLinkKL (HandleImageMouseUp), nhung ben
' trong da chuyen sang doc/ghi qua modResourceManager + modTheme,
' dung dung category "Icons" + resName = ten control — giong het quy
' uoc ma modResourceScanner.ScanControlsForImages da dung cho moi
' Image control khac trong toan bo project. Nghia la (4 thao tac deu
' CHI anh huong theme HIEN TAI, khong dong bo/anh huong theme khac):
'   - Ctrl+Click (Change) -> modResourceManager.ReplaceResourceImage:
'         CHI la thay TAM anh cua theme nay sang anh khac; anh GOC cua
'         theme duoc tu dong sao luu (lan doi dau tien) nen KHONG bi mat.
'   - Shift+Click (Reset) -> modResourceManager.RestoreOriginalResourceImage:
'         khoi phuc lai DUNG anh GOC cua CHINH theme dang chon (tu ban
'         sao luu noi tren) — KHONG phai lay tu theme Default (khac
'         RestoreDefaultResourceImage, van con dung rieng o Theme Manager).
'   - Shift+Right-Click (Clear) -> modResourceManager.ClearResourceImage:
'         CHI ngat LIEN KET HIEN THI (an icon di), KHONG xoa file anh
'         hay RelFile da gan — Change/Reset o tren se tu dong go an lai.
'   - Export -> luu lai file anh goc dang nam trong theme (giu dinh dang),
'                fallback SavePicture neu khong tim duoc file goc
' Sau moi thao tac, ve lai control qua modTheme.RefreshControlImage —
' cung 1 co che ma ApplyTheme/LoadControlImages dung cho TOAN BO form.
' =========================================================

Private Const IMG_CATEGORY As String = "Icons"

' =========================================================
' INIT — giu lai CHI de tuong thich nguoc voi cac form cu con goi
' InitImages Me trong UserForm_Activate/Initialize. KHONG con lam gi:
' viec nap anh theo theme da duoc ApplyTheme/LoadControlImages xu ly
' tu dong cho MOI Image control (qua category "Icons"), nen goi ham
' nay gio la NO-OP. Cac form MOI (vd frmLinkKL) khong can goi ham nay
' nua — chi can dam bao UserForm_Initialize co goi ApplyTheme Me.
' =========================================================
Public Sub InitImages(frm As Object)
    ' No-op: xem chu thich dau file.
End Sub
Public Sub HandleImageMouseUp(img As MSForms.Image, _
                               ByVal Button As Integer, _
                               ByVal Shift As Integer)

    If Button = 2 And (Shift And 2) <> 0 Then
        ExportImage img
        Exit Sub
    End If

    If Button = 2 And (Shift And 1) <> 0 Then
        ClearImage img
        img.Visible = False
        img.Visible = True
        Exit Sub
    End If

    If Button = 1 And (Shift And 1) <> 0 Then
        ResetImage img
        img.Visible = False
        img.Visible = True
        Exit Sub
    End If

    If Button = 1 And (Shift And 2) <> 0 Then
        ChangeImage img
        img.Visible = False
        img.Visible = True
        Exit Sub
    End If

End Sub


' =========================================================
' CHANGE IMAGE — chon file moi, luu vao Icons\ cua THEME HIEN TAI
' =========================================================
Public Sub ChangeImage(img As MSForms.Image)

    Dim fd As FileDialog
    Dim p As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .title = "Select image"
        .Filters.Clear
        .Filters.Add "Images", "*.jpg;*.jpeg;*.png;*.bmp;*.gif"
        If .Show <> -1 Then Exit Sub
        p = .SelectedItems(1)
    End With

    ' FIX: giai phong tham chieu Picture cua control TRUOC KHI ghi/xoa file
    ' vat ly (RemoveAllVariants + FileCopy ben trong ReplaceResourceImage).
    ' Neu khong, GDI+/OLE co the con giu khoa file anh dang hien thi, khien
    ' Kill/FileCopy de len dung file do that bai NGAM (khong loi) mot cach
    ' khong on dinh -> phai bam nhieu lan moi "an".
    On Error Resume Next
    Set img.Picture = Nothing
    On Error GoTo 0

    If Not ReplaceResourceImage(GetCurrentTheme(), IMG_CATEGORY, img.Name, p) Then
        ShowMsgCenter "Cannot update image.", vbExclamation
    End If
    RefreshControlImage img, IMG_CATEGORY

End Sub


' =========================================================
' RESET — khoi phuc anh GOC cua CHINH theme hien tai (anh da luu truoc
' lan Ctrl+Click dau tien), KHONG phai lay tu theme Default
' =========================================================
Public Sub ResetImage(img As MSForms.Image)

    ' FIX: xem chu thich trong ChangeImage o tren -- phai Set Picture =
    ' Nothing TRUOC khi RestoreOriginalResourceImage dung Kill/FileCopy
    ' ghi de dung file dang hien thi, tranh sharing-violation ngam khien
    ' Reset chap chon (phai bam nhieu lan, luc duoc luc khong).
    On Error Resume Next
    Set img.Picture = Nothing
    On Error GoTo 0

    ' Theo yeu cau: khong hien MsgBox nua du co gi de reset hay khong -
    ' im lang thoat neu khong co gi thay doi (chua tung Change/Clear).
    RestoreOriginalResourceImage GetCurrentTheme(), IMG_CATEGORY, img.Name
    RefreshControlImage img, IMG_CATEGORY

End Sub


' =========================================================
' CLEAR — an icon di, CHI trong theme hien tai (khong dong bo)
' =========================================================
Public Sub ClearImage(img As MSForms.Image)

    On Error Resume Next
    Set img.Picture = Nothing
    On Error GoTo 0

    ClearResourceImage GetCurrentTheme(), IMG_CATEGORY, img.Name
    RefreshControlImage img, IMG_CATEGORY

End Sub


' =========================================================
' EXPORT — luu anh dang hien thi ra file, uu tien file goc trong
' theme (giu dung dinh dang PNG/JPG) hon la SavePicture (luon ra BMP)
' =========================================================
Public Sub ExportImage(img As MSForms.Image)

    If img.Picture Is Nothing Then Exit Sub

    Dim srcPath As String
    srcPath = ResolveImagePath(img.Name, IMG_CATEGORY)

    Dim defaultExt As String
    defaultExt = ".bmp"
    If Len(srcPath) > 0 Then defaultExt = "." & GetFileExt(srcPath)

    Dim SavePath As Variant
    SavePath = Application.GetSaveAsFilename( _
                    InitialFileName:=img.Name & defaultExt, _
                    FileFilter:="Image (*" & defaultExt & "), *" & defaultExt)

    If SavePath = False Then Exit Sub

    If Len(srcPath) > 0 And ThemeFileExists(srcPath) Then
        FileCopy srcPath, CStr(SavePath)
    Else
        SavePicture img.Picture, CStr(SavePath)
    End If

    ShowMsgCenter "Image saved successfully"

End Sub


