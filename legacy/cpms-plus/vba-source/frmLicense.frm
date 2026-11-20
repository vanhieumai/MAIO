Attribute VB_Name = "frmLicense"
Attribute VB_Base = "0{91B3D025-81DB-46F8-B33C-1A8639B3B7FB}{00750953-78B3-4982-B29F-BFE11919D270}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'form frmLicense
'Tách t? form LinkKL — ch?a toàn b? n?i dung trang License
Option Explicit



Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
Alias "URLDownloadToFileA" ( _
ByVal pCaller As LongPtr, _
ByVal szURL As String, _
ByVal szFileName As String, _
ByVal dwReserved As LongPtr, _
ByVal lpfnCB As LongPtr) As Long
'Private mTargetWorkbook As Workbook

'==============================
' FORM INITIALIZE
'==============================
Private Sub UserForm_Initialize()
    ApplyTheme Me
    lblCPMS.ForeColor = RGB(30, 30, 30)
    lblCPMSdetail.ForeColor = RGB(30, 30, 30)
    lblPhone.ForeColor = RGB(30, 30, 30)
    lblDeviceID.ForeColor = RGB(30, 30, 30)
    lblActivationDate.ForeColor = RGB(30, 30, 30)
    lblExpiry.ForeColor = RGB(30, 30, 30)
    lblCopyright.ForeColor = RGB(30, 30, 30)
    
End Sub

Private Sub imgHM1_MouseUp(ByVal Button As Integer, _
                             ByVal Shift As Integer, _
                             ByVal X As Single, _
                             ByVal Y As Single)
    HandleImageMouseUp imgHM1, Button, Shift
End Sub

Private Sub imgHM2_MouseUp(ByVal Button As Integer, _
                             ByVal Shift As Integer, _
                             ByVal X As Single, _
                             ByVal Y As Single)
    HandleImageMouseUp imgHM2, Button, Shift
End Sub
'==============================
' FORM ACTIVATE
'==============================
Private Sub UserForm_Activate()
    Static UIloaded As Boolean
    ' [FIX] Capture workbook moi nhat moi lan form duoc kich hoat
    '       Phai dat TRUOC guard UIloaded de cap nhat ca khi form duoc tai su dung
    '       Khong loc ThisWorkbook: khi chay truc tiep thi chinh no la workbook can dung;
    '       khi la add-in (.xlam) thi file an nen ActiveWorkbook tu la workbook nguoi dung.
    'If Not ActiveWorkbook Is Nothing Then
    '    Set mTargetWorkbook = ActiveWorkbook
    'End If


    If UIloaded Then Exit Sub

    '===================================
    ' LOAD DEVICE ID
    '===================================
    lblDeviceID.Caption = GetMachineID()

    '===================================
    ' LOAD QR
    '===================================
    TaoQR_VaoFrame

    '===================================
    ' UPDATE LICENSE UI
    '===================================
    UpdateLicenseUI

    lblCopyright.Caption = _
    "© 2025–" & Year(Date) & " vanhieumai @ VNCN E&C"

    UIloaded = True

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload Me
    End If
    UnregisterForm Me
End Sub


'==============================
' DEVICE ID — double-click to copy / Ctrl+dblclick to clear
'==============================
Private Sub lblDeviceID_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    ' Ctrl + DoubleClick = clear only if activated
    If GetKeyState(vbKeyControl) < 0 Then

        If HasLicense Then
            ClearActivation
        Else
            ShowMsgCenter "Not activated"
        End If

        Exit Sub
    End If

    CopyTextToClipboard GetMachineID
    ShowMsgCenter "Device ID copied"

End Sub


'==============================
' CLEAR ACTIVATION
'==============================
Private Sub ClearActivation()

    If Not HasLicense Then Exit Sub

    ClearLicense

    ShowMsgCenter "Activation cleared"

    UpdateLicenseUI

End Sub


'==============================
' UPDATE LICENSE UI
'==============================
Private Sub UpdateLicenseUI()

    Dim ok As Boolean
    ok = IsActivated()

    If ok Then
        btnActive.Enabled = False
        btnActive.Caption = "Activated"
        btnActive.BackColor = RGB(200, 200, 200)

        Dim actDate As String, expDate As String, daysLeft As Long

        If GetLicenseInfo(actDate, expDate, daysLeft) Then

            ' Format activation date dd/mm/yyyy
            Dim actDisplay As String
            If actDate <> "" Then
                actDisplay = Format(CDate(actDate), "dd/mm/yyyy")
            Else
                actDisplay = "-"
            End If
            lblActivationDate.Caption = "Activated : " & actDisplay

            ' Format expiry date dd/mm/yyyy
            If daysLeft = -1 Then
                lblExpiry.Caption = "Expires   : Permanent"
                lblDaysLeft.Caption = ""
                lblDaysLeft.ForeColor = RGB(0, 150, 0)
            ElseIf daysLeft > 0 Then
                lblExpiry.Caption = "Expires   : " & Format(CDate(expDate), "dd/mm/yyyy")
                lblDaysLeft.Caption = daysLeft & " day(s) remaining"
                lblDaysLeft.ForeColor = IIf(daysLeft <= 7, RGB(220, 80, 0), RGB(0, 120, 0))
            Else
                lblExpiry.Caption = "Expires   : " & Format(CDate(expDate), "dd/mm/yyyy")
                lblDaysLeft.Caption = "LICENSE EXPIRED"
                lblDaysLeft.ForeColor = RGB(200, 0, 0)
                ShowMsgCenter "Your software license expires today. Please renew it to continue using all features."
            End If

        End If

    Else
        btnActive.Enabled = True
        btnActive.Caption = "Activate"
        btnActive.BackColor = RGB(0, 120, 215)
        lblActivationDate.Caption = "Activated : -"
        lblExpiry.Caption = "Expires   : -"
        lblDaysLeft.Caption = ""
    End If

End Sub


'==============================
' ACTIVATE BUTTON
'==============================
Private Sub btnActive_Click()

    ' -- L?y license key t? user --
    Dim k As String
    k = Trim$(InputBox("Enter License Key:", "Activation"))

    ' -- User b?m Cancel ho?c d? tr?ng --
    If Len(k) = 0 Then Exit Sub

    If TryActivate(k) Then

        ShowMsgCenter "Activated successfully!"
        UpdateLicenseUI

        ' -- Kích ho?t thành công: ch?y toàn b? init h? th?ng --
        ' EnsureCPMSSystem ki?m tra IsActivated() tru?c tiên;
        ' vì v?a activate xong nó s? = True và di th?ng vào
        ' EnsureSystemFolders > EnsureDataTemplate > v.v.
        EnsureCPMSSystem

    Else

        ShowMsgCenter "Invalid license key!"
        UpdateLicenseUI

    End If

End Sub


'==============================
' T?O QR VÀO FRAME
'==============================
Private Sub TaoQR_VaoFrame()

    Dim bank As String, acc As String, nameAcc As String
    Dim amount As String, note As String
    Dim URL As String
    Dim filePath As String

    bank = "MB"
    acc = "0904800002"
    nameAcc = "MAI VAN HIEU"
    amount = ""
    note = GetMachineID()

    ' QR ONLY (g?n)
    URL = "https://img.vietqr.io/image/" & _
          bank & "-" & acc & "-qr_only.png" & _
          "?amount=" & amount & _
          "&addInfo=" & WorksheetFunction.EncodeURL(note) & _
          "&accountName=" & WorksheetFunction.EncodeURL(nameAcc)

    filePath = Environ$("TEMP") & "\qr_frame.png"

    URLDownloadToFile 0, URL, filePath, 0, 0

    ' ===== LOAD VÀO IMAGE TRONG FRAME =====
    Me.imgQR.Picture = LoadPicture(filePath)

End Sub


'=========================================
' HIDDEN FEATURE: Ctrl+Alt+Click vào lblCopyright
' ? Hi?n/?n các sheet xlSheetVeryHidden
'=========================================
Private Sub lblCopyright_Click()

    ' Ch? kích ho?t khi gi? Ctrl + Alt d?ng th?i
    If GetKeyState(vbKeyControl) >= 0 Then Exit Sub
    If GetKeyState(vbKeyMenu) >= 0 Then Exit Sub

    ShowVeryHiddenSheets

End Sub


Private Sub ShowVeryHiddenSheets()

    ' [FIX] Dung GetUserWorkbook() thay vi defaultWorkbook (co the chua duoc set)
    '       GetUserWorkbook() xu ly dung ca 2 che do: XLAM add-in va direct-run
    Dim wb As Workbook
    Set wb = GetUserWorkbook()

    If wb Is Nothing Then
        ShowMsgCenter "No user workbook detected."
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim hiddenSheets() As String
    Dim hiddenTypes() As String
    Dim count As Long
    Dim i As Long
    Dim choice As String
    Dim msg As String

    count = 0
    ReDim hiddenSheets(1 To wb.sheets.count)
    ReDim hiddenTypes(1 To wb.sheets.count)

    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetHidden Or ws.Visible = xlSheetVeryHidden Then
            count = count + 1
            hiddenSheets(count) = ws.Name
            If ws.Visible = xlSheetVeryHidden Then
                hiddenTypes(count) = "[VH]"
            Else
                hiddenTypes(count) = "[H] "
            End If
        End If
    Next ws

    If count = 0 Then
        ShowMsgCenter "No hidden sheets found in " & wb.Name
        Exit Sub
    End If

    msg = "Hidden Sheets in [" & wb.Name & "] (" & count & "):" & vbCrLf & _
          "[H] = Hidden   [VH] = Very Hidden" & vbCrLf & vbCrLf
    For i = 1 To count
        msg = msg & i & ". " & hiddenTypes(i) & " " & hiddenSheets(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "Nhap so de unhide (0 = tat ca, trong = thoat):"

    choice = InputBox(msg, "Sheet Manager [Hidden]")

    If choice = "" Then Exit Sub

    If choice = "0" Then
        For i = 1 To count
            wb.Worksheets(hiddenSheets(i)).Visible = xlSheetVisible
        Next i
        ShowMsgCenter count & " sheet(s) unhidden."

    ElseIf IsNumeric(choice) Then
        Dim idx As Long
        idx = CLng(choice)
        If idx >= 1 And idx <= count Then
            wb.Worksheets(hiddenSheets(idx)).Visible = xlSheetVisible
            ShowMsgCenter "Sheet """ & hiddenSheets(idx) & """ is now visible."
        Else
            ShowMsgCenter "Invalid selection."
        End If
    Else
        On Error Resume Next
        wb.Worksheets(choice).Visible = xlSheetVisible
        If Err.Number = 0 Then
            ShowMsgCenter "Sheet """ & choice & """ is now visible."
        Else
            ShowMsgCenter "Sheet not found: " & choice
        End If
        On Error GoTo 0
    End If

End Sub


Public Sub VeryHideSheet(sheetName As String)
    Dim wb As Workbook
    Set wb = GetUserWorkbook()
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Worksheets(sheetName).Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub




