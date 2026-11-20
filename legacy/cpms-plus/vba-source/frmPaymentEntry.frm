Attribute VB_Name = "frmPaymentEntry"
Attribute VB_Base = "0{A17FEC92-2144-4343-9C5B-FD6237DCEF74}{8CD9A931-A21F-49CE-AB7F-5A9D1EF1D133}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'frmPaymentEhtry
Option Explicit

Private editingRow As Long   ' DÒNG ÐANG S?A
Private isEditMode As Boolean
Private lastPaymentIndex As Long
Private suppressPaymentSelect As Boolean
Private suppressContractChange As Boolean
Public selectedLedgerRow As Long



' ===== CONTRACT INFO (LABEL) =====
Public Property Let contractID(v As String)
    lblContractID.Caption = v
End Property

Public Property Get contractID() As String
    contractID = lblContractID.Caption
End Property

Public Property Let contractNo(v As String)
    'lblContractNO.Caption = v
End Property

Public Property Get contractNo() As String
    'contractNo = lblContractNO.Caption
End Property

Private Function ValidateAddPayment() As Boolean

    If GetRealText(txtPaymentNo) = "" Then
        MsgBox "Payment No is required.", vbExclamation
        txtPaymentNo.SetFocus
        Exit Function
    End If

    If GetRealText(txtPaymentRef) = "" Then
        MsgBox "Payment reference is required.", vbExclamation
        txtPaymentRef.SetFocus
        Exit Function
    End If

    If GetSafeNumber(txtPaymentRequest.value) <= 0 Then
        MsgBox "Payment request amount must be greater than 0.", vbExclamation
        txtPaymentRequest.SetFocus
        Exit Function
    End If

    Dim dtStr As String
    dtStr = Trim(GetRealText(txtPaymentDate))
    If dtStr = "" Then
        MsgBox "Payment date is required (dd/mm/yyyy).", vbExclamation
        txtPaymentDate.SetFocus
        Exit Function
    End If
    If Not IsValidDDMMYYYY(dtStr) Then
        MsgBox "Payment date must be in dd/mm/yyyy format." & vbCrLf & _
               "Example: 31/12/2024", vbExclamation
        txtPaymentDate.SetFocus
        Exit Function
    End If

    ValidateAddPayment = True

End Function


Private Sub btnClose_Click()

    Call ReleaseMemory

    Unload Me

End Sub

Private Sub btnBack_Click()

    ' Luu contractID TRU?C — ReleaseMemory s? xóa lblContractID.Caption t?i line 1468
    Dim savedCID As String
    savedCID = Trim$(lblContractID.Caption)

    Call ReleaseMemory
    Unload Me

    ' vbModeless: b?t bu?c, thi?u s? b? thoát form khi double-click
    frmPaymentSummary.Show vbModeless

    ' N?p data — form tr?ng vì chua qua InitFast khi vào th?ng t? btnPaymentEntryTab
    If Not gPaymentLoaded Then
        frmPaymentSummary.InitFast
    Else
        frmPaymentSummary.LoadData
    End If

    ' Nh?y t?i dúng contract v?a làm vi?c
    If savedCID <> "" Then
        frmPaymentSummary.SelectContract savedCID
    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        'Cancel = True
        Call ReleaseMemory

        Unload Me
    End If
    'ApplyTheme Me
    UnregisterForm Me
End Sub


Private Sub btnExportLedger_Click()

    Call ExportSelectedLedgerVertical

    'ShowMsgCenter "Ledger exported successfully.", vbInformation

End Sub


Private Sub btnReset_Click()

    ResetPaymentSelection

    ClearPaymentInput

    chkEditPayment.value = False

    lblLedgerID.Caption = ""

    If Trim$(lblContractID.Caption) <> "" Then

        LoadPaymentHistoryFast _
            lblContractID.Caption

        LoadLedgerSummaryFast _
            lblContractID.Caption

    End If

End Sub

Private Sub btnSave_Click()

    On Error GoTo ERR_HANDLER

    Dim ws As Worksheet
    Dim r As Long
    Dim rate As Double

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set ws = EnsurePaymentLedger

    '========================
    ' VALIDATE
    '========================
    If Not ValidateAddPayment Then GoTo SAFE_EXIT

    r = FindInsertRowForPayment(ws, lblContractID.Caption)

    If r <= 0 Then
        ShowMsgCenter _
            "Cannot find ledger position for this contract.", _
            vbCritical
        GoTo SAFE_EXIT
    End If

    '========================
    ' INSERT ROW
    '========================
    ws.rows(r).Insert Shift:=xlDown

    FormatPaymentRowSingle ws, r

    '========================
    ' BASIC INFO
    '========================
    ws.Cells(r, "A").value = lblContractID.Caption

    ws.Cells(r, "B").value = _
        "LG-" & _
        Format(Now, "yyyymmddhhnnss") & _
        "-" & _
        Format(Timer * 1000, "000")

    ws.Cells(r, "C").value = lblContractID.Caption

    ws.Cells(r, "E").value = NextItemIndex(lblContractID.Caption)

    '========================
    ' PAYMENT INFO
    '========================
    ws.Cells(r, "O").value = GetRealText(txtPaymentRef)
    ws.Cells(r, "P").value = GetRealText(txtPaymentNo)

    '========================
    ' EXCHANGE RATE
    '========================
    ws.Cells(r, "G").value = GetSafeNumber(txtExchangeRate.value)
    rate = val(ws.Cells(r, "G").value)
    If rate = 0 Then rate = 1
    ws.Cells(r, "G").value = rate

    '========================
    ' AMOUNT
    '========================
    ws.Cells(r, "J").value = GetSafeNumber(txtActualVolume.value)
    ws.Cells(r, "K").value = GetSafeNumber(txtAcceptedVolume.value)
    ws.Cells(r, "L").value = GetSafeNumber(txtPaymentRequest.value)
    ws.Cells(r, "M").value = GetSafeNumber(txtPaymentActual.value)
    ws.Cells(r, "N").value = ws.Cells(r, "L").value - ws.Cells(r, "M").value

    '========================
    ' PAYMENT DATE
    '========================
    If Trim(GetRealText(txtPaymentDate)) <> "" Then
        ws.Cells(r, "Q").value = ParseVNDate(GetRealText(txtPaymentDate))
        ws.Cells(r, "Q").NumberFormat = "dd/mm/yyyy"
    Else
        ws.Cells(r, "Q").ClearContents
    End If

    '========================
    ' INVOICE
    '========================
    ws.Cells(r, "R").value = GetSafeNumber(txtInvoiceValue.value)
    ws.Cells(r, "S").value = GetRealText(txtInvoiceRef)

    If Trim(GetRealText(txtInvoiceDate)) <> "" Then
        ws.Cells(r, "T").value = ParseVNDate(GetRealText(txtInvoiceDate))
        ws.Cells(r, "T").NumberFormat = "dd/mm/yyyy"
    Else
        ws.Cells(r, "T").ClearContents
    End If

    '========================
    ' NOTE
    '========================
    ws.Cells(r, "U").value = GetRealText(txtNote)
    ws.Cells(r, "V").value = Now

    '========================
    ' VND FORMULA
    '========================
    ws.Cells(r, "W").Formula = "=IFERROR($J" & r & "*$G" & r & ",0)"
    ws.Cells(r, "X").Formula = "=IFERROR($K" & r & "*$G" & r & ",0)"
    ws.Cells(r, "Y").Formula = "=IFERROR($L" & r & "*$G" & r & ",0)"
    ws.Cells(r, "Z").Formula = "=IFERROR($M" & r & "*$G" & r & ",0)"

    '========================
    ' *** FIX: RESTORE CALCULATION TRU?C KHI BUILD CACHE ***
    ' Ph?i tính toán xong các formula W-Z tru?c khi
    ' RefreshAllCache d?c l?i sheet, n?u không cache s? luu
    ' giá tr? 0 cho các c?t VND và form hi?n th? sai.
    '========================
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate   ' Force tính ngay các formula v?a ghi

    '========================
    ' REFRESH CACHE
    '========================
    RefreshAllCache

    '========================
    ' REFRESH OPEN FORMS
    '========================
    On Error Resume Next

    'If frmPaymentEntry.Visible Then
    '    frmPaymentEntry.Show
    'End If

    'If frmContractSearch.Visible Then
    '    frmContractSearch.LoadContractList
    'End If

    On Error GoTo ERR_HANDLER

    '========================
    ' REFRESH UI
    '========================
    LoadLedgerSummaryFast lblContractID.Caption
    LoadPaymentHistoryFast lblContractID.Caption

    ShowMsgCenter "Payment added successfully.", vbInformation
    'If frmPaymentEntry.Visible Then
        frmPaymentEntry.Show vbModeless
    'End If
    ClearPaymentInput
    InitPlaceholders

SAFE_EXIT:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Exit Sub

ERR_HANDLER:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "Error " & Err.Number & vbCrLf & _
        Err.Description & vbCrLf & _
        "Row: " & r, _
        vbCritical

End Sub



Public Function ParseVNDate(v As String) As Variant
    Dim s As String
    s = Trim(v)

    If s = "" Then
        ParseVNDate = Empty
        Exit Function
    End If

    ' Guard: neu khong co "/" thi khong phai ngay thang
    If InStr(s, "/") = 0 Then
        ParseVNDate = Empty
        Exit Function
    End If

    Dim arr() As String
    arr = Split(s, "/")

    If UBound(arr) <> 2 Then
        Err.Raise 1001, , "Invalid date format: " & s
    End If

    ParseVNDate = DateSerial(CLng(arr(2)), CLng(arr(1)), CLng(arr(0)))
End Function



Private Sub LoadPayeeFromContract(f As Object)

    cboPayee.Clear

    If Trim(f.txtCompany1.value) <> "" Then cboPayee.AddItem f.txtCompany1.value
    If Trim(f.txtCompany2.value) <> "" Then cboPayee.AddItem f.txtCompany2.value
    If Trim(f.txtCompany3.value) <> "" Then cboPayee.AddItem f.txtCompany3.value

End Sub


Private Sub ClearPaymentInput()

    txtPaymentRef.value = ""
    txtPaymentRequest.value = ""
    txtPaymentActual.value = ""
    txtPaymentDate.value = ""
    txtNote.value = ""
    'cboPayee.ListIndex = -1

    isEditMode = False
    editingRow = 0
    'btnSave.Caption = "Add Payment"
    chkEditPayment.value = False

End Sub




Private Sub btnRefreshTable_Click()
    Call FormatTableWithGroup(EnsurePaymentLedger)
    Call FormatTableWithGroup(EnsureContractDB)
    'ShowMsgCenter "Completed.", vbExclamation
End Sub

Private Sub btnUpdatePayment_Click()

    Dim ws As Worksheet
    Dim r As Long
    Dim rate As Double

    If editingRow <= 0 Then
        ShowMsgCenter "Please select a payment to update.", vbExclamation
        Exit Sub
    End If

    If Not chkEditPayment.value Then
        ShowMsgCenter "Please enable Edit mode.", vbExclamation
        Exit Sub
    End If

    Set ws = EnsurePaymentLedger
    'r = editingRow
    r = FindLedgerRowByID( _
        ws, _
        lblLedgerID.Caption)

    If r <= 0 Then
    
        ShowMsgCenter _
            "Cannot locate payment row.", _
            vbCritical
    
        Exit Sub
    
    End If
    'r = LedgerSheetRow(editingRow)

    ws.Cells(r, "G").value = GetSafeNumber(txtExchangeRate.value)
    rate = val(ws.Cells(r, "G").value)
    If rate = 0 Then rate = 1
    ws.Cells(r, "G").value = rate

    ws.Cells(r, "J").value = GetSafeNumber(txtActualVolume.value)
    ws.Cells(r, "K").value = GetSafeNumber(txtAcceptedVolume.value)
    ws.Cells(r, "L").value = GetSafeNumber(txtPaymentRequest.value)
    ws.Cells(r, "M").value = GetSafeNumber(txtPaymentActual.value)

    ws.Cells(r, "N").value = ws.Cells(r, "L").value - ws.Cells(r, "M").value

    ws.Cells(r, "F").value = GetRealText(txtPaymentRef)
    ws.Cells(r, "O").value = GetRealText(txtPaymentRef)
    ws.Cells(r, "P").value = GetRealText(txtPaymentNo)
    Dim updDtStr As String
    updDtStr = Trim(GetRealText(txtPaymentDate))
    If updDtStr <> "" Then
        If Not IsValidDDMMYYYY(updDtStr) Then
            ShowMsgCenter "Payment date must be in dd/mm/yyyy format." & vbCrLf & _
                          "Example: 31/12/2024", vbExclamation
            txtPaymentDate.SetFocus
            Exit Sub
        End If
        ws.Cells(r, "Q").value = GetSafeDate(updDtStr)
    Else
        ws.Cells(r, "Q").ClearContents
    End If
    
    ws.Cells(r, "Q").NumberFormat = "dd/mm/yyyy"

    ws.Cells(r, "R").value = GetSafeNumber(txtInvoiceValue.value)
    ws.Cells(r, "S").value = GetRealText(txtInvoiceRef)
    If Trim(txtInvoiceDate.value) <> "" Then

        ws.Cells(r, "T").value = _
            GetSafeDate(txtInvoiceDate.value)
    
    Else
    
        ws.Cells(r, "T").ClearContents
    
    End If
    
    ws.Cells(r, "Q").NumberFormat = "dd/mm/yyyy"

    'ws.Cells(r, "T").value = GetSafeDate(txtInvoiceDate.value)
    ws.Cells(r, "U").value = GetRealText(txtNote)
    
    '========================
    ' REFRESH ALL CACHE
    '========================
    RefreshAllCache
    'gPaymentEntryLoaded = False
    'gPaymentLoaded = False
    'gContractLoaded = False
    
    'BuildPaymentEntryCache
    'BuildPaymentCache
    'BuildGlobalContractCache
    
    '========================
    ' REFRESH OPEN FORMS
    '========================
    On Error Resume Next
    
    If frmPaymentSummary.Visible Then
        frmPaymentSummary.LoadData
    End If
    
    If frmContractSearch.Visible Then
        frmContractSearch.LoadContractList
    End If
    
    On Error GoTo 0
    
    '========================
    ' REFRESH UI
    '========================
    LoadLedgerSummaryFast _
        lblContractID.Caption
    
    LoadPaymentHistoryFast _
        lblContractID.Caption
    
    ShowMsgCenter _
        "Payment updated successfully.", _
        vbInformation
    
    '========================
    ' RESET STATE
    '========================
    chkEditPayment.value = False
    
    btnUpdatePayment.Enabled = False
    btnDeletePayment.Enabled = False
    
    editingRow = 0

End Sub

Private Sub chkEditPayment_Click()
    btnUpdatePayment.Enabled = chkEditPayment.value
    btnDeletePayment.Enabled = chkEditPayment.value
End Sub

Private Sub txtexchangerate_Enter(): TextBox_Enter txtExchangeRate: End Sub
Private Sub txtexchangerate_Exit(ByVal Cancel As MSForms.ReturnBoolean): TextBox_Exit txtExchangeRate: End Sub

Private Sub txtPaymentNo_Enter(): TextBox_Enter txtPaymentNo: End Sub
Private Sub txtPaymentNo_Exit(ByVal Cancel As MSForms.ReturnBoolean): TextBox_Exit txtPaymentNo: End Sub

Private Sub txtPaymentRef_Enter(): TextBox_Enter txtPaymentRef: End Sub
Private Sub txtPaymentRef_Exit(ByVal Cancel As MSForms.ReturnBoolean): TextBox_Exit txtPaymentRef: End Sub

Private Sub txtPaymentRequest_Enter(): TextBox_Enter txtPaymentRequest: End Sub
Private Sub txtPaymentRequest_Exit(ByVal Cancel As MSForms.ReturnBoolean): TextBox_Exit txtPaymentRequest: End Sub

Private Sub txtPaymentActual_Enter(): TextBox_Enter txtPaymentActual: End Sub
Private Sub txtPaymentActual_Exit(ByVal Cancel As MSForms.ReturnBoolean): TextBox_Exit txtPaymentActual: End Sub

Private Sub txtPaymentDate_Enter(): TextBox_Enter txtPaymentDate: End Sub
Private Sub txtPaymentDate_Exit(ByVal Cancel As MSForms.ReturnBoolean): TextBox_Exit txtPaymentDate: End Sub

Private Sub txtInvoiceValue_Enter(): TextBox_Enter txtInvoiceValue: End Sub
Private Sub txtInvoiceValue_Exit(ByVal Cancel As MSForms.ReturnBoolean): TextBox_Exit txtInvoiceValue: End Sub

Private Sub txtInvoiceRef_Enter(): TextBox_Enter txtInvoiceRef: End Sub
Private Sub txtInvoiceRef_Exit(ByVal Cancel As MSForms.ReturnBoolean): TextBox_Exit txtInvoiceRef: End Sub

Private Sub txtInvoiceDate_Enter(): TextBox_Enter txtInvoiceDate: End Sub
Private Sub txtInvoiceDate_Exit(ByVal Cancel As MSForms.ReturnBoolean): TextBox_Exit txtInvoiceDate: End Sub

Private Sub txtNote_Enter(): TextBox_Enter txtNote: End Sub
Private Sub txtNote_Exit(ByVal Cancel As MSForms.ReturnBoolean): TextBox_Exit txtNote: End Sub
Private Sub txtActualVolume_Enter()
    TextBox_Enter txtActualVolume
End Sub

Private Sub txtActualVolume_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    TextBox_Exit txtActualVolume
End Sub

Private Sub txtAcceptedVolume_Enter()
    TextBox_Enter txtAcceptedVolume
End Sub

Private Sub txtAcceptedVolume_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    TextBox_Exit txtAcceptedVolume
End Sub


Private Function GetRealText(tb As MSForms.TextBox) As String
    If tb.value = tb.Tag Then
        GetRealText = ""
    Else
        GetRealText = Trim(tb.value)
    End If
End Function

'====================================
' FORM INIT
'====================================
Private Sub UserForm_Initialize()

    If Not gPaymentEntryLoaded Then
        BuildPaymentEntryCache
    End If

    editingRow = 0

    InitPlaceholders

    chkEditPayment.value = False

    btnDeletePayment.Enabled = False
    btnUpdatePayment.Enabled = False

    With lstPayment

        .ColumnCount = 9

        .ColumnWidths = _
        "0 pt;25 pt;150 pt;80 pt;80 pt;80 pt;80 pt;80 pt;50 pt"

    End With
    LoadContractNoToComboFast
    BuildContractComboAll
    
    ApplyTheme Me
    'UnregisterForm Me
End Sub

Private Sub BuildContractComboAll()

    Dim i As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim contractID As String
    Dim contractNo As String

    Dim dictAdded As Object

    Set dictAdded = CreateObject("Scripting.Dictionary")

    suppressContractChange = True

    cboContractNo.Clear

    '========================
    ' NORMAL CONTRACT
    '========================
    For i = 1 To UBound(gArrContract, 1)

        contractID = _
            Trim$(CStr(gArrContract(i, 1)))

        contractNo = _
            Trim$(CStr(gArrContract(i, 2)))

        If contractID <> "" Then

            cboContractNo.AddItem contractNo

            cboContractNo.List( _
                cboContractNo.ListCount - 1, 1) = contractID

            dictAdded(contractID) = True

        End If

    Next i

    '========================
    ' CPB CONTRACT
    '========================
    Set ws = EnsurePaymentLedger

    lastRow = _
        ws.Cells(ws.rows.count, "C").End(xlUp).Row

    For r = 2 To lastRow

        '====================
        ' HEADER ROW ONLY
        '====================
        If Trim$(ws.Cells(r, "B").value) = "" Then

            contractID = _
                Trim$(ws.Cells(r, "C").value)

            If Left$(contractID, 6) = "B-CPB-" Then

                If Not dictAdded.Exists(contractID) Then

                    contractNo = _
                        Trim$(ws.Cells(r, "D").value)

                    cboContractNo.AddItem contractNo

                    cboContractNo.List( _
                        cboContractNo.ListCount - 1, 1) = contractID

                    dictAdded(contractID) = True

                End If

            End If

        End If

    Next r

    suppressContractChange = False

End Sub

Public Sub InitFromContractFast( _
    contractID As String)

    Dim i As Long

    LoadUnifiedSourceFast contractID

    LoadPaymentHistoryFast contractID

    LoadLedgerSummaryFast contractID

    '========================
    ' SELECT COMBO
    '========================
    suppressContractChange = True
    
    For i = 0 To cboContractNo.ListCount - 1
    
        If cboContractNo.List(i, 1) = contractID Then
    
            cboContractNo.ListIndex = i
            Exit For
    
        End If
    
    Next i
    
    suppressContractChange = False

End Sub

Private Sub LoadUnifiedSourceFast( _
    contractID As String)

    Dim r As Long

    '========================
    ' NORMAL CONTRACT
    '========================
    If gDictContract.Exists(contractID) Then

        r = gDictContract(contractID)

        lblContractID.Caption = _
            contractID

        lblLedgerID.Caption = ""

        lblCompany2.Caption = _
            gArrContract(r, 8)

        lblContractValue.Caption = _
            FastVN(gArrContract(r, 15))

        lblItemName.Caption = _
            gArrContract(r, 11)

    '========================
    ' CPB CONTRACT
    '========================
    ElseIf Left$(contractID, 6) = "B-CPB-" Then

        Dim ws As Worksheet
        Dim lastRow As Long

        Set ws = EnsurePaymentLedger

        lastRow = _
            ws.Cells(ws.rows.count, "C").End(xlUp).Row

        For r = 2 To lastRow

            '====================
            ' HEADER ROW ONLY
            '====================
            If Trim$(ws.Cells(r, "C").value) = contractID _
            And Trim$(ws.Cells(r, "B").value) = "" Then

                lblContractID.Caption = _
                    contractID

                lblLedgerID.Caption = ""

                lblCompany2.Caption = _
                    "Internal Cost"

                lblContractValue.Caption = _
                    FastVN(0)

                lblItemName.Caption = _
                    Trim$(ws.Cells(r, "F").value)

                Exit For

            End If

        Next r

    End If

    '========================
    ' SELECT COMBO
    '========================
    suppressContractChange = True

    For r = 0 To cboContractNo.ListCount - 1

        If Trim$(CStr( _
            cboContractNo.List(r, 1))) = _
            Trim$(contractID) Then

            cboContractNo.ListIndex = r

            Exit For

        End If

    Next r

    suppressContractChange = False

End Sub

Public Sub LoadPaymentHistoryFast( _
    contractID As String)

    Dim rows As Collection
    Dim idx As Variant

    Dim arrList() As Variant
    Dim finalArr() As Variant

    Dim total As Long
    Dim r As Long
    Dim c As Long

    lstPayment.Clear

    If Not gDictLedgerByContract.Exists(contractID) Then Exit Sub

    Set rows = gDictLedgerByContract(contractID)

    ReDim arrList(0 To rows.count + 5, 0 To 8)

    total = 0

    For Each idx In rows

        ' PAYMENT ROW ONLY
        If Trim(CStr(gArrLedger(idx, 2))) <> "" Then

            arrList(total, 0) = idx

            arrList(total, 1) = gArrLedger(idx, 16)

            arrList(total, 2) = gArrLedger(idx, 15)

            arrList(total, 3) = _
                FastVN(gArrLedger(idx, 10))

            arrList(total, 4) = _
                FastVN(gArrLedger(idx, 11))

            arrList(total, 5) = _
                FastVN(gArrLedger(idx, 12))

            arrList(total, 6) = _
                FastVN(gArrLedger(idx, 13))

            arrList(total, 7) = _
                FastVN(gArrLedger(idx, 14))

            If IsDate(gArrLedger(idx, 17)) Then

                arrList(total, 8) = _
                    Format(gArrLedger(idx, 17), _
                    "dd/mm/yyyy")

            End If

            total = total + 1

        End If

    Next idx

    If total > 0 Then
    
        ReDim finalArr(0 To total - 1, 0 To 8)
    
        For r = 0 To total - 1
    
            finalArr(r, 0) = arrList(r, 0)
            finalArr(r, 1) = arrList(r, 1)
            finalArr(r, 2) = arrList(r, 2)
            finalArr(r, 3) = arrList(r, 3)
            finalArr(r, 4) = arrList(r, 4)
            finalArr(r, 5) = arrList(r, 5)
            finalArr(r, 6) = arrList(r, 6)
            finalArr(r, 7) = arrList(r, 7)
            finalArr(r, 8) = arrList(r, 8)
    
        Next r
    
        lstPayment.List = finalArr
    
    End If

    ResetPaymentSelection

End Sub

Public Sub LoadLedgerSummaryFast( _
    contractID As String)

    Dim rows As Collection
    Dim idx As Variant

    lblSumJ.Caption = "0"
    lblSumK.Caption = "0"
    lblSumL.Caption = "0"
    lblSumM.Caption = "0"
    lblSumN.Caption = "0"

    If Not gDictLedgerByContract.Exists(contractID) Then Exit Sub

    Set rows = gDictLedgerByContract(contractID)

    For Each idx In rows

        If Trim(CStr(gArrLedger(idx, 2))) = "" _
        And Trim(CStr(gArrLedger(idx, 5))) = "" Then

            lblSumJ.Caption = _
                FastVN(gArrLedger(idx, 10))

            lblSumK.Caption = _
                FastVN(gArrLedger(idx, 11))

            lblSumL.Caption = _
                FastVN(gArrLedger(idx, 12))

            lblSumM.Caption = _
                FastVN(gArrLedger(idx, 13))

            lblSumN.Caption = _
                FastVN(gArrLedger(idx, 14))

            Exit Sub

        End If

    Next idx

End Sub





Private Sub cboContractNo_Change()

    If suppressContractChange Then Exit Sub

    If cboContractNo.ListIndex < 0 Then Exit Sub

    Dim contractID As String

    contractID = _
        cboContractNo.Column(1)

    If Trim(contractID) = "" Then Exit Sub

    '========================
    ' FAST CACHE LOAD
    '========================
    InitFromContractFast contractID

End Sub






Private Sub InitPlaceholders()
    Dim c As Control
    For Each c In Me.Controls
        If TypeName(c) = "TextBox" Then
            If Trim(c.Tag) <> "" Then
                c.value = c.Tag
                c.ForeColor = RGB(0, 0, 0)
            End If
        End If
    Next c
End Sub

Private Sub TextBox_Exit(ByVal tb As MSForms.TextBox)
    If Trim(tb.value) = "" Then
        tb.value = tb.Tag
        tb.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub TextBox_Enter(ByVal tb As MSForms.TextBox)
    If tb.value = tb.Tag Then
        tb.value = ""
        tb.ForeColor = RGB(0, 0, 0)
    End If
End Sub

'====================================
' LOAD PAYMENT HISTORY
'====================================




Private Sub FormatListHeader()

    Dim arrLabel As Variant, i As Long
    arrLabel = Array(lblH_ActualVolume, lblH_PaymentNo, lblH_PaymentRef, lblH_Amount, lblH_Date, lblH_Remaining, lblH_Request, lblSumJ, lblSumL, lblSumM, lblSumN, lblTotal, lblEmpty1, lblEmpty2)

    For i = 0 To UBound(arrLabel)
        With arrLabel(i)
            .BackColor = RGB(230, 230, 230)
            .BorderStyle = fmBorderStyleSingle
            .TextAlign = fmTextAlignCenter
            .Font.Bold = True
        End With
    Next i

End Sub





'====================================
' SELECT PAYMENT
'====================================
' ===== FORM CODE (ví d?: frmPayment) =====

Private Sub lstPayment_Click()

    If suppressPaymentSelect Then Exit Sub

    If lstPayment.ListIndex < 0 Then Exit Sub

    If lstPayment.ListIndex = lastPaymentIndex Then

        ResetPaymentSelection
        Exit Sub

    End If

    lastPaymentIndex = lstPayment.ListIndex

    editingRow = _
        CLng(lstPayment.List(lastPaymentIndex, 0))

    If editingRow <= 0 Then Exit Sub

    lblLedgerID.Caption = _
        gArrLedger(editingRow, 2)

    chkEditPayment.Enabled = True
    chkEditPayment.value = False

    selectedLedgerRow = editingRow

    LoadPaymentToForm editingRow

End Sub

Private Sub ResetPaymentSelection()

    suppressPaymentSelect = True

    lstPayment.ListIndex = -1
    lastPaymentIndex = -1
    editingRow = 0
    isEditMode = False
    'btnSave.Caption = "Add Payment"
    chkEditPayment.Enabled = True

    ClearPaymentInput
    InitPlaceholders

    suppressPaymentSelect = False

End Sub

Private Sub LoadPaymentToForm( _
    ByVal r As Long)

    txtExchangeRate.value = gArrLedger(r, 7)

    txtPaymentNo.value = gArrLedger(r, 16)

    txtPaymentRef.value = gArrLedger(r, 15)

    If IsDate(gArrLedger(r, 17)) Then

        txtPaymentDate.value = _
            Format(gArrLedger(r, 17), _
            "dd/mm/yyyy")

    Else

        txtPaymentDate.value = ""

    End If

    SetNumberToTextBox _
        txtActualVolume, gArrLedger(r, 10)

    SetNumberToTextBox _
        txtAcceptedVolume, gArrLedger(r, 11)

    SetNumberToTextBox _
        txtPaymentRequest, gArrLedger(r, 12)

    SetNumberToTextBox _
        txtPaymentActual, gArrLedger(r, 13)

    SetNumberToTextBox _
        txtInvoiceValue, gArrLedger(r, 18)

    txtInvoiceRef.value = _
        gArrLedger(r, 19)

    If IsDate(gArrLedger(r, 20)) Then

        txtInvoiceDate.value = _
            Format(gArrLedger(r, 20), _
            "dd/mm/yyyy")

    Else

        txtInvoiceDate.value = ""

    End If

    txtNote.value = _
        gArrLedger(r, 21)

End Sub


Public Sub SetNumberToTextBox(tb As MSForms.TextBox, v As Variant)

    On Error GoTo SAFE_EXIT

    tb.Enabled = True
    tb.Locked = False

    If IsNumeric(v) Then

        If CDbl(v) <> 0 Then

            tb.value = FastVN(v)

        Else

            tb.value = ""

        End If

    Else

        tb.value = ""

    End If

    tb.ForeColor = RGB(0, 0, 0)
    tb.BackColor = RGB(255, 255, 255)

    Exit Sub

SAFE_EXIT:

    tb.value = ""
    tb.ForeColor = RGB(0, 0, 0)
    tb.BackColor = RGB(255, 255, 255)

End Sub

Public Function FindLedgerRowByID( _
    ws As Worksheet, _
    ledgerID As String) As Long

    Dim lastRow As Long
    Dim r As Long

    If Trim$(ledgerID) = "" Then Exit Function

    'lastRow = _
        ws.Cells(ws.rows.Count, "B").End(xlUp).Row
    lastRow = ws.Cells.Find( _
        What:="*", _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious).Row
    For r = 2 To lastRow

        If Trim$(ws.Cells(r, "B").value) = _
           Trim$(ledgerID) Then

            FindLedgerRowByID = r
            Exit Function

        End If

    Next r

End Function

'====================================
' DELETE PAYMENT
'====================================
Private Sub btnDeletePayment_Click()

    Dim ws As Worksheet
    Dim r As Long
    Dim contractID As String

    If editingRow <= 0 Then

        ShowMsgCenter _
            "No payment selected.", _
            vbExclamation

        Exit Sub

    End If

    Set ws = EnsurePaymentLedger

    '========================
    ' FIND REAL SHEET ROW
    '========================
    r = FindLedgerRowByID( _
            ws, _
            lblLedgerID.Caption)

    If r <= 0 Then

        ShowMsgCenter _
            "Cannot locate payment row.", _
            vbCritical

        Exit Sub

    End If

    '========================
    ' SAFETY CHECK
    '========================
    If Not CanDeletePayment(ws, r) Then
        Exit Sub
    End If

    contractID = _
        Trim$(ws.Cells(r, "C").value)

    If MsgBox( _
        "Delete this payment?", _
        vbYesNo + vbCritical) <> vbYes Then

        Exit Sub

    End If

    On Error GoTo SAFE_EXIT

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '========================
    ' DELETE REAL ROW
    '========================
    ws.rows(r).Delete
    Application.Calculate
    DoEvents
    '========================
    ' RECALCULATE SUMMARY
    '========================
    'UpdateLedgerSummary contractID

    '========================
    ' REFRESH ALL CACHE
    '========================
    gPaymentEntryLoaded = False
    gPaymentLoaded = False
    gContractLoaded = False
    
    BuildPaymentEntryCache
    BuildPaymentCache
    BuildGlobalContractCache
    
    '========================
    ' REFRESH OPEN FORMS
    '========================
    On Error Resume Next
    
    If frmPaymentSummary.Visible Then
        frmPaymentSummary.LoadData
    End If
    
    If frmContractSearch.Visible Then
        frmContractSearch.LoadContractList
    End If
    
    On Error GoTo 0

    '========================
    ' RESET STATE
    '========================
    lastPaymentIndex = -1

    editingRow = 0
    selectedLedgerRow = 0

    isEditMode = False

    chkEditPayment.value = False

    btnUpdatePayment.Enabled = False
    btnDeletePayment.Enabled = False

    lblLedgerID.Caption = ""

    '========================
    ' REFRESH UI
    '========================
    If Trim$(contractID) <> "" Then

        LoadLedgerSummaryFast _
            contractID
            
        LoadPaymentHistoryFast _
            contractID


    End If

    frmPaymentSummary.LoadData

    ClearPaymentInput

    InitPlaceholders

    ShowMsgCenter _
        "Payment deleted successfully.", _
        vbInformation

SAFE_EXIT:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Private Function CanDeletePayment(ws As Worksheet, r As Long) As Boolean

    ' Không cho xóa skeleton
    If Trim(ws.Cells(r, "E").value) = "" Then
        ShowMsgCenter "Cannot delete ledger summary row.", vbCritical
        Exit Function
    End If

    CanDeletePayment = True

End Function

Private Function NextItemIndex(contractID As String) As Long
    Dim ws As Worksheet, i As Long, maxIdx As Long
    Set ws = EnsurePaymentLedger

    For i = 2 To ws.Cells(ws.rows.count, "C").End(xlUp).Row
        If ws.Cells(i, "C").value = contractID Then
            If ws.Cells(i, "E").value > maxIdx Then
                maxIdx = ws.Cells(i, "E").value
            End If
        End If
    Next i

    NextItemIndex = maxIdx + 1
End Function



Public Function SelectedContractID(cbo As MSForms.ComboBox) As String
    If cbo.ListIndex < 0 Then
        SelectedContractID = ""
    Else
        SelectedContractID = cbo.Column(1, cbo.ListIndex)
    End If
End Function

Public Sub SetComboByContractNo(cbo As MSForms.ComboBox, contractNo As String)

    Dim i As Long
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i, 0) = contractNo Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i

End Sub



Public Function ComboHasValue(cbo As MSForms.ComboBox, ByVal value As String) As Boolean
    Dim i As Long

    For i = 0 To cbo.ListCount - 1
        If Trim(cbo.List(i)) = Trim(value) Then
            ComboHasValue = True
            Exit Function
        End If
    Next i

    ComboHasValue = False
End Function



Public Sub LoadContractNoToComboFast()

    Dim arrList() As Variant
    Dim total As Long
    Dim i As Long

    cboContractNo.Clear

    cboContractNo.ColumnCount = 2
    cboContractNo.ColumnWidths = "150 pt;0 pt"

    If IsEmpty(gArrContract) Then Exit Sub

    ReDim arrList(0 To UBound(gArrContract, 1), 0 To 1)

    total = 0

    For i = 1 To UBound(gArrContract, 1)

        If Trim(CStr(gArrContract(i, 1))) <> "" Then

            arrList(total, 0) = _
                gArrContract(i, 2)

            arrList(total, 1) = _
                gArrContract(i, 1)

            total = total + 1

        End If

    Next i

    If total > 0 Then

        Dim finalArr() As Variant
        Dim r As Long
        Dim c As Long

        ReDim finalArr(0 To total - 1, 0 To 1)

        For r = 0 To total - 1
        
            finalArr(r, 0) = arrList(r, 0)
            finalArr(r, 1) = arrList(r, 1)
        
        Next r

        cboContractNo.List = finalArr

    End If

End Sub

'========================================
' RELEASE MEMORY
'========================================
Private Sub ReleaseMemory()

    On Error Resume Next

    Me.MousePointer = fmMousePointerHourGlass

    '========================
    ' CLEAR LIST
    '========================
    lstPayment.Clear

    cboContractNo.Clear

    '========================
    ' CLEAR LABEL
    '========================
    lblContractID.Caption = ""
    lblLedgerID.Caption = ""

    lblCompany2.Caption = ""
    lblContractValue.Caption = ""
    lblItemName.Caption = ""

    lblSumJ.Caption = ""
    lblSumK.Caption = ""
    lblSumL.Caption = ""
    lblSumM.Caption = ""
    lblSumN.Caption = ""

    '========================
    ' CLEAR INPUT
    '========================
    txtExchangeRate.value = ""

    txtPaymentNo.value = ""
    txtPaymentRef.value = ""

    txtPaymentDate.value = ""

    txtActualVolume.value = ""
    txtAcceptedVolume.value = ""

    txtPaymentRequest.value = ""
    txtPaymentActual.value = ""

    txtInvoiceValue.value = ""
    txtInvoiceRef.value = ""
    txtInvoiceDate.value = ""

    txtNote.value = ""

    '========================
    ' RESET STATE
    '========================
    editingRow = 0
    selectedLedgerRow = 0

    lastPaymentIndex = -1

    isEditMode = False

    suppressPaymentSelect = False
    suppressContractChange = False

    chkEditPayment.value = False

    btnDeletePayment.Enabled = False
    btnUpdatePayment.Enabled = False

    '========================
    ' OPTIONAL MEMORY RELEASE
    '========================
    Erase gArrLedger

    gPaymentEntryLoaded = False

    DoEvents

    Me.MousePointer = fmMousePointerDefault

End Sub

Private Sub UserForm_Terminate()

    On Error Resume Next

    lstPayment.Clear

    cboContractNo.Clear

End Sub

Private Function IsValidDDMMYYYY(dtStr As String) As Boolean
    Dim s As String
    s = Trim(dtStr)
    If s = "" Then Exit Function
    If InStr(s, "/") = 0 Then Exit Function

    Dim parts() As String
    parts = Split(s, "/")
    If UBound(parts) <> 2 Then Exit Function
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then Exit Function

    Dim dd As Long, mm As Long, yyyy As Long
    dd = CLng(parts(0))
    mm = CLng(parts(1))
    yyyy = CLng(parts(2))

    If dd < 1 Or dd > 31 Then Exit Function
    If mm < 1 Or mm > 12 Then Exit Function
    If yyyy < 1900 Or yyyy > 2100 Then Exit Function

    ' Ki?m tra tràn ngày (vd: 31/02/2024 ? DateSerial cho ra 02/03)
    Dim testDate As Date
    On Error Resume Next
    testDate = DateSerial(yyyy, mm, dd)
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0

    If Day(testDate) <> dd Or Month(testDate) <> mm Or Year(testDate) <> yyyy Then Exit Function

    IsValidDDMMYYYY = True
End Function
