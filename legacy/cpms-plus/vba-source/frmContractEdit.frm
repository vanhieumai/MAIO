Attribute VB_Name = "frmContractEdit"
Attribute VB_Base = "0{D3DA4FDF-C9F9-4422-A2CB-7219884A0BB1}{A2A51284-DA29-4648-AC21-7512DFA7ABC3}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'==============================================================================
' MODULE   : frmContractEdit (UserForm)
' PURPOSE  : Standalone Project / Contract form
' CONTROLS (d?t tên dúng nhu du?i dây trong UserForm Designer):
'
'   Labels         : lblContractID, lblContractNo, lblContractType,
'                    lblSignDate, lblEndDate, lblStatus,
'                    lblCompany1..3, lblProject, lblScope, lblLocation,
'                    lblContractValue, lblExchangeRate,
'                    lblPayee, lblBankAcc, lblBankName
'
'   TextBoxes      : txtContractID, txtContractNo,
'                    txtSignDate, txtEndDate,
'                    txtCompany1, txtCompany2, txtCompany3,
'                    txtProject, txtScope, txtLocation,
'                    txtContractValue,
'                    txtPayee, txtBankAcc, txtBankName
'
'   ComboBoxes     : cboContractType, cboStatus, cboExchangeRate
'
'   CheckBoxes     : chkEdit, chkAdd, chkAddAppendix
'
'   CommandButtons : btnSaveContract, btnDeleteContract, btnResetContract, btnClose,
'                    btnContract, btnDocument, btnPayment, btnExport
'
'   Frames         : fraContractInfo, fraPartners, fraPayment
'
'   (Tu? ch?n) Label: lblMsg  — hi?n th? thông báo inline
'==============================================================================
Private originalContractNo As String
Private mSearchForm As frmContractSearch

Option Explicit

' -- tr?ng thái n?i b? ------------------------------------------------------
Private isEditMode      As Boolean   ' dang ? ch? d? Edit
Private editingRow      As Long      ' dòng dang s?a trong sheet
Private baseContractID  As String    ' ID h?p d?ng cha (dùng cho Appendix)
Private isBusy          As Boolean   ' ch?n s? ki?n d? quy
Private mBaseContractType As String  ' type c?a parent khi Add Appendix ("Appendix" hay không)
Private mAppendixPrefix   As String  ' "APP-" (thêm cho contract) ho?c "PKG-" (thêm cho appendix)

Public Sub SetSearchRef(ByVal f As frmContractSearch)
    Set mSearchForm = f
End Sub

'==============================================================================
' KH?I T?O FORM
'==============================================================================
Private Sub UserForm_Initialize()

    isBusy = True

    '-- ContractType --------------------------
    With cboContractType
        .Clear
        .AddItem "Main"
        .AddItem "EPC"
        .AddItem "SUB"
        .AddItem "Appendix"
    End With

    '-- Status --------------------------------
    With cboStatus
        .Clear
        .AddItem "Active"
        .AddItem "Completed"
        .AddItem "Cancel"
        .value = "Active"
    End With

    '-- Exchange Rate -------------------------
    With cboExchangeRate
        .Clear
        .AddItem "1"
        .AddItem "USD"
        .AddItem "EUR"
        .AddItem "JPY"
        .AddItem "KRW"
        .AddItem "SGD"
        .value = "1"
    End With

    '-- Status --------------------------------
    With txtCategories
        .Clear
        .AddItem "Design & Build"
        .AddItem "Construction"
        .AddItem "Procurement"
        .AddItem "Consultancy"
        .AddItem "Design"
        .AddItem "Equipment"
        .AddItem "Service"
        .AddItem "Maintenance"
        .AddItem "Rental"
        
        .value = "Construction"
    End With
    
    '-- Tr?ng thái ban d?u --------------------
    txtContractID.Enabled = False     ' luôn khoá, do h? th?ng sinh

    With chkEdit
        .value = False
        .Enabled = False
        .Caption = "Edit existing contract"
    End With

    With chkAdd
        .value = False
        .Enabled = True
        .Caption = "Add contract"
    End With

    With chkAddAppendix
        .value = False
        .Enabled = False
        .Caption = "Add Appendix"
    End With

    btnSaveContract.Enabled = False
    btnSaveContract.Caption = "         Save"
    btnDeleteContract.Enabled = False

    Call EnableContractInputs(False)
    Dim sProjName As String, sLocation As String
    GetProjectInfo sProjName, sLocation
    lblDetailProject.Caption = IIf(sProjName <> "", sProjName, "")
    lblDetailLocation.Caption = IIf(sLocation <> "", sLocation, "")

    isBusy = False
    ApplyTheme Me
    'UnregisterForm Me
End Sub

'==============================================================================
' CHECKBOX: Edit existing contract
'==============================================================================
Private Sub chkEdit_Click()

    If isBusy Then Exit Sub

    If chkEdit.value Then
        chkAdd.value = False
        chkAddAppendix.value = False

        btnSaveContract.Enabled = True
        btnSaveContract.Caption = "        Edit & Save"
        btnDeleteContract.Enabled = True

        Call EnableContractInputs(True)
    Else
        btnSaveContract.Enabled = False
        btnDeleteContract.Enabled = False
        Call EnableContractInputs(False)
    End If

End Sub


'==============================================================================
' CHECKBOX: Add contract
'==============================================================================
Private Sub chkAdd_Click()

    If isBusy Then Exit Sub

    If chkAdd.value Then

        '-- h?i lo?i h?p d?ng ------------------
        Dim userChoice As VbMsgBoxResult
        userChoice = MsgBox( _
            "Which type of contract do you want to create?" & vbCrLf & vbCrLf & _
            "YES  —  Contract with INVESTOR (Group A)" & vbCrLf & _
            "NO   —  Contract with SUBCONTRACTOR (Group B)", _
            vbYesNoCancel + vbQuestion, "Select Contract Type")

        If userChoice = vbCancel Then
            isBusy = True
            chkAdd.value = False
            isBusy = False
            Exit Sub
        End If

        isBusy = True
        chkEdit.value = False
        chkEdit.Enabled = False
        chkAddAppendix.value = False
        chkAddAppendix.Enabled = False
        isBusy = False

        '-- setup theo lo?i --------------------
        If userChoice = vbYes Then
            Call SetupContractTypeByRole(True)          ' MAIN / EPC
            txtContractID.value = GenerateNextContractID("", "A")
        Else
            Call SetupContractTypeByRole(False)         ' SUB
            txtContractID.value = GenerateNextContractID("", "B")
        End If

        btnSaveContract.Enabled = True
        btnSaveContract.Caption = "         Add Contract"
        Call EnableContractInputs(True)

    Else
        btnSaveContract.Enabled = False
        chkEdit.Enabled = True
        chkAddAppendix.Enabled = (baseContractID <> "")
        Call EnableContractInputs(False)
    End If

End Sub


'==============================================================================
' CHECKBOX: Add Appendix
'==============================================================================
Private Sub chkAddAppendix_Click()

    If isBusy Then Exit Sub

    If chkAddAppendix.value Then

        If Trim(baseContractID) = "" Then
            MsgBox "Please search and select the parent contract first.", _
                   vbExclamation, "No Parent Contract"
            isBusy = True
            chkAddAppendix.value = False
            isBusy = False
            Exit Sub
        End If

        isBusy = True
        chkAdd.value = False
        chkEdit.value = False
        chkEdit.Enabled = False
        isBusy = False

        '-- Xác d?nh prefix theo lo?i parent ----------
        If UCase(mBaseContractType) = "APPENDIX" Then
            mAppendixPrefix = "PKG-"   ' thêm phu l?c cho appendix
        Else
            mAppendixPrefix = "APP-"   ' thêm phu l?c cho contract
        End If

        '-- set Appendix type ------------------
        cboContractType.Clear
        cboContractType.AddItem "Appendix"
        cboContractType.value = "Appendix"

        txtContractID.value = GenerateNextContractID(baseContractID)

        '-- Pre-fill prefix vào txtContractNo, user ch? c?n nh?p ph?n s? sau
        txtContractNo.value = mAppendixPrefix

        btnSaveContract.Enabled = True
        btnSaveContract.Caption = "          Add " & IIf(mAppendixPrefix = "PKG-", "Package", "Appendix")
        Call EnableContractInputs(True)
        txtContractID.Enabled = False     ' ID appendix không cho s?a

    Else
        btnSaveContract.Enabled = False
        chkEdit.Enabled = True
        Call EnableContractInputs(False)

        '-- Xóa prefix kh?i txtContractNo khi b? check
        If mAppendixPrefix <> "" Then
            If Left(txtContractNo.value, Len(mAppendixPrefix)) = mAppendixPrefix Then
                txtContractNo.value = mID(txtContractNo.value, Len(mAppendixPrefix) + 1)
            End If
        End If
        mAppendixPrefix = ""

        '-- khôi ph?c danh sách lo?i -----------
        With cboContractType
            .Clear
            .AddItem "Main"
            .AddItem "EPC"
            .AddItem "SUB"
            .AddItem "Appendix"
        End With
    End If

End Sub


'==============================================================================
' NÚT SAVE
'==============================================================================
Private Sub btnSaveContract_Click()

    On Error GoTo ERR_HANDLER

    Dim ws As Worksheet
    Set ws = EnsureContractDB

    '-- ki?m tra ch? d? ----------------------------
    If Not (chkEdit.value Or chkAdd.value Or chkAddAppendix.value) Then
        MsgBox "Please tick Edit / Add / Add Appendix before saving.", _
               vbExclamation, "No Mode Selected"
        Exit Sub
    End If

    '-- Contract No b?t bu?c -----------------------
    If Trim(txtContractNo.value) = "" Then
        ShowMsgCenter "Contract No. is required!", vbExclamation ', "Validation"
        txtContractNo.SetFocus
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim targetRow As Long
    Dim newID     As String
    Dim parentID As String

    '----------------------------------------------
    ' CH? Ð? EDIT
    '----------------------------------------------
    If chkEdit.value Then

        If editingRow <= 0 Then
            ShowMsgCenter "No contract loaded to edit.", vbExclamation ', "Edit"
            GoTo CLEAN_EXIT
        End If

        ' N?u dang edit appendix/PKG ? ch? check n?i b? parent, không check toàn sheet
        If UCase(Trim(cboContractType.value)) = "APPENDIX" Then
            ' Tìm parentID: b? ph?n "-N" cu?i cùng c?a contractID
            Dim editCID As String
            editCID = Trim(txtContractID.value)
            Dim lastDashPos As Long
            lastDashPos = 0
            Dim k As Long
            For k = Len(editCID) To 1 Step -1
                If mID(editCID, k, 1) = "-" Then lastDashPos = k: Exit For
            Next k
            Dim editParentID As String
            editParentID = IIf(lastDashPos > 0, Left(editCID, lastDashPos - 1), "")

            If editParentID <> "" Then
                If IsAppendixNoExistsForParent(ws, editParentID, txtContractNo.value, editingRow) Then
                    ShowMsgCenter "Contract No. already exists!", vbCritical
                    GoTo CLEAN_EXIT
                End If
            Else
                If IsContractNoExists(ws, txtContractNo.value, editingRow) Then
                    ShowMsgCenter "Contract No. already exists!", vbCritical
                    GoTo CLEAN_EXIT
                End If
            End If
        Else
            ' Contract thông th??ng: check toàn sheet (tr? chính nó)
            If IsContractNoExists(ws, txtContractNo.value, editingRow) Then
                ShowMsgCenter "Contract No. already exists!", vbCritical
                GoTo CLEAN_EXIT
            End If
        End If

        targetRow = editingRow
        newID = txtContractID.value

    '----------------------------------------------
    ' CH? Ð? ADD CONTRACT
    '----------------------------------------------
    ElseIf chkAdd.value Then

        If IsContractNoExists(ws, txtContractNo.value) Then
            ShowMsgCenter "Contract No. already exists!", vbExclamation ', "Duplicate"
            GoTo CLEAN_EXIT
        End If

        Dim prefix As String
        If cboContractType.value = "SUB" Then
            prefix = "B"
        Else
            prefix = "A"
        End If

        newID = GenerateNextContractID("", prefix)
        targetRow = FindInsertRowByGroup(ws, GetGroupByContractType(cboContractType.value))

        If targetRow <= 0 Then
            ShowMsgCenter "Cannot find insert position for contract.", _
                   vbCritical ', "Insert Error"
            GoTo CLEAN_EXIT
        End If

        ws.rows(targetRow).Insert Shift:=xlDown
        txtContractID.value = newID

    '----------------------------------------------
    ' CH? Ð? ADD APPENDIX
    '----------------------------------------------
    ElseIf chkAddAppendix.value Then

        If Trim(baseContractID) = "" Then
            ShowMsgCenter "Please load parent contract first.", vbExclamation
            GoTo CLEAN_EXIT
        End If

        '-- Ð?m b?o txtContractNo có prefix dúng
        If mAppendixPrefix <> "" Then
            If Left(txtContractNo.value, Len(mAppendixPrefix)) <> mAppendixPrefix Then
                txtContractNo.value = mAppendixPrefix & Trim(txtContractNo.value)
            End If
        End If

        '-- Ki?m tra trùng ch? trong n?i b? parent (không check toàn b? sheet)
        If IsAppendixNoExistsForParent(ws, baseContractID, txtContractNo.value) Then
            ShowMsgCenter "Appendix No. " & txtContractNo.value & " already exists under this " & _
                          IIf(UCase(mBaseContractType) = "APPENDIX", "appendix", "contract") & "!", vbExclamation
            GoTo CLEAN_EXIT
        End If

        newID = GenerateNextContractID(baseContractID)
        targetRow = FindInsertRow(baseContractID)

        If targetRow <= 0 Then
            ShowMsgCenter "Cannot find insert position for appendix.", _
                   vbCritical ', "Insert Error"
            GoTo CLEAN_EXIT
        End If

        ws.rows(targetRow).Insert Shift:=xlDown
        txtContractID.value = newID

    End If

    ' ==============================
    ' WRITE DATA
    ' ==============================
    'WriteContractToSheet ws, targetRow, txtContractID.value
    'Call UpdateLedgerFormulaByStatus(txtContractID.value)
    ' === SYNC LEDGER AFTER EDIT ===
    'Call SyncLedgerFromContract(txtContractID.value)

    'If chkAdd.value Then
    '    Call InsertLedgerForContract(txtContractID.value)
    
    'ElseIf chkAddAppendix.value Then
    '    Call InsertLedgerForAppendix( _
    '        parentID, _
    '        txtContractID.value)
    'End If
    
    '-- ghi d? li?u --------------------------------
    Call WriteContractToSheet(ws, targetRow, txtContractID.value)
    Call UpdateLedgerFormulaByStatus(txtContractID.value)
    ' === SYNC LEDGER AFTER EDIT ===
    Call SyncLedgerFromContract(txtContractID.value)

    If chkAdd.value Then
        Call InsertLedgerForContract(txtContractID.value)
    
    ElseIf chkAddAppendix.value Then
        ' Preserve the parent relationship for appendix ledger rows.
        parentID = baseContractID
        Call InsertLedgerForAppendix( _
            parentID, _
            txtContractID.value)
    End If
    
    'SyncLedgerFromContract (newID)
    FormatContractRowSingle ws, targetRow
    Call NormalizeSheetView(ws)
    Call NormalizeSheetView(EnsurePaymentLedger)

    ' Persist changes before another form can close the project workbook.
    Call SaveDataFile
   
    '========================
    ' REFRESH CACHE
    '========================
    RebuildContractSequence
    RefreshAllCache

    If Not mSearchForm Is Nothing Then
        mSearchForm.LoadContractList          ' Reload danh sach
        mSearchForm.SelectContract newID      ' ? THÊM: chon contract vua save/add
        mSearchForm.Show vbModeless           ' ? THÊM: dua form len truoc
    End If
    Call ResetForm
    Me.Hide
    
    ShowMsgCenter "Contract saved successfully!", vbInformation

    Call ResetForm
    Me.Hide
    
CLEAN_EXIT:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ERR_HANDLER:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ShowMsgCenter "Save failed: " & Err.Description, vbCritical ', "Error"

End Sub

Public Sub SyncLedgerFromContract(contractID As String)

    Dim wsLedger As Worksheet
    Dim wsDB As Worksheet
    Dim r As Long, lastRow As Long
    Dim dbRow As Long
    'Dim rate As Double

    If Trim(contractID) = "" Then Exit Sub

    Set wsLedger = EnsurePaymentLedger
    Set wsDB = EnsureContractDB

    ' === tìm contract trong DB ===
    dbRow = FindContractRowByID(wsDB, contractID)
    If dbRow = 0 Then Exit Sub

    lastRow = wsLedger.Cells(wsLedger.rows.count, "C").End(xlUp).Row

    ' === quét toàn b? ledger ===
    For r = 2 To lastRow

        ' ===== dúng contract family =====
        If (Trim(wsLedger.Cells(r, "C").value) = contractID) And (Trim(wsLedger.Cells(r, "B").value) = "") Then

            With wsLedger

                ' ===== CORE INFO =====
                '.Cells(r, "D").value = wsDB.Cells(dbRow, "B").value   ' ContractNo
                ' ===== ContractNo (FORCE TEXT) =====
                '.Cells(r, "D").NumberFormat = "@"
                '.Cells(r, "D").value = CStr(wsDB.Cells(dbRow, "B").value)
                Call WriteTextCell(.Cells(r, "D"), wsDB.Cells(dbRow, "B").value)
                .Cells(r, "F").value = wsDB.Cells(dbRow, "K").value   ' Scope / ItemName

                ' ===== FINANCIAL MASTER =====
                .Cells(r, "G").value = wsDB.Cells(dbRow, "M").value   ' ExchangeRate
                .Cells(r, "H").value = wsDB.Cells(dbRow, "N").value   ' EstimatedValue
                .Cells(r, "I").value = wsDB.Cells(dbRow, "O").value   ' ContractValue

                ' ===== BANKING / PAYEE =====
                '.Cells(r, "S").value = wsDB.Cells(dbRow, "S").value   ' Payee
                '.Cells(r, "T").value = wsDB.Cells(dbRow, "T").value   ' BankAcc
                '.Cells(r, "U").value = wsDB.Cells(dbRow, "U").value   ' BankName
                
                'rate = val(.Cells(r, "G").value)
                'If rate = 0 Then rate = 1
                
                ' ===== PRE-CALCULATED VND COLUMNS =====
                '.Cells(r, "W").value = val(ws.Cells(r, "J").value) * rate   ' Actual VND
                '.Cells(r, "X").value = val(ws.Cells(r, "K").value) * rate   ' Accepted VND
                '.Cells(r, "Y").value = val(ws.Cells(r, "L").value) * rate   ' Request VND
                '.Cells(r, "Z").value = val(ws.Cells(r, "M").value) * rate   ' Paid VND


            End With

        End If

    Next r

End Sub
'==============================================================================
' NÚT DELETE
'==============================================================================
Private Sub btnDeleteContract_Click()

    On Error GoTo ERR_HANDLER

    If Not isEditMode Or editingRow <= 0 Then
        ShowMsgCenter "No contract selected to delete.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = EnsureContractDB

    Dim contractNo As String
    Dim contractID As String
    contractNo = Trim(txtContractNo.value)
    contractID = Trim(txtContractID.value)

    If contractNo = "" Or contractID = "" Then
        ShowMsgCenter "Invalid contract data.", vbCritical
        Exit Sub
    End If

    '-- ki?m tra có appendix con không ------------
    If ContractFamilyHasPayment(contractID) Then
        ShowMsgCenter "This contract or its appendix already has payment records." & _
               vbCrLf & "Delete is NOT allowed.", vbCritical
        Exit Sub
    End If

    '-- xác nh?n ----------------------------------
    If MsgBox("Are you sure you want to DELETE this contract?" & vbCrLf & _
              "Contract No: " & contractNo & vbCrLf & _
              "This action CANNOT be undone.", _
              vbYesNo + vbCritical, "Confirm Delete") <> vbYes Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Call DeleteContractFamily(contractID)
    Call DeleteLedgerByContractFamily(contractID)
    '========================
    ' REFRESH CACHE
    '========================
    RebuildContractSequence
    RefreshAllCache
    If Not mSearchForm Is Nothing Then mSearchForm.LoadContractList  ' ? s?a

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ShowMsgCenter "Contract deleted successfully.", vbInformation

    Call ResetForm
    Exit Sub

ERR_HANDLER:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ShowMsgCenter "Delete failed: " & Err.Description, vbCritical

End Sub


'==============================================================================
' NÚT RESET
'==============================================================================
Private Sub btnResetContract_Click()
    Call ResetForm
End Sub


'==============================================================================
' NÚT CLOSE
'==============================================================================

'==============================================================================
' NÚT CONTRACT / DOCUMENT / PAYMENT / EXPORT
' (m? frmContractSearch ? tab tuong ?ng)
'==============================================================================
Private Sub btnContract_Click()
    Call OpenSearch(0)
End Sub



Private Sub OpenSearch(tabIndex As Long)
    Dim cid As String
    cid = Trim(txtContractID.value)
    Unload frmContractSearch
    Set frmContractSearch = Nothing
    frmContractSearch.Show vbModeless
    frmContractSearch.InitFast
    DoEvents
    If cid <> "" Then frmContractSearch.SelectContract cid
    frmContractSearch.tabMain.value = tabIndex
End Sub


'==============================================================================
' COMPANY ? C?P NH?T DANH SÁCH PAYEE  (mirrors txtCompany_Change)
'==============================================================================
Private Sub txtCompany1_Change()
    If Not isBusy Then
        Call UpdatePayeeList
        Call UpdateRepList     ' ? THÊM DÒNG NÀY
    End If
End Sub

Private Sub txtCompany2_Change()
    If Not isBusy Then
        Call UpdatePayeeList
        Call UpdateRepList     ' ? THÊM DÒNG NÀY
    End If
End Sub

Private Sub txtCompany3_Change()
    If Not isBusy Then
        Call UpdatePayeeList
        Call UpdateRepList     ' ? THÊM DÒNG NÀY
    End If
End Sub

Private Sub UpdatePayeeList()
    With txtPayee
        .Clear
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete

        ' Company1
        If Trim(txtCompany1.value) <> "" Then .AddItem txtCompany1.value
        If Trim(txtCompany2.value) <> "" Then .AddItem txtCompany2.value
        If Trim(txtCompany3.value) <> "" Then .AddItem txtCompany3.value

        ' Company2 có th? ch?a "Cty A / Cty B" (liên danh) — tách ra t?ng ph?n
        'If Trim(txtCompany2.value) <> "" Then
        '    Dim parts() As String
        '    Dim i       As Integer
        '    parts = Split(txtCompany2.value, " / ")
        '    For i = 0 To UBound(parts)
        '        If Trim(parts(i)) <> "" Then .AddItem Trim(parts(i))
        '    Next i
        'End If

        ' Company3 v?n thêm n?u dang nh?p tr?c ti?p (chua luu, chua ghép vào H)
        'If Trim(txtCompany3.value) <> "" Then
        '    If InStr(txtPayee.List(0), Trim(txtCompany3.value)) = 0 Then
        '        .AddItem Trim(txtCompany3.value)
        '    End If
        'End If
    End With
End Sub


'==============================================================================
' PAYEE ? T? ÐI?N BANK ACC + BANK NAME  (mirrors txtPayee_Change)
'==============================================================================
Private Sub txtPayee_Change()

    If isBusy Then Exit Sub

    Dim acc As String, bank As String

    If GetPayeeBankInfo(txtPayee.value, acc, bank) Then
        txtBankAcc.value = acc
        txtBankName.value = bank
    Else
        'txtBankAcc.value = ""
        'txtBankName.value = ""
    End If

End Sub

Private Function GetPayeeBankInfo(ByVal payeeName As String, _
                                  ByRef bankAcc As String, _
                                  ByRef bankName As String) As Boolean
    Dim ws       As Worksheet
    Dim lastRow  As Long
    Dim r        As Long

    Set ws = EnsureContractDB
    lastRow = ws.Cells(ws.rows.count, "S").End(xlUp).Row

    For r = 2 To lastRow
        If Trim(ws.Cells(r, "S").value) = payeeName _
           And Trim(ws.Cells(r, "T").value) <> "" Then
            bankAcc = ws.Cells(r, "T").value
            bankName = ws.Cells(r, "U").value
            GetPayeeBankInfo = True
            Exit Function
        End If
    Next r

End Function


'==============================================================================
' CONTRACT VALUE — x? lý format s? ki?u VN (###.###.###)
'==============================================================================
Private Sub txtContractValue_Enter()
    ' B? d?u ch?m khi b?t d?u nh?p
    If IsNumeric(Replace(txtContractValue.value, ".", "")) Then
        txtContractValue.value = Replace(txtContractValue.value, ".", "")
    End If
End Sub

Private Sub txtContractValue_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim raw As String
    Dim v   As Double

    raw = Trim(txtContractValue.value)
    If raw = "" Then Exit Sub

    raw = Replace(raw, ".", "")

    If Not IsNumeric(raw) Then
        MsgBox "Contract value must be numeric.", vbExclamation
        Cancel = True
        Exit Sub
    End If

    v = CDbl(raw)
    txtContractValue.value = Replace(Format(v, "#,##0"), ",", ".")
    txtContractValue.ForeColor = RGB(0, 0, 0)
End Sub


'==============================================================================
' GHI D? LI?U VÀO SHEET  (mirrors WriteContractToSheet)
'==============================================================================
Private Sub WriteContractToSheet(ws As Worksheet, _
                                  r As Long, _
                                  contractID As String)
    With ws
        .Cells(r, "A").value = contractID

        Call WriteTextCell(.Cells(r, "B"), txtContractNo.value)   ' force text

        .Cells(r, "C").value = cboContractType.value

        If IsDate(txtSignDate.value) Then
            .Cells(r, "D").value = CDate(txtSignDate.value)
            .Cells(r, "D").NumberFormat = "d/mm/yyyy"
        Else
            .Cells(r, "D").ClearContents
        End If

        If IsDate(txtEndDate.value) Then
            .Cells(r, "E").value = CDate(txtEndDate.value)
            .Cells(r, "E").NumberFormat = "d/mm/yyyy"
        Else
            .Cells(r, "E").ClearContents
        End If

        .Cells(r, "F").value = cboStatus.value

        .Cells(r, "G").value = txtCompany1.value
        .Cells(r, "H").value = txtCompany2.value
        .Cells(r, "I").value = txtCompany3.value
        
        ' Ghép Company3 vào c?t H — h? tr? thêm nhi?u bên B
        'Dim co2Merged As String
        'Dim co3Val    As String
        'co3Val = Trim(txtCompany3.value)

        'If chkEdit.value Then
        '    ' Edit mode: d?c th?ng t? c?t H d? n?i thêm, tránh m?t d? li?u cu
        '    co2Merged = Trim(.Cells(r, "H").value)
        '    If co3Val <> "" Then
        '        ' Ch? n?i n?u chua t?n t?i trong chu?i
        '        If InStr(1, co2Merged, co3Val, vbTextCompare) = 0 Then
        '            If co2Merged <> "" Then
        '                co2Merged = co2Merged & " / " & co3Val
        '            Else
        '                co2Merged = co3Val
        '            End If
        '        End If
        '    End If
        'Else
            ' Add / Appendix mode: ghép txtCompany2 + txtCompany3
        '    co2Merged = Trim(txtCompany2.value)
        '    If co3Val <> "" Then
        '        If co2Merged <> "" Then
        '            co2Merged = co2Merged & " / " & co3Val
        '        Else
        '            co2Merged = co3Val
        '        End If
        '    End If
        'End If

        '.Cells(r, "H").value = co2Merged
        '.Cells(r, "I").ClearContents
        .Cells(r, "J").value = txtRepB2.value
        .Cells(r, "K").value = txtScope.value
        '.Cells(r, "L").value = txtLocation.value
        .Cells(r, "M").value = cboExchangeRate.value
        .Cells(r, "P").value = txtRepA.value
        .Cells(r, "Q").value = txtRepB.value
        .Cells(r, "R").value = txtCategories.value
        .Cells(r, "V").value = txtContactPersonB.value

        Dim rawVal As String
        rawVal = Replace(Replace(Replace(txtContractValue.value, ".", ""), ",", ""), " ", "")
        If IsNumeric(rawVal) Then
            .Cells(r, "N").value = CDbl(rawVal)
            .Cells(r, "O").value = CDbl(rawVal)
        Else
            .Cells(r, "N").ClearContents
            .Cells(r, "O").ClearContents
        End If

        .Cells(r, "S").value = txtPayee.value
        .Cells(r, "T").value = txtBankAcc.value
        .Cells(r, "U").value = txtBankName.value

        .Cells(r, "W").value = Now   ' timestamp
    End With
End Sub


'==============================================================================
' CÀI Ð?T CONTRACT TYPE THEO LO?I  (mirrors SetupContractTypeByRole)
'==============================================================================
Public Sub SetupContractTypeByRole(isOwnerContract As Boolean)
    cboContractType.Clear
    If isOwnerContract Then
        cboContractType.AddItem "Main"
        cboContractType.AddItem "EPC"
        cboContractType.value = "EPC"
    Else
        cboContractType.AddItem "SUB"
        cboContractType.value = "SUB"
    End If
End Sub


'==============================================================================
' KI?M TRA TRÙNG S? PH? L?C TRONG N?I B? PARENT
' Ch? check các dòng là con tr?c ti?p c?a parentID (ID d?ng "parentID-N")
' Không check toàn b? sheet, tránh xung d?t v?i contract khác.
'==============================================================================
Private Function IsAppendixNoExistsForParent(ws As Worksheet, _
                                              parentID As String, _
                                              contractNo As String, _
                                              Optional excludeRow As Long = 0) As Boolean
    IsAppendixNoExistsForParent = False
    If parentID = "" Or contractNo = "" Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row

    Dim r As Long
    Dim rowPrefix As String
    rowPrefix = parentID & "-"   ' các con tr?c ti?p có ID b?t d?u = "parentID-"

    For r = 2 To lastRow
        If excludeRow > 0 And r = excludeRow Then GoTo SkipRow  ' b? qua chính nó khi edit
        Dim rowID As String
        rowID = Trim(CStr(ws.Cells(r, "A").value))
        ' Ch? xét con tr?c ti?p: ID b?t d?u b?ng prefix và không có thêm "-" sau dó
        If Left(rowID, Len(rowPrefix)) = rowPrefix Then
            Dim suffix As String
            suffix = mID(rowID, Len(rowPrefix) + 1)
            ' Con tr?c ti?p: ph?n còn l?i không ch?a "-" (ví d? A001-1 ch? không ph?i A001-1-2)
            If InStr(suffix, "-") = 0 Then
                If UCase(Trim(CStr(ws.Cells(r, "B").value))) = UCase(Trim(contractNo)) Then
                    IsAppendixNoExistsForParent = True
                    Exit Function
                End If
            End If
        End If
SkipRow:
    Next r
End Function


'==============================================================================
' B?T / T?T INPUTS  (mirrors EnableContractInputs)
' txtContractID luôn b? khoá — h? th?ng t? sinh
'==============================================================================
Private Sub EnableContractInputs(ByVal allowEdit As Boolean)
    Dim c As Control
    For Each c In Me.Controls
        Select Case TypeName(c)
            Case "TextBox"
                If c.Name = "txtContractID" Then
                    c.Enabled = False
                    c.Locked = True
                Else
                    c.Enabled = True
                    c.Locked = Not allowEdit
                End If
            Case "ComboBox"
                If c.Name <> "cboContractType" Then    ' type do chkAdd qu?n lý riêng
                    c.Enabled = allowEdit
                End If
        End Select
    Next c
End Sub


'==============================================================================
' RESET TOÀN B? FORM  (mirrors btnResetContract_Click)
'==============================================================================
Private Sub ResetForm()

    isBusy = True

    '-- xoá TextBox --------------------------
    Dim c As Control
    For Each c In Me.Controls
        If TypeName(c) = "TextBox" Then c.value = ""
    Next c

    '-- xoá ComboBox -------------------------
    For Each c In Me.Controls
        If TypeName(c) = "ComboBox" Then c.value = ""
    Next c

    '-- d?t l?i m?c d?nh ---------------------
    cboStatus.value = "Active"
    cboExchangeRate.value = "1"

    With cboContractType
        .Clear
        .AddItem "Main"
        .AddItem "EPC"
        .AddItem "SUB"
        .AddItem "Appendix"
    End With

    '-- tr?ng thái n?i b? --------------------
    isEditMode = False
    editingRow = 0
    baseContractID = ""
    mBaseContractType = ""
    mAppendixPrefix = ""

    '-- checkboxes ---------------------------
    chkEdit.value = False
    chkEdit.Enabled = False
    chkAdd.value = False
    chkAdd.Enabled = True
    chkAddAppendix.value = False
    chkAddAppendix.Enabled = False

    '-- buttons ------------------------------
    btnSaveContract.Enabled = False
    btnSaveContract.Caption = "         Save"
    btnDeleteContract.Enabled = False

    Call EnableContractInputs(False)

    isBusy = False

End Sub


'==============================================================================
' TI?N ÍCH
'==============================================================================
Private Function FormatDate(v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsDate(v) Then
        FormatDate = Format(v, "dd/mm/yyyy")
        Exit Function
    End If
    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function

    ' Hỗ trợ định dạng chuẩn ISO: YYYY-MM-DD
    If s Like "####-##-##*" Then
        Dim parts() As String
        parts = Split(Left$(s, 10), "-")
        FormatDate = parts(2) & "/" & parts(1) & "/" & parts(0)
        Exit Function
    End If

    ' Hỗ trợ định dạng DD/MM/YYYY
    If s Like "##/##/####*" Or s Like "#/##/####*" Or s Like "##/#/####*" Then
        FormatDate = Left$(s, InStr(s & " ", " ") - 1)
        Exit Function
    End If

    On Error Resume Next
    FormatDate = Format(CDate(s), "dd/mm/yyyy")
    On Error GoTo 0
End Function

Private Function FormatContractValue(v As Variant) As String
    If IsNumeric(v) And CDbl(v) <> 0 Then
        FormatContractValue = Replace(Format(CDbl(v), "#,##0"), ",", ".")
    Else
        FormatContractValue = ""
    End If
End Function

Public Sub LoadContractFromSearch(contractID As String)

    Dim ws As Worksheet
    'Dim wsLG As Worksheet
    Dim r As Variant

    Set ws = EnsureContractDB
    'Set wsLG = EnsurePaymentLedger
    r = Application.Match(contractID, ws.Columns(1), 0)

    If IsError(r) Then
        ShowMsgCenter "Contract not found.", vbCritical
        Exit Sub
    End If

    ' ====== Ð? D? LI?U ======
    baseContractID = ws.Cells(r, 1).value   ' luu ID g?c
    mBaseContractType = Trim(CStr(ws.Cells(r, 3).value))   ' luu type c?a parent

    With Me
        .txtContractID.value = baseContractID
        .txtContractNo.value = ws.Cells(r, 2).value
        .cboContractType.value = ws.Cells(r, 3).value

        .txtSignDate.value = FormatDate(ws.Cells(r, 4).value)
        .txtEndDate.value = FormatDate(ws.Cells(r, 5).value)

        .cboStatus.value = ws.Cells(r, 6).value

        .txtCompany1.value = ws.Cells(r, 7).value
        .txtCompany2.value = ws.Cells(r, 8).value
        .txtCompany3.value = ws.Cells(r, 9).value
        .txtRepB2.value = ws.Cells(r, 10).value
        .txtScope.value = ws.Cells(r, 11).value
        '.txtLocation.value = ws.Cells(r, 12).value
        .cboExchangeRate.value = ws.Cells(r, 13).value
        .txtRepA.value = ws.Cells(r, 16).value         ' c?t P
        .txtRepB.value = ws.Cells(r, 17).value         ' c?t Q
        .txtCategories.value = ws.Cells(r, 18).value   ' c?t R
        .txtContactPersonB = ws.Cells(r, 22).value
        
        If IsNumeric(ws.Cells(r, 14).value) Then
            txtContractValue.value = _
                Replace(Format(ws.Cells(r, 14).value, "#,##0"), ",", ".")
        Else
            txtContractValue.value = ""
        End If

        .txtPayee.value = ws.Cells(r, 19).value
        .txtBankAcc.value = ws.Cells(r, 20).value
        .txtBankName.value = ws.Cells(r, 21).value

    End With
    Call UpdateRepList

    ' ====== B?T EDIT MODE (ÐÚNG CH?) ======
    isEditMode = True
    editingRow = CLng(r)
    originalContractNo = ws.Cells(r, 2).value

    With chkEdit
        .Enabled = True
        .Visible = True
        .value = False
    End With
    With chkAdd
        .Enabled = True
        .Visible = True
        .value = False
    End With
        With chkAddAppendix
        .Enabled = True
        .Visible = True
        .value = False
    End With
    ' ===== VIEW MODE =====
    
    LockAllInput True
    btnSaveContract.Enabled = False

    btnSaveContract.Caption = "        Edit & Save"
    btnDeleteContract.Enabled = False
    txtContractID.Enabled = False
    EnableContractInputs False


End Sub

Private Sub LockAllInput(isLock As Boolean)
'tonew
    Dim ctrl As Control

    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl.Locked = isLock
        ElseIf TypeName(ctrl) = "ComboBox" Then
            ctrl.Enabled = Not isLock
        End If
    Next ctrl

End Sub
Private Sub UpdateRepList()

    Dim ws       As Worksheet
    Dim lastRow  As Long
    Dim r        As Long

    Set ws = EnsureContractDB
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row

    ' -- L?y tên các company hi?n t?i (dã nh?p) --
    Dim co1 As String, co2 As String, co3 As String
    co1 = Trim(txtCompany1.value)
    co2 = Trim(txtCompany2.value)
    co3 = Trim(txtCompany3.value)

    ' N?u chua nh?p company nào thì b? qua
    If co1 = "" And co2 = "" And co3 = "" Then Exit Sub

    ' -- Dùng Dictionary d? gom unique values --
    Dim dictA As Object, dictB As Object
    Set dictA = CreateObject("Scripting.Dictionary")
    Set dictB = CreateObject("Scripting.Dictionary")

    For r = 2 To lastRow
        Dim rowID As String
        rowID = Trim$(CStr(ws.Cells(r, "A").value))

        ' B? qua header group và dòng tr?ng
        If rowID = "" Or rowID = "GROUP-A" Or rowID = "GROUP-B" Then GoTo NEXT_ROW

        ' Ki?m tra company kh?p v?i b?t k? company nào dang nh?p
        Dim rowCo1 As String, rowCo2 As String, rowCo3 As String
        rowCo1 = Trim$(CStr(ws.Cells(r, "G").value))
        rowCo2 = Trim$(CStr(ws.Cells(r, "H").value))
        rowCo3 = Trim$(CStr(ws.Cells(r, "I").value))

        Dim isMatch As Boolean
        isMatch = False
        If co1 <> "" Then
            If rowCo1 = co1 Or rowCo2 = co1 Or rowCo3 = co1 Then isMatch = True
        End If
        If co2 <> "" And Not isMatch Then
            If rowCo1 = co2 Or rowCo2 = co2 Or rowCo3 = co2 Then isMatch = True
        End If
        If co3 <> "" And Not isMatch Then
            If rowCo1 = co3 Or rowCo2 = co3 Or rowCo3 = co3 Then isMatch = True
        End If

        If isMatch Then
            Dim repA As String, repB As String
            repA = Trim$(CStr(ws.Cells(r, "P").value))
            repB = Trim$(CStr(ws.Cells(r, "Q").value))
            If repA <> "" And Not dictA.Exists(repA) Then dictA(repA) = 1
            If repB <> "" And Not dictB.Exists(repB) Then dictB(repB) = 1
        End If

NEXT_ROW:
    Next r

    ' -- C?p nh?t txtRepA: gi? nguyên giá tr? dang có --
    Dim curA As String
    curA = txtRepA.value
    txtRepA.Clear
    Dim keyA As Variant
    For Each keyA In dictA.keys
        txtRepA.AddItem keyA
    Next keyA
    txtRepA.value = curA   ' khôi ph?c giá tr? cu n?u có

    ' -- C?p nh?t txtRepB: gi? nguyên giá tr? dang có --
    Dim curB As String
    curB = txtRepB.value
    txtRepB.Clear
    Dim keyB As Variant
    For Each keyB In dictB.keys
        txtRepB.AddItem keyB
    Next keyB
    txtRepB.value = curB   ' khôi ph?c giá tr? cu n?u có

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
        'Cancel = True
        Unload Me
        'Me.Hide
    End If
    'ApplyTheme Me
    UnregisterForm Me
End Sub
Public Sub OpenInAddMode(ByVal isGroupA As Boolean)

    On Error GoTo ERRH   ' B?t l?i — không có cái này l?i s? close h?t form

    isBusy = True   ' Ch?n t?t c? Change/Click events trong su?t quá trình

    '-- Hi?n và d?t tr?ng thái checkbox --
    chkEdit.value = False
    chkEdit.Enabled = False

    chkAdd.Visible = True
    chkAdd.value = True

    chkAddAppendix.Visible = True
    chkAddAppendix.value = False
    chkAddAppendix.Enabled = False

    '-- Setup ContractType theo group (v?n trong isBusy=True) --
    If isGroupA Then
        Call SetupContractTypeByRole(True)      ' Main / EPC

        ' Ð?m b?o cboContractType có value h?p l? (Group A)
        If cboContractType.value = "" Or _
           (UCase(cboContractType.value) <> "MAIN" And _
            UCase(cboContractType.value) <> "EPC") Then
            On Error Resume Next
            cboContractType.value = "EPC"
            If Err.Number <> 0 Then
                ' EPC không có trong list ? th? "Main"
                Err.Clear
                cboContractType.value = "Main"
            End If
            On Error GoTo ERRH
        End If

        txtContractID.value = GenerateNextContractID("", "A")

    Else
        Call SetupContractTypeByRole(False)     ' SUB

        ' Ð?m b?o cboContractType có value h?p l? (Group B)
        If cboContractType.value = "" Or _
           UCase(cboContractType.value) <> "SUB" Then
            On Error Resume Next
            cboContractType.value = "SUB"
            On Error GoTo ERRH
        End If

        txtContractID.value = GenerateNextContractID("", "B")
    End If

    '-- Kích ho?t Save và m? khóa inputs --
    btnSaveContract.Enabled = True
    btnSaveContract.Caption = "         Add Contract"

    isBusy = False  ' M? event tr? l?i tru?c EnableContractInputs

    Call EnableContractInputs(True)

    Exit Sub

ERRH:
    isBusy = False  ' Ð?m b?o isBusy du?c reset dù có l?i
    MsgBox "OpenInAddMode error at step [" & isGroupA & "]: " & _
           Err.Description, vbCritical, "Error"

End Sub




