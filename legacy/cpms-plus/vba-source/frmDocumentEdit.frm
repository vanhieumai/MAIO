Attribute VB_Name = "frmDocumentEdit"
Attribute VB_Base = "0{DEA4AF16-6FFF-4FDD-89CA-B1354CBE9421}{4BABFB22-2DEC-4D7D-8DA0-025E0F57B3C8}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'================================================================
' frmDocumentEdit — FULL CODE (Updated)
'
' THAY DOI SO VOI PHIEN BAN TRUOC:
'   [NEW] lstIssuers (ListBox) — them nhieu don vi ban hanh
'         (tuong tu lstRecipients cho don vi tiep nhan)
'         * Them vao form designer: ListBox ten lstIssuers
'   [MOD] Nut + va - dung chung cho ca lstIssuers va lstRecipients
'         Logic phan biet: mLastActiveList ("Issuer" / "Recip")
'         Khi cboIssuer thay doi -> mLastActiveList = "Issuer"
'         Khi cboAddRecip thay doi -> mLastActiveList = "Recip"
'         Nut +: them vao list dang active
'         Nut -: xoa khoi list co item duoc chon
'   [MOD] GetCompanyCodeList — lay tu sheet PROJECT_SETTINGS col B
'         va sheet CONTRACT_DB col Z (trich phan sau "=" neu co)
'         Fallback ve hardcoded neu ca 2 sheet deu trong
'   [MOD] TryGenerateDocNumber — lay issuer tu lstIssuers.List(0)
'         thay vi cboIssuer.Value (vi cboIssuer bay gio la "add" combo)
'   [MOD] btnSave — luu lstIssuers vao col Q phan cach ";"
'   [MOD] InitNew / InitEdit — clear va restore lstIssuers
'
' *** THEM VAO FORM DESIGNER TRUOC KHI DUNG CODE NAY ***
'   ListBox ten: lstIssuers  (dat canh hang "Issuer", tuong tu lstRecipients)
'
' FORMAT SO TAI LIEU (khong doi):
'   New:     YYMMDD.PROJ.Issuer.DocType.Recip1-Recip2.NNN
'   Revised: YYMMDD.PROJ.Issuer.DocType.Recip1-Recip2.NNN.X
'
' SCHEMA DOCUMENT_TRACKER:
'   A: DocumentID  B: ContractID  C: DocType    D: SoTaiLieu
'   E: HoSoName    F: Link        G: FromDate   H: ToDate
'   I: NgayNhan    J: NgayTra     K: NguoiNhan  L: (internal)
'   M: HoSoKem     N: LanVal      O: TinhTrang  P: GhiChu
'   Q: IssuerCode(s) (phan cach ";")   R: Recipients (phan cach ";")
'================================================================
Option Explicit

Private mDocumentID         As String
Private mContractID         As String
Private mInitDocType        As String
Private mIsNewMode          As Boolean
Private mContractNo         As String
Private mCallbackForm       As Object
Private mCallbackContractID As String
Private mIsInitializing     As Boolean
Private mIsUpdatingHoSoKem  As Boolean
Private mLockedFromLibrary  As Boolean

' [NEW] Theo doi list nao dang active de nut +/- biet them/xoa vao dau
' Gia tri: "Issuer" hoac "Recip"
Private mLastActiveList     As String

' [NEW] Luu serial goc (NNN) va version suffix khi mo edit mode
'       De TryGenerateDocNumber khong bi mat serial khi txtSo bi xoa
Private mOriginalSerial     As Long
Private mOriginalSuffix     As String


'================================================================
' DANH SACH MA CONG TY
' [MOD] Doc tu sheet PROJECT_SETTINGS cot B
'       va sheet CONTRACT_DB cot Z (trich code sau "=" neu co)
'       Fallback ve hardcoded neu ca 2 sheet trong
'================================================================
Private Function GetCompanyCodeList() As String()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    '--- 1. Doc tu PROJECT_SETTINGS, cot B (bo dong tieu de row 1) ---
    On Error Resume Next
    Dim ws1 As Worksheet
    Set ws1 = GetDataSheet("PROJECT_SETTINGS")
    If Not ws1 Is Nothing Then
        Dim lr1 As Long
        lr1 = ws1.Cells(ws1.rows.count, "B").End(xlUp).Row
        Dim i As Long
        For i = 2 To lr1
            Dim v1 As String
            Dim a1 As String
            a1 = UCase(Trim(CStr(ws1.Cells(i, "A").value)))
            If a1 = "CLIENT" Or a1 = "SUPCON" Or a1 = "MAINCON" Then
                v1 = UCase(Trim(CStr(ws1.Cells(i, "B").value)))
                If v1 <> "" And Not dict.Exists(v1) Then dict.Add v1, v1
            End If
        Next i
    End If
    On Error GoTo 0

    '--- 2. Doc tu CONTRACT_DB, cot Z ---
    '    Neu cell co dang "TYPE=CODE" (vi du "SUBCON=CTC") thi lay phan sau "="
    '    Neu khong co "=" thi dung nguyen gia tri cell
    On Error Resume Next
    Dim ws2 As Worksheet
    Set ws2 = GetDataSheet("CONTRACT_DB")
    If Not ws2 Is Nothing Then
        Dim lr2 As Long
        lr2 = ws2.Cells(ws2.rows.count, "Z").End(xlUp).Row
        Dim j As Long
        For j = 2 To lr2
            Dim v2 As String
            v2 = Trim(CStr(ws2.Cells(j, "Z").value))
            If v2 <> "" Then
                Dim eqPos As Integer
                eqPos = InStr(v2, "=")
                If eqPos > 0 Then
                    v2 = UCase(Trim(mID(v2, eqPos + 1)))
                Else
                    v2 = UCase(v2)
                End If
                If v2 <> "" And Not dict.Exists(v2) Then dict.Add v2, v2
            End If
        Next j
    End If
    On Error GoTo 0

    '--- Fallback neu khong lay duoc gi tu sheet ---
    If dict.count = 0 Then
        GetCompanyCodeList = Split("CLI;ART;VEC;CTC;NSH;KD;TID;HHG;HD;EC-KD;EC-PHC;EC-TG;EC-NSH", ";")
        Exit Function
    End If

    '--- Chuyen dict sang array ---
    Dim keys() As String
    ReDim keys(0 To dict.count - 1)
    Dim k As Long: k = 0
    Dim keyItem As Variant
    For Each keyItem In dict.keys
        keys(k) = CStr(keyItem): k = k + 1
    Next keyItem
    GetCompanyCodeList = keys
End Function


'================================================================
' PUBLIC: Mo form tu Library page (right-click label)
'================================================================
Public Sub SetFilter(ByVal contractID As String, ByVal docType As String)
    mContractID = contractID
    mInitDocType = UCase(Trim(docType))
    mLockedFromLibrary = True

    Dim contractNo    As String
    Dim contractValue As String
    Dim company       As String
    Dim hangMuc       As String
    Dim fromDate      As String
    Dim toDate        As String

    If Not gDictContract Is Nothing Then
        If gDictContract.Exists(mContractID) Then
            Dim r As Long
            r = gDictContract(mContractID)
            contractNo = Trim(CStr(gArrContract(r, 2)))
            contractValue = FastVN(gArrContract(r, 15))
            company = Trim(CStr(gArrContract(r, 8)))
            hangMuc = Trim(CStr(gArrContract(r, 11)))
            If IsDate(gArrContract(r, 4)) Then
                fromDate = Format(gArrContract(r, 4), "dd/mm/yyyy")
            End If
            If IsDate(gArrContract(r, 5)) Then
                toDate = Format(gArrContract(r, 5), "dd/mm/yyyy")
            End If
        End If
    End If

    InitNew mContractID, contractNo, contractValue, company, hangMuc, fromDate, toDate
    AutoSelectDocType mInitDocType
    LockDocumentTypeControls True
End Sub


'================================================================
' LOCK / UNLOCK cac control chon loai tai lieu
'================================================================
Private Sub LockDocumentTypeControls(ByVal doLock As Boolean)
    On Error Resume Next
    optHD.Enabled = Not doLock
    optIPC.Enabled = Not doLock
    optKhac.Enabled = Not doLock
    cboDocType.Enabled = Not doLock
    On Error GoTo 0
End Sub


'================================================================
' Tu dong chon optionbutton + cboDocType theo docType
'================================================================
Private Sub AutoSelectDocType(ByVal docType As String)
    Dim dt As String
    dt = UCase(Trim(docType))

    Select Case dt
        Case "HD", "CONTRACT", "CON"
            optHD.value = True
        Case "IPC"
            optIPC.value = True
        Case Else
            optKhac.value = True
            ApplyHoSoMode "Khac"
            On Error Resume Next
            cboDocType.Enabled = True
            cboDocType.value = dt
            If cboDocType.ListIndex = -1 Then
                cboDocType.AddItem dt
                cboDocType.value = dt
            End If
            On Error GoTo 0
            SetHoSoNameFromDocType dt
    End Select
End Sub


'================================================================
' Ap dung trang thai UI theo loai ho so
'================================================================
Private Sub ApplyHoSoMode(ByVal hoSo As String)
    Select Case UCase(Trim(hoSo))
        Case "HD"
            txtHoSoName.Enabled = True
            cboDocType.value = ""
            cboDocType.Enabled = False
            txtSo.Enabled = True
            txtNgayNhan.Enabled = True
            txtNgayTra.Enabled = False
            btnNgayNhanToday.Enabled = True
            btnNgayTraToday.Enabled = False
            cboNguoiNhan.Enabled = False
            cboTinhTrang.Enabled = False
            SetAttachedDocsEnabled False

        Case "IPC"
            txtHoSoName.Enabled = True
            cboDocType.value = ""
            cboDocType.Enabled = False
            txtSo.Enabled = True
            txtNgayNhan.Enabled = True
            txtNgayTra.Enabled = True
            btnNgayNhanToday.Enabled = True
            btnNgayTraToday.Enabled = True
            cboNguoiNhan.Enabled = True
            cboTinhTrang.Enabled = True
            SetAttachedDocsEnabled True

        Case Else
            cboDocType.Enabled = True
            txtHoSoName.Enabled = True
            txtSo.Enabled = True
            txtNgayNhan.Enabled = True
            txtNgayTra.Enabled = True
            btnNgayNhanToday.Enabled = True
            btnNgayTraToday.Enabled = True
            cboNguoiNhan.Enabled = True
            cboTinhTrang.Enabled = True
            SetAttachedDocsEnabled True
    End Select

    txtLink.Enabled = True
    txtGhiChu.Enabled = True
End Sub


'================================================================
' Dat txtHoSoName theo docType
'================================================================
Private Sub SetHoSoNameFromDocType(ByVal dt As String)
    On Error Resume Next
    txtHoSoName.value = BuildSuggestedHoSoName(dt)
    txtHoSoName.Enabled = True
    On Error GoTo 0
End Sub


'================================================================
' Xay dung ten goi y cho txtHoSoName
' Format: "[Ten day du DocType] - [HangMuc] ([SoHD])"
'================================================================
Private Function BuildSuggestedHoSoName(ByVal dt As String) As String
    Dim fullName As String
    fullName = GetDocTypeFullName(dt)

    Dim hangMuc As String
    hangMuc = Trim(lblHangMuc.Caption)

    Dim sohd As String
    sohd = Trim(lblSoHD.Caption)

    Dim result As String
    result = fullName
    If hangMuc <> "" Then result = result & " - " & hangMuc
    If sohd <> "" Then result = result & " (" & sohd & ")"
    BuildSuggestedHoSoName = result
End Function


'================================================================
' Ten day du cua docType code
'================================================================
Private Function GetDocTypeFullName(ByVal dt As String) As String
    Select Case UCase(Trim(dt))
        Case "CON":  GetDocTypeFullName = "Contract"
        Case "APP":  GetDocTypeFullName = "Appendix"
        Case "VO":   GetDocTypeFullName = "Variation Orders"
        Case "IPC":  GetDocTypeFullName = "Interim Payment Certificate"
        Case "ITP":  GetDocTypeFullName = "ITP (Inspection and Test Plan)"
        'Case "MIR":  GetDocTypeFullName = "MIR (Material Inspection Request)"
        'Case "WIR":  GetDocTypeFullName = "WIR (Work Inspection Request)"
        'Case "MAR":  GetDocTypeFullName = "MAR (Material Acceptance Record)"
        'Case "IAR":  GetDocTypeFullName = "IAR (Inspection Acceptance Record)"
        Case "RFA":  GetDocTypeFullName = "RFA (Request for Approval)"
        Case "RFC":  GetDocTypeFullName = "RFC (Request for Change)"
        Case "RFI":  GetDocTypeFullName = "RFI (Request for Information)"
        Case "SI":   GetDocTypeFullName = "SI (Site Instruction)"
        Case "NCR":  GetDocTypeFullName = "NCR (Non-Conformity Report)"
        Case "CAR":  GetDocTypeFullName = "CAR (Corrective Action Request)"
        Case "TRN":  GetDocTypeFullName = "TRN (Transmittal)"
        Case "MOM":  GetDocTypeFullName = "MOM (Minutes of Meeting)"
        Case "CMT":  GetDocTypeFullName = "CMT (Comment Sheet)"
        Case "LTR":  GetDocTypeFullName = "LTR (Letter)"
        Case "SR":   GetDocTypeFullName = "SR (Site Report)"
        Case "AD":   GetDocTypeFullName = "AD (Approved Drawings)"
        Case "SD":   GetDocTypeFullName = "SD (Shop Drawings)"
        Case "ABD":  GetDocTypeFullName = "ABD (As-Built Drawings)"
        Case "MOH":  GetDocTypeFullName = "MOH (Minutes of Handover)"
        Case "SPT":  GetDocTypeFullName = "SPT (Submission Plan and Tracking)"
        Case Else:   GetDocTypeFullName = dt
    End Select
End Function


'================================================================
' PUBLIC: SetCallback
'================================================================
Public Sub SetCallback(ByVal caller As Object, ByVal contractID As String)
    Set mCallbackForm = caller
    mCallbackContractID = contractID
End Sub


Private Sub NotifyAndClose()
    Dim savedID As String
    savedID = mDocumentID
    Me.Hide
    If Not mCallbackForm Is Nothing Then
        On Error Resume Next
        mCallbackForm.OnDocumentEditDone savedID
        On Error GoTo 0
        Set mCallbackForm = Nothing
    End If
End Sub


Public Property Get SavedDocumentID() As String
    SavedDocumentID = mDocumentID
End Property


'================================================================
' Xac dinh nhom hop dong (A / B / ...)
'================================================================
Private Function GetContractGroup(ByVal contractID As String) As String
    Dim hyphen As Integer
    hyphen = InStr(contractID, "-")
    If hyphen > 1 Then
        GetContractGroup = UCase(Trim(Left(contractID, hyphen - 1)))
    Else
        GetContractGroup = UCase(Trim(contractID))
    End If
End Function


'================================================================
' Dat caption label Ngay/Nguoi theo nhom
'================================================================
Private Sub SetDateLabels(ByVal isGroupA As Boolean)
    On Error Resume Next
    If isGroupA Then
        lblNgayNhan.Caption = "Processed Date"
        lblNgayTra.Caption = "Completed Date"
        lblNguoiNhan.Caption = "Send To"
    Else
        lblNgayNhan.Caption = "Received Date"
        lblNgayTra.Caption = "Returned Date"
        lblNguoiNhan.Caption = "Returned To"
    End If
    On Error GoTo 0
End Sub


'================================================================
' Nap danh sach Status theo nhom
'================================================================
Private Sub SetStatusList(ByVal isGroupA As Boolean)
    Dim curVal As String
    curVal = cboTinhTrang.value
    cboTinhTrang.Clear
    If isGroupA Then
        cboTinhTrang.AddItem "Processing"
        cboTinhTrang.AddItem "Completed"
    Else
        cboTinhTrang.AddItem "Checking"
        cboTinhTrang.AddItem "Returned"
        cboTinhTrang.AddItem "Accepted"
    End If
    On Error Resume Next
    cboTinhTrang.value = curVal
    On Error GoTo 0
End Sub


'================================================================
' Lay gia tri loai ho so tu optionbutton
'================================================================
Private Function GetHoSoValue() As String
    If optHD.value Then
        GetHoSoValue = "CON"
    ElseIf optIPC.value Then
        GetHoSoValue = "IPC"
    ElseIf optKhac.value Then
        GetHoSoValue = "Khac"
    Else
        GetHoSoValue = ""
    End If
End Function


'================================================================
' Lay type code de luu vao cot C
'================================================================
Private Function GetHoSoStoreValue() As String
    Select Case GetHoSoValue()
        Case "CON":  GetHoSoStoreValue = "CON"
        Case "IPC":  GetHoSoStoreValue = "IPC"
        Case "Khac"
            Dim dt As String
            dt = UCase(Trim(cboDocType.value))
            If dt <> "" Then
                GetHoSoStoreValue = dt
            Else
                GetHoSoStoreValue = "Other"
            End If
        Case Else:   GetHoSoStoreValue = ""
    End Select
End Function


'================================================================
' Dat optionbutton theo gia tri chuoi (dung khi load Edit)
'================================================================
Private Sub SetHoSoValue(ByVal hoSo As String)
    optHD.value = False
    optIPC.value = False
    optKhac.value = False

    Dim h As String
    h = UCase(Trim(hoSo))

    If h = "HD" Or Left(h, 3) = "CON" Then
        optHD.value = True
        txtHoSoName.Enabled = True
        cboDocType.value = ""
        cboDocType.Enabled = False

    ElseIf h = "IPC" Then
        optIPC.value = True
        txtHoSoName.Enabled = True
        cboDocType.value = ""
        cboDocType.Enabled = False

    Else
        optKhac.value = True
        cboDocType.Enabled = True
        If h = "PLHD" Then h = "APP"
        On Error Resume Next
        cboDocType.value = h
        If cboDocType.ListIndex = -1 And h <> "" Then
            cboDocType.AddItem h
            cboDocType.value = h
        End If
        On Error GoTo 0
        txtHoSoName.Enabled = True
    End If
End Sub


'================================================================
' INITIALIZE FORM
'================================================================
Private Sub UserForm_Initialize()
    mLockedFromLibrary = False
    mLastActiveList = "Issuer"

    With cboDocType
        .Clear
        .AddItem "APP"
        .AddItem "VO"
        .AddItem "SPT"
        .AddItem "ITP"
        .AddItem "RFA"
        .AddItem "RFC"
        .AddItem "RFI"
        .AddItem "SI"
        .AddItem "NCR"
        .AddItem "CAR"
        .AddItem "TRN"
        .AddItem "MOM"
        .AddItem "CMT"
        .AddItem "LTR"
        .AddItem "SR"
        .AddItem "AD"
        .AddItem "SD"
        .AddItem "ABD"
        .AddItem "MOH"
        .Enabled = False
    End With

    LoadCompanyCodes

    cboTinhTrang.Clear

    Dim i As Integer
    Dim cbo As Variant
    For Each cbo In Array(cboGTTT, cboTHKL, cboDGKL, cboNTCL, cboNTVL, cboHC, cboHSKhac)
        With cbo
            .Clear
            For i = 1 To 10
                .AddItem CStr(i)
            Next i
        End With
    Next cbo

    LoadNguoiNhanList

    txtSo.Locked = False
    ' Kh?i t?o 2 sub-type combobox — m?c d?nh disabled/empty
    On Error Resume Next
    cboSubType1.Clear:  cboSubType1.value = "":  cboSubType1.Enabled = False
    cboSubType2.Clear:  cboSubType2.value = "":  cboSubType2.Enabled = False
    On Error GoTo 0
    
    SetAttachedDocsEnabled False

    On Error Resume Next
    chkNewDoc.value = True
    chkNewDoc.Enabled = True
    On Error GoTo 0
    ApplyTheme Me
    'UnregisterForm Me
End Sub


Private Sub SetAttachedDocsEnabled(ByVal isEnabled As Boolean)
    chkGTTT.Enabled = isEnabled:    cboGTTT.Enabled = isEnabled
    chkTHKL.Enabled = isEnabled:    cboTHKL.Enabled = isEnabled
    chkDGKL.Enabled = isEnabled:    cboDGKL.Enabled = isEnabled
    chkNTCL.Enabled = isEnabled:    cboNTCL.Enabled = isEnabled
    chkNTVL.Enabled = isEnabled:    cboNTVL.Enabled = isEnabled
    chkHC.Enabled = isEnabled:      cboHC.Enabled = isEnabled
    chkHSKhac.Enabled = isEnabled:  cboHSKhac.Enabled = isEnabled
    txtHoSoKemNote.Enabled = isEnabled

    If Not isEnabled Then
        chkGTTT.value = False:    cboGTTT.value = ""
        chkTHKL.value = False:    cboTHKL.value = ""
        chkDGKL.value = False:    cboDGKL.value = ""
        chkNTCL.value = False:    cboNTCL.value = ""
        chkNTVL.value = False:    cboNTVL.value = ""
        chkHC.value = False:      cboHC.value = ""
        chkHSKhac.value = False:  cboHSKhac.value = ""
        txtHoSoKemNote.value = ""
    End If
End Sub


'================================================================
' Nap company codes vao cboIssuer (add-issuer combo) va cboAddRecip
'================================================================
Private Sub LoadCompanyCodes()
    Dim codes() As String
    codes = GetCompanyCodeList()
    On Error Resume Next
    cboIssuer.Clear
    cboAddRecip.Clear
    Dim i As Long
    For i = 0 To UBound(codes)
        Dim c As String
        c = UCase(Trim(codes(i)))
        If c <> "" Then
            cboIssuer.AddItem c
            cboAddRecip.AddItem c
        End If
    Next i
    On Error GoTo 0
End Sub


'================================================================
' Nap danh sach Nguoi Nhan tu DOCUMENT_TRACKER col K
'================================================================
Private Sub LoadNguoiNhanList()
    On Error Resume Next
    cboNguoiNhan.Clear
    Dim ws As Worksheet
    Set ws = GetDataSheet("DOCUMENT_TRACKER")
    If ws Is Nothing Then Exit Sub
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    Dim j As Long, v As String
    For j = 2 To lastRow
        v = Trim(CStr(ws.Cells(j, "K").value))
        If v <> "" And Not dict.Exists(v) Then dict.Add v, v
    Next j
    If dict.count = 0 Then Exit Sub
    Dim keys() As String
    Dim k As Long
    ReDim keys(0 To dict.count - 1)
    k = 0
    Dim keyItem As Variant
    For Each keyItem In dict.keys
        keys(k) = CStr(keyItem): k = k + 1
    Next keyItem
    Dim r1 As Long, r2 As Long, tmp As String
    For r1 = 0 To dict.count - 2
        For r2 = r1 + 1 To dict.count - 1
            If keys(r2) < keys(r1) Then
                tmp = keys(r1): keys(r1) = keys(r2): keys(r2) = tmp
            End If
        Next r2
    Next r1
    For k = 0 To dict.count - 1
        cboNguoiNhan.AddItem keys(k)
    Next k
    On Error GoTo 0
End Sub


'================================================================
' INIT NEW — tao moi document
' [MOD] Clear ca lstIssuers lan lstRecipients
'================================================================
Public Sub InitNew( _
    ByVal contractID As String, _
    Optional ByVal contractNo As String = "", _
    Optional ByVal contractValue As String = "", _
    Optional ByVal company As String = "", _
    Optional ByVal hangMuc As String = "", _
    Optional ByVal fromDate As String = "", _
    Optional ByVal toDate As String = "")

    mIsInitializing = True
    mIsNewMode = True
    mDocumentID = ""
    mContractID = contractID
    mContractNo = contractNo
    mLastActiveList = "Issuer"

    SetDateLabels (GetContractGroup(contractID) = "A")
    SetStatusList (GetContractGroup(contractID) = "A")

    lblSoHD.Caption = contractNo
    lblGiaTriHD.Caption = contractValue
    lblCongTy.Caption = company
    lblHangMuc.Caption = hangMuc
    lblTuNgay.Caption = fromDate
    lblDenNgay.Caption = toDate

    optHD.value = False:    optHD.Enabled = True
    optIPC.value = False:   optIPC.Enabled = True
    optKhac.value = False:  optKhac.Enabled = True

    On Error Resume Next
    optPLHD.Visible = False
    On Error GoTo 0

    txtHoSoName.value = ""
    txtHoSoName.Enabled = True
    cboDocType.value = ""
    cboDocType.Enabled = False

    txtSo.value = ""
    lblLanVal.Caption = "L1"
    txtLink.value = ""
    txtNgayNhan.value = ""
    txtNgayTra.value = ""
    cboNguoiNhan.value = ""
    ' [MOD] Mac dinh TinhTrang khi them moi: Processing (Group A) hoac Checking (cac nhom khac)
    On Error Resume Next
    If GetContractGroup(contractID) = "A" Then
        cboTinhTrang.value = "Processing"
    Else
        cboTinhTrang.value = "Checking"
    End If
    On Error GoTo 0
    txtGhiChu.value = ""
    btnNgayNhanToday.Enabled = True
    btnNgayTraToday.Enabled = False

    On Error Resume Next
    ' Clear ca lstIssuers lan lstRecipients
    lstIssuers.Clear
    cboIssuer.value = ""
    lstRecipients.Clear
    cboAddRecip.value = ""
    
    ' Clear sub-type comboboxes
    cboSubType1.Clear:  cboSubType1.value = "":  cboSubType1.Enabled = False
    cboSubType2.Clear:  cboSubType2.value = "":  cboSubType2.Enabled = False
    
    chkNewDoc.value = True
    chkNewDoc.Enabled = True
    On Error GoTo 0

    SetAttachedDocsEnabled False
    btnSave.Caption = "        Add"

    mIsInitializing = False

    ' [NEW] Tu dong dien Issuer mac dinh = Supervision Consultant (SUPCON)
    On Error Resume Next
    Dim defIssuer As String
    defIssuer = GetIssuerCodeForContract(contractID)
    'defIssuer = UCase(Trim(GetGlobalSetting("SUPCON")))
    'If defIssuer = "" Then defIssuer = "ART"
    lstIssuers.AddItem defIssuer

    ' [NEW] Tu dong dien Recipient mac dinh tu contract (company code)
    Dim defRecip As String
    defRecip = GetContractCompanyCode(contractID)
    If defRecip <> "" Then lstRecipients.AddItem defRecip
    On Error GoTo 0

    ' Cap nhat txtSo ngay neu docType da duoc chon (truong hop SetFilter goi InitNew)
    TryGenerateDocNumber
End Sub


'================================================================
' Lay company code mac dinh cua hop dong
' Uu tien: doc tu cot Z sheet CONTRACT_DB (dinh dang "ROLE=CODE")
'          hoac tu gArrContract index 26 (cot Z neu load full sheet)
' Fallback: GetGlobalSetting("MAINCON")
'================================================================
Private Function GetContractCompanyCode(ByVal contractID As String) As String
    GetContractCompanyCode = ""
    On Error Resume Next

    '--- Thu 1: dung global array neu da load ---
    If Not gDictContract Is Nothing Then
        If gDictContract.Exists(contractID) Then
            Dim r As Long
            r = gDictContract(contractID)
            ' Thu cot Z (index 26 neu array load full CONTRACT_DB)
            Dim rawZ As String
            rawZ = Trim(CStr(gArrContract(r, 26)))
            If rawZ <> "" Then
                Dim eq1 As Integer: eq1 = InStr(rawZ, "=")
                Dim code1 As String
                code1 = UCase(IIf(eq1 > 0, Trim(mID(rawZ, eq1 + 1)), rawZ))
                If code1 <> "" Then
                    GetContractCompanyCode = code1
                    Exit Function
                End If
            End If
        End If
    End If

    '--- Thu 2: doc truc tiep tu sheet CONTRACT_DB ---
    Dim ws As Worksheet
    Set ws = GetDataSheet("CONTRACT_DB")
    If Not ws Is Nothing Then
        Dim lastRow As Long
        lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
        Dim i As Long, c As Integer
        For i = 2 To lastRow
            ' Tim contractID trong 10 cot dau
            For c = 1 To 10
                If UCase(Trim(CStr(ws.Cells(i, c).value))) = UCase(Trim(contractID)) Then
                    Dim rawZ2 As String
                    rawZ2 = Trim(CStr(ws.Cells(i, "Z").value))
                    If rawZ2 <> "" Then
                        Dim eq2 As Integer: eq2 = InStr(rawZ2, "=")
                        Dim code2 As String
                        code2 = UCase(IIf(eq2 > 0, Trim(mID(rawZ2, eq2 + 1)), rawZ2))
                        If code2 <> "" Then
                            GetContractCompanyCode = code2
                            Exit Function
                        End If
                    End If
                    Exit For
                End If
            Next c
        Next i
    End If

    '--- Fallback: MAINCON tu project settings ---
    Dim maincon As String
    maincon = UCase(Trim(GetGlobalSetting("MAINCON")))
    If maincon <> "" Then GetContractCompanyCode = maincon

    On Error GoTo 0
End Function


'================================================================
' INIT EDIT — chinh sua document cu
' [MOD] Doc lstIssuers tu col Q (co the co nhieu gia tri phan cach ";")
' [MOD] Doc lstRecipients tu col R
'================================================================
Public Sub InitEdit(ByVal documentID As String)

    mIsInitializing = True
    mIsNewMode = False
    mDocumentID = documentID
    mLockedFromLibrary = False
    mLastActiveList = "Issuer"

    On Error Resume Next
    optPLHD.Visible = False
    On Error GoTo 0

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetDataSheet("DOCUMENT_TRACKER")
    On Error GoTo ERR_LOAD
    If ws Is Nothing Then
        ShowMsgCenter "Sheet DOCUMENT_TRACKER not found!"
        Me.Hide
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim i As Long

    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, "A").value)) = Trim(documentID) Then

            mContractID = Trim(CStr(ws.Cells(i, "B").value))
            SetDateLabels (GetContractGroup(mContractID) = "A")
            SetStatusList (GetContractGroup(mContractID) = "A")

            Dim contractNo    As String
            Dim contractValue As String
            Dim company       As String
            Dim hangMuc       As String
            Dim fromDate      As String
            Dim toDate        As String

            If Not gDictContract Is Nothing Then
                If gDictContract.Exists(mContractID) Then
                    Dim r As Long
                    r = gDictContract(mContractID)
                    contractNo = Trim(CStr(gArrContract(r, 2)))
                    contractValue = FastVN(gArrContract(r, 15))
                    company = Trim(CStr(gArrContract(r, 8)))
                    hangMuc = Trim(CStr(gArrContract(r, 11)))
                    If IsDate(gArrContract(r, 4)) Then
                        fromDate = Format(gArrContract(r, 4), "dd/mm/yyyy")
                    End If
                    If IsDate(gArrContract(r, 5)) Then
                        toDate = Format(gArrContract(r, 5), "dd/mm/yyyy")
                    End If
                End If
            End If

            mContractNo = contractNo
            lblSoHD.Caption = contractNo
            lblGiaTriHD.Caption = contractValue
            lblCongTy.Caption = company
            lblHangMuc.Caption = hangMuc
            lblTuNgay.Caption = fromDate
            lblDenNgay.Caption = toDate

            Dim hoSo As String
            hoSo = Trim(CStr(ws.Cells(i, "C").value))
            SetHoSoValue hoSo
            ApplyHoSoMode GetHoSoValue()

            txtHoSoName.value = Trim(CStr(ws.Cells(i, "E").value))
            txtHoSoName.Enabled = True

            txtSo.value = Trim(CStr(ws.Cells(i, "D").value))

            ' [NEW] Luu serial goc ngay khi load de dung lai trong TryGenerateDocNumber
            mOriginalSerial = 0
            mOriginalSuffix = ""
            Dim origSo As String: origSo = txtSo.value
            If origSo <> "" Then
                Dim oParts() As String: oParts = Split(origSo, ".")
                Dim op As Integer
                For op = 0 To UBound(oParts)
                    If Len(oParts(op)) = 3 Then
                        Dim oSn As Long: oSn = 0
                        On Error Resume Next
                        oSn = CLng(oParts(op))
                        On Error GoTo 0
                        If oSn > 0 Then
                            mOriginalSerial = oSn
                            If op + 1 <= UBound(oParts) Then
                                Dim oSuf As String
                                oSuf = UCase(Trim(oParts(op + 1)))
                                If Len(oSuf) = 1 And oSuf >= "A" And oSuf <= "Z" Then
                                    mOriginalSuffix = oSuf
                                End If
                            End If
                            Exit For
                        End If
                    End If
                Next op
            End If

            lblLanVal.Caption = Trim(CStr(ws.Cells(i, "N").value))
            txtLink.value = Trim(CStr(ws.Cells(i, "F").value))

            If IsDate(ws.Cells(i, "I").value) Then
                txtNgayNhan.value = Format(ws.Cells(i, "I").value, "dd/mm/yyyy")
            End If
            If IsDate(ws.Cells(i, "J").value) Then
                txtNgayTra.value = Format(ws.Cells(i, "J").value, "dd/mm/yyyy")
            End If

            cboNguoiNhan.value = Trim(CStr(ws.Cells(i, "K").value))
            cboTinhTrang.value = Trim(CStr(ws.Cells(i, "O").value))
            txtGhiChu.value = Trim(CStr(ws.Cells(i, "P").value))
            DecodeHoSoKem Trim(CStr(ws.Cells(i, "M").value))

            ' [MOD] Doc lstIssuers tu col Q (nhieu gia tri phan cach ";")
            On Error Resume Next
            Dim issuerStr As String
            issuerStr = Trim(CStr(ws.Cells(i, "Q").value))
            lstIssuers.Clear
            cboIssuer.value = ""
            If issuerStr <> "" Then
                Dim issuers() As String
                issuers = Split(issuerStr, ";")
                Dim ii As Long
                For ii = 0 To UBound(issuers)
                    Dim ic As String
                    ic = UCase(Trim(issuers(ii)))
                    If ic <> "" Then lstIssuers.AddItem ic
                Next ii
            End If

            ' Doc lstRecipients tu col R
            Dim recipStr As String
            recipStr = Trim(CStr(ws.Cells(i, "R").value))
            lstRecipients.Clear
            If recipStr <> "" Then
                Dim recips() As String
                recips = Split(recipStr, ";")
                Dim ri As Long
                For ri = 0 To UBound(recips)
                    Dim rC As String
                    rC = UCase(Trim(recips(ri)))
                    If rC <> "" Then lstRecipients.AddItem rC
                Next ri
            End If

            ' Ð?c cboSubType1 t? col S, cboSubType2 t? col T
            On Error Resume Next
            Dim st1Val As String: st1Val = Trim(CStr(ws.Cells(i, "S").value))
            Dim st2Val As String: st2Val = Trim(CStr(ws.Cells(i, "T").value))

            ' Populate cboSubType1 tru?c (d?a theo DocType dã load)
            Dim dtForSub As String
            dtForSub = UCase(Trim(ws.Cells(i, "C").value))
            UpdateSubType1 dtForSub
            If st1Val <> "" Then
                cboSubType1.value = st1Val
                If cboSubType1.ListIndex = -1 Then
                    cboSubType1.AddItem st1Val
                    cboSubType1.value = st1Val
                End If
            End If

            ' Populate cboSubType2 tru?c r?i set value
            If st1Val <> "" Then UpdateSubType2 st1Val
            If st2Val <> "" Then
                cboSubType2.value = st2Val
                If cboSubType2.ListIndex = -1 Then
                    cboSubType2.AddItem st2Val
                    cboSubType2.value = st2Val
                End If
            End If
            
            On Error GoTo 0

            Exit For
        End If
    Next i

    optHD.Enabled = False
    optIPC.Enabled = False
    optKhac.Enabled = False
    btnSave.Caption = "        Save"

    On Error Resume Next
    chkNewDoc.value = True
    chkNewDoc.Enabled = True
    On Error GoTo 0

    mIsInitializing = False

    ' Edit mode: giu nguyen so tai lieu da luu tu sheet, KHONG tu dong generate lai
    ' (txtSo da duoc set o tren tu ws.Cells(i, "D"))

    Exit Sub

ERR_LOAD:
    mIsInitializing = False
    ShowMsgCenter "InitEdit Error: " & Err.Description
End Sub


'================================================================
' Giai ma HoSo Kem tu chuoi encoded
'================================================================
Private Sub DecodeHoSoKem(ByVal encoded As String)
    If Trim(encoded) = "" Then Exit Sub
    Dim parts() As String
    parts = Split(encoded, "|")
    Dim p As String, k As String, v As String
    Dim sepPos As Integer, idx As Integer
    For idx = 0 To UBound(parts)
        p = Trim(parts(idx))
        sepPos = InStr(p, ":")
        If sepPos > 0 Then
            k = Left(p, sepPos - 1)
            v = mID(p, sepPos + 1)
        Else
            k = p: v = ""
        End If
        ' Set gia tri cbo VA tu dong check checkbox tuong ung neu co gia tri
        Select Case k
            Case "GTTT": cboGTTT.value = v:    chkGTTT.value = (Trim(v) <> "")
            Case "THKL": cboTHKL.value = v:    chkTHKL.value = (Trim(v) <> "")
            Case "DGKL": cboDGKL.value = v:    chkDGKL.value = (Trim(v) <> "")
            Case "NTCL": cboNTCL.value = v:    chkNTCL.value = (Trim(v) <> "")
            Case "NTVL": cboNTVL.value = v:    chkNTVL.value = (Trim(v) <> "")
            Case "HC":   cboHC.value = v:      chkHC.value = (Trim(v) <> "")
            Case "HSK":  cboHSKhac.value = v:  chkHSKhac.value = (Trim(v) <> "")
            Case "NOTE": txtHoSoKemNote.value = v
        End Select
    Next idx
End Sub


'================================================================
' EVENTS: HO SO KEM — cbo_Change tu dong check checkbox
'                      chk_Click tu dong set qty = 1 neu cbo chua co gia tri
'
' Quy tac:
'   - Khi chon gia tri trong cboGTTT (>0)   -> tu dong check chkGTTT
'   - Khi check chkGTTT ma cboGTTT chua chon -> tu dong chon "1"
'   - Khi bo check chkGTTT                   -> xoa gia tri cboGTTT
'   (Tuong tu cho 6 loai con lai)
'================================================================
Private Sub cboGTTT_Change()
    If mIsInitializing Then Exit Sub
    If Trim(cboGTTT.value) <> "" Then chkGTTT.value = True
End Sub

Private Sub chkGTTT_Click()
    If mIsInitializing Then Exit Sub
    If chkGTTT.value Then
        If Trim(cboGTTT.value) = "" Then cboGTTT.value = "1"
    Else
        cboGTTT.value = ""
    End If
End Sub

Private Sub cboTHKL_Change()
    If mIsInitializing Then Exit Sub
    If Trim(cboTHKL.value) <> "" Then chkTHKL.value = True
End Sub

Private Sub chkTHKL_Click()
    If mIsInitializing Then Exit Sub
    If chkTHKL.value Then
        If Trim(cboTHKL.value) = "" Then cboTHKL.value = "1"
    Else
        cboTHKL.value = ""
    End If
End Sub

Private Sub cboDGKL_Change()
    If mIsInitializing Then Exit Sub
    If Trim(cboDGKL.value) <> "" Then chkDGKL.value = True
End Sub

Private Sub chkDGKL_Click()
    If mIsInitializing Then Exit Sub
    If chkDGKL.value Then
        If Trim(cboDGKL.value) = "" Then cboDGKL.value = "1"
    Else
        cboDGKL.value = ""
    End If
End Sub

Private Sub cboNTCL_Change()
    If mIsInitializing Then Exit Sub
    If Trim(cboNTCL.value) <> "" Then chkNTCL.value = True
End Sub

Private Sub chkNTCL_Click()
    If mIsInitializing Then Exit Sub
    If chkNTCL.value Then
        If Trim(cboNTCL.value) = "" Then cboNTCL.value = "1"
    Else
        cboNTCL.value = ""
    End If
End Sub

Private Sub cboNTVL_Change()
    If mIsInitializing Then Exit Sub
    If Trim(cboNTVL.value) <> "" Then chkNTVL.value = True
End Sub

Private Sub chkNTVL_Click()
    If mIsInitializing Then Exit Sub
    If chkNTVL.value Then
        If Trim(cboNTVL.value) = "" Then cboNTVL.value = "1"
    Else
        cboNTVL.value = ""
    End If
End Sub

Private Sub cboHC_Change()
    If mIsInitializing Then Exit Sub
    If Trim(cboHC.value) <> "" Then chkHC.value = True
End Sub

Private Sub chkHC_Click()
    If mIsInitializing Then Exit Sub
    If chkHC.value Then
        If Trim(cboHC.value) = "" Then cboHC.value = "1"
    Else
        cboHC.value = ""
    End If
End Sub

Private Sub cboHSKhac_Change()
    If mIsInitializing Then Exit Sub
    If Trim(cboHSKhac.value) <> "" Then chkHSKhac.value = True
End Sub

Private Sub chkHSKhac_Click()
    If mIsInitializing Then Exit Sub
    If chkHSKhac.value Then
        If Trim(cboHSKhac.value) = "" Then cboHSKhac.value = "1"
    Else
        cboHSKhac.value = ""
    End If
End Sub


'================================================================
' EVENTS OPTIONBUTTON
'================================================================
Private Sub optHD_Click()
    If Not optHD.value Then Exit Sub
    ApplyHoSoMode "HD"
    Dim hdName As String
    hdName = "Contract"
    Dim hm As String:   hm = Trim(lblHangMuc.Caption)
    Dim sohd As String: sohd = Trim(lblSoHD.Caption)
    If hm <> "" Then hdName = hdName & " - " & hm
    If sohd <> "" Then hdName = hdName & " (" & sohd & ")"
    txtHoSoName.value = hdName
    txtSo.value = mContractNo
    lblLanVal.Caption = "L1"
    On Error Resume Next
    chkNewDoc.value = True
    chkNewDoc.Enabled = False
    On Error GoTo 0
End Sub


Private Sub optIPC_Click()
    If Not optIPC.value Then Exit Sub
    ApplyHoSoMode "IPC"
    If Not mIsInitializing Then
        Dim ipcName As String
        ipcName = "Interim Payment Certificate"
        Dim hm2 As String:   hm2 = Trim(lblHangMuc.Caption)
        Dim sohd2 As String: sohd2 = Trim(lblSoHD.Caption)
        If hm2 <> "" Then ipcName = ipcName & " - " & hm2
        If sohd2 <> "" Then ipcName = ipcName & " (" & sohd2 & ")"
        txtHoSoName.value = ipcName
        lblLanVal.Caption = "L1"
        TryGenerateDocNumber
    End If
End Sub


Private Sub optKhac_Click()
    If Not optKhac.value Then Exit Sub
    ApplyHoSoMode "Khac"
    If Not mIsInitializing Then
        cboDocType.value = ""
        txtHoSoName.value = ""
        txtHoSoName.Enabled = True
        lblLanVal.Caption = "L1"
        txtSo.value = ""
    End If
End Sub


'================================================================
' EVENT: cboDocType_Change
'================================================================
Private Sub cboDocType_Change()
    If mIsInitializing Then Exit Sub
    Dim dt As String
    dt = UCase(Trim(cboDocType.value))
    If dt = "" Then
        txtHoSoName.value = ""
        txtSo.value = ""
        UpdateSubType1 ""
    Else
        SetHoSoNameFromDocType dt
        SetAttachedDocsEnabled True
        TryGenerateDocNumber
        UpdateSubType1 dt
    End If
End Sub


'================================================================
' EVENT: cboIssuer_Change
' [MOD] cboIssuer bay gio la "add issuer" combo (khong luu truc tiep)
'       -> set mLastActiveList = "Issuer", KHONG goi TryGenerateDocNumber
'          (se goi sau khi nguoi dung bam nut + de them vao lstIssuers)
'================================================================
Private Sub cboIssuer_Change()
    If mIsInitializing Then Exit Sub
    mLastActiveList = "Issuer"
End Sub


'================================================================
' EVENT: cboAddRecip_Change
' [NEW] Set mLastActiveList = "Recip"
'================================================================
Private Sub cboAddRecip_Change()
    If mIsInitializing Then Exit Sub
    mLastActiveList = "Recip"
End Sub


'================================================================
' lstIssuers_Click — khi chon item trong lstIssuers:
'   bo chon lstRecipients de nut - biet xoa o dau
'================================================================
Private Sub lstIssuers_Click()
    On Error Resume Next
    lstRecipients.ListIndex = -1
    On Error GoTo 0
End Sub


'================================================================
' lstRecipients_Click — khi chon item trong lstRecipients:
'   bo chon lstIssuers de nut - biet xoa o dau
'================================================================
Private Sub lstRecipients_Click()
    On Error Resume Next
    lstIssuers.ListIndex = -1
    On Error GoTo 0
End Sub


'================================================================
' EVENT: chkNewDoc_Click
'================================================================
Private Sub chkNewDoc_Click()
    If mIsInitializing Then Exit Sub
    TryGenerateDocNumber
End Sub


'================================================================
' EVENT: btnAddRecip_Click  (NUT +)
' Them dong thoi vao CA HAI list:
'   cboIssuer  -> lstIssuers    (neu co gia tri)
'   cboAddRecip -> lstRecipients (neu co gia tri)
' Neu trung: im lang, bo qua, khong canh bao
' Neu ca 2 combo deu trong: moi thong bao
'================================================================
Private Sub btnAddRecip_Click()
    Dim issuerCode As String
    Dim recipCode  As String
    issuerCode = UCase(Trim(cboIssuer.value))
    recipCode = UCase(Trim(cboAddRecip.value))

    If issuerCode = "" And recipCode = "" Then
        MsgBox "Please select an Issuer and/or a Recipient before clicking +.", vbInformation
        Exit Sub
    End If

    '--- Them vao lstIssuers (neu co gia tri va chua trung) ---
    If issuerCode <> "" Then
        Dim isDupIssuer As Boolean: isDupIssuer = False
        Dim k As Long
        For k = 0 To lstIssuers.ListCount - 1
            If UCase(lstIssuers.List(k)) = issuerCode Then
                isDupIssuer = True: Exit For
            End If
        Next k
        If Not isDupIssuer Then lstIssuers.AddItem issuerCode
        cboIssuer.value = ""
    End If

    '--- Them vao lstRecipients (neu co gia tri va chua trung) ---
    If recipCode <> "" Then
        Dim isDupRecip As Boolean: isDupRecip = False
        Dim j As Long
        For j = 0 To lstRecipients.ListCount - 1
            If UCase(lstRecipients.List(j)) = recipCode Then
                isDupRecip = True: Exit For
            End If
        Next j
        If Not isDupRecip Then lstRecipients.AddItem recipCode
        cboAddRecip.value = ""
    End If

    TryGenerateDocNumber
End Sub


'================================================================
' EVENT: btnRemRecip_Click  (NUT -)
' Xoa 1 item khoi list dang co item duoc chon.
' Do lstIssuers_Click / lstRecipients_Click da tu bo chon list kia,
' nen moi luc chi co toi da 1 list co selection -> don gian.
'================================================================
Private Sub btnRemRecip_Click()
    On Error Resume Next
    Dim issuerSel As Long: issuerSel = lstIssuers.ListIndex
    Dim recipSel  As Long: recipSel = lstRecipients.ListIndex
    On Error GoTo 0

    If issuerSel >= 0 Then
        lstIssuers.RemoveItem issuerSel
    ElseIf recipSel >= 0 Then
        lstRecipients.RemoveItem recipSel
    Else
        MsgBox "Please click an item in the Issuer or Recipient list first, then click -.", vbInformation
        Exit Sub
    End If

    TryGenerateDocNumber
End Sub


'================================================================
' THU TAO SO TAI LIEU — TryGenerateDocNumber
'
' [MOD] Issuer lay tu lstIssuers.List(0) (phan tu dau trong list)
'       thay vi cboIssuer.Value
' Dieu kien: lstIssuers >= 1 item + DocTypeCode + >= 1 recipient
' HD (CON): dung so hop dong, skip
'================================================================
Private Sub TryGenerateDocNumber()
    On Error Resume Next

    If optHD.value Then
        txtSo.value = mContractNo
        chkNewDoc.value = True
        chkNewDoc.Enabled = False
        Exit Sub
    End If

    ' --- Lay danh sach issuers ---
    ' Uu tien lstIssuers; neu chua them control vao form thi fallback ve cboIssuer
    Dim issuerPart As String
    issuerPart = ""
    Dim issuer As String       ' phat hanh vien chinh (List(0)) de dem serial
    issuer = ""

    If lstIssuers.ListCount > 0 Then
        ' Dung lstIssuers (che do moi — nhieu issuer)
        Dim ii As Long
        For ii = 0 To lstIssuers.ListCount - 1
            Dim oneIssuer As String
            oneIssuer = UCase(Trim(lstIssuers.List(ii)))
            If oneIssuer <> "" Then
                If issuerPart <> "" Then issuerPart = issuerPart & "-"
                issuerPart = issuerPart & oneIssuer
            End If
        Next ii
        issuer = UCase(Trim(lstIssuers.List(0)))
    Else
        ' Fallback: doc tu cboIssuer (che do tuong thich — form chua co lstIssuers)
        issuer = UCase(Trim(cboIssuer.value))
        issuerPart = issuer
    End If

    If issuerPart = "" Then txtSo.value = "": Exit Sub
    If issuer = "" Then txtSo.value = "": Exit Sub

    Dim docTypeCode As String
    If optIPC.value Then
        docTypeCode = "IPC"
    ElseIf optKhac.value Then
        docTypeCode = UCase(Trim(cboDocType.value))
    Else
        txtSo.value = "": Exit Sub
    End If
    If docTypeCode = "" Then txtSo.value = "": Exit Sub

    ' Ghép APP/PKG suffix + subtype vào docTypeCode
    Dim fullDocType As String
    fullDocType = docTypeCode

    ' N?u dang nh?p cho APP ho?c PKG: chèn suffix vào gi?a
    ' VD: "APP-11" -> "IPC-APP11" | "PKG-03" -> "IPC-PKG03"
    Dim cno As String: cno = UCase(Trim(mContractNo))
    If UCase(Left(cno, 4)) = "APP-" Then
        fullDocType = fullDocType & "-" & Replace(cno, "-", "")
    ElseIf UCase(Left(cno, 4)) = "PKG-" Then
        fullDocType = fullDocType & "-" & Replace(cno, "-", "")
    End If

    On Error Resume Next
    Dim st1 As String: st1 = UCase(Trim(cboSubType1.value))
    Dim st2 As String: st2 = UCase(Trim(cboSubType2.value))
    On Error GoTo 0
    If st1 <> "" Then fullDocType = fullDocType & "-" & st1
    If st2 <> "" Then fullDocType = fullDocType & "-" & st2

    If lstRecipients.ListCount = 0 Then txtSo.value = "": Exit Sub

    ' Ghep tat ca recipients bang "-"
    Dim recipPart As String
    recipPart = ""
    Dim ri As Long
    For ri = 0 To lstRecipients.ListCount - 1
        Dim oneRecip As String
        oneRecip = UCase(Trim(lstRecipients.List(ri)))
        If oneRecip <> "" Then
            If recipPart <> "" Then recipPart = recipPart & "-"
            recipPart = recipPart & oneRecip
        End If
    Next ri
    If recipPart = "" Then txtSo.value = "": Exit Sub

    ' Don vi nhan chinh (List(0)) de dem serial
    Dim primaryRecip As String
    primaryRecip = UCase(Trim(lstRecipients.List(0)))

    Dim projCode As String
    projCode = GetGlobalSetting("PROJ")
    If Trim(projCode) = "" Then projCode = "PROJ"

    Dim dateStr As String
    dateStr = Format(Now(), "YYMMDD")

    ' ================================================================
    ' EDIT MODE: uu tien giu serial (NNN) neu so cu dung dinh dang
    '   - Co mOriginalSerial (so cu dung dinh dang): rebuild prefix, giu NNN
    '   - Khong co mOriginalSerial (so cu sai dinh dang, vi du "5"):
    '       fallthrough xuong NEW MODE de tinh serial moi binh thuong
    ' ================================================================
    If Not mIsNewMode Then
        If mOriginalSerial > 0 Then
            ' So cu dung dinh dang: giu nguyen serial, chi rebuild prefix
            If mOriginalSuffix <> "" Then
                txtSo.value = dateStr & "." & projCode & "." & issuerPart & "." & _
                              fullDocType & "." & recipPart & "." & Format(mOriginalSerial, "000") & "." & mOriginalSuffix
            Else
                txtSo.value = dateStr & "." & projCode & "." & issuerPart & "." & _
                              fullDocType & "." & recipPart & "." & Format(mOriginalSerial, "000")
            End If
            Exit Sub
        End If
        ' So cu sai dinh dang (khong co NNN hop le): tiep tuc xuong de tao so moi
    End If

    ' ================================================================
    ' NEW MODE: tinh serial tiep theo theo docType + issuer + recipient
    ' ================================================================
    Dim lastSerial As Long
        lastSerial = GetLastDocSerial(issuer, docTypeCode, primaryRecip, mContractID)

    If lastSerial = 0 Then
        chkNewDoc.value = True
        chkNewDoc.Enabled = False
        txtSo.value = dateStr & "." & projCode & "." & issuerPart & "." & _
                      fullDocType & "." & recipPart & ".001"
    Else
        chkNewDoc.Enabled = True

        If chkNewDoc.value Then
            Dim newSerial As Long
            newSerial = lastSerial + 1
            txtSo.value = dateStr & "." & projCode & "." & issuerPart & "." & _
                          fullDocType & "." & recipPart & "." & Format(newSerial, "000")
        Else
            Dim verSuffix As String
            verSuffix = GetNextVersionSuffix(issuer, docTypeCode, primaryRecip, lastSerial, mContractID)
            'verSuffix = GetNextVersionSuffix(issuer, docTypeCode, primaryRecip, lastSerial)
            txtSo.value = dateStr & "." & projCode & "." & issuerPart & "." & _
                          fullDocType & "." & recipPart & "." & Format(lastSerial, "000") & "." & verSuffix
        End If
    End If

    On Error GoTo 0
End Sub


'================================================================
' GET LAST DOC SERIAL
'================================================================
Private Function GetLastDocSerial(ByVal issuer As String, _
                                   ByVal docType As String, _
                                   ByVal recipient As String, _
                                   Optional ByVal contractID As String = "") As Long
    GetLastDocSerial = 0
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = GetDataSheet("DOCUMENT_TRACKER")
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim maxSerial As Long: maxSerial = 0
    Dim i As Long

    For i = 2 To lastRow
        ' L?c theo ContractID (col B) n?u du?c truy?n vào
        If contractID <> "" Then
            If Trim(CStr(ws.Cells(i, "B").value)) <> contractID Then GoTo SKIP_LS
        End If

        ' Kiem tra IssuerCode (col Q)
        Dim cellIssuerFull As String
        cellIssuerFull = Trim(CStr(ws.Cells(i, "Q").value))
        Dim cellPrimaryIssuer As String
        If InStr(cellIssuerFull, ";") > 0 Then
            cellPrimaryIssuer = UCase(Trim(Split(cellIssuerFull, ";")(0)))
        Else
            cellPrimaryIssuer = UCase(Trim(cellIssuerFull))
        End If
        If cellPrimaryIssuer <> UCase(issuer) Then GoTo SKIP_LS

        ' Kiem tra DocType (col C)
        If UCase(Trim(CStr(ws.Cells(i, "C").value))) <> UCase(docType) Then GoTo SKIP_LS

        ' Kiem tra PrimaryRecipient (col R)
        Dim cellRecipFull As String
        cellRecipFull = Trim(CStr(ws.Cells(i, "R").value))
        Dim cellPrimary As String
        If InStr(cellRecipFull, ";") > 0 Then
            cellPrimary = UCase(Trim(Split(cellRecipFull, ";")(0)))
        Else
            cellPrimary = UCase(Trim(cellRecipFull))
        End If
        If cellPrimary <> UCase(recipient) Then GoTo SKIP_LS

        ' Trich serial tu col D (SoTaiLieu)
        Dim docNo As String
        docNo = Trim(CStr(ws.Cells(i, "D").value))
        If docNo = "" Then GoTo SKIP_LS

        Dim parts() As String
        parts = Split(docNo, ".")
        Dim pi As Integer
        For pi = 0 To UBound(parts)
            If Len(parts(pi)) = 3 Then
                Dim sn As Long: sn = 0
                On Error Resume Next
                sn = CLng(parts(pi))
                On Error GoTo 0
                If sn > 0 Then
                    If sn > maxSerial Then maxSerial = sn
                    Exit For
                End If
            End If
        Next pi
SKIP_LS:
    Next i

    GetLastDocSerial = maxSerial
    On Error GoTo 0
End Function


'================================================================
' GET NEXT VERSION SUFFIX
'================================================================
Private Function GetNextVersionSuffix(ByVal issuer As String, _
                                       ByVal docType As String, _
                                       ByVal recipient As String, _
                                       ByVal targetSerial As Long, _
                                       Optional ByVal contractID As String = "") As String
    GetNextVersionSuffix = "A"
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = GetDataSheet("DOCUMENT_TRACKER")
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim usedSuffixes As String
    usedSuffixes = ""
    Dim i As Long

    For i = 2 To lastRow
        ' L?c theo ContractID (col B)
        If contractID <> "" Then
            If Trim(CStr(ws.Cells(i, "B").value)) <> contractID Then GoTo SKIP_VS
        End If

        Dim cellIssuerFull As String
        cellIssuerFull = Trim(CStr(ws.Cells(i, "Q").value))
        Dim cellPrimaryIssuer As String
        If InStr(cellIssuerFull, ";") > 0 Then
            cellPrimaryIssuer = UCase(Trim(Split(cellIssuerFull, ";")(0)))
        Else
            cellPrimaryIssuer = UCase(Trim(cellIssuerFull))
        End If
        If cellPrimaryIssuer <> UCase(issuer) Then GoTo SKIP_VS

        If UCase(Trim(CStr(ws.Cells(i, "C").value))) <> UCase(docType) Then GoTo SKIP_VS

        Dim cellRecipFull As String
        cellRecipFull = Trim(CStr(ws.Cells(i, "R").value))
        Dim cellPrimary As String
        If InStr(cellRecipFull, ";") > 0 Then
            cellPrimary = UCase(Trim(Split(cellRecipFull, ";")(0)))
        Else
            cellPrimary = UCase(Trim(cellRecipFull))
        End If
        If cellPrimary <> UCase(recipient) Then GoTo SKIP_VS

        Dim docNo As String
        docNo = Trim(CStr(ws.Cells(i, "D").value))
        If docNo = "" Then GoTo SKIP_VS

        Dim parts() As String
        parts = Split(docNo, ".")
        Dim serialIdx As Integer: serialIdx = -1
        Dim pi As Integer
        For pi = 0 To UBound(parts)
            If Len(parts(pi)) = 3 Then
                Dim sn As Long: sn = 0
                On Error Resume Next
                sn = CLng(parts(pi))
                On Error GoTo 0
                If sn = targetSerial Then
                    serialIdx = pi
                    Exit For
                End If
            End If
        Next pi

        If serialIdx >= 0 And serialIdx < UBound(parts) Then
            Dim suf As String
            suf = UCase(Trim(parts(serialIdx + 1)))
            If Len(suf) = 1 And suf >= "A" And suf <= "Z" Then
                If InStr(usedSuffixes, suf) = 0 Then
                    usedSuffixes = usedSuffixes & suf
                End If
            End If
        End If
SKIP_VS:
    Next i

    Dim n As Integer
    For n = 0 To 25
        Dim ch As String
        ch = Chr(65 + n)
        If InStr(usedSuffixes, ch) = 0 Then
            GetNextVersionSuffix = ch
            Exit For
        End If
    Next n

    On Error GoTo 0
End Function


'================================================================
' btnNgayNhanToday
'================================================================
Private Sub btnNgayNhanToday_Click()
    txtNgayNhan.value = Format(Now(), "dd/mm/yyyy")
    btnNgayTraToday.Enabled = True
End Sub


'================================================================
' btnNgayTraToday
'================================================================
Private Sub btnNgayTraToday_Click()
    txtNgayTra.value = Format(Now(), "dd/mm/yyyy")
    UpdateTinhTrangForNgayTra   ' [MOD] programmatic set khong fire _Change
End Sub


'================================================================
' HELPER: Cap nhat TinhTrang khi NgayTra duoc dien
' Group A -> Completed  |  Khac -> Accepted
'================================================================
Private Sub UpdateTinhTrangForNgayTra()
    If mIsInitializing Then Exit Sub
    If Trim(txtNgayTra.value) = "" Then Exit Sub
    On Error Resume Next
    If GetContractGroup(mContractID) = "A" Then
        cboTinhTrang.value = "Completed"
    Else
        cboTinhTrang.value = "Accepted"
    End If
    On Error GoTo 0
End Sub


'================================================================
' EVENT: txtNgayTra_Change
' Khi nguoi dung nhap ngay tra -> tu dong chuyen TinhTrang
'================================================================
Private Sub txtNgayTra_Change()
    UpdateTinhTrangForNgayTra
End Sub


'================================================================
' NUT LUU
' [MOD] Luu lstIssuers vao col Q (phan cach ";")
'       Luu lstRecipients vao col R (phan cach ";")
'================================================================
Private Sub btnSave_Click()

    On Error GoTo ERR_SAVE

    Dim hoSo As String
    hoSo = GetHoSoValue()

    If hoSo = "" Then
        ShowMsgCenter "Please select a document type (HD / IPC / Other)!"
        optHD.SetFocus
        Exit Sub
    End If

    If hoSo = "Khac" Then
        Dim storeVal As String
        storeVal = GetHoSoStoreValue()
        If storeVal = "Other" Or storeVal = "" Then
            If Trim(txtHoSoName.value) = "" Then
                ShowMsgCenter "Please select a document type or enter a name."
                cboDocType.SetFocus
                Exit Sub
            End If
        End If
    End If

    ' Kiem tra: phai co it nhat 1 issuer trong lstIssuers (tru truong hop HD)
    If hoSo <> "CON" Then
        If lstIssuers.ListCount = 0 Then
            ShowMsgCenter "Please add at least one Issuer."
            cboIssuer.SetFocus
            Exit Sub
        End If
    End If

    If hoSo <> "CON" Then
        If Trim(txtSo.value) = "" Then
            ShowMsgCenter "Document number not generated yet." & vbCrLf & _
                          "Please add at least one Issuer and one Recipient."
            cboIssuer.SetFocus
            Exit Sub
        End If
    End If

    If Trim(txtLink.value) = "" Then
        If MsgBox("No link entered. Save anyway?", vbQuestion + vbYesNo) = vbNo Then
            txtLink.SetFocus
            Exit Sub
        End If
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetDataSheet("DOCUMENT_TRACKER")
    On Error GoTo ERR_SAVE
    If ws Is Nothing Then
        ShowMsgCenter "Sheet DOCUMENT_TRACKER not found!"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim targetRow As Long

    If mIsNewMode Then
        targetRow = lastRow + 1
        mDocumentID = GenerateDocumentID(ws, lastRow)
        ws.Cells(targetRow, "A").value = mDocumentID
        ws.Cells(targetRow, "B").value = mContractID
    Else
        targetRow = 0
        Dim i As Long
        For i = 2 To lastRow
            If Trim(CStr(ws.Cells(i, "A").value)) = Trim(mDocumentID) Then
                targetRow = i: Exit For
            End If
        Next i
        If targetRow = 0 Then
            ShowMsgCenter "Record not found: " & mDocumentID
            Exit Sub
        End If
    End If

    ws.Cells(targetRow, "C").value = GetHoSoStoreValue()
    ws.Cells(targetRow, "D").value = Trim(txtSo.value)
    ws.Cells(targetRow, "E").value = Trim(txtHoSoName.value)
    ws.Cells(targetRow, "F").value = Trim(txtLink.value)

    If mIsNewMode Then
        If lblTuNgay.Caption <> "" And IsDate(lblTuNgay.Caption) Then
            ws.Cells(targetRow, "G").value = CDate(lblTuNgay.Caption)
            ws.Cells(targetRow, "G").NumberFormat = "dd/mm/yyyy"
        End If
        If lblDenNgay.Caption <> "" And IsDate(lblDenNgay.Caption) Then
            ws.Cells(targetRow, "H").value = CDate(lblDenNgay.Caption)
            ws.Cells(targetRow, "H").NumberFormat = "dd/mm/yyyy"
        End If
    End If

    If hoSo <> "HD" Then
        If Trim(txtNgayNhan.value) <> "" Then
            If IsDate(txtNgayNhan.value) Then
                ws.Cells(targetRow, "I").value = CDate(txtNgayNhan.value)
                ws.Cells(targetRow, "I").NumberFormat = "dd/mm/yyyy"
            Else
                ws.Cells(targetRow, "I").value = Trim(txtNgayNhan.value)
            End If
        Else
            ws.Cells(targetRow, "I").value = ""
        End If

        If Trim(txtNgayTra.value) <> "" Then
            If IsDate(txtNgayTra.value) Then
                ws.Cells(targetRow, "J").value = CDate(txtNgayTra.value)
                ws.Cells(targetRow, "J").NumberFormat = "dd/mm/yyyy"
            Else
                ws.Cells(targetRow, "J").value = Trim(txtNgayTra.value)
            End If
        Else
            ws.Cells(targetRow, "J").value = ""
        End If

        ws.Cells(targetRow, "K").value = Trim(cboNguoiNhan.value)
    End If

    ws.Cells(targetRow, "L").value = 1
    ws.Cells(targetRow, "M").value = EncodeHoSoKem()
    ws.Cells(targetRow, "N").value = lblLanVal.Caption

    If hoSo <> "HD" Then
        ws.Cells(targetRow, "O").value = Trim(cboTinhTrang.value)
    End If

    ws.Cells(targetRow, "P").value = Trim(txtGhiChu.value)

    ' [MOD] Col Q: IssuerCode(s) tu lstIssuers, phan cach ";"
    Dim issuerList As String
    issuerList = ""
    Dim ii As Long
    For ii = 0 To lstIssuers.ListCount - 1
        If ii > 0 Then issuerList = issuerList & ";"
        issuerList = issuerList & UCase(Trim(lstIssuers.List(ii)))
    Next ii
    ws.Cells(targetRow, "Q").value = issuerList

    ' Col R: Recipients phan cach ";"
    Dim recipList As String
    recipList = ""
    Dim ri As Long
    For ri = 0 To lstRecipients.ListCount - 1
        If ri > 0 Then recipList = recipList & ";"
        recipList = recipList & UCase(Trim(lstRecipients.List(ri)))
    Next ri
    ws.Cells(targetRow, "R").value = recipList
    ' Col S: cboSubType1,  Col T: cboSubType2
    ws.Cells(targetRow, "S").value = UCase(Trim(cboSubType1.value))
    ws.Cells(targetRow, "T").value = UCase(Trim(cboSubType2.value))
    BuildDocTrackerCache
    
    ShowMsgCenter "Saved successfully!"
    NotifyAndClose
    Exit Sub

ERR_SAVE:
    ShowMsgCenter "Save Error: " & Err.Description
End Sub


'================================================================
' Ma hoa HoSo Kem thanh chuoi luu sheet
'================================================================
Private Function EncodeHoSoKem() As String
    Dim parts(0 To 7) As String
    Dim cnt As Integer: cnt = 0
    If chkGTTT.value And cboGTTT.value <> "" Then parts(cnt) = "GTTT:" & cboGTTT.value:        cnt = cnt + 1
    If chkTHKL.value And cboTHKL.value <> "" Then parts(cnt) = "THKL:" & cboTHKL.value:        cnt = cnt + 1
    If chkDGKL.value And cboDGKL.value <> "" Then parts(cnt) = "DGKL:" & cboDGKL.value:        cnt = cnt + 1
    If chkNTCL.value And cboNTCL.value <> "" Then parts(cnt) = "NTCL:" & cboNTCL.value:        cnt = cnt + 1
    If chkNTVL.value And cboNTVL.value <> "" Then parts(cnt) = "NTVL:" & cboNTVL.value:        cnt = cnt + 1
    If chkHC.value And cboHC.value <> "" Then parts(cnt) = "HC:" & cboHC.value:                 cnt = cnt + 1
    If chkHSKhac.value And cboHSKhac.value <> "" Then parts(cnt) = "HSK:" & cboHSKhac.value:   cnt = cnt + 1
    Dim result As String: result = ""
    Dim j As Integer
    For j = 0 To cnt - 1
        If result <> "" Then result = result & "|"
        result = result & parts(j)
    Next j
    If Trim(txtHoSoKemNote.value) <> "" Then
        If result <> "" Then result = result & "|"
        result = result & "NOTE:" & Trim(txtHoSoKemNote.value)
    End If
    EncodeHoSoKem = result
End Function


'================================================================
' Tao Document ID noi bo
'================================================================
Private Function GenerateDocumentID( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long) As String

    Dim prefix As String
    prefix = "DOC-" & Format(Now, "YYYYMMDD") & "-"
    Dim maxNum As Long: maxNum = 0
    Dim i As Long, cellVal As String, numPart As Long
    For i = 2 To lastRow
        cellVal = Trim(CStr(ws.Cells(i, "A").value))
        If Left$(cellVal, Len(prefix)) = prefix Then
            On Error Resume Next
            numPart = CLng(mID$(cellVal, Len(prefix) + 1))
            On Error GoTo 0
            If numPart > maxNum Then maxNum = numPart
        End If
    Next i
    GenerateDocumentID = prefix & Format(maxNum + 1, "000")
End Function


'================================================================
' NUT REVERSE (nut doi chieu) — hoan doi noi dung lstIssuers va lstRecipients
' Dung khi can doi vai tro phat hanh / tiep nhan
' *** THEM VAO FORM DESIGNER: CommandButton ten btnReverse, Caption = "?" ***
'================================================================
Private Sub btnReverse_Click()
    On Error Resume Next

    ' --- Sao chep lstIssuers ra mang tam ---
    Dim issCount As Long: issCount = lstIssuers.ListCount
    Dim recCount As Long: recCount = lstRecipients.ListCount

    If issCount = 0 And recCount = 0 Then
        MsgBox "Both Issuer and Recipient lists are empty. Nothing to reverse.", vbInformation
        Exit Sub
    End If

    Dim issArr() As String
    ReDim issArr(0 To IIf(issCount > 0, issCount - 1, 0))
    Dim i As Long
    For i = 0 To issCount - 1
        issArr(i) = lstIssuers.List(i)
    Next i

    Dim recArr() As String
    ReDim recArr(0 To IIf(recCount > 0, recCount - 1, 0))
    Dim j As Long
    For j = 0 To recCount - 1
        recArr(j) = lstRecipients.List(j)
    Next j

    ' --- Ghi noi dung cua lstRecipients vao lstIssuers ---
    lstIssuers.Clear
    For j = 0 To recCount - 1
        lstIssuers.AddItem recArr(j)
    Next j

    ' --- Ghi noi dung cu cua lstIssuers vao lstRecipients ---
    lstRecipients.Clear
    For i = 0 To issCount - 1
        lstRecipients.AddItem issArr(i)
    Next i

    TryGenerateDocNumber
    On Error GoTo 0
End Sub


'================================================================
' NUT BROWSE (nut "...")
'
' Logic:
'   1. Neu txtSo co gia tri -> tim file co ten chua so tai lieu
'      trong thu muc du an (DOC_FOLDER setting hoac thu muc workbook)
'      - Tim thay duy nhat 1 file -> hoi xac nhan, neu Yes thi dien luon
'      - Tim thay nhieu file     -> hien danh sach, nguoi dung chon so thu tu
'      - Khong tim thay          -> mo dialog chon file binh thuong
'   2. Neu txtSo trong           -> mo dialog chon file binh thuong
'================================================================
Private Sub btnBrowse_Click()
    On Error GoTo ERR_BROWSE

    '--- [MOD] Uu tien: neu txtLink da co gia tri -> mo thu muc chua link do ---
    Dim existingLink As String
    existingLink = Trim(txtLink.value)
    If existingLink <> "" Then
        Dim fso0 As Object
        Set fso0 = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        If fso0.FileExists(existingLink) Then
            Shell "explorer.exe /select,""" & existingLink & """", vbNormalFocus
            On Error GoTo 0
            Application.Wait Now + TimeSerial(0, 0, 1)
            On Error Resume Next
            AppActivate "File Explorer"
            If Err.Number <> 0 Then AppActivate "Windows Explorer"
            On Error GoTo 0
            Exit Sub
        ElseIf fso0.FolderExists(existingLink) Then
            Shell "explorer.exe """ & existingLink & """", vbNormalFocus
            On Error GoTo 0
            Application.Wait Now + TimeSerial(0, 0, 1)
            On Error Resume Next
            AppActivate "File Explorer"
            If Err.Number <> 0 Then AppActivate "Windows Explorer"
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo ERR_BROWSE
    End If

    Dim docNo As String
    docNo = Trim(txtSo.value)

    '--- Co so tai lieu -> thu tim file khop ---
    If docNo <> "" Then
        Dim rootPath As String
        rootPath = Trim(GetGlobalSetting("DOC_FOLDER"))
        If rootPath = "" Then rootPath = Left(ThisWorkbook.fullName, _
            InStrRev(ThisWorkbook.fullName, Application.PathSeparator))

        If rootPath <> "" Then
            Dim matches As Collection
            Set matches = FindFilesByDocNumber(rootPath, docNo, 4)

            If matches.count = 1 Then
                '--- Duy nhat 1 ket qua: xac nhan nhanh ---
                Dim ans As VbMsgBoxResult
                ans = MsgBox("Found matching file:" & vbCrLf & vbCrLf & _
                             matches(1) & vbCrLf & vbCrLf & _
                             "Use this file?", vbQuestion + vbYesNo, "Auto-match")
                If ans = vbYes Then
                    txtLink.value = matches(1)
                    Exit Sub
                End If
                ' Nguoi dung chon No -> roi xuong mo dialog binh thuong

            ElseIf matches.count > 1 Then
                '--- Nhieu ket qua: hien danh sach de chon ---
                Dim listMsg As String
                listMsg = "Found " & matches.count & " matching files." & vbCrLf & _
                          "Enter the number to select, or 0 to browse manually:" & vbCrLf & vbCrLf
                Dim m As Long
                For m = 1 To Application.Min(matches.count, 15)
                    listMsg = listMsg & m & ". " & matches(m) & vbCrLf
                Next m
                If matches.count > 15 Then listMsg = listMsg & "..." & vbCrLf

                Dim choice As String
                choice = InputBox(listMsg, "Select File", "1")
                If choice = "" Then Exit Sub         ' Bam Cancel
                Dim choiceNum As Long
                choiceNum = val(choice)
                If choiceNum >= 1 And choiceNum <= matches.count Then
                    txtLink.value = matches(choiceNum)
                    Exit Sub
                End If
                ' 0 hoac ngoai khoang -> roi xuong mo dialog
            End If
            ' Count = 0 -> khong tim thay -> mo dialog
        End If
    End If

    '--- Mo dialog chon file binh thuong ---
    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Select file"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        .Filters.Add "PDF", "*.pdf"
        .Filters.Add "Word", "*.doc;*.docx"
        .Filters.Add "Excel", "*.xls;*.xlsx"
        If Trim(GetGlobalSetting("DOC_FOLDER")) <> "" Then
            On Error Resume Next
            .InitialFileName = Trim(GetGlobalSetting("DOC_FOLDER"))
            On Error GoTo ERR_BROWSE
        End If
        If .Show = True Then txtLink.value = .SelectedItems(1)
    End With
    Exit Sub

ERR_BROWSE:
    On Error Resume Next
    Dim fd2 As Object
    Set fd2 = Application.FileDialog(msoFileDialogFolderPicker)
    With fd2
        .title = "Select folder"
        If .Show = True Then txtLink.value = .SelectedItems(1)
    End With
End Sub


'================================================================
' TIM FILE THEO SO TAI LIEU — FindFilesByDocNumber
'
' Quyet duong trong rootPath (va cac thu muc con den maxDepth tang)
' Tim tat ca file co ten chua docNo (bat ky vi tri, khong phan biet hoa thuong)
' Tra ve Collection cac duong dan day du
'================================================================
Private Function FindFilesByDocNumber(ByVal rootPath As String, _
                                      ByVal docNo As String, _
                                      ByVal maxDepth As Integer) As Collection
    Dim result As Collection
    Set result = New Collection

    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then
        Set FindFilesByDocNumber = result
        Exit Function
    End If

    ' Chuan hoa duong dan
    If Right(rootPath, 1) <> Application.PathSeparator Then
        rootPath = rootPath & Application.PathSeparator
    End If

    SearchFolder fso, rootPath, UCase(docNo), maxDepth, 0, result
    Set FindFilesByDocNumber = result
    On Error GoTo 0
End Function


'================================================================
' TIM DE QUY THU MUC — SearchFolder (helper cho FindFilesByDocNumber)
'================================================================
Private Sub SearchFolder(ByVal fso As Object, _
                         ByVal folderPath As String, _
                         ByVal docNoUC As String, _
                         ByVal maxDepth As Integer, _
                         ByVal curDepth As Integer, _
                         ByRef result As Collection)
    On Error Resume Next
    If curDepth > maxDepth Then Exit Sub

    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    If folder Is Nothing Then Exit Sub

    ' Duyet file trong thu muc hien tai
    Dim f As Object
    For Each f In folder.Files
        Dim fName As String
        fName = UCase(f.Name)
        ' Kiem tra ten file chua so tai lieu (bo dau phan cach de tang kha nang khop)
        Dim docNorm As String
        docNorm = Replace(Replace(docNoUC, ".", ""), "-", "")
        Dim fnameNorm As String
        fnameNorm = Replace(Replace(fName, ".", ""), "-", "")
        If InStr(fName, docNoUC) > 0 Or InStr(fnameNorm, docNorm) > 0 Then
            result.Add f.path
        End If
    Next f

    ' De quy vao thu muc con
    If curDepth < maxDepth Then
        Dim subFolder As Object
        For Each subFolder In folder.SubFolders
            SearchFolder fso, subFolder.path, docNoUC, maxDepth, curDepth + 1, result
        Next subFolder
    End If
    On Error GoTo 0
End Sub


'================================================================
' NUT X (dong form)
'================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        NotifyAndClose
        Unload Me
        
    End If
    'ApplyTheme Me
    UnregisterForm Me
End Sub


'================================================================
' C?p nh?t cboSubType1 theo cboDocType
'================================================================
Private Sub UpdateSubType1(ByVal dt As String)
    On Error Resume Next
    cboSubType1.Clear
    cboSubType2.Clear
    cboSubType1.value = ""
    cboSubType2.value = ""
    cboSubType2.Enabled = False

    Select Case UCase(Trim(dt))
        Case "RFA"
            With cboSubType1
                .AddItem "ITP"
                .AddItem "MSA"
                .AddItem "HSE"
                .AddItem "SCH"
                .AddItem "GEN"
                .AddItem "MS"
                .AddItem "SD"
            End With
            cboSubType1.Enabled = True

        Case "MOM"
            With cboSubType1
                .AddItem "HSE"
                .AddItem "MEP"
                .AddItem "KOM"
            End With
            cboSubType1.Enabled = True

        Case Else
            cboSubType1.Enabled = False
    End Select
    On Error GoTo 0
End Sub


'================================================================
' C?p nh?t cboSubType2 theo cboSubType1
'================================================================
Private Sub UpdateSubType2(ByVal st1 As String)
    On Error Resume Next
    cboSubType2.Clear
    cboSubType2.value = ""

    Select Case UCase(Trim(st1))
        Case "GEN"
            cboSubType2.AddItem "FS"
            cboSubType2.AddItem "E"
            cboSubType2.Enabled = True

        Case "MSA"
            cboSubType2.AddItem "HVAC"
            cboSubType2.AddItem "E"
            cboSubType2.AddItem "FS"
            cboSubType2.AddItem "PL"
            cboSubType2.AddItem "FF"
            cboSubType2.Enabled = True

        Case "ITP"
            cboSubType2.AddItem "HVAC"
            cboSubType2.AddItem "E"
            cboSubType2.AddItem "PL"
            cboSubType2.AddItem "FF"
            cboSubType2.Enabled = True

        Case "SCH"
            cboSubType2.AddItem "E"
            cboSubType2.Enabled = True

        Case "MS"
            cboSubType2.AddItem "HVAC"
            cboSubType2.AddItem "E"
            cboSubType2.AddItem "PL"
            cboSubType2.AddItem "FF"
            cboSubType2.Enabled = True

        Case "SD"
            cboSubType2.AddItem "HVAC"
            cboSubType2.AddItem "E"
            cboSubType2.AddItem "ELV"
            cboSubType2.AddItem "FA"
            cboSubType2.AddItem "PL"
            cboSubType2.AddItem "FF"
            cboSubType2.Enabled = True

        Case Else
            ' HSE, MEP, KOM — không có cbo2
            cboSubType2.Enabled = False
    End Select
    On Error GoTo 0
End Sub
'================================================================
' EVENT: cboSubType1_Change
'================================================================
Private Sub cboSubType1_Change()
    If mIsInitializing Then Exit Sub
    UpdateSubType2 cboSubType1.value
    TryGenerateDocNumber
End Sub

Private Sub cboSubType2_Change()
    If mIsInitializing Then Exit Sub
    TryGenerateDocNumber
End Sub






