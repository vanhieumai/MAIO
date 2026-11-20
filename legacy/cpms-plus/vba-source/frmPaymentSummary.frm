Attribute VB_Name = "frmPaymentSummary"
Attribute VB_Base = "0{3D01F4AF-B18E-4FE0-8B75-D92EFECFA58D}{D85F33A3-03E5-45DB-A166-7B65F4B81D21}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'frmpaymentsummary
Option Explicit

Private suppressContractChange As Boolean
Private isLoading As Boolean
Private suppressContractClick As Boolean
Private suppressDNChange As Boolean
Private mDNRightClickTime As Double
Private mPayMonth As Long
Private mPayYear  As Long
Private mCategoryFilter As String

Private mReminderForm As frmReminder
Private mReminderRef  As Object

Public Sub LoadData()

    '========================
    ' BUILD CACHE FIRST
    '========================
    If Not gPaymentLoaded Then

        BuildPaymentCache

    End If

    '========================
    ' RESET BUTTON STATE
    '========================
    cmdAddInternalCost.Caption = _
        "Add Internal Cost"

    cmdAddInternalCost.Tag = ""

    'btnPaymentEntry.Enabled = False

    lstUnpaidContract.ListIndex = -1
    
    '========================
    ' N?P DANH SÁCH H?NG M?C (M?I)
    '========================
    PopulateCategoryCombo
    '========================
    ' LOAD UI
    '========================
    LoadGroupSummary

    LoadUnpaidContracts

End Sub


Private Sub LoadGroupSummary()

    Dim i As Long
    Dim r As Long

    Dim grp As String
    Dim contractID As String
    Dim statusText As String
    Dim catVal As String

    Dim estA As Double
    Dim totalA As Double
    Dim actA As Double
    Dim accA As Double
    Dim reqA As Double
    Dim paidA As Double
    Dim remA As Double

    Dim estB As Double
    Dim totalB As Double
    Dim actB As Double
    Dim accB As Double
    Dim reqB As Double
    Dim paidB As Double
    Dim remB As Double

    Dim paidA_alltime As Double
    Dim paidB_alltime As Double

    '========================
    ' RESET
    '========================
    lblGroupA_Estimated.Caption = "0"
    lblGroupA_Total.Caption = "0"
    lblGroupA_ActualVolume.Caption = "0"
    lblGroupA_AcceptedVolume.Caption = "0"
    lblGroupA_Requested.Caption = "0"
    lblGroupA_Paid.Caption = "0"
    lblGroupA_Remain.Caption = "0"

    lblGroupB_Estimated.Caption = "0"
    lblGroupB_Total.Caption = "0"
    lblGroupB_ActualVolume.Caption = "0"
    lblGroupB_AcceptedVolume.Caption = "0"
    lblGroupB_Requested.Caption = "0"
    lblGroupB_Paid.Caption = "0"
    lblGroupB_Remain.Caption = "0"

    lbl_PaidRemaining.Caption = "0"

    If IsEmpty(gArrPayment) Then Exit Sub

    '========================
    ' Ð?C TIME FILTER
    '========================
    Dim useMonthly As Boolean, useYearly As Boolean
    On Error Resume Next
    useMonthly = optPayMonthly.value
    useYearly = optPayYearly.value
    On Error GoTo 0
    Dim useTimeFilter As Boolean
    useTimeFilter = (useMonthly Or useYearly)

    '========================
    ' Ð?C CATEGORY FILTER (M?I) — ngu?n: gArrContract c?t R
    '========================
    Dim catFilter As String
    catFilter = GetCurrentCategoryFilter()
    Dim useCatFilter As Boolean
    useCatFilter = (catFilter <> "")

    Dim dCatByCID As Object
    Set dCatByCID = BuildCategoryDictFromContract()

    '========================
    ' LOOP CONTRACT HEADER (trong gArrPayment)
    '========================
    For i = 1 To UBound(gArrPayment, 1)

        contractID = Trim$(CStr(gArrPayment(i, 3)))

        If (Left$(contractID, 2) = "A-" Or Left$(contractID, 2) = "B-") _
        And Trim$(CStr(gArrPayment(i, 2))) = "" Then

            ' Skip CANCEL
            If gDictContract.Exists(contractID) Then
                r = gDictContract(contractID)
                statusText = UCase$(Trim$(CStr(gArrContract(r, 6))))
                If statusText = "CANCEL" Then GoTo NEXT_I
            End If

            grp = Trim$(CStr(gArrPayment(i, 1)))
            If grp <> "GROUP-A" And grp <> "GROUP-B" Then GoTo NEXT_I

            ' ---- L?c Category (M?I, tra tu dict Contract-DB) ----
            If useCatFilter Then
                catVal = ""
                If dCatByCID.Exists(contractID) Then catVal = CStr(dCatByCID(contractID))
                If LCase(catVal) <> LCase(catFilter) Then GoTo NEXT_I
            End If

            Dim exRateH As Double
            exRateH = val(gArrPayment(i, 7))
            If exRateH <= 0 Then exRateH = 1

            If grp = "GROUP-A" Then
                estA = estA + val(gArrPayment(i, 8)) * exRateH     ' col H – g?c ? nhân t? giá OK
                totalA = totalA + val(gArrPayment(i, 9)) * exRateH  ' col I – g?c ? nhân t? giá OK
                paidA_alltime = paidA_alltime + val(gArrPayment(i, 13)) ' col M – dã VND ? KHÔNG nhân
                If Not useTimeFilter Then
                    actA = actA + val(gArrPayment(i, 10))    ' col J – dã VND
                    accA = accA + val(gArrPayment(i, 11))    ' col K – dã VND
                    reqA = reqA + val(gArrPayment(i, 12))    ' col L – dã VND
                    paidA = paidA + val(gArrPayment(i, 13))  ' col M – dã VND
                    remA = remA + val(gArrPayment(i, 14))    ' col N – dã VND
                End If
            ElseIf grp = "GROUP-B" Then
                estB = estB + val(gArrPayment(i, 8)) * exRateH
                totalB = totalB + val(gArrPayment(i, 9)) * exRateH
                paidB_alltime = paidB_alltime + val(gArrPayment(i, 13))
                If Not useTimeFilter Then
                    actB = actB + val(gArrPayment(i, 10))
                    accB = accB + val(gArrPayment(i, 11))
                    reqB = reqB + val(gArrPayment(i, 12))
                    paidB = paidB + val(gArrPayment(i, 13))
                    remB = remB + val(gArrPayment(i, 14))
                End If
            End If

        End If

NEXT_I:
    Next i

    '========================
    ' TIME FILTER: t?ng t? payment rows trong k?
    '========================
    If useTimeFilter Then

        Dim dGrp As Object
        Set dGrp = CreateObject("Scripting.Dictionary")
        Dim hCID As String, hGrp As String
        For i = 1 To UBound(gArrPayment, 1)
            hCID = Trim$(CStr(gArrPayment(i, 3)))
            If (Left$(hCID, 2) = "A-" Or Left$(hCID, 2) = "B-") _
            And Trim$(CStr(gArrPayment(i, 2))) = "" _
            And Not dGrp.Exists(hCID) Then

                Dim hst As String: hst = ""
                If gDictContract.Exists(hCID) Then
                    hst = UCase$(Trim$(CStr(gArrContract(gDictContract(hCID), 6))))
                End If

                ' ---- L?c Category cho contract này (M?I, tra tu dict) ----
                Dim hCat As String: hCat = ""
                If dCatByCID.Exists(hCID) Then hCat = CStr(dCatByCID(hCID))

                If hst <> "CANCEL" _
                And (Not useCatFilter Or LCase(hCat) = LCase(catFilter)) Then
                    If Left$(hCID, 2) = "A-" Then
                        dGrp(hCID) = "GROUP-A"
                    Else
                        dGrp(hCID) = "GROUP-B"
                    End If
                End If
            End If
        Next i

        Dim pCID As String, pgrp As String
        Dim payDate As Variant
        Dim SkipRow As Boolean
        Dim pd As Date, pd2 As Date

        For i = 1 To UBound(gArrPayment, 1)
            If Trim$(CStr(gArrPayment(i, 2))) <> "" Then

                payDate = gArrPayment(i, 17)
                SkipRow = True

                If IsValidDMY(payDate) Then
                    pd = ParseDMY(payDate)
                    If useMonthly Then
                        If Month(pd) = mPayMonth And Year(pd) = mPayYear Then SkipRow = False
                    Else
                        If Year(pd) = mPayYear Then SkipRow = False
                    End If
                ElseIf IsDate(payDate) Then
                    pd2 = CDate(payDate)
                    If useMonthly Then
                        If Month(pd2) = mPayMonth And Year(pd2) = mPayYear Then SkipRow = False
                    Else
                        If Year(pd2) = mPayYear Then SkipRow = False
                    End If
                End If

                If SkipRow Then GoTo NEXT_PAY

                pCID = Trim$(CStr(gArrPayment(i, 3)))
                If dGrp.Exists(pCID) Then
                    pgrp = dGrp(pCID)
                    Dim exRateP As Double
                    exRateP = val(gArrPayment(i, 7))
                    If exRateP <= 0 Then exRateP = 1
                    If pgrp = "GROUP-A" Then
                        actA = actA + val(gArrPayment(i, 10)) * exRateP
                        accA = accA + val(gArrPayment(i, 11)) * exRateP
                        reqA = reqA + val(gArrPayment(i, 12)) * exRateP
                        paidA = paidA + val(gArrPayment(i, 13)) * exRateP
                        remA = remA + val(gArrPayment(i, 14)) * exRateP
                    ElseIf pgrp = "GROUP-B" Then
                        actB = actB + val(gArrPayment(i, 10)) * exRateP
                        accB = accB + val(gArrPayment(i, 11)) * exRateP
                        reqB = reqB + val(gArrPayment(i, 12)) * exRateP
                        paidB = paidB + val(gArrPayment(i, 13)) * exRateP
                        remB = remB + val(gArrPayment(i, 14)) * exRateP
                    End If
                End If

NEXT_PAY:
            End If
        Next i

        Set dGrp = Nothing

    End If

    Set dCatByCID = Nothing

    '========================
    ' DISPLAY
    '========================
    lblGroupA_Estimated.Caption = FastVN(estA)
    lblGroupA_Total.Caption = FastVN(totalA)
    lblGroupA_ActualVolume.Caption = FastVN(actA)
    lblGroupA_AcceptedVolume.Caption = FastVN(accA)
    lblGroupA_Requested.Caption = FastVN(reqA)
    lblGroupA_Paid.Caption = FastVN(paidA)
    lblGroupA_Remain.Caption = FastVN(remA)

    lblGroupB_Estimated.Caption = FastVN(estB)
    lblGroupB_Total.Caption = FastVN(totalB)
    lblGroupB_ActualVolume.Caption = FastVN(actB)
    lblGroupB_AcceptedVolume.Caption = FastVN(accB)
    lblGroupB_Requested.Caption = FastVN(reqB)
    lblGroupB_Paid.Caption = FastVN(paidB)
    lblGroupB_Remain.Caption = FastVN(remB)

    lbl_PaidRemaining.Caption = FastVN(paidA - paidB)

End Sub




Public Sub LoadUnpaidContracts( _
    Optional mode As PaymentFilterMode = pfall, _
    Optional dnFilter As String = "", _
    Optional categoryFilter As String = "")

    Dim arrList() As Variant
    Dim finalArr() As Variant
    Dim tmp(0 To 11) As Variant
    Dim total As Long, i As Long, r As Long, c As Long
    Dim r1 As Long, r2 As Long
    Dim currentGroup As String, lastShownGroup As String
    Dim cellC As String, sTT As String
    Dim sortOrder As Long, statusText As String

    Me.MousePointer = fmMousePointerHourGlass
    lstUnpaidContract.Clear
    currentGroup = ""
    lastShownGroup = ""

    If IsEmpty(gArrPayment) Then GoTo SAFE_EXIT
    If Not IsArray(gArrPayment) Then GoTo SAFE_EXIT

    ' ====================================
    ' Ð?C TIME FILTER T? CONTROLS
    ' ====================================
    Dim useMonthly As Boolean, useYearly As Boolean
    On Error Resume Next
    useMonthly = optPayMonthly.value
    useYearly = optPayYearly.value
    On Error GoTo 0
    Dim useTimeFilter As Boolean
    useTimeFilter = (useMonthly Or useYearly)

    ' ====================================
    ' B? L?C CATEGORY (M?I) — cboCategories
    ' Ngu?n: gArrContract c?t R (18), ánh x? qua ContractID c?t A (1)
    ' Áp d?ng d? l?c gArrPayment theo c?t C (3) = ContractID
    ' ====================================
    Dim catKwLow As String
    catKwLow = LCase(Trim(categoryFilter))
    Dim useCat As Boolean
    useCat = (catKwLow <> "")

    Dim dCatByCID As Object
    Set dCatByCID = BuildCategoryDictFromContract()

    ' ====================================
    ' PRE-PASS: xây dict t?ng theo contractID
    ' Kích ho?t khi có DN filter HO?C time filter
    ' ====================================
    Dim useDN As Boolean
    Dim dnKwLow As String
    dnKwLow = LCase(Trim(dnFilter))
    useDN = (dnKwLow <> "")

    Dim useFilter As Boolean
    useFilter = (useDN Or useTimeFilter)

    Dim dAct As Object, dAcc As Object, dReq As Object
    Dim dPd  As Object, dRm  As Object
    Dim cid As String
    If useFilter Then
        Set dAct = CreateObject("Scripting.Dictionary")
        Set dAcc = CreateObject("Scripting.Dictionary")
        Set dReq = CreateObject("Scripting.Dictionary")
        Set dPd = CreateObject("Scripting.Dictionary")
        Set dRm = CreateObject("Scripting.Dictionary")
        For i = 1 To UBound(gArrPayment, 1)
            If Trim(CStr(gArrPayment(i, 2))) <> "" Then   ' payment row

                cid = Trim(CStr(gArrPayment(i, 3)))  ' cot C = ContractID

                ' -- Category filter (M?I, tra Contract-DB theo ContractID) --
                If useCat Then
                    If Not dCatByCID.Exists(cid) Then GoTo NEXT_PRE
                    If LCase(Trim(CStr(dCatByCID(cid)))) <> catKwLow Then GoTo NEXT_PRE
                End If

                ' -- DN filter --
                If useDN Then
                    If LCase(Trim(CStr(gArrPayment(i, 16)))) <> dnKwLow Then GoTo NEXT_PRE
                End If
                ' -- Time filter --
                If useTimeFilter Then
                    Dim preDateVal As Variant
                    preDateVal = gArrPayment(i, 17)
                    Dim preSkip As Boolean: preSkip = True
                    If IsValidDMY(preDateVal) Then
                        Dim prePd As Date: prePd = ParseDMY(preDateVal)
                        If useMonthly Then
                            If Month(prePd) = mPayMonth And Year(prePd) = mPayYear Then preSkip = False
                        Else
                            If Year(prePd) = mPayYear Then preSkip = False
                        End If
                    ElseIf IsDate(preDateVal) Then
                        Dim prePd2 As Date: prePd2 = CDate(preDateVal)
                        If useMonthly Then
                            If Month(prePd2) = mPayMonth And Year(prePd2) = mPayYear Then preSkip = False
                        Else
                            If Year(prePd2) = mPayYear Then preSkip = False
                        End If
                    End If
                    If preSkip Then GoTo NEXT_PRE
                End If
                ' -- T?ng h?p theo contractID --
                If cid <> "" Then
                    Dim exRatePre As Double
                    exRatePre = val(gArrPayment(i, 7))
                    If exRatePre <= 0 Then exRatePre = 1
                    dAct(cid) = val(dAct(cid)) + val(gArrPayment(i, 10)) * exRatePre
                    dAcc(cid) = val(dAcc(cid)) + val(gArrPayment(i, 11)) * exRatePre
                    dReq(cid) = val(dReq(cid)) + val(gArrPayment(i, 12)) * exRatePre
                    dPd(cid) = val(dPd(cid)) + val(gArrPayment(i, 13)) * exRatePre
                    dRm(cid) = val(dRm(cid)) + val(gArrPayment(i, 14)) * exRatePre
                End If
            End If
NEXT_PRE:
        Next i
        If dAct.count = 0 Then
            lstPaymentDetail.Clear
            GoTo SAFE_EXIT
        End If
    End If

    ' ====================================
    ' PREALLOCATE
    ' ====================================
    ReDim arrList(0 To UBound(gArrPayment, 1) * 2, 0 To 11)
    total = 0

    ' ====================================
    ' BUILD DATA
    ' ====================================
    For i = 1 To UBound(gArrPayment, 1)

        cellC = Trim$(CStr(gArrPayment(i, 3)))   ' cot C = ContractID

        If cellC Like "GROUP-*" Then
            currentGroup = cellC
            GoTo NEXT_I
        End If

        If (Left$(cellC, 2) = "A-" Or Left$(cellC, 2) = "B-") _
        And Trim$(CStr(gArrPayment(i, 2))) = "" Then

            ' ---- L?c Category theo ContractID (M?I, tra Contract-DB) ----
            If useCat Then
                If Not dCatByCID.Exists(cellC) Then GoTo NEXT_I
                If LCase(Trim$(CStr(dCatByCID(cellC)))) <> catKwLow Then GoTo NEXT_I
            End If

            ' ---- Filter: ch? gi? contract có d? li?u trong k? l?c ----
            If useFilter Then
                If Not dAct.Exists(cellC) Then GoTo NEXT_I
            End If

            ' ---- Unpaid filter ----
            If mode = pfUnpaid Then
                If useFilter Then
                    If val(dRm(cellC)) <= 0 Then GoTo NEXT_I
                Else
                    If val(gArrPayment(i, 14)) <= 0 Then GoTo NEXT_I
                End If
            End If

            sTT = ""
            sortOrder = 999999

            If Left$(cellC, 6) = "B-CPB-" Then
                currentGroup = "GROUP-B"
                sTT = "B-CPB"
                sortOrder = 0
            ElseIf gDictContract.Exists(cellC) Then
                r = gDictContract(cellC)
                sTT = Trim$(CStr(gArrContract(r, 24)))
                statusText = UCase$(Trim$(CStr(gArrContract(r, 6))))
                If statusText = "CANCEL" Then sTT = sTT & "(x)"
                If IsNumeric(gArrContract(r, 25)) Then
                    sortOrder = CLng(gArrContract(r, 25))
                End If
            End If

            arrList(total, 0) = cellC
            arrList(total, 1) = currentGroup
            arrList(total, 2) = sTT
            arrList(total, 3) = gArrPayment(i, 4)
            arrList(total, 4) = gArrPayment(i, 6)

            Dim exRateC As Double
            exRateC = val(gArrPayment(i, 7))
            If exRateC <= 0 Then exRateC = 1

            If useFilter Then
                arrList(total, 5) = FastVN(val(gArrPayment(i, 9)) * exRateC)   ' ContractValue – g?c OK
                arrList(total, 6) = FastVN(dAct(cellC))
                arrList(total, 7) = FastVN(dAcc(cellC))
                arrList(total, 8) = FastVN(dReq(cellC))
                arrList(total, 9) = FastVN(dPd(cellC))
                arrList(total, 10) = FastVN(dRm(cellC))
            Else
                arrList(total, 5) = FastVN(val(gArrPayment(i, 9)) * exRateC)   ' ContractValue – g?c OK
                arrList(total, 6) = FastVN(val(gArrPayment(i, 10)))   ' col J – dã VND
                arrList(total, 7) = FastVN(val(gArrPayment(i, 11)))   ' col K – dã VND
                arrList(total, 8) = FastVN(val(gArrPayment(i, 12)))   ' col L – dã VND
                arrList(total, 9) = FastVN(val(gArrPayment(i, 13)))   ' col M – dã VND
                arrList(total, 10) = FastVN(val(gArrPayment(i, 14)))  ' col N – dã VND
            End If
            
            arrList(total, 11) = sortOrder
            total = total + 1

        End If
NEXT_I:
    Next i

    Set dCatByCID = Nothing

    If total <= 0 Then
        lstPaymentDetail.Clear
        GoTo SAFE_EXIT
    End If

    ' ====================================
    ' SORT (gi? nguyên logic cu)
    ' ====================================
    For r1 = 0 To total - 2
        For r2 = r1 + 1 To total - 1
            Dim grp1 As String, grp2 As String
            Dim ord1 As Long, ord2 As Long
            grp1 = Trim$(CStr(arrList(r1, 1)))
            grp2 = Trim$(CStr(arrList(r2, 1)))
            ord1 = CLng(arrList(r1, 11))
            ord2 = CLng(arrList(r2, 11))
            Dim swapIt As Boolean: swapIt = False
            If grp1 <> grp2 Then
                If grp1 = "GROUP-B" And grp2 = "GROUP-A" Then swapIt = True
            Else
                If ord2 < ord1 Then swapIt = True
            End If
            If swapIt Then
                For c = 0 To 11
                    tmp(c) = arrList(r1, c)
                    arrList(r1, c) = arrList(r2, c)
                    arrList(r2, c) = tmp(c)
                Next c
            End If
        Next r2
    Next r1

    ' ====================================
    ' BUILD FINAL WITH SEPARATORS
    ' ====================================
    ReDim finalArr(0 To total * 2, 0 To 11)
    total = 0
    lastShownGroup = ""
    Dim groupKey As String

    For i = 0 To UBound(arrList, 1)
        If Trim$(CStr(arrList(i, 0))) <> "" Then
            groupKey = IIf(Left$(Trim$(CStr(arrList(i, 2))), 2) = "A-", "A", "B")
            If lastShownGroup <> "" And groupKey <> lastShownGroup Then
                finalArr(total, 0) = "0"
                finalArr(total, 2) = "++++++++++++++++++++++"
                finalArr(total, 3) = "+++++++++++++++++++++++++"
                finalArr(total, 4) = "++++++++++++++++"
                finalArr(total, 5) = "++++++++++++++++"
                finalArr(total, 6) = "++++++++++++++++"
                finalArr(total, 7) = "++++++++++++++++"
                finalArr(total, 8) = "++++++++++++++++"
                finalArr(total, 9) = "++++++++++++++++"
                finalArr(total, 10) = "++++++++++++++++"
                total = total + 1
            End If
            lastShownGroup = groupKey
            For c = 0 To 11
                finalArr(total, c) = arrList(i, c)
            Next c
            total = total + 1
        End If
    Next i

    Dim outArr() As Variant
    ReDim outArr(0 To total - 1, 0 To 11)
    Dim rr As Long, cc As Long
    For rr = 0 To total - 1
        For cc = 0 To 11
            outArr(rr, cc) = finalArr(rr, cc)
        Next cc
    Next rr

    lstUnpaidContract.List = outArr

    ' Auto select first contract
    Dim firstRow As Long
    For firstRow = 0 To lstUnpaidContract.ListCount - 1
        If Trim$(lstUnpaidContract.List(firstRow, 0)) <> "" _
        And Trim$(lstUnpaidContract.List(firstRow, 0)) <> "0" Then
            suppressContractClick = True
            lstUnpaidContract.ListIndex = firstRow
            LoadContractDetail CStr(lstUnpaidContract.List(firstRow, 0))
            btnPaymentEntry.Enabled = True
            suppressContractClick = False
            Exit For
        End If
    Next firstRow

SAFE_EXIT:
    Me.MousePointer = fmMousePointerDefault
End Sub



Private Sub btnAskAI_Click()
    frmAskAI.Show vbModeless
End Sub

Private Sub btnHome_Click()

    Application.ScreenUpdating = False
    Application.Cursor = xlDefault

    Me.Hide

    On Error Resume Next
    Unload frmHomePage
    On Error GoTo 0

    frmHomePage.Show vbModeless
    frmHomePage.RefreshList

    Application.Cursor = xlDefault
    Application.ScreenUpdating = True

End Sub

'================================================================
' REMINDER BUTTON
'================================================================
Private Sub btnReminder_Click()

    On Error Resume Next
    If Not gContractLoaded Then BuildGlobalContractCache
    If Not gPaymentLoaded Then BuildPaymentCache
    On Error GoTo 0

    ' N?u frmReminder dang m? thì ch? refresh + Show l?i
    If Not mReminderForm Is Nothing Then
        mReminderForm.LoadPendingList
        mReminderForm.Show vbModeless
        Exit Sub
    End If

    ' T?o m?i hoàn toàn (ch?y l?i UserForm_Initialize c?a frmReminder)
    Set mReminderForm = New frmReminder
    mReminderForm.SetContractSearchRef Me
    mReminderForm.Show vbModeless

End Sub

'================================================================
' Ðu?c g?i b?i frmReminder khi form dó b? Unload
' ? Xoá zombie reference, l?n b?m sau t?o m?i du?c
'================================================================
Public Sub ClearReminderRef()
    Set mReminderForm = Nothing
End Sub

'================================================================
' Reference ngu?c d?n frmReminder
' (du?c set khi frmPaymentSummary m? ra t? frmReminder)
'================================================================
Public Sub SetReminderRef(ByVal frm As Object)
    Set mReminderRef = frm
End Sub

Private Sub txtDNSearch_Enter()
    If txtDNSearch.Text = "DN" Then
        suppressDNChange = True
        txtDNSearch.Text = ""
        txtDNSearch.ForeColor = RGB(0, 0, 0)
        suppressDNChange = False
    End If
End Sub

Private Sub txtDNSearch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(txtDNSearch.Text) = "" Then
        suppressDNChange = True
        txtDNSearch.Text = "DN"
        txtDNSearch.ForeColor = RGB(150, 150, 150)
        suppressDNChange = False
    End If
End Sub


Private Sub UserForm_Activate()
    On Error Resume Next
    lstUnpaidContract.SetFocus
    On Error GoTo 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    ' Thông báo frmReminder gi?i phóng reference ngu?c
    If Not mReminderRef Is Nothing Then
        On Error Resume Next
        mReminderRef.ClearContractSearchRef
        On Error GoTo 0
        Set mReminderRef = Nothing
    End If

    If CloseMode = vbFormControlMenu Then
        'Cancel = True
        Call ReleaseMemory
        Unload Me
    End If
    'ApplyTheme Me
    UnregisterForm Me
End Sub


Private Sub btnPaymentEntry_Click()

    Dim contractID As String

    If lstUnpaidContract.ListIndex < 0 Then

        ShowMsgCenter _
            "Please select a contract first.", _
            vbExclamation

        Exit Sub

    End If

    contractID = _
        Trim$(lstUnpaidContract.List( _
        lstUnpaidContract.ListIndex, 0))

    If contractID = "" _
    Or contractID = "0" Then Exit Sub

    On Error Resume Next

    Me.MousePointer = fmMousePointerHourGlass

    '========================
    ' FULL RELEASE OLD FORM
    '========================
    Unload frmPaymentEntry
    Set frmPaymentEntry = Nothing

    DoEvents

    '========================
    ' OPEN CLEAN FORM
    '========================
    frmPaymentEntry.Show vbModeless

    frmPaymentEntry.InitFromContractFast _
        contractID
    'isRunning = False
    Me.Hide

SAFE_EXIT:

    Me.MousePointer = _
        fmMousePointerDefault

End Sub


Private Sub UpdatePayControlState()
    Dim useMonthly As Boolean, useYearly As Boolean
    On Error Resume Next
    useMonthly = optPayMonthly.value
    useYearly = optPayYearly.value
    btnPayThisMonth.Enabled = useMonthly
    btnPayMonthPrev.Enabled = useMonthly
    txtPayMonth.Enabled = useMonthly
    btnPayMonthNext.Enabled = useMonthly
    btnPayThisYear.Enabled = (useMonthly Or useYearly)
    btnPayYearPrev.Enabled = (useMonthly Or useYearly)
    txtPayYear.Enabled = (useMonthly Or useYearly)
    btnPayYearNext.Enabled = (useMonthly Or useYearly)
    On Error GoTo 0
End Sub

Private Sub RefreshPaymentDetail()
    If lstUnpaidContract.ListIndex < 0 Then Exit Sub
    Dim cid As String
    cid = Trim$(CStr(lstUnpaidContract.List(lstUnpaidContract.ListIndex, 0)))
    If cid = "" Or cid = "0" Then Exit Sub
    LoadContractDetail cid
End Sub

' Reload c? lstUnpaidContract theo b? l?c hi?n t?i,
' c? g?ng gi? nguyên contract dang ch?n.
Private Sub RefreshContractListAndDetail()
    Dim savedID As String: savedID = ""
    If lstUnpaidContract.ListIndex >= 0 Then
        savedID = Trim$(CStr(lstUnpaidContract.List(lstUnpaidContract.ListIndex, 0)))
        If savedID = "0" Then savedID = ""
    End If

    Dim dnKw As String
    On Error Resume Next
    dnKw = Trim(txtDNSearch.Text)
    On Error GoTo 0
    If dnKw = "DN" Then dnKw = ""

    Dim mode As PaymentFilterMode
    On Error Resume Next
    If chkUnpaid.value Then mode = pfUnpaid Else mode = pfall
    On Error GoTo 0

    LoadUnpaidContracts mode, dnKw, GetCurrentCategoryFilter()

    LoadGroupSummary

    If savedID <> "" Then
        Dim i As Long
        For i = 0 To lstUnpaidContract.ListCount - 1
            If Trim$(CStr(lstUnpaidContract.List(i, 0))) = savedID Then
                suppressContractClick = True
                lstUnpaidContract.ListIndex = i
                suppressContractClick = False
                LoadContractDetail savedID
                Exit For
            End If
        Next i
    End If
End Sub



' ---- Radio buttons ----
Private Sub optPayAllTime_Click():  UpdatePayControlState: RefreshContractListAndDetail: End Sub
Private Sub optPayMonthly_Click():  UpdatePayControlState: RefreshContractListAndDetail: End Sub
Private Sub optPayYearly_Click():   UpdatePayControlState: RefreshContractListAndDetail: End Sub

' ---- This Month / This Year ----
Private Sub btnPayThisMonth_Click()
    mPayMonth = Month(Now): mPayYear = Year(Now)
    On Error Resume Next
    txtPayMonth.Text = CStr(mPayMonth): txtPayYear.Text = CStr(mPayYear)
    On Error GoTo 0
    RefreshContractListAndDetail
End Sub
Private Sub btnPayThisYear_Click()
    mPayYear = Year(Now)
    On Error Resume Next: txtPayYear.Text = CStr(mPayYear): On Error GoTo 0
    RefreshContractListAndDetail
End Sub

' ---- Month navigation ----
Private Sub btnPayMonthPrev_Click()
    mPayMonth = mPayMonth - 1
    If mPayMonth < 1 Then mPayMonth = 12: mPayYear = mPayYear - 1
    On Error Resume Next: txtPayMonth.Text = CStr(mPayMonth): txtPayYear.Text = CStr(mPayYear): On Error GoTo 0
    RefreshContractListAndDetail
End Sub
Private Sub btnPayMonthNext_Click()
    mPayMonth = mPayMonth + 1
    If mPayMonth > 12 Then mPayMonth = 1: mPayYear = mPayYear + 1
    On Error Resume Next: txtPayMonth.Text = CStr(mPayMonth): txtPayYear.Text = CStr(mPayYear): On Error GoTo 0
    RefreshContractListAndDetail
End Sub

' ---- Year navigation ----
Private Sub btnPayYearPrev_Click()
    mPayYear = mPayYear - 1
    On Error Resume Next: txtPayYear.Text = CStr(mPayYear): On Error GoTo 0
    RefreshContractListAndDetail
End Sub
Private Sub btnPayYearNext_Click()
    mPayYear = mPayYear + 1
    On Error Resume Next: txtPayYear.Text = CStr(mPayYear): On Error GoTo 0
    RefreshContractListAndDetail
End Sub

' ---- Manual input ----
Private Sub txtPayMonth_Change()
    On Error Resume Next
    Dim m As Long: m = CLng(txtPayMonth.Text)
    If m >= 1 And m <= 12 Then mPayMonth = m: RefreshContractListAndDetail
    On Error GoTo 0
End Sub
Private Sub txtPayYear_Change()
    On Error Resume Next
    Dim Y As Long: Y = CLng(txtPayYear.Text)
    If Y >= 2000 And Y <= 2100 Then mPayYear = Y: RefreshContractListAndDetail
    On Error GoTo 0
End Sub

' ---- DN Search ----
Private Sub txtDNSearch_Change()
    If suppressDNChange Then Exit Sub
    Dim dnKw As String
    dnKw = Trim(txtDNSearch.Text)
    If dnKw = "DN" Then dnKw = ""

    Dim mode As PaymentFilterMode
    On Error Resume Next
    If chkUnpaid.value Then mode = pfUnpaid Else mode = pfall
    On Error GoTo 0

    LoadUnpaidContracts mode, dnKw, GetCurrentCategoryFilter()
End Sub


' ---- DN Double Right-Click: clear search ----
Private Sub txtDNSearch_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then  ' Right click
        If (Timer - mDNRightClickTime) < 0.5 And (Timer - mDNRightClickTime) >= 0 Then
            suppressDNChange = True
            txtDNSearch.Text = "DN"
            txtDNSearch.ForeColor = RGB(150, 150, 150)
            suppressDNChange = False
            Dim mode As PaymentFilterMode
            On Error Resume Next
            If chkUnpaid.value Then mode = pfUnpaid Else mode = pfall
            On Error GoTo 0
            LoadUnpaidContracts mode, "", GetCurrentCategoryFilter()
        End If
        mDNRightClickTime = Timer
    End If
End Sub



Private Sub chkUnpaid_Click()
    suppressDNChange = True
    txtDNSearch.Text = "DN"
    txtDNSearch.ForeColor = RGB(150, 150, 150)
    suppressDNChange = False
    If chkUnpaid.value Then
        LoadUnpaidContracts pfUnpaid, "", GetCurrentCategoryFilter()
    Else
        LoadUnpaidContracts pfall, "", GetCurrentCategoryFilter()
    End If
End Sub



Public Function NextInternalCostID() As String

    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim maxID As Long
    Dim tmp As String

    Set ws = EnsurePaymentLedger

    lastRow = ws.Cells(ws.rows.count, "C").End(xlUp).Row

    For r = 2 To lastRow

        tmp = ws.Cells(r, "C").value

        If Left(tmp, 6) = "B-CPB-" Then

            tmp = Replace(tmp, "B-CPB-", "")
            If IsNumeric(tmp) Then
                maxID = Application.Max(maxID, CLng(tmp))
            End If

        End If

    Next r

    NextInternalCostID = "B-CPB-" & Format(maxID + 1, "000")

End Function

Private Sub btnAddInternalCost_Click()

    If cmdAddInternalCost.Tag <> "" Then
        EditOrDeleteCPB cmdAddInternalCost.Tag
    Else
        AddNewCPB
    End If

End Sub

Private Sub DeleteCPB(contractID As String)

    Dim ws As Worksheet
    Set ws = EnsurePaymentLedger

    Dim r As Long, startRow As Long, endRow As Long
    Dim lastRow As Long
    Dim hasPayment As Boolean

    lastRow = ws.Cells(ws.rows.count, "C").End(xlUp).Row

'========================
' FIND FAMILY + CHECK PAYMENT
'========================
    For r = 2 To lastRow

        If ws.Cells(r, "C").value = contractID Then

            If startRow = 0 Then startRow = r

            ' payment row
            If Trim(ws.Cells(r, "B").value) <> "" Then
                hasPayment = True
            End If

            endRow = r

        ElseIf startRow > 0 Then
            Exit For
        End If

    Next r

'========================
' BLOCK DELETE IF HAS PAYMENT
'========================
    If hasPayment Then
        ShowMsgCenter "Cannot delete. This CPB already has payments.", vbCritical
        Exit Sub
    End If

'========================
' DELETE
'========================
    If MsgBox("Delete this CPB?", vbYesNo + vbCritical) <> vbYes Then Exit Sub

    ws.rows(startRow & ":" & endRow).Delete

    gPaymentLoaded = False
    BuildPaymentCache
    
    LoadUnpaidContracts pfall, "", GetCurrentCategoryFilter()
    LoadGroupSummary

    
End Sub

Private Sub EditCPBName(contractID As String)

    Dim ws As Worksheet
    Set ws = EnsurePaymentLedger

    Dim r As Long
    Dim newName As String

    newName = InputBox("Edit internal cost name:")
    If newName = "" Then Exit Sub

    For r = 2 To ws.Cells(ws.rows.count, "C").End(xlUp).Row

        If ws.Cells(r, "C").value = contractID _
           And Trim(ws.Cells(r, "B").value) = "" Then

            ws.Cells(r, "F").value = newName
            Exit For

        End If

    Next r

    gPaymentLoaded = False
    BuildPaymentCache
    LoadUnpaidContracts pfall, "", GetCurrentCategoryFilter()

        
End Sub

Private Sub EditOrDeleteCPB(contractID As String)

    Dim ans As VbMsgBoxResult

    ans = MsgBox( _
        "Yes = Edit name" & vbCrLf & _
        "No  = Delete CPB (no payments only)", _
        vbYesNoCancel + vbQuestion, _
        "CPB Action")

    If ans = vbCancel Then Exit Sub

    If ans = vbYes Then
        EditCPBName contractID
    Else
        DeleteCPB contractID
    End If
    cmdAddInternalCost.Caption = "Add Internal Cost"
    cmdAddInternalCost.Tag = ""
End Sub

Private Sub AddNewCPB()

    Dim ws As Worksheet
    Set ws = EnsurePaymentLedger

    Dim costName As String
    Dim insertRow As Long
    Dim headerRow As Long
    Dim r As Long
    Dim lastRow As Long
    Dim idVal As String
    Dim newID As String

'========================
' INPUT NAME
'========================
    costName = InputBox("Internal Cost Name:", "Internal Cost")
    If costName = "" Then Exit Sub

'========================
' AUTO GENERATE ID
'========================
    newID = NextInternalCostID()

'========================
' FIND GROUP-B HEADER
'========================
    lastRow = ws.Cells(ws.rows.count, "C").End(xlUp).Row

    For r = 2 To lastRow
        If ws.Cells(r, "C").value = "GROUP-B" _
           And Trim(ws.Cells(r, "B").value) = "" Then
            headerRow = r
            insertRow = r + 1
            Exit For
        End If
    Next r

    If insertRow = 0 Then Exit Sub

'========================
' MOVE AFTER ALL CPB
'========================
    Do While insertRow <= lastRow

        idVal = Trim(ws.Cells(insertRow, "C").value)

        If Left(idVal, 6) = "B-CPB-" Then
            insertRow = insertRow + 1
        Else
            Exit Do
        End If

    Loop

'========================
' INSERT ROW ONLY
'========================
    ws.rows(insertRow).Insert

'========================
' WRITE VALUES
'========================
    With ws

        .Cells(insertRow, "A").value = "GROUP-B"
        .Cells(insertRow, "C").value = newID
        .Cells(insertRow, "D").value = newID
        .Cells(insertRow, "F").value = costName
        .Cells(insertRow, "G").value = 1
        .Cells(insertRow, "H").value = 0
        .Cells(insertRow, "I").value = 0

        '====================
        ' SUMMARY FORMULAS (HEADER TYPE)
        '====================
        .Cells(insertRow, "W").Formula = "=SUMIFS($W:$W,$A:$A,C" & insertRow & ")"
        .Cells(insertRow, "X").Formula = "=SUMIFS($X:$X,$A:$A,C" & insertRow & ")"
        .Cells(insertRow, "Y").Formula = "=SUMIFS($Y:$Y,$A:$A,C" & insertRow & ")"
        .Cells(insertRow, "Z").Formula = "=SUMIFS($Z:$Z,$A:$A,C" & insertRow & ")"
        
        .Cells(insertRow, "J").Formula = "=W" & insertRow
        .Cells(insertRow, "K").Formula = "=X" & insertRow
        .Cells(insertRow, "L").Formula = "=Y" & insertRow
        .Cells(insertRow, "M").Formula = "=Z" & insertRow
        .Cells(insertRow, "N").Formula = "=L" & insertRow & "-M" & insertRow


    End With

'========================
' FORMAT LIKE HEADER (NO COPY)
'========================
    Dim lastCol As String
    lastCol = "Z"
    
    With ws.Range("A" & insertRow & ":" & lastCol & insertRow)
    
        .Interior.Color = RGB(250, 220, 250) ' cam nh?t
        .Font.Bold = False
        .Borders.LineStyle = xlContinuous
    
    End With


'========================
' REFRESH
'========================
    
    gPaymentLoaded = False
    BuildPaymentCache
    LoadGroupSummary
    LoadUnpaidContracts pfall, "", GetCurrentCategoryFilter()

    ShowMsgCenter "Created: " & newID, vbInformation

End Sub



'================================================================
' NUT CONTRACT: chuyen sang frmContractSearch, giu contract ID
'================================================================
Private Sub btnContract_Click()

    ' Tat nut ngay lap tuc — block moi click tiep theo du nhanh den dau
    btnContract.Enabled = False

    Dim cid As String
    cid = ""

    On Error Resume Next
    If lstUnpaidContract.ListIndex >= 0 Then
        cid = Trim$(CStr(lstUnpaidContract.List(lstUnpaidContract.ListIndex, 0)))
        If cid = "0" Then cid = ""
    End If
    On Error GoTo 0

    Unload frmContractSearch
    Set frmContractSearch = Nothing

    frmContractSearch.Show vbModeless
    frmContractSearch.InitFast
    DoEvents

    If cid <> "" Then
        On Error Resume Next
        frmContractSearch.SelectContract cid
        On Error GoTo 0
    End If

    ' Bat lai nut TRUOC khi Hide — de lan sau form duoc show lai thi nut van hoat dong
    btnContract.Enabled = True

    On Error Resume Next
    Me.Hide
    On Error GoTo 0

End Sub



Private Sub lstUnpaidContract_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim idx As Long
    idx = lstUnpaidContract.ListIndex
    If idx < 0 Then Exit Sub

    ' 0 = Group, 1 = Contract
    If lstUnpaidContract.List(idx, 0) = 0 Then
        ShowMsgCenter "Please select a contract item, not a group header."
        Cancel = True
        Exit Sub
    End If

    btnPaymentEntry_Click
End Sub

Private Sub lstUnpaidContract_Click()

    If suppressContractClick Then Exit Sub

    On Error GoTo SAFE_EXIT

    Dim contractID As String

    If lstUnpaidContract.ListIndex < 0 Then Exit Sub

    contractID = _
        Trim$(CStr( _
        lstUnpaidContract.List( _
        lstUnpaidContract.ListIndex, 0)))

    If contractID = "" _
    Or contractID = "0" Then Exit Sub

    suppressContractClick = True

    '================================
    ' CPB MODE
    '================================
    If Left$(contractID, 6) = "B-CPB-" Then

        cmdAddInternalCost.Caption = _
            "Edit / Delete CPB"

        cmdAddInternalCost.Tag = _
            contractID

    Else

        cmdAddInternalCost.Caption = _
            "Add Internal Cost"

        cmdAddInternalCost.Tag = ""

    End If

    '================================
    ' LOAD DETAIL
    '================================
    LoadContractDetail contractID

    btnPaymentEntry.Enabled = True

SAFE_EXIT:

    suppressContractClick = False

End Sub



Private Sub UserForm_Initialize()
    
    With lstPaymentDetail

        .ColumnCount = 8

        .ColumnWidths = _
        "25 pt;60 pt;185 pt;90 pt;90 pt;90 pt;90 pt;90 pt"

    End With

    With lstUnpaidContract

        .ColumnCount = 12

        .ColumnWidths = _
        "0 pt;0 pt;35 pt;145 pt;0 pt;90 pt;90 pt;90 pt;90 pt;90 pt;90 pt;0 pt"

    End With

    On Error Resume Next
    cboCategories.Style = fmStyleDropDownList
    cboCategories.Clear
    cboCategories.AddItem "(All)"
    cboCategories.ListIndex = 0
    On Error GoTo 0
    
    ' Kh?i t?o b? l?c th?i gian
    mPayMonth = Month(Now)
    mPayYear = Year(Now)
    On Error Resume Next
    optPayAllTime.value = True
    txtPayMonth.Text = CStr(mPayMonth)
    txtPayYear.Text = CStr(mPayYear)
    On Error GoTo 0
    UpdatePayControlState
    
    txtDNSearch.Text = "DN"
    txtDNSearch.ForeColor = RGB(150, 150, 150)

    cmdAddInternalCost.Caption = "Add Internal Cost"

    cmdAddInternalCost.Tag = ""

    'btnPaymentEntry.Enabled = False
    ApplyTheme Me
    'UnregisterForm Me
End Sub

Public Sub InitFast()

    If isLoading Then Exit Sub

    On Error GoTo SAFE_EXIT

    isLoading = True

    'Me.Visible = False

    '========================
    ' BUILD CACHE
    '========================
    If Not gPaymentLoaded Then

        BuildPaymentCache

    End If

    DoEvents

    '========================
    ' ALWAYS RELOAD UI
    '========================
    LoadData

    'Me.Visible = True

SAFE_EXIT:

    isLoading = False

    'Me.Visible = True

End Sub

Private Sub LoadContractDetail(contractID As String)

    Dim r As Long
    Dim rows As Collection
    Dim idx As Variant

    Dim arrList() As Variant
    Dim finalArr() As Variant
    Dim total As Long
    Dim rr As Long, cc As Long

    Me.MousePointer = fmMousePointerHourGlass

    ' Reset
    lblDetailContractNo.Caption = ""
    lblDetailSigningDate.Caption = ""
    lblDetailCompany2.Caption = ""
    'lblDetailCompany3.Caption = ""
    lblDetailScope.Caption = ""
    lblDetailValue.Caption = ""
    lstPaymentDetail.Clear

    ' Contract info
    If gDictContract.Exists(contractID) Then
        r = gDictContract(contractID)
        lblDetailContractNo.Caption = BuildDisplayContractNo(contractID)
        If IsDate(gArrContract(r, 4)) Then
            lblDetailSigningDate.Caption = Format(gArrContract(r, 4), "dd/mm/yyyy")
        End If
        lblDetailCompany2.Caption = gArrContract(r, 8)
        'lblDetailCompany3.Caption = gArrContract(r, 9)
        lblDetailScope.Caption = gArrContract(r, 11)
        lblDetailValue.Caption = FastVN(gArrContract(r, 15))
    End If

    If Not gDictPaymentRows.Exists(contractID) Then
        Me.MousePointer = fmMousePointerDefault
        Exit Sub
    End If

    ' ---- L?y b? l?c ----
    Dim useMonthly As Boolean, useYearly As Boolean
    Dim dnKw As String
    On Error Resume Next
    useMonthly = optPayMonthly.value
    useYearly = optPayYearly.value
    dnKw = LCase(Trim(txtDNSearch.Text))
    If dnKw = "dn" Then dnKw = ""   ' b? qua placeholder
    On Error GoTo 0

    Set rows = gDictPaymentRows(contractID)
    ReDim arrList(0 To rows.count + 5, 0 To 7)
    total = 0

    Dim sumAct As Double, sumAcc As Double
    Dim sumReq As Double, sumPd As Double, sumRm As Double
    Dim hasHeaderRow As Boolean
    hasHeaderRow = False

    For Each idx In rows

        If Trim(CStr(gArrPayment(idx, 2))) = "" Then
            ' ---- Header row (DN / T?ng c?ng) ----
            ' Gi? l?i, s? c?p nh?t giá tr? sau vòng l?p
            hasHeaderRow = True
            arrList(0, 0) = "DN"
            arrList(0, 2) = "Total"
            ' Giá tr? s? du?c tính l?i bên du?i
            ' Separator
            arrList(1, 2) = "--------------------------"
            If total < 2 Then total = 2

        Else
            ' ---- Payment row ----

            ' 1. L?c th?i gian
            If useMonthly Or useYearly Then
                Dim payDate As Variant
                payDate = gArrPayment(idx, 17)
                Dim SkipRow As Boolean: SkipRow = False
                If IsValidDMY(payDate) Then
                    Dim pd As Date
                    pd = ParseDMY(payDate)
                    If useMonthly Then
                        If Month(pd) <> mPayMonth Or Year(pd) <> mPayYear Then SkipRow = True
                    ElseIf useYearly Then
                        If Year(pd) <> mPayYear Then SkipRow = True
                    End If
                ElseIf IsDate(payDate) Then
                    Dim pd2 As Date
                    pd2 = CDate(payDate)
                    If useMonthly Then
                        If Month(pd2) <> mPayMonth Or Year(pd2) <> mPayYear Then SkipRow = True
                    ElseIf useYearly Then
                        If Year(pd2) <> mPayYear Then SkipRow = True
                    End If
                Else
                    SkipRow = True
                End If
                If SkipRow Then GoTo NEXT_IDX
            End If

            ' 2. L?c DN search (c?t P = index 16 = payment No)
            If dnKw <> "" Then
                Dim descVal As String
                descVal = LCase(Trim(CStr(gArrPayment(idx, 16)))) ' c?t P
                If InStr(descVal, dnKw) = 0 Then GoTo NEXT_IDX
            End If

            ' 3. Ghi vào m?ng
            If total > UBound(arrList, 1) - 1 Then
                ReDim Preserve arrList(0 To UBound(arrList, 1) + 200, 0 To 7)
            End If

            arrList(total, 0) = gArrPayment(idx, 16)  ' No

            If IsDate(gArrPayment(idx, 17)) Then
                arrList(total, 1) = Format(gArrPayment(idx, 17), "dd/mm/yyyy")
            ElseIf IsValidDMY(gArrPayment(idx, 17)) Then
                arrList(total, 1) = Format(ParseDMY(gArrPayment(idx, 17)), "dd/mm/yyyy")
            End If

            arrList(total, 2) = gArrPayment(idx, 15)           ' Description

            ' Ty gia ngoai te: cot G (index 7). Chuyen ve VND truoc khi hien thi
            Dim exRatePay As Double
            exRatePay = val(gArrPayment(idx, 7))
            If exRatePay <= 0 Then exRatePay = 1

            arrList(total, 3) = FastVN(val(gArrPayment(idx, 10)) * exRatePay)   ' Actual (VND)
            arrList(total, 4) = FastVN(val(gArrPayment(idx, 11)) * exRatePay)   ' Accepted (VND)
            arrList(total, 5) = FastVN(val(gArrPayment(idx, 12)) * exRatePay)   ' Requested (VND)
            arrList(total, 6) = FastVN(val(gArrPayment(idx, 13)) * exRatePay)   ' Paid (VND)
            arrList(total, 7) = FastVN(val(gArrPayment(idx, 14)) * exRatePay)   ' Remaining (VND)
            ' C?ng vào t?ng DN
            On Error Resume Next
            sumAct = sumAct + val(gArrPayment(idx, 10)) * exRatePay
            sumAcc = sumAcc + val(gArrPayment(idx, 11)) * exRatePay
            sumReq = sumReq + val(gArrPayment(idx, 12)) * exRatePay
            sumPd = sumPd + val(gArrPayment(idx, 13)) * exRatePay
            sumRm = sumRm + val(gArrPayment(idx, 14)) * exRatePay

            total = total + 1

        End If
NEXT_IDX:
    Next idx

    ' C?p nh?t DN row (t?ng theo b? l?c)
    If hasHeaderRow Then
        If useMonthly Or useYearly Then
            ' Dùng t?ng dã tính l?i t? các dòng l?c
            arrList(0, 3) = FastVN(sumAct)
            arrList(0, 4) = FastVN(sumAcc)
            arrList(0, 5) = FastVN(sumReq)
            arrList(0, 6) = FastVN(sumPd)
            arrList(0, 7) = FastVN(sumRm)
        Else
            ' All Time: l?y t?ng t? header g?c trong cache
            Dim hIdx As Variant
            For Each hIdx In gDictPaymentRows(contractID)
                If Trim(CStr(gArrPayment(hIdx, 2))) = "" Then
                    ' Các c?t J–N (index 10–14) trên header row dã là VND
                    ' (công th?c =W, =X, =Y, =Z, =L-M t?ng h?p t? các dòng con)
                    ' ? KHÔNG nhân thêm t? giá
                    arrList(0, 3) = FastVN(val(gArrPayment(hIdx, 10)))  ' col J – dã VND
                    arrList(0, 4) = FastVN(val(gArrPayment(hIdx, 11)))  ' col K – dã VND
                    arrList(0, 5) = FastVN(val(gArrPayment(hIdx, 12)))  ' col L – dã VND
                    arrList(0, 6) = FastVN(val(gArrPayment(hIdx, 13)))  ' col M – dã VND
                    arrList(0, 7) = FastVN(val(gArrPayment(hIdx, 14)))  ' col N – dã VND
                End If
            Next hIdx
        End If
    End If

    If total <= 0 Then
        Me.MousePointer = fmMousePointerDefault
        Exit Sub
    End If

    ReDim finalArr(0 To total - 1, 0 To 7)
    For rr = 0 To total - 1
        For cc = 0 To 7
            finalArr(rr, cc) = arrList(rr, cc)
        Next cc
    Next rr

    lstPaymentDetail.List = finalArr

    Me.MousePointer = fmMousePointerDefault

End Sub


'========================================
' RELEASE MEMORY
'========================================
Private Sub ReleaseMemory()

    On Error Resume Next

    Me.MousePointer = fmMousePointerHourGlass

    '========================
    ' CLEAR LISTBOX
    '========================
    lstUnpaidContract.Clear
    lstPaymentDetail.Clear

    '========================
    ' RESET BUTTON
    '========================
    'btnPaymentEntry.Enabled = False

    cmdAddInternalCost.Caption = _
        "Add Internal Cost"

    cmdAddInternalCost.Tag = ""

    '========================
    ' RESET CHECKBOX
    '========================
    chkUnpaid.value = False

    '========================
    ' RESET LABELS
    '========================
    lblDetailContractNo.Caption = ""
    lblDetailSigningDate.Caption = ""

    lblDetailCompany2.Caption = ""
    'lblDetailCompany3.Caption = ""

    lblDetailScope.Caption = ""
    lblDetailValue.Caption = ""

    '========================
    ' RESET SUMMARY
    '========================
    lblGroupA_Estimated.Caption = ""
    lblGroupA_Total.Caption = ""
    lblGroupA_ActualVolume.Caption = ""
    lblGroupA_AcceptedVolume.Caption = ""
    lblGroupA_Requested.Caption = ""
    lblGroupA_Paid.Caption = ""
    lblGroupA_Remain.Caption = ""

    lblGroupB_Estimated.Caption = ""
    lblGroupB_Total.Caption = ""
    lblGroupB_ActualVolume.Caption = ""
    lblGroupB_AcceptedVolume.Caption = ""
    lblGroupB_Requested.Caption = ""
    lblGroupB_Paid.Caption = ""
    lblGroupB_Remain.Caption = ""

    lbl_PaidRemaining.Caption = ""

    '========================
    ' RESET FLAGS
    '========================
    suppressContractChange = False
    suppressContractClick = False
    isLoading = False

    DoEvents

    Me.MousePointer = fmMousePointerDefault

End Sub

Private Sub UserForm_Terminate()

    On Error Resume Next

    lstUnpaidContract.Clear
    lstPaymentDetail.Clear

End Sub

'================================================================
' PUBLIC: Jump den contract cu the trong lstUnpaidContract
' Goi tu form cha (LinkKL) khi mo frmPaymentSummary
'================================================================
Public Sub SelectContract(ByVal contractID As String)

    ' Guard: contractID rong
    If Trim(contractID) = "" Then Exit Sub

    ' Guard: list chua co du lieu
    If lstUnpaidContract.ListCount = 0 Then Exit Sub

    Dim i As Long

    For i = 0 To lstUnpaidContract.ListCount - 1

        Dim cid As String
        cid = Trim$(CStr(lstUnpaidContract.List(i, 0)))

        If cid = Trim(contractID) Then

            suppressContractClick = True

            lstUnpaidContract.ListIndex = i

            suppressContractClick = False

            ' Goi giong lstUnpaidContract_Click
            If Left$(contractID, 8) = "B-CPB-" Then
                cmdAddInternalCost.Caption = "Edit / Delete CPB"
                cmdAddInternalCost.Tag = contractID
            Else
                cmdAddInternalCost.Caption = "Add Internal Cost"
                cmdAddInternalCost.Tag = ""
            End If

            LoadContractDetail contractID

            btnPaymentEntry.Enabled = True

            On Error Resume Next
            lstUnpaidContract.SetFocus
            On Error GoTo 0

            Exit For

        End If

    Next i

End Sub

Private Function GetCurrentCategoryFilter() As String
    Dim v As String
    On Error Resume Next
    v = Trim$(CStr(cboCategories.value))
    On Error GoTo 0
    If v = "" Or v = "(All)" Then
        GetCurrentCategoryFilter = ""
    Else
        GetCurrentCategoryFilter = v
    End If
End Function


'### 2.3 Thêm hàm n?p danh sách H?ng m?c vào combobox

Private Sub PopulateCategoryCombo()

    Dim i As Long
    Dim catVal As String
    Dim dSeen As Object
    Dim curSel As String

    Set dSeen = CreateObject("Scripting.Dictionary")

    ' Gi? l?i l?a ch?n hi?n t?i (n?u có) tru?c khi build l?i list
    On Error Resume Next
    curSel = cboCategories.value
    On Error GoTo 0

    cboCategories.Clear
    cboCategories.AddItem "(All)"

    If Not IsEmpty(gArrContract) Then
        If IsArray(gArrContract) Then

            For i = 1 To UBound(gArrContract, 1)

                catVal = Trim$(CStr(gArrContract(i, 18)))  ' cot R = Category

                If catVal <> "" Then
                    If Not dSeen.Exists(catVal) Then
                        dSeen(catVal) = 1
                        cboCategories.AddItem catVal
                    End If
                End If

            Next i

        End If
    End If

    ' Khôi ph?c l?a ch?n cu n?u còn t?n t?i, ngu?c l?i v? "(All)"
    Dim k As Long
    Dim found As Boolean: found = False
    If curSel <> "" Then
        For k = 0 To cboCategories.ListCount - 1
            If cboCategories.List(k) = curSel Then
                found = True
                Exit For
            End If
        Next k
    End If

    If found Then
        cboCategories.value = curSel
    Else
        cboCategories.ListIndex = 0
    End If

    Set dSeen = Nothing

End Sub


'### 2.4 Thêm s? ki?n Change c?a combobox (dây là noi kích ho?t b? l?c)

Private Sub cboCategories_Change()

    If isLoading Then Exit Sub

    mCategoryFilter = GetCurrentCategoryFilter()

    Dim dnKw As String
    On Error Resume Next
    dnKw = Trim(txtDNSearch.Text)
    On Error GoTo 0
    If dnKw = "DN" Then dnKw = ""

    Dim mode As PaymentFilterMode
    On Error Resume Next
    If chkUnpaid.value Then mode = pfUnpaid Else mode = pfall
    On Error GoTo 0

    LoadGroupSummary
    LoadUnpaidContracts mode, dnKw, mCategoryFilter

End Sub

Private Function BuildCategoryDictFromContract() As Object
    Dim dCat As Object
    Set dCat = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(gArrContract) Then
        If IsArray(gArrContract) Then
            Dim i As Long
            Dim cid As String
            Dim catVal As String
            For i = 1 To UBound(gArrContract, 1)
                cid = Trim$(CStr(gArrContract(i, 1)))      ' cot A = ContractID
                catVal = Trim$(CStr(gArrContract(i, 18)))  ' cot R = Category
                If cid <> "" Then
                    dCat(cid) = catVal
                End If
            Next i
        End If
    End If
    Set BuildCategoryDictFromContract = dCat
End Function







