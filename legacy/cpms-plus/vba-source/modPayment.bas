Attribute VB_Name = "modPayment"
Option Explicit

Public Enum PaymentFilterMode
    pfall = 0
    pfUnpaid = 1
End Enum


Public Function EnsurePaymentLedger() As Worksheet
    Set EnsurePaymentLedger = GetDataSheet("PAYMENT_LEDGER")
End Function

Public Function FastVN(ByVal v As Variant) As String

    If IsNumeric(v) Then

        If CDbl(v) <> 0 Then

            FastVN = _
                Replace$( _
                Format$(CDbl(v), "#,##0"), _
                ",", ".")

        Else

            FastVN = ""

        End If

    Else

        FastVN = ""

    End If

End Function


Public Function FormatVN(v As Variant) As String
    If IsNumeric(v) And v <> 0 Then
        FormatVN = Replace(Format(v, "#,##0"), ",", ".")
    Else
        FormatVN = ""
    End If
End Function

Public Sub BuildLedgerFromContractDB(Optional dummy As Variant = 0)

    Dim wsC As Worksheet, wsL As Worksheet
    Dim rC As Long, rL As Long
    Dim lastC As Long
    Dim curGroupID As String

    Set wsC = EnsureContractDB
    Set wsL = EnsurePaymentLedger

    ' ===== CLEAR LEDGER DATA (GI? HEADER) =====
    If wsL.Cells(wsL.rows.count, 1).End(xlUp).Row >= 2 Then
        wsL.rows("2:" & wsL.rows.count).ClearContents
        wsL.rows("2:" & wsL.rows.count).Font.Bold = False
    End If

    rL = 2
    lastC = wsC.Cells(wsC.rows.count, 1).End(xlUp).Row
    curGroupID = ""

    For rC = 2 To lastC

        Dim cid As String
        Dim cno As String
        Dim signDate As Variant

        cid = Trim(wsC.Cells(rC, "A").value)
        cno = Trim(wsC.Cells(rC, "B").value)
        signDate = wsC.Cells(rC, "D").value

        ' =========================
        ' GROUP HEADER
        ' Condition:
        ' - ContractID có d? li?u
        ' - SigningDate tr?ng
        ' =========================
        If cid <> "" And Trim(signDate) = "" Then

            curGroupID = cid

            wsL.Cells(rL, 3).value = curGroupID          ' ContractID (group)
            wsL.Cells(rL, 6).value = cno                 ' Group name
            'wsL.Cells(rL, 22).Value = Now

            wsL.rows(rL).Font.Bold = True
            rL = rL + 1

        ' =========================
        ' CONTRACT ROW
        ' =========================
        ElseIf cid <> "" And IsDate(signDate) Then

            wsL.Cells(rL, 3).value = cid                 ' ContractID
            wsL.Cells(rL, 4).value = cno                 ' ContractNo
            wsL.Cells(rL, 6).value = wsC.Cells(rC, 11).value  ' ScopeOfWork
            wsL.Cells(rL, 7).value = wsC.Cells(rC, 13).value                  ' Ratio (tính sau)
            wsL.Cells(rL, 8).value = wsC.Cells(rC, 14).value ' EstimatedValue
            wsL.Cells(rL, 9).value = wsC.Cells(rC, 15).value ' ContractValue
            wsL.Cells(rL, 22).value = Now

            rL = rL + 1

        End If

    Next rC
    Call NormalizeSheetView(wsL)
    Call FormatPaymentLedger(wsL)
    Call FormatTableWithGroup(wsL)


End Sub




Public Sub FormatPaymentLedger(ws As Worksheet)

    With ws
        ' ===== TEXT =====
        .Columns("B:D").NumberFormat = "@"
        .Columns("F:F").WrapText = True
        .Columns("O:P").NumberFormat = "@"
        .Columns("S:S").NumberFormat = "@"
        .Columns("U:U").NumberFormat = "@"

        ' ===== NUMBER =====
        .Columns("E:E").NumberFormat = "0"
        .Columns("G:N").NumberFormat = "#,##0"
        .Columns("R:R").NumberFormat = "#,##0"

        ' ===== PERCENT =====
        '.Columns("G:G").NumberFormat = "0.00%"

        ' ===== DATE =====
        .Columns("Q:Q").NumberFormat = "dd/mm/yyyy"
        .Columns("T:T").NumberFormat = "dd/mm/yyyy"
        .Columns("V:V").NumberFormat = "dd/mm/yyyy hh:mm"

        ' ===== VISUAL =====
        .rows(1).Font.Bold = True
        .rows(1).HorizontalAlignment = xlCenter
        .Columns.AutoFit
    End With

End Sub



Public Function GetSafeDouble(val As Variant) As Double
    If IsNumeric(val) Then
        GetSafeDouble = CDbl(val)
    Else
        GetSafeDouble = 0
    End If
End Function

Public Sub NormalizeSheetView(ws As Worksheet)

    With ws
        ' T?t wrap text toàn sheet
        .Cells.WrapText = False

        ' AutoFit c?t
        .Columns.AutoFit

        ' Can ch?nh l?i cho d?p (tu? ch?n)
        .Cells.VerticalAlignment = xlCenter
    End With

End Sub




Public Function IsGroupRow(ws As Worksheet, r As Long) As Boolean

    ' CONTRACT SHEET:
    ' ContractID có d? li?u + SigningDate tr?ng
    If ws.Name = EnsureContractDB.Name Then
        If Trim(ws.Cells(r, 1).value) <> "" _
           And Trim(ws.Cells(r, 4).value) = "" Then
            IsGroupRow = True
            Exit Function
        End If
    End If

    ' LEDGER SHEET:
    ' ContractID b?t d?u b?ng GROUP-
    If ws.Name = EnsurePaymentLedger.Name Then
        If Trim(ws.Cells(r, 3).value) <> "" _
           And Trim(ws.Cells(r, 1).value) = "" Then
            IsGroupRow = True
            Exit Function
        End If
    End If

    IsGroupRow = False

End Function



Public Sub FormatGroupRow(rowRange As Range)

    With rowRange
        .Font.Bold = True
        .Interior.Color = RGB(255, 228, 225) ' Rose nh?t 80%
        .Borders.LineStyle = xlContinuous
    End With

End Sub



Public Function GetLevelSoftBlue(level As Long) As Long
    Select Case level
        Case 1: GetLevelSoftBlue = RGB(248, 250, 252)
        Case 2: GetLevelSoftBlue = RGB(208, 226, 242)
        Case 3: GetLevelSoftBlue = RGB(228, 238, 248)
        Case 4: GetLevelSoftBlue = RGB(218, 232, 245)
        Case Else
            GetLevelSoftBlue = RGB(238, 244, 250)
    End Select
End Function




Public Function FindLedgerInsertRow(ledgerGroup As String) As Long

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim inGroup As Boolean
    Dim curID As String

    Set ws = EnsurePaymentLedger
    lastRow = ws.Cells(ws.rows.count, 3).End(xlUp).Row

    FindLedgerInsertRow = 0
    inGroup = False

    For r = 2 To lastRow

        curID = Trim(ws.Cells(r, 3).value)

        ' === G?P GROUP HEADER ===
        If curID = ledgerGroup Then
            inGroup = True
            FindLedgerInsertRow = r + 1
            GoTo NEXT_R
        End If

        ' === ÐANG TRONG GROUP ===
        If inGroup Then
            ' g?p group khác ? d?ng
            If curID Like "GROUP-*" Then Exit Function

            ' contract / appendix ? c?p nh?t v? trí cu?i
            FindLedgerInsertRow = r + 1
        End If

NEXT_R:
    Next r

End Function



Public Sub InsertLedgerForAppendix(parentID As String, appendixID As String)

    Dim wsLedger As Worksheet
    Dim wsDB As Worksheet
    Dim r As Long
    Dim dbRow As Long

    Set wsLedger = EnsurePaymentLedger
    Set wsDB = EnsureContractDB

    dbRow = FindContractRowByID(wsDB, appendixID)
    If dbRow <= 0 Then Exit Sub

    r = FindLedgerInsertRowByParent(parentID)
    If r <= 0 Then Exit Sub

    wsLedger.rows(r).Insert Shift:=xlDown

    With wsLedger
        ' ===== PARENT ID (C?T A) =====
        .Cells(r, "A").value = parentID

        ' ===== CORE =====
        .Cells(r, "C").value = appendixID
        .Cells(r, "D").value = wsDB.Cells(dbRow, 2).value
        .Cells(r, "F").value = wsDB.Cells(dbRow, 11).value

        ' ===== FINANCIAL =====
        .Cells(r, "G").value = wsDB.Cells(dbRow, 13).value
        .Cells(r, "H").value = wsDB.Cells(dbRow, 14).value
        .Cells(r, "I").value = wsDB.Cells(dbRow, 15).value
        .Cells(r, "V").value = wsDB.Cells(dbRow, 23).value

        .rows(r).Font.Bold = False
    End With

    ' ===== ÁP FORMULA CHO HEADER APPENDIX =====
    ApplyFormulaToHeader wsLedger, r
    FormatLedgerRowSingle wsLedger, r

End Sub




Public Sub InsertLedgerForContract(contractID As String)

    Dim wsLedger As Worksheet
    Dim wsDB As Worksheet
    Dim ledgerGroup As String
    Dim r As Long
    Dim dbRow As Long

    Set wsLedger = EnsurePaymentLedger
    Set wsDB = EnsureContractDB

    ' === xác d?nh GROUP g?c ===
    ledgerGroup = GetRootLedgerGroupFromContractDB(contractID)
    If ledgerGroup = "" Then Exit Sub

    Call EnsureLedgerGroup(ledgerGroup, GetGroupName(Right(ledgerGroup, 1)))

    ' === tìm contract trong DB ===
    dbRow = FindContractRowByID(wsDB, contractID)
    If dbRow <= 0 Then Exit Sub

    ' === tìm v? trí insert ===
    r = FindLedgerInsertRow(ledgerGroup)
    If r <= 0 Then Exit Sub

    wsLedger.rows(r).Insert Shift:=xlDown

    With wsLedger
        ' ===== PARENT ID (C?T A) =====
        .Cells(r, "A").value = ledgerGroup

        ' ===== CORE =====
        .Cells(r, "C").value = contractID
        .Cells(r, "D").value = wsDB.Cells(dbRow, 2).value
        .Cells(r, "F").value = wsDB.Cells(dbRow, 11).value

        ' ===== FINANCIAL =====
        .Cells(r, "G").value = wsDB.Cells(dbRow, 13).value
        .Cells(r, "H").value = wsDB.Cells(dbRow, 14).value
        .Cells(r, "I").value = wsDB.Cells(dbRow, 15).value
        .Cells(r, "V").value = wsDB.Cells(dbRow, 23).value

        .rows(r).Font.Bold = False
    End With

    ' ===== ÁP FORMULA CHO CHÍNH DÒNG NÀY =====
    ApplyFormulaToHeader wsLedger, r
    FormatLedgerRowSingle wsLedger, r

End Sub




Public Function GetGroupName(groupCode As String) As String
    Select Case groupCode

        ' H?p d?ng v?i ch? d?u tu
        Case "A"
            GetGroupName = _
                ChrW(72) & ChrW(7907) & "p " & _
                ChrW(273) & ChrW(7891) & "ng v" & _
                ChrW(7899) & "i ch" & ChrW(7911) & " " & _
                ChrW(273) & ChrW(7847) & "u t" & ChrW(432)

        ' H?p d?ng v?i nhà th?u
        Case "B"
            GetGroupName = _
                ChrW(72) & ChrW(7907) & "p " & _
                ChrW(273) & ChrW(7891) & "ng v" & _
                ChrW(7899) & "i nh" & ChrW(224) & " th" & ChrW(7847) & "u"

        Case Else
            GetGroupName = "Contract Group"

    End Select
End Function

Public Sub EnsureLedgerGroup(groupID As String, groupName As String)

    Dim ws As Worksheet, r As Long
    Set ws = EnsurePaymentLedger

    ' Ðã t?n t?i ? thoát
    If GroupExists(ws, groupID, 3) Then Exit Sub

    r = ws.Cells(ws.rows.count, 3).End(xlUp).Row + 1
    If r < 2 Then r = 2

    With ws
        .Cells(r, 3).value = groupID     ' ContractID
        .Cells(r, 6).value = groupName   ' ItemName
        .rows(r).Font.Bold = True
    End With

End Sub
Public Function GroupExists(ws As Worksheet, groupID As String, idCol As Long) As Boolean
    Dim r As Long
    For r = 2 To ws.Cells(ws.rows.count, idCol).End(xlUp).Row
        If Trim(ws.Cells(r, idCol).value) = groupID Then
            GroupExists = True
            Exit Function
        End If
    Next r
    GroupExists = False
End Function

Public Function GetLedgerGroupByContractID(contractID As String) As String
    If contractID Like "GROUP-A*" Or contractID Like "GROUP-B*" Then
        GetLedgerGroupByContractID = contractID
    Else
        ' root contract ? d?a vào MAIN/EPC/SUB dã map s?n
        If contractID <> "" Then
            ' l?y group t? ContractDB
            GetLedgerGroupByContractID = GetRootLedgerGroupFromContractDB(contractID)
        End If
    End If
End Function

Public Function GetRootLedgerGroupFromContractDB(contractID As String) As String

    Dim ws As Worksheet, r As Long
    Set ws = EnsureContractDB

    For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).Row

        If ws.Cells(r, 1).value = contractID Then
            ' di ngu?c lên tìm GROUP-*
            Do While r >= 2
                If ws.Cells(r, 1).value Like "GROUP-*" Then
                    GetRootLedgerGroupFromContractDB = ws.Cells(r, 1).value
                    Exit Function
                End If
                r = r - 1
            Loop
        End If

    Next r

End Function

Public Function FindLedgerInsertRowByParent(parentID As String) As Long

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim inTree As Boolean
    Dim curID As String

    Set ws = EnsurePaymentLedger
    lastRow = ws.Cells(ws.rows.count, 3).End(xlUp).Row

    FindLedgerInsertRowByParent = 0
    inTree = False

    For r = 2 To lastRow

        curID = Trim(ws.Cells(r, 3).value)

        ' === g?p parent ===
        If curID = parentID Then
            inTree = True
            FindLedgerInsertRowByParent = r + 1
            GoTo NEXT_R
        End If

        ' === dang trong subtree ===
        If inTree Then
            If curID Like parentID & "-*" Then
                FindLedgerInsertRowByParent = r + 1
            Else
                Exit Function
            End If
        End If

NEXT_R:
    Next r

End Function


Public Function FindInsertRowForPayment(ws As Worksheet, contractID As String) As Long

    Dim lastRow As Long, r As Long
    Dim foundContract As Boolean

    lastRow = ws.Cells(ws.rows.count, "C").End(xlUp).Row
    FindInsertRowForPayment = 0
    foundContract = False

    For r = 2 To lastRow

        ' G?p ledger contract / appendix
        If Trim(ws.Cells(r, "C").value) = contractID _
           And ws.Cells(r, "B").value = "" Then
            foundContract = True
            FindInsertRowForPayment = r + 1
            GoTo NEXT_R
        End If

        ' Ðang trong payment c?a contract dó
        If foundContract Then
            ' G?p contract khác ? d?ng
            If Trim(ws.Cells(r, "C").value) <> contractID Then Exit Function
            FindInsertRowForPayment = r + 1
        End If

NEXT_R:
    Next r

End Function




Public Function GetContractLevel(contractID As String) As Long
    If contractID Like "GROUP-*" Then
        GetContractLevel = 0
    Else
        GetContractLevel = UBound(Split(contractID, "-"))
    End If
End Function

Public Sub ApplyFormulaToHeader(ws As Worksheet, headerRow As Long)

    Dim id As String
    id = Trim(ws.Cells(headerRow, "C").value)
    If id = "" Then Exit Sub

    ' ===== CH? HEADER (KHÔNG PH?I PAYMENT) =====
    If Trim(ws.Cells(headerRow, "B").value) <> "" Then Exit Sub

    ' =================================================
    ' W : Actual VND (SUM FROM CHILDREN)
    ' =================================================
    ws.Cells(headerRow, "W").Formula = _
        "=SUMIFS($W:$W,$A:$A,C" & headerRow & ")"

    ' =================================================
    ' X : Accepted VND
    ' =================================================
    ws.Cells(headerRow, "X").Formula = _
        "=SUMIFS($X:$X,$A:$A,C" & headerRow & ")"

    ' =================================================
    ' Y : Request VND
    ' =================================================
    ws.Cells(headerRow, "Y").Formula = _
        "=SUMIFS($Y:$Y,$A:$A,C" & headerRow & ")"

    ' =================================================
    ' Z : Paid VND
    ' =================================================
    ws.Cells(headerRow, "Z").Formula = _
        "=SUMIFS($Z:$Z,$A:$A,C" & headerRow & ")"

    ' =================================================
    ' J K L M N (HI?N TH? CHÍNH)
    ' =================================================
    ws.Cells(headerRow, "J").Formula = "=W" & headerRow
    ws.Cells(headerRow, "K").Formula = "=X" & headerRow
    ws.Cells(headerRow, "L").Formula = "=Y" & headerRow
    ws.Cells(headerRow, "M").Formula = "=Z" & headerRow
    ws.Cells(headerRow, "N").Formula = "=L" & headerRow & "-M" & headerRow

End Sub



Public Sub ApplyFormulaToAllHeaders(Optional dummy As Variant = 0)

    Dim ws As Worksheet
    Set ws = EnsurePaymentLedger

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.rows.count, "C").End(xlUp).Row

    For r = 2 To lastRow
        If Trim(ws.Cells(r, "C").value) <> "" _
           And Trim(ws.Cells(r, "B").value) = "" Then

            Call ApplyFormulaToHeader(ws, r)

        End If
    Next r

End Sub




Public Function FindLedgerRowByContractID(ws As Worksheet, contractID As String) As Long

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.rows.count, "C").End(xlUp).Row

    For r = 2 To lastRow
        If Trim(ws.Cells(r, "C").value) = contractID _
           And Trim(ws.Cells(r, "B").value) = "" Then
            FindLedgerRowByContractID = r
            Exit Function
        End If
    Next r

    FindLedgerRowByContractID = 0

End Function

Public Function GetParentContractID(contractID As String) As String

    Dim p As Long

    contractID = Trim(contractID)
    If contractID = "" Then Exit Function

    ' GROUP không có parent
    If contractID Like "GROUP-*" Then Exit Function

    ' tìm d?u "-" cu?i cùng
    p = InStrRev(contractID, "-")
    If p <= 0 Then Exit Function

    GetParentContractID = Left(contractID, p - 1)

End Function

Public Sub NormalizeLedgerParentID(Optional dummy As Variant = 0)

    Dim ws As Worksheet
    Set ws = EnsurePaymentLedger

    Dim lastRow As Long, r As Long
    Dim cid As String, pid As String
    Dim dashCount As Long

    lastRow = ws.Cells(ws.rows.count, "C").End(xlUp).Row

    Application.ScreenUpdating = False

    For r = 2 To lastRow

        cid = Trim(ws.Cells(r, "C").value)

        ' ===== TR?NG =====
        If cid = "" Then
            ws.Cells(r, "A").value = ""
            GoTo NEXT_R
        End If

        ' ===== GROUP HEADER =====
        If cid Like "GROUP-*" Then
            ws.Cells(r, "A").value = ""
            GoTo NEXT_R
        End If

        ' ===== PAYMENT ROW (CÓ LEDGERID) =====
        If Trim(ws.Cells(r, "B").value) <> "" Then
            ws.Cells(r, "A").value = cid
            GoTo NEXT_R
        End If

        ' ===== HEADER CONTRACT / APPENDIX =====
        dashCount = CountDash(cid)

        Select Case dashCount

            ' ===== CONTRACT C?P 1 =====
            Case 2
                ws.Cells(r, "A").value = GetLedgerGroupByContractID(cid)

            ' ===== APPENDIX C?P >= 2 =====
            Case Is >= 3
                ws.Cells(r, "A").value = Left(cid, InStrRev(cid, "-") - 1)

            Case Else
                ws.Cells(r, "A").value = ""

        End Select

NEXT_R:
    Next r

    Application.ScreenUpdating = True

End Sub


Public Function CountDash(ByVal s As String) As Long
    CountDash = Len(s) - Len(Replace(s, "-", ""))
End Function

'================================================================
' HELPER: Viet chi tiet thanh toan cho mot nhom (A hoac B)
' Co dong subtotal cuoi moi thang
' Tra ve row tiep theo sau khi viet xong
'================================================================
Private Function WriteDetailGroup( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByRef detailArr() As Variant, _
    ByVal detailCount As Long, _
    ByVal grpFilter As String, _
    ByRef groupTotal As Double) As Long

    '  Col layout:
    '  1=#   2=Date   3=Contract No   4=Description   5=Ref No
    '  6=Paid Amount (VND)   7=Ty gia   8=Goc (USD)

    Dim i As Long
    Dim r As Long

    Dim lastMonthShown As String
    Dim monthTotal As Double
    Dim altRow As Long
    Dim seqNo As Long

    Dim dYM As String
    Dim yr As Long
    Dim mo As Long
    Dim prevYr As Long
    Dim prevMo As Long

    r = startRow
    lastMonthShown = ""
    monthTotal = 0
    groupTotal = 0
    altRow = 0
    seqNo = 0

    For i = 0 To detailCount - 1

        If CStr(detailArr(i, 1)) <> grpFilter Then GoTo NEXT_D

        dYM = CStr(detailArr(i, 0))

        '====================
        ' MONTH CHANGED -> SUBTOTAL
        '====================
        If dYM <> lastMonthShown And lastMonthShown <> "" Then

            With ws.Range(ws.Cells(r, 1), ws.Cells(r, 8))
                .Interior.Color = RGB(255, 242, 204)
                .Font.Bold = True
                .VerticalAlignment = xlVAlignCenter
            End With

            prevYr = CLng(Left$(lastMonthShown, 4))
            prevMo = CLng(Right$(lastMonthShown, 2))

            ws.Cells(r, 1).NumberFormat = "@"
            ws.Cells(r, 1).value = _
                "Total " & Format(DateSerial(prevYr, prevMo, 1), "mmm-yyyy")
            ws.Cells(r, 1).Font.Italic = True

            ws.Cells(r, 6).value = monthTotal
            ws.Cells(r, 6).NumberFormat = "#,##0"
            ws.Cells(r, 6).Font.Bold = True

            r = r + 1
            monthTotal = 0
            altRow = 0
            seqNo = 0

        End If

        '====================
        ' MONTH SEPARATOR
        '====================
        If dYM <> lastMonthShown Then

            yr = CLng(Left$(dYM, 4))
            mo = CLng(Right$(dYM, 2))

            With ws.Range(ws.Cells(r, 1), ws.Cells(r, 8))
                .Interior.Color = RGB(189, 215, 238)
                .Font.Bold = True
                .Font.Color = RGB(0, 0, 0)
                .VerticalAlignment = xlVAlignCenter
            End With

            ws.Cells(r, 1).NumberFormat = "@"
            ws.Cells(r, 1).value = _
                Format(DateSerial(yr, mo, 1), "mmmm yyyy")

            r = r + 1
            lastMonthShown = dYM

        End If

        '====================
        ' DATA ROW
        '====================
        seqNo = seqNo + 1

        ' Col 1: # (sequence in month)
        ws.Cells(r, 1).value = seqNo
        ws.Cells(r, 1).HorizontalAlignment = xlCenter
        ws.Cells(r, 1).VerticalAlignment = xlVAlignCenter

        ' Col 2: Date — center aligned
        ws.Cells(r, 2).value = detailArr(i, 6)
        ws.Cells(r, 2).NumberFormat = "dd/mm/yyyy"
        ws.Cells(r, 2).HorizontalAlignment = xlCenter
        ws.Cells(r, 2).VerticalAlignment = xlVAlignCenter

        ' Col 3: Contract No — [FIX 3] ep format text tranh Excel tu dong doi date-like string
        With ws.Cells(r, 3)
            .NumberFormat = "@"
            .value = CStr(detailArr(i, 3))
            .VerticalAlignment = xlVAlignCenter
        End With

        ' Col 4: Description — [FIX 4] luu dinh dang text, wrap text, middle align
        With ws.Cells(r, 4)
            .NumberFormat = "@"
            .value = detailArr(i, 4)
            .WrapText = True
            .VerticalAlignment = xlVAlignCenter
        End With

        ' Col 5: Ref No
        ws.Cells(r, 5).value = detailArr(i, 7)
        ws.Cells(r, 5).HorizontalAlignment = xlCenter
        ws.Cells(r, 5).VerticalAlignment = xlVAlignCenter

        ' Col 6: Paid Amount (VND — da quy doi neu can)
        ws.Cells(r, 6).value = detailArr(i, 5)
        ws.Cells(r, 6).NumberFormat = "#,##0"
        ws.Cells(r, 6).VerticalAlignment = xlVAlignCenter

        ' Col 7 & 8: Ty gia va Goc (chi hien thi neu la ngoai te)
        If CDbl(detailArr(i, 8)) > 1 Then

            ' Col 7: Ty gia
            ws.Cells(r, 7).value = detailArr(i, 8)
            ws.Cells(r, 7).NumberFormat = "#,##0.000"
            ws.Cells(r, 7).HorizontalAlignment = xlRight
            ws.Cells(r, 7).VerticalAlignment = xlVAlignCenter

            ' Col 8: So tien goc (USD)
            ws.Cells(r, 8).value = detailArr(i, 9)
            ws.Cells(r, 8).NumberFormat = "#,##0.##"
            ws.Cells(r, 8).HorizontalAlignment = xlRight
            ws.Cells(r, 8).VerticalAlignment = xlVAlignCenter

        Else

            ws.Cells(r, 7).value = ""
            ws.Cells(r, 8).value = ""

        End If

        If (altRow Mod 2) = 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)) _
                .Interior.Color = RGB(235, 241, 252)
        End If

        monthTotal = monthTotal + CDbl(detailArr(i, 5))
        groupTotal = groupTotal + CDbl(detailArr(i, 5))
        altRow = altRow + 1
        r = r + 1

NEXT_D:
    Next i

    '====================
    ' SUBTOTAL LAST MONTH
    '====================
    If lastMonthShown <> "" And monthTotal > 0 Then

        With ws.Range(ws.Cells(r, 1), ws.Cells(r, 8))
            .Interior.Color = RGB(255, 242, 204)
            .Font.Bold = True
            .VerticalAlignment = xlVAlignCenter
        End With

        prevYr = CLng(Left$(lastMonthShown, 4))
        prevMo = CLng(Right$(lastMonthShown, 2))

        ws.Cells(r, 1).NumberFormat = "@"
        ws.Cells(r, 1).value = _
            "Total " & Format(DateSerial(prevYr, prevMo, 1), "mmm-yyyy")
        ws.Cells(r, 1).Font.Italic = True

        ws.Cells(r, 6).value = monthTotal
        ws.Cells(r, 6).NumberFormat = "#,##0"
        ws.Cells(r, 6).Font.Bold = True

        r = r + 1

    End If

    WriteDetailGroup = r

End Function

'================================================================
' HELPER: Xac dinh nhom (A/B) cua mot contractID bang cach
' quet qua gArrContract va tim dong GROUP-A / GROUP-B gan nhat
' phia tren. Dung khi contractID khong bat dau bang "A-"/"B-"
' (vi du payment duoc nhap voi so hop dong thay vi internal ID).
' Tra ve "A", "B", hoac "" neu khong tim thay.
'================================================================
Public Function GetGroupFromContractArrold(ByVal contractID As String) As String

    GetGroupFromContractArr = ""

    If IsEmpty(gArrContract) Then Exit Function
    If gDictContract Is Nothing Then Exit Function

    ' Buoc 1: tim dong cua contractID trong gArrContract
    If Not gDictContract.Exists(contractID) Then Exit Function

    Dim targetRow As Long
    targetRow = gDictContract(contractID)

    ' Buoc 2: quet nguoc tu targetRow len tren de tim marker GROUP-A/GROUP-B
    Dim j As Long
    Dim id As String

    For j = targetRow To 1 Step -1

        id = Trim$(CStr(gArrContract(j, 1)))

        If id = "GROUP-A" Then
            GetGroupFromContractArr = "A"
            Exit Function
        ElseIf id = "GROUP-B" Then
            GetGroupFromContractArr = "B"
            Exit Function
        End If

    Next j

End Function

'================================================================
' HELPER: Xac dinh nhom (A/B) cua mot contractID bang cach
' quet qua gArrContract va tim dong GROUP-A / GROUP-B gan nhat
' phia tren. Dung khi contractID khong bat dau bang "A-"/"B-"
' Tra ve "A", "B", hoac "" neu khong tim thay.
'================================================================
Public Function GetGroupFromContractArr(ByVal contractID As String) As String

    GetGroupFromContractArr = ""

    If IsEmpty(gArrContract) Then Exit Function
    If gDictContract Is Nothing Then Exit Function

    If Not gDictContract.Exists(contractID) Then Exit Function

    Dim targetRow As Long
    targetRow = gDictContract(contractID)

    Dim j As Long
    Dim id As String

    For j = targetRow To 1 Step -1

        id = Trim$(CStr(gArrContract(j, 1)))

        If id = "GROUP-A" Then
            GetGroupFromContractArr = "A"
            Exit Function
        ElseIf id = "GROUP-B" Then
            GetGroupFromContractArr = "B"
            Exit Function
        End If

    Next j

End Function

'================================================================
' HELPER: Lay Status cua hop dong tu CONTRACT_DB
' Tra ve chuoi uppercase: "ACTIVE", "CANCEL", "COMPLETED", hoac ""
' Col F (index 6) cua gArrContract = Status
'================================================================
Public Function GetContractStatus(ByVal contractID As String) As String

    GetContractStatus = ""
    If IsEmpty(gArrContract) Then Exit Function
    If gDictContract Is Nothing Then Exit Function
    If Not gDictContract.Exists(contractID) Then Exit Function

    Dim rowIdx As Long
    rowIdx = gDictContract(contractID)
    GetContractStatus = UCase$(Trim$(CStr(gArrContract(rowIdx, 6))))

End Function


