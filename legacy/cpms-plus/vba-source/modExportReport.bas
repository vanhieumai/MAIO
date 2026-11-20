Attribute VB_Name = "modExportReport"
'Attribute VB_Name = "modExportReport"
'================================================================
' MODULE: modExportReport  [MOD v8 - fix compile errors]
'
' Layout Payment Detail nhieu hop dong (cCount > 1) — A:L 12 cot:
'   A=# | B=Date | C:E=Contract No(merge 3) | F:H=Description(merge 3) |
'   I=Ref No | J=Paid Amount(VND) | K=Exchange Rate | L=Foreign Amount
'   paidColMain = 10 (cot J)
'
' Layout Payment Detail 1 hop dong (cCount = 1) — A:L 12 cot:
'   A=# | B=Date | C:H=Description(merge 6) |
'   I=Ref No | J=Paid Amount(VND) | K=Exchange Rate | L=Foreign Amount
'   paidColMain = 10 (cot J)
'================================================================
Option Explicit

Private Const NCOLS_SINGLE As Long = 11
Private Const NCOLS_MULTI  As Long = 12

' Shadow (temp) cols: ghi bong noi dung merge de AutoFit tinh dung chieu cao dong
Private mShCI_A    As Long   ' CI multi  : CompanyA   F:G -> col 13 (M)  w=36
Private mShCI_B    As Long   ' CI multi  : CompanyB   H:I -> col 14 (N)  w=36
Private mShCI_SOW  As Long   ' CI multi  : ScopeOfWork J:K-> col 15 (O)  w=36
Private mShCI_Desc As Long   ' CI single : Description C:J-> col 13 (M)  w=144
Private mShDT_Note As Long   ' DocTracking Note J:K -> multi=16(P) single=14(N) w=36
Private mShPay_CNo As Long   ' Payment multi : ContractNo C:E-> col 17 (Q) w=54
Private mShPay_Dsc As Long   ' Payment: Desc -> multi=18(R) w=54 / single=15(O) w=108
Private mShLastCol As Long   ' col cuoi cung da dung (de Clear sau AutoFit)

'================================================================
' HELPER: Ke vien mong cho mot dong du lieu
'================================================================
Private Sub ApplyDataRowBorders( _
    ByRef wsOut As Worksheet, _
    ByVal r As Long, _
    ByVal NCOLS As Long, _
    ParamArray colSplit() As Variant)

    Dim bc As Long
    bc = RGB(160, 160, 160)

    With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS))
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = bc
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = bc
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = bc
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = bc
        End With
    End With

    Dim v As Variant
    For Each v In colSplit
        Dim sc As Long
        sc = CLng(v)
        If sc >= 1 And sc < NCOLS Then
            With wsOut.Cells(r, sc).Borders(xlEdgeRight)
                .LineStyle = xlContinuous: .Weight = xlThin: .Color = bc
            End With
        End If
    Next v

End Sub

'================================================================
' HELPER: Do rong cot
'================================================================
'Private Sub SetSingleContractColWidths(ByRef ws As Worksheet)
'    Dim c As Long
'    ws.Columns("A").ColumnWidth = 5
'    For c = 2 To 11
'        ws.Columns(c).ColumnWidth = 20
'    Next c
'End Sub

Private Sub SetMultiContractColWidths(ByRef ws As Worksheet)
    Dim c As Long
    ws.Columns("A").ColumnWidth = 5
    For c = 2 To 12
        ws.Columns(c).ColumnWidth = 18
    Next c
End Sub

'================================================================
' HELPER: Phan tich ngay dd/mm/yyyy
'================================================================
Public Function ParseDMY(v As Variant) As Date
    Dim vt As Integer
    vt = VarType(v)
    If vt = vbDate Or vt = vbDouble Or vt = vbSingle Or vt = vbLong Or vt = vbInteger Then
        ParseDMY = CDate(v)
        Exit Function
    End If
    Dim s As String
    s = Trim$(CStr(v))
    Dim p() As String
    p = Split(s, "/")
    If UBound(p) = 2 Then
        ParseDMY = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
    Else
        Err.Raise 13, "ParseDMY", "Khong the phan tich ngay: " & s
    End If
End Function

Public Function IsValidDMY(v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then
        IsValidDMY = False
        Exit Function
    End If
    Dim s As String
    s = Trim$(CStr(v))
    If s = "" Or s = "0" Then
        IsValidDMY = False
        Exit Function
    End If
    On Error Resume Next
    Dim d As Date
    d = ParseDMY(v)
    IsValidDMY = (Err.Number = 0)
    On Error GoTo 0
End Function

'================================================================
' HELPER: Sort key loai tai lieu
'================================================================
Private Function DocTypeSortKey(ByVal docType As String) As Long
    Select Case UCase$(Trim$(docType))
        Case "HD", "HOP DONG":       DocTypeSortKey = 1
        Case "PLHD":                  DocTypeSortKey = 2
        Case "IPC", "THANH TOAN":    DocTypeSortKey = 3
        Case Else:                    DocTypeSortKey = 4
    End Select
End Function

'================================================================
' HELPER: Lay nhom tu contract ID
'================================================================
Private Function GetGrpFromID(ByVal cid As String) As String
    If Left$(cid, 2) = "A-" Then
        GetGrpFromID = "A"
    ElseIf Left$(cid, 2) = "B-" Then
        GetGrpFromID = "B"
    Else
        GetGrpFromID = GetGroupFromContractArr(cid)
    End If
End Function

'================================================================
' SECTION 1: CONTRACT INFORMATION
'================================================================
Public Function ExportContractInfoSection( _
    ByRef wsOut As Worksheet, _
    ByVal startRow As Long, _
    contractList() As String, _
    ByVal cCount As Long) As Long

    Dim r As Long
    r = startRow

    If gDictContract Is Nothing Then
        ExportContractInfoSection = r
        Exit Function
    End If

    Dim nColsUsed As Long
    If cCount = 1 Then
        nColsUsed = NCOLS_MULTI
    Else
        nColsUsed = NCOLS_MULTI
    End If

    With wsOut.Cells(r, 1)
        .value = "I. CONTRACT INFORMATION"
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
    End With
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, nColsUsed)).Interior.Color = RGB(31, 73, 125)
    r = r + 1

    '------------------------------------------------------------
    ' 1 HOP DONG — bang doc
    '------------------------------------------------------------
    If cCount = 1 Then
        'SetSingleContractColWidths wsOut

        Dim cid1 As String
        cid1 = contractList(0)
        If Not gDictContract.Exists(cid1) Then
            ExportContractInfoSection = r
            Exit Function
        End If

        Dim rr1 As Long
        rr1 = gDictContract(cid1)
        wsOut.Cells(1, 1).value = Trim$(CStr(gArrContract(rr1, 2)))

        Dim grp1 As String
        grp1 = GetGrpFromID(cid1)
        If grp1 = "" Then grp1 = "?"

        With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI))
            .Interior.Color = RGB(68, 114, 196)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .VerticalAlignment = xlVAlignCenter
        End With
        wsOut.Cells(r, 1).value = "GROUP " & grp1
        r = r + 1

        With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI))
            .Interior.Color = RGB(31, 73, 125)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .VerticalAlignment = xlVAlignCenter
        End With
        wsOut.Cells(r, 1).value = "#"
        wsOut.Cells(r, 1).HorizontalAlignment = xlCenter
        wsOut.Cells(r, 2).value = "Name"
        wsOut.Cells(r, 2).HorizontalAlignment = xlCenter

        Application.DisplayAlerts = False
        wsOut.Range(wsOut.Cells(r, 3), wsOut.Cells(r, 10)).UnMerge
        wsOut.Range(wsOut.Cells(r, 3), wsOut.Cells(r, 10)).Merge
        wsOut.Range(wsOut.Cells(r, 11), wsOut.Cells(r, 12)).UnMerge
        wsOut.Range(wsOut.Cells(r, 11), wsOut.Cells(r, 12)).Merge
        Application.DisplayAlerts = True

        wsOut.Cells(r, 3).value = "Description"
        wsOut.Cells(r, 3).HorizontalAlignment = xlCenter
        wsOut.Cells(r, 11).value = "Note"
        wsOut.Cells(r, 11).HorizontalAlignment = xlCenter
        ApplyDataRowBorders wsOut, r, NCOLS_MULTI, 1, 2, 10
        r = r + 1

        Dim fSigDate As String
        Dim fEndDate As String
        fSigDate = ""
        fEndDate = ""
        If IsValidDMY(gArrContract(rr1, 4)) Then
            fSigDate = Format(ParseDMY(gArrContract(rr1, 4)), "dd/mm/yyyy")
        End If
        If IsValidDMY(gArrContract(rr1, 5)) Then
            fEndDate = Format(ParseDMY(gArrContract(rr1, 5)), "dd/mm/yyyy")
        End If

        Dim fValRaw As Variant
        fValRaw = gArrContract(rr1, 15)
        Dim fContractValue As String
        If IsNumeric(fValRaw) Then
            fContractValue = Format(CDbl(fValRaw), "#,##0")
        ElseIf IsNumeric(Replace(CStr(fValRaw), ",", "")) Then
            fContractValue = Format(CDbl(Replace(CStr(fValRaw), ",", "")), "#,##0")
        Else
            fContractValue = Trim$(CStr(fValRaw))
        End If

        Dim fLabels(0 To 10) As String
        Dim fValues(0 To 10) As Variant
        fLabels(0) = "Contract No":     fValues(0) = Trim$(CStr(gArrContract(rr1, 2)))
        fLabels(1) = "Signing Date":    fValues(1) = fSigDate
        fLabels(2) = "End Date":        fValues(2) = fEndDate
        fLabels(3) = "Status":          fValues(3) = Trim$(CStr(gArrContract(rr1, 6)))
        fLabels(4) = "Company A":       fValues(4) = Trim$(CStr(gArrContract(rr1, 7)))
        fLabels(5) = "Company B":       fValues(5) = Trim$(CStr(gArrContract(rr1, 8)))
        fLabels(6) = "Scope of Work":   fValues(6) = Trim$(CStr(gArrContract(rr1, 11)))
        fLabels(7) = "Contract Value":  fValues(7) = fContractValue
        fLabels(8) = "Payee":           fValues(8) = Trim$(CStr(gArrContract(rr1, 19)))
        fLabels(9) = "Account Number":  fValues(9) = Trim$(CStr(gArrContract(rr1, 20)))
        fLabels(10) = "Bank Name":           fValues(10) = Trim$(CStr(gArrContract(rr1, 21)))

        Dim fi As Long
        Application.DisplayAlerts = False
        For fi = 0 To 10
            Dim fColor As Long
            If fi Mod 2 = 0 Then
                fColor = RGB(255, 255, 255)
            Else
                fColor = RGB(235, 241, 252)
            End If
            wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI)).Interior.Color = fColor

            wsOut.Cells(r, 1).value = fi + 1
            wsOut.Cells(r, 1).HorizontalAlignment = xlCenter
            wsOut.Cells(r, 1).VerticalAlignment = xlVAlignCenter

            wsOut.Cells(r, 2).value = fLabels(fi)
            wsOut.Cells(r, 2).VerticalAlignment = xlVAlignCenter

            wsOut.Range(wsOut.Cells(r, 3), wsOut.Cells(r, 10)).UnMerge
            wsOut.Range(wsOut.Cells(r, 3), wsOut.Cells(r, 10)).Merge
            With wsOut.Cells(r, 3)
                .value = fValues(fi)
                .WrapText = True
                .VerticalAlignment = xlVAlignCenter
                .HorizontalAlignment = xlLeft
            End With
            If mShCI_Desc > 0 Then
                wsOut.Cells(r, mShCI_Desc).value = fValues(fi)
                wsOut.Cells(r, mShCI_Desc).WrapText = True
            End If

            wsOut.Range(wsOut.Cells(r, 11), wsOut.Cells(r, 12)).UnMerge
            wsOut.Range(wsOut.Cells(r, 11), wsOut.Cells(r, 12)).Merge
            wsOut.Cells(r, 11).value = ""
            wsOut.Cells(r, 11).VerticalAlignment = xlVAlignCenter

            ApplyDataRowBorders wsOut, r, NCOLS_MULTI, 1, 2, 10
            r = r + 1
        Next fi
        Application.DisplayAlerts = True

        wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI)).Interior.Color = RGB(31, 73, 125)
        r = r + 1
        ExportContractInfoSection = r
        Exit Function
    End If

    '------------------------------------------------------------
    ' NHIEU HOP DONG — bang ngang 12 cot (A:L)
    ' A=# | B=ContractNo | C=SignDate | D=EndDate | E=Status |
    ' F:G=CompanyA | H:I=CompanyB | J:K=ScopeOfWork | L=Value
    '------------------------------------------------------------
    SetMultiContractColWidths wsOut

    Dim colInfo() As Variant
    ReDim colInfo(0 To cCount - 1, 0 To 3)
    Dim colCount As Long
    colCount = 0

    Dim i As Long
    For i = 0 To cCount - 1
        Dim cid As String
        cid = contractList(i)
        If Not gDictContract.Exists(cid) Then GoTo CI_NEXT

        Dim rr As Long
        rr = gDictContract(cid)

        Dim grp As String
        grp = GetGrpFromID(cid)
        If grp = "" Then GoTo CI_NEXT

        Dim sigDbl As Double
        sigDbl = 0
        If IsValidDMY(gArrContract(rr, 4)) Then
            sigDbl = CDbl(ParseDMY(gArrContract(rr, 4)))
        End If

        colInfo(colCount, 0) = cid
        colInfo(colCount, 1) = rr
        colInfo(colCount, 2) = sigDbl
        colInfo(colCount, 3) = grp
        colCount = colCount + 1
CI_NEXT:
    Next i

    Dim s1 As Long, s2 As Long, c As Long
    Dim tmpV(0 To 3) As Variant
    For s1 = 0 To colCount - 2
        For s2 = s1 + 1 To colCount - 1
            If colInfo(s1, 3) = colInfo(s2, 3) Then
                Dim d1 As Double
                Dim d2 As Double
                d1 = CDbl(colInfo(s1, 2))
                d2 = CDbl(colInfo(s2, 2))
                Dim doSwap As Boolean
                doSwap = False
                If d1 = 0 And d2 > 0 Then
                    doSwap = True
                ElseIf d2 > 0 And d2 < d1 Then
                    doSwap = True
                End If
                If doSwap Then
                    For c = 0 To 3: tmpV(c) = colInfo(s1, c): Next c
                    For c = 0 To 3: colInfo(s1, c) = colInfo(s2, c): Next c
                    For c = 0 To 3: colInfo(s2, c) = tmpV(c): Next c
                End If
            End If
        Next s2
    Next s1

    Dim totalA_cnt As Long, totalB_cnt As Long
    totalA_cnt = 0: totalB_cnt = 0
    For i = 0 To colCount - 1
        If colInfo(i, 3) = "A" Then
            totalA_cnt = totalA_cnt + 1
        Else
            totalB_cnt = totalB_cnt + 1
        End If
    Next i

    Dim grpPass As Long
    For grpPass = 1 To 2
        Dim curGrp As String
        Dim grpLabel As String
        If grpPass = 1 Then
            curGrp = "A"
            grpLabel = "GROUP A  (" & totalA_cnt & " contract(s))"
        Else
            curGrp = "B"
            grpLabel = "GROUP B  (" & totalB_cnt & " contract(s))"
        End If

        Dim hasGrp As Boolean
        hasGrp = False
        For i = 0 To colCount - 1
            If colInfo(i, 3) = curGrp Then
                hasGrp = True
                Exit For
            End If
        Next i
        If Not hasGrp Then GoTo CI_NEXTGRP

        With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI))
            .Interior.Color = RGB(68, 114, 196)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .VerticalAlignment = xlVAlignCenter
        End With
        wsOut.Cells(r, 1).value = grpLabel
        r = r + 1

        With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI))
            .Interior.Color = RGB(31, 73, 125)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlCenter
        End With
        wsOut.Cells(r, 1).value = "#"
        wsOut.Cells(r, 2).value = "Contract No"
        wsOut.Cells(r, 3).value = "Signing Date"
        wsOut.Cells(r, 4).value = "End Date"
        wsOut.Cells(r, 5).value = "Status"

        Application.DisplayAlerts = False
        wsOut.Range(wsOut.Cells(r, 6), wsOut.Cells(r, 7)).UnMerge
        wsOut.Range(wsOut.Cells(r, 6), wsOut.Cells(r, 7)).Merge
        wsOut.Cells(r, 6).value = "Company A"

        wsOut.Range(wsOut.Cells(r, 8), wsOut.Cells(r, 9)).UnMerge
        wsOut.Range(wsOut.Cells(r, 8), wsOut.Cells(r, 9)).Merge
        wsOut.Cells(r, 8).value = "Company B"

        wsOut.Range(wsOut.Cells(r, 10), wsOut.Cells(r, 11)).UnMerge
        wsOut.Range(wsOut.Cells(r, 10), wsOut.Cells(r, 11)).Merge
        wsOut.Cells(r, 10).value = "Scope of Work"
        Application.DisplayAlerts = True

        wsOut.Cells(r, 12).value = "Contract Value"
        ApplyDataRowBorders wsOut, r, NCOLS_MULTI, 1, 2, 3, 4, 5, 7, 9, 11
        r = r + 1

        Dim grpIDs() As String
        Dim grpIDCount As Long
        grpIDCount = 0
        ReDim grpIDs(0 To colCount - 1)
        Dim pi As Long
        For pi = 0 To colCount - 1
            If colInfo(pi, 3) = curGrp Then
                grpIDs(grpIDCount) = CStr(colInfo(pi, 0))
                grpIDCount = grpIDCount + 1
            End If
        Next pi

        Dim seqNo As Long
        Dim altRow As Boolean
        Dim grpValueSum As Double
        seqNo = 0: altRow = False: grpValueSum = 0

        Application.DisplayAlerts = False
        For i = 0 To colCount - 1
            If colInfo(i, 3) <> curGrp Then GoTo CI_NEXTROW

            Dim contractID As String
            contractID = CStr(colInfo(i, 0))
            Dim arrRow As Long
            arrRow = CLng(colInfo(i, 1))

            Dim sigDateV As Variant
            Dim endDateV As Variant
            sigDateV = ""
            endDateV = ""
            If IsValidDMY(gArrContract(arrRow, 4)) Then sigDateV = ParseDMY(gArrContract(arrRow, 4))
            If IsValidDMY(gArrContract(arrRow, 5)) Then endDateV = ParseDMY(gArrContract(arrRow, 5))

            Dim rowColor As Long
            If altRow Then
                rowColor = RGB(235, 241, 252)
            Else
                rowColor = RGB(255, 255, 255)
            End If
            altRow = Not altRow
            seqNo = seqNo + 1

            wsOut.Cells(r, 1).value = seqNo
            wsOut.Cells(r, 1).HorizontalAlignment = xlCenter
            wsOut.Cells(r, 1).VerticalAlignment = xlVAlignCenter

            wsOut.Cells(r, 2).NumberFormat = "@"
            wsOut.Cells(r, 2).value = Trim$(CStr(gArrContract(arrRow, 2)))
            wsOut.Cells(r, 2).WrapText = True
            wsOut.Cells(r, 2).VerticalAlignment = xlVAlignCenter

            With wsOut.Cells(r, 3)
                If sigDateV <> "" Then
                    .value = sigDateV
                    .NumberFormat = "dd/mm/yyyy"
                End If
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlVAlignCenter
            End With

            With wsOut.Cells(r, 4)
                If endDateV <> "" Then
                    .value = endDateV
                    .NumberFormat = "dd/mm/yyyy"
                End If
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlVAlignCenter
            End With

            wsOut.Cells(r, 5).value = Trim$(CStr(gArrContract(arrRow, 6)))
            wsOut.Cells(r, 5).HorizontalAlignment = xlCenter
            wsOut.Cells(r, 5).VerticalAlignment = xlVAlignCenter

            wsOut.Range(wsOut.Cells(r, 6), wsOut.Cells(r, 7)).UnMerge
            wsOut.Range(wsOut.Cells(r, 6), wsOut.Cells(r, 7)).Merge
            wsOut.Cells(r, 6).value = Trim$(CStr(gArrContract(arrRow, 7)))
            wsOut.Cells(r, 6).WrapText = True
            wsOut.Cells(r, 6).VerticalAlignment = xlVAlignCenter
            If mShCI_A > 0 Then
                wsOut.Cells(r, mShCI_A).value = wsOut.Cells(r, 6).value
                wsOut.Cells(r, mShCI_A).WrapText = True
            End If

            wsOut.Range(wsOut.Cells(r, 8), wsOut.Cells(r, 9)).UnMerge
            wsOut.Range(wsOut.Cells(r, 8), wsOut.Cells(r, 9)).Merge
            wsOut.Cells(r, 8).value = Trim$(CStr(gArrContract(arrRow, 8)))
            wsOut.Cells(r, 8).WrapText = True
            wsOut.Cells(r, 8).VerticalAlignment = xlVAlignCenter
            If mShCI_B > 0 Then
                wsOut.Cells(r, mShCI_B).value = wsOut.Cells(r, 8).value
                wsOut.Cells(r, mShCI_B).WrapText = True
            End If

            wsOut.Range(wsOut.Cells(r, 10), wsOut.Cells(r, 11)).UnMerge
            wsOut.Range(wsOut.Cells(r, 10), wsOut.Cells(r, 11)).Merge
            wsOut.Cells(r, 10).value = Trim$(CStr(gArrContract(arrRow, 11)))
            wsOut.Cells(r, 10).WrapText = True
            wsOut.Cells(r, 10).VerticalAlignment = xlVAlignCenter
            If mShCI_SOW > 0 Then
                wsOut.Cells(r, mShCI_SOW).value = wsOut.Cells(r, 10).value
                wsOut.Cells(r, mShCI_SOW).WrapText = True
            End If

            Dim valRaw As Variant
            valRaw = gArrContract(arrRow, 15)
            Dim valNum As Double
            valNum = 0
            With wsOut.Cells(r, 12)
                If IsNumeric(valRaw) Then
                    valNum = CDbl(valRaw)
                    .value = valNum
                    .NumberFormat = "#,##0"
                ElseIf IsNumeric(Replace(CStr(valRaw), ",", "")) Then
                    valNum = CDbl(Replace(CStr(valRaw), ",", ""))
                    .value = valNum
                    .NumberFormat = "#,##0"
                Else
                    .value = Trim$(CStr(valRaw))
                    .NumberFormat = "@"
                End If
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlVAlignCenter
            End With

            Dim statusVal As String
            statusVal = UCase$(Trim$(CStr(gArrContract(arrRow, 6))))

            Dim isLeaf As Boolean
            isLeaf = True
            Dim li As Long
            For li = 0 To grpIDCount - 1
                If grpIDs(li) <> contractID Then
                    If Left$(grpIDs(li), Len(contractID) + 1) = contractID & "-" Then
                        isLeaf = False
                        Exit For
                    End If
                End If
            Next li
            If isLeaf And statusVal <> "CANCEL" Then
                grpValueSum = grpValueSum + valNum
            End If

            wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI)).Interior.Color = rowColor

            If statusVal = "CANCEL" Then
                wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI)).Font.Color = RGB(192, 0, 0)
                wsOut.Cells(r, 5).Font.Bold = True
            ElseIf statusVal = "COMPLETED" Then
                wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI)).Font.Color = RGB(0, 112, 192)
            End If

            ApplyDataRowBorders wsOut, r, NCOLS_MULTI, 1, 2, 3, 4, 5, 7, 9, 11
            r = r + 1
CI_NEXTROW:
        Next i
        Application.DisplayAlerts = True

        With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI))
            .Interior.Color = RGB(31, 73, 125)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .VerticalAlignment = xlVAlignCenter
        End With
        wsOut.Cells(r, 1).value = "Total Group " & curGrp
        wsOut.Cells(r, 12).value = grpValueSum
        wsOut.Cells(r, 12).NumberFormat = "#,##0"
        wsOut.Cells(r, 12).HorizontalAlignment = xlRight
        r = r + 1

        If grpPass = 1 And totalB_cnt > 0 Then r = r + 1
CI_NEXTGRP:
    Next grpPass

    ExportContractInfoSection = r
End Function

'================================================================
' SECTION 2: DOCUMENT TRACKING
'================================================================
Public Function ExportDocTrackingSection( _
    ByRef wsOut As Worksheet, _
    ByVal startRow As Long, _
    contractList() As String, _
    ByVal cCount As Long, _
    ByVal filterMode As String, _
    ByVal filterMonth As Long, _
    ByVal filterYear As Long) As Long

    ' Layout moi: luon dung 12 cot (A:L) cho ca single va multi
    Const DT_NCOLS As Long = 12
    ' Col constants
    Const COL_STT       As Long = 1   ' A
    Const COL_DOCTYPE   As Long = 2   ' B
    Const COL_NUMBER    As Long = 3   ' C
    Const COL_VERSION   As Long = 4   ' D
    Const COL_DOCNAME   As Long = 5   ' E
    Const COL_DATE1     As Long = 6   ' F
    Const COL_DATE2     As Long = 7   ' G
    Const COL_SENDTO    As Long = 8   ' H
    Const COL_STATUS    As Long = 9   ' I
    Const COL_NOTE      As Long = 10  ' J (merge voi K=11)
    Const COL_LINK      As Long = 12  ' L

    Dim r As Long
    r = startRow

    Dim wsDoc As Worksheet
    On Error Resume Next
    Set wsDoc = GetDataSheet("DOCUMENT_TRACKER")
    On Error GoTo 0
    If wsDoc Is Nothing Then
        wsOut.Cells(r, 1).value = "[DOCUMENT_TRACKER sheet not found]"
        ExportDocTrackingSection = r + 2
        Exit Function
    End If

    Dim lastDocRow As Long
    lastDocRow = wsDoc.Cells(wsDoc.rows.count, "A").End(xlUp).Row
    If lastDocRow < 2 Then
        ExportDocTrackingSection = r + 2
        Exit Function
    End If

    Dim arrDoc As Variant
    arrDoc = wsDoc.Range(wsDoc.Cells(2, 1), wsDoc.Cells(lastDocRow, 16)).value

    ' -- Section title ----------------------------------------
    With wsOut.Cells(r, 1)
        .value = "II. DOCUMENT TRACKING"
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
    End With
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, DT_NCOLS)).Interior.Color = RGB(31, 73, 125)
    r = r + 1

    Dim filterNote As String
    Select Case filterMode
        Case "monthly": filterNote = "Period: Month " & filterMonth & "/" & filterYear
        Case "yearly":  filterNote = "Period: Year " & filterYear
        Case Else:      filterNote = "Period: All time"
    End Select
    With wsOut.Cells(r, 1)
        .value = filterNote
        .Font.Italic = True
        .Font.Color = RGB(89, 89, 89)
    End With
    r = r + 1

    ' -- Group counting ---------------------------------------
    Dim cidGrp() As String
    ReDim cidGrp(0 To cCount - 1)
    Dim cntA As Long, cntB As Long
    cntA = 0: cntB = 0
    Dim ci As Long
    For ci = 0 To cCount - 1
        Dim cidTemp As String
        cidTemp = contractList(ci)
        Dim grpTemp As String
        grpTemp = GetGrpFromID(cidTemp)
        cidGrp(ci) = grpTemp
        If grpTemp = "A" Then
            cntA = cntA + 1
        ElseIf grpTemp = "B" Then
            cntB = cntB + 1
        End If
    Next ci

    Dim grandTotal As Long
    grandTotal = 0

    ' -- Loop qua Group A roi Group B -------------------------
    Dim grpPass As Long
    For grpPass = 1 To 2
        Dim curGrp As String
        Dim grpContractCount As Long
        If grpPass = 1 Then
            curGrp = "A": grpContractCount = cntA
        Else
            curGrp = "B": grpContractCount = cntB
        End If
        If grpContractCount = 0 Then GoTo DT_NEXTGRP

        ' GROUP header row (placeholder, cap nhat sau)
        Dim grpHeaderRow As Long
        grpHeaderRow = r
        With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, DT_NCOLS))
            .Interior.Color = RGB(68, 114, 196)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .VerticalAlignment = xlVAlignCenter
        End With
        r = r + 1

        ' -- Column header row --------------------------------
        Dim date1Lbl As String, date2Lbl As String, sendToLbl As String
        If curGrp = "A" Then
            date1Lbl = "Processing Date"
            date2Lbl = "Completed Date"
            sendToLbl = "Send To"
        Else
            date1Lbl = "Received Date"
            date2Lbl = "Returned Date"
            sendToLbl = "Returned To"
        End If

        Dim hdrLabels(0 To 10) As String
        hdrLabels(0) = "#"
        hdrLabels(1) = "Document Type"
        hdrLabels(2) = "Number"
        hdrLabels(3) = "Version"
        hdrLabels(4) = "Document Name"
        hdrLabels(5) = date1Lbl
        hdrLabels(6) = date2Lbl
        hdrLabels(7) = sendToLbl
        hdrLabels(8) = "Status"
        hdrLabels(9) = "Note"        ' se merge J:K
        hdrLabels(10) = "Link"       ' cot L

        ' Mang anh xa cot Excel cho tung header label (bo qua K=11 vi merge vao J)
        Dim hdrCols(0 To 10) As Long
        hdrCols(0) = 1:  hdrCols(1) = 2:  hdrCols(2) = 3
        hdrCols(3) = 4:  hdrCols(4) = 5:  hdrCols(5) = 6
        hdrCols(6) = 7:  hdrCols(7) = 8:  hdrCols(8) = 9
        hdrCols(9) = 10: hdrCols(10) = 12

        Dim hIdx As Long
        Application.DisplayAlerts = False
        wsOut.Range(wsOut.Cells(r, COL_NOTE), wsOut.Cells(r, COL_NOTE + 1)).UnMerge
        wsOut.Range(wsOut.Cells(r, COL_NOTE), wsOut.Cells(r, COL_NOTE + 1)).Merge
        Application.DisplayAlerts = True

        For hIdx = 0 To 10
            With wsOut.Cells(r, hdrCols(hIdx))
                .value = hdrLabels(hIdx)
                .Font.Bold = True
                .Interior.Color = RGB(31, 73, 125)
                .Font.Color = RGB(255, 255, 255)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlVAlignCenter
            End With
        Next hIdx
        ' Note header (da merge J:K o tren)
        wsOut.Cells(r, COL_NOTE).HorizontalAlignment = xlCenter

        ApplyDataRowBorders wsOut, r, DT_NCOLS, 1, 2, 3, 4, 5, 6, 7, 8, 9, 11
        r = r + 1

        ' -- Loop qua tung contract ---------------------------
        Dim seqNo As Long
        Dim contractSeqNo As Long
        Dim grpDocCount As Long
        Dim contractColorToggle As Boolean
        seqNo = 0: contractSeqNo = 0: grpDocCount = 0
        contractColorToggle = False
        Dim contractsWithDoc As Long
        contractsWithDoc = 0

        Dim i As Long
        For i = 0 To cCount - 1
            If cidGrp(i) <> curGrp Then GoTo DT_NEXTCON

            Dim contractID As String
            contractID = contractList(i)
            Dim contractNo As String
            contractNo = contractID
            Dim companyB As String
            companyB = ""
            If Not gDictContract Is Nothing Then
                If gDictContract.Exists(contractID) Then
                    Dim rIdx As Long
                    rIdx = gDictContract(contractID)
                    contractNo = Trim$(CStr(gArrContract(rIdx, 2)))
                    companyB = Trim$(CStr(gArrContract(rIdx, 8)))
                End If
            End If

            ' Scan + filter documents
            Dim bufSize As Long
            bufSize = 64
            Dim docBuf() As Variant
            ReDim docBuf(0 To bufSize - 1, 0 To 2)
            Dim bufCount As Long
            bufCount = 0

            Dim j As Long
            Dim arrRows As Long
            arrRows = UBound(arrDoc, 1)
            For j = 1 To arrRows
                If Trim$(CStr(arrDoc(j, 2))) <> contractID Then GoTo DT_SCANNEXT
                If filterMode <> "all" Then
                    Dim recDV As Variant
                    recDV = arrDoc(j, 9)
                    If Not IsValidDMY(recDV) Then GoTo DT_SCANNEXT
                    Dim recDP As Date
                    recDP = ParseDMY(recDV)
                    If filterMode = "monthly" Then
                        If Month(recDP) <> filterMonth Or Year(recDP) <> filterYear Then GoTo DT_SCANNEXT
                    ElseIf filterMode = "yearly" Then
                        If Year(recDP) <> filterYear Then GoTo DT_SCANNEXT
                    End If
                End If
                If bufCount >= bufSize Then
                    bufSize = bufSize * 2
                    ReDim Preserve docBuf(0 To bufSize - 1, 0 To 2)
                End If
                Dim dtKey As Long
                dtKey = DocTypeSortKey(CStr(arrDoc(j, 3)))
                Dim recDbl As Double
                recDbl = 0
                If IsValidDMY(arrDoc(j, 9)) Then recDbl = CDbl(ParseDMY(arrDoc(j, 9)))
                docBuf(bufCount, 0) = j
                docBuf(bufCount, 1) = dtKey
                docBuf(bufCount, 2) = recDbl
                bufCount = bufCount + 1
DT_SCANNEXT:
            Next j

            If bufCount = 0 Then GoTo DT_NEXTCON

            ' Sort documents
            Dim sv1 As Long, sv2 As Long
            Dim tv0 As Variant, tv1 As Long, tv2 As Double
            For sv1 = 0 To bufCount - 2
                For sv2 = sv1 + 1 To bufCount - 1
                    Dim needSwap As Boolean
                    needSwap = False
                    If docBuf(sv2, 1) < docBuf(sv1, 1) Then
                        needSwap = True
                    ElseIf docBuf(sv2, 1) = docBuf(sv1, 1) Then
                        If docBuf(sv2, 2) < docBuf(sv1, 2) Then needSwap = True
                    End If
                    If needSwap Then
                        tv0 = docBuf(sv1, 0): tv1 = docBuf(sv1, 1): tv2 = docBuf(sv1, 2)
                        docBuf(sv1, 0) = docBuf(sv2, 0): docBuf(sv1, 1) = docBuf(sv2, 1): docBuf(sv1, 2) = docBuf(sv2, 2)
                        docBuf(sv2, 0) = tv0: docBuf(sv2, 1) = tv1: docBuf(sv2, 2) = tv2
                    End If
                Next sv2
            Next sv1

            contractsWithDoc = contractsWithDoc + 1
            contractColorToggle = Not contractColorToggle
            contractSeqNo = contractSeqNo + 1
            seqNo = 0

            ' -- CONTRACT HEADER ROW --------------------------
            Dim hdrBc As Long
            hdrBc = RGB(160, 160, 160)
            With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, DT_NCOLS))
                .Interior.Color = RGB(189, 215, 238)
                .Font.Bold = True
                .Font.Italic = False
                .Font.Color = RGB(31, 73, 125)
                .VerticalAlignment = xlVAlignCenter
            End With

            ' Col A: so thu tu contract
            With wsOut.Cells(r, 1)
                .value = contractSeqNo
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlVAlignCenter
            End With

            ' Col B:K merged: "ContractNo - CompanyB"
            Application.DisplayAlerts = False
            wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, 11)).UnMerge
            wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, 11)).Merge
            Application.DisplayAlerts = True
            With wsOut.Cells(r, 2)
                .NumberFormat = "@"
                If companyB <> "" Then
                    .value = contractNo & " - " & companyB
                Else
                    .value = contractNo
                End If
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlVAlignCenter
                .IndentLevel = 1
            End With

            ' Col L: empty
            wsOut.Cells(r, COL_LINK).value = ""

            ' Border mong xam dong nhat voi row data
            With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, DT_NCOLS))
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous: .Weight = xlThin: .Color = hdrBc
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous: .Weight = xlThin: .Color = hdrBc
                End With
                With .Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous: .Weight = xlThin: .Color = hdrBc
                End With
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous: .Weight = xlThin: .Color = hdrBc
                End With
            End With
            With wsOut.Cells(r, 1).Borders(xlEdgeRight)
                .LineStyle = xlContinuous: .Weight = xlThin: .Color = hdrBc
            End With
            r = r + 1
            ' -- KET THUC CONTRACT HEADER ROW -----------------

            ' -- DATA ROWS ------------------------------------
            Dim bi As Long
            For bi = 0 To bufCount - 1
                Dim aRow As Long
                aRow = CLng(docBuf(bi, 0))

                Dim nhanNgay As Variant
                Dim traNgay As Variant
                nhanNgay = ""
                traNgay = ""
                If IsValidDMY(arrDoc(aRow, 9)) Then nhanNgay = ParseDMY(arrDoc(aRow, 9))
                If IsValidDMY(arrDoc(aRow, 10)) Then traNgay = ParseDMY(arrDoc(aRow, 10))

                Dim hoSoName As String
                hoSoName = Trim$(CStr(arrDoc(aRow, 5)))
                If hoSoName = "" Then hoSoName = Trim$(CStr(arrDoc(aRow, 3)))

                Dim rowColor As Long
                rowColor = IIf(contractColorToggle, RGB(235, 241, 252), RGB(255, 255, 255))

                seqNo = seqNo + 1
                grpDocCount = grpDocCount + 1
                grandTotal = grandTotal + 1

                ' Col A: STT dang "1.1", "1.2"...
                wsOut.Cells(r, COL_STT).NumberFormat = "@"
                wsOut.Cells(r, COL_STT).value = contractSeqNo & "." & seqNo
                wsOut.Cells(r, COL_STT).HorizontalAlignment = xlCenter
                wsOut.Cells(r, COL_STT).VerticalAlignment = xlVAlignCenter

                ' Col B: Document Type
                wsOut.Cells(r, COL_DOCTYPE).value = Trim$(CStr(arrDoc(aRow, 3)))
                wsOut.Cells(r, COL_DOCTYPE).HorizontalAlignment = xlCenter
                wsOut.Cells(r, COL_DOCTYPE).VerticalAlignment = xlVAlignCenter

                ' Col C: Number
                wsOut.Cells(r, COL_NUMBER).value = Trim$(CStr(arrDoc(aRow, 4)))
                wsOut.Cells(r, COL_NUMBER).WrapText = True
                wsOut.Cells(r, COL_NUMBER).HorizontalAlignment = xlCenter
                wsOut.Cells(r, COL_NUMBER).VerticalAlignment = xlVAlignCenter

                ' Col D: Version
                wsOut.Cells(r, COL_VERSION).value = Trim$(CStr(arrDoc(aRow, 14)))
                wsOut.Cells(r, COL_VERSION).HorizontalAlignment = xlCenter
                wsOut.Cells(r, COL_VERSION).VerticalAlignment = xlVAlignCenter

                ' Col E: Document Name
                wsOut.Cells(r, COL_DOCNAME).value = hoSoName
                wsOut.Cells(r, COL_DOCNAME).WrapText = True
                wsOut.Cells(r, COL_DOCNAME).VerticalAlignment = xlVAlignCenter

                ' Col F: Date 1
                With wsOut.Cells(r, COL_DATE1)
                    If nhanNgay <> "" Then .value = nhanNgay: .NumberFormat = "dd/mm/yyyy"
                    .HorizontalAlignment = xlCenter: .VerticalAlignment = xlVAlignCenter
                End With

                ' Col G: Date 2
                With wsOut.Cells(r, COL_DATE2)
                    If traNgay <> "" Then .value = traNgay: .NumberFormat = "dd/mm/yyyy"
                    .HorizontalAlignment = xlCenter: .VerticalAlignment = xlVAlignCenter
                End With

                ' Col H: Returned To / Send To
                wsOut.Cells(r, COL_SENDTO).value = Trim$(CStr(arrDoc(aRow, 11)))
                wsOut.Cells(r, COL_SENDTO).WrapText = True
                wsOut.Cells(r, COL_SENDTO).VerticalAlignment = xlVAlignCenter

                ' Col I: Status
                wsOut.Cells(r, COL_STATUS).value = Trim$(CStr(arrDoc(aRow, 15)))
                wsOut.Cells(r, COL_STATUS).HorizontalAlignment = xlCenter
                wsOut.Cells(r, COL_STATUS).VerticalAlignment = xlVAlignCenter

                ' Col J:K merged — Note
                Application.DisplayAlerts = False
                wsOut.Range(wsOut.Cells(r, COL_NOTE), wsOut.Cells(r, COL_NOTE + 1)).UnMerge
                wsOut.Range(wsOut.Cells(r, COL_NOTE), wsOut.Cells(r, COL_NOTE + 1)).Merge
                Application.DisplayAlerts = True
                wsOut.Cells(r, COL_NOTE).value = Trim$(CStr(arrDoc(aRow, 16)))
                wsOut.Cells(r, COL_NOTE).WrapText = True
                wsOut.Cells(r, COL_NOTE).VerticalAlignment = xlVAlignCenter
                If mShDT_Note > 0 Then
                    wsOut.Cells(r, mShDT_Note).value = wsOut.Cells(r, COL_NOTE).value
                    wsOut.Cells(r, mShDT_Note).WrapText = True
                End If

                ' Col L: Link
                Dim lnkVal As String
                lnkVal = Trim$(CStr(arrDoc(aRow, 6)))
                If lnkVal <> "" Then
                    wsOut.Hyperlinks.Add Anchor:=wsOut.Cells(r, COL_LINK), _
                        Address:=lnkVal, TextToDisplay:="Open"
                    wsOut.Cells(r, COL_LINK).HorizontalAlignment = xlCenter
                    wsOut.Cells(r, COL_LINK).VerticalAlignment = xlVAlignCenter
                    wsOut.Cells(r, COL_LINK).Font.Color = RGB(0, 112, 192)
                    wsOut.Cells(r, COL_LINK).Font.Underline = xlUnderlineStyleSingle
                End If

                ApplyDataRowBorders wsOut, r, DT_NCOLS, 1, 2, 3, 4, 5, 6, 7, 8, 9, 11

                wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, DT_NCOLS)).Interior.Color = rowColor
                r = r + 1
            Next bi

            ' Ke duong dam cuoi contract
            With wsOut.Range(wsOut.Cells(r - 1, 1), wsOut.Cells(r - 1, DT_NCOLS)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .Color = RGB(31, 73, 125)
            End With
DT_NEXTCON:
        Next i

        ' Cap nhat GROUP header label
        Dim grpLbl As String
        grpLbl = "GROUP " & curGrp & "  (" & contractsWithDoc & " contract(s),  " & grpDocCount & " document(s))"
        wsOut.Cells(grpHeaderRow, 1).value = grpLbl
        wsOut.Cells(grpHeaderRow, 1).Font.Size = 10

        ' Total row
        With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, DT_NCOLS))
            .Interior.Color = RGB(31, 73, 125)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .VerticalAlignment = xlVAlignCenter
        End With
        wsOut.Cells(r, 1).value = "Total Group " & curGrp & ": " & grpDocCount & " document(s)"
        r = r + 1

        If grpPass = 1 And cntB > 0 Then r = r + 1
DT_NEXTGRP:
    Next grpPass

    ' Grand total
    With wsOut.Cells(r, 1)
        .value = "TOTAL: " & grandTotal & " document(s)"
        .Font.Bold = True
        .Font.Italic = True
        .Font.Color = RGB(31, 73, 125)
    End With
    r = r + 1

    ExportDocTrackingSection = r

End Function

'================================================================
' SECTION 3: PAYMENT REPORT
'================================================================
Public Function ExportPaymentSection( _
    ByRef wsOut As Worksheet, _
    ByVal startRow As Long, _
    contractList() As String, _
    ByVal cCount As Long, _
    ByVal filterMode As String, _
    ByVal filterMonth As Long, _
    ByVal filterYear As Long, _
    ByVal inclCancel As Boolean) As Long

    Dim r As Long
    r = startRow

    If Not gPaymentLoaded Then BuildPaymentCache
    If IsEmpty(gArrPayment) Then
        wsOut.Cells(r, 1).value = "[No payment data found]"
        ExportPaymentSection = r + 2
        Exit Function
    End If

    Dim dictFilter As Object
    Set dictFilter = CreateObject("Scripting.Dictionary")
    Dim ci As Long
    For ci = 0 To cCount - 1
        dictFilter(contractList(ci)) = 1
    Next ci

    '------------------------------------------------------------
    ' Tu dong nap B-CPB contracts vao dictFilter
    ' (B-CPB co the khong nam trong contractList truyen vao)
    '------------------------------------------------------------
    If Not gBCPBLoaded Then BuildPaymentCache
    If Not gDictBCPBRows Is Nothing Then
        Dim bcpbKey As Variant
        For Each bcpbKey In gDictBCPBRows.keys
            If Not dictFilter.Exists(bcpbKey) Then
                dictFilter(bcpbKey) = 1
            End If
        Next bcpbKey
    End If

    Dim showSummary As Boolean
    Dim showA As Boolean
    Dim showB As Boolean

    If cCount = 1 Then
        showSummary = False
        Dim singleGrp As String
        singleGrp = GetGrpFromID(contractList(0))
        showA = (singleGrp = "A")
        showB = (singleGrp = "B")
        If Not showA And Not showB Then
            showA = True
            showB = True
        End If
    Else
        showSummary = True
        showA = True
        showB = True
    End If

    Dim showContractCol As Boolean
    showContractCol = (cCount > 1)

    ' Layout:
    '   cCount=1 : 12 cot (A:L), paidColMain=10 (cot J)
    '   cCount>1 : 12 cot (A:L), paidColMain=10 (cot J)  [MOD v8]
    Dim payNCOLS As Long
    Dim paidColMain As Long
    If showContractCol Then
        payNCOLS = NCOLS_MULTI
        paidColMain = 10
    Else
        payNCOLS = NCOLS_MULTI
        paidColMain = 10
    End If

    Dim sectionTitle As String
    Select Case filterMode
        Case "monthly": sectionTitle = "III. PAYMENT REPORT — MONTH " & filterMonth & "/" & filterYear
        Case "yearly":  sectionTitle = "III. PAYMENT REPORT — YEAR " & filterYear
        Case Else:      sectionTitle = "III. PAYMENT REPORT (ALL TIME)"
    End Select

    With wsOut.Cells(r, 1)
        .value = sectionTitle
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
    End With
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, payNCOLS)).Interior.Color = RGB(31, 73, 125)
    r = r + 1

    Dim dictContractHasLedger As Object
    Set dictContractHasLedger = CreateObject("Scripting.Dictionary")
    Dim iScan As Long
    Dim cidScan As String
    For iScan = 1 To UBound(gArrPayment, 1)
        cidScan = Trim$(CStr(gArrPayment(iScan, 3)))
        If cidScan = "" Then GoTo SL_NEXT
        If Not dictFilter.Exists(cidScan) Then GoTo SL_NEXT
        If Trim$(CStr(gArrPayment(iScan, 2))) = "" Then GoTo SL_NEXT
        If Not dictContractHasLedger.Exists(cidScan) Then dictContractHasLedger(cidScan) = 1
SL_NEXT:
    Next iScan

    Dim detailArr() As Variant
    Dim detailCount As Long
    ReDim detailArr(0 To UBound(gArrPayment, 1) + 200, 0 To 9)
    detailCount = 0

    Dim dictA As Object
    Dim dictB As Object
    Set dictA = CreateObject("Scripting.Dictionary")
    Set dictB = CreateObject("Scripting.Dictionary")
    Dim dictContractHasTxn As Object
    Set dictContractHasTxn = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim contractID As String
    Dim contractNo As String
    Dim payDate As Variant
    Dim paidAmt As Double
    Dim originalAmt As Double
    Dim exRate As Double
    Dim grp As String
    Dim ym As String
    Dim rr As Long

    For i = 1 To UBound(gArrPayment, 1)
        contractID = Trim$(CStr(gArrPayment(i, 3)))
        If contractID = "" Then GoTo PY_NEXT
        payDate = gArrPayment(i, 17)
        If Not IsValidDMY(payDate) Then GoTo PY_NEXT
        paidAmt = val(CStr(gArrPayment(i, 13)))
        If paidAmt = 0 Then GoTo PY_NEXT
        If Not dictFilter.Exists(contractID) Then GoTo PY_NEXT
        If Not inclCancel Then
            If GetContractStatus(contractID) = "CANCEL" Then GoTo PY_NEXT
        End If
        ym = Format(ParseDMY(payDate), "yyyy-mm")
        If filterMode = "monthly" Then
            If CLng(Left$(ym, 4)) <> filterYear Or CLng(Right$(ym, 2)) <> filterMonth Then GoTo PY_NEXT
        ElseIf filterMode = "yearly" Then
            If CLng(Left$(ym, 4)) <> filterYear Then GoTo PY_NEXT
        End If
        originalAmt = paidAmt
        exRate = val(CStr(gArrPayment(i, 7)))
        If exRate > 1 Then paidAmt = originalAmt * exRate Else exRate = 0

        grp = GetGrpFromID(contractID)
        If grp = "" Then GoTo PY_NEXT

        If grp = "A" Then
            If Not dictA.Exists(ym) Then dictA(ym) = 0#
            dictA(ym) = dictA(ym) + paidAmt
        Else
            If Not dictB.Exists(ym) Then dictB(ym) = 0#
            dictB(ym) = dictB(ym) + paidAmt
        End If

        If Not dictContractHasTxn.Exists(contractID) Then dictContractHasTxn(contractID) = 1

        contractNo = contractID
        If Not gDictContract Is Nothing Then
            If gDictContract.Exists(contractID) Then
                rr = gDictContract(contractID)
                contractNo = Trim$(CStr(gArrContract(rr, 2)))
            End If
        End If

        detailArr(detailCount, 0) = ym
        detailArr(detailCount, 1) = grp
        detailArr(detailCount, 2) = contractID
        detailArr(detailCount, 3) = contractNo
        detailArr(detailCount, 4) = Trim$(CStr(gArrPayment(i, 15)))
        detailArr(detailCount, 5) = paidAmt
        detailArr(detailCount, 6) = ParseDMY(payDate)
        detailArr(detailCount, 7) = Trim$(CStr(gArrPayment(i, 16)))
        detailArr(detailCount, 8) = exRate
        detailArr(detailCount, 9) = originalAmt
        detailCount = detailCount + 1
PY_NEXT:
    Next i

    Dim dictPass2Done As Object
    Set dictPass2Done = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(gArrPayment, 1)
        contractID = Trim$(CStr(gArrPayment(i, 3)))
        If contractID = "" Then GoTo P2_NEXT
        If Not dictFilter.Exists(contractID) Then GoTo P2_NEXT
        If dictContractHasTxn.Exists(contractID) Then GoTo P2_NEXT
        If Not dictContractHasLedger.Exists(contractID) Then GoTo P2_NEXT
        If Trim$(CStr(gArrPayment(i, 2))) <> "" Then GoTo P2_NEXT
        If dictPass2Done.Exists(contractID) Then GoTo P2_NEXT
        paidAmt = val(CStr(gArrPayment(i, 13)))
        If paidAmt = 0 Then GoTo P2_NEXT
        If Not inclCancel Then
            If GetContractStatus(contractID) = "CANCEL" Then GoTo P2_NEXT
        End If
        grp = GetGrpFromID(contractID)
        If grp = "" Then GoTo P2_NEXT

        Dim hdrDate As Date
        Dim hdrDateFound As Boolean
        hdrDateFound = False
        Dim hh As Long
        For hh = 1 To UBound(gArrPayment, 1)
            If Trim$(CStr(gArrPayment(hh, 3))) = contractID Then
                If IsValidDMY(gArrPayment(hh, 17)) Then
                    hdrDate = ParseDMY(gArrPayment(hh, 17))
                    hdrDateFound = True
                    Exit For
                End If
            End If
        Next hh
        If Not hdrDateFound Then
            If filterMode = "monthly" Then
                hdrDate = DateSerial(filterYear, filterMonth, 1)
            ElseIf filterMode = "yearly" Then
                hdrDate = DateSerial(filterYear, 6, 1)
            Else
                hdrDate = DateSerial(Year(Now), 1, 1)
            End If
        End If

        ym = Format(hdrDate, "yyyy-mm")
        If filterMode = "monthly" Then
            If CLng(Left$(ym, 4)) <> filterYear Or CLng(Right$(ym, 2)) <> filterMonth Then GoTo P2_NEXT
        ElseIf filterMode = "yearly" Then
            If CLng(Left$(ym, 4)) <> filterYear Then GoTo P2_NEXT
        End If

        originalAmt = paidAmt
        exRate = val(CStr(gArrPayment(i, 7)))
        If exRate > 1 Then paidAmt = originalAmt * exRate Else exRate = 0

        If grp = "A" Then
            If Not dictA.Exists(ym) Then dictA(ym) = 0#
            dictA(ym) = dictA(ym) + paidAmt
        Else
            If Not dictB.Exists(ym) Then dictB(ym) = 0#
            dictB(ym) = dictB(ym) + paidAmt
        End If

        contractNo = contractID
        If Not gDictContract Is Nothing Then
            If gDictContract.Exists(contractID) Then
                rr = gDictContract(contractID)
                contractNo = Trim$(CStr(gArrContract(rr, 2)))
            End If
        End If

        detailArr(detailCount, 0) = ym
        detailArr(detailCount, 1) = grp
        detailArr(detailCount, 2) = contractID
        detailArr(detailCount, 3) = contractNo
        detailArr(detailCount, 4) = Trim$(CStr(gArrPayment(i, 15)))
        detailArr(detailCount, 5) = paidAmt
        detailArr(detailCount, 6) = hdrDate
        detailArr(detailCount, 7) = ""
        detailArr(detailCount, 8) = exRate
        detailArr(detailCount, 9) = originalAmt
        detailCount = detailCount + 1
        dictPass2Done(contractID) = 1
P2_NEXT:
    Next i

    If detailCount = 0 Then
        wsOut.Cells(r, 1).value = "[No payment records found for the selected period]"
        wsOut.Cells(r, 1).Font.Italic = True
        ExportPaymentSection = r + 2
        Exit Function
    End If

    Dim allMonths As Object
    Set allMonths = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In dictA.keys
        If Not allMonths.Exists(k) Then allMonths(k) = 1
    Next k
    For Each k In dictB.keys
        If Not allMonths.Exists(k) Then allMonths(k) = 1
    Next k

    Dim mCount As Long
    mCount = allMonths.count
    Dim monthArr() As String
    ReDim monthArr(0 To mCount - 1)
    Dim mi As Long
    mi = 0
    For Each k In allMonths.keys
        monthArr(mi) = CStr(k)
        mi = mi + 1
    Next k

    Dim m1 As Long, m2 As Long
    Dim tmpM As String
    For m1 = 0 To mCount - 2
        For m2 = m1 + 1 To mCount - 1
            If monthArr(m2) < monthArr(m1) Then
                tmpM = monthArr(m1)
                monthArr(m1) = monthArr(m2)
                monthArr(m2) = tmpM
            End If
        Next m2
    Next m1

    Dim d1 As Long, d2 As Long, dd As Long
    Dim tmpD(0 To 9) As Variant
    For d1 = 0 To detailCount - 2
        For d2 = d1 + 1 To detailCount - 1
            If CDbl(detailArr(d2, 6)) < CDbl(detailArr(d1, 6)) Then
                For dd = 0 To 9
                    tmpD(dd) = detailArr(d1, dd)
                    detailArr(d1, dd) = detailArr(d2, dd)
                    detailArr(d2, dd) = tmpD(dd)
                Next dd
            End If
        Next d2
    Next d1

    '================================================================
    ' SUMMARY BY MONTH (chi khi cCount > 1)
    '================================================================
    If showSummary Then
        With wsOut.Cells(r, 1)
            .value = "SUMMARY BY MONTH"
            .Font.Bold = True
            .Font.Size = 11
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 73, 125)
        End With
        wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI)).Interior.Color = RGB(31, 73, 125)
        r = r + 1

        With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI))
            .Interior.Color = RGB(31, 73, 125)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlCenter
        End With
        wsOut.Cells(r, 1).value = "#"
        Application.DisplayAlerts = False
        wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, 5)).UnMerge
        wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, 5)).Merge
        wsOut.Cells(r, 2).value = "Month"
        wsOut.Range(wsOut.Cells(r, 11), wsOut.Cells(r, 12)).UnMerge
        wsOut.Range(wsOut.Cells(r, 11), wsOut.Cells(r, 12)).Merge
        wsOut.Cells(r, 11).value = "Note"
        Application.DisplayAlerts = True
        wsOut.Cells(r, 6).value = "Group-A Paid"
        wsOut.Cells(r, 7).value = "Group-B Paid"
        wsOut.Cells(r, 8).value = "Net (A - B)"
        wsOut.Cells(r, 9).value = "Cumulative A"
        wsOut.Cells(r, 10).value = "Cumulative B"
        ApplyDataRowBorders wsOut, r, NCOLS_MULTI, 1, 5, 6, 7, 8, 9, 10
        r = r + 1

        Dim cumA As Double, cumB As Double
        cumA = 0: cumB = 0
        Dim mPaidA As Double, mPaidB As Double, mKey As String
        Dim yr As Long, mo As Long
        Dim lastYr As Long, yrSumA As Double, yrSumB As Double, yearSeq As Long
        lastYr = 0: yrSumA = 0: yrSumB = 0: yearSeq = 0
        Dim c As Long

        Application.DisplayAlerts = False
        For m1 = 0 To mCount - 1
            mKey = monthArr(m1)
            mPaidA = 0: mPaidB = 0
            If dictA.Exists(mKey) Then mPaidA = dictA(mKey)
            If dictB.Exists(mKey) Then mPaidB = dictB(mKey)
            yr = CLng(Left$(mKey, 4))
            mo = CLng(Right$(mKey, 2))

            If yr <> lastYr Then
                If lastYr <> 0 Then
                    With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI))
                        .Interior.Color = RGB(198, 224, 180)
                        .Font.Bold = True
                        .VerticalAlignment = xlVAlignCenter
                    End With
                    wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, 5)).UnMerge
                    wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, 5)).Merge
                    wsOut.Cells(r, 2).NumberFormat = "@"
                    wsOut.Cells(r, 2).value = "Total " & lastYr
                    wsOut.Cells(r, 2).Font.Italic = True
                    wsOut.Cells(r, 2).HorizontalAlignment = xlCenter
                    wsOut.Cells(r, 6).value = yrSumA
                    wsOut.Cells(r, 7).value = yrSumB
                    wsOut.Cells(r, 8).value = yrSumA - yrSumB
                    For c = 6 To 10: wsOut.Cells(r, c).NumberFormat = "#,##0": Next c
                    ApplyDataRowBorders wsOut, r, NCOLS_MULTI, 1, 5, 6, 7, 8, 9, 10
                    r = r + 1: yrSumA = 0: yrSumB = 0: yearSeq = 0
                End If
                With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI))
                    .Interior.Color = RGB(68, 114, 196)
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .VerticalAlignment = xlVAlignCenter
                End With
                wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, 5)).UnMerge
                wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, 5)).Merge
                wsOut.Cells(r, 2).NumberFormat = "@"
                wsOut.Cells(r, 2).value = CStr(yr)
                wsOut.Cells(r, 2).HorizontalAlignment = xlLeft
                r = r + 1
            End If

            cumA = cumA + mPaidA: cumB = cumB + mPaidB
            yrSumA = yrSumA + mPaidA: yrSumB = yrSumB + mPaidB
            lastYr = yr: yearSeq = yearSeq + 1

            wsOut.Cells(r, 1).value = yearSeq
            wsOut.Cells(r, 1).HorizontalAlignment = xlCenter
            wsOut.Cells(r, 1).VerticalAlignment = xlVAlignCenter
            wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, 5)).UnMerge
            wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, 5)).Merge
            wsOut.Cells(r, 2).NumberFormat = "@"
            wsOut.Cells(r, 2).value = Format(DateSerial(yr, mo, 1), "mmm-yyyy")
            wsOut.Cells(r, 2).HorizontalAlignment = xlCenter
            wsOut.Cells(r, 2).VerticalAlignment = xlVAlignCenter
            wsOut.Cells(r, 6).value = mPaidA
            wsOut.Cells(r, 7).value = mPaidB
            wsOut.Cells(r, 8).value = mPaidA - mPaidB
            wsOut.Cells(r, 9).value = cumA
            wsOut.Cells(r, 10).value = cumB
            For c = 6 To 10
                wsOut.Cells(r, c).NumberFormat = "#,##0"
                wsOut.Cells(r, c).VerticalAlignment = xlVAlignCenter
            Next c
            wsOut.Range(wsOut.Cells(r, 11), wsOut.Cells(r, 12)).UnMerge
            wsOut.Range(wsOut.Cells(r, 11), wsOut.Cells(r, 12)).Merge
            If (m1 Mod 2) = 0 Then
                wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI)).Interior.Color = RGB(235, 241, 252)
            End If
            ApplyDataRowBorders wsOut, r, NCOLS_MULTI, 1, 5, 6, 7, 8, 9, 10
            r = r + 1
        Next m1
        Application.DisplayAlerts = True

        If lastYr <> 0 Then
            With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI))
                .Interior.Color = RGB(198, 224, 180)
                .Font.Bold = True
                .VerticalAlignment = xlVAlignCenter
            End With
            Application.DisplayAlerts = False
            wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, 5)).UnMerge
            wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, 5)).Merge
            Application.DisplayAlerts = True
            wsOut.Cells(r, 2).NumberFormat = "@"
            wsOut.Cells(r, 2).value = "Total " & lastYr
            wsOut.Cells(r, 2).Font.Italic = True
            wsOut.Cells(r, 2).HorizontalAlignment = xlCenter
            wsOut.Cells(r, 6).value = yrSumA
            wsOut.Cells(r, 7).value = yrSumB
            wsOut.Cells(r, 8).value = yrSumA - yrSumB
            For c = 6 To 10: wsOut.Cells(r, c).NumberFormat = "#,##0": Next c
            ApplyDataRowBorders wsOut, r, NCOLS_MULTI, 1, 5, 6, 7, 8, 9, 10
            r = r + 1
        End If

        With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS_MULTI))
            .Interior.Color = RGB(31, 73, 125)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .VerticalAlignment = xlVAlignCenter
        End With
        Application.DisplayAlerts = False
        wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, 5)).UnMerge
        wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, 5)).Merge
        Application.DisplayAlerts = True
        wsOut.Cells(r, 2).value = "TOTAL"
        wsOut.Cells(r, 2).HorizontalAlignment = xlCenter
        wsOut.Cells(r, 6).value = cumA
        wsOut.Cells(r, 7).value = cumB
        wsOut.Cells(r, 8).value = cumA - cumB
        wsOut.Cells(r, 9).value = cumA
        wsOut.Cells(r, 10).value = cumB
        For c = 6 To 10: wsOut.Cells(r, c).NumberFormat = "#,##0": Next c
        r = r + 2
    End If

    '================================================================
    ' PAYMENT DETAIL — GROUP A
    '================================================================
    Dim totalA As Double
    Dim totalB As Double

    If showA Then
        Dim hasA As Boolean
        hasA = False
        Dim chkA As Long
        For chkA = 0 To detailCount - 1
            If CStr(detailArr(chkA, 1)) = "A" Then
                hasA = True
                Exit For
            End If
        Next chkA

        If hasA Then
            With wsOut.Cells(r, 1)
                .value = "GROUP-A — PAYMENT DETAIL"
                .Font.Bold = True
                .Font.Size = 11
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(31, 73, 125)
            End With
            wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, payNCOLS)).Interior.Color = RGB(31, 73, 125)
            r = r + 1
            r = WritePayHeaderRow(wsOut, r, showContractCol, payNCOLS)
            r = WriteDetailGroup(wsOut, r, detailArr, detailCount, "A", totalA, showContractCol, payNCOLS, paidColMain)

            With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, payNCOLS))
                .Interior.Color = RGB(31, 73, 125)
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
                .VerticalAlignment = xlVAlignCenter
            End With
            ' Merge B:(paidColMain-1) cho nhan TOTAL A, gia tri tai paidColMain
            Application.DisplayAlerts = False
            wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, paidColMain - 1)).UnMerge
            wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, paidColMain - 1)).Merge
            Application.DisplayAlerts = True
            With wsOut.Cells(r, 2)
                .value = "TOTAL A"
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlVAlignCenter
            End With
            wsOut.Cells(r, paidColMain).value = totalA
            wsOut.Cells(r, paidColMain).NumberFormat = "#,##0"
            wsOut.Cells(r, paidColMain).HorizontalAlignment = xlRight
            r = r + 2
        End If
    End If

    '================================================================
    ' PAYMENT DETAIL — GROUP B
    '================================================================
    If showB Then
        Dim hasB As Boolean
        hasB = False
        Dim chkB As Long
        For chkB = 0 To detailCount - 1
            If CStr(detailArr(chkB, 1)) = "B" Then
                hasB = True
                Exit For
            End If
        Next chkB

        If hasB Then
            With wsOut.Cells(r, 1)
                .value = "GROUP-B — PAYMENT DETAIL"
                .Font.Bold = True
                .Font.Size = 11
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(31, 73, 125)
            End With
            wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, payNCOLS)).Interior.Color = RGB(31, 73, 125)
            r = r + 1
            r = WritePayHeaderRow(wsOut, r, showContractCol, payNCOLS)
            r = WriteDetailGroup(wsOut, r, detailArr, detailCount, "B", totalB, showContractCol, payNCOLS, paidColMain)

            With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, payNCOLS))
                .Interior.Color = RGB(31, 73, 125)
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
                .VerticalAlignment = xlVAlignCenter
            End With
            ' Merge B:(paidColMain-1) cho nhan TOTAL B, gia tri tai paidColMain
            Application.DisplayAlerts = False
            wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, paidColMain - 1)).UnMerge
            wsOut.Range(wsOut.Cells(r, 2), wsOut.Cells(r, paidColMain - 1)).Merge
            Application.DisplayAlerts = True
            With wsOut.Cells(r, 2)
                .value = "TOTAL B"
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlVAlignCenter
            End With
            wsOut.Cells(r, paidColMain).value = totalB
            wsOut.Cells(r, paidColMain).NumberFormat = "#,##0"
            wsOut.Cells(r, paidColMain).HorizontalAlignment = xlRight
            r = r + 1
        End If
    End If

    ExportPaymentSection = r
End Function

'================================================================
' HELPER: Dong header bang Payment Detail
'
' showContractCol=True (cCount > 1) — A:L 12 cot:
'   A=# | B=Date | C:E=Contract No(merge 3) | F:H=Description(merge 3) |
'   I=Ref No | J=Paid Amount(VND) | K=Exchange Rate | L=Foreign Amount
'   Borders: sau A(1) B(2) E(5) H(8) I(9) J(10) K(11)
'
' showContractCol=False (cCount = 1) — A:L 12 cot:
'   A=# | B=Date | C:H=Description(merge 6) |
'   I=Ref No | J=Paid Amount(VND) | K=Exchange Rate | L=Foreign Amount
'   Borders: sau A(1) B(2) H(8) I(9) J(10) K(11)
'================================================================
Private Function WritePayHeaderRow( _
    ByRef wsOut As Worksheet, _
    ByVal r As Long, _
    ByVal showContractCol As Boolean, _
    ByVal payNCOLS As Long) As Long

    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, payNCOLS)).ClearContents
    With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, payNCOLS))
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlCenter
    End With
    wsOut.Cells(r, 1).value = "#"
    wsOut.Cells(r, 2).value = "Date"

    Application.DisplayAlerts = False

    If showContractCol Then
        ' C:E = Contract No (merge 3)
        wsOut.Range(wsOut.Cells(r, 3), wsOut.Cells(r, 5)).UnMerge
        wsOut.Range(wsOut.Cells(r, 3), wsOut.Cells(r, 5)).Merge
        wsOut.Cells(r, 3).value = "Contract No"

        ' F:H = Description (merge 3)
        wsOut.Range(wsOut.Cells(r, 6), wsOut.Cells(r, 8)).UnMerge
        wsOut.Range(wsOut.Cells(r, 6), wsOut.Cells(r, 8)).Merge
        wsOut.Cells(r, 6).value = "Description"

        wsOut.Cells(r, 9).value = "Ref No"
        wsOut.Cells(r, 10).value = "Paid Amount (VND)"
        wsOut.Cells(r, 11).value = "Exchange Rate"
        wsOut.Cells(r, 12).value = "Foreign Amount"

        ApplyDataRowBorders wsOut, r, payNCOLS, 1, 2, 5, 8, 9, 10, 11
    Else
        ' C:H = Description (merge 6)
        wsOut.Range(wsOut.Cells(r, 3), wsOut.Cells(r, 8)).UnMerge
        wsOut.Range(wsOut.Cells(r, 3), wsOut.Cells(r, 8)).Merge
        wsOut.Cells(r, 3).value = "Description"

        wsOut.Cells(r, 9).value = "Ref No"
        wsOut.Cells(r, 10).value = "Paid Amount (VND)"
        wsOut.Cells(r, 11).value = "Exchange Rate"
        wsOut.Cells(r, 12).value = "Foreign Amount"

        ApplyDataRowBorders wsOut, r, payNCOLS, 1, 2, 8, 9, 10, 11
    End If

    Application.DisplayAlerts = True
    WritePayHeaderRow = r + 1
End Function

'================================================================
' HELPER: Viet chi tiet thanh toan cho 1 nhom
'
' showContractCol=True:
'   C:E=Contract No | F:H=Description | I=RefNo | J=PaidAmt(paidCol=10) |
'   K=ExRate | L=ForeignAmt
'   Dong thang header : merge B:L
'   Dong tong thang   : merge B:I, gia tri tai paidCol (J)
'================================================================
Private Function WriteDetailGroup( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByRef detailArr() As Variant, _
    ByVal detailCount As Long, _
    ByVal grpFilter As String, _
    ByRef groupTotal As Double, _
    ByVal showContractCol As Boolean, _
    ByVal NCOLS As Long, _
    ByVal paidCol As Long) As Long

    Dim i As Long
    Dim r As Long
    Dim lastMonthShown As String
    Dim monthTotal As Double
    Dim altRow As Long
    Dim seqNo As Long
    Dim dYM As String
    Dim yr As Long, mo As Long
    Dim prevYr As Long, prevMo As Long

    r = startRow
    lastMonthShown = ""
    monthTotal = 0
    groupTotal = 0
    altRow = 0
    seqNo = 0

    Application.DisplayAlerts = False

    For i = 0 To detailCount - 1
        If CStr(detailArr(i, 1)) <> grpFilter Then GoTo DG_NEXT
        dYM = CStr(detailArr(i, 0))

        ' --- Dong tong thang cu ---
        If dYM <> lastMonthShown And lastMonthShown <> "" Then
            prevYr = CLng(Left$(lastMonthShown, 4))
            prevMo = CLng(Right$(lastMonthShown, 2))
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, NCOLS))
                .Interior.Color = RGB(198, 224, 180)
                .Font.Bold = True
                .VerticalAlignment = xlVAlignCenter
            End With
            ' Merge B:(paidCol-1) cho nhan tong thang
            ws.Range(ws.Cells(r, 2), ws.Cells(r, paidCol - 1)).UnMerge
            ws.Range(ws.Cells(r, 2), ws.Cells(r, paidCol - 1)).Merge
            ws.Cells(r, 2).NumberFormat = "@"
            ws.Cells(r, 2).value = "Total " & Format(DateSerial(prevYr, prevMo, 1), "mmm-yyyy")
            ws.Cells(r, 2).Font.Italic = True
            ws.Cells(r, 2).HorizontalAlignment = xlCenter
            ws.Cells(r, 2).VerticalAlignment = xlVAlignCenter
            ws.Cells(r, paidCol).value = monthTotal
            ws.Cells(r, paidCol).NumberFormat = "#,##0"
            ws.Cells(r, paidCol).Font.Bold = True
            ws.Cells(r, paidCol).HorizontalAlignment = xlRight
            r = r + 1
            monthTotal = 0
            altRow = 0
            seqNo = 0
        End If

        ' --- Dong header thang moi ---
        If dYM <> lastMonthShown Then
            yr = CLng(Left$(dYM, 4))
            mo = CLng(Right$(dYM, 2))
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, NCOLS))
                .Interior.Color = RGB(68, 114, 196)
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .VerticalAlignment = xlVAlignCenter
            End With
            ' Merge B:nCols cho ten thang (toan bo chieu rong)
            ws.Range(ws.Cells(r, 2), ws.Cells(r, NCOLS)).UnMerge
            ws.Range(ws.Cells(r, 2), ws.Cells(r, NCOLS)).Merge
            ws.Cells(r, 2).NumberFormat = "@"
            ws.Cells(r, 2).value = Format(DateSerial(yr, mo, 1), "mmmm yyyy")
            ws.Cells(r, 2).HorizontalAlignment = xlLeft
            ws.Cells(r, 2).VerticalAlignment = xlVAlignCenter
            r = r + 1
            lastMonthShown = dYM
        End If

        ' --- Dong du lieu ---
        seqNo = seqNo + 1
        ws.Cells(r, 1).value = seqNo
        ws.Cells(r, 1).HorizontalAlignment = xlCenter
        ws.Cells(r, 1).VerticalAlignment = xlVAlignCenter

        ws.Cells(r, 2).value = detailArr(i, 6)
        ws.Cells(r, 2).NumberFormat = "dd/mm/yyyy"
        ws.Cells(r, 2).HorizontalAlignment = xlCenter
        ws.Cells(r, 2).VerticalAlignment = xlVAlignCenter

        If showContractCol Then
            ' C:E = Contract No (merge 3)
            ws.Range(ws.Cells(r, 3), ws.Cells(r, 5)).UnMerge
            ws.Range(ws.Cells(r, 3), ws.Cells(r, 5)).Merge
            With ws.Cells(r, 3)
                .NumberFormat = "@"
                .value = CStr(detailArr(i, 3))
                .WrapText = True
                .VerticalAlignment = xlVAlignCenter
                .HorizontalAlignment = xlLeft
            End With
            If mShPay_CNo > 0 Then
                ws.Cells(r, mShPay_CNo).value = CStr(detailArr(i, 3))
                ws.Cells(r, mShPay_CNo).WrapText = True
            End If

            ' F:H = Description (merge 3)
            ws.Range(ws.Cells(r, 6), ws.Cells(r, 8)).UnMerge
            ws.Range(ws.Cells(r, 6), ws.Cells(r, 8)).Merge
            With ws.Cells(r, 6)
                .NumberFormat = "@"
                .value = detailArr(i, 4)
                .WrapText = True
                .VerticalAlignment = xlVAlignCenter
                .HorizontalAlignment = xlLeft
            End With
            If mShPay_Dsc > 0 Then
                ws.Cells(r, mShPay_Dsc).value = CStr(detailArr(i, 4))
                ws.Cells(r, mShPay_Dsc).WrapText = True
            End If

            ' I = Ref No
            ws.Cells(r, 9).value = detailArr(i, 7)
            ws.Cells(r, 9).HorizontalAlignment = xlCenter
            ws.Cells(r, 9).VerticalAlignment = xlVAlignCenter

            ' J = Paid Amount (paidCol = 10)
            ws.Cells(r, paidCol).value = detailArr(i, 5)
            ws.Cells(r, paidCol).NumberFormat = "#,##0"
            ws.Cells(r, paidCol).HorizontalAlignment = xlRight
            ws.Cells(r, paidCol).VerticalAlignment = xlVAlignCenter

            ' K = Exchange Rate, L = Foreign Amount
            If CDbl(detailArr(i, 8)) > 1 Then
                ws.Cells(r, 11).value = detailArr(i, 8)
                ws.Cells(r, 11).NumberFormat = "#,##0.000"
                ws.Cells(r, 11).HorizontalAlignment = xlRight
                ws.Cells(r, 11).VerticalAlignment = xlVAlignCenter
                ws.Cells(r, 12).value = detailArr(i, 9)
                ws.Cells(r, 12).NumberFormat = "#,##0.##"
                ws.Cells(r, 12).HorizontalAlignment = xlRight
                ws.Cells(r, 12).VerticalAlignment = xlVAlignCenter
            Else
                ws.Cells(r, 11).value = ""
                ws.Cells(r, 12).value = ""
            End If

            ' Borders: sau A(1) B(2) E(5) H(8) I(9) J(10) K(11)
            ApplyDataRowBorders ws, r, NCOLS, 1, 2, 5, 8, 9, 10, 11
        Else
            ' C:H = Description (merge 6)
            ws.Range(ws.Cells(r, 3), ws.Cells(r, 8)).UnMerge
            ws.Range(ws.Cells(r, 3), ws.Cells(r, 8)).Merge
            With ws.Cells(r, 3)
                .NumberFormat = "@"
                .value = detailArr(i, 4)
                .WrapText = True
                .VerticalAlignment = xlVAlignCenter
                .HorizontalAlignment = xlLeft
            End With
            If mShPay_Dsc > 0 Then
                ws.Cells(r, mShPay_Dsc).value = CStr(detailArr(i, 4))
                ws.Cells(r, mShPay_Dsc).WrapText = True
            End If
            ws.Cells(r, 9).value = detailArr(i, 7)
            ws.Cells(r, 9).HorizontalAlignment = xlCenter
            ws.Cells(r, 9).VerticalAlignment = xlVAlignCenter
            ws.Cells(r, paidCol).value = detailArr(i, 5)
            ws.Cells(r, paidCol).NumberFormat = "#,##0"
            ws.Cells(r, paidCol).HorizontalAlignment = xlRight
            ws.Cells(r, paidCol).VerticalAlignment = xlVAlignCenter
            If CDbl(detailArr(i, 8)) > 1 Then
                ws.Cells(r, 11).value = detailArr(i, 8)
                ws.Cells(r, 11).NumberFormat = "#,##0.000"
                ws.Cells(r, 11).HorizontalAlignment = xlRight
                ws.Cells(r, 11).VerticalAlignment = xlVAlignCenter
                ws.Cells(r, 12).value = detailArr(i, 9)
                ws.Cells(r, 12).NumberFormat = "#,##0.##"
                ws.Cells(r, 12).HorizontalAlignment = xlRight
                ws.Cells(r, 12).VerticalAlignment = xlVAlignCenter
            Else
                ws.Cells(r, 11).value = ""
                ws.Cells(r, 12).value = ""
            End If
            ApplyDataRowBorders ws, r, NCOLS, 1, 2, 8, 9, 10, 11
        End If

        If (altRow Mod 2) = 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, NCOLS)).Interior.Color = RGB(235, 241, 252)
        End If

        monthTotal = monthTotal + CDbl(detailArr(i, 5))
        groupTotal = groupTotal + CDbl(detailArr(i, 5))
        altRow = altRow + 1
        r = r + 1
DG_NEXT:
    Next i

    ' --- Dong tong thang cuoi cung ---
    If lastMonthShown <> "" And monthTotal > 0 Then
        prevYr = CLng(Left$(lastMonthShown, 4))
        prevMo = CLng(Right$(lastMonthShown, 2))
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, NCOLS))
            .Interior.Color = RGB(198, 224, 180)
            .Font.Bold = True
            .VerticalAlignment = xlVAlignCenter
        End With
        ws.Range(ws.Cells(r, 2), ws.Cells(r, paidCol - 1)).UnMerge
        ws.Range(ws.Cells(r, 2), ws.Cells(r, paidCol - 1)).Merge
        ws.Cells(r, 2).NumberFormat = "@"
        ws.Cells(r, 2).value = "Total " & Format(DateSerial(prevYr, prevMo, 1), "mmm-yyyy")
        ws.Cells(r, 2).Font.Italic = True
        ws.Cells(r, 2).HorizontalAlignment = xlCenter
        ws.Cells(r, 2).VerticalAlignment = xlVAlignCenter
        ws.Cells(r, paidCol).value = monthTotal
        ws.Cells(r, paidCol).NumberFormat = "#,##0"
        ws.Cells(r, paidCol).Font.Bold = True
        ws.Cells(r, paidCol).HorizontalAlignment = xlRight
        r = r + 1
    End If

    Application.DisplayAlerts = True
    WriteDetailGroup = r
End Function





'================================================================
' HELPER: AutoFit chieu cao dong cho o merge ngang (WrapText)
' [v3 - NHANH NHAT]
'
' Cai tien so voi v2:
'   1. Dung SpecialCells(-4143) -> 1 COM call
'      lay tat ca merged ranges, KHONG duyet tung o.
'   2. Ghi HET cac o tam 1 luot.
'   3. Chi goi EntireRow.AutoFit DUY NHAT 1 LAN cho ca sheet
'      (v2 goi N lan, 1 lan moi dong co merge -> rat cham).
'   4. Xoa o tam bang Union range -> it COM call hon.
'================================================================
'================================================================
' SetupShadowCols: Dat chieu rong cot tam va khoi tao bien shadow
' Goi TRUOC khi xuat bieu. Cot tam bat dau tu col 13 (M).
'================================================================
Public Sub SetupShadowCols(ByRef ws As Worksheet, ByVal cCount As Long)
    Const w As Double = 18
    ws.Columns.ColumnWidth = 18
    ws.Columns(1).ColumnWidth = 5
    If cCount > 1 Then
        mShCI_A = 13: ws.Columns(13).ColumnWidth = 2 * w
        mShCI_B = 14: ws.Columns(14).ColumnWidth = 2 * w
        mShCI_SOW = 15: ws.Columns(15).ColumnWidth = 2 * w
        mShDT_Note = 16: ws.Columns(16).ColumnWidth = 2 * w
        mShPay_CNo = 17: ws.Columns(17).ColumnWidth = 3 * w
        mShPay_Dsc = 18: ws.Columns(18).ColumnWidth = 3 * w
        mShCI_Desc = 0: mShLastCol = 18
    Else
        mShCI_Desc = 13: ws.Columns(13).ColumnWidth = 8 * w
        mShDT_Note = 14: ws.Columns(14).ColumnWidth = 2 * w
        mShPay_Dsc = 15: ws.Columns(15).ColumnWidth = 6 * w
        mShCI_A = 0: mShCI_B = 0: mShCI_SOW = 0: mShPay_CNo = 0
        mShLastCol = 15
    End If

End Sub

'================================================================
' ClearShadowCols: Xoa cac cot tam sau khi AutoFit xong
'================================================================
Public Sub ClearShadowCols(ByRef ws As Worksheet)

    Dim r As Long
    Dim lastRow As Long
    Dim RowHeights() As Double

    If mShLastCol < 13 Then Exit Sub

    'Luu chi?u cao dòng hi?n t?i
    lastRow = ws.UsedRange.rows(ws.UsedRange.rows.count).Row
    ReDim RowHeights(1 To lastRow)

    For r = 1 To lastRow
        RowHeights(r) = ws.rows(r).RowHeight
    Next r

    'Xóa các c?t shadow
    ws.Range(ws.Columns(13), ws.Columns(mShLastCol)).Clear
    ws.Range(ws.Columns(13), ws.Columns(mShLastCol)).ColumnWidth = 18

    'Khôi ph?c chi?u cao dòng
    For r = 1 To lastRow
        ws.rows(r).RowHeight = RowHeights(r)
    Next r

    mShCI_A = 0
    mShCI_B = 0
    mShCI_SOW = 0
    mShCI_Desc = 0
    mShDT_Note = 0
    mShPay_CNo = 0
    mShPay_Dsc = 0
    mShLastCol = 0

End Sub


Public Sub ExportConsolidatedSummary( _
    ByRef wsOut As Worksheet, _
    contractList() As String, _
    ByVal cCount As Long)

    Const NCOLS     As Long = 13   ' A:M
    Const SH_BENB   As Long = 14   ' shadow col N (cho F:G merge)
    Const SH_HANGMUC As Long = 15  ' shadow col O (cho H:I merge)

    ' Cot du lieu trong gArrContract — chinh o day neu can
    Const COL_CONTRACTNO   As Long = 2
    Const COL_SIGNDATE     As Long = 4
    Const COL_STATUS       As Long = 6
    Const COL_PERSON       As Long = 22  ' Nguoi phu trach chinh
    Const COL_COMPANYB     As Long = 8
    Const COL_SCOPE        As Long = 11
    Const COL_VALUE        As Long = 15
    'Const COL_IPC_STATUS    As Long = 15
    'Const COL_NOTE_CONTRACT As Long = 22 ' Ghi chu hop dong

    If gDictContract Is Nothing Then Exit Sub
    If IsEmpty(gArrContract) Then Exit Sub
    If Not gPaymentLoaded Then BuildPaymentCache

    Application.DisplayAlerts = False

    Dim r As Long
    r = 1

    '================================================================
    ' 1. CHUOI TIENG VIET (dung ChrW de dam bao Unicode chinh xac)
    '================================================================
    ' Tieu de chinh: BANG TONG HOP THEO DOI
    Dim sTitle As String
    sTitle = "B" & ChrW(7842) & "NG T" & ChrW(7892) & "NG H" & ChrW(7906) & _
             "P"

    ' Ten cot header
    Dim sNo      As String: sNo = "STT"
    Dim sHDS     As String  ' Hop dong so
    Dim sNgKy    As String  ' Ngay ky
    Dim sTT      As String  ' Tinh trang HD
    Dim sNPT     As String  ' Nguoi phu trach chinh
    Dim sBenB    As String  ' Ben B
    Dim sHangMuc As String  ' Hang muc
    Dim sGTHD    As String  ' Gia tri HD
    Dim sDTT     As String  ' Da thanh toan
    Dim sHSTT    As String  ' Ho so thanh toan da duoc xac nhan
    Dim sGC      As String  ' Ghi chu

    sHDS = "H" & ChrW(7907) & "p " & ChrW(273) & ChrW(7891) & "ng s" & ChrW(7889)
    sNgKy = "Ng" & ChrW(224) & "y k" & ChrW(253)
    sTT = "T" & ChrW(236) & "nh tr" & ChrW(7841) & "ng HD"
    sNPT = "Ng" & ChrW(432) & ChrW(7901) & "i ph" & ChrW(7909) & _
               " tr" & ChrW(225) & "ch ch" & ChrW(237) & "nh"
    sBenB = "Nh" & ChrW(224) & " th" & ChrW(7847) & "u/" & _
        "Nh" & ChrW(224) & " cung c" & ChrW(7845) & "p"
    sHangMuc = "H" & ChrW(7841) & "ng m" & ChrW(7909) & "c"
    sGTHD = "Gi" & ChrW(225) & " tr" & ChrW(7883) & " HD"
    sDTT = ChrW(272) & ChrW(227) & " thanh to" & ChrW(225) & "n"
    sHSTT = "H" & ChrW(7891) & " s" & ChrW(417) & " thanh to" & ChrW(225) & "n " & _
               ChrW(273) & ChrW(227) & " " & ChrW(273) & ChrW(432) & ChrW(7907) & _
               "c x" & ChrW(225) & "c nh" & ChrW(7853) & "n"
    sGC = "Ghi ch" & ChrW(250)

    ' Nhan thong tin du an
    Dim sDuAn   As String: sDuAn = "D" & ChrW(7921) & " " & ChrW(225) & "n:"
    Dim sDiaChi As String: sDiaChi = ChrW(272) & ChrW(7883) & "a " & ChrW(273) & "i" & ChrW(7875) & "m:"
    Dim sHopDong As String: sHopDong = "Lo" & ChrW(7841) & "i h" & ChrW(7907) & "p " & ChrW(273) & ChrW(7891) & "ng:"

    ' Nhan tong hop nhom
    Dim sTong    As String: sTong = "T" & ChrW(7893) & "ng"
    Dim sTongCong As String: sTongCong = "T" & ChrW(7893) & "ng c" & ChrW(7897) & "ng"
    Dim sHopDongShort As String
    sHopDongShort = "h" & ChrW(7907) & "p " & ChrW(273) & ChrW(7891) & "ng"

    '================================================================
    ' 2. DO RONG COT + COT SHADOW
    '================================================================
    wsOut.Columns(1).ColumnWidth = 5      ' A: No
    wsOut.Columns(2).ColumnWidth = 18     ' B: Hop dong so
    wsOut.Columns(3).ColumnWidth = 18     ' C: Ngay ky
    wsOut.Columns(4).ColumnWidth = 18     ' D: Tinh trang HD
    wsOut.Columns(5).ColumnWidth = 18     ' E: Nguoi phu trach
    wsOut.Columns(6).ColumnWidth = 18     ' F: Ben B (merge F:G)
    wsOut.Columns(7).ColumnWidth = 18     ' G: (merged voi F)
    wsOut.Columns(8).ColumnWidth = 18     ' H: Hang muc (merge H:I)
    wsOut.Columns(9).ColumnWidth = 18     ' I: (merged voi H)
    wsOut.Columns(10).ColumnWidth = 18    ' J: Gia tri HD
    wsOut.Columns(11).ColumnWidth = 18    ' K: Da thanh toan
    wsOut.Columns(12).ColumnWidth = 18    ' L: Ho so TT
    wsOut.Columns(13).ColumnWidth = 18    ' M: Ghi chu
    ' Shadow cols (tu dong xoa sau AutoFit)
    wsOut.Columns(SH_BENB).ColumnWidth = 36     ' N: shadow Ben B (= col6+col7)
    wsOut.Columns(SH_HANGMUC).ColumnWidth = 36  ' O: shadow Hang muc (= col8+col9)

    '================================================================
    ' 3. DOC THONG TIN DU AN (tu sheet "PROJECT_SETTINGS" neu co)
    '================================================================
    Dim sProjName As String: sProjName = ""
    Dim sLocation     As String: sLocation = ""
    Dim sContType As String: sContType = ""
    
    GetProjectInfo sProjName, sLocation
    'lblDetailProject.Caption = IIf(sProjName <> "", sProjName, "")
    'lblDetailLocation.Caption = IIf(sLocation <> "", sLocation, "")
    
    'On Error Resume Next
    'Dim wsConfig As Worksheet
    'Set wsConfig = GetDataSheet("PROJECT_SETTINGS")
    'If Not wsConfig Is Nothing Then
    '    sProjName = Trim$(CStr(wsConfig.Cells(6, 2).value))
    '    sAddr = Trim$(CStr(wsConfig.Cells(7, 2).value))
    'End If
    'Set wsConfig = Nothing
    'On Error GoTo 0

    ' Lay sContType tu cboCategories dang duoc chon tren frmContractSearch
    On Error Resume Next
    sContType = Trim$(CStr(frmContractSearch.cboCategories.Text))
    On Error GoTo 0

    '================================================================
    ' 4. HANG HEADER DU AN (hang 1-3)
    '================================================================
    Dim hBorderColor As Long: hBorderColor = RGB(160, 160, 160)

    ' --- Hang 1: Du an ---
    'wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS)).UnMerge
    'wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS)).Merge
    With wsOut.Cells(r, 1)
        .value = sDuAn & " " & sProjName
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    wsOut.rows(r).RowHeight = 18
    r = r + 1

    ' --- Hang 2: Dia chi ---
    'wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS)).UnMerge
    'wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS)).Merge
    With wsOut.Cells(r, 1)
        .value = sDiaChi & " " & sLocation
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    wsOut.rows(r).RowHeight = 16
    r = r + 1

    ' --- Hang 3: Hop dong ---
    'wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS)).UnMerge
    'wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 7)).Merge
    With wsOut.Cells(r, 1)
        .value = sHopDong & " " & sContType
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    ' --- Hang 4: Tieu de chinh (merged tat ca cot, centered) ---
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS)).UnMerge
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS)).Merge
    With wsOut.Cells(r, 1)
        .value = sTitle
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(31, 73, 125)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    wsOut.rows(r).RowHeight = 32
    r = r + 1
    ' --- Hang 5: export
    With wsOut.Cells(r, NCOLS)
        .value = "Exported: " & Format(Now, "dd/mm/yyyy hh:mm")
        .HorizontalAlignment = xlRight
        .Font.Color = RGB(128, 128, 128)
        .Font.Size = 9
        .VerticalAlignment = xlVAlignCenter
    End With
    wsOut.rows(r).RowHeight = 16
    r = r + 1
    '================================================================
    ' 5. HANG HEADER COT (dark blue, wrap text, middle align)
    '================================================================
    wsOut.rows(r).RowHeight = 48  ' header du cao cho 2-3 dong wrap

    With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS))
        .Interior.Color = RGB(31, 73, 125)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .WrapText = True
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlCenter
    End With

    wsOut.Cells(r, 1).value = sNo

    wsOut.Cells(r, 2).value = sHDS
    wsOut.Cells(r, 2).HorizontalAlignment = xlCenter

    wsOut.Cells(r, 3).value = sNgKy
    wsOut.Cells(r, 4).value = sTT
    wsOut.Cells(r, 5).value = sNPT

    ' F:G merge -> Ben B
    wsOut.Range(wsOut.Cells(r, 6), wsOut.Cells(r, 7)).UnMerge
    wsOut.Range(wsOut.Cells(r, 6), wsOut.Cells(r, 7)).Merge
    wsOut.Cells(r, 6).value = sBenB
    wsOut.Cells(r, 6).HorizontalAlignment = xlCenter

    ' H:I merge -> Hang muc
    wsOut.Range(wsOut.Cells(r, 8), wsOut.Cells(r, 9)).UnMerge
    wsOut.Range(wsOut.Cells(r, 8), wsOut.Cells(r, 9)).Merge
    wsOut.Cells(r, 8).value = sHangMuc
    wsOut.Cells(r, 8).HorizontalAlignment = xlCenter

    wsOut.Cells(r, 10).value = sGTHD
    wsOut.Cells(r, 11).value = sDTT
    ' Doi ten header col L va M
    'Dim sIPCDone As String
    'sIPCDone = "IPC (Completed / Accepted)"    ' ten cot L
    Dim sRatio As String
    sRatio = "T" & ChrW(7927) & " l" & ChrW(7879) & " (%)"  ' "T? l? (%)"

    wsOut.Cells(r, 12).value = sHSTT
    wsOut.Cells(r, 12).WrapText = True
    wsOut.Cells(r, 13).value = sRatio

    ' Ke vien header
    ApplyDataRowBorders wsOut, r, NCOLS, 1, 2, 3, 4, 5, 7, 9, 10, 11, 12
    r = r + 1

    '================================================================
    ' 6. XAY DUNG PAYMENT SUMMARY / IPC COUNT
    '    dictPaidTotal(cid) = tong da thanh toan (VND)  <- tu gArrPayment
    '    dictIPCCount(cid)  = so IPC Completed/Accepted <- tu gArrDoc
    '================================================================
    If Not gDocLoaded Then BuildDocTrackerCache

    Dim dictPaidTotal As Object
    Dim dictIPCCount  As Object
    Set dictPaidTotal = CreateObject("Scripting.Dictionary")
    Set dictIPCCount = CreateObject("Scripting.Dictionary")

    ' --- Build dictPaidTotal tu gArrPayment ---
    If Not IsEmpty(gArrPayment) Then
        Dim pi As Long
        For pi = 1 To UBound(gArrPayment, 1)
            Dim pCID As String
            pCID = Trim$(CStr(gArrPayment(pi, 3)))
            If pCID = "" Then GoTo PAY_NEXT

            ' Chi tinh cac ledger row (co LedgerID o cot 2)
            If Trim$(CStr(gArrPayment(pi, 2))) = "" Then GoTo PAY_NEXT

            Dim pAmt As Double
            pAmt = val(CStr(gArrPayment(pi, 13)))
            If pAmt = 0 Then GoTo PAY_NEXT

            Dim pExRate As Double
            pExRate = val(CStr(gArrPayment(pi, 7)))
            If pExRate <= 1 Then pExRate = 1

            If Not dictPaidTotal.Exists(pCID) Then dictPaidTotal(pCID) = 0#
            dictPaidTotal(pCID) = dictPaidTotal(pCID) + pAmt * pExRate
PAY_NEXT:
        Next pi
    End If

    ' --- Build dictIPCCount tu gArrDoc (DOCUMENT_TRACKER) ---
    ' COL_DOC_TYPE   = 3  (cot C: IPC/CON/APP)
    ' COL_DOC_STATUS = 15 (cot O: Tinh trang)
    If Not IsEmpty(gArrDoc) And Not gDictDocByContract Is Nothing Then
        Dim dci As Long
        For dci = 0 To cCount - 1
            Dim dcid As String
            dcid = contractList(dci)
            If gDictDocByContract.Exists(dcid) Then
                Dim dColItem As Variant
                For Each dColItem In gDictDocByContract(dcid)
                    Dim di As Long
                    di = CLng(dColItem)
                    If UCase$(Trim$(CStr(gArrDoc(di, 3)))) = "IPC" Then
                        Dim dst As String
                        dst = UCase$(Trim$(CStr(gArrDoc(di, 15))))
                        If dst = "COMPLETED" Or dst = "ACCEPTED" Then
                            If Not dictIPCCount.Exists(dcid) Then dictIPCCount(dcid) = 0&
                            dictIPCCount(dcid) = dictIPCCount(dcid) + 1
                        End If
                    End If
                Next dColItem
            End If
        Next dci
    End If
    '================================================================
    ' 7. DATA ROWS — nhom A truoc, nhom B sau
    '================================================================
    Dim seqNo    As Long:  seqNo = 0
    Dim totalJ   As Double: totalJ = 0    ' grand total Gia tri HD
    Dim totalK   As Double: totalK = 0    ' grand total Da thanh toan

    Dim grpPass As Long
    For grpPass = 1 To 2

        Dim curGrp As String
        If grpPass = 1 Then curGrp = "A" Else curGrp = "B"

        ' Dem so hop dong trong nhom
        Dim grpCount As Long: grpCount = 0
        Dim ci As Long
        For ci = 0 To cCount - 1
            If GetGrpFromID(contractList(ci)) = curGrp Then
                grpCount = grpCount + 1
            End If
        Next ci
        If grpCount = 0 Then GoTo CONS_NEXTGRP

        ' --- Group header row ---
        Dim hdrColor As Long
        If grpPass = 1 Then
            hdrColor = RGB(68, 114, 196)
        Else
            hdrColor = RGB(0, 112, 192)
        End If

        wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS)).UnMerge
        wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS)).Merge
        With wsOut.Cells(r, 1)
            .value = "GROUP " & curGrp & "  (" & grpCount & " " & sHopDongShort & ")"
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = hdrColor
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlVAlignCenter
        End With
        wsOut.rows(r).RowHeight = 18
        r = r + 1

        ' --- Data rows trong nhom ---
        Dim altRow   As Boolean: altRow = False
        Dim grpJ     As Double:  grpJ = 0
        Dim grpK     As Double:  grpK = 0

        For ci = 0 To cCount - 1
            Dim cid As String
            cid = contractList(ci)
            If GetGrpFromID(cid) <> curGrp Then GoTo CONS_NEXT
            If Not gDictContract.Exists(cid) Then GoTo CONS_NEXT

            Dim arrRow As Long
            arrRow = CLng(gDictContract(cid))
            seqNo = seqNo + 1

            ' Mau xen ke
            Dim rowColor As Long
            If altRow Then
                rowColor = RGB(235, 241, 252)
            Else
                rowColor = RGB(255, 255, 255)
            End If
            altRow = Not altRow

            ' Dat mau nen truoc
            wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS)).Interior.Color = rowColor

            ' Kiem tra trang thai CANCEL
            Dim statusStr As String
            statusStr = UCase$(Trim$(CStr(gArrContract(arrRow, COL_STATUS))))

            ' === Col A: So thu tu ===
            With wsOut.Cells(r, 1)
                .value = seqNo
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlVAlignCenter
            End With

            ' === Col B: Hop dong so — ep kieu text ===
            With wsOut.Cells(r, 2)
                .NumberFormat = "@"
                .value = Trim$(CStr(gArrContract(arrRow, COL_CONTRACTNO)))
                .WrapText = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlVAlignCenter
            End With

            ' === Col C: Ngay ky — ep kieu ngay dd/mm/yyyy ===
            Dim sigV As Variant
            sigV = gArrContract(arrRow, COL_SIGNDATE)
            With wsOut.Cells(r, 3)
                If IsValidDMY(sigV) Then
                    .value = ParseDMY(sigV)
                    .NumberFormat = "dd/mm/yyyy"
                Else
                    .value = ""
                    .NumberFormat = "@"
                End If
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlVAlignCenter
            End With

            ' === Col D: Tinh trang HD ===
            With wsOut.Cells(r, 4)
                .value = Trim$(CStr(gArrContract(arrRow, COL_STATUS)))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlVAlignCenter
            End With

            ' === Col E: Nguoi phu trach chinh ===
            ' (doc tu COL_PERSON = 9; chinh hang CONST neu can)
            With wsOut.Cells(r, 5)
                On Error Resume Next
                .value = Trim$(CStr(gArrContract(arrRow, COL_PERSON)))
                On Error GoTo 0
                .WrapText = True
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlVAlignCenter
            End With

            ' === Col F:G (merge): Ben B — ghi shadow sang col N ===
            wsOut.Range(wsOut.Cells(r, 6), wsOut.Cells(r, 7)).UnMerge
            wsOut.Range(wsOut.Cells(r, 6), wsOut.Cells(r, 7)).Merge
            Dim sBenBVal As String
            sBenBVal = Trim$(CStr(gArrContract(arrRow, COL_COMPANYB)))
            With wsOut.Cells(r, 6)
                .value = sBenBVal
                .WrapText = True
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlVAlignCenter
            End With
            ' Shadow col N de AutoFit tinh chieu cao dung
            With wsOut.Cells(r, SH_BENB)
                .value = sBenBVal
                .WrapText = True
            End With

            ' === Col H:I (merge): Hang muc — ghi shadow sang col O ===
            wsOut.Range(wsOut.Cells(r, 8), wsOut.Cells(r, 9)).UnMerge
            wsOut.Range(wsOut.Cells(r, 8), wsOut.Cells(r, 9)).Merge
            Dim sHMVal As String
            sHMVal = Trim$(CStr(gArrContract(arrRow, COL_SCOPE)))
            With wsOut.Cells(r, 8)
                .value = sHMVal
                .WrapText = True
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlVAlignCenter
            End With
            ' Shadow col O
            With wsOut.Cells(r, SH_HANGMUC)
                .value = sHMVal
                .WrapText = True
            End With

            ' === Col J: Gia tri HD — ep kieu so #,##0 ===
            Dim valRaw As Variant
            valRaw = gArrContract(arrRow, COL_VALUE)
            Dim valNum As Double: valNum = 0
            With wsOut.Cells(r, 10)
                If IsNumeric(valRaw) Then
                    valNum = CDbl(valRaw)
                    .value = valNum
                    .NumberFormat = "#,##0"
                ElseIf IsNumeric(Replace(CStr(valRaw), ",", "")) Then
                    valNum = CDbl(Replace(CStr(valRaw), ",", ""))
                    .value = valNum
                    .NumberFormat = "#,##0"
                Else
                    .value = 0
                    .NumberFormat = "#,##0"
                End If
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlVAlignCenter
            End With
            If statusStr <> "CANCEL" Then grpJ = grpJ + valNum

            ' === Col K: Da thanh toan — ep kieu so #,##0 ===
            Dim paidTotal As Double: paidTotal = 0
            If dictPaidTotal.Exists(cid) Then
                paidTotal = CDbl(dictPaidTotal(cid))
            End If
            With wsOut.Cells(r, 11)
                .value = paidTotal
                .NumberFormat = "#,##0"
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlVAlignCenter
            End With
            If statusStr <> "CANCEL" Then grpK = grpK + paidTotal

            ' === Col L: Ho so thanh toan da duoc xac nhan (IPC Completed/Accepted) ===
            Dim ipcCount As Long: ipcCount = 0
            If dictIPCCount.Exists(cid) Then
                ipcCount = CLng(dictIPCCount(cid))
            End If
            With wsOut.Cells(r, 12)
                .value = ipcCount
                .NumberFormat = "00"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlVAlignCenter
            End With

            ' === Col M: Ty le % = K / J ===
            '     Neu J=0 va K>0 -> mac dinh 100%
            Dim ratioVal As Double
            If valNum = 0 Then
                If paidTotal > 0 Then
                    ratioVal = 1    ' 100%
                Else
                    ratioVal = 0
                End If
            Else
                ratioVal = paidTotal / valNum
            End If
            With wsOut.Cells(r, 13)
                .value = ratioVal
                .NumberFormat = "0.00%"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlVAlignCenter
            End With

            ' To mau do + in dam trang thai neu CANCEL; xanh neu COMPLETED
            If statusStr = "CANCEL" Then
                wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS)) _
                     .Font.Color = RGB(192, 0, 0)
                wsOut.Cells(r, 4).Font.Bold = True
            ElseIf statusStr = "COMPLETED" Then
                wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS)) _
                     .Font.Color = RGB(0, 112, 192)
            End If

            ' Ke vien dong
            ApplyDataRowBorders wsOut, r, NCOLS, 1, 2, 3, 4, 5, 7, 9, 10, 11, 12
            wsOut.rows(r).RowHeight = 35  ' chieu cao toi thieu (AutoFit se tang them)
            r = r + 1
CONS_NEXT:
        Next ci

        totalJ = totalJ + grpJ
        totalK = totalK + grpK

        ' --- Group subtotal row ---
        wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS)).UnMerge
        wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 9)).Merge
        With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS))
            .Interior.Color = RGB(198, 224, 180)
            .Font.Bold = True
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlLeft
        End With
        With wsOut.Cells(r, 1)
            .value = sTong & " GROUP " & curGrp & " (" & grpCount & " " & sHopDongShort & ")"
            .HorizontalAlignment = xlLeft
        End With
        With wsOut.Cells(r, 10)
            .value = grpJ
            .NumberFormat = "#,##0"
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlVAlignCenter
        End With
        With wsOut.Cells(r, 11)
            .value = grpK
            .NumberFormat = "#,##0"
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlVAlignCenter
        End With
        ' --- Group subtotal: them ty le nhom o col M ---
        Dim grpRatio As Double
        If grpJ = 0 Then
            grpRatio = IIf(grpK > 0, 1, 0)
        Else
            grpRatio = grpK / grpJ
        End If
        With wsOut.Cells(r, 13)
            .value = grpRatio
            .NumberFormat = "0.00%"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlVAlignCenter
            .Font.Bold = True
            .Interior.Color = RGB(198, 224, 180)
        End With
        
        ApplyDataRowBorders wsOut, r, NCOLS, 1, 9, 10, 11, 12
        wsOut.rows(r).RowHeight = 20
        r = r + 1

        ' Spacing giua hai nhom
        If grpPass = 1 Then
            wsOut.rows(r).RowHeight = 8
            r = r + 1
        End If

CONS_NEXTGRP:
    Next grpPass

    '================================================================
    ' 8. GRAND TOTAL ROW
    '================================================================
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS)).UnMerge
    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 9)).Merge
    With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, NCOLS))
        .Interior.Color = RGB(31, 73, 125)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlVAlignCenter
    End With
    With wsOut.Cells(r, 1)
        .value = sTongCong & " (" & cCount & " " & sHopDongShort & ")"
        .HorizontalAlignment = xlLeft
    End With
    With wsOut.Cells(r, 10)
        .value = totalJ
        .NumberFormat = "#,##0"
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlVAlignCenter
    End With
    With wsOut.Cells(r, 11)
        .value = totalK
        .NumberFormat = "#,##0"
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlVAlignCenter
    End With
    ' --- Grand total: them ty le tong cong o col M ---
    Dim totalRatio As Double
    If totalJ = 0 Then
        totalRatio = IIf(totalK > 0, 1, 0)
    Else
        totalRatio = totalK / totalJ
    End If
    With wsOut.Cells(r, 13)
        .value = totalRatio
        .NumberFormat = "0.00%"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
    End With
    ApplyDataRowBorders wsOut, r, NCOLS, 1, 9, 10, 11, 12
    wsOut.rows(r).RowHeight = 22
    r = r + 1

    '================================================================
    ' 9. AUTOFIT CHIEU CAO DONG (xu ly merge cell)
    '    AutoFit toan bo sheet -> chieu cao duoc tinh dua tren
    '    shadow cols (N, O) thay cho cac merge cells (F:G, H:I).
    '    Sau do xoa shadow cols va khoi phuc chieu cao.
    '================================================================
        
    Dim lastRow As Long
    lastRow = wsOut.UsedRange.rows(wsOut.UsedRange.rows.count).Row

    ' Luu chieu cao truoc khi xoa shadow (cho cac dong khong phu thuoc shadow)
    Dim rh() As Double
    ReDim rh(1 To lastRow)
    Dim rIdx As Long
    For rIdx = 1 To lastRow
        rh(rIdx) = wsOut.rows(rIdx).RowHeight
    Next rIdx

    ' AutoFit toan bo (khi shadow cols con ton tai)
    wsOut.UsedRange.EntireRow.AutoFit

    ' Luu chieu cao sau AutoFit (chieu cao dua vao ca shadow)
    Dim rhFit() As Double
    ReDim rhFit(1 To lastRow)
    For rIdx = 1 To lastRow
        rhFit(rIdx) = wsOut.rows(rIdx).RowHeight
    Next rIdx

    ' Xoa shadow cols
    wsOut.Range(wsOut.Columns(SH_BENB), wsOut.Columns(SH_HANGMUC)).Clear
    wsOut.Range(wsOut.Columns(SH_BENB), wsOut.Columns(SH_HANGMUC)).ColumnWidth = 0.5

    ' Khoi phuc chieu cao da fit (giu nguyen ket qua AutoFit)
    For rIdx = 1 To lastRow
        wsOut.rows(rIdx).RowHeight = rhFit(rIdx)
    Next rIdx

    ' An shadow cols
    wsOut.Columns(SH_BENB).hidden = True
    wsOut.Columns(SH_HANGMUC).hidden = True

    Application.DisplayAlerts = True

End Sub


