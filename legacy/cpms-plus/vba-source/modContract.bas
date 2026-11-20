Attribute VB_Name = "modContract"
Option Explicit

'================================
' GLOBAL IMPORT WORKBOOK (CACHE)
'================================
Public gImportWb As Workbook
Public gImportFilePath As String

'================================
' ENSURE CONTRACT_DB EXISTS
'================================
Public Function EnsureContractDB() As Worksheet
    Set EnsureContractDB = GetDataSheet("CONTRACT_DB")
End Function


'================================
' ENSURE DOCUMENT_TRACKER EXISTS
'================================
Public Function EnsureDocumentTracker() As Worksheet
    Set EnsureDocumentTracker = GetDataSheet("DOCUMENT_TRACKER")
    
End Function



'================================
' COPY ONE CONTRACT ROW
'================================
Public Sub CopyContractRow(wsFrom As Worksheet, rFrom As Long, _
                           wsTo As Worksheet, rTo As Long)

    Dim c As Long
    Dim contractID As String

    '==============================
    ' CONTRACT ID
    '==============================
    contractID = Trim(wsFrom.Cells(rFrom, 1).value)

    If contractID = "" Then
        contractID = GenerateNextContractID
    End If

    wsTo.Cells(rTo, 1).value = contractID

    '==============================
    ' OTHER COLUMNS
    '==============================
    For c = 2 To 22
        wsTo.Cells(rTo, c).value = wsFrom.Cells(rFrom, c).value
    Next c

    wsTo.Cells(rTo, 23).value = Now

End Sub


Public Function GenerateNextContractID(Optional ByVal parentID As String = "", _
                                        Optional prefix As String = "A") As String

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim curID As String
    Dim maxNo As Long
    Dim suffix As String
    Dim contractPrefix As String
    contractPrefix = prefix & "-"

    Set ws = EnsureContractDB
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    maxNo = 0

    For i = 2 To lastRow
        curID = Trim(ws.Cells(i, 1).value)
        If curID = "" Then GoTo NEXT_I

        If parentID = "" Then
            If curID Like contractPrefix & Year(Date) & "-##" Then
                suffix = Right(curID, 2)
                If IsNumeric(suffix) Then
                    maxNo = Application.Max(maxNo, CLng(suffix))
                End If
            End If
        Else
            If curID Like parentID & "-##" Then
                suffix = Right(curID, 2)
                If IsNumeric(suffix) Then
                    maxNo = Application.Max(maxNo, CLng(suffix))
                End If
            End If
        End If

NEXT_I:
    Next i

    If parentID = "" Then
        GenerateNextContractID = contractPrefix & Year(Date) & "-" & Format(maxNo + 1, "00")
    Else
        GenerateNextContractID = parentID & "-" & Format(maxNo + 1, "00")
    End If

End Function


Public Function GetSafeDate(v As Variant) As Variant

    Dim s As String
    Dim arr() As String

    s = Trim(CStr(v))

    If s = "" Then
        GetSafeDate = ""
        Exit Function
    End If

    On Error GoTo INVALID_DATE

    If InStr(s, "/") > 0 Then
        arr = Split(s, "/")
        If UBound(arr) = 2 Then
            GetSafeDate = DateSerial( _
                CLng(arr(2)), _
                CLng(arr(1)), _
                CLng(arr(0)))
            Exit Function
        End If
    End If

    If IsDate(s) Then
        GetSafeDate = CDate(s)
    Else
        GetSafeDate = ""
    End If

    Exit Function

INVALID_DATE:
    GetSafeDate = ""

End Function


Public Function GetSafeText(v As Variant) As String
    If IsNull(v) Or v = "" Then
        GetSafeText = ""
    Else
        GetSafeText = CStr(v)
    End If
End Function


Public Function GetSafeNumber(v As Variant) As Double
    Dim s As String

    If Trim(v & "") = "" Then
        GetSafeNumber = 0
        Exit Function
    End If

    s = CStr(v)
    s = Replace(s, ".", "")
    s = Replace(s, ",", "")

    If IsNumeric(s) Then
        GetSafeNumber = CDbl(s)
    Else
        GetSafeNumber = 0
    End If
End Function


Public Function GetSafeDouble(v As Variant) As Double
    On Error Resume Next
    GetSafeDouble = CDbl(v)
    If Err.Number <> 0 Then GetSafeDouble = 0
    On Error GoTo 0
End Function


Public Sub TextBox_Enter(tb As MSForms.TextBox)
    If tb.value = tb.Tag Then
        tb.value = ""
        tb.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Public Sub TextBox_Exit(ByVal tb As MSForms.TextBox)

    Dim v As Double
    Dim fmt As String

    If Trim(tb.value) = "" Then
        tb.value = tb.Tag
        tb.ForeColor = RGB(150, 150, 150)
        Exit Sub
    End If

    If Not IsNumeric(Replace(tb.value, ",", "")) Then Exit Sub

    fmt = tb.Tag
    v = CDbl(Replace(tb.value, ",", ""))
    tb.value = Format(v, fmt)
    tb.ForeColor = RGB(0, 0, 0)

End Sub


Public Function FindContractRowByID(ws As Worksheet, contractID As String) As Long

    Dim lastRow As Long
    Dim r As Long

    FindContractRowByID = 0
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If Trim(ws.Cells(r, 1).value) = contractID Then
            FindContractRowByID = r
            Exit Function
        End If
    Next r

End Function


Public Function GetUniqueValuesFromContractDB(ColIndex As Long) As Collection
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim col As New Collection
    Dim val As String

    Set ws = EnsureContractDB
    lastRow = ws.Cells(ws.rows.count, ColIndex).End(xlUp).Row

    On Error Resume Next
    For r = 2 To lastRow
        val = Trim(ws.Cells(r, ColIndex).value)
        If val <> "" Then col.Add val, val
    Next r
    On Error GoTo 0

    Set GetUniqueValuesFromContractDB = col
End Function


Public Sub LoadComboFromContractDB(cbo As MSForms.ComboBox, ColIndex As Long)
    Dim col As Collection
    Dim i As Long

    Set col = GetUniqueValuesFromContractDB(ColIndex)

    With cbo
        .Clear
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        For i = 1 To col.count
            .AddItem col(i)
        Next i
    End With
End Sub


Public Sub LoadComboFromContractDB_2Cols(cbo As ComboBox, col1 As Long, col2 As Long)

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim dict As Object
    Dim v As String

    Set ws = EnsureContractDB
    Set dict = CreateObject("Scripting.Dictionary")

    cbo.Clear

    lastRow = ws.Cells(ws.rows.count, col1).End(xlUp).Row

    For r = 2 To lastRow
        v = Trim(ws.Cells(r, col1).value)
        If v <> "" Then
            If Not dict.Exists(v) Then dict.Add v, 1
        End If

        v = Trim(ws.Cells(r, col2).value)
        If v <> "" Then
            If Not dict.Exists(v) Then dict.Add v, 1
        End If
    Next r

    Dim key As Variant
    For Each key In dict.keys
        cbo.AddItem key
    Next key

End Sub


Public Function GetAllContractFamily(contractID As String) As Collection

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim col As New Collection
    Dim curID As String

    Set ws = EnsureContractDB
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row

    For r = 2 To lastRow
        curID = Trim(ws.Cells(r, 1).value)
        If curID <> "" Then
            If curID = contractID _
               Or curID Like contractID & "-*" Then
                col.Add curID
            End If
        End If
    Next r

    Set GetAllContractFamily = col

End Function


Public Function ContractHasPayment(contractID As String) As Boolean

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim curID As String

    Set ws = EnsurePaymentLedger
    lastRow = ws.Cells(ws.rows.count, 3).End(xlUp).Row

    For r = 2 To lastRow
        curID = Trim(ws.Cells(r, 3).value)

        If curID = contractID Then
            If GetSafeDouble(ws.Cells(r, 12).value) <> 0 _
               Or GetSafeDouble(ws.Cells(r, 10).value) <> 0 Then
                ContractHasPayment = True
                Exit Function
            End If
        End If
    Next r

    ContractHasPayment = False

End Function


Public Function ContractFamilyHasPayment(contractID As String) As Boolean

    Dim col As Collection
    Dim id As Variant

    Set col = GetAllContractFamily(contractID)

    For Each id In col
        If ContractHasPayment(CStr(id)) Then
            ContractFamilyHasPayment = True
            Exit Function
        End If
    Next id

    ContractFamilyHasPayment = False

End Function


Public Sub DeleteContractFamily(contractID As String)

    Dim ws As Worksheet
    Dim r As Long

    Set ws = EnsureContractDB

    For r = ws.Cells(ws.rows.count, 1).End(xlUp).Row To 2 Step -1
        If Trim(ws.Cells(r, 1).value) = contractID _
           Or Trim(ws.Cells(r, 1).value) Like contractID & "-*" Then
            ws.rows(r).Delete
        End If
    Next r

End Sub


Public Sub DeleteLedgerByContractFamily(contractID As String)

    Dim ws As Worksheet
    Dim r As Long

    Set ws = EnsurePaymentLedger

    For r = ws.Cells(ws.rows.count, 3).End(xlUp).Row To 2 Step -1
        If Trim(ws.Cells(r, 3).value) = contractID _
           Or Trim(ws.Cells(r, 3).value) Like contractID & "-*" Then
            ws.rows(r).Delete
        End If
    Next r

End Sub


Public Sub SortContractDBBySigningDate(Optional dummy As Variant = 0)

    Dim ws As Worksheet
    Set ws = EnsureContractDB

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim groups As Variant
    groups = Array("GROUP-A", "GROUP-B")

    Dim g As Variant

    For Each g In groups
        Call SortOneGroup(ws, CStr(g))
    Next g

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub


Private Sub SortOneGroup(ws As Worksheet, groupID As String)

    Dim groupRow As Long
    groupRow = FindContractRowByID(ws, groupID)
    If groupRow = 0 Then Exit Sub

    Dim rows As Collection
    Set rows = CollectContractsUnderGroup(ws, groupRow)

    If rows.count <= 1 Then Exit Sub

    Call SortContractBlock(ws, rows)

End Sub


Private Function CollectContractsUnderGroup(ws As Worksheet, groupRow As Long) As Collection

    Dim col As New Collection
    Dim r As Long, lastRow As Long
    Dim id As String

    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row

    For r = groupRow + 1 To lastRow
        id = Trim(ws.Cells(r, 1).value)

        If id Like "GROUP-*" Then Exit For

        If CountIDLevel(id) = 3 Then
            col.Add id
        End If
    Next r

    Set CollectContractsUnderGroup = col

End Function


Private Sub SortContractBlock(ws As Worksheet, contracts As Collection)

    Dim arr() As Variant
    Dim i As Long

    ReDim arr(1 To contracts.count, 1 To 2)

    For i = 1 To contracts.count
        arr(i, 1) = contracts(i)
        arr(i, 2) = GetSigningDate(ws, contracts(i))
    Next i

    Call QuickSortByDate(arr, 1, UBound(arr, 1))

    For i = 1 To UBound(arr, 1)
        Call MoveContractFamily(ws, arr(i, 1))
        Call SortAppendixRecursive(ws, arr(i, 1))
    Next i

End Sub


Private Sub SortAppendixRecursive(ws As Worksheet, parentID As String)

    Dim children As Collection
    Set children = GetDirectChildren(ws, parentID)

    If children.count = 0 Then Exit Sub

    Dim arr() As Variant
    Dim i As Long

    ReDim arr(1 To children.count, 1 To 2)

    For i = 1 To children.count
        arr(i, 1) = children(i)
        arr(i, 2) = GetSigningDate(ws, children(i))
    Next i

    Call QuickSortByDate(arr, 1, UBound(arr, 1))

    For i = 1 To UBound(arr, 1)
        Call MoveContractFamily(ws, arr(i, 1))
        Call SortAppendixRecursive(ws, arr(i, 1))
    Next i

End Sub


Private Function CountIDLevel(id As String) As Long
    CountIDLevel = UBound(Split(id, "-")) + 1
End Function


Private Function GetSigningDate(ws As Worksheet, contractID As String) As Date
    Dim r As Long
    r = FindContractRowByID(ws, contractID)
    If r = 0 Then
        GetSigningDate = DateSerial(1900, 1, 1)
    Else
        GetSigningDate = ws.Cells(r, SIGNING_DATE_COL).value
    End If
End Function


Public Sub SortLedgerByContractOrder(Optional dummy As Variant = 0)

    Dim wsC As Worksheet, wsL As Worksheet
    Set wsC = EnsureContractDB
    Set wsL = EnsurePaymentLedger

    Dim order As Collection
    Set order = GetContractIDOrder(wsC)

    Application.ScreenUpdating = False

    Dim id As Variant
    For Each id In order
        Call MoveLedgerBlock(wsL, CStr(id))
    Next id

    Application.ScreenUpdating = True

End Sub


Private Function GetContractIDOrder(ws As Worksheet) As Collection

    Dim col As New Collection
    Dim r As Long, lastRow As Long
    Dim id As String

    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row

    For r = 2 To lastRow
        id = Trim(ws.Cells(r, 1).value)
        If id <> "" And Not id Like "GROUP-*" Then
            col.Add id
        End If
    Next r

    Set GetContractIDOrder = col

End Function


Public Sub UpdateLedgerFormulaByStatus( _
    ByVal contractID As String)

    Dim wsDB As Worksheet
    Dim wsLG As Worksheet

    Dim rDB As Long
    Dim lastRow As Long
    Dim r As Long

    Dim statusText As String
    Dim isCancel As Boolean

    Set wsDB = EnsureContractDB
    Set wsLG = EnsurePaymentLedger

    rDB = FindContractRowByID(wsDB, contractID)

    If rDB <= 0 Then Exit Sub

    statusText = _
        UCase$(Trim$(wsDB.Cells(rDB, "F").value))

    isCancel = (statusText = "CANCEL")

    lastRow = wsLG.Cells( _
        wsLG.rows.count, "C").End(xlUp).Row

    For r = 2 To lastRow

        If Trim$(wsLG.Cells(r, "C").value) = contractID Then

            If Trim$(wsLG.Cells(r, "E").value) = "" Then

                If isCancel Then
                    wsLG.Cells(r, "W").value = 0
                    wsLG.Cells(r, "X").value = 0
                    wsLG.Cells(r, "Y").value = 0
                    wsLG.Cells(r, "Z").value = 0
                Else
                    wsLG.Cells(r, "W").Formula = _
                        "=SUMIFS($W:$W,$A:$A,$C" & r & ")"
                    wsLG.Cells(r, "X").Formula = _
                        "=SUMIFS($X:$X,$A:$A,$C" & r & ")"
                    wsLG.Cells(r, "Y").Formula = _
                        "=SUMIFS($Y:$Y,$A:$A,$C" & r & ")"
                    wsLG.Cells(r, "Z").Formula = _
                        "=SUMIFS($Z:$Z,$A:$A,$C" & r & ")"
                End If

                Exit For

            End If

        End If

    Next r

End Sub

' Module: modSettings
' Hàm tr? v? giá tr? t? sheet PROJECT_SETTINGS theo key
Public Function GetProjectSetting(ByVal sKey As String) As String
    Const SHEET_NAME As String = "PROJECT_SETTINGS"

    Dim wsPS As Worksheet
    On Error Resume Next
    Set wsPS = GetDataSheet(SHEET_NAME)
    On Error GoTo 0
    If wsPS Is Nothing Then Exit Function

    Dim i As Long
    Dim lastR As Long
    lastR = wsPS.Cells(wsPS.rows.count, 1).End(xlUp).Row

    For i = 1 To lastR
        If UCase$(Trim$(CStr(wsPS.Cells(i, 1).value))) = UCase$(Trim$(sKey)) Then
            GetProjectSetting = Trim$(CStr(wsPS.Cells(i, 2).value))
            Exit Function
        End If
    Next i
End Function

' Hàm ti?n ích: tr? v? c? 2 giá tr? cùng lúc qua tham s? ByRef
Public Sub GetProjectInfo(ByRef sProjName As String, ByRef sLocation As String)
    sProjName = GetProjectSetting("PROJ_NAME")
    sLocation = GetProjectSetting("LOCATION")
End Sub
Public Function BuildDisplayContractNo(ByVal contractID As String) As String

    BuildDisplayContractNo = ""

    If contractID = "" Or contractID = "0" Then Exit Function
    If gDictContract Is Nothing Then Exit Function
    If Not gDictContract.Exists(contractID) Then Exit Function

    Dim r As Long
    r = gDictContract(contractID)

    Dim contractNo As String
    contractNo = Trim(CStr(gArrContract(r, 2)))
    If contractNo = "" Then
        BuildDisplayContractNo = contractID
        Exit Function
    End If

    '------------------------------------------------------------
    ' Xây dict ngu?c: chu?i seq (c?t 24) -> ch? s? dòng
    '------------------------------------------------------------
    Dim dictSeq As Object
    Set dictSeq = CreateObject("Scripting.Dictionary")

    Dim iSeq As Long
    Dim seqK As String
    For iSeq = 1 To UBound(gArrContract, 1)
        seqK = Trim(CStr(gArrContract(iSeq, 24)))
        If seqK <> "" Then dictSeq(seqK) = iSeq
    Next iSeq

    '------------------------------------------------------------
    ' Duy?t chu?i t? tiên (ancestor chain) qua cidSeq (c?t 24)
    ' Ví d? cidSeq = "1.2.3" ? t? tiên trung gian: "1.2", g?c: "1"
    '------------------------------------------------------------
    Dim cidSeq  As String: cidSeq = Trim(CStr(gArrContract(r, 24)))
    Dim rootNo  As String: rootNo = ""
    Dim anChain As String: anChain = ""

    If InStr(cidSeq, ".") > 0 Then
        Dim tmpSeq As String: tmpSeq = cidSeq

        Do While InStr(tmpSeq, ".") > 0
            Dim ldp As Long: ldp = 0
            Dim kk  As Long
            For kk = Len(tmpSeq) To 1 Step -1
                If mID(tmpSeq, kk, 1) = "." Then ldp = kk: Exit For
            Next kk
            If ldp = 0 Then Exit Do

            tmpSeq = Left(tmpSeq, ldp - 1)

            If InStr(tmpSeq, ".") > 0 Then
                ' Còn d?u ch?m ? dây là c?p cha trung gian
                If dictSeq.Exists(tmpSeq) Then
                    Dim anR As Long:   anR = dictSeq(tmpSeq)
                    Dim anN As String: anN = Trim(CStr(gArrContract(anR, 2)))
                    If anN <> "" Then
                        If anChain = "" Then
                            anChain = anN
                        Else
                            anChain = anChain & "." & anN
                        End If
                    End If
                End If
            Else
                ' Không còn d?u ch?m ? dây là HÐ g?c c?p cao nh?t
                If dictSeq.Exists(tmpSeq) Then
                    Dim rootR As Long: rootR = dictSeq(tmpSeq)
                    rootNo = Trim(CStr(gArrContract(rootR, 2)))
                End If
            End If
        Loop
    End If

    Set dictSeq = Nothing

    '------------------------------------------------------------
    ' Ghép chu?i hi?n th?
    '------------------------------------------------------------
    Dim displayNo As String
    displayNo = contractNo

    If anChain <> "" Then
        displayNo = contractNo & " (" & anChain & ")"
    End If

    Dim isAP As Boolean
    isAP = (UCase(Left(contractNo, 4)) = "APP-" Or _
            UCase(Left(contractNo, 4)) = "PKG-")

    If isAP And rootNo <> "" Then
        displayNo = displayNo & "  " & ChrW(8212) & "  " & rootNo
    End If

    BuildDisplayContractNo = displayNo

End Function




