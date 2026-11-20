Attribute VB_Name = "modGlobal"
Option Explicit

Public gArrContract As Variant
Public gDictContract As Object
Public gContractLoaded As Boolean

Public gArrPayment As Variant
'Public gArrContract As Variant

Public gPaymentLoaded As Boolean

'Public gDictContract As Object
Public gDictPaymentRows As Object

Public gPaymentEntryLoaded As Boolean

Public gArrLedger As Variant
'Public gArrContract As Variant

'Public gDictContract As Object
Public gDictLedgerByContract As Object
Public gDictContractRow As Object

Public gDictBCPBRows As Object       ' B-CPB map tu BuildPaymentCache
Public gDictBCPBLedger As Object     ' B-CPB map tu BuildPaymentEntryCache
Public gArrDoc          As Variant
Public gDictDocByContract As Object   ' contractID -> Collection of row indices in gArrDoc
Public gDocLoaded       As Boolean
Public gBCPBLoaded       As Boolean
Public gCacheProjectID As String
Public mLastBrowseDir As String   ' remembers last picked folder

Public Sub BuildDocTrackerCache(Optional dummy As Variant = 0)

    Dim wsD     As Worksheet
    Dim lastRow As Long
    Dim i       As Long
    Dim key     As String

    On Error Resume Next
    Set wsD = GetDataSheet("DOCUMENT_TRACKER")
    On Error GoTo 0

    ' *** KH?I T?O DICT TRU?C ***
    Set gDictDocByContract = CreateObject("Scripting.Dictionary")

    If wsD Is Nothing Then
        gDocLoaded = False
        Exit Sub
    End If

    lastRow = wsD.Cells(wsD.rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then
        gArrDoc = Empty
        gDocLoaded = True
        Exit Sub
    End If

    gArrDoc = wsD.Range("A2:S" & lastRow).value

    For i = 1 To UBound(gArrDoc, 1)
        key = Trim$(CStr(gArrDoc(i, 2)))
        If key <> "" Then
            If Not gDictDocByContract.Exists(key) Then
                Set gDictDocByContract(key) = New Collection
            End If
            gDictDocByContract(key).Add i
        End If
    Next i

    gDocLoaded = True

End Sub
Public Sub BuildGlobalContractCache(Optional dummy As Variant = 0)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = EnsureContractDB

    ' *** KH?I T?O DICT TRU?C ***
    Set gDictContract = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then
        gArrContract = Empty
        gContractLoaded = True
        Exit Sub
    End If

    gArrContract = ws.Range("A2:Z" & lastRow).value

    For i = 1 To UBound(gArrContract, 1)
        gDictContract(gArrContract(i, 1)) = i
    Next i

    gContractLoaded = True

End Sub

Public Sub BuildPaymentCache(Optional dummy As Variant = 0)

    Dim wsC As Worksheet
    Dim wsP As Worksheet
    Dim lastRowC As Long
    Dim lastRowP As Long
    Dim i As Long
    Dim key As String

    Set wsC = EnsureContractDB
    Set wsP = EnsurePaymentLedger

    ' *** KH?I T?O DICT TRU?C — tránh dict cu còn sót khi exit s?m ***
    Set gDictContract = CreateObject("Scripting.Dictionary")
    Set gDictPaymentRows = CreateObject("Scripting.Dictionary")
    Set gDictBCPBRows = CreateObject("Scripting.Dictionary")

    lastRowC = wsC.Cells(wsC.rows.count, "A").End(xlUp).Row
    If lastRowC < 2 Then
        gArrContract = Empty
    Else
        gArrContract = wsC.Range("A2:Z" & lastRowC).value
    End If

    lastRowP = wsP.Cells(wsP.rows.count, "C").End(xlUp).Row
    If lastRowP < 2 Then
        gArrPayment = Empty
        gPaymentLoaded = True
        Exit Sub   ' Dict dã là empty dict s?ch — an toàn
    End If

    gArrPayment = wsP.Range("A2:Z" & lastRowP).value

    ' CONTRACT MAP
    For i = 1 To UBound(gArrContract, 1)
        gDictContract(gArrContract(i, 1)) = i
    Next i

    ' PAYMENT MAP
    For i = 1 To UBound(gArrPayment, 1)
        key = Trim(CStr(gArrPayment(i, 3)))
        If key <> "" Then
            If Not gDictPaymentRows.Exists(key) Then
                Set gDictPaymentRows(key) = New Collection
            End If
            gDictPaymentRows(key).Add i
        End If
    Next i

    ' B-CPB MAP
    For i = 1 To UBound(gArrPayment, 1)
        key = Trim(CStr(gArrPayment(i, 3)))
        If Left$(UCase$(key), 5) = "B-CPB" Then
            If Not gDictBCPBRows.Exists(key) Then
                Set gDictBCPBRows(key) = New Collection
            End If
            gDictBCPBRows(key).Add i
        End If
    Next i

    gBCPBLoaded = True
    gPaymentLoaded = True

End Sub
Public Sub BuildPaymentEntryCache(Optional dummy As Variant = 0)

    Dim wsL As Worksheet
    Dim wsC As Worksheet
    Dim lastRowL As Long
    Dim lastRowC As Long
    Dim i As Long
    Dim key As String

    Set wsL = EnsurePaymentLedger
    Set wsC = EnsureContractDB

    ' *** KH?I T?O DICT TRU?C — tránh dict cu sót ***
    Set gDictContract = CreateObject("Scripting.Dictionary")
    Set gDictLedgerByContract = CreateObject("Scripting.Dictionary")
    Set gDictContractRow = CreateObject("Scripting.Dictionary")
    Set gDictBCPBLedger = CreateObject("Scripting.Dictionary")

    lastRowL = wsL.Cells(wsL.rows.count, "C").End(xlUp).Row
    lastRowC = wsC.Cells(wsC.rows.count, "A").End(xlUp).Row

    If lastRowL < 2 Then
        gArrLedger = Empty
        gPaymentEntryLoaded = True
        Exit Sub
    End If

    gArrLedger = wsL.Range("A2:Z" & lastRowL).value
    gArrContract = wsC.Range("A2:Z" & lastRowC).value

    ' CONTRACT MAP
    For i = 1 To UBound(gArrContract, 1)
        key = Trim(CStr(gArrContract(i, 1)))
        If key <> "" Then
            gDictContract(key) = i
            gDictContractRow(key) = i
        End If
    Next i

    ' LEDGER MAP
    For i = 1 To UBound(gArrLedger, 1)
        key = Trim(CStr(gArrLedger(i, 3)))
        If key <> "" Then
            If Not gDictLedgerByContract.Exists(key) Then
                Set gDictLedgerByContract(key) = New Collection
            End If
            gDictLedgerByContract(key).Add i
        End If
    Next i

    ' B-CPB LEDGER MAP
    For i = 1 To UBound(gArrLedger, 1)
        key = Trim(CStr(gArrLedger(i, 3)))
        If Left$(UCase$(key), 5) = "B-CPB" Then
            If Not gDictBCPBLedger.Exists(key) Then
                Set gDictBCPBLedger(key) = New Collection
            End If
            gDictBCPBLedger(key).Add i
        End If
    Next i

    gPaymentEntryLoaded = True

End Sub

'================================================================
' CONTRACT META — Luu/doc 5 truong meta trong col Z CONTRACT_DB
'
' Format col Z: "PROJ=CHC1|CLIENT=CLI|MAINCON=VEC|SUPCON=ART|SUBCON=CTC"
' 5 field keys: PROJ, CLIENT, MAINCON, SUPCON, SUBCON
'================================================================

'----------------------------------------------------------------
' Doc mot truong meta tu cache (gArrContract col 25 = col Z)
' fieldKey: "PROJ" | "CLIENT" | "MAINCON" | "SUPCON" | "SUBCON"
'----------------------------------------------------------------
Public Function GetContractMetaField( _
        ByVal contractID As String, _
        ByVal fieldKey As String) As String

    GetContractMetaField = ""
    If gDictContract Is Nothing Then Exit Function
    If Not gDictContract.Exists(contractID) Then Exit Function

    Dim r As Long
    r = gDictContract(contractID)
    Dim metaStr As String
    metaStr = Trim$(CStr(gArrContract(r, 26)))
    If metaStr = "" Then Exit Function

    GetContractMetaField = ParseMetaField(metaStr, UCase$(fieldKey))
End Function

Private Function ParseMetaField( _
        ByVal metaStr As String, _
        ByVal fieldKey As String) As String

    ParseMetaField = ""
    Dim parts() As String
    parts = Split(metaStr, "|")
    Dim p As Variant
    For Each p In parts
        Dim kv() As String
        kv = Split(CStr(p), "=")
        If UBound(kv) >= 1 Then
            If UCase$(Trim$(kv(0))) = UCase$(fieldKey) Then
                ParseMetaField = Trim$(kv(1))
                Exit Function
            End If
        End If
    Next p
End Function

Public Sub SaveContractMetaField( _
        ByVal contractID As String, _
        ByVal fieldKey As String, _
        ByVal newValue As String)

    On Error GoTo ErrSave

    Dim ws As Worksheet
    Set ws = EnsureContractDB

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim targetRow As Long: targetRow = 0
    Dim i As Long
    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, "A").value)) = contractID Then
            targetRow = i
            Exit For
        End If
    Next i
    If targetRow = 0 Then
        MsgBox "Contract not found: " & contractID, vbExclamation
        Exit Sub
    End If

    Dim metaStr As String
    metaStr = Trim$(CStr(ws.Cells(targetRow, "Z").value))

    metaStr = UpdateMetaField(metaStr, UCase$(fieldKey), newValue)

    ws.Cells(targetRow, "Z").value = metaStr

    If Not gDictContract Is Nothing Then
        If gDictContract.Exists(contractID) Then
            Dim r As Long
            r = gDictContract(contractID)
            gArrContract(r, 26) = metaStr
        End If
    End If

    Exit Sub
ErrSave:
    MsgBox "SaveContractMetaField: " & Err.Description, vbExclamation
End Sub

Private Function UpdateMetaField( _
        ByVal metaStr As String, _
        ByVal fieldKey As String, _
        ByVal newValue As String) As String

    Dim parts() As String
    Dim rebuilt As String
    Dim found As Boolean: found = False

    If metaStr <> "" Then
        parts = Split(metaStr, "|")
        Dim p As Variant
        For Each p In parts
            Dim kv() As String
            kv = Split(CStr(p), "=")
            If UBound(kv) >= 1 Then
                If UCase$(Trim$(kv(0))) = fieldKey Then
                    If rebuilt <> "" Then rebuilt = rebuilt & "|"
                    rebuilt = rebuilt & fieldKey & "=" & newValue
                    found = True
                Else
                    If rebuilt <> "" Then rebuilt = rebuilt & "|"
                    rebuilt = rebuilt & CStr(p)
                End If
            Else
                If Trim$(CStr(p)) <> "" Then
                    If rebuilt <> "" Then rebuilt = rebuilt & "|"
                    rebuilt = rebuilt & CStr(p)
                End If
            End If
        Next p
    End If

    If Not found Then
        If rebuilt <> "" Then rebuilt = rebuilt & "|"
        rebuilt = rebuilt & fieldKey & "=" & newValue
    End If

    UpdateMetaField = rebuilt
End Function


Public Sub RefreshAllCache(Optional ByVal projectID As String = "")

    gContractLoaded = False
    gPaymentLoaded = False
    gPaymentEntryLoaded = False
    gBCPBLoaded = False
    gDocLoaded = False

    BuildGlobalContractCache
    BuildPaymentCache
    BuildPaymentEntryCache
    BuildDocTrackerCache

    ' ? Gán projectID SAU KHI cache build xong — d?c t? file, không doán
    gCacheProjectID = Trim$(ReadCurrentProjectValue("ProjectID"))

End Sub



'================================================================
' FIX — RebuildContractSequence
'
' NGUYÊN NHÂN BUG:
'   GetContractLevel("A-2026-01-01") tr? v? 3 (có 3 d?u "-")
'   GetContractLevel("B-2026-02")    tr? v? 2 (có 2 d?u "-")
'
'   Code cu ch? coi level=2 là "main contract" ? GROUP-A luôn b? b? qua
'   vì ID c?a nó có 3 d?u "-" (dài hon GROUP-B 1 do?n).
'
' CÁCH S?A:
'   Không dùng GetContractLevel d? xác d?nh main/child.
'   Thay b?ng: m?t contract là MAIN n?u không có contract nào khác
'   trong cùng sheet mà khi thêm "-" vào cu?i thì là ti?n t? c?a nó.
'   (VD: "A-2026-01-01" là main vì không có "A-2026-01" trong sheet)
'
' THAY TH? TOÀN B? hàm RebuildContractSequence b?ng do?n du?i dây.
'================================================================

Public Sub RebuildContractSequence(Optional dummy As Variant = 0)

    On Error GoTo ERR_HANDLER

    Dim ws       As Worksheet
    Dim lastRow  As Long
    Dim i        As Long

    Set ws = EnsureContractDB
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    '==================================================================
    ' BU?C 1: Thu th?p t?t c? contract (lo?i tr? GROUP header)
    '==================================================================
    Dim total As Long
    total = 0
    For i = 2 To lastRow
        Dim cid As String
        cid = Trim$(CStr(ws.Cells(i, "A").value))
        If cid <> "" And cid <> "GROUP-A" And cid <> "GROUP-B" Then
            total = total + 1
        End If
    Next i
    If total = 0 Then Exit Sub

    ' arrC(n, 1) = contractID
    ' arrC(n, 2) = sheet row
    ' arrC(n, 3) = SignDate (col D)
    ' arrC(n, 4) = group ("GROUP-A" / "GROUP-B")
    Dim arrC() As Variant
    ReDim arrC(1 To total, 1 To 4)

    Dim dictID As Object
    Set dictID = CreateObject("Scripting.Dictionary")

    total = 0
    Dim grp As String
    grp = ""
    For i = 2 To lastRow
        cid = Trim$(CStr(ws.Cells(i, "A").value))
        If cid = "GROUP-A" Or cid = "GROUP-B" Then
            grp = cid
        ElseIf cid <> "" Then
            total = total + 1
            arrC(total, 1) = cid
            arrC(total, 2) = i
            arrC(total, 3) = ws.Cells(i, "D").value
            arrC(total, 4) = grp
            dictID(cid) = total   ' index trong arrC
        End If
    Next i

    '==================================================================
    ' BU?C 2: Xác d?nh PARENT c?a t?ng contract
    '         (tìm contract-khác trong sheet mà khi thêm "-" là ti?n t?)
    '==================================================================
    ' arrParent(n) = index trong arrC c?a parent, ho?c 0 n?u là main
    Dim arrParent() As Long
    ReDim arrParent(1 To total)

    Dim n As Long
    For n = 1 To total
        Dim id As String
        id = CStr(arrC(n, 1))
        Dim parentIdx As Long
        parentIdx = 0

        ' Th? l?n lu?t t?ng ti?n t? b?ng cách bóc do?n cu?i
        Dim pos As Long
        pos = InStrRev(id, "-")
        Do While pos > 1
            Dim candidate As String
            candidate = Left$(id, pos - 1)
            If dictID.Exists(candidate) Then
                parentIdx = dictID(candidate)
                Exit Do
            End If
            pos = InStrRev(id, "-", pos - 1)
        Loop

        arrParent(n) = parentIdx
    Next n

    '==================================================================
    ' BU?C 3: Clear c?t X, Y
    '==================================================================
    ws.Range("X2:Y" & lastRow).ClearContents

    '==================================================================
    ' BU?C 4: L?y danh sách MAIN contracts (parent = 0)
    '         Tách GROUP-A và GROUP-B, sort theo SignDate trong m?i nhóm
    '==================================================================
    Dim mainA() As Long, cntA As Long
    Dim mainB() As Long, cntB As Long
    ReDim mainA(1 To total)
    ReDim mainB(1 To total)
    cntA = 0: cntB = 0

    For n = 1 To total
        If arrParent(n) = 0 Then
            If CStr(arrC(n, 4)) = "GROUP-A" Then
                cntA = cntA + 1
                mainA(cntA) = n
            Else
                cntB = cntB + 1
                mainB(cntB) = n
            End If
        End If
    Next n

    ' Sort mainA theo SignDate
    Call SortIndexByDate(arrC, mainA, cntA)
    ' Sort mainB theo SignDate
    Call SortIndexByDate(arrC, mainB, cntB)

    '==================================================================
    ' BU?C 5: Ghi sequence — GROUP-A tru?c, GROUP-B sau
    '==================================================================
    Dim sortNo   As Long
    Dim mainNoA  As Long
    Dim mainNoB  As Long
    sortNo = 0: mainNoA = 0: mainNoB = 0

    ' --- GROUP-A ---
    Dim g As Long
    For g = 1 To cntA
        Dim mi As Long
        mi = mainA(g)
        mainNoA = mainNoA + 1
        Dim sTT As String
        sTT = "A-" & mainNoA
        sortNo = sortNo + 1
        ws.Cells(CLng(arrC(mi, 2)), "X").value = sTT
        ws.Cells(CLng(arrC(mi, 2)), "Y").value = sortNo
        Call WriteChildren(ws, arrC, arrParent, total, mi, sTT, sortNo)
    Next g

    ' --- GROUP-B ---
    For g = 1 To cntB
        mi = mainB(g)
        mainNoB = mainNoB + 1
        sTT = "B-" & mainNoB
        sortNo = sortNo + 1
        ws.Cells(CLng(arrC(mi, 2)), "X").value = sTT
        ws.Cells(CLng(arrC(mi, 2)), "Y").value = sortNo
        Call WriteChildren(ws, arrC, arrParent, total, mi, sTT, sortNo)
    Next g

    Exit Sub

ERR_HANDLER:
    MsgBox "RebuildContractSequence: " & Err.Description, vbCritical

End Sub


'================================================================
' HELPER 1 — Ghi d? quy con c?p 1 và c?p 2
'================================================================
Private Sub WriteChildren( _
    ByVal ws As Worksheet, _
    ByRef arrC() As Variant, _
    ByRef arrParent() As Long, _
    ByVal total As Long, _
    ByVal parentIdx As Long, _
    ByVal parentStt As String, _
    ByRef sortNo As Long)

    ' L?y danh sách con tr?c ti?p
    Dim children() As Long
    Dim childCnt As Long
    ReDim children(1 To total)
    childCnt = 0

    Dim n As Long
    For n = 1 To total
        If arrParent(n) = parentIdx Then
            childCnt = childCnt + 1
            children(childCnt) = n
        End If
    Next n

    If childCnt = 0 Then Exit Sub

    ' Sort theo SignDate
    Call SortIndexByDate(arrC, children, childCnt)

    ' Ghi và d? quy
    Dim c As Long
    For c = 1 To childCnt
        Dim ci As Long
        ci = children(c)
        Dim childStt As String
        childStt = parentStt & "." & c
        sortNo = sortNo + 1
        ws.Cells(CLng(arrC(ci, 2)), "X").value = childStt
        ws.Cells(CLng(arrC(ci, 2)), "Y").value = sortNo
        ' Ð? quy cho con c?p sâu hon
        Call WriteChildren(ws, arrC, arrParent, total, ci, childStt, sortNo)
    Next c

End Sub


'================================================================
' HELPER 2 — Insertion sort m?ng index theo SignDate (c?t 3 trong arrC)
'================================================================
Private Sub SortIndexByDate( _
    ByRef arrC() As Variant, _
    ByRef idx() As Long, _
    ByVal cnt As Long)

    Dim i As Long, j As Long, tmp As Long
    For i = 2 To cnt
        tmp = idx(i)
        Dim keyDate As Variant
        keyDate = arrC(tmp, 3)
        j = i - 1
        Do While j >= 1
            Dim cmpDate As Variant
            cmpDate = arrC(idx(j), 3)
            ' So sánh ngày: dùng CDate n?u h?p l?, ngu?c l?i gi? nguyên th? t?
            Dim shouldSwap As Boolean
            shouldSwap = False
            If IsDate(cmpDate) And IsDate(keyDate) Then
                If CDate(cmpDate) > CDate(keyDate) Then shouldSwap = True
            ElseIf CStr(cmpDate) > CStr(keyDate) Then
                shouldSwap = True
            End If
            If shouldSwap Then
                idx(j + 1) = idx(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        idx(j + 1) = tmp
    Next i

End Sub



Public Function VN_NoSign(ByVal s As String) As String

    s = LCase$(Trim$(s))

    ' a
    s = Replace(s, ChrW(224), "a")
    s = Replace(s, ChrW(225), "a")
    s = Replace(s, ChrW(7841), "a")
    s = Replace(s, ChrW(7843), "a")
    s = Replace(s, ChrW(227), "a")

    s = Replace(s, ChrW(226), "a")
    s = Replace(s, ChrW(7847), "a")
    s = Replace(s, ChrW(7845), "a")
    s = Replace(s, ChrW(7853), "a")
    s = Replace(s, ChrW(7849), "a")
    s = Replace(s, ChrW(7851), "a")

    s = Replace(s, ChrW(259), "a")
    s = Replace(s, ChrW(7857), "a")
    s = Replace(s, ChrW(7855), "a")
    s = Replace(s, ChrW(7863), "a")
    s = Replace(s, ChrW(7859), "a")
    s = Replace(s, ChrW(7861), "a")

    ' e
    s = Replace(s, ChrW(232), "e")
    s = Replace(s, ChrW(233), "e")
    s = Replace(s, ChrW(7865), "e")
    s = Replace(s, ChrW(7867), "e")
    s = Replace(s, ChrW(7869), "e")

    s = Replace(s, ChrW(234), "e")
    s = Replace(s, ChrW(7873), "e")
    s = Replace(s, ChrW(7871), "e")
    s = Replace(s, ChrW(7879), "e")
    s = Replace(s, ChrW(7875), "e")
    s = Replace(s, ChrW(7877), "e")

    ' i
    s = Replace(s, ChrW(236), "i")
    s = Replace(s, ChrW(237), "i")
    s = Replace(s, ChrW(7883), "i")
    s = Replace(s, ChrW(7881), "i")
    s = Replace(s, ChrW(297), "i")

    ' o
    s = Replace(s, ChrW(242), "o")
    s = Replace(s, ChrW(243), "o")
    s = Replace(s, ChrW(7885), "o")
    s = Replace(s, ChrW(7887), "o")
    s = Replace(s, ChrW(245), "o")

    s = Replace(s, ChrW(244), "o")
    s = Replace(s, ChrW(7891), "o")
    s = Replace(s, ChrW(7889), "o")
    s = Replace(s, ChrW(7897), "o")
    s = Replace(s, ChrW(7893), "o")
    s = Replace(s, ChrW(7895), "o")

    s = Replace(s, ChrW(417), "o")
    s = Replace(s, ChrW(7901), "o")
    s = Replace(s, ChrW(7899), "o")
    s = Replace(s, ChrW(7907), "o")
    s = Replace(s, ChrW(7903), "o")
    s = Replace(s, ChrW(7905), "o")

    ' u
    s = Replace(s, ChrW(249), "u")
    s = Replace(s, ChrW(250), "u")
    s = Replace(s, ChrW(7909), "u")
    s = Replace(s, ChrW(7911), "u")
    s = Replace(s, ChrW(361), "u")

    s = Replace(s, ChrW(432), "u")
    s = Replace(s, ChrW(7915), "u")
    s = Replace(s, ChrW(7913), "u")
    s = Replace(s, ChrW(7921), "u")
    s = Replace(s, ChrW(7917), "u")
    s = Replace(s, ChrW(7919), "u")

    ' y
    s = Replace(s, ChrW(7923), "y")
    s = Replace(s, ChrW(253), "y")
    s = Replace(s, ChrW(7925), "y")
    s = Replace(s, ChrW(7927), "y")
    s = Replace(s, ChrW(7929), "y")

    ' d
    s = Replace(s, ChrW(273), "d")

    VN_NoSign = s

End Function

Public Function PickExcelFile(ByVal title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = title
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls"
        .AllowMultiSelect = False
        If mLastBrowseDir <> "" Then
            .InitialFileName = mLastBrowseDir & "\"
        Else
            .InitialFileName = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\"
        End If
        If .Show <> -1 Then Exit Function
        PickExcelFile = .SelectedItems(1)
        mLastBrowseDir = Left(.SelectedItems(1), InStrRev(.SelectedItems(1), "\") - 1)
    End With
End Function

