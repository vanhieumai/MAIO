Attribute VB_Name = "modLinkEngine"
Option Explicit

Public srcWorkbook As Workbook
Public srcWbName As String
Public srcWbPath As String
Public isUsingExternalSource As Boolean
Public defaultWorkbook As Workbook


Public Function Nz(v As Variant) As String
    If IsNull(v) Or v = "" Then
        Nz = ""
    Else
        Nz = CStr(v)
    End If
End Function



Public Function IsWorkbookOpen(wb As Workbook) As Boolean
    Dim w As Workbook
    If wb Is Nothing Then Exit Function

    For Each w In Application.Workbooks
        If w Is wb Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next w
End Function

Public Function IsSourceWorkbookValid() As Boolean

    Dim wb As Workbook

    If srcWbPath = "" Then
        IsSourceWorkbookValid = False
        Exit Function
    End If

    For Each wb In Application.Workbooks
        If wb.fullName = srcWbPath Then
            Set srcWorkbook = wb
            IsSourceWorkbookValid = True
            Exit Function
        End If
    Next wb

    '--- không còn t?n t?i
    Set srcWorkbook = Nothing
    srcWbName = ""
    srcWbPath = ""

    IsSourceWorkbookValid = False

End Function


Public Function wsColumnIndex(colRef As String) As Long
    ' Nh?n "A", "B", "AA", "$C$1", "C1" ? tr? v? s? c?t

    On Error GoTo ErrHandle
    wsColumnIndex = Range(colRef & "1").Column
    Exit Function

ErrHandle:
    wsColumnIndex = 0
End Function

'==============================
' TÌM WORKBOOK C?A USER
'
' H? tr? 2 ch? d?:
'
'   A) XLAM mode (ThisWorkbook.IsAddin = True):
'      Code n?m trong add-in (.xlam). Workbook c?a user là file
'      dang active (không ph?i addin). Uu tiên ActiveWorkbook,
'      fallback sang workbook d?u tiên không ph?i addin.
'      Tr? v? Nothing n?u không có file nào dang m?.
'
'   B) Direct-run mode (ThisWorkbook.IsAddin = False):
'      Code ch?y th?ng t? file .xlsm/.xlsb (ví d?: CPMS AR.xlsm).
'      Luôn tr? v? CHÍNH ThisWorkbook — không tìm ki?m workbook khác.
'
'      LÝ DO: Trong ch? d? này có th? t?n t?i workbook ?n (ví d?
'      CPMS_Data.xlsb) trong Application.Workbooks. Workbook ?n dó
'      là external source, du?c load riêng qua "Load External" ?
'      srcWorkbook. N?u loop qua Workbooks s? nh?t nh?m workbook ?n
'      làm defaultWorkbook. Vì v?y KHÔNG loop — tr? ThisWorkbook ngay.
'==============================
Public Function GetUserWorkbook() As Workbook
    Dim wb As Workbook

    ' =====================================================
    ' MODE A: Ch?y qua XLAM add-in
    ' ThisWorkbook là addin ? user workbook n?m ? file khác
    ' =====================================================
    If ThisWorkbook.IsAddin Then

        ' Uu tiên ActiveWorkbook n?u không ph?i addin
        If Not ActiveWorkbook Is Nothing Then
            If Not ActiveWorkbook.IsAddin Then
                Set GetUserWorkbook = ActiveWorkbook
                Exit Function
            End If
        End If

        ' Fallback: workbook d?u tiên không ph?i addin
        For Each wb In Application.Workbooks
            If Not wb.IsAddin Then
                Set GetUserWorkbook = wb
                Exit Function
            End If
        Next wb

        ' Không có file nào dang m? ? tr? Nothing
        Exit Function
    End If

    ' =====================================================
    ' MODE B: Ch?y tr?c ti?p t? file ch?a code (không ph?i XLAM)
    ' ? Luôn dùng ThisWorkbook, không tìm workbook khác.
    '   Workbook ?n (data file) là external source riêng bi?t.
    ' =====================================================
    Set GetUserWorkbook = ThisWorkbook

End Function

Public Function GetSourceWorkbook() As Workbook

    Dim wb As Workbook

    '--- n?u không có source ngoài ? internal (defaultWorkbook)
    If srcWbPath = "" Then
        Set GetSourceWorkbook = defaultWorkbook
        Exit Function
    End If

    '--- ki?m tra source còn m? không
    For Each wb In Application.Workbooks
        If wb.fullName = srcWbPath Then
            Set srcWorkbook = wb
            Set GetSourceWorkbook = wb
            Exit Function
        End If
    Next wb

    '--- source dã b? dóng
    Set srcWorkbook = Nothing
    srcWbName = ""
    srcWbPath = ""

    Set GetSourceWorkbook = Nothing

End Function


Public Function GetGroupByContractType(contractType As String) As String
    Select Case UCase(Trim(contractType))
        Case "MAIN", "EPC"
            GetGroupByContractType = "A"
        Case "SUB"
            GetGroupByContractType = "B"
        Case Else
            GetGroupByContractType = ""
    End Select
End Function

Public Sub EnsureGroupMarkers(ws As Worksheet)

    Dim lastRow   As Long
    Dim r         As Long
    Dim hasGroupA As Boolean
    Dim hasGroupB As Boolean

    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row

    '-- Quét tìm marker dã t?n t?i --
    For r = 2 To lastRow
        If Trim(ws.Cells(r, 1).value) = "GROUP-A" Then hasGroupA = True
        If Trim(ws.Cells(r, 1).value) = "GROUP-B" Then hasGroupB = True
    Next r

    '-- Ðã d? c? 2 ? thoát --
    If hasGroupA And hasGroupB Then Exit Sub

    If Not hasGroupA And Not hasGroupB Then
        '-- Sheet tr?ng (ch? có header) ? chèn 2 dòng sau header --
        ws.rows(2).Insert Shift:=xlDown
        ws.rows(2).Insert Shift:=xlDown
        ws.Cells(2, 1).value = "GROUP-A"
        ws.Cells(3, 1).value = "GROUP-B"

    ElseIf Not hasGroupA Then
        '-- Ch? thi?u GROUP-A ? chèn tru?c GROUP-B --
        Dim rB As Long
        For rB = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).Row
            If Trim(ws.Cells(rB, 1).value) = "GROUP-B" Then
                ws.rows(rB).Insert Shift:=xlDown
                ws.Cells(rB, 1).value = "GROUP-A"
                Exit For
            End If
        Next rB

    ElseIf Not hasGroupB Then
        '-- Ch? thi?u GROUP-B ? thêm ? cu?i --
        Dim lastUsed As Long
        lastUsed = ws.Cells(ws.rows.count, 1).End(xlUp).Row + 1
        ws.Cells(lastUsed, 1).value = "GROUP-B"

    End If

End Sub


'------------------------------------------------------------------------------
' FindInsertRowByGroup — GI? NGUYÊN logic g?c
' Ch? thêm 1 dòng: g?i EnsureGroupMarkers tru?c khi tìm
'------------------------------------------------------------------------------
Public Function FindInsertRowByGroup(ws As Worksheet, groupCode As String) As Long

    '-- Ð?m b?o marker t?n t?i (t?o n?u chua có) --
    EnsureGroupMarkers ws

    Dim lastRow As Long, r As Long
    Dim curID   As String
    Dim inGroup As Boolean

    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    FindInsertRowByGroup = 0
    inGroup = False

    For r = 2 To lastRow

        curID = Trim(ws.Cells(r, 1).value)

        If curID = "GROUP-" & groupCode Then
            inGroup = True
            FindInsertRowByGroup = r + 1
            GoTo NEXT_R
        End If

        If inGroup Then
            If curID Like "GROUP-*" Then Exit Function
            FindInsertRowByGroup = r + 1
        End If

NEXT_R:
    Next r

End Function


Public Function FindInsertRow(parentID As String) As Long

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim curID As String
    Dim foundParent As Boolean

    Set ws = EnsureContractDB
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row

    FindInsertRow = 0
    foundParent = False

    For i = 2 To lastRow

        curID = Trim(ws.Cells(i, 1).value)
        If curID = "" Then GoTo NEXT_I

        ' G?p dúng parent
        If curID = parentID Then
            foundParent = True
            FindInsertRow = i + 1
            GoTo NEXT_I
        End If

        ' Ðang trong subtree
        If foundParent Then
            If curID Like parentID & "-*" Then
                FindInsertRow = i + 1
            Else
                Exit Function
            End If
        End If

NEXT_I:
    Next i

End Function



Public Function GetRootContractID(contractID As String) As String
    If InStr(contractID, "-") > 0 Then
        GetRootContractID = Split(contractID, "-")(0)
    Else
        GetRootContractID = contractID
    End If
End Function

Public Function IsContractNoExists( _
    ws As Worksheet, _
    contractNo As String, _
    Optional excludeRow As Long = 0 _
) As Boolean

    Dim r As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).Row

    For r = 2 To lastRow
        If r <> excludeRow Then
            If Trim(ws.Cells(r, 2).value) = Trim(contractNo) Then
                IsContractNoExists = True
                Exit Function
            End If
        End If
    Next r

    IsContractNoExists = False
End Function

Public Function GetGroupByContractID(contractID As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = EnsureContractDB

    For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).Row
        If ws.Cells(r, 1).value Like "GROUP-*" Then
            GetGroupByContractID = ws.Cells(r, 1).value
        ElseIf ws.Cells(r, 1).value = contractID Then
            Exit Function
        End If
    Next r
End Function


