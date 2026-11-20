Attribute VB_Name = "modFormat"


Public Sub FormatContractRowSingle(ws As Worksheet, r As Long)

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    Dim contractID As String
    Dim lvl As Long
    Dim idCol As Long

    ' ===== XÁC Ð?NH C?T CONTRACTID THEO C?U TRÚC =====
    ' N?u c?t C có giá tr? gi?ng ContractID (ch?a "-") ? dùng c?t C
    ' Ngu?c l?i ? dùng c?t A
    contractID = Trim(ws.Cells(1, "C").value)
    If contractID = "ContractID" Then
        idCol = 3   ' Payment Ledger
    Else
        idCol = 1   ' Contract DB
    End If

    contractID = Trim(ws.Cells(r, idCol).value)
    If contractID = "" Then Exit Sub
    If contractID Like "GROUP-*" Then Exit Sub

    lvl = GetContractLevel(contractID)
    If lvl <= 0 Then Exit Sub

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
        .Font.Bold = False
        .Interior.Color = GetLevelSoftBlue(lvl)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    'FormatRowBorderSingle ws, r

End Sub


Public Sub FormatPaymentRowSingle(ws As Worksheet, r As Long)

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
        .Font.Bold = False
        .Interior.Color = RGB(252, 245, 225) ' vàng nh?t
        
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlHairline
    End With
    'FormatRowBorderSingle ws, r
    
End Sub

Public Sub FormatRowBorderSingle(ws As Worksheet, r As Long)

    Dim lastCol As Long, lastRow As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    lastRow = ws.UsedRange.rows(ws.UsedRange.rows.count).Row

    Dim rowRange As Range
    Set rowRange = ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))

    ' ===== VI?N NGOÀI =====
    With rowRange
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

End Sub


Public Sub FormatLedgerRowSingle(ws As Worksheet, r As Long)

    If ws.Name <> "PAYMENT_LEDGER" Then Exit Sub

    Dim contractID As String
    Dim ledgerID As String

    contractID = Trim(ws.Cells(r, "C").value)
    ledgerID = Trim(ws.Cells(r, "B").value)

    ' ===== GROUP HEADER =====
    If contractID Like "GROUP-*" And ledgerID = "" Then
        'FormatGroupRowSingle ws, r
        Exit Sub
    End If

    ' ===== PAYMENT ROW =====
    If ledgerID <> "" Then
        FormatPaymentRowSingle ws, r
        Exit Sub
    End If

    ' ===== CONTRACT / APPENDIX HEADER =====
    If contractID <> "" And ledgerID = "" Then
        FormatContractRowSingle ws, r
        Exit Sub
    End If

End Sub

Public Sub WriteTextCell(c As Range, ByVal v As String)
    c.NumberFormat = "@"
    c.value = CStr(v)
End Sub

Public Sub FormatTableWithGroup(ws As Worksheet)

    Dim lastRow As Long, lastCol As Long
    Dim r As Long
    Dim rng As Range, rowRange As Range

    If ws Is Nothing Then Exit Sub

    lastRow = ws.UsedRange.rows(ws.UsedRange.rows.count).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    
  
    
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' ===== VIEW =====
    rng.WrapText = False
    ws.Columns.AutoFit

    ' ===== DEFAULT BORDER (DATA AREA ONLY) =====
    With rng
        '.Borders(xlEdgeLeft).LineStyle = xlContinuous
        '.Borders(xlEdgeRight).LineStyle = xlContinuous
        '.Borders(xlEdgeTop).LineStyle = xlContinuous
        '.Borders(xlEdgeBottom).LineStyle = xlContinuous

        '.Borders(xlInsideVertical).LineStyle = xlContinuous
        '.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Font.Bold = False
    End With

    ' ===== TITLE ROW =====
    Set rowRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    With rowRange
        .Font.Bold = True
        .Interior.Color = RGB(230, 230, 250) ' Lavender 80%
        '.Borders.LineStyle = xlContinuous
        
        '.Borders(xlEdgeLeft).Weight = xlThin
        '.Borders(xlEdgeRight).Weight = xlThin
        '.Borders(xlInsideVertical).Weight = xlThin
        '.Borders(xlInsideHorizontal).Weight = xlThin
    End With

    ' ===== GROUP ROW =====
    For r = 2 To lastRow

        If IsGroupRow(ws, r) Then
            Set rowRange = ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
            FormatGroupRow rowRange
        End If

    Next r
    
    ' ===== CONTRACT LEVEL COLOR =====
    Dim idCol As Long

    ' Detect ContractID column
    If ws.Name = "PAYMENT_LEDGER" Then
        idCol = 3   ' Ledger
    Else
        idCol = 1   ' Contract DB
    End If

    For r = 2 To lastRow

        ' Skip group
        If IsGroupRow(ws, r) Then GoTo NEXT_R

        ' Skip empty
        If Trim(ws.Cells(r, idCol).value) = "" Then GoTo NEXT_R

        Dim lvl As Long
        lvl = GetContractLevel(ws.Cells(r, idCol).value)

        If lvl > 0 Then
            Set rowRange = ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
            rowRange.Interior.Color = GetLevelSoftBlue(lvl)
        End If

NEXT_R:
    Next r

    ' ===== PAYMENT LEDGER ROW COLOR =====
    If ws.Name = "PAYMENT_LEDGER" Then

        For r = 2 To lastRow

            ' Có LedgerID => Payment row
            If Trim(ws.Cells(r, "B").value) <> "" Then

                Set rowRange = ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
                rowRange.Interior.Color = RGB(252, 245, 225) ' vàng nh?t
                rowRange.Borders.LineStyle = xlContinuous
                'rowRange.Borders.Weight = xlHairline
                rowRange.Borders(xlEdgeLeft).Weight = xlThin
                rowRange.Borders(xlEdgeRight).Weight = xlThin
                rowRange.Borders(xlEdgeTop).Weight = xlHairline
                rowRange.Borders(xlEdgeBottom).Weight = xlThin
                'rowRange.Borders(xlInsideVertical).Weight = xlThin
                'rowRange.Borders(xlInsideHorizontal).Weight = xlHairline
            End If

        Next r

    End If


End Sub
