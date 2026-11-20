Option Explicit
Private mSource As Worksheet, mTarget As Worksheet
Private mSourceKey As Long, mTargetKey As Long
Private mSourceCols As Variant, mTargetCols As Variant
Private mRows As Variant, mKeys As Variant
Private mLast As Long

Public Sub Configure(ByVal source As Worksheet, ByVal target As Worksheet, ByVal sourceKey As Long, ByVal targetKey As Long, ByVal sourceCols As String, ByVal targetCols As String)
    Set mSource = source
    Set mTarget = target
    mSourceKey = sourceKey
    mTargetKey = targetKey
    mSourceCols = Split(sourceCols, ",")
    mTargetCols = Split(targetCols, ",")
    Validate
End Sub

Private Function WorkbookOpen(ByVal candidate As Workbook) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb Is candidate Then WorkbookOpen = True: Exit Function
    Next wb
End Function

Public Sub Validate()
    Dim i As Long, col As Long
    If mSource Is Nothing Or mTarget Is Nothing Then Err.Raise 5, , "Missing worksheet."
    If Not WorkbookOpen(mSource.parent) Or Not WorkbookOpen(mTarget.parent) Then Err.Raise 5, , "Source or target workbook has been closed. Reopen Preview."
    If mSourceKey < 1 Or mSourceKey > mSource.columns.count Or mTargetKey < 1 Or mTargetKey > mTarget.columns.count Then Err.Raise 5, , "Invalid matching column."
    If UBound(mSourceCols) <> UBound(mTargetCols) Then Err.Raise 5, , "Source and target link column counts differ."
    For i = 0 To UBound(mSourceCols)
        col = CLng(mSourceCols(i))
        If col < 1 Or col > mSource.columns.count Then Err.Raise 5, , "Invalid source column."
        col = CLng(mTargetCols(i))
        If col < 1 Or col > mTarget.columns.count Then Err.Raise 5, , "Invalid target column."
    Next i
End Sub

Private Function GetKeyColumnLastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long, lastR As Long
    On Error Resume Next
    lastR = ws.Cells(ws.rows.count, col).End(xlUp).Row

    ' Nếu cột bị merge với cột khác (ví dụ merge ngang B:C), End(xlUp) có thể bị sót dòng.
    ' Quét ngược từ UsedRange để tìm dòng cuối thực tế có dữ liệu trong ô hoặc merge area.
    Dim usedLast As Long
    If Not ws.UsedRange Is Nothing Then
        usedLast = ws.UsedRange.Row + ws.UsedRange.rows.count - 1
        If usedLast > lastR Then
            For r = usedLast To lastR + 1 Step -1
                Dim c As Range
                Set c = ws.Cells(r, col)
                If c.MergeCells Then
                    If Not IsEmpty(c.mergeArea.Cells(1, 1).value) Then
                        lastR = r
                        Exit For
                    End If
                ElseIf Not IsEmpty(c.value) Then
                    lastR = r
                    Exit For
                End If
            Next r
        End If
    End If
    If lastR < 1 Then lastR = 1
    GetKeyColumnLastRow = lastR
End Function

Public Property Get LastTargetRow() As Long
    Validate
    LastTargetRow = GetKeyColumnLastRow(mTarget, mTargetKey)
End Property

Public Property Get Description() As String
    Description = mSource.parent.Name & " / " & mSource.Name & " -> " & mTarget.parent.Name & " / " & mTarget.Name
End Property

Public Property Get MappingDescription() As String
    Dim i As Long, result As String
    For i = 0 To UBound(mSourceCols)
        If i > 0 Then result = result & "; "
        result = result & mSource.Cells(1, CLng(mSourceCols(i))).Address(False, False) & " -> " & mTarget.Cells(1, CLng(mTargetCols(i))).Address(False, False)
    Next i
    MappingDescription = result
End Property

Private Function KeyText(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    KeyText = Trim$(CStr(v))
End Function

Private Function ReadKeyColumn(ByVal ws As Worksheet, _
                              ByVal col As Long, _
                              ByVal lastRow As Long) As Variant
    Dim data As Variant
    Dim r As Long
    Dim c As Range
    Dim val As Variant
    Dim rngCheck As Range
    Dim hasMerge As Variant

    If lastRow < 1 Then lastRow = 1

    If lastRow = 1 Then
        Dim oneCellData(1 To 1, 1 To 1) As Variant
        Set c = ws.Cells(1, col)
        If c.MergeCells Then
            oneCellData(1, 1) = c.mergeArea.Cells(1, 1).value
        Else
            oneCellData(1, 1) = c.value
        End If
        ReadKeyColumn = oneCellData
        Exit Function
    End If

    Set rngCheck = ws.Range(ws.Cells(1, col), ws.Cells(lastRow, col))
    data = rngCheck.value

    ' Nếu dải ô có chứa ô merge (ngang hoặc dọc), lấy giá trị từ ô gốc top-left của merge area
    hasMerge = rngCheck.MergeCells
    If IsNull(hasMerge) Or hasMerge = True Then
        For r = 1 To lastRow
            Set c = ws.Cells(r, col)
            If c.MergeCells Then
                val = c.mergeArea.Cells(1, 1).value
                If Not IsEmpty(val) Then
                    data(r, 1) = val
                End If
            End If
        Next r
    End If

    ReadKeyColumn = data
End Function

Private Sub BuildMatches(ByVal firstRow As Long, ByVal lastRow As Long)
    Validate
    If firstRow < 1 Or lastRow < firstRow Or lastRow > LastTargetRow Then Err.Raise 5, , "Invalid target row bounds."
    Dim sourceData As Variant, targetData As Variant, lastSource As Long
    Dim dict As Object, used As Object, bucket As Collection
    Dim r As Long, key As String, matched() As Long
    lastSource = GetKeyColumnLastRow(mSource, mSourceKey)
    sourceData = ReadKeyColumn(mSource, mSourceKey, lastSource)
    targetData = ReadKeyColumn(mTarget, mTargetKey, lastRow)
    Set dict = CreateObject("Scripting.Dictionary")
    Set used = CreateObject("Scripting.Dictionary")
    ' Binary comparison preserves the original Link Data matching rule.
    For r = 1 To lastSource
        key = KeyText(sourceData(r, 1))
        If key <> "" Then
            If Not dict.exists(key) Then
                Set bucket = New Collection
                dict.Add key, bucket
            End If
            dict(key).Add r
        End If
    Next r
    ReDim matched(1 To lastRow)
    ' Count occurrences before firstRow as well: narrowing the range must not re-pair duplicates.
    For r = 1 To lastRow
        key = KeyText(targetData(r, 1))
        If key <> "" Then
            If dict.exists(key) Then
                If Not used.exists(key) Then used.Add key, CLng(0)
                used(key) = CLng(used(key)) + 1
                Set bucket = dict(key)
                If used(key) <= bucket.count Then matched(r) = bucket(used(key))
            End If
        End If
    Next r
    mRows = matched
    mKeys = targetData
    mLast = lastRow
End Sub

Public Sub FillPreview(ByVal lb As Object, ByVal firstRow As Long, ByVal lastRow As Long)
    BuildMatches firstRow, lastRow
    Dim result() As Variant, r As Long, n As Long, j As Long, idx As Long, v As Variant
    For r = firstRow To lastRow
        If KeyText(mKeys(r, 1)) <> "" Then n = n + 1
    Next r
    lb.Clear
    lb.ColumnCount = 8
    lb.ColumnWidths = "38 pt;180 pt;65 pt;65 pt;65 pt;65 pt;70 pt;0 pt"
    If n = 0 Then Exit Sub
    ReDim result(0 To n - 1, 0 To 7)
    For r = firstRow To lastRow
        If KeyText(mKeys(r, 1)) <> "" Then
            result(idx, 0) = r
            result(idx, 1) = KeyText(mKeys(r, 1))
            For j = 0 To 3
                If j <= UBound(mSourceCols) And mRows(r) > 0 Then
                    Dim srcCell As Range
                    Set srcCell = mSource.Cells(mRows(r), CLng(mSourceCols(j)))
                    If srcCell.MergeCells Then Set srcCell = srcCell.mergeArea.Cells(1, 1)
                    v = srcCell.value
                    If Not IsError(v) Then
                        If IsNumeric(v) And Not IsEmpty(v) Then
                            result(idx, j + 2) = Format$(CDbl(v), "#,##0.00")
                        Else
                            result(idx, j + 2) = KeyText(v)
                        End If
                    End If
                End If
            Next j
            result(idx, 6) = IIf(mRows(r) > 0, "OK", "Not Found")
            result(idx, 7) = mRows(r)
            idx = idx + 1
        End If
    Next r
    lb.List = result
End Sub

Public Sub NavigateSource(ByVal sourceRow As Long)
    Validate
    If sourceRow < 1 Then Err.Raise 5, , "No source cell for this row."
    Application.GoTo mSource.Cells(sourceRow, CLng(mSourceCols(0))), True
End Sub

Public Function Execute(ByVal firstRow As Long, ByVal lastRow As Long) As Boolean
    Dim oldEvents As Boolean, oldScreen As Boolean, oldCalc As XlCalculation
    Dim stateSet As Boolean, written As Long, failure As String
    On Error GoTo Failed
    BuildMatches firstRow, lastRow
    If mTarget.parent.ReadOnly Or mTarget.ProtectContents Then Err.Raise 5, , "Target is read-only or protected."
    Dim j As Long, k As Long, r As Long, targetRange As Range, overlap As Range, merged As Variant
    Dim targetCol As Long, targetCell As Range, hasMergeTarget As Boolean
    Dim formulaVal As Variant

    For j = 0 To UBound(mTargetCols)
        targetCol = CLng(mTargetCols(j))
        If mSource Is mTarget Then
            If targetCol = mSourceKey Then Err.Raise 5, , "Target would overwrite the source matching column."
            For k = 0 To UBound(mSourceCols)
                If targetCol = CLng(mSourceCols(k)) Then Err.Raise 5, , "Source and target link columns overlap on the same sheet."
            Next k
        End If
    Next j

    If MsgBox("Link data to target rows " & firstRow & " - " & lastRow & " (inclusive)?" & vbCrLf & _
        Description & vbCrLf & MappingDescription & vbCrLf & vbCrLf & _
        "Existing contents in the selected target columns and rows will be replaced. Unmatched or empty source values become blank, as in Link Data. Other rows are unchanged.", _
        vbYesNo + vbQuestion + vbDefaultButton2, "Link Data") <> vbYes Then Exit Function

    Dim outputs As Collection, outArr() As Variant, output As Variant, cell As Range, value As Variant, reference As String
    Set outputs = New Collection
    For j = 0 To UBound(mTargetCols)
        ReDim outArr(1 To lastRow - firstRow + 1, 1 To 1)
        For r = firstRow To lastRow
            If mRows(r) > 0 Then
                Set cell = mSource.Cells(mRows(r), CLng(mSourceCols(j)))
                If cell.MergeCells Then Set cell = cell.mergeArea.Cells(1, 1)
                value = cell.value
                If KeyText(value) <> "" Then
                    If mSource.parent Is mTarget.parent Then
                        reference = mSource.Name
                    Else
                        reference = "[" & mSource.parent.Name & "]" & mSource.Name
                    End If
                    outArr(r - firstRow + 1, 1) = "='" & Replace(reference, "'", "''") & "'!" & cell.Address(False, False)
                End If
            End If
        Next r
        outputs.Add outArr
    Next j

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    stateSet = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For j = 0 To UBound(mTargetCols)
        targetCol = CLng(mTargetCols(j))
        Set targetRange = mTarget.Range(mTarget.Cells(firstRow, targetCol), mTarget.Cells(lastRow, targetCol))
        merged = targetRange.MergeCells
        hasMergeTarget = (IsNull(merged) Or (merged = True))

        output = outputs(j + 1)

        If Not hasMergeTarget Then
            targetRange.Formula = output
        Else
            For r = firstRow To lastRow
                Set targetCell = mTarget.Cells(r, targetCol)
                If mRows(r) > 0 Then
                    formulaVal = output(r - firstRow + 1, 1)
                    If Not IsEmpty(formulaVal) And formulaVal <> "" Then
                        If targetCell.MergeCells Then
                            targetCell.mergeArea.Cells(1, 1).Formula = formulaVal
                        Else
                            targetCell.Formula = formulaVal
                        End If
                    Else
                        If Not targetCell.MergeCells Then
                            targetCell.ClearContents
                        End If
                    End If
                Else
                    If Not targetCell.MergeCells Then
                        targetCell.ClearContents
                    End If
                End If
            Next r
        End If
        written = written + 1
    Next j

    Execute = True

Clean:
    If stateSet Then
        On Error Resume Next
        Application.Calculation = oldCalc
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
        On Error GoTo 0
    End If
    If failure <> "" Then MsgBox "Link stopped. Columns already written: " & written & vbCrLf & failure, vbExclamation, "Link Data"
    If Execute Then MsgBox "Link completed. Target rows " & firstRow & " - " & lastRow & ".", vbInformation, "Link Data"
    Exit Function

Failed:
    failure = Err.Description
    Resume Clean
End Function
