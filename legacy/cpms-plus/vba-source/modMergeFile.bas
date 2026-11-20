Attribute VB_Name = "modMergeFile"
Option Explicit

' ==============================


' ==============================
' CHECK FOR DUPLICATE SHEET NAMES
' ==============================
Public Function HasDuplicateSheets( _
        wbTarget As Workbook, _
        wbSource As Workbook, _
        lbSheets As MSForms.ListBox, _
        dupList As String) As Boolean

    Dim i As Long
    Dim shName As String
    dupList = ""

    For i = 0 To lbSheets.ListCount - 1
        If lbSheets.Selected(i) Then
            shName = lbSheets.List(i)
            If SheetExists(shName, wbTarget) Then
                dupList = dupList & "  " & shName & vbCrLf
            End If
        End If
    Next i

    HasDuplicateSheets = (dupList <> "")

End Function


Public Function SheetExists(sName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function




