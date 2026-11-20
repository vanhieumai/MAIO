Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'================================================================
' PASTE VÀO MODULE "ThisWorkbook" (double-click "ThisWorkbook"
' trong c?a s? VBA Project Explorer)
'
' Tác d?ng: Ba sheet CONTRACT_DB, PAYMENT_LEDGER, DOCUMENT_TRACKER
'           s? t? ?n hoàn toàn (xlSheetVeryHidden) m?i khi ngu?i
'           dùng chuy?n sang sheet khác.
'           Các hàm Ensure* s? t? b? ?n tru?c khi truy c?p data.
'================================================================

Option Explicit


'----------------------------------------------------------------
' ?n t?t c? 3 sheet data ngay l?p t?c (dùng khi m? file)
'----------------------------------------------------------------
Private Sub HideAllDataSheets()

    Dim sheetNames As Variant
    Dim sName As Variant
    Dim ws As Worksheet

    sheetNames = Array("CONTRACT_DB", "PAYMENT_LEDGER", "DOCUMENT_TRACKER")

    For Each sName In sheetNames
        On Error Resume Next
        Set ws = Me.Worksheets(CStr(sName))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ' Ch? ?n n?u dây không ph?i sheet dang active
            If ws.Name <> Me.ActiveSheet.Name Then
                ws.Visible = xlSheetVeryHidden
            End If
        End If
        Set ws = Nothing
    Next sName

End Sub


