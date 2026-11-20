VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQSMSourcePicker 
   Caption         =   "UserForm1"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12765
   OleObjectBlob   =   "frmQSMSourcePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmQSMSourcePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCandidates As Collection
Private mSelected As Workbook
Public Sub Configure(ByVal candidates As Collection)
    Dim wb As Workbook, i As Long
    Set mCandidates = candidates
    Set mSelected = Nothing
    lstSources.Clear
    For i = 1 To candidates.count
        Set wb = candidates(i)
        lstSources.AddItem wb.Name
        If Len(wb.path) = 0 Then
            lstSources.List(i - 1, 1) = "(Unsaved workbook)"
        Else
            lstSources.List(i - 1, 1) = wb.fullName
        End If
    Next i
    If lstSources.ListCount > 0 Then lstSources.ListIndex = 0
End Sub
Public Property Get SelectedWorkbook() As Workbook
    Set SelectedWorkbook = mSelected
End Property
Private Sub cmdSelect_Click()
    If lstSources.ListIndex < 0 Then Exit Sub
    Dim candidate As Workbook, wb As Workbook
    Set candidate = mCandidates(lstSources.ListIndex + 1)
    For Each wb In Application.Workbooks
        If wb Is candidate Then
            Set mSelected = wb
            Me.Hide
            Exit Sub
        End If
    Next wb
    MsgBox "This workbook has been closed. Please reopen the source list.", vbExclamation
End Sub
Private Sub lstSources_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdSelect_Click
End Sub
Private Sub cmdCancel_Click()
    Set mSelected = Nothing
    Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        cmdCancel_Click
    End If
End Sub
