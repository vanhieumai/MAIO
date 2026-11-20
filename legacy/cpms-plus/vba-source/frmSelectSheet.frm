Attribute VB_Name = "frmSelectSheet"
Attribute VB_Base = "0{3BE1BCC6-24A5-4A0A-9679-5B1671D4B75C}{99E9EFB7-51E2-4045-A9DD-859E6C35478F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public SelectedSheetName As String
Public IsCanceled As Boolean

Private Sub UserForm_Initialize()
    IsCanceled = True
    ApplyTheme Me
    'UnregisterForm Me
End Sub

Private Sub btnOK_Click()

    If lstSheet.ListIndex = -1 Then
        ShowMsgCenter "Please select a sheet!", vbExclamation
        Exit Sub
    End If

    SelectedSheetName = lstSheet.List(lstSheet.ListIndex)
    IsCanceled = False
    Me.Hide

End Sub

Private Sub btnCancel_Click()
    IsCanceled = True
    Me.Hide
End Sub

Private Sub lstSheet_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnOK_Click
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        'Cancel = True

        IsCanceled = True

        Unload Me
    End If
    'ApplyTheme Me
    UnregisterForm Me
End Sub

