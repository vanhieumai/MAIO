Attribute VB_Name = "frmPreviewLink"
Attribute VB_Base = "0{D520175D-4539-4B30-8A57-6541627973DC}{506D057A-D962-49F6-A340-6C1E4E708A52}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private arrSrcSheet() As String
Private arrSrcAddress() As String
Private srcWorkbook As Workbook

Public Sub LoadPreview(arrData As Variant)

    Dim i As Long

    lbPreview.Clear

    If IsEmpty(arrData) Then Exit Sub

    For i = LBound(arrData) To UBound(arrData)
        lbPreview.AddItem arrData(i)
    Next i

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ApplyTheme Me
    'UnregisterForm Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        'Cancel = True
        Unload Me
    End If
    'ApplyTheme Me
    UnregisterForm Me
End Sub

Private Sub lbPreview_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim idx As Long
    idx = lbPreview.ListIndex

    If idx < 0 Then Exit Sub

    If arrSrcSheet(idx) = "" Or arrSrcAddress(idx) = "" Then
        ShowMsgCenter "No source cell for this row"
        Exit Sub
    End If

    On Error Resume Next
    srcWorkbook.Activate
    srcWorkbook.Worksheets(arrSrcSheet(idx)).Activate
    srcWorkbook.Worksheets(arrSrcSheet(idx)).Range(arrSrcAddress(idx)).Select
    On Error GoTo 0

    'Unload Me   ' n?u mu?n dóng preview sau khi nh?y

End Sub

Public Sub SetSourceWorkbook(wb As Workbook)
    Set srcWorkbook = wb
End Sub

Public Sub SetPreviewSource(shArr() As String, addrArr() As String)
    arrSrcSheet = shArr
    arrSrcAddress = addrArr
End Sub

