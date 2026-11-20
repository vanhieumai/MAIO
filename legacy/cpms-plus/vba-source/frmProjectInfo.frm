Attribute VB_Name = "frmProjectInfo"
Attribute VB_Base = "0{CB3AE3CC-0E05-4DC2-864D-BCF0B9572242}{1B5B6521-B924-40A5-800B-1EB8FBB9E60E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'==========================================================
' frmProjectInfo
'==========================================================

Private m_Cancelled As Boolean
Private m_ProjName  As String
Private m_Location  As String

'==========================================================
' PUBLIC INTERFACE
'==========================================================
Public Sub Init(ByVal DefaultName As String, _
                ByVal DefaultLocation As String)

    Me.Caption = "New Project"
    m_Cancelled = True
    txtName.value = DefaultName
    txtLocation.value = DefaultLocation
    txtName.SetFocus   ' <-- SetFocus SAU khi set value, khong phai trong Initialize

End Sub

Public Sub InitEdit(ByVal CurrentName As String, _
                    ByVal CurrentLocation As String)

    Me.Caption = "Edit Project"
    m_Cancelled = True
    txtName.value = CurrentName
    txtLocation.value = CurrentLocation
    txtName.SetFocus   ' <-- SetFocus SAU khi set value

End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property

Public Property Get projName() As String
    projName = m_ProjName
End Property

Public Property Get location() As String
    location = m_Location
End Property

'==========================================================
' INITIALIZE — chi khoi tao bien, KHONG SetFocus o day
' (SetFocus trong Initialize gay flash vi form chua co data)
'==========================================================
Private Sub UserForm_Initialize()
    m_Cancelled = True
    ' KHONG goi txtName.SetFocus o day
    ApplyTheme Me
    'UnregisterForm Me
    ' SetFocus duoc goi o cuoi Init / InitEdit sau khi da set value
End Sub

'==========================================================
' OK
'==========================================================
Private Sub btnOK_Click()

    Dim sName As String
    Dim sLoc  As String

    sName = Trim(txtName.value)
    sLoc = Trim(txtLocation.value)

    If Len(sName) = 0 Then
        MsgBox "Project Name is required.", vbExclamation
        txtName.SetFocus
        Exit Sub
    End If

    m_ProjName = sName
    m_Location = sLoc
    m_Cancelled = False

    Me.Hide

End Sub

'==========================================================
' CANCEL
'==========================================================
Private Sub btnCancel_Click()
    m_Cancelled = True
    Me.Hide
End Sub

'==========================================================
' ENTER = OK, ESC = Cancel
'==========================================================
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                              ByVal Shift As Integer)
    Select Case KeyCode
        Case 13: btnOK_Click
        Case 27: btnCancel_Click
    End Select
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'ApplyTheme Me
    UnregisterForm Me
End Sub
