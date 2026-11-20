Attribute VB_Name = "frmMsg2"
Attribute VB_Base = "0{DA41F5C2-C3D9-431D-82A8-93CD772BC60D}{BEC03803-7DBF-475E-8C5B-A9AA2F14EEE1}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private mResult As Boolean
Private mResultStr As String   ' "YES" | "NO" | "CANCEL"
Private mDone As Boolean       ' MOI: co bao "nguoi dung da bam xong"
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
#If VBA7 Then
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
#Else
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
#End If

#If VBA7 Then
    Private mCallerHwnd As LongPtr
#Else
    Private mCallerHwnd As Long
#End If

' ==========================================
' GOI TU BEN NGOAI
' ==========================================

' Thay MsgBox vbYesNoCancel — tr? v? "YES", "NO", ho?c "CANCEL"
Public Function ShowYesNoCancel(ByVal msg As String, Optional ByVal title As String = "") As String
    Dim ttl As String
    If title = "" Then ttl = "X" & ChrW(225) & "c nh" & ChrW(7853) & "n" Else ttl = title
    lblTitle.Caption = ttl
    lblMessage.value = vbCrLf & msg
    btnOK.Enabled = False
    btnYes.Enabled = True
    btnNo.Enabled = True
    btnCancel.Enabled = True
    mResultStr = "CANCEL"
    mDone = False
    CenterOnExcel
    
    mCallerHwnd = GetForegroundWindow()
    Me.Show vbModeless

    ' MOI: cho den khi nguoi dung bam 1 trong cac nut
    Do While Not mDone
        DoEvents
        Sleep 10
    Loop

    ShowYesNoCancel = mResultStr
End Function

' Thay MsgBox thông thu?ng (ch? có OK)
Public Sub ShowInfo(ByVal msg As String, Optional ByVal title As String = "")
    Dim ttl As String
    If title = "" Then ttl = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o" Else ttl = title
    lblTitle.Caption = ttl
    lblMessage.value = vbCrLf & msg
    btnOK.Enabled = True
    btnYes.Enabled = False
    btnNo.Enabled = False
    btnCancel.Enabled = False
    mDone = False
    CenterOnExcel
    
    mCallerHwnd = GetForegroundWindow()
    'Me.Show vbModal
    Me.Show vbModeless
    ' MOI: cho den khi nguoi dung bam OK
    Do While Not mDone
        DoEvents
        Sleep 10
    Loop

End Sub

' Thay MsgBox vbYesNo — tr? v? True n?u ch?n Có
Public Function ShowQuestion(ByVal msg As String, Optional ByVal title As String = "") As Boolean
    Dim ttl As String
    If title = "" Then ttl = "X" & ChrW(225) & "c nh" & ChrW(7853) & "n" Else ttl = title
    lblTitle.Caption = ttl
    lblMessage.value = vbCrLf & msg
    btnOK.Enabled = False
    btnYes.Enabled = True
    btnNo.Enabled = True
    btnCancel.Enabled = False
    mResult = False
    mDone = False
    CenterOnExcel
    Me.Show vbModeless

    ' MOI: cho den khi nguoi dung bam Yes/No
    Do While Not mDone
        DoEvents
        Sleep 10
    Loop

    ShowQuestion = mResult
End Function

' ==========================================
' XU LY NUT
' ==========================================
Private Sub btnOK_Click()
    mDone = True    ' MOI
    CloseAndReturnFocus
End Sub

Private Sub btnYes_Click()
    mResult = True
    mResultStr = "YES"
    mDone = True    ' MOI
    CloseAndReturnFocus
End Sub

Private Sub btnNo_Click()
    mResult = False
    mResultStr = "NO"
    mDone = True    ' MOI
    CloseAndReturnFocus
End Sub

Private Sub btnCancel_Click()
    mResult = False
    mResultStr = "CANCEL"
    mDone = True    ' MOI
    CloseAndReturnFocus
End Sub

' ==========================================
' CANG GIUA EXCEL
' ==========================================
Private Sub CenterOnExcel()
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub
   
Private Sub UserForm_Initialize()
    ApplyTheme Me
    'UnregisterForm Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
       If Not mDone Then
           mResult = False
           mResultStr = "CANCEL"
           mDone = True
       End If
    'ApplyTheme Me
    UnregisterForm Me
End Sub

Private Sub CloseAndReturnFocus()
    Unload Me                              ' huy han form
    On Error Resume Next
    If mCallerHwnd <> 0 Then
        SetForegroundWindow mCallerHwnd    ' MOI: tra focus DUNG ve form da goi minh
    Else
        SetForegroundWindow Application.hwnd   ' du phong neu khong lay duoc handle
    End If
    On Error GoTo 0
End Sub

