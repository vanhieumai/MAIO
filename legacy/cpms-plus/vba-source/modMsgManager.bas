Attribute VB_Name = "modMsgManager"
Option Explicit
Public MsgForms As Collection

'Public Sub CloseAllMsgForms()

'    Dim f As Object

'    If MsgForms Is Nothing Then Exit Sub

'    For Each f In MsgForms
'        On Error Resume Next
'        Unload f
'        On Error GoTo 0
'    Next f

'    Set MsgForms = Nothing

'End Sub

Public Sub ShowMsgCenter1(ByVal txt As String, Optional ByVal MsgType As VbMsgBoxStyle = vbInformation)
    
    If MsgForms Is Nothing Then
        Set MsgForms = New Collection
    End If
    
    Dim f As New frmMsg
        
    With f
        .MsgText = txt
        .MsgType = MsgType
        '.Show vbModeless
        .Show vbModal
    End With
    
    'MsgForms.Add f
    
    Set f = Nothing

End Sub

