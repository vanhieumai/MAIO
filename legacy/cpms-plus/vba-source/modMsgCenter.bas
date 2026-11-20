Attribute VB_Name = "modMsgCenter"
'Attribute VB_Name = "modMsgCenter"
Option Explicit

'==============================================================================
' modMsgBoxCenter
' Wrapper dung frmMsg2 de thay the MsgBox / UniMsgBox trong toan bo project.
'
' API cong khai:
'   MsgCenter(msg, buttons, title) As Long   ' drop-in cho MsgBox
'   ShowMsgCenter msg, icon                  ' thong bao don gian (OK)
'==============================================================================

'------------------------------------------------------------------------------
' MsgCenter — thay the 1-1 cho MsgBox
'
' Vi du:
'   MsgCenter "Da luu!", vbInformation, "OK"
'   If MsgCenter("Xoa?", vbYesNo + vbQuestion, "Xac nhan") = vbYes Then
'   Select Case MsgCenter("Tiep tuc?", vbYesNoCancel, "")
'------------------------------------------------------------------------------
Public Function MsgCenter( _
    ByVal Prompt As String, _
    Optional ByVal Buttons As Long = vbOKOnly, _
    Optional ByVal title As String = "") As Long

    Dim f As frmMsg2
    Set f = New frmMsg2

    ' Chi lay phan button-type (3 bit thap), bo icon/default-button flags
    Dim btnType As Long
    btnType = Buttons And &HF

    Select Case btnType

        Case vbYesNo        ' 4
            If f.ShowQuestion(Prompt, title) Then
                MsgCenter = vbYes
            Else
                MsgCenter = vbNo
            End If

        Case vbYesNoCancel  ' 3
            Select Case f.ShowYesNoCancel(Prompt, title)
                Case "YES":    MsgCenter = vbYes
                Case "NO":     MsgCenter = vbNo
                Case Else:     MsgCenter = vbCancel
            End Select

        Case Else           ' vbOKOnly(0), vbOKCancel(1), vbRetryCancel(5)...
            f.ShowInfo Prompt, title
            MsgCenter = vbOK

    End Select

    Unload f
    Set f = Nothing
End Function

'------------------------------------------------------------------------------
' ShowMsgCenter — thong bao don gian, khong can gia tri tra ve
'
' Tuong duong: MsgBox "...", vbInformation / vbCritical / vbExclamation
'
' Vi du:
'   ShowMsgCenter "Da luu hop dong.", vbInformation
'   ShowMsgCenter "Khong tim thay du lieu.", vbCritical
'------------------------------------------------------------------------------
Public Sub ShowMsgCenter( _
    ByVal Prompt As String, _
    Optional ByVal Icon As Long = vbInformation, _
    Optional ByVal title As String = "")

    ' Xac dinh title mac dinh theo loai icon
    If title = "" Then
        Select Case Icon
            Case vbCritical:    title = "L" & ChrW(7895) & "i"
            Case vbExclamation: title = "C" & ChrW(7843) & "nh b" & ChrW(225) & "o"
            Case Else:          title = "Th" & ChrW(244) & "ng b" & ChrW(225) & "o"
        End Select
    End If

    Dim f As frmMsg2
    Set f = New frmMsg2
    f.ShowInfo Prompt, title
    Unload f
    Set f = Nothing
End Sub


