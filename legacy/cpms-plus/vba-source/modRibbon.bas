Attribute VB_Name = "modRibbon"
Option Explicit

Private mStandaloneReminder As frmReminder

'==========================================================
' Unicode Helper
'==========================================================
Public Function Uni(ByVal s As String) As String

    Dim i As Long
    Dim tmp As String

    For i = 1 To Len(s)
        tmp = tmp & "ChrW(" & AscW(mID$(s, i, 1)) & ") & "
    Next i

    Uni = Left$(tmp, Len(tmp) - 3)

End Function

'==========================================================
' LICENSE
'==========================================================
Public Sub RibbonLicense(Control As Object)

    On Error GoTo ErrHandle

    frmLicense.Show vbModeless
    'LinkKL.MultiPage1.value = 1

    Exit Sub

ErrHandle:

    MsgBox "Unable to open License." & vbCrLf & Err.Description, vbCritical

End Sub

'==========================================================
' QMS
'==========================================================
Public Sub RibbonQMS(Control As Object)

    On Error GoTo ErrHandle

    frmQSMToolkiT.Show vbModeless
    'LinkKL.MultiPage1.value = 0

    Exit Sub

ErrHandle:

    MsgBox "Unable to open QMS." & vbCrLf & Err.Description, vbCritical

End Sub

'==========================================================
' ASK AI
'==========================================================
Public Sub RibbonAskAI(Control As Object)

    On Error GoTo ErrHandle
    If Not EnsureCPMSSystem() Then Exit Sub

    If Not EnsureAISystem() Then Exit Sub

    frmAskAI.Show vbModeless

    Exit Sub

ErrHandle:
    MsgBox "Unable to open Ask AI." & vbCrLf & Err.Description, vbCritical

End Sub

'==========================================================
' REMINDER
'==========================================================
Public Sub RibbonReminder(Control As Object)

    On Error GoTo ErrHandle

    If Not EnsureCPMSSystem() Then Exit Sub

    ' Ki?m tra có project dang m? không tru?c khi làm b?t c? di?u gì
    Dim DataFile As String
    DataFile = ReadCurrentProjectValue("DataFile")

    If Len(DataFile) = 0 Or Dir(DataFile) = "" Then
        ShowMsgCenter "No projects found." & vbCrLf & vbCrLf & "Please create or import a project first.", vbExclamation, "Reminder"
        Exit Sub
    End If

    On Error Resume Next
    If Not gContractLoaded Then BuildGlobalContractCache
    If Not gPaymentLoaded Then BuildPaymentCache
    On Error GoTo 0

    If frmContractSearch.Visible Then
        frmContractSearch.OpenReminderForm
        Exit Sub
    End If

    If Not mStandaloneReminder Is Nothing Then
        On Error Resume Next
        mStandaloneReminder.LoadPendingList
        mStandaloneReminder.Show vbModeless
        On Error GoTo 0
        Exit Sub
    End If

    Set mStandaloneReminder = New frmReminder
    mStandaloneReminder.Show vbModeless

    Exit Sub

ErrHandle:
    MsgBox "Unable to open Reminder." & vbCrLf & Err.Description, vbCritical

End Sub
'==========================================================
' Callback t? frmReminder khi standalone instance b? Unload
'==========================================================
Public Sub ClearStandaloneReminderRef(Optional dummy As Variant = 0)
    Set mStandaloneReminder = Nothing
End Sub
'==========================================================
' CPMS
'
' Open current project
' If no project -> create new project
'==========================================================
Public Sub RibbonCPMS(Control As Object)

    On Error GoTo ErrHandle

    If Not EnsureCPMSSystem() Then Exit Sub

    frmHomePage.Show vbModeless

    Exit Sub

ErrHandle:
    MsgBox "Unable to start CPMS." & vbCrLf & Err.Description, vbCritical

End Sub

'==========================================================
' CURRENT PROJECT
'
' Same behaviour as CPMS
'==========================================================
Public Sub RibbonCurrent(Control As Object)

    On Error GoTo ErrHandle

    If Not EnsureCPMSSystem() Then Exit Sub

    OpenCurrentProject

    Exit Sub

ErrHandle:

    MsgBox Err.Description, vbCritical

End Sub

'==========================================================
' CREATE NEW PROJECT
'==========================================================
Public Sub RibbonCreateProject(Control As Object)

    On Error GoTo ErrHandle

    If Not EnsureCPMSSystem() Then Exit Sub

    CreateNewProject

    Exit Sub

ErrHandle:

    MsgBox Err.Description, vbCritical

End Sub

Public Sub RibbonThemeManager(Control As Object)

    On Error GoTo ErrHandle
    If Not EnsureCPMSSystem() Then Exit Sub

    frmThemeManager.Show vbModeless
    Exit Sub

ErrHandle:

    MsgBox Err.Description, vbCritical

End Sub


'==========================================================
' IMPORT PROJECT
'==========================================================
Public Sub RibbonImportProject(Control As Object)

    On Error GoTo ErrHandle

    If Not EnsureCPMSSystem() Then Exit Sub

    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
        FileFilter:="CPMS Project (*.cpms),*.cpms," & _
                    "Excel Workbook (*.xlsx;*.xlsm),*.xlsx;*.xlsm", _
        title:="Import Project", _
        MultiSelect:=False)

    ' User b?m Cancel
    If VarType(filePath) = vbBoolean Then Exit Sub
    Dim fPath As String: fPath = CStr(filePath)
    If Len(fPath) = 0 Then Exit Sub

    Dim ext As String
    ext = LCase$(mID$(fPath, InStrRev(fPath, ".") + 1))

    Select Case ext
        Case "cpms"
            ImportProjectFromFileWithMeta fPath
        Case "xlsx", "xlsm"
            ImportProjectFromExcel fPath
        Case Else
            MsgBox "Unsupported file type: ." & ext, vbExclamation
            Exit Sub
    End Select

    ' Ch? refresh n?u form dang hi?n th? (tránh auto-instantiate khi dã Unload)
    If frmHomePage.Visible Then
        On Error Resume Next
        frmHomePage.RefreshList
        On Error GoTo 0
    End If

    Exit Sub

ErrHandle:
    MsgBox "RibbonImportProject error: " & Err.Description, vbCritical

End Sub
