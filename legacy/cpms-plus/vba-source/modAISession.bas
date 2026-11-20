Attribute VB_Name = "modAISession"
Option Explicit

'==========================================================
' modAISession
' Qu?n lý session làm vi?c v?i AI_Data.xlsb
'==========================================================

Private m_Wb       As Workbook
Private m_RefCount As Long

'----------------------------------------------------------
' M? SESSION — g?i trong UserForm_Initialize
'----------------------------------------------------------
Public Sub OpenAISession()

    Dim aiFile As String
    aiFile = GetAIDataFile()

    If Len(aiFile) = 0 Or Dir(aiFile) = "" Then
        Err.Raise vbObjectError + 300, "OpenAISession", _
            "AI data file not found. Run EnsureAISystem first."
    End If

    If Not m_Wb Is Nothing Then
        If LCase$(m_Wb.fullName) = LCase$(aiFile) Then
            m_RefCount = m_RefCount + 1
            Exit Sub
        End If
        ForceCloseAI
    End If

    Dim wb As Workbook
    For Each wb In Workbooks
        If LCase$(wb.fullName) = LCase$(aiFile) Then
            Set m_Wb = wb
            Exit For
        End If
    Next wb

    If m_Wb Is Nothing Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Set m_Wb = Workbooks.Open(aiFile)
        ' ?N TRU?C khi restore ScreenUpdating — tránh flash
        On Error Resume Next
        m_Wb.Windows(1).Visible = False
        On Error GoTo 0
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    Else
        ' Ðã m? t? tru?c, d?m b?o v?n ?n
        On Error Resume Next
        m_Wb.Windows(1).Visible = False
        On Error GoTo 0
    End If

    m_RefCount = m_RefCount + 1

End Sub

'----------------------------------------------------------
' ÐÓNG SESSION — g?i trong UserForm_QueryClose
'----------------------------------------------------------
Public Sub CloseAISession(Optional ByVal SaveChanges As Boolean = True)

    If m_RefCount > 0 Then m_RefCount = m_RefCount - 1
    If m_RefCount <= 0 Then ForceCloseAI SaveChanges

End Sub

'----------------------------------------------------------
' L?Y WORKSHEET
'----------------------------------------------------------
Public Function GetAISheet(ByVal sheetName As String) As Worksheet

    If m_Wb Is Nothing Then
        Err.Raise vbObjectError + 301, "GetAISheet", _
            "AI session not open. Call OpenAISession first."
    End If

    Set GetAISheet = m_Wb.Worksheets(sheetName)

End Function

'----------------------------------------------------------
' L?Y WORKBOOK
'----------------------------------------------------------
Public Function GetAIWorkbook() As Workbook

    If m_Wb Is Nothing Then
        Err.Raise vbObjectError + 302, "GetAIWorkbook", _
            "AI session not open. Call OpenAISession first."
    End If

    Set GetAIWorkbook = m_Wb

End Function

'----------------------------------------------------------
' LUU FILE
'----------------------------------------------------------
Public Sub SaveAIFile()

    If Not m_Wb Is Nothing Then
        Application.DisplayAlerts = False
        m_Wb.Save
        Application.DisplayAlerts = True
    End If

End Sub

'----------------------------------------------------------
' TR?NG THÁI
'----------------------------------------------------------
Public Function IsAIOpen() As Boolean
    IsAIOpen = Not (m_Wb Is Nothing)
End Function

Public Function AIRefCount() As Long
    AIRefCount = m_RefCount
End Function

'----------------------------------------------------------
' RESET — dùng khi workbook b? dóng bên ngoài
'----------------------------------------------------------
Public Sub ResetAISessionRef()
    Set m_Wb = Nothing
    m_RefCount = 0
End Sub

'==========================================================
' PRIVATE
'==========================================================
Private Sub ForceCloseAI(Optional ByVal SaveChanges As Boolean = True)

    If Not m_Wb Is Nothing Then
        Application.DisplayAlerts = False
        On Error Resume Next
        m_Wb.Close SaveChanges:=SaveChanges
        On Error GoTo 0
        Application.DisplayAlerts = True
        Set m_Wb = Nothing
    End If

    m_RefCount = 0

End Sub

