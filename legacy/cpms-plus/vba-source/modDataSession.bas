Attribute VB_Name = "modDataSession"
Option Explicit

'==========================================================
' modDataSession
'
' Qu?n lý session làm vi?c v?i project data file.
' M?i form/module dùng chung — không t? m? file riêng l?.
'
' Cách dùng trong m?i form:
'   UserForm_Initialize : DataSession.Open
'   UserForm_QueryClose : DataSession.Close
'   L?y sheet           : DataSession.Sheet("CONTRACT_DB")
'   L?y workbook        : DataSession.Workbook
'==========================================================

Private m_Wb            As Workbook
Private m_RefCount      As Long     ' Ð?m s? form dang dùng session
Private m_DataFile      As String   ' Path dang m?

'----------------------------------------------------------
' M? SESSION (g?i trong UserForm_Initialize)
' T? d?ng m? ?n n?u chua m?, tang ref count n?u dã m?
'----------------------------------------------------------
Public Sub OpenSession(Optional dummy As Variant = 0)

    Dim DataFile As String
    DataFile = ReadCurrentProjectValue("DataFile")

    If Len(DataFile) = 0 Then
        Err.Raise vbObjectError + 200, "DataSession.OpenSession", _
            "No project is currently open."
    End If

    ' A workbook can be closed externally while the object reference
    ' remains non-Nothing.  Do not reuse such a stale reference.
    Dim cachedWorkbookIsOpen As Boolean
    Dim openWb As Workbook
    If Not m_Wb Is Nothing Then
        For Each openWb In Workbooks
            If openWb Is m_Wb Then
                cachedWorkbookIsOpen = True
                Exit For
            End If
        Next openWb

        If Not cachedWorkbookIsOpen Then
            Set m_Wb = Nothing
            m_DataFile = ""
            m_RefCount = 0
        End If
    End If

    ' N?u dang m? ÐÚNG file (so sánh full path) ? ch? tang ref count
    If Not m_Wb Is Nothing Then
        If LCase$(m_Wb.fullName) = LCase$(DataFile) Then   ' <<< S?A: dùng FullName
            m_RefCount = m_RefCount + 1
            Exit Sub
        End If
        ' Ðang m? sai file ? dóng l?i
        ForceClose
    End If

    ' Tìm trong Workbooks dang m? b?ng FULL PATH              <<< S?A
    Dim wb As Workbook
    For Each wb In Workbooks
        If LCase$(wb.fullName) = LCase$(DataFile) Then
            Set m_Wb = wb
            Exit For
        End If
    Next wb

    ' Chua có ? m? m?i
    If m_Wb Is Nothing Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Set m_Wb = Workbooks.Open(DataFile)
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End If

    ' Ð?m b?o luôn ?n
    On Error Resume Next
    m_Wb.Windows(1).Visible = False
    On Error GoTo 0

    m_DataFile = DataFile
    m_RefCount = m_RefCount + 1

End Sub

'----------------------------------------------------------
' ÐÓNG SESSION (g?i trong UserForm_QueryClose)
' Gi?m ref count — ch? dóng file khi không còn form nào dùng
'----------------------------------------------------------
Public Sub CloseSession(Optional ByVal SaveChanges As Boolean = True)

    If m_RefCount > 0 Then m_RefCount = m_RefCount - 1

    If m_RefCount <= 0 Then
        ForceClose SaveChanges
    End If

End Sub

'----------------------------------------------------------
' L?Y WORKSHEET (thay cho ThisWorkbook.Worksheets / GetProjectSheet)
'----------------------------------------------------------
Public Function GetDataSheet(ByVal sheetName As String) As Worksheet
    EnsureOpen
    Set GetDataSheet = m_Wb.Worksheets(sheetName)   ' ? gán vào tên hàm
End Function

Public Function GetDataWorkbook() As Workbook
    EnsureOpen
    Set GetDataWorkbook = m_Wb                      ' ? gán vào tên hàm
End Function


'----------------------------------------------------------
' LUU FILE DATA (không dóng)
'----------------------------------------------------------
Public Sub SaveDataFile(Optional dummy As Variant = 0)

    If Not m_Wb Is Nothing Then
        Application.DisplayAlerts = False
        m_Wb.Save
        Application.DisplayAlerts = True
    End If

End Sub

'----------------------------------------------------------
' TR?NG THÁI
'----------------------------------------------------------
Public Function IsOpen() As Boolean
    IsOpen = Not (m_Wb Is Nothing)
End Function

Public Function RefCount() As Long
    RefCount = m_RefCount
End Function

'==========================================================
' PRIVATE
'==========================================================
Private Sub EnsureOpen(Optional dummy As Variant = 0)

    If m_Wb Is Nothing Then
        Err.Raise vbObjectError + 201, "DataSession", _
            "Data session is not open. Call DataSession.OpenSession first."
    End If

End Sub

Private Sub ForceClose(Optional ByVal SaveChanges As Boolean = True)

    If Not m_Wb Is Nothing Then
        Application.DisplayAlerts = False
        On Error Resume Next
        m_Wb.Close SaveChanges:=SaveChanges
        On Error GoTo 0
        Application.DisplayAlerts = True
        Set m_Wb = Nothing
    End If

    m_DataFile = ""
    m_RefCount = 0

End Sub
'----------------------------------------------------------
' RESET THAM CHIEU — dung khi workbook da bi dong ben ngoai
' (vi du: ImportProjectFromFile dong thu cong roi FileCopy)
'----------------------------------------------------------
Public Sub ResetSessionRef()
    Set m_Wb = Nothing
    m_DataFile = ""
    m_RefCount = 0
End Sub
'==========================================================
' L?Y WORKSHEET T? DATA FILE C?A PROJECT ÐANG M?
' Dùng thay cho ThisWorkbook.Worksheets(name)
'==========================================================
Public Function GetProjectSheet(ByVal sheetName As String) As Worksheet

    Dim DataFile As String
    DataFile = ReadCurrentProjectValue("DataFile")

    If Len(DataFile) = 0 Or Dir(DataFile) = "" Then
        Err.Raise vbObjectError + 100, "GetProjectSheet", _
            "No project is currently open."
    End If

    Dim wb As Workbook
    Set wb = GetOrOpenWorkbook(DataFile)

    Set GetProjectSheet = wb.Worksheets(sheetName)

End Function

'==========================================================
' L?Y WORKBOOK DATA C?A PROJECT ÐANG M?
'==========================================================
Public Function GetProjectWorkbook() As Workbook

    Dim DataFile As String
    DataFile = ReadCurrentProjectValue("DataFile")

    If Len(DataFile) = 0 Or Dir(DataFile) = "" Then
        Err.Raise vbObjectError + 101, "GetProjectWorkbook", _
            "No project is currently open."
    End If

    Set GetProjectWorkbook = GetOrOpenWorkbook(DataFile)

End Function



