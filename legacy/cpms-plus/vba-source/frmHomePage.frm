Attribute VB_Name = "frmHomePage"
Attribute VB_Base = "0{C21ED8FB-7765-4410-B1C4-6C7FAD3A72C2}{9AE6CCE8-0FBC-4254-8124-B08705B27147}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'==========================================================
' frmHomePage
'
' THAY Ð?I SO V?I B?N CU:
'   - OpenSelectedProject g?i EnsureCacheForProject(selID)
'     sau khi m? workbook ? cache luôn dúng project
'   - btnNew_Click g?i InvalidateCacheForProject tru?c khi
'     t?o project m?i ? cache không b? dùng nh?m
'   - SwitchContextIfNeeded g?i InvalidateCacheForProject
'     khi chuy?n project ? bu?c rebuild cache m?i
'==========================================================

Private m_AllRows() As Variant
Private m_RowCount  As Long

Private Sub btnExport_Click()

    Dim idx As Long
    idx = lstProjects.ListIndex

    If idx < 0 Then
        ShowMsgCenter "Please select a project to export.", vbExclamation
        Exit Sub
    End If

    Dim selID    As String
    Dim DataFile As String
    Dim projName As String
    Dim i        As Long

    selID = lstProjects.List(idx, 0)

    For i = 1 To m_RowCount
        If m_AllRows(i, 1) = selID Then
            projName = m_AllRows(i, 2)
            DataFile = m_AllRows(i, 5)
            Exit For
        End If
    Next i

    ExportSelectedProject DataFile, selID

End Sub

Private Sub btnImport_Click()

    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
        FileFilter:="CPMS Project (*.cpms),*.cpms," & _
                    "Excel Workbook (*.xlsx;*.xlsm),*.xlsx;*.xlsm", _
        title:="Import Project", _
        MultiSelect:=False)

    If VarType(filePath) = vbBoolean Then Exit Sub
    Dim fPath As String: fPath = CStr(filePath)
    If Len(fPath) = 0 Then Exit Sub

    Dim ext As String
    ext = LCase$(mID$(fPath, InStrRev(fPath, ".") + 1))

    Select Case ext
        Case "cpms"
            'ImportProjectFromFile fPath
            ImportProjectFromFileWithMeta fPath
        Case "xlsx", "xlsm"
            ImportProjectFromExcel fPath
        Case Else
            ShowMsgCenter "Unsupported file type: ." & ext, vbExclamation
            Exit Sub
    End Select

    RefreshList

End Sub


Private Sub UserForm_Activate()
    RefreshList
End Sub

'==========================================================
' INITIALIZE
'==========================================================
Private Sub UserForm_Initialize()

    Me.Caption = "CPMS — Project Manager"

    ConfigureListBox

    txtProjectInfo.Font.Name = "Corbel"
    txtProjectInfo.Font.Size = 12

    RefreshList

    If lstProjects.ListCount = 0 And Not gSuppressHomePageAutoCreate Then
        frmProjectInfo.Init "", ""
        frmProjectInfo.Show vbModal

        If Not frmProjectInfo.Cancelled Then
            CreateProject frmProjectInfo.projName, frmProjectInfo.location
            RefreshList   ' ? l?i HomePage, refresh d? th?y project m?i
        End If
    End If

    gSuppressHomePageAutoCreate = False
    ApplyTheme Me
    'UnregisterForm Me
End Sub

Private Sub ConfigureListBox()

    With lstProjects
        .ColumnCount = 5
        .ColumnHeads = False
        .ColumnWidths = "90;220;0;75;55"
        .MultiSelect = fmMultiSelectSingle
    End With

End Sub

'==========================================================
' REFRESH
'==========================================================
Public Sub RefreshList()

    LoadAllProjects
    PopulateListBox
    HighlightCurrentProject
    ShowSelectedProjectInfo

End Sub

'==========================================================
' LOAD ALL PROJECTS t? ProjectList.xlsb
'==========================================================
Private Sub LoadAllProjects()

    Dim wb  As Workbook
    Dim ws  As Worksheet
    Dim lr  As Long
    Dim i   As Long

    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.Cursor = xlDefault

    Set wb = Workbooks.Open(GetProjectListFile)
    Set ws = wb.Worksheets("PROJECT_LIST")

    lr = ws.Cells(ws.rows.count, 1).End(xlUp).Row

    If lr < 2 Then
        wb.Close False
        Application.Cursor = xlDefault
        m_RowCount = 0
        Application.ScreenUpdating = True
        SetStatus "No projects found."
        Exit Sub
    End If

    m_RowCount = lr - 1
    ReDim m_AllRows(1 To m_RowCount, 1 To 8)

    For i = 1 To m_RowCount
        m_AllRows(i, 1) = ws.Cells(i + 1, 1).value
        m_AllRows(i, 2) = ws.Cells(i + 1, 2).value
        m_AllRows(i, 3) = ws.Cells(i + 1, 3).value
        m_AllRows(i, 4) = ws.Cells(i + 1, 4).value
        m_AllRows(i, 5) = ws.Cells(i + 1, 5).value
        m_AllRows(i, 6) = Format(ws.Cells(i + 1, 6).value, "dd-mm-yyyy")
        m_AllRows(i, 7) = Format(ws.Cells(i + 1, 7).value, "dd-mm-yyyy")
        m_AllRows(i, 8) = ws.Cells(i + 1, 8).value
    Next i

    wb.Close False
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True

    Exit Sub

EH:
    If Not wb Is Nothing Then wb.Close False
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    SetStatus "Error loading projects: " & Err.Description

End Sub

'==========================================================
' POPULATE LISTBOX
'==========================================================
Private Sub PopulateListBox()

    Dim i As Long

    lstProjects.Clear

    For i = 1 To m_RowCount
        If LCase$(m_AllRows(i, 8)) <> "deleted" Then
            With lstProjects
                .AddItem m_AllRows(i, 1)
                .List(.ListCount - 1, 1) = m_AllRows(i, 2)
                .List(.ListCount - 1, 2) = m_AllRows(i, 3)
                .List(.ListCount - 1, 3) = m_AllRows(i, 6)
                .List(.ListCount - 1, 4) = m_AllRows(i, 8)
            End With
        End If
    Next i

    SetStatus "Total Projects : " & lstProjects.ListCount

End Sub

'==========================================================
' HIGHLIGHT current project
'==========================================================
Private Sub HighlightCurrentProject()

    Dim curID As String
    Dim i     As Long

    curID = ReadCurrentProjectValue("ProjectID")

    If Len(curID) = 0 Then Exit Sub

    For i = 0 To lstProjects.ListCount - 1
        If lstProjects.List(i, 0) = curID Then
            lstProjects.ListIndex = i
            Exit For
        End If
    Next i

End Sub

'==========================================================
' SHOW INFO c?a project dang ch?n
'==========================================================
Private Sub ShowSelectedProjectInfo()

    Dim idx As Long
    idx = lstProjects.ListIndex

    If idx < 0 Then
        txtProjectInfo.value = "(No project selected)"
        Exit Sub
    End If

    Dim selID As String
    Dim i     As Long
    selID = lstProjects.List(idx, 0)

    For i = 1 To m_RowCount
        If m_AllRows(i, 1) = selID Then
            txtProjectInfo.value = _
                InfoRow("Project ID", m_AllRows(i, 1)) & _
                InfoRow("Name", m_AllRows(i, 2)) & _
                InfoRow("Location", m_AllRows(i, 3)) & _
                InfoRow("Created", m_AllRows(i, 6)) & _
                InfoRow("Last Access", m_AllRows(i, 7)) & _
                InfoRow("Status", m_AllRows(i, 8))
            Exit For
        End If
    Next i

End Sub

'==========================================================
' LIST EVENTS
'==========================================================
Private Sub lstProjects_Click()
    ShowSelectedProjectInfo
End Sub

Private Sub lstProjects_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    OpenSelectedProject
End Sub

'==========================================================
' BUTTONS
'==========================================================
Private Sub btnOpen_Click()
    OpenSelectedProject
End Sub

Private Sub btnNew_Click()

    If Not IsActivated() Then
        ShowMsgCenter "Software is not activated or license is invalid." & vbCrLf & _
                      "Please contact your vendor to activate.", vbCritical
        Exit Sub
    End If

    frmProjectInfo.Init "", ""
    frmProjectInfo.Show vbModal

    If frmProjectInfo.Cancelled Then Exit Sub

    CreateProject frmProjectInfo.projName, frmProjectInfo.location

    ' ? l?i frmHomePage — refresh list d? th?y project v?a t?o
    RefreshList

End Sub

Private Sub btnRename_Click()

    If lstProjects.ListIndex < 0 Then
        ShowMsgCenter "Please select a project to rename.", vbExclamation
        Exit Sub
    End If

    RenameCurrentProject lstProjects.List(lstProjects.ListIndex, 0)
    RefreshList

End Sub

Private Sub btnClone_Click()

    If lstProjects.ListIndex < 0 Then
        ShowMsgCenter "Please select a project to clone.", vbExclamation
        Exit Sub
    End If

    CloneCurrentProject lstProjects.List(lstProjects.ListIndex, 0)
    RefreshList

End Sub


Private Sub btnDelete_Click()

    Dim idx As Long
    idx = lstProjects.ListIndex

    If idx < 0 Then
        ShowMsgCenter "Please select a project to delete.", vbExclamation
        Exit Sub
    End If

    Dim selName As String
    Dim selID   As String

    selName = lstProjects.List(idx, 1)
    selID = lstProjects.List(idx, 0)

    Dim ans As Long
    ans = UniMsgBox( _
        "Delete project """ & selName & """?" & vbCrLf & vbCrLf & _
        "YES = Delete permanently (remove ALL data files from disk)." & vbCrLf & _
        "NO  = Cancel, do nothing.", _
        vbYesNo + vbCritical + vbDefaultButton2, "Delete Project")

    If ans = vbNo Then Exit Sub

    DeleteProject selID, True

    RefreshList
    SelectMostRecentProject

End Sub

'==========================================================
' OPEN SELECTED PROJECT
'   *** THAY Ð?I CHÍNH: g?i EnsureCacheForProject(selID)
'       sau khi m? workbook d? cache kh?p dúng project ***
'==========================================================
Private Sub OpenSelectedProject()

    Dim idx As Long
    idx = lstProjects.ListIndex

    If idx < 0 Then
        ShowMsgCenter "Please select a project to open.", vbExclamation
        Exit Sub
    End If

    If Not IsActivated() Then
        ShowMsgCenter "Software is not activated or license is invalid." & vbCrLf & _
                      "Please contact your vendor to activate.", vbCritical
        Exit Sub
    End If

    Dim selID     As String
    Dim DataFile  As String
    Dim selName   As String
    Dim selFolder As String
    Dim selLoc    As String
    Dim i         As Long

    selID = lstProjects.List(idx, 0)

    For i = 1 To m_RowCount
        If m_AllRows(i, 1) = selID Then
            DataFile = m_AllRows(i, 5)
            selName = m_AllRows(i, 2)
            selFolder = m_AllRows(i, 4)
            selLoc = m_AllRows(i, 3)
            Exit For
        End If
    Next i

    If Dir(DataFile) = "" Then
        ShowMsgCenter "Data file not found:" & vbCrLf & DataFile, vbCritical
        Exit Sub
    End If

    ' *** Ðóng workbook c?a project dang m? (n?u khác project m?i ch?n) ***
    ' T?t c? project dùng cùng tên CPMS_Data.xlsb — n?u d? m? cùng lúc,
    ' Workbooks.Open s? fail silently và DataSession.m_Wb = Nothing ? crash.
    Dim oldDataFile As String
    oldDataFile = ReadCurrentProjectValue("DataFile")
    If Len(oldDataFile) > 0 And LCase$(oldDataFile) <> LCase$(DataFile) Then
        Dim wbOld As Workbook
        For Each wbOld In Workbooks
            If LCase$(wbOld.fullName) = LCase$(oldDataFile) Then
                Application.DisplayAlerts = False
                wbOld.Close False
                Application.DisplayAlerts = True
                Exit For
            End If
        Next wbOld
    End If

    SaveCurrentProject selID, selName, selFolder, DataFile
    UpdateLastOpened selID

    OpenProjectWorkbook DataFile

    EnsureProjectInfoInSheet DataFile, selID, selName, selLoc

    ' Rebuild the contract/payment caches from the workbook just opened.
    ' Otherwise the list can display a stale in-memory contract that is
    ' no longer present in the physical project file.
    InvalidateProjectRuntimeCaches

    Me.Hide
    ShowContractSearch

End Sub

'==========================================================
' INVALIDATE RUNTIME CACHES AFTER REOPENING A PROJECT FILE
'==========================================================
Private Sub InvalidateProjectRuntimeCaches()

    On Error Resume Next
    gContractLoaded = False
    gPaymentLoaded = False
    gPaymentEntryLoaded = False
    gBCPBLoaded = False
    gDocLoaded = False
    gCacheProjectID = ""
    On Error GoTo 0

End Sub

'==========================================================
' SHOW CONTRACT SEARCH
'==========================================================
Private Sub ShowContractSearch()

    Application.ScreenUpdating = False
    Application.Cursor = xlDefault

    On Error Resume Next
    Unload frmContractSearch
    Set frmContractSearch = Nothing
    On Error GoTo 0

    frmContractSearch.Show vbModeless
    frmContractSearch.InitFast

    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    DoEvents

End Sub

'==========================================================
' DELETE PROJECT
'==========================================================
Private Sub DeleteProject(ByVal projectID As String, ByVal deleteFiles As Boolean)

    Dim wb            As Workbook
    Dim ws            As Worksheet
    Dim f             As Range
    Dim projectFolder As String
    Dim DataFile      As String

    On Error GoTo EH

    Application.ScreenUpdating = False

    Set wb = GetOrOpenWorkbook(GetProjectListFile)
    Set ws = wb.Worksheets("PROJECT_LIST")

    Set f = ws.Columns(1).Find(What:=projectID, LookAt:=xlWhole, LookIn:=xlValues)

    If Not f Is Nothing Then
        projectFolder = f.Offset(0, 3).value
        DataFile = f.Offset(0, 4).value

        If deleteFiles Then
            f.EntireRow.Delete
        Else
            f.Offset(0, 7).value = "Deleted"
        End If
    End If

    wb.Save
    wb.Close False
    Application.ScreenUpdating = True

    If deleteFiles And Len(projectFolder) > 0 Then

        ' *** FIX: tìm workbook b?ng FULL PATH thay vì tên ng?n ***
        Dim wbData  As Workbook
        Dim wbCheck As Workbook
        For Each wbCheck In Workbooks
            If LCase$(wbCheck.fullName) = LCase$(DataFile) Then
                Set wbData = wbCheck
                Exit For
            End If
        Next wbCheck

        If Not wbData Is Nothing Then
            Application.DisplayAlerts = False
            wbData.Close False
            Application.DisplayAlerts = True
            Set wbData = Nothing
        End If

        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
        DoEvents

        Dim deleted As Boolean
        deleted = False

        If FolderExists(projectFolder) Then
            Dim fso As Object
            Set fso = CreateObject("Scripting.FileSystemObject")
            On Error Resume Next
            fso.DeleteFolder projectFolder, True
            deleted = (Err.Number = 0)
            On Error GoTo 0
        Else
            deleted = True
        End If

        If Not deleted Then
            Dim cleanPath As String
            cleanPath = projectFolder
            If Right$(cleanPath, 1) = "\" Then
                cleanPath = Left$(cleanPath, Len(cleanPath) - 1)
            End If
            Shell "cmd /c rd /s /q """ & cleanPath & """"
        End If

    End If

    ' N?u project b? xóa dang là current ? xóa INI (SelectMostRecentProject s? ghi l?i)
    If ReadCurrentProjectValue("ProjectID") = projectID Then
        SaveCurrentProject "", "", "", ""
    End If

    Exit Sub

EH:
    If Not wb Is Nothing Then wb.Close False
    Application.ScreenUpdating = True
    ShowMsgCenter "Error deleting project: " & Err.Description, vbCritical

End Sub
'==========================================================
' SELECT MOST RECENT PROJECT
' Sau khi delete: t? d?ng ch?n project m? g?n nh?t còn l?i
' trong lstProjects, c?p nh?t txtProjectInfo và CurrentProject.ini
'==========================================================
Private Sub SelectMostRecentProject()

    ' Không còn project nào ? xóa INI, xóa info
    If lstProjects.ListCount = 0 Then
        txtProjectInfo.value = "(No projects available)"
        SaveCurrentProject "", "", "", ""
        Exit Sub
    End If

    ' Tìm project có LastOpened l?n nh?t trong m_AllRows
    ' m_AllRows(i, 7) = Format(LastOpened, "dd-mm-yyyy")
    ' Ð?i sang "yyyymmdd" d? so sánh chu?i dúng th? t? th?i gian
    Dim bestRowIdx  As Long
    Dim bestDateStr As String
    Dim i           As Long

    bestRowIdx = 0
    bestDateStr = ""

    For i = 1 To m_RowCount

        ' B? qua project dã b? dánh d?u Deleted
        If LCase$(CStr(m_AllRows(i, 8))) = "deleted" Then GoTo NextRow

        ' "dd-mm-yyyy" ? "yyyymmdd"
        Dim raw  As String
        Dim comp As String
        raw = CStr(m_AllRows(i, 7))

        If Len(raw) = 10 Then
            comp = mID$(raw, 7, 4) & mID$(raw, 4, 2) & Left$(raw, 2)
        Else
            comp = "00000000"
        End If

        If comp > bestDateStr Then
            bestDateStr = comp
            bestRowIdx = i
        End If

NextRow:
    Next i

    If bestRowIdx = 0 Then
        txtProjectInfo.value = "(No projects available)"
        SaveCurrentProject "", "", "", ""
        Exit Sub
    End If

    ' Ch?n project dó trong lstProjects
    Dim targetID As String
    targetID = CStr(m_AllRows(bestRowIdx, 1))

    Dim j As Long
    For j = 0 To lstProjects.ListCount - 1
        If lstProjects.List(j, 0) = targetID Then
            lstProjects.ListIndex = j
            Exit For
        End If
    Next j

    ' Hi?n th? thông tin vào txtProjectInfo
    ShowSelectedProjectInfo

    ' C?p nh?t CurrentProject.ini
    SaveCurrentProject _
        CStr(m_AllRows(bestRowIdx, 1)), _
        CStr(m_AllRows(bestRowIdx, 2)), _
        CStr(m_AllRows(bestRowIdx, 4)), _
        CStr(m_AllRows(bestRowIdx, 5))

End Sub
'==========================================================
' XÓA FOLDER Ð? QUY
'==========================================================
Private Sub DeleteFolderRecursive(ByVal folderPath As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    fso.DeleteFolder folderPath, True
    On Error GoTo 0

End Sub

'==========================================================
' UPDATE LAST OPENED
'==========================================================
Private Sub UpdateLastOpened(ByVal projectID As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim f  As Range

    On Error GoTo EH

    Application.ScreenUpdating = False

    Set wb = Workbooks.Open(GetProjectListFile)
    Set ws = wb.Worksheets("PROJECT_LIST")

    Set f = ws.Columns(1).Find(What:=projectID, LookAt:=xlWhole, LookIn:=xlValues)

    If Not f Is Nothing Then
        f.Offset(0, 6).value = Now
    End If

    wb.Save
    wb.Close False

    Application.ScreenUpdating = True

    Exit Sub

EH:
    If Not wb Is Nothing Then wb.Close False
    Application.ScreenUpdating = True

End Sub

'==========================================================
' STATUS BAR
'==========================================================
Private Sub SetStatus(ByVal msg As String)
    lblStatus.Caption = msg
End Sub

'==========================================================
' ESC Ð? ÐÓNG
'==========================================================
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                              ByVal Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

'==========================================================
' EXPORT PROJECT ÐU?C CH?N
'==========================================================
Public Sub ExportSelectedProject(ByVal DataFile As String, ByVal projName As String)

    If Len(DataFile) = 0 Or Dir(DataFile) = "" Then
        ShowMsgCenter "Data file not found:" & vbCrLf & DataFile, vbExclamation
        Exit Sub
    End If

    ' Persist changes before the export routine closes the workbook
    ' with Close False.
    On Error GoTo EH
    SaveDataFile
    On Error GoTo 0

    ' -- Read PROJ code directly from DataFile --
    Dim projCode As String
    projCode = Trim$(GetSettingFromDataFile(DataFile, "PROJ"))

    If projCode = "" Then
        Dim entered As String
        entered = InputBox("Project Code (PROJ) is not set in this project." & vbCrLf & _
                           "Enter a project code to continue (e.g. CHC1):", _
                           "Enter Project Code", "")
        If StrPtr(entered) = 0 Or Trim$(entered) = "" Then
            ShowMsgCenter "Export cancelled: no project code provided.", vbExclamation
            Exit Sub
        End If
        projCode = Trim$(entered)
        SetSettingInDataFile DataFile, "PROJ", projCode
    End If

    ' -- Create export folder in Documents --
    Dim exportDir As String
    exportDir = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\CPMS_Exports"
    If Not FolderExists(exportDir) Then
        On Error Resume Next
        MkDir exportDir
        On Error GoTo 0
    End If

    ' -- Filename: PROJ_YYYYMMDD_HHMMSS.cpms --
    Dim SafeCode As String
    SafeCode = projCode
    Dim illChars As Variant, ch As Variant
    illChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In illChars
        SafeCode = Join(Split(SafeCode, CStr(ch)), "_")
    Next ch

    Dim defaultFile As String
    defaultFile = exportDir & "\" & SafeCode & "_" & Format$(Now, "YYYYMMDD_HHMMSS") & ".cpms"

    ' -- Save dialog --
    Dim SavePath As Variant
    On Error Resume Next
    ChDrive Left$(exportDir, 2)
    ChDir exportDir
    On Error GoTo 0

    SavePath = Application.GetSaveAsFilename( _
                    InitialFileName:=defaultFile, _
                    FileFilter:="CPMS Project File (*.cpms),*.cpms", _
                    title:="Export Project — choose folder and filename")

    If VarType(SavePath) = vbBoolean Then Exit Sub

    Dim saveStr As String
    saveStr = CStr(SavePath)
    If LCase$(Right$(saveStr, 5)) <> ".cpms" Then saveStr = saveStr & ".cpms"

    ' -- Close workbook if open, copy, reopen --
    Dim wbData  As Workbook
    Dim wasOpen As Boolean
    For Each wbData In Workbooks
        If LCase$(wbData.fullName) = LCase$(DataFile) Then
            wasOpen = True: Exit For
        End If
    Next wbData

    If wasOpen Then
        Application.DisplayAlerts = False
        wbData.Close False
        Application.DisplayAlerts = True
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
        DoEvents
    End If

    Dim errDesc As String
    On Error GoTo EH
    FileCopy DataFile, saveStr
    On Error GoTo 0

    If wasOpen Then OpenProjectWorkbook DataFile
    ShowMsgCenter "Export successful!" & vbCrLf & vbCrLf & saveStr, vbInformation, "Export Project"
    Exit Sub

EH:
    errDesc = Err.Description
    On Error GoTo 0
    If wasOpen Then OpenProjectWorkbook DataFile
    ShowMsgCenter "Export failed: " & errDesc, vbCritical

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'ApplyTheme Me
    UnregisterForm Me
End Sub


