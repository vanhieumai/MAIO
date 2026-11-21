Attribute VB_Name = "modCPMSBootstrap"
Option Explicit

'==========================================================
' CPMS Bootstrap
'
' Responsible for:
'   - License checking
'   - System folders
'   - Data Template
'   - Project List
'
' NOTE:
'   GetAppFolder() is provided by modLicense
'==========================================================
'==========================================================
' Windows API — ch?ng flash khi chuy?n form
'==========================================================
#If VBA7 Then
    Private Declare PtrSafe Function LockWindowUpdate Lib "user32" _
        (ByVal hwndLock As LongPtr) As Long
#Else
    Private Declare Function LockWindowUpdate Lib "user32" _
        (ByVal hwndLock As Long) As Long
#End If

'==========================================================
' UNICODE MSGBOX — dùng thay MsgBox khi có ti?ng Vi?t
'==========================================================
#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpText As LongPtr, _
        ByVal lpCaption As LongPtr, _
        ByVal uType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal lpText As Long, _
        ByVal lpCaption As Long, _
        ByVal uType As Long) As Long
#End If



Private Const TEMPLATE_FOLDER As String = "DataTemplate\"
Private Const PROJECT_FOLDER As String = "Projects\"

Private Const TEMPLATE_FILE As String = "CPMS_DataTemplate.xlsb"
Private Const PROJECTLIST_FILE As String = "ProjectList.xlsb"
Private Const CURRENTPROJECT_FILE As String = "CurrentProject.ini"
Private Const TEMPLATE_VERSION      As String = "1.0"
Private Const TEMPLATE_VERSION_FILE As String = "DataTemplateVersion.ini"
Public gSuppressHomePageAutoCreate As Boolean

'==========================================================
' AI SYSTEM — Constants & Paths
'==========================================================
Private Const AI_FOLDER    As String = "AI\"
Private Const AI_DATA_FILE As String = "AI_Data.xlsb"

Public Function GetAIFolder() As String
    GetAIFolder = GetAppFolder & AI_FOLDER
End Function

Public Function GetAIDataFile() As String
    GetAIDataFile = GetAIFolder & AI_DATA_FILE
End Function

'==========================================================
' ENSURE AI SYSTEM
' T?o thu m?c AI\ và AI_Data.xlsb n?u chua có
'==========================================================
Public Function EnsureAISystem() As Boolean

    On Error GoTo EH

    EnsureFolder GetAIFolder

    ' T?o m?i n?u chua có
    If Dir(GetAIDataFile) = "" Then
        BuildAIDataFile
    Else
        ' File có nhung chua ch?c d? sheet ? validate
        If Not ValidateAIDataFile() Then
            ' Xóa file cu, build l?i
            On Error Resume Next
            Kill GetAIDataFile
            On Error GoTo 0
            BuildAIDataFile
        End If
    End If

    EnsureAISystem = True
    Exit Function

EH:
    MsgBox "AI System init error: " & vbCrLf & Err.Description, vbCritical

End Function

'----------------------------------------------------------
' Ki?m tra AI_Data.xlsb có d? 2 sheet c?n thi?t không
'----------------------------------------------------------
Private Function ValidateAIDataFile() As Boolean

    Dim wb  As Workbook
    Dim ws  As Worksheet
    Dim ok  As Boolean

    On Error GoTo bad

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wb = Workbooks.Open(GetAIDataFile)

    On Error Resume Next
    Set ws = wb.Worksheets("SETTINGS")
    ok = Not (ws Is Nothing)
    If ok Then
        Set ws = Nothing
        Set ws = wb.Worksheets("AI_NOTES")
        ok = Not (ws Is Nothing)
    End If
    On Error GoTo 0

    wb.Close False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ValidateAIDataFile = ok
    Exit Function

bad:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ValidateAIDataFile = False

End Function

'==========================================================
' BUILD AI DATA FILE
' Copy 2 sheet SETTING và AI_NOTES t? ThisWorkbook
' vào AI_Data.xlsb trong thu m?c AI\
'==========================================================
Private Sub BuildAIDataFile()

    Dim wbNew As Workbook
    Dim ws    As Worksheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo EH_Build

    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    wbNew.Windows(1).Visible = False

    ' --- T?o sheet SETTINGS t? d?u ---
    Set ws = wbNew.Worksheets(1)
    ws.Name = "SETTINGS"
    ws.Range("A1").value = "key"
    ws.Range("B1").value = "value"
    ws.rows(1).Font.Bold = True

    ' --- T?o sheet AI_NOTES t? d?u ---
    Set ws = wbNew.Worksheets.Add( _
        After:=wbNew.Worksheets(wbNew.Worksheets.count))
    ws.Name = "AI_NOTES"
    ws.Range("A1").value = "datetime"
    ws.Range("B1").value = "question"
    ws.Range("C1").value = "answer"
    ws.Range("D1").value = "model"
    ws.rows(1).Font.Bold = True

    wbNew.SaveAs fileName:=GetAIDataFile, fileFormat:=xlExcel12
    wbNew.Close False

CleanUp_Build:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

EH_Build:
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close False
    On Error GoTo 0
    Resume CleanUp_Build

End Sub
Public Function UniMsgBox(ByVal Prompt As String, _
                           ByVal Buttons As Long, _
                           ByVal title As String) As Long
    UniMsgBox = MessageBoxW(0, StrPtr(Prompt), StrPtr(title), Buttons)
End Function
Public Function GetTemplateVersionFile() As String
    GetTemplateVersionFile = GetAppFolder & TEMPLATE_VERSION_FILE
End Function
'==========================================================
' Main Entry
'==========================================================
'==========================================================
' ENSURE CPMS SYSTEM — có loading form
'==========================================================
Public Function EnsureCPMSSystem() As Boolean

    On Error GoTo EH

    If Not IsActivated Then
        'LinkKL.Show vbModal
        frmLicense.Show vbModeless
        If Not IsActivated Then Exit Function
    End If

    Dim step As String

    step = "EnsureSystemFolders"
    EnsureSystemFolders

    step = "EnsureProjectList"
    EnsureProjectList

    step = "EnsureDataTemplate"
    EnsureDataTemplate

    step = "EnsureCurrentProjectFile"
    EnsureCurrentProjectFile

    EnsureCPMSSystem = True
    Exit Function

EH:
    MsgBox "Failed at step: [" & step & "]" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical

End Function


'==========================================================
' Create folders
'==========================================================
Private Sub EnsureSystemFolders(Optional dummy As Variant = 0)

    ' Original folders
    EnsureFolder GetTemplateFolder
    EnsureFolder GetProjectsFolder

    ' NEW: ensure app root + Themes folder exist
    EnsureFolder GetCPMSAppFolder()
    EnsureFolder GetThemesFolder()

    ' NEW: create Settings.ini with defaults if missing
    EnsureSettingsFile

End Sub

Private Sub EnsureFolder(ByVal folderPath As String)

    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If

End Sub

'==========================================================
' Folder Path
'==========================================================
Public Function GetTemplateFolder() As String

    GetTemplateFolder = GetAppFolder & TEMPLATE_FOLDER

End Function

Public Function GetProjectsFolder() As String

    GetProjectsFolder = GetAppFolder & PROJECT_FOLDER

End Function

'==========================================================
' File Path
'==========================================================
Public Function GetTemplateFile() As String

    GetTemplateFile = GetTemplateFolder & TEMPLATE_FILE

End Function

Public Function GetProjectListFile() As String

    GetProjectListFile = GetAppFolder & PROJECTLIST_FILE

End Function

Public Function GetCurrentProjectFile() As String

    GetCurrentProjectFile = GetAppFolder & CURRENTPROJECT_FILE

End Function

'==========================================================
' Ensure Data Template
'==========================================================
Private Sub EnsureDataTemplate(Optional dummy As Variant = 0)

    ' -- No template found -> build fresh --
    If Dir(GetTemplateFile) = "" Then

        ' Bring Excel to the foreground so the dialog is not
        ' hidden behind a modeless form open at startup.
        On Error Resume Next
        AppActivate Application.Caption
        On Error GoTo 0
        DoEvents

        Dim confirmMsg As String
        confirmMsg = "No CPMS data was found on this computer." & vbCrLf & vbCrLf & _
                     "Do you want to initialize the data now?" & vbCrLf & vbCrLf & _
                     "  - A new data template will be created." & vbCrLf & _
                     "  - Any .cpmstheme files found in the same" & vbCrLf & _
                     "    folder as this program will be installed" & vbCrLf & _
                     "    automatically as default themes." & vbCrLf & vbCrLf & _
                     "Click YES to initialize, NO to skip."

        ' vbDefaultButton2 = default is No, so accidental
        ' Enter does NOT trigger a silent Yes.
        Dim ans As Long
        ans = MsgBox(confirmMsg, _
                     vbYesNo + vbQuestion + vbDefaultButton2, _
                     "Initialize CPMS Data")
        If ans <> vbYes Then Exit Sub

        BuildDataTemplateWithProgress
        SaveTemplateVersion

        ' Auto-install bundled .cpmstheme files then activate
        ' the first one so colors load correctly on first open.
        InstallBundledThemes

        ' --- H?i ngu?i dùng có mu?n luu file này thành .xlam ---
        OfferConvertToXlam

        Exit Sub
    End If

    ' -- Template exists -> check version --
    Dim savedVer As String
    savedVer = ReadTemplateVersion()
    If savedVer = TEMPLATE_VERSION Then
        Exit Sub   ' Same version -> skip entirely
    End If

    ' -- Different version -> ask user --
    Dim ansUpd As Integer
    ansUpd = MsgBox( _
        "A newer data template is available." & vbCrLf & vbCrLf & _
        "Installed : " & IIf(Len(savedVer) = 0, "(unknown)", savedVer) & vbCrLf & _
        "Available : " & TEMPLATE_VERSION & vbCrLf & vbCrLf & _
        "Update now? Existing project data will NOT be affected.", _
        vbYesNo + vbQuestion, "Template Update")
    If ansUpd <> vbYes Then Exit Sub

    ' Delete old template -> rebuild
    On Error Resume Next
    Kill GetTemplateFile
    On Error GoTo 0
    BuildDataTemplateWithProgress
    SaveTemplateVersion

End Sub


'==========================================================
' INSTALL BUNDLED THEMES  (NEW SUB)
'
' Scans every *.cpmstheme file in the same folder as
' ThisWorkbook (i.e. the folder that contains the running
' .xlsm / .xlam file) and installs each one automatically
' via InstallThemeFromFileEx.
'
' Uses mode "silent" so no per-file dialog is shown.
' If a theme already exists, InstallThemeFromFileEx skips it.
'
' Shows a single summary MsgBox at the end (only when at
' least one theme was successfully installed).
'==========================================================
Private Sub InstallBundledThemes()

    On Error GoTo EH_Bundle

    Dim searchDir As String
    searchDir = ThisWorkbook.path   ' folder containing the .xlsm / .xlam

    Dim fileName As String
    fileName = Dir(searchDir & "\*.cpmstheme")

    If Len(fileName) = 0 Then Exit Sub   ' nothing to install

    Dim installedCount As Long
    Dim skippedCount   As Long
    Dim installedNames As String
    Dim firstThemeName As String   ' used to activate after loop

    Do While Len(fileName) > 0

        Dim fullPath As String
        fullPath = searchDir & "\" & fileName

        Dim themeName As String
        themeName = InstallThemeFromFileEx(fullPath, "ask")

        If Len(themeName) > 0 Then
            installedCount = installedCount + 1
            ' Remember the first successfully installed theme
            If installedCount = 1 Then firstThemeName = themeName
            If Len(installedNames) > 0 Then installedNames = installedNames & vbCrLf
            installedNames = installedNames & "  - " & themeName
        Else
            skippedCount = skippedCount + 1
            Debug.Print "InstallBundledThemes: skipped -> " & fullPath
        End If

        fileName = Dir()

    Loop

    ' -------------------------------------------------------
    ' ACTIVATE the first installed theme if no theme is
    ' currently set (Settings.ini CurrentTheme is empty or
    ' points to a folder that doesn't exist yet).
    '
    ' This makes SetCurrentTheme + SetDefaultTheme do the
    ' same work as clicking Apply in Theme Manager:
    '   1. Write theme name to Settings.ini
    '   2. Set m_CurrentTheme (module-level var in modTheme)
    '   3. Clear + reload the color cache from colors.ini
    ' Without this step, ApplyTheme on the first form open
    ' would find an empty/invalid m_CurrentTheme and fall
    ' back to hardcoded RGB values from FallbackColor().
    ' -------------------------------------------------------
    If installedCount > 0 And Len(firstThemeName) > 0 Then
        ' Check whether a valid theme is already active
        Dim activeTheme As String
        activeTheme = GetCurrentTheme()   ' returns fallback "Default" if Settings empty

        Dim activeFolder As String
        activeFolder = GetThemeFolder(activeTheme)

        If Not ThemeFolderExists(activeFolder) Then
            ' No valid theme active — promote the first installed theme
            On Error Resume Next
            SetDefaultTheme firstThemeName   ' writes DefaultTheme= to Settings.ini
            SetCurrentTheme firstThemeName   ' writes CurrentTheme=, loads cache
            On Error GoTo 0
            Debug.Print "InstallBundledThemes: activated """ & firstThemeName & """ as default"
        End If
    End If
    ' -------------------------------------------------------

    ' Summary MsgBox (only when at least one theme was installed)
    If installedCount > 0 Then
        Dim resultMsg As String
        resultMsg = installedCount & " default theme(s) installed from the setup folder:" & _
                    vbCrLf & vbCrLf & installedNames

        If skippedCount > 0 Then
            resultMsg = resultMsg & vbCrLf & vbCrLf & _
                        "(" & skippedCount & " file(s) skipped - " & _
                        "theme already exists or install failed.)"
        End If

        MsgBox resultMsg, vbInformation, "Auto Theme Install"
    End If

    Exit Sub

EH_Bundle:
    On Error Resume Next
    Debug.Print "InstallBundledThemes error: " & Err.Description
    On Error GoTo 0

End Sub

'==========================================================
' OFFER CONVERT TO XLAM
' G?i sau khi kh?i t?o d? li?u xong. H?i ngu?i dùng có mu?n
' luu file hi?n t?i thành Excel Add-In (.xlam) và cài d?t luôn.
'==========================================================
Private Sub OfferConvertToXlam()

    ' B? qua n?u file dã là xlam / xla r?i
    If ThisWorkbook.IsAddin Then Exit Sub

    Dim ext As String
    ext = LCase(mID$(ThisWorkbook.Name, _
                     InStrRev(ThisWorkbook.Name, ".") + 1))
    If ext = "xlam" Or ext = "xla" Then Exit Sub

    ' ----- Tên file xlam s? du?c t?o -----
    Dim baseName As String
    baseName = Left$(ThisWorkbook.Name, _
                     InStrRev(ThisWorkbook.Name, ".") - 1)

    ' UserLibraryPath là thu m?c chu?n c?a Excel (%APPDATA%\Microsoft\AddIns\)
    ' Ðây là noi duy nh?t Excel tì m add-in d? hi?n trong h?p tho?i Add-Ins.
    ' UserLibraryPath luôn k?t thúc b?ng d?u \ nên n?i tr?c ti?p không c?n thêm \.
    Dim xlamPath As String
    xlamPath = Application.UserLibraryPath & baseName & ".xlam"

    ' ----- H?i ngu?i dùng -----
    Dim ans As Long
    ans = MsgBox( _
        "Data initialization is complete." & vbCrLf & vbCrLf & _
        "Do you want to save this file as an Excel Add-In (.xlam)?" & vbCrLf & vbCrLf & _
        "  - The add-in will be installed to:" & vbCrLf & _
        "    " & xlamPath & vbCrLf & _
        "  - It will appear in Excel's Add-Ins list automatically." & vbCrLf & _
        "  - CPMS will load every time Excel starts." & vbCrLf & vbCrLf & _
        "Click YES to convert now, NO to keep the current format.", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "Convert to Add-In (.xlam)")

    If ans <> vbYes Then Exit Sub

    On Error GoTo EH_Xlam

    Application.DisplayAlerts = False

    ' Xóa file xlam cu n?u dã t?n t?i
    If Dir(xlamPath) <> "" Then
        On Error Resume Next
        Kill xlamPath
        On Error GoTo EH_Xlam
    End If

    ' Ð?t IsAddin = True: ?n sheet, chuy?n ch? d? add-in
    ThisWorkbook.IsAddin = True

    ' Luu thành .xlam (xlOpenXMLAddIn = 55)
    ThisWorkbook.SaveAs _
        fileName:=xlamPath, _
        fileFormat:=xlOpenXMLAddIn

    Application.DisplayAlerts = True

    ' Cài d?t vào danh sách add-in c?a Excel và b?t lên
    Dim ai As AddIn
    On Error Resume Next
    Set ai = Application.AddIns.Add(xlamPath)
    If Not ai Is Nothing Then ai.installed = True
    On Error GoTo EH_Xlam

    MsgBox "CPMS has been saved and installed as an Excel add-in." & vbCrLf & vbCrLf & _
           "File: " & xlamPath & vbCrLf & vbCrLf & _
           "CPMS will now load automatically every time Excel starts.", _
           vbInformation, "Add-In Installed"

    Exit Sub

EH_Xlam:
    Application.DisplayAlerts = True
    MsgBox "Could not convert to add-in." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "You can do this manually via: File > Save As > Excel Add-In (.xlam).", _
           vbExclamation, "Convert to Add-In Failed"

End Sub

Private Function ReadTemplateVersion() As String
    Dim vf As String
    vf = GetTemplateVersionFile()
    If Dir(vf) = "" Then Exit Function   ' Tr? v? ""
    Dim ff As Integer
    ff = FreeFile
    Open vf For Input As #ff
    If Not EOF(ff) Then Line Input #ff, ReadTemplateVersion
    Close #ff
    ReadTemplateVersion = Trim(ReadTemplateVersion)
End Function
Private Sub SaveTemplateVersion(Optional dummy As Variant = 0)
    Dim ff As Integer
    ff = FreeFile
    Open GetTemplateVersionFile() For Output As #ff
    Print #ff, TEMPLATE_VERSION
    Close #ff
End Sub
'==========================================================
' BUILD DATA TEMPLATE — version ?n d?nh
'==========================================================
Public Sub BuildDataTemplateWithProgress(Optional dummy As Variant = 0)

    Dim wbNew         As Workbook
    Dim wbTemp        As Workbook
    Dim wsSrc         As Worksheet
    Dim wsNew         As Worksheet
    Dim arrSheets     As Variant
    Dim i             As Long
    Dim prevVis       As XlSheetVisibility
    Dim pct           As Integer
    Dim wbCountBefore As Integer

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' -- Show loading --
    On Error Resume Next
    frmLoading.StartLoading "Building Data..."
    On Error GoTo 0

    arrSheets = ProjectTemplateSheets()

    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    If wbNew Is Nothing Then
        Err.Raise vbObjectError + 1, , "Could not create new workbook for template."
    End If

    Dim totalSheets As Long
    totalSheets = UBound(arrSheets) - LBound(arrSheets) + 1

    For i = LBound(arrSheets) To UBound(arrSheets)

        pct = 10 + CLng(80 * (i - LBound(arrSheets) + 1) / totalSheets)

        ' FIX: dùng SetProgress thay vì gán th?ng lblXxx.Caption
        ' SetProgress g?i ForceRepaint (toggle ScreenUpdating + DoEvents)
        ' nên form th?c s? v? l?i sau m?i bu?c.
        On Error Resume Next
        frmLoading.SetProgress pct, "Initializing application data..."
        On Error GoTo 0

        Set wsSrc = Nothing
        On Error Resume Next
        Set wsSrc = ThisWorkbook.Worksheets(arrSheets(i))
        On Error GoTo 0

        If Not wsSrc Is Nothing Then

            prevVis = wsSrc.Visible
            wsSrc.Visible = xlSheetVisible

            wbCountBefore = Workbooks.count

            On Error Resume Next
            wsSrc.Copy
            On Error GoTo 0

            wsSrc.Visible = prevVis

            If Workbooks.count > wbCountBefore Then

                Set wbTemp = Workbooks(Workbooks.count)

                On Error Resume Next
                wbTemp.sheets(1).Move _
                    After:=wbNew.Worksheets(wbNew.Worksheets.count)
                On Error GoTo EH

                Set wsNew = wbNew.Worksheets(wbNew.Worksheets.count)
                PrepareTemplateSheet wsNew

            End If

        End If

    Next i

    ' -- Validate sheet b?t bu?c --
    Dim missingSheet As Boolean
    missingSheet = False

    Dim requiredSheets As Variant
    requiredSheets = Array("PROJECT_SETTINGS", "CONTRACT_DB")

    Dim j As Long
    For j = 0 To UBound(requiredSheets)
        Set wsNew = Nothing
        On Error Resume Next
        Set wsNew = wbNew.Worksheets(requiredSheets(j))
        On Error GoTo 0
        If wsNew Is Nothing Then missingSheet = True
    Next j

    If missingSheet Then
        Err.Raise vbObjectError + 1, , _
            "Template build failed: required sheets are missing." & vbCrLf & _
            "Please verify source sheets exist in ThisWorkbook " & _
            "and are not protected."
    End If

    ' -- Xóa sheet blank d?u tiên --
    If wbNew.Worksheets.count > 1 Then
        wbNew.Worksheets(1).Delete
    End If

    ' -- Luu --
    On Error Resume Next
    frmLoading.SetProgress 95, "Saving..."
    On Error GoTo 0

    wbNew.SaveAs _
        fileName:=GetTemplateFile, _
        fileFormat:=xlExcel12

    wbNew.Close False

    On Error Resume Next
    frmLoading.SetProgress 100, "Complete!"
    On Error GoTo 0

    Application.Wait Now + TimeValue("00:00:01")

Cleanup:
    On Error Resume Next
    Unload frmLoading
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub

EH:
    Dim msg As String
    msg = Err.Description
    Resume Cleanup

    MsgBox "Template build error: " & msg, vbCritical

End Sub

'==========================================================
' VALIDATE — ki?m tra sheet b?t bu?c ph?i có
'==========================================================
Private Sub ValidateTemplateSheetsold(wb As Workbook)

    Dim required As Variant
    Dim i        As Long
    Dim missing  As String

    required = Array("PROJECT_SETTINGS", "CONTRACT_DB", "PAYMENT_LEDGER")

    For i = LBound(required) To UBound(required)

        Dim ws As Worksheet
        Set ws = Nothing

        On Error Resume Next
        Set ws = wb.Worksheets(required(i))
        On Error GoTo 0

        If ws Is Nothing Then
            missing = missing & "  - " & required(i) & vbCrLf
        End If

    Next i

    If Len(missing) > 0 Then
        ' Ðóng workbook không h?p l?
        Application.DisplayAlerts = False
        wb.Close False
        Application.DisplayAlerts = True

        Err.Raise vbObjectError + 1, , _
            "Template build incomplete. Missing sheets:" & vbCrLf & missing & _
            vbCrLf & "Please check if source sheets exist in ThisWorkbook."
    End If

End Sub
'==========================================================
' Ensure Project List
'==========================================================
Private Sub EnsureProjectList(Optional dummy As Variant = 0)

    If Dir(GetProjectListFile) <> "" Then Exit Sub

    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = Workbooks.Add(xlWBATWorksheet)

    Set ws = wb.Worksheets(1)

    ws.Name = "PROJECT_LIST"

    ws.Range("A1:H1").value = Array( _
        "ProjectID", _
        "PROJ_NAME", _
        "Location", _
        "ProjectFolder", _
        "DataFile", _
        "CreatedDate", _
        "LastOpened", _
        "Status")

    ws.rows(1).Font.Bold = True

    ws.Columns.AutoFit

    Application.DisplayAlerts = False

    wb.SaveAs _
        fileName:=GetProjectListFile, _
        fileFormat:=xlExcel12

    wb.Close False

    Application.DisplayAlerts = True

End Sub

'==========================================================
' Ensure Current Project File
'==========================================================
Private Sub EnsureCurrentProjectFile(Optional dummy As Variant = 0)

    If Dir(GetCurrentProjectFile) <> "" Then Exit Sub

    Dim ff As Integer

    ff = FreeFile

    Open GetCurrentProjectFile For Output As #ff

        Print #ff, "ProjectID="
        Print #ff, "PROJ_NAME="
        Print #ff, "ProjectFolder="
        Print #ff, "DataFile="

    Close #ff

End Sub

'==========================================================
' Helpers
'==========================================================
Public Function FileExists(ByVal fileName As String) As Boolean

    FileExists = (Dir(fileName) <> "")

End Function

Public Function FolderExists(ByVal folderName As String) As Boolean

    FolderExists = (Len(Dir(folderName, vbDirectory)) > 0)

End Function


Private Function ProjectTemplateSheets() As Variant

    ProjectTemplateSheets = Array( _
        "SYS_INFO", _
        "PROJECT_SETTINGS", _
        "CONTRACT_DB", _
        "PAYMENT_LEDGER", _
        "DOCUMENT_TRACKER", _
        "SETTINGS", _
        "LIBRARY_DB", _
        "AI_NOTES")

End Function
Private Sub PrepareTemplateSheet(ws As Worksheet)

    Dim shp As Shape

    ws.Visible = xlSheetVisible

    If ws.AutoFilterMode Then

        On Error Resume Next
        ws.ShowAllData
        On Error GoTo 0

    End If

    If ws.rows.count > 1 Then

        Dim lastRow As Long
        
        lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row
        
        If lastRow > 1 Then
            ws.rows("2:" & lastRow).ClearContents
        End If

    End If

    For Each shp In ws.Shapes
        shp.Delete
    Next shp

End Sub


'==========================================================
' Open Current Project
'==========================================================
Public Sub OpenCurrentProject(Optional dummy As Variant = 0)

    On Error GoTo EH

    Dim DataFile As String
    Dim step     As String

    step = "ReadCurrentProjectValue"
    DataFile = ReadCurrentProjectValue("DataFile")

    If Len(DataFile) = 0 Then
        step = "CreateNewProject"
        CreateNewProject
        Exit Sub
    End If

    If Dir(DataFile) = "" Then
        MsgBox "File not found: " & DataFile, vbExclamation
        LockWindowUpdate Application.hwnd
        frmHomePage.Show vbModeless
        LockWindowUpdate 0
        Exit Sub
    End If

    step = "OpenProjectWorkbook"
    OpenProjectWorkbook DataFile

    LockWindowUpdate Application.hwnd

    step = "Unload frmContractSearch"
    On Error Resume Next
    Unload frmContractSearch
    On Error GoTo EH

    step = "frmContractSearch.Show"
    frmContractSearch.Show vbModeless

    step = "frmContractSearch.InitFast"
    On Error Resume Next
    frmContractSearch.InitFast
    On Error GoTo EH

    LockWindowUpdate 0
    DoEvents

    Exit Sub

EH:
    LockWindowUpdate 0
    MsgBox "OpenCurrentProject failed at: [" & step & "]" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical

End Sub

Public Sub CreateNewProject(Optional dummy As Variant = 0)

    frmProjectInfo.Init "", ""
    frmProjectInfo.Show vbModal

    If frmProjectInfo.Cancelled Then
        LockWindowUpdate Application.hwnd
        gSuppressHomePageAutoCreate = True
        frmHomePage.Show vbModeless
        LockWindowUpdate 0
        Exit Sub
    End If

    LockWindowUpdate Application.hwnd

    CreateProject _
        frmProjectInfo.projName, _
        frmProjectInfo.location

    ' Sau t?o m?i ? quay v? frmHomePage (refresh list d? th?y project v?a t?o)
    gSuppressHomePageAutoCreate = True
    frmHomePage.Show vbModeless
    frmHomePage.RefreshList

    LockWindowUpdate 0

End Sub

Public Sub CreateProject( _
            ByVal PROJ_NAME As String, _
            ByVal location As String)

    Dim projectID As String
    projectID = GenerateProjectID()

    Dim projectFolder As String
    projectFolder = GetProjectsFolder & projectID & "\"
    MkDir projectFolder

    Dim newDataFile As String
    newDataFile = projectFolder & "CPMS_Data.xlsb"
    FileCopy GetTemplateFile, newDataFile

    ' *** QUAN TR?NG: Ðóng workbook c?a project cu tru?c khi m? file m?i.
    ' T?t c? project d?u dùng tên file "CPMS_Data.xlsb" — n?u 2 file cùng tên
    ' cùng m? m?t lúc, Workbooks.Open s? fail silently khi DisplayAlerts = False. ***
    Dim oldDataFile As String
    oldDataFile = ReadCurrentProjectValue("DataFile")
    If Len(oldDataFile) > 0 And LCase$(oldDataFile) <> LCase$(newDataFile) Then
        Dim wbOld As Workbook
        For Each wbOld In Workbooks
            If LCase$(wbOld.fullName) = LCase$(oldDataFile) Then
                Application.DisplayAlerts = False
                wbOld.Close False   ' Ðóng hidden workbook cu, không c?n save
                Application.DisplayAlerts = True
                Exit For
            End If
        Next wbOld
    End If

    ' Thêm vào ProjectList.xlsb
    AddProjectToList projectID, PROJ_NAME, location, projectFolder, newDataFile

    ' Luu CurrentProject.ini (ghi tru?c d? ReadCurrentProjectValue sau này dúng)
    SaveCurrentProject projectID, PROJ_NAME, projectFolder, newDataFile

    ' M? workbook m?i (không còn xung d?t tên vì dã dóng file cu)
    OpenProjectWorkbook newDataFile

    ' Ghi PROJECT_SETTINGS vào workbook dang m?
    WriteProjectInfo newDataFile, projectID, PROJ_NAME, location

End Sub
Private Function GenerateProjectID() As String

    GenerateProjectID = _
        "CP" & Format(Now, "yymmddhhmmss")

End Function

Public Sub SafeCopyDataFile(ByVal srcPath As String, ByVal destPath As String)
    Dim wbSrc As Workbook
    Dim isConverted As Boolean
    isConverted = False

    ' Nếu destPath đang mở, đóng lại trước
    Dim wb As Workbook
    For Each wb In Workbooks
        If LCase$(wb.fullName) = LCase$(destPath) Then
            Application.DisplayAlerts = False
            wb.Close False
            Application.DisplayAlerts = True
            Exit For
        End If
    Next wb

    ' Thử mở file nguồn bằng Excel để chuẩn hóa sang định dạng xlExcel12 (.xlsb)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wbSrc = Workbooks.Open(fileName:=srcPath, ReadOnly:=True)
    If Not wbSrc Is Nothing Then
        wbSrc.SaveAs fileName:=destPath, fileFormat:=50 ' 50 = xlExcel12 (.xlsb)
        wbSrc.Close False
        isConverted = True
    End If
    Err.Clear
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    ' Nếu mở trực tiếp không được hoặc có lỗi, fallback dùng FileCopy thông thường
    If Not isConverted Then
        On Error Resume Next
        If Dir(destPath) <> "" Then Kill destPath
        On Error GoTo 0
        FileCopy srcPath, destPath
    End If
End Sub

Public Sub OpenProjectWorkbook(ByVal fileName As String)

    Dim wb As Workbook

    ' *** QUAN TR?NG: Ðóng t?t c? CPMS_Data.xlsb dang m? mà KHÔNG ph?i file c?n m?.
    ' T?t c? project dùng cùng tên file ? ch? du?c phép có 1 file m? t?i m?t th?i di?m.
    ' Dùng vòng l?p + m?ng t?m vì không du?c duy?t và Close trong cùng For Each. ***
    Dim closeCandidates() As Workbook
    Dim nClose As Long
    nClose = 0

    For Each wb In Workbooks
        Dim wbShortName As String
        wbShortName = LCase$(mID$(wb.fullName, InStrRev(wb.fullName, "\") + 1))
        If wbShortName = "cpms_data.xlsb" Then
            If LCase$(wb.fullName) <> LCase$(fileName) Then
                nClose = nClose + 1
                ReDim Preserve closeCandidates(1 To nClose)
                Set closeCandidates(nClose) = wb
            End If
        End If
    Next wb

    Dim j As Long
    For j = 1 To nClose
        On Error Resume Next
        Application.DisplayAlerts = False
        closeCandidates(j).Close False
        Application.DisplayAlerts = True
        On Error GoTo 0
    Next j

    ' N?u file m?c tiêu dã m? dúng r?i thì không m? l?i
    For Each wb In Workbooks
        If LCase$(wb.fullName) = LCase$(fileName) Then
            On Error Resume Next
            wb.Windows(1).Visible = False
            On Error GoTo 0
            Exit Sub
        End If
    Next wb

    ' Mở file mới (kèm cơ chế tự động phục hồi nếu file bị mismatch format XLSX/XLSB)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(fileName)

    If wb Is Nothing Or Err.Number <> 0 Then
        Err.Clear
        ' Nếu file bị mismatch định dạng (ví dụ nội dung XLSX OpenXML nhưng mang đuôi .xlsb)
        ' Thử đổi tên tạm sang .xlsx để Excel nhận diện, mở ra và lưu lại chuẩn xlExcel12 (.xlsb)
        Dim tempRepairFile As String
        tempRepairFile = Left$(fileName, InStrRev(fileName, ".")) & "temp_repair.xlsx"
        On Error Resume Next
        If Dir(tempRepairFile) <> "" Then Kill tempRepairFile
        FileCopy fileName, tempRepairFile
        Err.Clear
        Set wb = Workbooks.Open(tempRepairFile)
        If Not wb Is Nothing Then
            wb.SaveAs fileName, 50 ' 50 = xlExcel12 (.xlsb)
            wb.Close False
            Kill tempRepairFile
            Err.Clear
            Set wb = Workbooks.Open(fileName)
        Else
            If Dir(tempRepairFile) <> "" Then Kill tempRepairFile
        End If
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    On Error Resume Next
    wb.Windows(1).Visible = False
    On Error GoTo 0

End Sub
'==========================================================
' SAVE CURRENT PROJECT — UTF-8
'==========================================================
Public Sub SaveCurrentProject( _
        ByVal projectID As String, _
        ByVal PROJ_NAME As String, _
        ByVal folder As String, _
        ByVal DataFile As String)

    Dim content As String
    content = "ProjectID=" & projectID & vbCrLf & _
              "PROJ_NAME=" & PROJ_NAME & vbCrLf & _
              "ProjectFolder=" & folder & vbCrLf & _
              "DataFile=" & DataFile & vbCrLf

    WriteUTF8 GetCurrentProjectFile(), content

End Sub

'==========================================================
' READ CURRENT PROJECT VALUE — UTF-8
'==========================================================
Public Function ReadCurrentProjectValue(ByVal key As String) As String

    Dim filePath As String

    On Error Resume Next
    filePath = GetCurrentProjectFile()
    On Error GoTo 0

    If Len(filePath) = 0 Or Dir(filePath) = "" Then Exit Function

    Dim content As String
    content = ReadUTF8(filePath)

    Dim lines()  As String
    Dim i        As Long
    Dim s        As String

    lines = Split(content, vbCrLf)

    For i = 0 To UBound(lines)
        s = lines(i)
        If LCase$(Left$(s, Len(key) + 1)) = LCase$(key & "=") Then
            ReadCurrentProjectValue = mID$(s, Len(key) + 2)
            Exit Function
        End If
    Next i

End Function
'==========================================================
' GHI FILE UTF-8 (h? tr? ti?ng Vi?t)
'==========================================================
Private Sub WriteUTF8(ByVal filePath As String, ByVal content As String)

    Dim st As Object
    Set st = CreateObject("ADODB.Stream")

    With st
        .Type = 2           ' adTypeText
        .Charset = "UTF-8"
        .Open
        .WriteText content
        .SaveToFile filePath, 2   ' adSaveCreateOverWrite
        .Close
    End With

End Sub

'==========================================================
' Ð?C FILE UTF-8
'==========================================================
Private Function ReadUTF8(ByVal filePath As String) As String

    If Dir(filePath) = "" Then Exit Function

    Dim st As Object
    Set st = CreateObject("ADODB.Stream")

    With st
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        ReadUTF8 = .ReadText
        .Close
    End With

End Function

Public Sub WriteProjectInfo( _
    ByVal DataFile As String, _
    ByVal projectID As String, _
    ByVal PROJ_NAME As String, _
    ByVal location As String)

    Dim wb          As Workbook
    Dim ws          As Worksheet
    Dim alreadyOpen As Boolean

    Application.ScreenUpdating = False

    Set wb = GetOrOpenWorkbook(DataFile, alreadyOpen)

    ' Guard: không m? du?c workbook
    If wb Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "WriteProjectInfo: Không m? du?c file d? li?u." & vbCrLf & DataFile, _
               vbCritical, "CPMS Error"
        Exit Sub
    End If

    ' Guard: sheet chua t?n t?i (template cu ho?c file l?i)
    On Error Resume Next
    Set ws = wb.Worksheets("PROJECT_SETTINGS")
    On Error GoTo 0
    If ws Is Nothing Then
        If Not alreadyOpen Then wb.Close False
        Application.ScreenUpdating = True
        MsgBox "WriteProjectInfo: Sheet 'PROJECT_SETTINGS' không tìm th?y trong:" & vbCrLf & DataFile, _
               vbCritical, "CPMS Error"
        Exit Sub
    End If

    SaveSettingField ws, "ProjectID", projectID
    SaveSettingField ws, "PROJ_NAME", PROJ_NAME
    SaveSettingField ws, "Location", location
    SaveSettingField ws, "CreatedDate", Format(Now, "yyyy-mm-dd")
    SaveSettingField ws, "LastOpened", Format(Now, "yyyy-mm-dd")
    SaveSettingField ws, "DataVersion", "1.0"

    wb.Save
    If Not alreadyOpen Then wb.Close False

    Application.ScreenUpdating = True

End Sub
Private Sub SaveSettingField( _
            ws As Worksheet, _
            ByVal fieldName As String, _
            ByVal FieldValue As String)

    Dim f As Range

    Set f = ws.Columns(1).Find( _
                What:=fieldName, _
                LookAt:=xlWhole, _
                LookIn:=xlValues)

    If f Is Nothing Then

        Dim r As Long

        r = ws.Cells(ws.rows.count, 1).End(xlUp).Row + 1

        ws.Cells(r, 1).value = fieldName
        ws.Cells(r, 2).value = FieldValue

    Else

        f.Offset(0, 1).value = FieldValue

    End If

End Sub
Public Sub AddProjectToList( _
            ByVal projectID As String, _
            ByVal PROJ_NAME As String, _
            ByVal location As String, _
            ByVal projectFolder As String, _
            ByVal DataFile As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim r As Long

    Application.ScreenUpdating = False

    Set wb = Workbooks.Open(GetProjectListFile)

    Set ws = wb.Worksheets("PROJECT_LIST")

    r = ws.Cells(ws.rows.count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).value = projectID
    ws.Cells(r, 2).value = PROJ_NAME
    ws.Cells(r, 3).value = location
    ws.Cells(r, 4).value = projectFolder
    ws.Cells(r, 5).value = DataFile
    ws.Cells(r, 6).value = Now
    ws.Cells(r, 7).value = Now
    ws.Cells(r, 8).value = "Active"

    wb.Save
    wb.Close False

    Application.ScreenUpdating = True

End Sub

'==========================================================
' RENAME CURRENT PROJECT
'
' - Ð?i tên trong ProjectList.xlsb
' - Ð?i tên trong PROJECT_SETTINGS c?a data file
' - C?p nh?t CurrentProject.ini
'==========================================================
Public Sub RenameCurrentProject(Optional ByVal projectID As String = "")

    Dim oldName       As String
    Dim OldLocation   As String
    Dim DataFile      As String
    Dim projectFolder As String

    If Len(Trim(projectID)) = 0 Then
        projectID = ReadCurrentProjectValue("ProjectID")
        oldName = ReadCurrentProjectValue("PROJ_NAME")
        DataFile = ReadCurrentProjectValue("DataFile")
        projectFolder = ReadCurrentProjectValue("ProjectFolder")
        OldLocation = GetProjectFieldFromList(projectID, 3)
    Else
        oldName = GetProjectFieldFromList(projectID, 2)
        OldLocation = GetProjectFieldFromList(projectID, 3)
        projectFolder = GetProjectFieldFromList(projectID, 4)
        DataFile = GetProjectFieldFromList(projectID, 5)
    End If

    If Len(projectID) = 0 Then
        MsgBox "No project is currently open.", vbExclamation
        Exit Sub
    End If

    frmProjectInfo.InitEdit oldName, OldLocation
    frmProjectInfo.Show vbModal

    If frmProjectInfo.Cancelled Then Exit Sub

    Dim newName     As String
    Dim NewLocation As String
    newName = frmProjectInfo.projName
    NewLocation = frmProjectInfo.location

    If newName = oldName And NewLocation = OldLocation Then
        MsgBox "No changes made.", vbInformation
        Exit Sub
    End If

    ' C?p nh?t ProjectList (Name + Location)
    UpdateProjectFieldsInList projectID, newName, NewLocation

    ' C?p nh?t PROJECT_SETTINGS trong data file
    If Dir(DataFile) <> "" Then

        Application.ScreenUpdating = False

        ' *** FIX: Ðóng workbook nào dang chi?m tên "CPMS_Data.xlsb" (n?u không ph?i
        '     chính file c?n rename) d? Workbooks.Open không b? conflict ***
        Dim wbConflict As Workbook
        Dim reopenFile As String   ' workbook nào b? dóng t?m ? m? l?i sau

        For Each wbConflict In Workbooks
            If LCase$(mID$(wbConflict.fullName, InStrRev(wbConflict.fullName, "\") + 1)) _
                    = "cpms_data.xlsb" Then
                If LCase$(wbConflict.fullName) <> LCase$(DataFile) Then
                    reopenFile = wbConflict.fullName   ' nh? d? m? l?i sau
                    Application.DisplayAlerts = False
                    wbConflict.Close False
                    Application.DisplayAlerts = True
                    Exit For
                End If
            End If
        Next wbConflict

        ' M? file c?n rename (không còn conflict tên)
        Dim wb          As Workbook
        Dim ws          As Worksheet
        Dim alreadyOpen As Boolean

        Set wb = GetOrOpenWorkbook(DataFile, alreadyOpen)

        If wb Is Nothing Then
            Application.ScreenUpdating = True
            ' M? l?i workbook v?a dóng (n?u có) tru?c khi báo l?i
            If Len(reopenFile) > 0 Then
                On Error Resume Next
                OpenProjectWorkbook reopenFile
                On Error GoTo 0
            End If
            MsgBox "RenameCurrentProject: Không m? du?c file d? li?u." & vbCrLf & _
                   DataFile, vbCritical, "CPMS Error"
            Exit Sub
        End If

        On Error Resume Next
        Set ws = wb.Worksheets("PROJECT_SETTINGS")
        On Error GoTo 0

        If Not ws Is Nothing Then
            SaveSettingField ws, "PROJ_NAME", newName
            SaveSettingField ws, "Location", NewLocation
            wb.Save
        End If

        If Not alreadyOpen Then wb.Close False

        ' M? l?i workbook dã dóng t?m d? nhu?ng ch? (n?u có)
        If Len(reopenFile) > 0 Then
            On Error Resume Next
            OpenProjectWorkbook reopenFile
            On Error GoTo 0
        End If

        Application.ScreenUpdating = True

    End If

    ' Ch? c?p nh?t CurrentProject.ini n?u dây CHÍNH LÀ project dang m?
    If ReadCurrentProjectValue("ProjectID") = projectID Then
        SaveCurrentProject projectID, newName, projectFolder, DataFile
    End If

    MsgBox "Project updated successfully.", vbInformation

End Sub

'==========================================================
' L?y giá tr? 1 field t? ProjectList theo column index
'==========================================================
Public Function GetProjectFieldFromList( _
                    ByVal projectID As String, _
                    ByVal ColIndex As Long) As String

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim f  As Range

    On Error GoTo EH

    Application.ScreenUpdating = False

    Set wb = Workbooks.Open(GetProjectListFile)
    Set ws = wb.Worksheets("PROJECT_LIST")

    Set f = ws.Columns(1).Find( _
                What:=projectID, LookAt:=xlWhole, LookIn:=xlValues)

    If Not f Is Nothing Then
        GetProjectFieldFromList = f.Offset(0, ColIndex - 1).value
    End If

    wb.Close False
    Application.ScreenUpdating = True
    Exit Function

EH:
    If Not wb Is Nothing Then wb.Close False
    Application.ScreenUpdating = True

End Function

'==========================================================
' C?p nh?t Name + Location trong ProjectList
'==========================================================
Private Sub UpdateProjectFieldsInList( _
                ByVal projectID As String, _
                ByVal newName As String, _
                ByVal NewLocation As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim f  As Range

    On Error GoTo EH

    Application.ScreenUpdating = False

    Set wb = Workbooks.Open(GetProjectListFile)
    Set ws = wb.Worksheets("PROJECT_LIST")

    Set f = ws.Columns(1).Find( _
                What:=projectID, LookAt:=xlWhole, LookIn:=xlValues)

    If Not f Is Nothing Then
        f.Offset(0, 1).value = newName      ' Col B = PROJ_NAME
        f.Offset(0, 2).value = NewLocation  ' Col C = Location
    End If

    wb.Save
    wb.Close False
    Application.ScreenUpdating = True
    Exit Sub

EH:
    If Not wb Is Nothing Then wb.Close False
    Application.ScreenUpdating = True
    Err.Raise Err.Number, , "UpdateProjectFieldsInList: " & Err.Description

End Sub

'==========================================================
' Update PROJ_NAME in ProjectList.xlsb
'==========================================================
Private Sub UpdateProjectNameInList( _
            ByVal projectID As String, _
            ByVal newName As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim f  As Range

    On Error GoTo EH

    Application.ScreenUpdating = False

    Set wb = Workbooks.Open(GetProjectListFile)
    Set ws = wb.Worksheets("PROJECT_LIST")

    Set f = ws.Columns(1).Find( _
                What:=projectID, _
                LookAt:=xlWhole, _
                LookIn:=xlValues)

    If Not f Is Nothing Then
        f.Offset(0, 1).value = newName   ' Column B = ProjectName
    End If

    wb.Save
    wb.Close False

    Application.ScreenUpdating = True

    Exit Sub

EH:
    If Not wb Is Nothing Then wb.Close False
    Application.ScreenUpdating = True
    Err.Raise Err.Number, , "UpdateProjectNameInList: " & Err.Description

End Sub
'==========================================================
' CLONE CURRENT PROJECT
'
' - T?o ProjectID m?i
' - T?o folder m?i trong Projects\
' - Copy data file (dóng t?m n?u dang m?)
' - C?p nh?t PROJECT_SETTINGS trong b?n sao
' - Thêm vào ProjectList.xlsb
' - KHÔNG thay d?i CurrentProject.ini
'==========================================================
Public Sub CloneCurrentProject(Optional ByVal projectID As String = "")

    Dim SrcProjectID   As String
    Dim SrcName        As String
    Dim SrcDataFile    As String
    Dim SrcLocation    As String

    If Len(Trim(projectID)) = 0 Then
        SrcProjectID = ReadCurrentProjectValue("ProjectID")
        SrcName = ReadCurrentProjectValue("PROJ_NAME")
        SrcDataFile = ReadCurrentProjectValue("DataFile")
    Else
        SrcProjectID = projectID
        SrcName = GetProjectFieldFromList(projectID, 2)
        SrcDataFile = GetProjectFieldFromList(projectID, 5)
    End If

    If Len(SrcProjectID) = 0 Then
        MsgBox "No project is currently open.", vbExclamation
        Exit Sub
    End If

    If Dir(SrcDataFile) = "" Then
        MsgBox "Source data file not found:" & vbCrLf & SrcDataFile, vbCritical
        Exit Sub
    End If

    SrcLocation = GetProjectFieldFromList(SrcProjectID, 3)

    frmProjectInfo.InitEdit SrcName & " (Copy)", SrcLocation
    frmProjectInfo.Show vbModal
    If frmProjectInfo.Cancelled Then Exit Sub

    Dim newName     As String
    Dim NewLocation As String
    newName = Trim(frmProjectInfo.projName)
    NewLocation = Trim(frmProjectInfo.location)

    If Len(newName) = 0 Then
        MsgBox "Project name cannot be empty.", vbExclamation
        Exit Sub
    End If

    Dim NewProjectID     As String
    Dim NewProjectFolder As String
    Dim newDataFile      As String

    NewProjectID = GenerateProjectID()
    NewProjectFolder = GetProjectsFolder & NewProjectID & "\"
    newDataFile = NewProjectFolder & "CPMS_Data.xlsb"

    MkDir NewProjectFolder

    ' *** FIX 1: Tìm workbook ngu?n b?ng FULL PATH (không ph?i tên ng?n) ***
    Dim wbSrc      As Workbook
    Dim srcWasOpen As Boolean

    For Each wbSrc In Workbooks
        If LCase$(wbSrc.fullName) = LCase$(SrcDataFile) Then
            srcWasOpen = True
            Exit For
        End If
    Next wbSrc

    ' Ðóng workbook ngu?n d? FileCopy không b? lock
    If srcWasOpen Then
        Application.DisplayAlerts = False
        wbSrc.Save
        wbSrc.Close False
        Application.DisplayAlerts = True
        Set wbSrc = Nothing
    End If

    ' Sao chép file
    On Error GoTo EH_Copy
    FileCopy SrcDataFile, newDataFile
    On Error GoTo 0

    ' *** FIX 2: M? file M?I tru?c khi m? l?i ngu?n ***
    ' Th? t? cu: re-open source ? WriteProjectInfo(new) ? name conflict!
    ' Th? t? m?i: open new ? write ? close new ? re-open source (không conflict)

    OpenProjectWorkbook newDataFile
    WriteProjectInfo newDataFile, NewProjectID, newName, NewLocation

    ' Ðóng file m?i (dã ghi xong), r?i m?i m? l?i ngu?n
    Dim wbNew As Workbook
    For Each wbNew In Workbooks
        If LCase$(wbNew.fullName) = LCase$(newDataFile) Then
            Application.DisplayAlerts = False
            wbNew.Close False
            Application.DisplayAlerts = True
            Exit For
        End If
    Next wbNew

    ' M? l?i workbook ngu?n n?u tru?c dó dang m?
    If srcWasOpen Then
        On Error Resume Next
        OpenProjectWorkbook SrcDataFile
        On Error GoTo 0
    End If

    AddProjectToList _
        NewProjectID, newName, NewLocation, NewProjectFolder, newDataFile

    Dim msg As String
    msg = "Project cloned successfully!" & vbCrLf & vbCrLf & _
          "New Project : " & newName & vbCrLf & _
          "ID          : " & NewProjectID & vbCrLf
    ShowMsgCenter msg

    ' *** FIX 3: B? frmContractSearch.InitFast ***
    ' V?i lu?ng m?i, sau clone v?n ? frmHomePage — không c?n refresh frmContractSearch.
    ' G?i InitFast trên form chua kh?i t?o (OpenSession chua ch?y) gây error 800400c9.

    Exit Sub

EH_Copy:
    If srcWasOpen Then
        On Error Resume Next
        OpenProjectWorkbook SrcDataFile
        On Error GoTo 0
    End If
    MsgBox "Clone failed — could not copy data file." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical

End Sub

'==========================================================
' Read a single field from PROJECT_SETTINGS of any data file
' Used internally — does not change CurrentProject.ini
'==========================================================
Private Function ReadProjectSettingField( _
                    ByVal DataFile As String, _
                    ByVal fieldName As String) As String

    Dim wb          As Workbook
    Dim ws          As Worksheet
    Dim f           As Range
    Dim alreadyOpen As Boolean

    On Error GoTo EH

    Application.ScreenUpdating = False

    Set wb = GetOrOpenWorkbook(DataFile, alreadyOpen)
    Set ws = wb.Worksheets("PROJECT_SETTINGS")

    Set f = ws.Columns(1).Find( _
                What:=fieldName, _
                LookAt:=xlWhole, _
                LookIn:=xlValues)

    If Not f Is Nothing Then
        ReadProjectSettingField = f.Offset(0, 1).value
    End If

    If Not alreadyOpen Then wb.Close False

    Application.ScreenUpdating = True
    Exit Function

EH:
    If Not wb Is Nothing Then
        If Not alreadyOpen Then wb.Close False
    End If
    Application.ScreenUpdating = True

End Function

'==========================================================
' GET OR OPEN WORKBOOK — không h?i l?i n?u dã m?
' Returns: wb = workbook reference, wasAlreadyOpen = True n?u dã m? s?n
'==========================================================
Public Function GetOrOpenWorkbook(ByVal fileName As String, _
                                   Optional ByRef wasAlreadyOpen As Boolean = False) As Workbook

    Dim wb As Workbook

    ' Tìm b?ng FULL PATH d? không nh?m khi 2 project cùng tên file
    For Each wb In Workbooks
        If LCase$(wb.fullName) = LCase$(fileName) Then
            wasAlreadyOpen = True
            Set GetOrOpenWorkbook = wb
            Exit Function
        End If
    Next wb

    ' Chua m? — m? m?i, có error handling
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(fileName)
    Application.DisplayAlerts = True
    On Error GoTo 0

    wasAlreadyOpen = False
    Set GetOrOpenWorkbook = wb   ' Tr? v? Nothing n?u Open th?t b?i (caller t? x? lý)

End Function

'==========================================================
' Helper — d?c 1 field t? PROJECT_SETTINGS trong file b?t k?
' (dùng riêng cho Import, không thay d?i CurrentProject.ini)
'==========================================================
Private Function ReadFieldFromFile(ByVal filePath As String, _
                                    ByVal fieldName As String) As String

    If Dir(filePath) = "" Then Exit Function

    Dim wb          As Workbook
    Dim ws          As Worksheet
    Dim f           As Range
    Dim alreadyOpen As Boolean

    On Error Resume Next
    Set wb = GetOrOpenWorkbook(filePath, alreadyOpen)
    If wb Is Nothing Then Exit Function

    Set ws = wb.Worksheets("PROJECT_SETTINGS")
    If ws Is Nothing Then
        If Not alreadyOpen Then wb.Close False
        Exit Function
    End If

    Set f = ws.Columns(1).Find(What:=fieldName, LookAt:=xlWhole, LookIn:=xlValues)
    If Not f Is Nothing Then ReadFieldFromFile = f.Offset(0, 1).value
    On Error GoTo 0

    If Not alreadyOpen Then wb.Close False

End Function


'==========================================================
' IMPORT TU FILE .XLSX / .XLSM
' - Mo workbook nguon
' - Tam thoi unhide T?T CA sheet (ke ca xlSheetVeryHidden)
' - SaveAs .xlsb de dung lam nguon import
' - Goi ImportProjectFromFile
'==========================================================
Public Sub ImportProjectFromExcel(ByVal sourcePath As String)

    If Len(sourcePath) = 0 Or Dir(sourcePath) = "" Then
        MsgBox "File not found: " & vbCrLf & sourcePath, vbCritical
        Exit Sub
    End If

    Dim wbSrc    As Workbook
    Dim wbClean  As Workbook
    Dim tempXlsb As String
    Dim tempCpms As String

    On Error GoTo EH_Excel

    Application.ScreenUpdating = False
    Application.Cursor = xlDefault
    Application.DisplayAlerts = False
    Application.EnableEvents = False   ' ngan xlsm tu dong very-hide khi Open

    ' Mo file nguon (ReadOnly de an toan)
    Set wbSrc = Workbooks.Open(sourcePath, _
                               UpdateLinks:=False, _
                               ReadOnly:=True, _
                               IgnoreReadOnlyRecommended:=True)

    ' Unhide TAT CA sheets (ke ca xlSheetVeryHidden)
    Dim sh As Worksheet
    For Each sh In wbSrc.Worksheets
        On Error Resume Next
        sh.Visible = xlSheetVisible
        On Error GoTo EH_Excel
    Next sh

    ' Doc project metadata truoc khi copy
    Dim projName As String: projName = ""
    Dim location As String: location = ""
    On Error Resume Next
    Dim wsPS As Worksheet
    Set wsPS = wbSrc.Worksheets("PROJECT_SETTINGS")
    If Not wsPS Is Nothing Then
        Dim r As Long
        For r = 1 To wsPS.Cells(wsPS.rows.count, 1).End(xlUp).Row
            Dim k As String
            k = UCase(Trim(CStr(wsPS.Cells(r, 1).value)))
            If k = "PROJ_NAME" Then projName = Trim(CStr(wsPS.Cells(r, 2).value))
            If k = "LOCATION" Then location = Trim(CStr(wsPS.Cells(r, 2).value))
        Next r
    End If
    Set wsPS = Nothing
    On Error GoTo EH_Excel

    '--------------------------------------------------------------
    ' COPY TUNG SHEET DATA sang workbook SACH (khong co VBA code)
    ' -> Tranh crash khi SaveCopyAs convert xlsm -> xlsb
    '--------------------------------------------------------------
    ' Copy sheet dau tien de tao workbook moi
    wbSrc.Worksheets(1).Copy   ' Excel tu tao ActiveWorkbook moi
    Set wbClean = ActiveWorkbook

    ' Copy cac sheet con lai
    Dim i As Long
    For i = 2 To wbSrc.Worksheets.count
        On Error Resume Next
        wbSrc.Worksheets(i).Copy After:=wbClean.sheets(wbClean.sheets.count)
        On Error GoTo EH_Excel
    Next i

    ' Dong workbook nguon (da copy xong)
    wbSrc.Close False
    Set wbSrc = Nothing

    ' Luu workbook sach thanh .xlsb (khong co VBA -> khong crash)
    tempXlsb = Environ("TEMP") & "\cpms_import_" & _
               Format(Now, "yyyymmddHHmmss") & ".xlsb"
    wbClean.SaveAs tempXlsb, xlExcel12
    wbClean.Close False
    Set wbClean = Nothing

    ' Khoi phuc Excel state
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True

    ' Doi sang .cpms
    tempCpms = Left$(tempXlsb, Len(tempXlsb) - 4) & "cpms"
    FileCopy tempXlsb, tempCpms
    Kill tempXlsb

    ' Import voi metadata da doc duoc (pre-fill dialog)
    ImportProjectFromFileWithMeta tempCpms, projName, location

    On Error Resume Next: Kill tempCpms: On Error GoTo 0
    Exit Sub

EH_Excel:
    Dim errMsg As String: errMsg = Err.Description
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    If Not wbSrc Is Nothing Then
        On Error Resume Next: wbSrc.Close False: On Error GoTo 0
    End If
    If Not wbClean Is Nothing Then
        On Error Resume Next: wbClean.Close False: On Error GoTo 0
    End If
    On Error Resume Next
    If Len(tempXlsb) > 0 Then Kill tempXlsb
    If Len(tempCpms) > 0 Then Kill tempCpms
    On Error GoTo 0
    MsgBox "Excel import failed: " & errMsg, vbCritical

End Sub

'==========================================================
' IMPORT TU .CPMS — co the pre-fill ten/location tu metadata
'==========================================================
Public Sub ImportProjectFromFileWithMeta( _
        ByVal sourcePath As String, _
        Optional ByVal prefilledName As String = "", _
        Optional ByVal prefilledLocation As String = "")

    If Len(sourcePath) = 0 Or Dir(sourcePath) = "" Then
        MsgBox "File not found: " & vbCrLf & sourcePath, vbCritical
        Exit Sub
    End If

    ' T? d?c metadata t? file n?u không du?c truy?n t? bên ngoài
    If Len(Trim(prefilledName)) = 0 Then
        prefilledName = ReadFieldFromFile(sourcePath, "PROJ_NAME")
    End If
    If Len(Trim(prefilledLocation)) = 0 Then
        prefilledLocation = ReadFieldFromFile(sourcePath, "Location")
    End If

    ' Hi?n dialog d? user xác nh?n / s?a tên
    frmProjectInfo.Init prefilledName, prefilledLocation
    frmProjectInfo.Show vbModal
    If frmProjectInfo.Cancelled Then Exit Sub

    Dim projName As String: projName = frmProjectInfo.projName
    Dim location As String: location = frmProjectInfo.location

    '----------------------------------------------------------
    '  KI?M TRA TRÙNG TÊN PROJECT
    '----------------------------------------------------------
    Dim matchID       As String
    Dim matchFolder   As String
    Dim matchDataFile As String

    If FindProjectByNameInList(projName, matchID, matchFolder, matchDataFile) Then

        ' --- Ð?c th?ng kê file import (m? t?m read-only) ---
        Dim impCon As Long, impPay As Long, impDoc As Long
        Dim impConDate As String, impPayDate As String, impDocDate As String

        Dim wbImp As Workbook
        On Error Resume Next
        Application.DisplayAlerts = False
        Set wbImp = Workbooks.Open(sourcePath, ReadOnly:=True)
        Application.DisplayAlerts = True
        On Error GoTo 0

        If Not wbImp Is Nothing Then
            GetSheetStats wbImp, "CONTRACT_DB", impCon, impConDate
            GetSheetStats wbImp, "PAYMENT_LEDGER", impPay, impPayDate
            GetSheetStats wbImp, "DOCUMENT_TRACKER", impDoc, impDocDate
            wbImp.Close False
            Set wbImp = Nothing
        End If

        ' --- Ð?c th?ng kê project hi?n có ---
        Dim exCon As Long, exPay As Long, exDoc As Long
        Dim exConDate As String, exPayDate As String, exDocDate As String

        Dim wbEx As Workbook
        Dim alreadyOpenEx As Boolean

        ' Ðóng conflict tru?c khi m? existing project
        CloseConflictingDataFile matchDataFile

        Set wbEx = GetOrOpenWorkbook(matchDataFile, alreadyOpenEx)
        If Not wbEx Is Nothing Then
            GetSheetStats wbEx, "CONTRACT_DB", exCon, exConDate
            GetSheetStats wbEx, "PAYMENT_LEDGER", exPay, exPayDate
            GetSheetStats wbEx, "DOCUMENT_TRACKER", exDoc, exDocDate
            If Not alreadyOpenEx Then wbEx.Close False
            Set wbEx = Nothing
        End If

        ' --- B?ng so sánh ---
        Dim msg As String
        msg = "Project """ & projName & """ already exists!" & vbCrLf & vbCrLf & _
              "                    Existing       Import file" & vbCrLf & _
                String(49, "-") & vbCrLf & _
              "Contracts:      " & rj(exCon, 7) & "       " & rj(impCon, 7) & vbCrLf & _
              "  Latest:    " & PadR(exConDate, 13) & "  " & impConDate & vbCrLf & _
              "Payments:       " & rj(exPay, 7) & "       " & rj(impPay, 7) & vbCrLf & _
              "  Latest:    " & PadR(exPayDate, 13) & "  " & impPayDate & vbCrLf & _
              "Documents:      " & rj(exDoc, 7) & "       " & rj(impDoc, 7) & vbCrLf & _
              "  Latest:    " & PadR(exDocDate, 13) & "  " & impDocDate & vbCrLf & _
               String(49, "-") & vbCrLf & vbCrLf & _
              "YES    - Replace existing project data" & vbCrLf & _
              "           (old file auto-backed up)" & vbCrLf & _
              "NO     - Import as a NEW separate project" & vbCrLf & _
              "CANCEL - Abort"

        Dim ans As Long
        ans = UniMsgBox(msg, vbYesNoCancel + vbQuestion, "Import - Duplicate Project Name")

        Select Case ans
            Case vbCancel
                Exit Sub

            Case vbYes
                ReplaceExistingProjectData _
                    matchID, matchFolder, matchDataFile, _
                    sourcePath, projName, location
                On Error Resume Next
                If frmHomePage.Visible Then frmHomePage.RefreshList
                On Error GoTo 0
                Exit Sub

            Case vbNo
                ' Ti?p t?c import thành project m?i (fall through)
        End Select

    End If

    '----------------------------------------------------------
    '  IMPORT THÀNH PROJECT M?I
    '  (không trùng tên, ho?c user ch?n NO ? trên)
    '----------------------------------------------------------
    CreateProject projName, location

    Dim newDataFile As String
    newDataFile = ReadCurrentProjectValue("DataFile")

    If Len(newDataFile) = 0 Then
        MsgBox "Failed to create project structure.", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    ResetSessionRef
    On Error GoTo 0

    ' Ðóng workbook template v?a t?o tru?c khi overwrite b?ng source
    Dim wbCheck As Workbook
    For Each wbCheck In Workbooks
        If LCase$(wbCheck.fullName) = LCase$(newDataFile) Then
            Application.DisplayAlerts = False
            wbCheck.Close False
            Application.DisplayAlerts = True
            Exit For
        End If
    Next wbCheck

    On Error GoTo EH_Meta
    SafeCopyDataFile sourcePath, newDataFile
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.Cursor = xlDefault
    OpenProjectWorkbook newDataFile

    On Error Resume Next
    gContractLoaded = False
    gPaymentLoaded = False
    gDocLoaded = False
    BuildGlobalContractCache
    BuildPaymentCache
    BuildDocTrackerCache
    On Error GoTo 0

    On Error Resume Next
    If frmHomePage.Visible Then frmHomePage.Hide
    On Error GoTo 0

    On Error Resume Next
    Unload frmContractSearch
    Set frmContractSearch = Nothing
    On Error GoTo 0

    frmContractSearch.Show vbModeless
    frmContractSearch.InitFast

    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    DoEvents
    Exit Sub

EH_Meta:
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    MsgBox "Import failed: " & Err.Description, vbCritical

End Sub

Public Sub ImportProjectFromFileWithMetaold( _
        ByVal sourcePath As String, _
        Optional ByVal prefilledName As String = "", _
        Optional ByVal prefilledLocation As String = "")

    If Len(sourcePath) = 0 Or Dir(sourcePath) = "" Then
        MsgBox "File not found: " & vbCrLf & sourcePath, vbCritical
        Exit Sub
    End If

    ' T? d?c metadata t? file n?u không du?c truy?n t? bên ngoài
    If Len(Trim(prefilledName)) = 0 Then
        prefilledName = ReadFieldFromFile(sourcePath, "PROJ_NAME")
    End If
    If Len(Trim(prefilledLocation)) = 0 Then
        prefilledLocation = ReadFieldFromFile(sourcePath, "Location")
    End If

    ' Hi?n th? dialog v?i tên/location dã du?c di?n s?n
    frmProjectInfo.Init prefilledName, prefilledLocation
    frmProjectInfo.Show vbModal
    If frmProjectInfo.Cancelled Then Exit Sub

    Dim projName As String: projName = frmProjectInfo.projName
    Dim location As String: location = frmProjectInfo.location

    ' --- Ph?n còn l?i GI? NGUYÊN t? dây ---
    CreateProject projName, location

    Dim newDataFile As String
    newDataFile = ReadCurrentProjectValue("DataFile")

    If Len(newDataFile) = 0 Then
        MsgBox "Failed to create project structure.", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    ResetSessionRef
    On Error GoTo 0

    Dim wbName As String
    wbName = mID$(newDataFile, InStrRev(newDataFile, "\") + 1)
    Dim wbCheck As Workbook
    On Error Resume Next
    Set wbCheck = Workbooks(wbName)
    On Error GoTo 0
    If Not wbCheck Is Nothing Then
        Application.DisplayAlerts = False
        wbCheck.Close False
        Application.DisplayAlerts = True
        Set wbCheck = Nothing
    End If

    On Error GoTo EH_Meta
    SafeCopyDataFile sourcePath, newDataFile
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.Cursor = xlDefault
    OpenProjectWorkbook newDataFile

    On Error Resume Next
    gContractLoaded = False
    gPaymentLoaded = False
    gDocLoaded = False
    BuildGlobalContractCache
    BuildPaymentCache
    BuildDocTrackerCache
    On Error GoTo 0

    On Error Resume Next
    If frmHomePage.Visible Then frmHomePage.Hide
    On Error GoTo 0

    On Error Resume Next
    Unload frmContractSearch
    Set frmContractSearch = Nothing
    On Error GoTo 0

    frmContractSearch.Show vbModeless
    frmContractSearch.InitFast

    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    DoEvents

    Exit Sub

EH_Meta:
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    MsgBox "Import failed: " & Err.Description, vbCritical

End Sub

'==========================================================
' ENSURE PROJECT INFO IN SHEET
' Ghi PROJ_NAME/Location vào PROJECT_SETTINGS c?a data workbook
' n?u PROJ_NAME còn tr?ng. G?i khi Open project d? d?ng b?
' d? li?u t? ProjectList vào sheet (phòng tru?ng h?p sheet cu/r?ng).
'==========================================================
Public Sub EnsureProjectInfoInSheet( _
            ByVal DataFile As String, _
            ByVal projectID As String, _
            ByVal PROJ_NAME As String, _
            ByVal location As String)

    Dim wb          As Workbook
    Dim alreadyOpen As Boolean

    On Error Resume Next
    Set wb = GetOrOpenWorkbook(DataFile, alreadyOpen)
    On Error GoTo 0
    If wb Is Nothing Then Exit Sub

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("PROJECT_SETTINGS")
    On Error GoTo 0
    If ws Is Nothing Then
        If Not alreadyOpen Then wb.Close False
        Exit Sub
    End If

    ' Ki?m tra PROJ_NAME hi?n t?i trong sheet
    Dim i As Long, lastR As Long, curName As String
    lastR = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    For i = 1 To lastR
        If UCase$(Trim$(CStr(ws.Cells(i, 1).value))) = "PROJ_NAME" Then
            curName = Trim$(CStr(ws.Cells(i, 2).value))
            Exit For
        End If
    Next i

    ' N?u dã có tên thì không c?n ghi dè
    If curName <> "" Then
        If Not alreadyOpen Then wb.Close False
        Exit Sub
    End If

    ' Ghi d?y d? qua WriteProjectInfo (tái dùng logic SaveSettingField + Save)
    If Not alreadyOpen Then wb.Close False   ' WriteProjectInfo t? m? l?i
    WriteProjectInfo DataFile, projectID, PROJ_NAME, location

End Sub

'==========================================================
' TÌM PROJECT THEO TÊN TRONG PROJECTLIST.XLSB
'==========================================================
Private Function FindProjectByNameInList( _
                    ByVal searchName As String, _
                    ByRef outID As String, _
                    ByRef outFolder As String, _
                    ByRef outDataFile As String) As Boolean
    outID = "": outFolder = "": outDataFile = ""
    FindProjectByNameInList = False

    Dim listFile As String: listFile = GetProjectListFile()
    If Dir(listFile) = "" Then Exit Function

    Dim wb As Workbook, ws As Worksheet, alreadyOpen As Boolean
    Set wb = GetOrOpenWorkbook(listFile, alreadyOpen)
    If wb Is Nothing Then Exit Function

    On Error Resume Next
    Set ws = wb.Worksheets("PROJECT_LIST")
    On Error GoTo 0

    If ws Is Nothing Then
        If Not alreadyOpen Then wb.Close False
        Exit Function
    End If

    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row

    For i = 2 To lastRow
        ' Col A=ID | B=PROJ_NAME | C=Location | D=Folder | E=DataFile
        If LCase$(Trim$(CStr(ws.Cells(i, 2).value))) = LCase$(Trim$(searchName)) Then
            outID = CStr(ws.Cells(i, 1).value)
            outFolder = CStr(ws.Cells(i, 4).value)
            outDataFile = CStr(ws.Cells(i, 5).value)
            FindProjectByNameInList = True
            Exit For
        End If
    Next i

    If Not alreadyOpen Then wb.Close False
End Function

'==========================================================
' Ð?C S? DÒNG DATA + NGÀY M?I NH?T T? 1 SHEET
' T? d?ng tìm c?t date b?ng scan header (không hardcode)
'==========================================================
Private Sub GetSheetStats(ByVal wb As Workbook, _
                          ByVal sheetName As String, _
                          ByRef rowCount As Long, _
                          ByRef latestStr As String)
    rowCount = 0: latestStr = "(empty)"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    rowCount = lastRow - 1   ' b? header row 1

    ' Tìm c?t date: scan header row, l?y c?t d?u tiên có "date" trong tên
    Dim dateCol As Long: dateCol = 0
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        If InStr(1, LCase$(CStr(ws.Cells(1, c).value)), "date") > 0 Then
            dateCol = c: Exit For
        End If
    Next c
    If dateCol = 0 Then Exit Sub   ' Không tìm th?y c?t date

    ' Tìm ngày m?i nh?t
    Dim latestDate As Date, hasDate As Boolean: hasDate = False
    Dim r As Long, v As Variant
    For r = 2 To lastRow
        v = ws.Cells(r, dateCol).value
        If IsDate(v) Then
            If Not hasDate Then
                latestDate = CDate(v): hasDate = True
            ElseIf CDate(v) > latestDate Then
                latestDate = CDate(v)
            End If
        End If
    Next r

    If hasDate Then latestStr = Format$(latestDate, "DD/MM/YYYY")
End Sub

'==========================================================
' ÐÓNG CPMS_Data.xlsb ÐANG M? (b?t k? file nào KHÁC targetFile)
'==========================================================
Private Sub CloseConflictingDataFile(ByVal targetFile As String)
    Dim wb As Workbook
    Dim candidates() As Workbook
    Dim n As Long: n = 0

    For Each wb In Workbooks
        If LCase$(mID$(wb.fullName, InStrRev(wb.fullName, "\") + 1)) = "cpms_data.xlsb" Then
            If LCase$(wb.fullName) <> LCase$(targetFile) Then
                n = n + 1
                ReDim Preserve candidates(1 To n)
                Set candidates(n) = wb
            End If
        End If
    Next wb

    Dim j As Long
    For j = 1 To n
        On Error Resume Next
        Application.DisplayAlerts = False
        candidates(j).Close False
        Application.DisplayAlerts = True
        On Error GoTo 0
    Next j
End Sub

'==========================================================
' THAY TH? DATA FILE C?A PROJECT HI?N CÓ
' (g?i khi user ch?n YES trong dialog so sánh)
'==========================================================
Private Sub ReplaceExistingProjectData( _
            ByVal projID As String, _
            ByVal projFolder As String, _
            ByVal existDataFile As String, _
            ByVal srcFile As String, _
            ByVal projName As String, _
            ByVal location As String)

    ' 1. Ðóng workbook cu n?u dang m?
    Dim wb As Workbook
    For Each wb In Workbooks
        If LCase$(wb.fullName) = LCase$(existDataFile) Then
            Application.DisplayAlerts = False
            wb.Close False
            Application.DisplayAlerts = True
            Exit For
        End If
    Next wb

    ' 2. T?o thu m?c backup b?ng FileSystemObject (x? lý Unicode t?t hon MkDir)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Tìm thu m?c Documents th?c t? c?a PC (k? c? khi dã chuy?n sang ? khác)
    Dim backupDir As String
    backupDir = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\CPMS_Backups"

    On Error Resume Next
    If Not fso.FolderExists(backupDir) Then fso.CreateFolder backupDir
    On Error GoTo 0

    If Not fso.FolderExists(backupDir) Then
        MsgBox "Cannot create backup folder:" & vbCrLf & backupDir, vbCritical
        Exit Sub
    End If

    ' 3. Ð?t tên file backup dùng ProjectID (tránh ký t? Unicode t? tên project)
    Dim backupFile As String
    backupFile = backupDir & "\" & projID & "_" & Format$(Now, "YYYYMMDD_HHMMSS") & ".xlsb"

    ' 4. Backup + ghi dè
    On Error GoTo EH_Replace
    FileCopy existDataFile, backupFile   ' Copy sang Documents (cross-drive OK)
    Kill existDataFile                   ' Xóa file g?c
    SafeCopyDataFile srcFile, existDataFile      ' Copy file import vào (chuẩn hóa sang xlsb an toàn)
    On Error GoTo 0

    ' 5. Ghi dúng ProjectID / PROJ_NAME vào file m?i
    WriteProjectInfo existDataFile, projID, projName, location

    ' 6. C?p nh?t LastOpened
    UpdateLastOpenedInList projID

    Set fso = Nothing
    UniMsgBox "Data for project """ & projName & """ replaced successfully." & vbCrLf & vbCrLf & _
              "Backup saved as:" & vbCrLf & backupFile, _
              vbInformation, "Import - Replace OK"
    Exit Sub

EH_Replace:
    Set fso = Nothing
    MsgBox "Replace failed at step [" & Err.Number & "]: " & Err.Description & vbCrLf & vbCrLf & _
           "existDataFile: " & existDataFile & vbCrLf & _
           "backupFile:    " & backupFile & vbCrLf & _
           "srcFile:       " & srcFile, _
           vbCritical, "Replace Error"
End Sub

'==========================================================
' C?P NH?T LASTOPENED TRONG PROJECTLIST
'==========================================================
Private Sub UpdateLastOpenedInList(ByVal projID As String)
    Dim wb As Workbook, ws As Worksheet, f As Range
    On Error GoTo EH_ULO

    Set wb = Workbooks.Open(GetProjectListFile())
    Set ws = wb.Worksheets("PROJECT_LIST")
    Set f = ws.Columns(1).Find(What:=projID, LookAt:=xlWhole, LookIn:=xlValues)

    If Not f Is Nothing Then
        f.Offset(0, 6).value = Now   ' Col G = LastOpened (index 7, offset 6 t? col A)
    End If

    wb.Save
    wb.Close False
    Exit Sub

EH_ULO:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Sub

'==========================================================
' TI?N ÍCH CANG CH? CHO B?NG SO SÁNH
'==========================================================
Private Function rj(ByVal v As Variant, ByVal w As Long) As String
    rj = Right$(Space(w) & CStr(v), w)
End Function

Private Function PadR(ByVal s As String, ByVal w As Long) As String
    PadR = Left$(s & Space(w), w)
End Function

'==========================================================
' PUBLIC WRAPPERS — d?c/ghi 1 field PROJECT_SETTINGS
' t? b?t k? DataFile nào (dùng cho Export, v.v.)
'==========================================================
Public Function GetSettingFromDataFile( _
                    ByVal DataFile As String, _
                    ByVal key As String) As String
    GetSettingFromDataFile = ReadProjectSettingField(DataFile, key)
End Function

Public Sub SetSettingInDataFile( _
               ByVal DataFile As String, _
               ByVal key As String, _
               ByVal keyValue As String)

    Dim wb          As Workbook
    Dim ws          As Worksheet
    Dim alreadyOpen As Boolean

    Set wb = GetOrOpenWorkbook(DataFile, alreadyOpen)
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    Set ws = wb.Worksheets("PROJECT_SETTINGS")
    On Error GoTo 0

    If Not ws Is Nothing Then
        SaveSettingField ws, key, keyValue
        wb.Save
    End If

    If Not alreadyOpen Then wb.Close False

End Sub



