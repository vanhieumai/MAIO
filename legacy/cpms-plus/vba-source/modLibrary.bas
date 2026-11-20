Attribute VB_Name = "modLibrary"
'Attribute VB_Name = "modLibrary"
'================================================================
' MODULE: modLibrary
' Purpose: Helper functions cho Library page cua frmContractSearch
'
' Ham cong khai:
'   EnsureLibraryDB        - Tao sheet LIBRARY_DB neu chua co
'   GetTemplatePath        - Doc duong dan template tu LIBRARY_DB
'   SaveTemplatePath       - Luu duong dan template vao LIBRARY_DB
'   GetFolderPath          - Doc duong dan thu muc tu LIBRARY_DB
'   SaveFolderPath         - Luu duong dan thu muc vao LIBRARY_DB
'   OpenTemplate           - Mo file template (hoac cho chon neu chua co)
'   OpenTemplateFolder     - Mo thu muc chua template
'   GetLibraryRecordCount  - Dem so ban ghi trong DOCUMENT_TRACKER
'   GetDocTypeName         - Lay ten day du tu ma loai ho so
'   ShowDocumentManagement - Mo form quan ly tai lieu
'   BuildDocumentPackage   - Tao goi ho so theo loai duoc chon
'================================================================
Option Explicit

Private Const LIB_SHEET  As String = "LIBRARY_DB"
Private Const DOC_SHEET  As String = "DOCUMENT_TRACKER"
Private Const SETTINGS_SHEET As String = "PROJECT_SETTINGS"


'================================================================
' ENSURE LIBRARY_DB SHEET
' Tao sheet LIBRARY_DB neu chua ton tai, them header
'================================================================
Public Sub EnsureLibraryDB(Optional dummy As Variant = 0)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetDataSheet(LIB_SHEET)
    On Error GoTo 0

End Sub

'================================================================
' GET TEMPLATE PATH
' Tra ve duong dan file template cua docType, "" neu chua co
'================================================================
Public Function GetTemplatePath(ByVal docType As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetDataSheet(LIB_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If UCase(Trim(CStr(ws.Cells(i, 1).value))) = UCase(Trim(docType)) Then
            GetTemplatePath = Trim(CStr(ws.Cells(i, 3).value))
            Exit Function
        End If
    Next i
End Function

'================================================================
' SAVE TEMPLATE PATH
' Cap nhat hoac them moi ban ghi TemplatePath cho docType
'================================================================
Public Sub SaveTemplatePath(ByVal docType As String, ByVal filePath As String)
    EnsureLibraryDB
    Dim ws As Worksheet
    Set ws = GetDataSheet(LIB_SHEET)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If UCase(Trim(CStr(ws.Cells(i, 1).value))) = UCase(Trim(docType)) Then
            ws.Cells(i, 3).value = filePath
            ws.Cells(i, 5).value = Now()
            Exit Sub
        End If
    Next i

    ' Them dong moi neu docType chua co
    Dim newRow As Long
    newRow = lastRow + 1
    ws.Cells(newRow, 1).value = UCase(Trim(docType))
    ws.Cells(newRow, 2).value = GetDocTypeName(docType)
    ws.Cells(newRow, 3).value = filePath
    ws.Cells(newRow, 5).value = Now()
End Sub

'================================================================
' GET FOLDER PATH
'================================================================
Public Function GetFolderPath(ByVal docType As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetDataSheet(LIB_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If UCase(Trim(CStr(ws.Cells(i, 1).value))) = UCase(Trim(docType)) Then
            GetFolderPath = Trim(CStr(ws.Cells(i, 4).value))
            Exit Function
        End If
    Next i
End Function

'================================================================
' SAVE FOLDER PATH
'================================================================
Public Sub SaveFolderPath(ByVal docType As String, ByVal folderPath As String)
    EnsureLibraryDB
    Dim ws As Worksheet
    Set ws = GetDataSheet(LIB_SHEET)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If UCase(Trim(CStr(ws.Cells(i, 1).value))) = UCase(Trim(docType)) Then
            ws.Cells(i, 4).value = folderPath
            ws.Cells(i, 5).value = Now()
            Exit Sub
        End If
    Next i

    Dim newRow As Long
    newRow = lastRow + 1
    ws.Cells(newRow, 1).value = UCase(Trim(docType))
    ws.Cells(newRow, 2).value = GetDocTypeName(docType)
    ws.Cells(newRow, 4).value = folderPath
    ws.Cells(newRow, 5).value = Now()
End Sub

'================================================================
' OPEN TEMPLATE
' Logic:
'   1. Lay TemplatePath tu LIBRARY_DB
'   2. Neu chua co -> mo File Picker de chon
'   3. Neu la thu muc -> mo Explorer
'   4. Neu la file -> mo bang Shell
'   5. Neu file khong ton tai -> thong bao va cho chon lai
'================================================================
Public Sub OpenTemplate(ByVal docType As String)
    Dim tplPath As String
    tplPath = GetTemplatePath(docType)

    If tplPath = "" Then
        ' Chua co template -> cho chon file moi
        Dim newPath As String
        newPath = PickLibraryFile(GetDocTypeName(docType))
        If newPath = "" Then Exit Sub
        SaveTemplatePath docType, newPath
        tplPath = newPath
    End If

    ' Kiem tra la thu muc hay file
    ' [FIX] Dung FSO thay Dir() de ho tro duong dan co ky tu tieng Viet / Unicode
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(tplPath) Then
        ' La thu muc -> mo Explorer
        Shell "explorer.exe """ & tplPath & """", vbNormalFocus
    ElseIf fso.FileExists(tplPath) Then

        ' La file ton tai -> mo bang Shell
        On Error Resume Next
        Shell "cmd /c start """" """ & tplPath & """", vbHide
        If Err.Number <> 0 Then
            MsgBox "Cannot open file: " & Err.Description, vbExclamation, "Open Error"
        End If
        On Error GoTo 0
    Else
        ' File khong con ton tai
        Dim msg As String
        msg = "Template file not found:" & vbCrLf & tplPath & vbCrLf & vbCrLf & _
              "Please select a new template."
        MsgBox msg, vbExclamation, "File Not Found"

        Dim reselect As String
        reselect = PickLibraryFile(GetDocTypeName(docType))
        If reselect <> "" Then
            SaveTemplatePath docType, reselect
        End If
    End If
End Sub

'================================================================
' OPEN TEMPLATE FOLDER
' Mo thu muc chua file template
'================================================================
Public Sub OpenTemplateFolder(ByVal docType As String)
    Dim fPath As String
    fPath = GetFolderPath(docType)

    ' Fallback: neu khong co FolderPath, lay thu muc cua TemplatePath
    If fPath = "" Then
        Dim tplPath As String
        tplPath = GetTemplatePath(docType)
        If tplPath <> "" Then
            If Len(Dir(tplPath, vbDirectory)) = 0 Then
                ' La file -> lay thu muc cha
                Dim slash As Long
                slash = InStrRev(tplPath, "\")
                If slash > 0 Then fPath = Left(tplPath, slash - 1)
            Else
                fPath = tplPath
            End If
        End If
    End If

    If fPath = "" Then
        MsgBox "No folder configured for: " & GetDocTypeName(docType), _
               vbInformation, "No Folder"
        Exit Sub
    End If

    If Len(Dir(fPath, vbDirectory)) > 0 Then
        Shell "explorer.exe """ & fPath & """", vbNormalFocus
    Else
        MsgBox "Folder not found: " & fPath, vbExclamation, "Folder Not Found"
    End If
End Sub

'================================================================
' FILE PICKER DIALOG
' Hien thi hop thoai chon file, tra ve duong dan hoac ""
'================================================================
Private Function PickLibraryFile(ByVal displayName As String) As String
    PickLibraryFile = ""
    On Error GoTo ErrPick

    Dim fd As Object
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker

    With fd
        .title = "Select Template: " & displayName
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        .Filters.Add "Word Documents", "*.doc;*.docx"
        .Filters.Add "Excel Files", "*.xls;*.xlsx;*.xlsm"
        .Filters.Add "PDF Files", "*.pdf"
        .AllowMultiSelect = False
        If .Show = True Then
            PickLibraryFile = .SelectedItems(1)
        End If
    End With
    Exit Function

ErrPick:
    MsgBox "File picker error: " & Err.Description, vbExclamation
End Function

'================================================================
' COUNT RECORDS IN DOCUMENT_TRACKER
' Dem so ban ghi co col C (DocType) khop voi docType
'================================================================
'================================================================
' DEM SO LUONG RECORD TRONG LIBRARY_DB THEO DOCTYPE
' [FIX] Them tham so contractID de loc theo hop dong cu the
'       Neu contractID = "" thi dem tat ca (hanh vi cu)
'================================================================
Public Function GetLibraryRecordCount(ByVal docType As String, _
                                       Optional ByVal contractID As String = "") As Long
    On Error Resume Next
    GetLibraryRecordCount = 0

    Dim ws As Worksheet
    Set ws = Nothing
    Set ws = GetDataSheet("LIBRARY_DB")
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim cnt As Long: cnt = 0
    Dim i As Long
    Dim cidFilter As String
    cidFilter = Trim$(contractID)

    For i = 2 To lastRow
        ' Col A = ContractID, Col B = DocType (hoac tuy schema thuc te)
        Dim rowType As String
        rowType = UCase$(Trim$(CStr(ws.Cells(i, "B").value)))  ' dieu chinh cot neu can
        If rowType <> UCase$(docType) Then GoTo NEXT_ROW

        ' [FIX] Neu co loc theo contract thi kiem tra cot ContractID
        If cidFilter <> "" Then
            Dim rowCID As String
            rowCID = Trim$(CStr(ws.Cells(i, "A").value))       ' dieu chinh cot neu can
            If rowCID <> cidFilter Then GoTo NEXT_ROW
        End If

        cnt = cnt + 1
NEXT_ROW:
    Next i

    GetLibraryRecordCount = cnt
    On Error GoTo 0
End Function

'================================================================
' GET DOCTYPE DISPLAY NAME
' Chuyen ma docType thanh ten hien thi day du
'================================================================
Public Function GetDocTypeName(ByVal docType As String) As String
    Select Case UCase(Trim(docType))
        Case "CONTRACT": GetDocTypeName = "Contract"
        Case "APPENDIX": GetDocTypeName = "Appendix"
        Case "VO":       GetDocTypeName = "Variation Orders"
        Case "IPC":      GetDocTypeName = "IPC (Interim Payment Certificate)"
        Case "ITP":      GetDocTypeName = "ITP (Inspection and Test Plan)"
        Case "MIR":      GetDocTypeName = "MIR (Material Inspection Request)"
        Case "WIR":      GetDocTypeName = "WIR (Work Inspection Request)"
        Case "MAR":      GetDocTypeName = "MAR (Material Acceptance Record)"
        Case "IAR":      GetDocTypeName = "IAR (Inspection Acceptance Record)"
        Case "RFA":      GetDocTypeName = "RFA (Request for Approval)"
        Case "RFC":      GetDocTypeName = "RFC (Request for Change)"
        Case "RFI":      GetDocTypeName = "RFI (Request for Information)"
        Case "SI":       GetDocTypeName = "SI (Site Instruction)"
        Case "NCR":      GetDocTypeName = "NCR (Non-Conformity Report)"
        Case "CAR":      GetDocTypeName = "CAR (Corrective Action Request)"
        Case "TRN":      GetDocTypeName = "TRN (Transmittal)"
        Case "MOM":      GetDocTypeName = "MOM (Minutes of Meeting)"
        Case "CMT":      GetDocTypeName = "CMT (Comment Sheet)"
        Case "LTR":      GetDocTypeName = "LTR (Letter)"
        Case "SR":       GetDocTypeName = "SR (Site Report)"
        Case "AD":       GetDocTypeName = "AD (Approved Drawings)"
        Case "SD":       GetDocTypeName = "SD (Shop Drawings)"
        Case "ABD":      GetDocTypeName = "ABD (As-Built Drawings)"
        Case "MOH":      GetDocTypeName = "MOH (Minutes of Handover)"
        Case "SPT":      GetDocTypeName = "SPT (Submission Plan and Tracking)"
        Case Else:       GetDocTypeName = docType
    End Select
End Function

'================================================================
' SHOW DOCUMENT MANAGEMENT FORM
' Mo frmDocumentManagement voi contractID va docType da chon
'================================================================
Public Sub ShowDocumentManagementold(ByVal contractID As String, _
                                   ByVal docType As String)
    On Error GoTo ErrHandler

    If contractID = "" Or contractID = "0" Then
        MsgBox "Please select a contract first.", vbInformation, "No Contract"
        Exit Sub
    End If

    Dim frm As frmDocumentEdit
    Set frm = New frmDocumentEdit
    frm.SetFilter contractID, docType
    frm.Show vbModeless
    Exit Sub

ErrHandler:
    MsgBox "Cannot open Document Management:" & vbCrLf & Err.Description, _
           vbExclamation, "Error"
End Sub

'================================================================
' BUILD DOCUMENT PACKAGE
' Tao goi ho so theo cac checkbox duoc chon trong Library page
' Xuat ra mot sheet moi tom tat tat ca tai lieu can chuyen bi
'================================================================
Public Sub BuildDocumentPackage(ByVal contractID As String, _
                                 ByVal inclContract As Boolean, _
                                 ByVal inclPayment As Boolean, _
                                 ByVal inclQuality As Boolean, _
                                 ByVal inclTechnical As Boolean, _
                                 ByVal inclDrawings As Boolean, _
                                 ByVal inclDocControl As Boolean)

    If contractID = "" Or contractID = "0" Then
        MsgBox "Please select a contract first.", vbInformation, "No Contract"
        Exit Sub
    End If

    If Not (inclContract Or inclPayment Or inclQuality Or _
            inclTechnical Or inclDrawings Or inclDocControl) Then
        MsgBox "Please select at least one document category.", _
               vbInformation, "No Category Selected"
        Exit Sub
    End If

    ' Xac dinh cac DocType can bao gom
    Dim typeList() As String
    ReDim typeList(0 To 30)
    Dim typeCount As Long
    typeCount = 0

    If inclContract Then
        typeList(typeCount) = "CONTRACT":  typeCount = typeCount + 1
        typeList(typeCount) = "APPENDIX":  typeCount = typeCount + 1
        typeList(typeCount) = "VO":        typeCount = typeCount + 1
    End If
    If inclPayment Then
        typeList(typeCount) = "IPC":       typeCount = typeCount + 1
    End If
    If inclQuality Then
        typeList(typeCount) = "SPT":       typeCount = typeCount + 1
        typeList(typeCount) = "ITP":       typeCount = typeCount + 1
        typeList(typeCount) = "MIR":       typeCount = typeCount + 1
        typeList(typeCount) = "WIR":       typeCount = typeCount + 1
        typeList(typeCount) = "MAR":       typeCount = typeCount + 1
        typeList(typeCount) = "IAR":       typeCount = typeCount + 1
    End If
    If inclTechnical Then
        typeList(typeCount) = "RFA":       typeCount = typeCount + 1
        typeList(typeCount) = "RFC":       typeCount = typeCount + 1
        typeList(typeCount) = "RFI":       typeCount = typeCount + 1
        typeList(typeCount) = "SI":        typeCount = typeCount + 1
        typeList(typeCount) = "NCR":       typeCount = typeCount + 1
        typeList(typeCount) = "CAR":       typeCount = typeCount + 1
        typeList(typeCount) = "TRN":       typeCount = typeCount + 1
    End If
    If inclDocControl Then
        typeList(typeCount) = "MOM":       typeCount = typeCount + 1
        typeList(typeCount) = "CMT":       typeCount = typeCount + 1
        typeList(typeCount) = "LTR":       typeCount = typeCount + 1
        typeList(typeCount) = "SR":        typeCount = typeCount + 1
        typeList(typeCount) = "MOH":       typeCount = typeCount + 1
    End If
    If inclDrawings Then
        typeList(typeCount) = "AD":        typeCount = typeCount + 1
        typeList(typeCount) = "SD":        typeCount = typeCount + 1
        typeList(typeCount) = "ABD":       typeCount = typeCount + 1
    End If

    If typeCount = 0 Then Exit Sub

    ' Doc du lieu tu DOCUMENT_TRACKER
    Dim wsDoc As Worksheet
    On Error Resume Next
    Set wsDoc = GetDataSheet(DOC_SHEET)
    On Error GoTo 0
    If wsDoc Is Nothing Then
        MsgBox "Sheet " & DOC_SHEET & " not found.", vbExclamation
        Exit Sub
    End If

    ' Tao Dictionary de tra cuu nhanh cac docType can xuat
    Dim dictTypes As Object
    Set dictTypes = CreateObject("Scripting.Dictionary")
    Dim t As Long
    For t = 0 To typeCount - 1
        dictTypes(typeList(t)) = 1
    Next t

    ' Quet DOCUMENT_TRACKER lay cac ban ghi khop
    Dim lastRow As Long
    lastRow = wsDoc.Cells(wsDoc.rows.count, "A").End(xlUp).Row

    Dim outArr() As Variant
    ReDim outArr(0 To lastRow, 0 To 7)
    Dim outCount As Long
    outCount = 0

    Dim i As Long
    For i = 2 To lastRow
        Dim rowCID As String
        rowCID = Trim(CStr(wsDoc.Cells(i, "B").value))
        Dim rowType As String
        rowType = UCase(Trim(CStr(wsDoc.Cells(i, "C").value)))

        If rowCID = contractID And dictTypes.Exists(rowType) Then
            outArr(outCount, 0) = outCount + 1
            outArr(outCount, 1) = wsDoc.Cells(i, "C").value
            outArr(outCount, 2) = wsDoc.Cells(i, "D").value
            outArr(outCount, 3) = wsDoc.Cells(i, "I").value
            outArr(outCount, 4) = wsDoc.Cells(i, "J").value
            outArr(outCount, 5) = wsDoc.Cells(i, "N").value
            outArr(outCount, 6) = wsDoc.Cells(i, "O").value
            outArr(outCount, 7) = wsDoc.Cells(i, "A").value
            outCount = outCount + 1
        End If
    Next i

    If outCount = 0 Then
        MsgBox "No documents found for the selected categories.", _
               vbInformation, "No Records"
        Exit Sub
    End If

    ' Tao sheet xuat ket qua
    Dim pkgSheetName As String
    pkgSheetName = "PKG_" & contractID & "_" & Format(Now(), "yyyymmdd_HHmm")
    If Len(pkgSheetName) > 31 Then pkgSheetName = Left(pkgSheetName, 31)

    Dim wsPkg As Worksheet
    If defaultWorkbook Is Nothing Then Set defaultWorkbook = GetUserWorkbook()
    Set wsPkg = defaultWorkbook.Worksheets.Add( _
        After:=defaultWorkbook.Worksheets(defaultWorkbook.Worksheets.count))
    wsPkg.Name = pkgSheetName

    ' Header
    With wsPkg.rows(1)
        .Cells(1, 1).value = "No."
        .Cells(1, 2).value = "Type"
        .Cells(1, 3).value = "Document Name"
        .Cells(1, 4).value = "Date Received"
        .Cells(1, 5).value = "Date Returned"
        .Cells(1, 6).value = "Status"
        .Cells(1, 7).value = "Remark"
        .Cells(1, 8).value = "Doc ID"
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Ghi du lieu
    For i = 0 To outCount - 1
        Dim r As Long
        r = i + 2
        wsPkg.Cells(r, 1).value = outArr(i, 0)
        wsPkg.Cells(r, 2).value = outArr(i, 1)
        wsPkg.Cells(r, 3).value = outArr(i, 2)
        wsPkg.Cells(r, 4).value = outArr(i, 3)
        wsPkg.Cells(r, 5).value = outArr(i, 4)
        wsPkg.Cells(r, 6).value = outArr(i, 5)
        wsPkg.Cells(r, 7).value = outArr(i, 6)
        wsPkg.Cells(r, 8).value = outArr(i, 7)
    Next i

    wsPkg.Columns("A:H").AutoFit

    ' Them tieu de o tren
    wsPkg.rows(1).Insert
    wsPkg.Cells(1, 1).value = "DOCUMENT PACKAGE: " & contractID & _
                               " | Generated: " & Format(Now(), "dd/mm/yyyy HH:mm")
    wsPkg.Cells(1, 1).Font.Bold = True
    wsPkg.Cells(1, 1).Font.Size = 11

    defaultWorkbook.Activate
    wsPkg.Activate
    MsgBox "Package created: " & outCount & " document(s)." & vbCrLf & _
           "Sheet: " & pkgSheetName, vbInformation, "Package Ready"

    Set dictTypes = Nothing
End Sub

'================================================================
' DEM SO LUONG RECORDS TRONG DOCUMENT_TRACKER
' theo ContractID (col B) va DocType (col C = "Ho so")
'
' [FIX] LIBRARY_DB khong co ContractID — template dung chung.
'       Records thuc te theo tung hop dong nam trong DOCUMENT_TRACKER.
'
' - contractID = "" -> dem tat ca (khong loc theo contract)
' - docType: so sanh voi col C, co ho tro alias (CONTRACT -> HD, IPC -> THANH TOAN, v.v.)
'================================================================
Public Function CountDocTrackerByType(ByVal contractID As String, _
                                        ByVal docType As String) As Long
    On Error Resume Next
    CountDocTrackerByType = 0

    Dim ws As Worksheet
    Set ws = Nothing
    Set ws = GetDataSheet("DOCUMENT_TRACKER")
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim cnt As Long: cnt = 0
    Dim cidFilter As String: cidFilter = Trim$(contractID)
    Dim dtUpper As String: dtUpper = UCase$(Trim$(docType))

    ' Alias: LibDocType -> gia tri luu trong col C cua DOCUMENT_TRACKER
    Dim alias1 As String, alias2 As String
    alias1 = ""
    alias2 = ""
    Select Case dtUpper
        Case "CONTRACT":  alias1 = "HD": alias2 = "CON"
        Case "APPENDIX":  alias1 = "APP": alias2 = "PLHD"  ' tuy du lieu thuc te
        Case "IPC":       alias1 = "IPC": alias2 = "THANH TOAN"
        Case Else:        alias1 = dtUpper
    End Select

    Dim i As Long
    For i = 2 To lastRow
        ' Loc theo ContractID (col B)
        If cidFilter <> "" Then
            If Trim$(CStr(ws.Cells(i, "B").value)) <> cidFilter Then GoTo NEXT_ROW
        End If

        ' Loc theo DocType (col C), co alias
        Dim rowType As String
        rowType = UCase$(Trim$(CStr(ws.Cells(i, "C").value)))

        Dim matched As Boolean
        matched = (rowType = dtUpper) _
               Or (alias1 <> "" And rowType = UCase$(alias1)) _
               Or (alias2 <> "" And rowType = UCase$(alias2))
        If Not matched Then GoTo NEXT_ROW

        cnt = cnt + 1
NEXT_ROW:
    Next i

    CountDocTrackerByType = cnt
    On Error GoTo 0
End Function

'================================================================
' FORCE PICK TEMPLATE
' Duoc goi boi Ctrl+Click: luon mo File Picker de chon lai,
' bo qua template hien tai (neu co)
'================================================================
Public Sub ForcePickTemplate(ByVal docType As String)
    Dim newPath As String
    newPath = PickLibraryFile(GetDocTypeName(docType))
    If newPath <> "" Then
        SaveTemplatePath docType, newPath
    End If
End Sub
'================================================================
' PROJECT_SETTINGS — Luu cac truong dung chung cho toan bo du an
' (Project Code, Client, Main Contractor, Supervision Consultant)
'
' Sheet "PROJECT_SETTINGS": 2 cot
'   Col A: Key   (PROJ | CLIENT | MAINCON | SUPCON)
'   Col B: Value
'================================================================

'----------------------------------------------------------------
' Tao sheet PROJECT_SETTINGS neu chua co
'----------------------------------------------------------------
Private Sub EnsureProjectSettings(Optional dummy As Variant = 0)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetDataSheet(SETTINGS_SHEET)
    On Error GoTo 0
    'If ws Is Nothing Then
    '    Set ws = GetDataSheet.Add( _
    '        After:=GetDataSheet(GetDataSheet.Count))
    '    ws.Name = SETTINGS_SHEET
    '    ws.Cells(1, 1).value = "Key"
    '    ws.Cells(1, 2).value = "Value"
    '    With ws.rows(1)
    '        .Font.Bold = True
    '        .Interior.Color = RGB(68, 114, 196)
    '        .Font.Color = RGB(255, 255, 255)
    '    End With
    'End If
End Sub

'----------------------------------------------------------------
' Doc gia tri global setting theo key, "" neu chua co
'----------------------------------------------------------------
Public Function GetGlobalSetting(ByVal key As String) As String
    GetGlobalSetting = ""
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = GetDataSheet(SETTINGS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(i, 1).value))) = UCase$(key) Then
            GetGlobalSetting = Trim$(CStr(ws.Cells(i, 2).value))
            Exit Function
        End If
    Next i
End Function

'----------------------------------------------------------------
' Luu global setting (cap nhat hoac them moi)
'----------------------------------------------------------------
Public Sub SaveGlobalSetting(ByVal key As String, ByVal value As String)
    EnsureProjectSettings
    On Error GoTo ErrSave
    Dim ws As Worksheet
    Set ws = GetDataSheet(SETTINGS_SHEET)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(i, 1).value))) = UCase$(key) Then
            ws.Cells(i, 2).value = value
            Exit Sub
        End If
    Next i
    ' Them moi
    Dim newRow As Long: newRow = lastRow + 1
    ws.Cells(newRow, 1).value = UCase$(key)
    ws.Cells(newRow, 2).value = value
    Exit Sub
ErrSave:
    MsgBox "SaveGlobalSetting: " & Err.Description, vbExclamation
End Sub

'================================================================
' THEM VAO MODULE: modLibrary
' (Dan 3 sub/function nay vao cuoi file modLibrary hien tai)
'
' [NEW] OpenDocumentEditFromLibrary  - Ham tong hop goi khi right-click label
' [NEW] GetGeneralCodeForContract    - Lay ma cong ty (project-level)
'================================================================


'================================================================
' [NEW] MO frmDocumentEdit TU LIBRARY (right-click label)
'
' Tham so:
'   contractID - ID hop dong hien dang chon
'   docType    - Ma loai tai lieu (RFA, RFC, IPC, ...)
'
' Hanh vi:
'   1. Kiem tra contractID hop le
'   2. Tao frmDocumentEdit moi
'   3. Goi SetFilter (tu dong chon + lock docType)
'   4. Show vbModeless
'================================================================
Public Sub OpenDocumentEditFromLibrary(ByVal contractID As String, _
                                        ByVal docType As String)
    On Error GoTo ErrHandler

    If Trim(contractID) = "" Or contractID = "0" Then
        MsgBox "Please select a contract first.", vbInformation, "No Contract"
        Exit Sub
    End If

    If Trim(docType) = "" Then
        MsgBox "Invalid document type.", vbInformation, "Error"
        Exit Sub
    End If

    Dim frm As New frmDocumentEdit
    ' SetFilter: dat contractID + docType, goi InitNew + AutoSelectDocType
    '            va TU DONG LOCK cac option doctype
    frm.SetFilter contractID, docType
    frm.Show vbModeless
    Exit Sub

ErrHandler:
    MsgBox "Cannot open Document Edit:" & vbCrLf & Err.Description, _
           vbExclamation, "Error"
End Sub


'================================================================
' [NEW] GET GENERAL CODE FOR CONTRACT
' Lay ma cong ty theo loai hop dong (dung de pre-fill cboIssuer)
' Co the mo rong neu co mapping cu the
'================================================================
Public Function GetIssuerCodeForContract(ByVal contractID As String) As String
    GetIssuerCodeForContract = ""
    ' Vi du: hop dong A-* -> issuer la VEC (tong thau)
    '        hop dong B-* -> issuer la ART
    ' Mo rong theo du lieu thuc te cua du an
    Dim grp As String
    Dim hyphen As Integer
    hyphen = InStr(contractID, "-")
    If hyphen > 1 Then
        grp = UCase(Trim(Left(contractID, hyphen - 1)))
    Else
        grp = UCase(Trim(contractID))
    End If

    Select Case grp
        Case "A": GetIssuerCodeForContract = GetGlobalSetting("CLIENT")
        Case "B": GetIssuerCodeForContract = GetGlobalSetting("MAINCON")
        Case Else: GetIssuerCodeForContract = ""
    End Select
End Function


'================================================================
' CAI DAT COMPANY_CODES TRONG PROJECT_SETTINGS
' Goi ham nay 1 lan de khoi tao danh sach ma cong ty
' (Hoac nguoi dung co the sua truc tiep trong sheet PROJECT_SETTINGS)
'
' Format: chuoi phan cach bang ";" (vi du "CLI;ART;VEC;CTC;NSH;KD")
'================================================================
Public Sub SetupCompanyCodes(ByVal codesCSV As String)
    ' Vi du goi: SetupCompanyCodes "CLI;ART;VEC;CTC;NSH;KD;TID;HHG"
    SaveGlobalSetting "COMPANY_CODES", codesCSV
    MsgBox "Company codes saved: " & codesCSV, vbInformation, "Done"
End Sub

Public Function InfoRow(ByVal lbl As String, ByVal val As String) As String
    Const LW As Integer = 13
    Dim padded As String
    If Len(lbl) >= LW Then
        padded = Left$(lbl, LW)
    Else
        padded = lbl & Space$(LW - Len(lbl))
    End If
    InfoRow = "  " & padded & Chr(9) & ": " & val & vbCrLf
End Function

'Public Function Row(ByVal lbl As String, ByVal val As String) As String
'    Row = "  " & lbl & Chr(9) & ": " & val & vbCrLf
'End Function
