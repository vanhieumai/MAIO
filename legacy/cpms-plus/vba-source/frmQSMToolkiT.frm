Attribute VB_Name = "frmQSMToolkiT"
Attribute VB_Base = "0{AF7B236F-8B6B-4EB2-BA3A-09699393DA31}{6361A17E-78D2-4DDF-BA67-BF2A0E808071}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'form LinkKL
'Ðã tách trang License ra form frmLicense riêng.
'Form này ch? còn n?i dung Toolkit (b? MultiPage1).
Option Explicit
Private isBusy As Boolean
Private srcWorkbook As Workbook


Private Sub btnExpandRows_Click()
    If defaultWorkbook Is Nothing Then
        ShowMsgCenter "No workbook loaded. Please open a workbook first.", vbExclamation
        Exit Sub
    End If

    Dim f As frmExpandRows
    Set f = New frmExpandRows
    f.Init defaultWorkbook, cbSheetGT.value
    f.Show vbModeless
    Set f = Nothing
End Sub

Private Sub btnExportSheet_Click()
    'frmExportSlx.Show vbModeless
    frmExportnewSlx.Show vbModeless
End Sub



Private Sub imgVanhieumai_MouseUp(ByVal Button As Integer, _
                             ByVal Shift As Integer, _
                             ByVal X As Single, _
                             ByVal Y As Single)
    HandleImageMouseUp imgVanhieumai, Button, Shift
End Sub

Private Sub imgVNCN_MouseUp(ByVal Button As Integer, _
                             ByVal Shift As Integer, _
                             ByVal X As Single, _
                             ByVal Y As Single)
    HandleImageMouseUp imgVNCN, Button, Shift
End Sub





Private Sub UpdateSourceLabel()

    Dim userWbName As String
    If defaultWorkbook Is Nothing Then
        userWbName = "(no workbook open)"
    Else
        userWbName = defaultWorkbook.Name
    End If

    If srcWorkbook Is Nothing Then
        ' Chua ch?n source ngoài ? TH l?y t? file user
        lblSourceInfo.Caption = "Source: " & userWbName
    Else
        ' Ðã ch?n external source cho sheet TH
        lblSourceInfo.Caption = "Source (TH): External [" & srcWorkbook.Name & "]" & _
                                "  |  Target (GT): " & userWbName
    End If

End Sub

Private Sub RefreshSheetCombos()

    Dim ws As Worksheet

    cbSheetTH.Clear
    cbSheetGT.Clear

    ' N?u chua detect du?c workbook c?a user thì th? l?i
    If defaultWorkbook Is Nothing Then
        Set defaultWorkbook = GetUserWorkbook()
    End If

    If defaultWorkbook Is Nothing Then Exit Sub   ' Không có workbook nào m?

    For Each ws In defaultWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            cbSheetTH.AddItem ws.Name
            cbSheetGT.AddItem ws.Name
        End If
    Next ws

End Sub


Private Sub cbSheetTH_Change()

    If isBusy Then Exit Sub
    If cbSheetTH.ListIndex = -1 Then Exit Sub
    If Trim(cbSheetTH.value) = "" Then Exit Sub

    UpdateSourceLabel

    Dim wb As Workbook
    If srcWorkbook Is Nothing Then
        Set wb = defaultWorkbook
    Else
        Set wb = srcWorkbook
    End If

    '--- DOUBLE CHECK: sheet có t?n t?i không
    On Error Resume Next
    LoadColumns wb.Worksheets(cbSheetTH.value), cbCotTenTH, lbCotKL_TH, 1
    On Error GoTo 0

End Sub



Private Sub cbSheetGT_Change()

    If isBusy Then Exit Sub
    If cbSheetGT.ListIndex = -1 Then Exit Sub
    If Trim(cbSheetGT.value) = "" Then Exit Sub

    LoadColumns defaultWorkbook.Worksheets(cbSheetGT.value), cbCotTenGT, lbCotKL_GT, 1

End Sub



Private Sub btnExternalSource_Click()

    Dim fd As FileDialog
    Dim wbSrc As Workbook
    Dim wbTarget As Workbook
    Dim selectedPath As String
    Dim wb As Workbook

    Set wbTarget = defaultWorkbook

    Set srcWorkbook = Nothing
    srcWbName = ""
    srcWbPath = ""

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .title = "Select source Excel file"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
    End With

    selectedPath = fd.SelectedItems(1)

    ' Ki?m tra file dã m? chua qua Workbooks collection (KHÔNG dùng GetObject)
    For Each wb In Application.Workbooks
        If wb.fullName = selectedPath Then
            Set wbSrc = wb
            Exit For
        End If
    Next wb

    ' N?u dã m? nhung không ph?i read-only ? dóng và m? l?i
    If Not wbSrc Is Nothing Then
        If Not wbSrc.ReadOnly Then
            wbSrc.Close SaveChanges:=False
            Set wbSrc = Nothing
        End If
    End If

    ' M? file n?u chua m?
    If wbSrc Is Nothing Then

        Application.ScreenUpdating = False
        Application.AskToUpdateLinks = False

        Set wbSrc = Workbooks.Open( _
            fileName:=selectedPath, _
            UpdateLinks:=0, _
            ReadOnly:=True, _
            IgnoreReadOnlyRecommended:=True, _
            AddToMru:=False _
        )

        Application.AskToUpdateLinks = True
        Application.ScreenUpdating = True

    End If

    If wbSrc Is Nothing Then
        ShowMsgCenter "Cannot open source file"
        UpdateSourceLabel
        Exit Sub
    End If

    If wbSrc.Windows.count > 0 Then
        wbSrc.Windows(1).Visible = True
    End If

    Set srcWorkbook = wbSrc
    srcWbName = wbSrc.Name
    srcWbPath = wbSrc.fullName
    isUsingExternalSource = True

    LoadSourceSheets
    UpdateSourceLabel

    wbTarget.Activate

    ShowMsgCenter "External file loaded as Source (Read-Only)"

End Sub


Private Sub LoadSourceSheets()

    Dim ws As Worksheet

    cbSheetTH.Clear

    For Each ws In srcWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            cbSheetTH.AddItem ws.Name
        End If
    Next ws

End Sub

Private Sub btnWorkbookManager_Click()

    On Error GoTo EH

    '==============================
    ' CH? M? FORM — m?i hành d?ng
    ' (unhide/rehide sheets, merge,
    '  luu file) du?c x? lý bên trong
    '  form frmMergeSheets.
    '==============================
    Dim f As frmMergeSheets
    Set f = New frmMergeSheets
    f.Show vbModeless
    Set f = Nothing

    '==============================
    ' REFRESH SHEET LIST (n?u c?n)
    '==============================
    RefreshSheetCombos

    Exit Sub

EH:
    ShowMsgCenter _
        "An error occurred while opening Merge tool." & vbCrLf & _
        "Please try again.", vbCritical

End Sub



Private Sub btnPreviewLink_Click()

    Dim wsTH As Worksheet, wsGT As Worksheet
    Dim colTenTH As Long, colTenGT As Long
    Dim colKL_TH() As Long
    Dim rGT As Long, rTH As Long
    Dim i As Long, idx As Long, nKL As Long
    Dim firstRow As Long: firstRow = 1
    Dim lastRowGT As Long

    '==============================
    ' CHECK SOURCE
    '==============================
    If Not IsExternalSourceAlive() Then
        ShowMsgCenter _
            "Source file has been closed" & vbCrLf & _
            "Please click RESET and select source again"
        Exit Sub
    End If

    '==============================
    ' CHECK INPUT
    '==============================
    If cbSheetTH.value = "" Or cbSheetGT.value = "" _
        Or cbCotTenTH.value = "" Or cbCotTenGT.value = "" Then
        ShowMsgCenter "Please select all required sheets and columns before preview"
        Exit Sub
    End If

    '==============================
    ' GET KL COLUMNS (MAX 4)
    '==============================
    nKL = -1
    For i = 0 To lbCotKL_TH.ListCount - 1
        If lbCotKL_TH.Selected(i) Then
            nKL = nKL + 1
            ReDim Preserve colKL_TH(nKL)
            colKL_TH(nKL) = wsColumnIndex(lbCotKL_TH.List(i))
            If nKL = 3 Then Exit For
        End If
    Next i

    If nKL < 0 Then
        ShowMsgCenter "Please select at least one Quantity column"
        Exit Sub
    End If

    '==============================
    ' SET WORKSHEETS
    '==============================
    Dim wbSrc As Workbook
    Set wbSrc = GetSourceWorkbook()
    If wbSrc Is Nothing Then
        ShowMsgCenter "Source file not available"
        Exit Sub
    End If

    Set wsTH = wbSrc.Worksheets(cbSheetTH.value)
    Set wsGT = defaultWorkbook.Worksheets(cbSheetGT.value)

    colTenTH = wsTH.Range(cbCotTenTH.value).Column
    colTenGT = wsGT.Range(cbCotTenGT.value).Column

    lastRowGT = wsGT.Cells(wsGT.rows.count, colTenGT).End(xlUp).Row

    '==============================
    ' CREATE PREVIEW FORM
    '==============================
    Dim f As frmPreviewLink
    Set f = New frmPreviewLink
    f.SetSourceWorkbook wbSrc

    With f.lbPreview
        .Clear
        .ColumnCount = 6
    End With

    '==============================
    ' BUILD PREVIEW
    '==============================
    Dim arrSheet() As String
    Dim arrAddress() As String
    Dim p As Long: p = -1

    For rGT = firstRow To lastRowGT

        Dim itemName As String
        itemName = Trim(wsGT.Cells(rGT, colTenGT).value)
        If itemName = "" Then GoTo NextGT

        rTH = FindMatchingRow(wsTH, wsGT, rGT, colTenTH, colTenGT)

        p = p + 1
        ReDim Preserve arrSheet(p)
        ReDim Preserve arrAddress(p)

        With f.lbPreview
            .AddItem itemName

            If rTH > 0 Then
                arrSheet(p) = wsTH.Name
                arrAddress(p) = wsTH.Cells(rTH, colKL_TH(0)).Address
            Else
                arrSheet(p) = ""
                arrAddress(p) = ""
            End If

            For idx = 0 To 3
                If idx <= nKL And rTH > 0 Then
                    .List(.ListCount - 1, idx + 1) = _
                        FormatPreviewValue(wsTH.Cells(rTH, colKL_TH(idx)).value)
                Else
                    .List(.ListCount - 1, idx + 1) = ""
                End If
            Next idx

            .List(.ListCount - 1, 5) = IIf(rTH > 0, "OK", "Not Found")
        End With

NextGT:
    Next rGT

    If f.lbPreview.ListCount = 0 Then
        ShowMsgCenter "No data available for preview"
        Unload f
        Exit Sub
    End If

    '==============================
    ' PASS DATA & SHOW
    '==============================
    f.SetPreviewSource arrSheet, arrAddress
    f.Show vbModeless
    Set f = Nothing

End Sub




Private Function FindMatchingRow(wsTH As Worksheet, wsGT As Worksheet, _
                                 rGT As Long, colTenTH As Long, colTenGT As Long) As Long
    Dim i As Long, rTH As Long
    Dim hm As String
    Dim countGT As Long, countTH As Long
    Dim firstRow As Long: firstRow = 10

    hm = Trim(wsGT.Cells(rGT, colTenGT).value)

    For i = firstRow To rGT
        If Trim(wsGT.Cells(i, colTenGT).value) = hm Then countGT = countGT + 1
    Next i

    For rTH = firstRow To wsTH.Cells(wsTH.rows.count, colTenTH).End(xlUp).Row
        If Trim(wsTH.Cells(rTH, colTenTH).value) = hm Then
            countTH = countTH + 1
            If countTH = countGT Then
                FindMatchingRow = rTH
                Exit Function
            End If
        End If
    Next rTH
End Function

Private Function FormatPreviewValue(v As Variant) As String

    ' 1. Error value (#N/A, #VALUE!, ...)
    If IsError(v) Then
        FormatPreviewValue = ""
        Exit Function
    End If

    ' 2. Empty / Null
    If IsEmpty(v) Or IsNull(v) Then
        FormatPreviewValue = ""
        Exit Function
    End If

    ' 3. Empty string
    If Trim(CStr(v)) = "" Then
        FormatPreviewValue = ""
        Exit Function
    End If

    ' 4. Numeric
    If IsNumeric(v) Then
        FormatPreviewValue = Format(CDbl(v), "#,##0.00")
        Exit Function
    End If

    ' 5. Default
    FormatPreviewValue = CStr(v)

End Function



Private Function ValidateInputForLink() As Boolean

    Dim i As Long
    Dim cntTH As Long, cntGT As Long

    '--- CHECK SHEET
    If cbSheetTH.value = "" Or cbSheetGT.value = "" Then
        ShowMsgCenter "Please select source sheet and target sheet", vbExclamation
        Exit Function
    End If

    '--- CHECK C?T TÊN
    If cbCotTenTH.value = "" Or cbCotTenGT.value = "" Then
        ShowMsgCenter "Please select matching column", vbCritical
        Exit Function
    End If

    '--- COUNT C?T KL
    For i = 0 To lbCotKL_TH.ListCount - 1
        If lbCotKL_TH.Selected(i) Then cntTH = cntTH + 1
    Next i

    For i = 0 To lbCotKL_GT.ListCount - 1
        If lbCotKL_GT.Selected(i) Then cntGT = cntGT + 1
    Next i

    If cntTH = 0 Or cntGT = 0 Then
        ShowMsgCenter "Please select the data column", vbCritical
        Exit Function
    End If

    If cntTH <> cntGT Then
        ShowMsgCenter "The number of data columns in the source and target must be the same", vbExclamation
        Exit Function
    End If

    ValidateInputForLink = True

End Function



'==============================
' FORM ACTIVATE
'==============================
Private Sub UserForm_Activate()

    Static UIloaded As Boolean

    If UIloaded Then Exit Sub

    isBusy = True

    '===================================
    ' LOAD INTERNAL SHEETS
    '===================================
    RefreshSheetCombos

    cbSheetTH.ListIndex = -1
    cbSheetGT.ListIndex = -1
    UpdateSourceLabel

    '===================================
    ' LOAD IMAGES
    ' NÂNG C?P: Image1..Image6 gi? là ?nh theo THEME (category
    ' "Icons", resName = tên control), du?c ApplyTheme Me
    ' (UserForm_Initialize) n?p t? d?ng qua modTheme.LoadControlImages
    '===================================

    isBusy = False

    UIloaded = True

End Sub


'==============================
' FORM INITIALIZE
'==============================
Private Sub UserForm_Initialize()

    isBusy = True
    ApplyTheme Me
    'UnregisterForm Me

    '===================================
    ' DETECT WORKBOOK C?A USER (file dang m? khi user g?i form)
    '===================================
    Set defaultWorkbook = GetUserWorkbook()

    '===================================
    ' WORKING DIRECTORY — dùng path c?a user workbook, không ph?i XLAM
    '===================================
    On Error Resume Next
    If Not defaultWorkbook Is Nothing Then
        ChDrive defaultWorkbook.path
        ChDir defaultWorkbook.path
    End If
    On Error GoTo 0

    '===================================
    ' Bây gi? m?i b?t error handler chính
    '===================================
    On Error GoTo SAFE_EXIT

    '===================================
    ' RESET SOURCE CONTEXT
    '===================================
    Set srcWorkbook = Nothing
    srcWbName = ""
    srcWbPath = ""
    isUsingExternalSource = False

    UpdateSourceLabel

    '===================================
    ' CLEAR COMBO / LIST (KHÔNG LOAD)
    '===================================
    cbSheetTH.Clear
    cbSheetGT.Clear
    cbCotTenTH.Clear
    cbCotTenGT.Clear
    lbCotKL_TH.Clear
    lbCotKL_GT.Clear

SAFE_EXIT:
    isBusy = False

End Sub


Private Function IsSourceWorkbookOpen() As Boolean

    Dim wb As Workbook

    If srcWbName = "" Then Exit Function

    For Each wb In Application.Workbooks

        If StrComp(wb.Name, srcWbName, vbTextCompare) = 0 Then
            Set srcWorkbook = wb
            IsSourceWorkbookOpen = True
            Exit Function
        End If

    Next wb

    ' source dã dóng
    Set srcWorkbook = Nothing

End Function

Private Function CheckSourceOrReset() As Boolean

    If srcWbName = "" Then
        CheckSourceOrReset = True
        Exit Function
    End If

    If Not IsSourceWorkbookOpen Then

        srcWbName = ""
        srcWbPath = ""

        ShowMsgCenter "Source workbook has been closed.", vbExclamation

        btnResetLink_Click
        Exit Function

    End If

    CheckSourceOrReset = True

End Function

Private Sub cbSheetTH_AfterUpdate()

    Dim ws As Worksheet
    Dim wbSrc As Workbook

    If Not CheckSourceOrReset Then Exit Sub

    ' N?u có external source ? dùng nó, không thì dùng file user
    If srcWorkbook Is Nothing Then
        Set wbSrc = defaultWorkbook
    Else
        Set wbSrc = srcWorkbook
    End If

    If wbSrc Is Nothing Then Exit Sub

    On Error Resume Next
    Set ws = wbSrc.Worksheets(cbSheetTH.value)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    LoadColumns ws, cbCotTenTH, lbCotKL_TH, 1

End Sub

Private Sub cbSheetGT_AfterUpdate()

    Dim ws As Worksheet

    If defaultWorkbook Is Nothing Then Exit Sub

    On Error Resume Next
    Set ws = defaultWorkbook.Worksheets(cbSheetGT.value)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    LoadColumns ws, cbCotTenGT, lbCotKL_GT, 1

End Sub

Private Sub cbSheetTH_DropButtonClick()

    If Not CheckSourceOrReset Then Exit Sub

End Sub

Private Sub cbSheetGT_DropButtonClick()

    If Not CheckSourceOrReset Then Exit Sub

End Sub




'==============================
' LOAD COLUMNS (SAFE – NO TYPE ERROR)
'==============================
Private Sub LoadColumns(ws As Worksheet, cbo As Object, lb As Object, _
                        Optional extraCols As Long = 0)

    Dim i As Long
    Dim lastCol As Long

    cbo.Clear
    lb.Clear
    lb.MultiSelect = fmMultiSelectMulti

    '=== TÌM C?T CU?I CÓ DATA (B?T K? ROW NÀO) ===
    ' Dùng Find thay vì UsedRange (tránh th?a do ô có format không có value)
    ' và thay vì End(xlToLeft) trên row 1 (b? sót khi row 1 tr?ng)
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", _
                                 After:=ws.Cells(1, 1), _
                                 LookIn:=xlFormulas, _
                                 SearchOrder:=xlByColumns, _
                                 SearchDirection:=xlPrevious)
    On Error GoTo 0

    If lastCell Is Nothing Then
        If extraCols = 0 Then Exit Sub
        lastCol = 0
    Else
        lastCol = lastCell.Column
    End If

    For i = 1 To lastCol + extraCols
        cbo.AddItem ws.Cells(1, i).Address(False, False)
        lb.AddItem ws.Cells(1, i).Address(False, False)
    Next i

End Sub


Private Sub btnLink_Click()
    Dim oldPath As String
    oldPath = CurDir
    ChDrive Environ$("TEMP")
    ChDir Environ$("TEMP")

    If Not ValidateInputForLink Then Exit Sub
    If Not IsExternalSourceAlive() Then
        ShowMsgCenter "Source file has been closed." & vbCrLf & _
                      "Please click RESET and select source again."
        Exit Sub
    End If

    Dim wsTH As Worksheet, wsGT As Worksheet
    Dim firstRow As Long: firstRow = 1

    Dim colTenTH As Long, colTenGT As Long
    Dim lastRowTH As Long, lastRowGT As Long

    Dim arrTH() As Long, arrGT() As Long
    Dim i As Long, j As Long, n As Long

    Dim dict As Object, usedIdx As Object

    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

'====================================================
' SPEED MODE
'====================================================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo SafeExit

'====================================================
' SHEETS
'====================================================
    If srcWorkbook Is Nothing Then
        Set wsTH = defaultWorkbook.Worksheets(cbSheetTH.value)
    Else
        Set wsTH = srcWorkbook.Worksheets(cbSheetTH.value)
    End If
    Set wsGT = defaultWorkbook.Worksheets(cbSheetGT.value)

'====================================================
' KEY COLUMNS
'====================================================
    colTenTH = wsTH.Range(cbCotTenTH.value).Column
    colTenGT = wsGT.Range(cbCotTenGT.value).Column

    lastRowTH = wsTH.Cells(wsTH.rows.count, colTenTH).End(xlUp).Row
    lastRowGT = wsGT.Cells(wsGT.rows.count, colTenGT).End(xlUp).Row

'====================================================
' LOAD SELECTED LINK COLUMNS
'====================================================
    Dim ctH As Long, cGT As Long

    ctH = -1
    For i = 0 To lbCotKL_TH.ListCount - 1
        If lbCotKL_TH.Selected(i) Then
            ctH = ctH + 1
            ReDim Preserve arrTH(ctH)
            arrTH(ctH) = wsTH.Range(lbCotKL_TH.List(i)).Column
        End If
    Next i

    cGT = -1
    For i = 0 To lbCotKL_GT.ListCount - 1
        If lbCotKL_GT.Selected(i) Then
            cGT = cGT + 1
            ReDim Preserve arrGT(cGT)
            arrGT(cGT) = wsGT.Range(lbCotKL_GT.List(i)).Column
        End If
    Next i

    If ctH <> cGT Then
        ShowMsgCenter "Source and Target column count mismatch", vbCritical
        GoTo SafeExit
    End If

    n = ctH

'====================================================
' LOAD KEY ARRAYS (FAST)
'====================================================
    Dim dataTH, dataGT
    dataTH = wsTH.Columns(colTenTH).value
    dataGT = wsGT.Columns(colTenGT).value

'====================================================
' BUILD DICTIONARY (key -> rows)
'====================================================
    Set dict = CreateObject("Scripting.Dictionary")

    Dim key As String

    For i = firstRow To lastRowTH

        key = Trim(dataTH(i, 1))

        If key <> "" Then
            If Not dict.Exists(key) Then
                dict.Add key, New Collection
            End If
            dict(key).Add i
        End If

    Next i

'====================================================
' LINK BY COLUMN (FAST + SAFE)
'====================================================
    Dim outArr() As Variant
    Dim colIdx As Long
    Dim rowList As Collection
    Dim rSource As Long
    Dim f As String
    Dim srcVal

    If wsGT.AutoFilterMode Then
        If wsGT.filterMode Then wsGT.ShowAllData
    End If

    For j = 0 To n

        colIdx = arrGT(j)

        ReDim outArr(1 To lastRowGT - firstRow + 1, 1 To 1)

        Set usedIdx = CreateObject("Scripting.Dictionary")

        For i = firstRow To lastRowGT

            key = Trim(dataGT(i, 1))

            If key <> "" Then

                If dict.Exists(key) Then

                    Set rowList = dict(key)

                    If Not usedIdx.Exists(key) Then
                        usedIdx(key) = 1
                    Else
                        usedIdx(key) = usedIdx(key) + 1
                    End If

                    If usedIdx(key) <= rowList.count Then

                        rSource = rowList(usedIdx(key))

                        Dim srcCell As Range

                        Set srcCell = wsTH.Cells(rSource, arrTH(j))
                        srcVal = srcCell.value

                        If Not IsError(srcVal) Then

                            If Len(Trim(CStr(srcVal))) > 0 Then

                                If wsTH.parent Is defaultWorkbook Then

                                    f = "='" & wsTH.Name & "'!" & _
                                        srcCell.Address(False, False)

                                Else

                                    f = "='[" & wsTH.parent.Name & "]" & _
                                        wsTH.Name & "'!" & _
                                        srcCell.Address(False, False)

                                End If

                                outArr(i - firstRow + 1, 1) = f

                            End If

                        End If

                    End If

                End If

            End If

        Next i

        wsGT.Range(wsGT.Cells(firstRow, colIdx), _
                   wsGT.Cells(lastRowGT, colIdx)).Formula = outArr

    Next j


    'ResetSheetSelection

    ChDrive oldPath
    ChDir oldPath
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ShowMsgCenter "Link completed (Turbo Formula mode)", vbInformation

'====================================================
' RESTORE
'====================================================
SafeExit:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload Me
    End If
    UnregisterForm Me
End Sub

Private Sub ResetSheetSelection()

    isBusy = True

    ' reset sheet selection
    cbSheetTH.ListIndex = -1
    cbSheetGT.ListIndex = -1

    ' clear column controls
    cbCotTenTH.Clear
    cbCotTenGT.Clear

    lbCotKL_TH.Clear
    lbCotKL_GT.Clear

    isBusy = False

End Sub


Private Sub btnResetLink_Click()

    Dim i As Long
    isBusy = True

    '==============================
    ' CLEAR SOURCE WORKBOOK (FILE NGOÀI)
    '==============================
    Set srcWorkbook = Nothing
    srcWbName = ""
    srcWbPath = ""
    isUsingExternalSource = False

    '==============================
    ' RELOAD SHEET LIST = THISWORKBOOK
    '==============================
    Call RefreshSheetCombos

    '==============================
    ' RESET SELECTION (KHÔNG CLEAR LIST)
    '==============================
    cbSheetTH.ListIndex = -1
    cbSheetGT.ListIndex = -1

    '==============================
    ' CLEAR COLUMN COMBO
    '==============================
    cbCotTenTH.Clear
    cbCotTenGT.Clear

    '==============================
    ' CLEAR KL LISTBOX
    '==============================
    lbCotKL_TH.Clear
    lbCotKL_GT.Clear

    '==============================
    ' UPDATE SOURCE LABEL
    '==============================
    UpdateSourceLabel

    isBusy = False

End Sub




Private Function IsExternalSourceAlive() As Boolean
    Dim wb As Workbook

    If srcWorkbook Is Nothing Then
        IsExternalSourceAlive = True   ' internal source
        Exit Function
    End If

    For Each wb In Application.Workbooks
        If wb Is srcWorkbook Then
            IsExternalSourceAlive = True
            Exit Function
        End If
    Next wb

    IsExternalSourceAlive = False
End Function




Private Function HasAnyContract(ws As Worksheet) As Boolean
    Dim r As Long
    For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).Row
        If Trim(ws.Cells(r, 1).value) <> "" _
           And Left(ws.Cells(r, 1).value, 5) <> "GROUP" Then
            HasAnyContract = True
            Exit Function
        End If
    Next r
    HasAnyContract = False
End Function




Public Sub SyncLedgerFromContract(contractID As String)

    Dim wsLedger As Worksheet
    Dim wsDB As Worksheet
    Dim r As Long, lastRow As Long
    Dim dbRow As Long

    If Trim(contractID) = "" Then Exit Sub

    Set wsLedger = EnsurePaymentLedger
    Set wsDB = EnsureContractDB

    ' === tìm contract trong DB ===
    dbRow = FindContractRowByID(wsDB, contractID)
    If dbRow = 0 Then Exit Sub

    lastRow = wsLedger.Cells(wsLedger.rows.count, "C").End(xlUp).Row

    ' === quét toàn b? ledger ===
    For r = 2 To lastRow

        ' ===== dúng contract family =====
        If (Trim(wsLedger.Cells(r, "C").value) = contractID) And (Trim(wsLedger.Cells(r, "B").value) = "") Then

            With wsLedger

                ' ===== CORE INFO =====
                Call WriteTextCell(.Cells(r, "D"), wsDB.Cells(dbRow, "B").value)
                .Cells(r, "F").value = wsDB.Cells(dbRow, "K").value   ' Scope / ItemName

                ' ===== FINANCIAL MASTER =====
                .Cells(r, "G").value = wsDB.Cells(dbRow, "M").value   ' ExchangeRate
                .Cells(r, "H").value = wsDB.Cells(dbRow, "N").value   ' EstimatedValue
                .Cells(r, "I").value = wsDB.Cells(dbRow, "O").value   ' ContractValue

            End With

        End If

    Next r

End Sub




