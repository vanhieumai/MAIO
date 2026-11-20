Attribute VB_Name = "frmExportnewSlx"
Attribute VB_Base = "0{BBE08A53-F69C-413E-9083-1CBB9941E3E7}{365285F5-460B-49A5-8633-B9B18B757DBD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'frmExportnewSlx
Private Sub chkShowHidden_Click()

    Dim oldSelect As Boolean

    oldSelect = chkSelectAll.value

    LoadSheetList

    If oldSelect Then

        chkSelectAll.value = True

        Call chkSelectAll_Click

    End If

End Sub

Private Sub UserForm_Initialize()

    LoadSheetList
    ApplyTheme Me
    'UnregisterForm Me
End Sub



Private Sub chkSelectAll_Click()

    Dim i As Long
    
    For i = 0 To lstSheets.ListCount - 1
        lstSheets.Selected(i) = chkSelectAll.value
    Next i

End Sub

Private Sub cmdExport_Click()

    Dim i As Long
    Dim ws As Worksheet
    Dim total As Long
    Dim stepDone As Long
    Dim keepSheets As Object
    Dim wbNew As Workbook
    Dim fPath As Variant
    Dim ext As String
    Dim fileFormat As Long
    Dim tempPath As String
    Dim delSheets As New Collection
    Dim shName As Variant
    
    On Error GoTo EH
    
    Set keepSheets = CreateObject("Scripting.Dictionary")
    
    ' L?y danh sách sheet du?c ch?n
    For i = 0 To lstSheets.ListCount - 1
        If lstSheets.Selected(i) Then
            keepSheets(Replace(lstSheets.List(i), " (Hidden)", "")) = True
        End If
    Next i
    
    If keepSheets.count = 0 Then
        ShowMsgCenter "Please select sheet"
        Exit Sub
    End If
    
    ' Ch?n noi luu
    fPath = Application.GetSaveAsFilename( _
                FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xls), *.xlsx;*.xlsm;*.xls", _
                title:="Save Export File")
                
    If fPath = False Then Exit Sub
    
    ' T?t c?p nh?t
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Clone workbook
    Dim srcExt As String
    srcExt = LCase$(mID$(defaultWorkbook.Name, InStrRev(defaultWorkbook.Name, ".") + 1))
    If srcExt = "" Then srcExt = "xlsm"
    
    tempPath = Environ$("TEMP") & "\clone_" & Format(Now, "yyyymmdd_hhnnss") & "." & srcExt
    defaultWorkbook.SaveCopyAs tempPath
    
    Set wbNew = Workbooks.Open(tempPath)
    
    total = wbNew.Worksheets.count
    stepDone = 0
    UpdateProgress 0, total
    
    ' Gom sheet c?n xóa
    For Each ws In wbNew.Worksheets
        
        If Not keepSheets.Exists(ws.Name) Then
            delSheets.Add ws.Name
        End If
        
        stepDone = stepDone + 1
        If stepDone Mod 3 = 0 Or stepDone = total Then
            UpdateProgress stepDone, total
        End If
        
    Next ws
    
    ' Không du?c xóa h?t workbook
    If delSheets.count >= wbNew.Worksheets.count Then
        Err.Raise vbObjectError + 1000, , "Cannot delete all worksheets."
    End If
    
    ' Xóa sheet: n?u VeryHidden/Hidden thì b?t Visible tru?c
    For i = delSheets.count To 1 Step -1
        
        Set ws = wbNew.Worksheets(CStr(delSheets(i)))
        
        ' n?u là very hidden / hidden thì chuy?n sang visible tru?c khi delete
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If
        
        ws.Delete
    Next i
    
    ' Xác d?nh format
    ext = LCase$(mID$(CStr(fPath), InStrRev(CStr(fPath), ".") + 1))
    
    Select Case ext
        Case "xlsx"
            fileFormat = xlOpenXMLWorkbook   ' b? macro
        Case "xlsm"
            fileFormat = xlOpenXMLWorkbookMacroEnabled
        Case "xls"
            fileFormat = xlExcel8
        Case Else
            fileFormat = xlOpenXMLWorkbook
    End Select
    
    wbNew.SaveAs CStr(fPath), fileFormat:=fileFormat
    wbNew.Close SaveChanges:=False
    
    On Error Resume Next
    Kill tempPath
    On Error GoTo 0
    
    ' Restore app
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    UpdateProgress total, total
    ShowMsgCenter "Export Successfully"
    Exit Sub

EH:
    On Error Resume Next
    
    If Not wbNew Is Nothing Then
        wbNew.Close SaveChanges:=False
    End If
    
    If Len(tempPath) > 0 Then
        Kill tempPath
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ShowMsgCenter "Export failed: " & Err.Description
End Sub


Private Sub UpdateProgress(curStep As Long, totalStep As Long)

    Dim pct As Double
    
    If totalStep = 0 Then Exit Sub
    
    pct = curStep / totalStep

    lblProgressBar.Width = fraProgress.Width * pct
    lblPercent.Caption = Format(pct, "0%")

    DoEvents
    Me.Repaint

End Sub

Private Sub LoadSheetList()

    Dim ws As Worksheet

    lstSheets.Clear

    If defaultWorkbook Is Nothing Then
        ShowMsgCenter "No user workbook detected."
        Exit Sub
    End If

    For Each ws In defaultWorkbook.Worksheets    ' <-- Ð?I

        If chkShowHidden.value Then

            lstSheets.AddItem ws.Name

            If ws.Visible <> xlSheetVisible Then
                lstSheets.List(lstSheets.ListCount - 1) = ws.Name & " (Hidden)"
            End If

        Else

            If ws.Visible = xlSheetVisible Then
                lstSheets.AddItem ws.Name
            End If

        End If

    Next ws

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        'Cancel = True
        'NotifyAndClose
        Unload Me
    End If
    'ApplyTheme Me
    UnregisterForm Me
End Sub
