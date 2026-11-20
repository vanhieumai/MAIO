Attribute VB_Name = "frmAskAI"
Attribute VB_Base = "0{54D8F600-8BAC-47AE-AA41-BC86D4DA4EF2}{2B437E6A-432B-4D74-813D-4FEC88F21E80}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Const APIKEY_PREFIX As String = "apikey_"
Private m_ImagePath   As String
Private m_ImageBase64 As String
Private m_ImageMime   As String


' ============================================================
' CONVERT IMAGE TO BASE64
' ============================================================
Private Function ImageToBase64(filePath As String) As String
    On Error GoTo EH

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile filePath

    Dim bytes() As Byte
    bytes = stream.Read
    stream.Close

    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    Dim node As Object
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    ImageToBase64 = Replace(node.Text, vbLf, "")
    Exit Function
EH:
    ImageToBase64 = ""
End Function


' ============================================================
' FORM INIT — an toàn khi m? l?i
' ============================================================
Private Sub UserForm_Initialize()
    On Error GoTo EH
    ApplyTheme Me
    'UnregisterForm Me
    ' Reset module-level variables tru?c (phòng tru?ng h?p form b? unload gi?a ch?ng)
    m_ImagePath = ""
    m_ImageBase64 = ""
    m_ImageMime = ""

    ' G?i OpenAISession v?i fallback: n?u l?i thì close session cu r?i th? l?i
    On Error Resume Next
    OpenAISession
    Dim sessionErr As Long: sessionErr = Err.Number
    On Error GoTo EH

    If sessionErr <> 0 Then
        On Error Resume Next
        CloseAISession   ' D?n s?ch state cu
        Err.Clear
        OpenAISession    ' Th? l?i l?n 2
        On Error GoTo EH
    End If

    txtQuestion.Text = ""
    lblRespone.value = ""
    lblRespone.WordWrap = True
    lblRespone.AutoSize = False
    Me.Caption = "Ask AI"
    CenterUploadLabel
    UpdateProviderUI
    
    
    'UnregisterForm Me
    Exit Sub
EH:
    MsgBox "AskAI: " & Err.Description, vbCritical
    Unload Me
End Sub

Private Sub imgPreview_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                  ByVal X As Single, ByVal Y As Single)
    If Button <> 2 Then Exit Sub
    If m_ImageBase64 = "" Then Exit Sub
    If MsgBox("Remove the selected image?", vbQuestion + vbYesNo, "Ask AI") = vbYes Then
        ClearImage
    End If
End Sub

' ============================================================
' C?P NH?T TOOLTIP
' ============================================================
Private Sub UserForm_Activate()
    txtQuestion.SetFocus
    imgPreview.ControlTipText = _
        "Click to change image." & vbCrLf & _
        "Right-click to remove it."
    lblClickUpload.ControlTipText = "Click to select an image."
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    CloseAISession
    UnregisterForm Me
End Sub

' ============================================================
' SETTINGS SHEET
' ============================================================
Private Function GetSettingsSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetAISheet("SETTINGS")
    On Error GoTo 0
    Set GetSettingsSheet = ws
End Function

' ============================================================
' GET / SAVE APP SETTING
' ============================================================
Public Function GetAppSetting(settingKey As String) As String
    Dim ws As Worksheet
    Set ws = GetSettingsSheet()
    If ws Is Nothing Then Exit Function
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, 1).value)) = settingKey Then
            GetAppSetting = Trim(CStr(ws.Cells(i, 2).value))
            Exit Function
        End If
    Next i
End Function

Public Sub SaveAppSetting(settingKey As String, settingValue As String)
    Dim ws As Worksheet
    Set ws = GetSettingsSheet()
    If ws Is Nothing Then Exit Sub
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, 1).value)) = settingKey Then
            ws.Cells(i, 2).value = settingValue
            SaveAIFile
            Exit Sub
        End If
    Next i
    ws.Cells(lastRow + 1, 1).value = settingKey
    ws.Cells(lastRow + 1, 2).value = settingValue
    SaveAIFile
End Sub

Public Sub SaveApiKeyToHistory(newKey As String)
    newKey = Trim(newKey)
    If newKey = "" Then Exit Sub
    Dim ws As Worksheet
    Set ws = GetSettingsSheet()
    If ws Is Nothing Then Exit Sub
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim maxIdx As Long: maxIdx = 0
    Dim i As Long
    For i = 2 To lastRow
        Dim cellKey As String
        cellKey = Trim(CStr(ws.Cells(i, 1).value))
        If Left(cellKey, Len(APIKEY_PREFIX)) = APIKEY_PREFIX Then
            If Trim(CStr(ws.Cells(i, 2).value)) = newKey Then Exit Sub
            Dim idx As Long
            On Error Resume Next
            idx = CLng(mID(cellKey, Len(APIKEY_PREFIX) + 1))
            On Error GoTo 0
            If idx > maxIdx Then maxIdx = idx
        End If
    Next i
    ws.Cells(lastRow + 1, 1).value = APIKEY_PREFIX & (maxIdx + 1)
    ws.Cells(lastRow + 1, 2).value = newKey
    SaveAIFile
End Sub

' ============================================================
' API KEY
' ============================================================
Function GetApiKey() As String
    GetApiKey = GetAppSetting("api_key")
End Function

Sub SetApiKey(newKey As String)
    SaveAppSetting "api_key", Trim(newKey)
End Sub

' ============================================================
' UPDATE PROVIDER UI
' ============================================================
Private Sub UpdateProviderUI()
    Dim key As String
    key = GetApiKey()
    Dim provider As String
    provider = DetectProvider(key)

    lblProviderName.Caption = GetProviderDisplayName(provider)

    If key = "" Then
        lblKeyValue.Caption = "Click to enter API Key"
    ElseIf Len(key) > 12 Then
        lblKeyValue.Caption = Left(key, 6) & "..." & Right(key, 4)
    Else
        lblKeyValue.Caption = key
    End If

    cboModel.Clear
    Dim models() As String
    models = GetProviderModels(provider)
    Dim i As Integer
    For i = 0 To UBound(models)
        cboModel.AddItem models(i)
    Next i

    Dim savedModel As String
    savedModel = LoadSelectedModel()
    Dim found As Boolean
    found = False
    For i = 0 To cboModel.ListCount - 1
        If cboModel.List(i) = savedModel Then
            cboModel.ListIndex = i
            found = True
            Exit For
        End If
    Next i
    If Not found And cboModel.ListCount > 0 Then
        cboModel.ListIndex = 0
        SaveSelectedModel cboModel.List(0)
    End If
    UpdateVisionButton
End Sub

Private Sub cboModel_Change()
    If cboModel.ListIndex >= 0 Then
        SaveSelectedModel cboModel.List(cboModel.ListIndex)
    End If
    UpdateVisionButton
End Sub

' ============================================================
' CLICK KEY LABEL TO MANAGE KEYS
' ============================================================
Private Sub lblKeyValue_Click()
    Dim savedKeys() As String
    savedKeys = GetSavedApiKeys()

    Dim hasSaved As Boolean
    hasSaved = False
    If UBound(savedKeys) >= 0 Then
        If savedKeys(0) <> "" Then hasSaved = True
    End If

    Dim menu As String
    Dim i As Integer

    If hasSaved Then
        menu = "=== SAVED KEYS ===" & vbCrLf
        For i = 0 To UBound(savedKeys)
            If savedKeys(i) <> "" Then
                Dim provider As String
                provider = DetectProvider(savedKeys(i))
                menu = menu & "  " & (i + 1) & ". [" & GetProviderDisplayName(provider) & "] " & MaskKey(savedKeys(i)) & vbCrLf
            End If
        Next i
        menu = menu & vbCrLf
        menu = menu & "Enter number (1-" & (UBound(savedKeys) + 1) & ") to select key." & vbCrLf
        menu = menu & "Enter N to add new key." & vbCrLf
        menu = menu & "Enter D1, D2... to delete key."
    Else
        menu = "No keys saved." & vbCrLf & "Enter N to add new key."
    End If

    Dim choice As String
    choice = Trim(InputBox(menu, "Manage API Key"))
    If choice = "" Then Exit Sub

    If UCase(choice) = "N" Then
        Dim newKey As String
        newKey = Trim(InputBox( _
            "Enter new API Key:" & vbCrLf & vbCrLf & _
            "  gsk_...    -> Groq (free)" & vbCrLf & _
            "  nvapi-...  -> NVIDIA NIM (free)" & vbCrLf & _
            "  sk-or-...  -> OpenRouter (free)" & vbCrLf & _
            "  AIza...    -> Google Gemini" & vbCrLf & _
            "  sk-...     -> OpenAI", _
            "Add New API Key"))
        If newKey = "" Then Exit Sub
        SetApiKey newKey
        SaveApiKeyToHistory newKey
        If IsHexKey32(newKey) Then
            Dim azureEp As String
            azureEp = Trim(InputBox( _
                "Azure key detected (32 hex chars)." & vbCrLf & vbCrLf & _
                "Enter Azure endpoint URL:", _
                "Azure Endpoint", GetAzureEndpoint()))
            If azureEp <> "" Then SaveAzureEndpoint azureEp
        End If

    ElseIf UCase(Left(choice, 1)) = "D" Then
        Dim delIdxStr As String
        delIdxStr = mID(choice, 2)
        Dim delIdx As Long
        On Error Resume Next
        delIdx = CLng(delIdxStr) - 1
        On Error GoTo 0
        If delIdx >= 0 And delIdx <= UBound(savedKeys) And savedKeys(delIdx) <> "" Then
            Dim delKey As String
            delKey = savedKeys(delIdx)
            If GetApiKey() = delKey Then SetApiKey ""
            DeleteApiKeyFromHistory delKey
            MsgBox "Key deleted: " & MaskKey(delKey), vbInformation
        Else
            MsgBox "Invalid number.", vbExclamation
        End If

    Else
        Dim selIdx As Long
        On Error Resume Next
        selIdx = CLng(choice) - 1
        On Error GoTo 0
        If selIdx >= 0 And selIdx <= UBound(savedKeys) And savedKeys(selIdx) <> "" Then
            SetApiKey savedKeys(selIdx)
        Else
            MsgBox "Invalid number.", vbExclamation
            Exit Sub
        End If
    End If

    UpdateProviderUI
    lblRespone.value = ""
End Sub

' ============================================================
' TEXTBOX EVENTS
' ============================================================
Private Sub txtQuestion_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                   ByVal X As Single, ByVal Y As Single)
    With txtQuestion
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtQuestion_Enter()
    With txtQuestion
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtQuestion_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 0 Then
        KeyCode = 0
        btnAskAI_Click
    End If
End Sub

' ============================================================
' ASK AI BUTTON
' ============================================================
Private Sub btnAsk_Click()
    Dim question As String
    question = Trim(txtQuestion.Text)

    If question = "" And m_ImageBase64 = "" Then
        MsgBox "Please enter a question or select an image!", vbInformation
        txtQuestion.SetFocus
        Exit Sub
    End If

    If GetApiKey() = "" Then
        MsgBox "No API key! Click the key field to enter one.", vbInformation
        Exit Sub
    End If

    Dim selectedModel As String
    If cboModel.ListIndex >= 0 Then
        selectedModel = cboModel.List(cboModel.ListIndex)
    ElseIf cboModel.ListCount > 0 Then
        cboModel.ListIndex = 0
        selectedModel = cboModel.List(0)
    End If

    lblRespone.value = "AI processing..."
    Me.MousePointer = fmMousePointerHourGlass
    btnAskAI.Enabled = False
    DoEvents

    On Error GoTo ErrHandler

    Dim raw As String
    raw = AskAIWithImage(question, selectedModel, m_ImageBase64, m_ImageMime)
    If raw = "" Then GoTo Cleanup

    If InStr(raw, """error""") > 0 Or _
       (InStr(raw, """message""") > 0 And InStr(raw, """code""") > 0) Then
        Dim errMsg As String
        errMsg = ExtractJsonField(raw, "message")
        lblRespone.value = IIf(errMsg <> "", "API Error: " & errMsg, "API Error:" & vbCrLf & raw)
        GoTo Cleanup
    End If

    Dim answer As String
    answer = ParseAIContent(raw)
    If answer = "" Then
        lblRespone.value = "Could not parse response. Raw:" & vbCrLf & Left(raw, 300)
        GoTo Cleanup
    End If

    lblRespone.value = "AI: " & GetAIInfo(selectedModel) & vbCrLf & vbCrLf & answer
    GoTo Cleanup

ErrHandler:
    lblRespone.value = "Error: " & Err.Number & " - " & Err.Description

Cleanup:
    Me.MousePointer = fmMousePointerDefault
    btnAskAI.Enabled = True
    txtQuestion.SetFocus
End Sub

' ============================================================
' HISTORY BUTTON
' ============================================================
Private Sub btnHistory_Click()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetAISheet("AI_NOTES")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "No history saved.", vbInformation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No history saved.", vbInformation
        Exit Sub
    End If

    Dim startRow As Long
    startRow = IIf(lastRow > 51, lastRow - 49, 2)

    Dim menu As String
    menu = "=== SAVED QUESTION HISTORY ===" & vbCrLf & vbCrLf

    Dim entries() As Long
    Dim entryCount As Long
    ReDim entries(lastRow - startRow)

    Dim r As Long
    For r = lastRow To startRow Step -1
        Dim q As String
        q = Trim(CStr(ws.Cells(r, 2).value))
        If q <> "" Then
            entries(entryCount) = r
            entryCount = entryCount + 1
            Dim preview As String
            preview = Left(q, 60) & IIf(Len(q) > 60, "...", "")
            Dim timeStr As String
            On Error Resume Next
            timeStr = Format(ws.Cells(r, 1).value, "dd/mm hh:mm")
            On Error GoTo 0
            menu = menu & entryCount & ". [" & timeStr & "] " & preview & vbCrLf
        End If
    Next r

    If entryCount = 0 Then
        MsgBox "No history saved.", vbInformation
        Exit Sub
    End If

    menu = menu & vbCrLf & "Enter number to reload (1-" & entryCount & "):" & vbCrLf
    menu = menu & "Enter X to delete an entry (e.g. X3)."

    Dim choice As String
    choice = Trim(InputBox(menu, "Question History"))
    If choice = "" Then Exit Sub

    If UCase(Left(choice, 1)) = "X" Then
        Dim delNum As Long
        On Error Resume Next
        delNum = CLng(mID(choice, 2))
        On Error GoTo 0
        If delNum >= 1 And delNum <= entryCount Then
            Dim delRow As Long
            delRow = entries(delNum - 1)
            If MsgBox("Delete this entry?" & vbCrLf & Left(Trim(CStr(ws.Cells(delRow, 2).value)), 40), _
                      vbQuestion + vbYesNo) = vbYes Then
                ws.rows(delRow).Delete
                SaveAIFile
                MsgBox "Deleted.", vbInformation
            End If
        Else
            MsgBox "Invalid number.", vbExclamation
        End If
        Exit Sub
    End If

    Dim selNum As Long
    On Error Resume Next
    selNum = CLng(choice)
    On Error GoTo 0

    If selNum >= 1 And selNum <= entryCount Then
        Dim selRow As Long
        selRow = entries(selNum - 1)
        txtQuestion.Text = Trim(CStr(ws.Cells(selRow, 2).value))
        Dim loadedA As String
        loadedA = Trim(CStr(ws.Cells(selRow, 3).value))
        Dim dispProv As String
        Dim lProv As String: lProv = Trim(CStr(ws.Cells(selRow, 4).value))
        Dim lModel As String: lModel = Trim(CStr(ws.Cells(selRow, 5).value))
        If lProv <> "" And lModel <> "" Then
            dispProv = lProv & "  (" & lModel & ")"
        ElseIf lProv <> "" Then
            dispProv = lProv
        Else
            dispProv = "?"
        End If
        lblRespone.value = "AI: " & dispProv & vbCrLf & vbCrLf & loadedA
    Else
        MsgBox "Invalid number.", vbExclamation
    End If
End Sub

' ============================================================
' RIGHT-CLICK TO SAVE
' ============================================================
Private Sub lblRespone_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                  ByVal X As Single, ByVal Y As Single)
    If Button <> 2 Then Exit Sub
    If Trim(lblRespone.value) = "" Or lblRespone.value = "AI processing..." Then Exit Sub
    If MsgBox("Save this question and answer to notes?", vbQuestion + vbYesNo, "Save Note") = vbNo Then Exit Sub
    SaveAINoteToSheet
End Sub

' ============================================================
' SAVE TO AI_NOTES SHEET
' ============================================================
Private Sub SaveAINoteToSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetAISheet("AI_NOTES")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim selectedModel As String
    If cboModel.ListIndex >= 0 Then selectedModel = cboModel.List(cboModel.ListIndex)

    Dim fullText As String:     fullText = lblRespone.value
    Dim questionText As String: questionText = Trim(txtQuestion.Text)
    Dim answerText As String
    Dim sepPos As Long: sepPos = InStr(fullText, vbCrLf & vbCrLf)
    If sepPos > 0 Then
        answerText = Trim(mID(fullText, sepPos + 4))
    Else
        answerText = Trim(fullText)
    End If

    Dim newRow As Long
    newRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row + 1
    ws.Cells(newRow, 1).value = Now()
    ws.Cells(newRow, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Cells(newRow, 2).value = questionText
    ws.Cells(newRow, 3).value = answerText
    ws.Cells(newRow, 4).value = GetProviderDisplayName(DetectProvider(GetApiKey()))
    ws.Cells(newRow, 5).value = selectedModel
    ws.Cells(newRow, 2).WrapText = True
    ws.Cells(newRow, 3).WrapText = True
    SaveAIFile
    MsgBox "Saved to AI_NOTES!", vbInformation
End Sub

Public Sub DeleteApiKeyFromHistory(targetKey As String)
    targetKey = Trim(targetKey)
    Dim ws As Worksheet
    Set ws = GetSettingsSheet()
    If ws Is Nothing Then Exit Sub
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = lastRow To 2 Step -1
        Dim cellKey As String: cellKey = Trim(CStr(ws.Cells(i, 1).value))
        If Left(cellKey, Len(APIKEY_PREFIX)) = APIKEY_PREFIX Then
            If Trim(CStr(ws.Cells(i, 2).value)) = targetKey Then
                ws.rows(i).Delete
                SaveAIFile
                Exit For
            End If
        End If
    Next i
End Sub

Public Function GetSavedApiKeys() As String()
    Dim ws As Worksheet
    Set ws = GetSettingsSheet()

    Dim emptyKeys(0) As String
    emptyKeys(0) = ""

    If ws Is Nothing Then
        GetSavedApiKeys = emptyKeys
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, "A").End(xlUp).Row

    Dim count As Long: count = 0
    Dim i As Long
    For i = 2 To lastRow
        If Left(Trim(CStr(ws.Cells(i, 1).value)), Len(APIKEY_PREFIX)) = APIKEY_PREFIX Then
            count = count + 1
        End If
    Next i

    If count = 0 Then
        GetSavedApiKeys = emptyKeys
        Exit Function
    End If

    Dim keys() As String
    ReDim keys(count - 1)
    Dim k As Long: k = 0
    For i = 2 To lastRow
        If Left(Trim(CStr(ws.Cells(i, 1).value)), Len(APIKEY_PREFIX)) = APIKEY_PREFIX Then
            keys(k) = Trim(CStr(ws.Cells(i, 2).value))
            k = k + 1
        End If
    Next i

    GetSavedApiKeys = keys
End Function

' ============================================================
' HELPER FUNCTIONS
' ============================================================
Public Function MaskKey(key As String) As String
    If Len(key) <= 10 Then MaskKey = key Else MaskKey = Left(key, 6) & "..." & Right(key, 4)
End Function

Function LoadSelectedModel() As String
    LoadSelectedModel = GetAppSetting("selected_model_" & DetectProvider(GetApiKey()))
End Function

Sub SaveSelectedModel(m As String)
    SaveAppSetting "selected_model_" & DetectProvider(GetApiKey()), m
End Sub

Public Function GetAzureEndpoint() As String
    GetAzureEndpoint = GetAppSetting("azure_endpoint")
End Function

Public Sub SaveAzureEndpoint(endpoint As String)
    SaveAppSetting "azure_endpoint", endpoint
End Sub

Public Function IsHexKey32(s As String) As Boolean
    If Len(s) <> 32 Then Exit Function
    Dim i As Long
    For i = 1 To 32
        Dim c As String: c = mID(s, i, 1)
        If Not ((c >= "0" And c <= "9") Or (c >= "a" And c <= "f") Or (c >= "A" And c <= "F")) Then Exit Function
    Next i
    IsHexKey32 = True
End Function

' ============================================================
' DETECT PROVIDER FROM KEY FORMAT
' ============================================================
Public Function DetectProvider(apiKey As String) As String
    If IsHexKey32(apiKey) And GetAzureEndpoint() <> "" Then
        DetectProvider = "AZURE"
    ElseIf Left(apiKey, 4) = "csk-" Then
        DetectProvider = "CEREBRAS"
    ElseIf Left(apiKey, 11) = "github_pat_" Or Left(apiKey, 4) = "ghp_" Then
        DetectProvider = "GITHUB"
    ElseIf Left(apiKey, 4) = "gsk_" Then
        DetectProvider = "GROQ"
    ElseIf Left(apiKey, 6) = "nvapi-" Then
        DetectProvider = "NVIDIA"
    ElseIf Left(apiKey, 6) = "sk-or-" Then
        DetectProvider = "OPENROUTER"
    ElseIf Left(apiKey, 4) = "AIza" Then
        DetectProvider = "GEMINI"
    ElseIf Left(apiKey, 8) = "sk-proj-" Then
        DetectProvider = "OPENAI"
    ElseIf Left(apiKey, 3) = "sk-" Then
        DetectProvider = "OPENAI"
    Else
        DetectProvider = "GROQ"
    End If
End Function

Public Function GetProviderDisplayName(provider As String) As String
    Select Case provider
        Case "GROQ":       GetProviderDisplayName = "GROQ"
        Case "NVIDIA":     GetProviderDisplayName = "NVIDIA NIM"
        Case "OPENROUTER": GetProviderDisplayName = "OPENROUTER"
        Case "GEMINI":     GetProviderDisplayName = "GOOGLE GEMINI"
        Case "OPENAI":     GetProviderDisplayName = "OPENAI"
        Case "AZURE":      GetProviderDisplayName = "AZURE OPENAI"
        Case "CEREBRAS":   GetProviderDisplayName = "CEREBRAS"
        Case "GITHUB":     GetProviderDisplayName = "GITHUB MODELS"
        Case Else:         GetProviderDisplayName = "UNKNOWN"
    End Select
End Function

Public Function GetProviderModels(provider As String) As String()
    Dim models() As String
    Select Case provider
        Case "GROQ"
            models = Split("llama-4-scout-17b-16e-instruct|llama-4-maverick-17b-128e-instruct|llama-3.3-70b-versatile|llama-3.1-8b-instant|llama-3.2-11b-vision-preview|compound-beta|qwen-qwq-32b|gemma2-9b-it", "|")
        Case "NVIDIA"
            models = Split("meta/llama-4-scout-17b-16e-instruct|meta/llama-4-maverick-17b-128e-instruct|meta/llama-3.3-70b-instruct|meta/llama-3.1-8b-instruct|deepseek-ai/deepseek-r1|google/gemma-3-27b-it|qwen/qwq-32b", "|")
        Case "OPENROUTER"
            models = Split("meta-llama/llama-4-scout:free|meta-llama/llama-4-maverick:free|deepseek/deepseek-r1:free|deepseek/deepseek-v3:free|google/gemma-3-27b-it:free|qwen/qwq-32b:free|microsoft/phi-4:free|meta-llama/llama-3.3-70b-instruct:free", "|")
        Case "GEMINI"
            models = Split("gemini-2.5-flash|gemini-2.5-pro|gemini-2.0-flash|gemini-2.0-flash-lite|gemini-1.5-flash|gemini-1.5-pro", "|")
        Case "OPENAI"
            models = Split("gpt-4.1-mini|gpt-4.1|gpt-4.1-nano|gpt-4o|gpt-4o-mini|o4-mini|o3-mini", "|")
        Case "AZURE"
            models = Split("gpt-4.1|gpt-4o|gpt-4o-mini|gpt-4-turbo|gpt-35-turbo", "|")
        Case "CEREBRAS"
            models = Split("llama-4-scout-17b-16e-instruct|llama3.3-70b|llama3.1-70b|qwen-3-32b|llama3.1-8b", "|")
        Case "GITHUB"
            models = Split("gpt-4.1|gpt-4.1-mini|gpt-4o|o3-mini|Meta-Llama-3.3-70B-Instruct|Llama-4-Scout|DeepSeek-R1|Phi-4", "|")
        Case Else
            models = Split("llama-4-scout-17b-16e-instruct", "|")
    End Select
    GetProviderModels = models
End Function

Function AskAI(Prompt As String, Optional model As String = "") As String
    Dim apiKey As String: apiKey = GetApiKey()
    If apiKey = "" Then
        MsgBox "No API key!", vbExclamation
        AskAI = "": Exit Function
    End If
    Dim provider As String: provider = DetectProvider(apiKey)
    If model = "" Then
        Dim models() As String: models = GetProviderModels(provider): model = models(0)
    End If
    Select Case provider
        Case "GROQ":       AskAI = CallOpenAICompat(Prompt, apiKey, "https://api.groq.com/openai/v1/chat/completions", model)
        Case "NVIDIA":     AskAI = CallOpenAICompat(Prompt, apiKey, "https://integrate.api.nvidia.com/v1/chat/completions", model)
        Case "OPENROUTER": AskAI = CallOpenAICompat(Prompt, apiKey, "https://openrouter.ai/api/v1/chat/completions", model)
        Case "OPENAI":     AskAI = CallOpenAICompat(Prompt, apiKey, "https://api.openai.com/v1/chat/completions", model)
        Case "GEMINI":     AskAI = CallGemini(Prompt, apiKey, model)
        Case "CEREBRAS":   AskAI = CallOpenAICompat(Prompt, apiKey, "https://api.cerebras.ai/v1/chat/completions", model)
        Case "GITHUB":     AskAI = CallOpenAICompat(Prompt, apiKey, "https://models.inference.ai.azure.com/chat/completions", model)
        Case "AZURE"
            Dim azureUrl As String: azureUrl = GetAzureEndpoint()
            AskAI = IIf(azureUrl = "", "{""error"":""Azure endpoint not configured""}", _
                        CallOpenAICompat(Prompt, apiKey, azureUrl, model))
        Case Else: AskAI = "{""error"":""Unknown provider""}"
    End Select
End Function

Private Function CallOpenAICompat(Prompt As String, apiKey As String, URL As String, model As String) As String
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    Dim body As String
    body = "{""model"":""" & model & """,""messages"":[{""role"":""user"",""content"":""" & EscapeJson(Prompt) & """}]}"
    On Error GoTo ErrHTTP
    http.Open "POST", URL, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send body
    If http.status <> 200 Then
        CallOpenAICompat = "{""error"":""HTTP " & http.status & """,""message"":""" & EscapeJson(http.responseText) & """}"
        Exit Function
    End If
    CallOpenAICompat = http.responseText
    Exit Function
ErrHTTP:
    CallOpenAICompat = "{""error"":""" & EscapeJson(Err.Description) & """}"
End Function

Private Function CallGemini(Prompt As String, apiKey As String, model As String) As String
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    Dim URL As String
    URL = "https://generativelanguage.googleapis.com/v1beta/models/" & model & ":generateContent?key=" & apiKey
    Dim body As String
    body = "{""contents"":[{""parts"":[{""text"":""" & EscapeJson(Prompt) & """}]}]}"
    On Error GoTo ErrHTTP
    http.Open "POST", URL, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send body
    CallGemini = http.responseText
    Exit Function
ErrHTTP:
    CallGemini = "{""error"":""" & EscapeJson(Err.Description) & """}"
End Function

Function GetAIInfo(Optional model As String = "") As String
    Dim apiKey As String: apiKey = GetApiKey()
    If apiKey = "" Then GetAIInfo = "No API key": Exit Function
    Dim provider As String: provider = DetectProvider(apiKey)
    If model = "" Then
        Dim models() As String: models = GetProviderModels(provider): model = models(0)
    End If
    GetAIInfo = GetProviderDisplayName(provider) & "  (" & model & ")"
End Function

Function ParseAIContent(jsonText As String) As String
    Dim result As String: result = ""
    Dim marker As String: marker = """content"":"""
    Dim p1 As Long: p1 = InStr(jsonText, marker)
    If p1 = 0 Then marker = """text"":""": p1 = InStr(jsonText, marker)
    If p1 = 0 Then ParseAIContent = "": Exit Function
    p1 = p1 + Len(marker)
    Dim pos As Long: pos = p1
    Dim ch As String
    Do While pos <= Len(jsonText)
        ch = mID(jsonText, pos, 1)
        If ch = "\" Then
            pos = pos + 1
            If pos > Len(jsonText) Then Exit Do
            Select Case mID(jsonText, pos, 1)
                Case "n": result = result & vbCrLf
                Case "t": result = result & " "
                Case """": result = result & """"
                Case "\": result = result & "\"
                Case "u"
                    If pos + 4 <= Len(jsonText) Then
                        On Error Resume Next
                        result = result & ChrW(CLng("&H" & mID(jsonText, pos + 1, 4)))
                        On Error GoTo 0
                        pos = pos + 4
                    End If
            End Select
        ElseIf ch = """" Then
            Exit Do
        Else
            result = result & ch
        End If
        pos = pos + 1
    Loop
    ParseAIContent = result
End Function

Function EscapeJson(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr(10), "\n")
    s = Replace(s, Chr(13), "")
    s = Replace(s, Chr(9), " ")
    EscapeJson = s
End Function

Function ExtractJsonField(json As String, fieldName As String) As String
    Dim marker As String: marker = """" & fieldName & """:"""
    Dim p1 As Long: p1 = InStr(json, marker)
    If p1 = 0 Then ExtractJsonField = "": Exit Function
    p1 = p1 + Len(marker)
    Dim p2 As Long: p2 = InStr(p1, json, """")
    If p2 = 0 Then ExtractJsonField = "": Exit Function
    ExtractJsonField = mID(json, p1, p2 - p1)
End Function

Function AskAIWithImage(Prompt As String, model As String, _
                         Optional imgBase64 As String = "", _
                         Optional imgMime As String = "image/jpeg") As String
    Dim apiKey As String: apiKey = GetApiKey()
    If apiKey = "" Then AskAIWithImage = "": Exit Function

    Dim provider As String: provider = DetectProvider(apiKey)
    If model = "" Then
        Dim models() As String: models = GetProviderModels(provider)
        model = models(0)
    End If

    Select Case provider
        Case "GEMINI"
            AskAIWithImage = CallGeminiVision(Prompt, apiKey, model, imgBase64, imgMime)
        Case "GROQ", "NVIDIA", "OPENROUTER", "OPENAI", "AZURE", "CEREBRAS", "GITHUB"
            Dim URL As String
            Select Case provider
                Case "GROQ":       URL = "https://api.groq.com/openai/v1/chat/completions"
                Case "NVIDIA":     URL = "https://integrate.api.nvidia.com/v1/chat/completions"
                Case "OPENROUTER": URL = "https://openrouter.ai/api/v1/chat/completions"
                Case "OPENAI":     URL = "https://api.openai.com/v1/chat/completions"
                Case "CEREBRAS":   URL = "https://api.cerebras.ai/v1/chat/completions"
                Case "GITHUB":     URL = "https://models.inference.ai.azure.com/chat/completions"
                Case "AZURE":      URL = GetAzureEndpoint()
            End Select
            AskAIWithImage = CallOpenAICompatVision(Prompt, apiKey, URL, model, imgBase64, imgMime)
        Case Else
            AskAIWithImage = "{""error"":""Unknown provider""}"
    End Select
End Function

Private Function CallOpenAICompatVision(Prompt As String, apiKey As String, _
                                         URL As String, model As String, _
                                         imgBase64 As String, imgMime As String) As String
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")

    Dim textPart As String
    Dim imgPart As String
    Dim contentVal As String
    Dim body As String

    textPart = "{""type"":""text"",""text"":""" & EscapeJson(Prompt) & """}"

    If imgBase64 <> "" Then
        imgPart = "{""type"":""image_url"",""image_url"":{""url"":""data:" & imgMime & ";base64," & imgBase64 & """}}"
        contentVal = "[" & textPart & "," & imgPart & "]"
    Else
        contentVal = """" & EscapeJson(Prompt) & """"
    End If

    body = "{""model"":""" & model & """,""messages"":[{""role"":""user"",""content"":" & contentVal & "}]}"

    On Error GoTo ErrHTTP
    http.Open "POST", URL, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send body

    If http.status <> 200 Then
        CallOpenAICompatVision = "{""error"":""HTTP " & http.status & """,""message"":""" & EscapeJson(http.responseText) & """}"
        Exit Function
    End If
    CallOpenAICompatVision = http.responseText
    Exit Function
ErrHTTP:
    CallOpenAICompatVision = "{""error"":""" & EscapeJson(Err.Description) & """}"
End Function

Private Function CallGeminiVision(Prompt As String, apiKey As String, model As String, _
                                   imgBase64 As String, imgMime As String) As String
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    Dim URL As String
    URL = "https://generativelanguage.googleapis.com/v1beta/models/" & model & ":generateContent?key=" & apiKey

    Dim body As String
    Dim textPart As String
    Dim imgPart As String
    Dim partsList As String

    textPart = "{""text"":""" & EscapeJson(Prompt) & """}"

    If imgBase64 <> "" Then
        imgPart = "{""inline_data"":{""mime_type"":""" & imgMime & """,""data"":""" & imgBase64 & """}}"
        partsList = "[" & textPart & "," & imgPart & "]"
    Else
        partsList = "[" & textPart & "]"
    End If

    body = "{""contents"":[{""parts"":" & partsList & "}]}"

    On Error GoTo ErrHTTP
    http.Open "POST", URL, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send body
    CallGeminiVision = http.responseText
    Exit Function
ErrHTTP:
    CallGeminiVision = "{""error"":""" & EscapeJson(Err.Description) & """}"
End Function

' ============================================================
' VISION SUPPORT CHECK
' ============================================================
Public Function ModelSupportsVision(provider As String, model As String) As Boolean
    Select Case provider
        Case "GEMINI"
            ModelSupportsVision = True

        Case "OPENAI"
            Select Case model
                Case "gpt-4.1", "gpt-4.1-mini", "gpt-4.1-nano", "gpt-4o", "gpt-4o-mini", "o4-mini"
                    ModelSupportsVision = True
                Case Else
                    ModelSupportsVision = False
            End Select

        Case "AZURE"
            Select Case model
                Case "gpt-4.1", "gpt-4o", "gpt-4o-mini", "gpt-4-turbo"
                    ModelSupportsVision = True
                Case Else
                    ModelSupportsVision = False
            End Select

        Case "GROQ"
            Select Case model
                Case "llama-3.2-11b-vision-preview", "llama-4-scout-17b-16e-instruct", "llama-4-maverick-17b-128e-instruct"
                    ModelSupportsVision = True
                Case Else
                    ModelSupportsVision = False
            End Select

        Case "NVIDIA"
            Select Case model
                Case "meta/llama-4-scout-17b-16e-instruct", "meta/llama-4-maverick-17b-128e-instruct"
                    ModelSupportsVision = True
                Case Else
                    ModelSupportsVision = False
            End Select

        Case "OPENROUTER"
            Select Case model
                Case "meta-llama/llama-4-scout:free", "meta-llama/llama-4-maverick:free", _
                     "google/gemma-3-27b-it:free"
                    ModelSupportsVision = True
                Case Else
                    ModelSupportsVision = False
            End Select

        Case "GITHUB"
            Select Case model
                Case "gpt-4.1", "gpt-4.1-mini", "gpt-4o", "Llama-4-Scout"
                    ModelSupportsVision = True
                Case Else
                    ModelSupportsVision = False
            End Select

        Case Else
            ' CEREBRAS: no vision support
            ModelSupportsVision = False
    End Select
End Function

' ============================================================
' HELPER: CHECK IF CURRENT MODEL SUPPORTS VISION
' ============================================================
Private Function CurrentModelSupportsVision() As Boolean
    Dim provider As String
    provider = DetectProvider(GetApiKey())
    Dim selectedModel As String
    If cboModel.ListIndex >= 0 Then
        selectedModel = cboModel.List(cboModel.ListIndex)
    End If
    CurrentModelSupportsVision = ModelSupportsVision(provider, selectedModel)
End Function

' ============================================================
' UPDATE VISION BUTTON / LABEL STATE
' ============================================================
' Trong UpdateVisionButton — s?a l?i:
Private Sub UpdateVisionButton()
    Dim supportsVision As Boolean
    supportsVision = CurrentModelSupportsVision()

    If supportsVision Then
        If m_ImageBase64 <> "" Then
            imgPreview.Visible = True
            lblClickUpload.Visible = False
        Else
            imgPreview.Visible = False
            CenterUploadLabel    ' ? can gi?a tru?c khi hi?n
            lblClickUpload.Visible = True
        End If
    Else
        imgPreview.Visible = False
        lblClickUpload.Visible = False
        ClearImage
    End If
End Sub

' ============================================================
' IMAGE PREVIEW — CLICK TO UPLOAD (ch? khi h? tr? vision)
' ============================================================
Private Sub imgPreview_Click()
    If Not CurrentModelSupportsVision() Then Exit Sub
    AttachImage
End Sub

' ============================================================
' ATTACH IMAGE — fix update ?nh l?n 2 tr? di
' ============================================================
Private Sub AttachImage()
    Dim fp As Variant
    fp = Application.GetOpenFilename( _
            FileFilter:="Image Files (*.jpg;*.jpeg;*.png;*.gif;*.webp),*.jpg;*.jpeg;*.png;*.gif;*.webp", _
            title:="Select image")
    If VarType(fp) = vbBoolean Then Exit Sub

    m_ImagePath = CStr(fp)
    m_ImageBase64 = ImageToBase64(m_ImagePath)

    If m_ImageBase64 = "" Then
        MsgBox "Cannot read image.", vbExclamation
        m_ImagePath = ""
        Exit Sub
    End If

    Select Case LCase$(mID$(m_ImagePath, InStrRev(m_ImagePath, ".") + 1))
        Case "jpg", "jpeg": m_ImageMime = "image/jpeg"
        Case "png":         m_ImageMime = "image/png"
        Case "gif":         m_ImageMime = "image/gif"
        Case "webp":        m_ImageMime = "image/webp"
        Case Else:          m_ImageMime = "image/jpeg"
    End Select

    lblImageName.Caption = mID$(m_ImagePath, InStrRev(m_ImagePath, "\") + 1)

    ' *** KEY FIX: ?n control ? clear ? DoEvents ? load m?i ? hi?n l?i ***
    imgPreview.Visible = False
    lblClickUpload.Visible = False

    On Error Resume Next
    imgPreview.Picture = LoadPicture("")   ' Clear ?nh cu
    On Error GoTo 0
    DoEvents                               ' Flush render buffer

    imgPreview.PictureSizeMode = 3         ' fmPictureSizeModeZoom
    imgPreview.PictureAlignment = 2        ' fmPictureAlignmentCenter

    On Error Resume Next
    imgPreview.Picture = LoadPicture(m_ImagePath)   ' Load ?nh m?i
    Dim loadErr As Long: loadErr = Err.Number
    On Error GoTo 0

    If loadErr <> 0 Then
        MsgBox "Cannot display image preview.", vbExclamation
        m_ImagePath = ""
        m_ImageBase64 = ""
        m_ImageMime = ""
        lblImageName.Caption = "(No image selected)"
        CenterUploadLabel
        lblClickUpload.Visible = CurrentModelSupportsVision()
        Exit Sub
    End If

    imgPreview.Visible = True   ' Hi?n l?i SAU KHI dã load xong
    DoEvents
End Sub
' Trong ClearImage — s?a l?i:
Private Sub ClearImage()
    m_ImagePath = ""
    m_ImageBase64 = ""
    m_ImageMime = ""
    lblImageName.Caption = "(No image selected)"
    On Error Resume Next
    imgPreview.Picture = LoadPicture("")
    On Error GoTo 0
    imgPreview.Visible = False
    If CurrentModelSupportsVision() Then
        CenterUploadLabel    ' ? can gi?a tru?c khi hi?n
        lblClickUpload.Visible = True
    Else
        lblClickUpload.Visible = False
    End If
End Sub
Private Sub lblClickUpload_Click()
    If Not CurrentModelSupportsVision() Then Exit Sub
    AttachImage
End Sub

Private Sub CenterUploadLabel()
    ' Tính s? dòng tr?ng c?n thêm d? text xu?ng gi?a
    ' Font m?c d?nh ~10pt/dòng trong VBA form units
    Dim linesAbove As Integer
    linesAbove = Int(lblClickUpload.Height / 8 / 2) - 1
    If linesAbove < 0 Then linesAbove = 0
    lblClickUpload.Caption = String(linesAbove, vbCrLf) & "Click to upload"
End Sub
