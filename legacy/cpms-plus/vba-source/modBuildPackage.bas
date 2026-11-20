Attribute VB_Name = "modBuildPackage"
'Attribute VB_Name = "modBuildPackage"
Option Explicit

'================================================================
' COLUMN INDEX trong gArrDoc (range A2:S c?a DOCUMENT_TRACKER)
' Ði?u ch?nh n?u c?u trúc sheet khác
'================================================================
Private Const COL_DOC_CONTRACTID As Long = 2   ' Col B = contractID
Private Const COL_DOC_DOCTYPE    As Long = 3   ' Col C = docType code
Private Const COL_DOC_FILEPATH   As Long = 6  ' Col S = file path

'================================================================
' MAP DOCTYPE -> SUBFOLDER trong thu m?c contract
' APP và VO n?m TRONG Contract (cùng v?i CON)
'================================================================
Private Function DocTypeToSubFolder(ByVal dt As String) As String
    Select Case UCase$(Trim$(dt))
        Case "CON", "APP", "VO"
            DocTypeToSubFolder = "01_Contract"          ' Appendix n?m trong Contract
        Case "IPC"
            DocTypeToSubFolder = "02_Payment"
        Case "SPT", "ITP", "MS", "MAS"
            DocTypeToSubFolder = "03_Planning"
        Case "RFA", "RFC", "RFI", "SI", "NCR", "CAR", "TRN"
            DocTypeToSubFolder = "04_Technical"
        Case "AD", "SD", "ABD"
            DocTypeToSubFolder = "05_Drawings"
        Case "MOM", "LTR", "CMT", "SR", "MOH"
            DocTypeToSubFolder = "06_Correspondence"
        Case Else
            DocTypeToSubFolder = ""
    End Select
End Function

'================================================================
' KI?M TRA DOCTYPE CÓ ÐU?C CH?N
'================================================================
Private Function DocTypeIncluded( _
        ByVal dt As String, _
        ByVal inclContract As Boolean, _
        ByVal inclPayment As Boolean, _
        ByVal inclQuality As Boolean, _
        ByVal inclTech As Boolean, _
        ByVal inclDraw As Boolean, _
        ByVal inclDocCtrl As Boolean) As Boolean

    Select Case UCase$(Trim$(dt))
        Case "CON", "APP", "VO": DocTypeIncluded = inclContract
        Case "IPC": DocTypeIncluded = inclPayment
        Case "SPT", "ITP", "MS", "MAS": DocTypeIncluded = inclQuality
        Case "RFA", "RFC", "RFI", "SI", "NCR", "CAR", "TRN": DocTypeIncluded = inclTech
        Case "AD", "SD", "ABD": DocTypeIncluded = inclDraw
        Case "MOM", "LTR", "CMT", "SR", "MOH": DocTypeIncluded = inclDocCtrl
        Case Else: DocTypeIncluded = False
    End Select
End Function

'================================================================
' SANITIZE CHU?I Ð? DÙNG LÀM TÊN THU M?C
'================================================================
Private Function MakeSafeFileName(ByVal s As String) As String
    Dim result As String
    result = s
    Dim badChars As String
    badChars = "\/:*?""<>|"
    Dim k As Long
    For k = 1 To Len(badChars)
        result = Join(Split(result, mID$(badChars, k, 1)), "-")
    Next k
    ' C?t kho?ng tr?ng th?a và d?u - th?a
    Do While InStr(result, "  ") > 0
        result = Join(Split(result, "  "), " ")
    Loop
    Do While InStr(result, "--") > 0
        result = Join(Split(result, "--"), "-")
    Loop
    MakeSafeFileName = Trim$(result)
End Function

'================================================================
' T?O THU M?C Ð? QUY (t?o c? parent n?u chua có)
'================================================================
Private Sub EnsureFolder(ByVal fso As Object, ByVal folderPath As String)
    If fso.FolderExists(folderPath) Then Exit Sub
    Dim parent As String
    parent = fso.GetParentFolderName(folderPath)
    If Not fso.FolderExists(parent) Then EnsureFolder fso, parent
    fso.CreateFolder folderPath
End Sub

'================================================================
' X? LÝ TRÙNG TÊN FILE
'================================================================
Private Function UniqueFilePath( _
        ByVal fso As Object, _
        ByVal folder As String, _
        ByVal fName As String) As String

    Dim candidate As String
    candidate = folder & fName
    If Not fso.FileExists(candidate) Then
        UniqueFilePath = candidate
        Exit Function
    End If

    Dim base As String, ext As String
    Dim dotPos As Long
    dotPos = InStrRev(fName, ".")
    If dotPos > 0 Then
        base = Left$(fName, dotPos - 1)
        ext = mID$(fName, dotPos)
    Else
        base = fName: ext = ""
    End If

    Dim n As Long: n = 1
    Do
        candidate = folder & base & "_" & Format(n, "00") & ext
        If Not fso.FileExists(candidate) Then
            UniqueFilePath = candidate
            Exit Function
        End If
        n = n + 1
    Loop
End Function

'================================================================
' ZIP B?NG POWERSHELL Compress-Archive (Windows 10+)
'================================================================
Private Function ZipFolderWithPowerShell( _
        ByVal srcFolder As String, _
        ByVal zipPath As String) As Boolean

    Dim src As String
    src = srcFolder
    If Right$(src, 1) = "\" Then src = Left$(src, Len(src) - 1)

    On Error Resume Next
    Kill zipPath
    On Error GoTo 0

    Dim psCmd As String
    psCmd = "Compress-Archive -Path """ & src & "\*"" " & _
            "-DestinationPath """ & zipPath & """ -Force"

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "cmd /c powershell.exe -NoProfile -NonInteractive -Command """ & _
            psCmd & """", 0, True
    Set wsh = Nothing

    Dim fso2 As Object
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    ZipFolderWithPowerShell = fso2.FileExists(zipPath)
    Set fso2 = Nothing
End Function

'================================================================
' HÀM CHÍNH: BUILD ZIP PACKAGE
'================================================================
Public Sub BuildDocumentZipPackage( _
        ByVal outFolder As String, _
        ByVal lstOrderMap As Object, _
        ByVal selectAllContracts As Boolean, _
        ByVal currentContractID As String, _
        ByVal catFilter As String, _
        ByVal inclContract As Boolean, _
        ByVal inclPayment As Boolean, _
        ByVal inclQuality As Boolean, _
        ByVal inclTech As Boolean, _
        ByVal inclDraw As Boolean, _
        ByVal inclDocCtrl As Boolean)

    On Error GoTo ERR_HANDLER

    '================================================
    ' Ð?M B?O CACHE ÐÃ LOAD
    '================================================
    If Not gDocLoaded Then BuildDocTrackerCache
    If Not gContractLoaded Then BuildGlobalContractCache

    If IsEmpty(gArrDoc) Then
        ShowMsgCenter "No documents found in DOCUMENT_TRACKER.", vbExclamation
        Exit Sub
    End If

    '================================================
    ' XÁC Ð?NH DANH SÁCH CONTRACT C?N XU?T
    ' Key = contractID, Value = "folderName" (A-1 (ContractNo))
    '================================================
    Dim exportContracts As Object
    Set exportContracts = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim cid As String
    Dim cno As String
    Dim st  As String
    Dim catVal As String
    Dim folderName As String
    Dim parts() As String

    If selectAllContracts Then
        '--- Xu?t t?t c? contract theo th? t? lstResult ---
        ' lstOrderMap: contractID -> "seqIdx|||contractNo"
        Dim keys As Variant
        keys = lstOrderMap.keys
        Dim vals As Variant
        vals = lstOrderMap.items

        ' Duy?t gArrContract d? l?c CANCEL và category
        Dim r As Long
        Dim cidMap As Object
        Set cidMap = CreateObject("Scripting.Dictionary")

        For i = 1 To UBound(gArrContract, 1)
            cid = Trim$(CStr(gArrContract(i, 1)))
            If cid = "" Then GoTo NEXT_CON_ALL

            ' B? CANCEL
            st = UCase$(Trim$(CStr(gArrContract(i, 6))))
            If st = "CANCEL" Then GoTo NEXT_CON_ALL

            ' L?c category
            If catFilter <> "" Then
                catVal = Trim$(CStr(gArrContract(i, 18)))
                If LCase$(catVal) <> LCase$(catFilter) Then GoTo NEXT_CON_ALL
            End If

            cidMap(cid) = True

NEXT_CON_ALL:
        Next i

        ' L?y theo th? t? lstOrderMap
        Dim k As Long
        For k = 0 To UBound(keys)
            cid = CStr(keys(k))
            If cidMap.Exists(cid) Then
                parts = Split(CStr(vals(k)), "|||")
                cno = IIf(UBound(parts) >= 1, parts(1), "")
                folderName = BuildFolderName(cid, cno)
                exportContracts(cid) = folderName
            End If
        Next k

        Set cidMap = Nothing

    Else
        '--- Ch? xu?t contract hi?n t?i ---
        cid = Trim$(currentContractID)
        If cid = "" Then
            ShowMsgCenter "No contract selected.", vbExclamation
            Exit Sub
        End If

        ' L?y contractNo t? lstOrderMap
        cno = ""
        If lstOrderMap.Exists(cid) Then
            parts = Split(CStr(lstOrderMap(cid)), "|||")
            If UBound(parts) >= 1 Then cno = parts(1)
        End If
        If cno = "" And gDictContract.Exists(cid) Then
            r = gDictContract(cid)
            cno = Trim$(CStr(gArrContract(r, 2)))
        End If

        folderName = BuildFolderName(cid, cno)
        exportContracts(cid) = folderName
    End If

    If exportContracts.count = 0 Then
        ShowMsgCenter "No contracts match the current filter.", vbExclamation
        Exit Sub
    End If

    '================================================
    ' T?O THU M?C T?M
    '================================================
    Dim pkgName As String
    pkgName = "DocPackage_" & Format(Now, "yyyymmdd_hhnnss")

    Dim tmpBase As String
    tmpBase = outFolder & pkgName & "\"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder tmpBase

    ' T?o s?n GROUP-A và GROUP-B
    fso.CreateFolder tmpBase & "GROUP-A\"
    fso.CreateFolder tmpBase & "GROUP-B\"

    '================================================
    ' DUY?T gArrDoc, COPY FILE VÀO ÐÚNG THU M?C
    '================================================
    Dim folderCache As Object
    Set folderCache = CreateObject("Scripting.Dictionary")

    Dim copiedCount  As Long
    Dim missingCount As Long
    copiedCount = 0
    missingCount = 0

    Dim docType   As String
    Dim filePath  As String
    Dim subFolder As String
    Dim tgtFolder As String
    Dim tgtPath   As String
    Dim fName     As String
    Dim grpFolder As String

    For i = 1 To UBound(gArrDoc, 1)

        cid = Trim$(CStr(gArrDoc(i, COL_DOC_CONTRACTID)))
        If Not exportContracts.Exists(cid) Then GoTo NEXT_DOC

        docType = Trim$(CStr(gArrDoc(i, COL_DOC_DOCTYPE)))
        If Not DocTypeIncluded(docType, inclContract, inclPayment, _
                               inclQuality, inclTech, inclDraw, inclDocCtrl) Then
            GoTo NEXT_DOC
        End If

        filePath = Trim$(CStr(gArrDoc(i, COL_DOC_FILEPATH)))
        If filePath = "" Then GoTo NEXT_DOC

        If Not fso.FileExists(filePath) Then
            missingCount = missingCount + 1
            GoTo NEXT_DOC
        End If

        ' Xác d?nh GROUP
        If Left$(UCase$(cid), 2) = "A-" Then
            grpFolder = "GROUP-A"
        ElseIf Left$(UCase$(cid), 2) = "B-" Then
            grpFolder = "GROUP-B"
        Else
            grpFolder = "GROUP-A"   ' fallback
        End If

        subFolder = DocTypeToSubFolder(docType)
        If subFolder = "" Then GoTo NEXT_DOC

        ' tmpBase\GROUP-A\A-1 (ContractNo)\01_Contract\
        tgtFolder = tmpBase & grpFolder & "\" & _
                    CStr(exportContracts(cid)) & "\" & subFolder & "\"

        If Not folderCache.Exists(tgtFolder) Then
            EnsureFolder fso, tgtFolder
            folderCache(tgtFolder) = 1
        End If

        fName = fso.GetFileName(filePath)
        tgtPath = UniqueFilePath(fso, tgtFolder, fName)

        fso.CopyFile filePath, tgtPath
        copiedCount = copiedCount + 1

NEXT_DOC:
    Next i

    '================================================
    ' NÉN ZIP
    '================================================
    Dim zipPath As String
    zipPath = outFolder & pkgName & ".zip"

    Dim zipOK As Boolean
    zipOK = ZipFolderWithPowerShell(tmpBase, zipPath)

    '================================================
    ' XÓA THU M?C T?M
    '================================================
    On Error Resume Next
    fso.DeleteFolder tmpBase, True
    On Error GoTo 0

    '================================================
    ' BÁO K?T QU?
    '================================================
    If zipOK Then
        Dim msg As String
        msg = "Package created successfully!" & vbCrLf & vbCrLf & _
              "File: " & zipPath & vbCrLf & _
              "Contracts: " & exportContracts.count & vbCrLf & _
              "Files copied: " & copiedCount
        If missingCount > 0 Then
            msg = msg & vbCrLf & "Files not found (skipped): " & missingCount
        End If
        ShowMsgCenter msg, vbInformation, "Build Package"
    Else
        ShowMsgCenter _
            "ZIP creation failed." & vbCrLf & _
            "Requires PowerShell + Windows 10+." & vbCrLf & _
            "Staged folder (not deleted): " & outFolder, _
            vbCritical, "Build Package"
    End If

Cleanup:
    Set fso = Nothing
    Set exportContracts = Nothing
    Set folderCache = Nothing
    Exit Sub

ERR_HANDLER:
    On Error Resume Next
    If Not fso Is Nothing Then
        If fso.FolderExists(tmpBase) Then fso.DeleteFolder tmpBase, True
    End If
    On Error GoTo 0
    Set fso = Nothing
    ShowMsgCenter "Error " & Err.Number & ": " & Err.Description, _
                  vbCritical, "Build Package"
    Resume Cleanup

End Sub

'================================================================
' T?O TÊN THU M?C: "A-1 (ContractNo)"
'================================================================
Private Function BuildFolderName(ByVal cid As String, _
                                  ByVal cno As String) As String
    Dim safeCno As String
    safeCno = MakeSafeFileName(cno)
    If safeCno <> "" Then
        BuildFolderName = cid & " (" & safeCno & ")"
    Else
        BuildFolderName = cid
    End If
End Function

