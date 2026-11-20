Attribute VB_Name = "modLicense"
Option Explicit

' =========================================
' LICENSE SYSTEM (PRO VERSION - TIMED)
' Key format: yyyymmdd + 8-char hash = 16 chars
' =========================================
#If VBA7 Then

    Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal cb As LongPtr)
    Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

#End If

Private Const SECRET_KEY   As String = "CBIMS2025"
Private Const LICENSE_FILE As String = "license.key"
Private Const APP_FOLDER   As String = "\CPMS_R\"
Private Const PERM_DATE    As String = "99991231"

Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE  As Long = &H2

Public Const DUR_7DAYS   As String = "7D"
Public Const DUR_3MONTHS As String = "3M"
Public Const DUR_6MONTHS As String = "6M"
Public Const DUR_1YEAR   As String = "1Y"
Public Const DUR_PERM    As String = "PERM"

' =========================================
' PATHS
' =========================================
Private Function LicensePath() As String
    LicensePath = GetAppFolder & LICENSE_FILE
End Function

Public Function GetAppFolder() As String
    Dim p As String
    p = Environ$("LOCALAPPDATA") & APP_FOLDER
    If Dir(p, vbDirectory) = "" Then MkDir p
    GetAppFolder = p
End Function

Public Function HasLicense() As Boolean
    HasLicense = (Dir(LicensePath) <> "")
End Function

' =========================================
' KEY GENERATION
' =========================================
Public Function GenerateLicenseKey(sMachineID As String, expiryDate As Date) As String
    Dim expStr As String
    expStr = Format(expiryDate, "yyyymmdd")
    GenerateLicenseKey = expStr & Left$(MD5Hash(sMachineID & SECRET_KEY & expStr), 8)
End Function

Public Function GenerateKeyForDuration(sMachineID As String, _
                                       durationCode As String) As String
    Dim expDate As Date
    Select Case durationCode
        Case DUR_7DAYS:   expDate = Date + 7
        Case DUR_3MONTHS: expDate = DateAdd("m", 3, Date)
        Case DUR_6MONTHS: expDate = DateAdd("m", 6, Date)
        Case DUR_1YEAR:   expDate = DateAdd("yyyy", 1, Date)
        Case Else:        expDate = CDate("9999-12-31")
    End Select
    GenerateKeyForDuration = GenerateLicenseKey(sMachineID, expDate)
End Function

' =========================================
' VALIDATE KEY FORMAT
' Returns expiry string "yyyymmdd" if valid, "" if not
' =========================================
Private Function ValidateKeyFormat(sMachineID As String, _
                                   userKey As String) As String
    If Len(userKey) <> 16 Then Exit Function

    Dim expStr   As String: expStr = Left$(userKey, 8)
    Dim hashPart As String: hashPart = Right$(userKey, 8)

    If hashPart <> Left$(MD5Hash(sMachineID & SECRET_KEY & expStr), 8) Then Exit Function

    On Error Resume Next
    Dim d As Date
    d = CDate(Left$(expStr, 4) & "-" & mID(expStr, 5, 2) & "-" & Right$(expStr, 2))
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0

    ValidateKeyFormat = expStr
End Function

' =========================================
' ACTIVATE
' =========================================
Public Function TryActivate(userKey As String) As Boolean
    Dim sMachineID As String: sMachineID = GetMachineID()
    Dim expStr     As String: expStr = ValidateKeyFormat(sMachineID, userKey)

    If expStr = "" Then Exit Function

    If expStr <> PERM_DATE Then
        Dim expDate As Date
        expDate = CDate(Left$(expStr, 4) & "-" & mID(expStr, 5, 2) & "-" & Right$(expStr, 2))
        If expDate < Date Then Exit Function
    End If

    SaveLicense userKey, expStr
    TryActivate = True
End Function

' =========================================
' SAVE LICENSE
' =========================================
Private Sub SaveLicense(key As String, expStr As String)
    Dim actDate As String: actDate = Format(Now(), "yyyy-mm-dd")
    Dim expFormatted As String

    If expStr = PERM_DATE Then
        expFormatted = "PERMANENT"
    Else
        expFormatted = Left$(expStr, 4) & "-" & mID(expStr, 5, 2) & "-" & Right$(expStr, 2)
    End If

    Dim raw As String
    raw = GetMachineID() & "|" & key & "|" & actDate & "|" & expFormatted

    Dim enc As String: enc = XorCrypt(raw)
    Dim f   As Integer: f = FreeFile
    Open LicensePath For Binary As #f
        Put #f, , enc
    Close #f
End Sub

' =========================================
' CLEAR LICENSE
' =========================================
Public Sub ClearLicense(Optional dummy As Variant = 0)
    If Dir(LicensePath) <> "" Then Kill LicensePath
End Sub

' =========================================
' READ LICENSE
' =========================================
Private Function ReadLicense(ByRef sMachineID As String, _
                             ByRef key As String, _
                             ByRef actDate As String, _
                             ByRef expDate As String) As Boolean
    On Error GoTo BAD_FILE

    If Dir(LicensePath) = "" Then Exit Function

    Dim enc As String, f As Integer
    f = FreeFile
    Open LicensePath For Binary As #f
        enc = Space$(LOF(f))
        Get #f, , enc
    Close #f

    Dim raw As String: raw = XorCrypt(enc)
    Dim arr() As String: arr = Split(raw, "|")

    If UBound(arr) = 3 Then
        sMachineID = arr(0): key = arr(1)
        actDate = arr(2): expDate = arr(3)
        ReadLicense = True
    Else
        GoTo BAD_FILE
    End If
    Exit Function

BAD_FILE:
    On Error Resume Next
    Kill LicensePath
    ReadLicense = False
End Function

' =========================================
' CORE CHECK
' =========================================
Public Function IsActivated() As Boolean
    Dim sMachineID As String, key     As String
    Dim actDate    As String, expDate As String

    If Not ReadLicense(sMachineID, key, actDate, expDate) Then Exit Function
    If sMachineID <> GetMachineID() Then Exit Function

    Dim expStr As String: expStr = ValidateKeyFormat(sMachineID, key)
    If expStr = "" Then Exit Function

    If expStr = PERM_DATE Then
        IsActivated = True
    Else
        Dim expDateFromKey As Date
        expDateFromKey = CDate(Left$(expStr, 4) & "-" & _
                               mID(expStr, 5, 2) & "-" & _
                               Right$(expStr, 2))
        If expDateFromKey >= Date Then IsActivated = True
    End If
End Function

' =========================================
' GET LICENSE INFO (for UI display)
' =========================================
Public Function GetLicenseInfo(ByRef actDate As String, _
                               ByRef expDate As String, _
                               ByRef daysLeft As Long) As Boolean
    Dim sMachineID As String, key As String

    If Not ReadLicense(sMachineID, key, actDate, expDate) Then Exit Function
    If sMachineID <> GetMachineID() Then Exit Function

    Dim expStr As String: expStr = ValidateKeyFormat(sMachineID, key)
    If expStr = "" Then Exit Function

    If expStr = PERM_DATE Then
        expDate = "PERMANENT"
        daysLeft = -1
    Else
        Dim expDateFromKey As Date
        expDateFromKey = CDate(Left$(expStr, 4) & "-" & _
                               mID(expStr, 5, 2) & "-" & _
                               Right$(expStr, 2))
        expDate = Format(expDateFromKey, "yyyy-mm-dd")
        daysLeft = DateDiff("d", Date, expDateFromKey)
    End If

    GetLicenseInfo = True
End Function




' =========================================
' CLIPBOARD
' =========================================
Public Sub CopyTextToClipboard(ByVal txt As String)
    Dim hMem As LongPtr, pMem As LongPtr, bytes As LongPtr
    bytes = LenB(txt) + 2
    hMem = GlobalAlloc(GMEM_MOVEABLE, bytes)
    pMem = GlobalLock(hMem)
    CopyMemory pMem, StrPtr(txt), LenB(txt)
    GlobalUnlock hMem
    OpenClipboard 0
    EmptyClipboard
    SetClipboardData CF_UNICODETEXT, hMem
    CloseClipboard
End Sub

' =========================================
' XOR CRYPT
' =========================================
Private Function XorCrypt(txt As String) As String
    Dim i As Long, s As String
    For i = 1 To Len(txt)
        s = s & Chr(Asc(mID(txt, i, 1)) Xor _
                     Asc(mID(SECRET_KEY, (i Mod Len(SECRET_KEY)) + 1, 1)))
    Next
    XorCrypt = s
End Function

