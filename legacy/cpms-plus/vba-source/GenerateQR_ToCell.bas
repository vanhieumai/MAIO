Attribute VB_Name = "GenerateQR_ToCell"
Option Explicit

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
Alias "URLDownloadToFileA" ( _
ByVal pCaller As LongPtr, _
ByVal szURL As String, _
ByVal szFileName As String, _
ByVal dwReserved As LongPtr, _
ByVal lpfnCB As LongPtr) As Long



'========================================
' PUBLIC – L?Y MACHINE ID DUY NH?T
'========================================
Function GetMachineID() As String

    Dim raw As String

    raw = GetDriveSerial() & "|" & _
          GetBIOSSerial() & "|" & _
          Environ$("COMPUTERNAME")

    GetMachineID = MD5Hash(raw)

End Function


Public Sub InsertQR(filePath As String)

    Dim ws As Worksheet
    Dim pic As Picture

    Set ws = ActiveSheet

    For Each pic In ws.Pictures
        If pic.Name Like "QR_IMAGE*" Then pic.Delete
    Next

    Set pic = ws.Pictures.Insert(filePath)

    With pic
        .Name = "QR_IMAGE"
        .Left = ws.Range("G2").Left
        .Top = ws.Range("G2").Top
        .Width = 170
        .Height = 170
    End With

End Sub

Sub TaoQR_VietQR_Only(Optional dummy As Variant = 0)

    Dim bank As String, acc As String, nameAcc As String
    Dim amount As String, note As String
    Dim machineID As String

    bank = "MB"
    acc = "0904800002"
    nameAcc = "MAI VAN HIEU"
    amount = ""
    note = GetMachineID


    Dim URL As String
    Dim filePath As String

    ' ? template qr_only
    URL = "https://img.vietqr.io/image/" & _
          bank & "-" & acc & "-qr_only.png" & _
          "?amount=" & amount & _
          "&addInfo=" & WorksheetFunction.EncodeURL(note) & _
          "&accountName=" & WorksheetFunction.EncodeURL(nameAcc)

    filePath = Environ$("TEMP") & "\qr_only.png"

    URLDownloadToFile 0, URL, filePath, 0, 0

    InsertQR filePath

End Sub

Public Function GetDriveSerial() As String

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    GetDriveSerial = Hex(fso.GetDrive("C:").SerialNumber)

End Function

Public Function GetBIOSSerial() As String

    Dim objWMI As Object, col As Object, itm As Object

    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set col = objWMI.ExecQuery("Select * from Win32_BIOS")

    For Each itm In col
        GetBIOSSerial = itm.SerialNumber
        Exit For
    Next

End Function

Public Function MD5Hash(ByVal txt As String) As String

    Dim enc As Object
    Dim bytes() As Byte
    Dim hash() As Byte
    Dim i As Long
    Dim s As String

    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

    bytes = StrConv(txt, vbFromUnicode)
    hash = enc.ComputeHash_2(bytes)

    For i = 0 To UBound(hash)
        s = s & LCase(Right$("0" & Hex(hash(i)), 2))
    Next

    ' l?y 16 ký t? cho g?n
    MD5Hash = UCase(Left$(s, 16))

End Function


'==============================
' T?O QR CHUY?N KHO?N NGÂN HÀNG (VIETQR)
'========================================
Sub TaoQR_ChuyenKhoan(Optional dummy As Variant = 0)

    Dim bank As String
    Dim acc As String
    Dim nameAcc As String
    Dim amount As String
    Dim note As String

    ' ===== L?Y D? LI?U T? SHEET =====
    bank = "MB"          ' mã bank (VCB, TCB, BIDV...)
    acc = "0904800002"           ' s? tài kho?n
    nameAcc = "Mai Van Hieu"      ' tên tài kho?n
    amount = ""        ' s? ti?n (có th? tr?ng)
    note = Range("E2").value          ' n?i dung

    If bank = "" Or acc = "" Then
        MsgBox "Thi?u Bank ho?c S? TK", vbExclamation
        Exit Sub
    End If

    ' ===== ENCODE URL =====
    Dim URL As String

    URL = "https://img.vietqr.io/image/" & _
          bank & "-" & acc & "-print.png" & _
          "?amount=" & amount & _
          "&addInfo=" & WorksheetFunction.EncodeURL(note) & _
          "&accountName=" & WorksheetFunction.EncodeURL(nameAcc)

    ' ===== XÓA ?NH CU =====
    Dim pic As Picture
    For Each pic In ActiveSheet.Pictures
        If pic.Name = "QR_IMAGE" Then pic.Delete
    Next

    ' ===== CHÈN ?NH QR =====
    Set pic = ActiveSheet.Pictures.Insert(URL)

    With pic
        .Name = "QR_IMAGE"
        .Left = Range("G2").Left
        .Top = Range("G2").Top
        .Width = 400
        .Height = 400
    End With

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("A2:E2")) Is Nothing Then
        TaoQR_ChuyenKhoan
    End If
End Sub

Sub TaoQR_OnlyCode(Optional dummy As Variant = 0)

    Dim txt As String
    txt = Range("A2").value
    If txt = "" Then Exit Sub

    Dim URL As String
    Dim filePath As String

    URL = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" & _
          WorksheetFunction.EncodeURL(txt)

    filePath = Environ$("TEMP") & "\qr_only.png"

    URLDownloadToFile 0, URL, filePath, 0, 0

    InsertQR filePath

End Sub


Sub TaoQR_VietQR(Optional dummy As Variant = 0)

    Dim bank As String, acc As String, nameAcc As String
    Dim amount As String, note As String

    bank = "MB"
    acc = "0904800002"
    nameAcc = "Mai Van Hieu"
    amount = Range("D2").value
    note = Range("E2").value

    If bank = "" Or acc = "" Then Exit Sub

    Dim URL As String
    Dim filePath As String

    URL = "https://img.vietqr.io/image/" & _
          bank & "-" & acc & "-print.png" & _
          "?amount=" & amount & _
          "&addInfo=" & WorksheetFunction.EncodeURL(note) & _
          "&accountName=" & WorksheetFunction.EncodeURL(nameAcc)

    filePath = Environ$("TEMP") & "\qr_bank.png"

    URLDownloadToFile 0, URL, filePath, 0, 0

    InsertQR filePath

End Sub


