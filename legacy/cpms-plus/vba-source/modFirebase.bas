Attribute VB_Name = "modFirebase"
Function CheckLicenseFirebase(licenseKey As String) As Boolean

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    Dim URL As String
    URL = "https://firestore.googleapis.com/v1/projects/qsmtoolkit/databases/(default)/documents/licenses/" & licenseKey

    http.Open "GET", URL, False
    http.send

    If http.status <> 200 Then
        CheckLicenseFirebase = False
        Exit Function
    End If

    Dim res As String
    res = http.responseText

    If InStr(res, """booleanValue"": true") > 0 Then
        CheckLicenseFirebase = True
    Else
        CheckLicenseFirebase = False
    End If

End Function

