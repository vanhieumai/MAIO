Attribute VB_Name = "modAssets"
Option Explicit

'==========================================================
' CPMS Asset Manager
' Download Background & Icons from Internet
'==========================================================

#If VBA7 Then

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
Alias "URLDownloadToFileA" ( _
    ByVal pCaller As LongPtr, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As LongPtr) As Long

#Else

Private Declare Function URLDownloadToFile Lib "urlmon" _
Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

#End If

'==========================================================
' CONFIG
'==========================================================

Public Const APP_FOLDER As String = "\CPMS\Assets\"

'Ð?i thành website c?a b?n
Public Const ASSET_URL As String = "https://yourdomain.com/assets/"

'==========================================================
' LOCAL PATH
'==========================================================

Public Function AssetFolder() As String

    AssetFolder = Environ$("LOCALAPPDATA") & APP_FOLDER

End Function

'==========================================================
' CREATE LOCAL FOLDER
'==========================================================

Public Sub CreateAssetFolder()

    Dim root As String
    Dim folder As String

    root = Environ$("LOCALAPPDATA") & "\CPMS"
    folder = AssetFolder

    If Dir(root, vbDirectory) = "" Then
        MkDir root
    End If

    If Dir(folder, vbDirectory) = "" Then
        MkDir folder
    End If

End Sub

'==========================================================
' DOWNLOAD ONE FILE
'==========================================================

Public Function DownloadFile(ByVal URL As String, _
                             ByVal SaveFile As String) As Boolean

    DownloadFile = _
        (URLDownloadToFile(0, URL, SaveFile, 0, 0) = 0)

End Function

'==========================================================
' DOWNLOAD IF FILE NOT EXISTS
'==========================================================

Public Sub DownloadAsset(ByVal fileName As String, _
                         Optional ByVal folder As String = "")

    Dim URL As String
    Dim SavePath As String

    If folder <> "" Then

        URL = ASSET_URL & folder & "/" & fileName

    Else

        URL = ASSET_URL & fileName

    End If

    SavePath = AssetFolder & fileName

    If Dir(SavePath) = "" Then

        DownloadFile URL, SavePath

    End If

End Sub

'==========================================================
' DOWNLOAD ALL ICONS
'==========================================================

Public Sub DownloadIcons()

    Dim arr
    Dim i As Long

    arr = Array( _
        "Create.png", _
        "Rename.png", _
        "Delete.png", _
        "Clone.png", _
        "Import.png", _
        "Export.png", _
        "Open.png", _
        "AskAI.png")

    For i = LBound(arr) To UBound(arr)

        DownloadAsset arr(i), "icons"

    Next i

End Sub

'==========================================================
' DOWNLOAD BACKGROUNDS
'==========================================================

Public Sub DownloadBackgrounds()

    DownloadAsset "Home.jpg", "backgrounds"

    DownloadAsset "AskAI.jpg", "backgrounds"

End Sub

'==========================================================
' DOWNLOAD EVERYTHING
'==========================================================

Public Sub CheckAssets()

    CreateAssetFolder

    DownloadIcons

    DownloadBackgrounds

End Sub

