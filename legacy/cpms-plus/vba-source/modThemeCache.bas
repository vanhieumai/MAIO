Attribute VB_Name = "modThemeCache"
Option Explicit

'===========================================================
' modCache — In-memory caches cho Theme Engine
' ThemeInfo | Colors | Images
'
' CAP NHAT (gop resources.ini + ComponentManifest.ini): cache
' "Resources" cu (m_ResourceCache, doc tu resources.ini) da bi XOA.
' Du lieu resource (ten anh + kich thuoc) gio nam chung trong
' ComponentManifest.ini CUA TUNG THEME va duoc modResourceManager tu
' cache/invalidate rieng (xem InvalidateManifestCache). ClearThemeCache/
' ClearAllCache goi qua ham do de dam bao dong bo, khong con giu cache
' resource o day nua.
'===========================================================

Private m_ThemeInfoCache As Object   ' themeName -> Dict(key->value)
Private m_ColorCache     As Object   ' themeName -> Dict(colorKey->Long)
Private m_ImageCache     As Object   ' absPath   -> Object (StdPicture)
Private m_Ready          As Boolean

' -- Khoi tao ---------------------------------------------
Public Sub InitCache()
    Set m_ThemeInfoCache = MakeDict()
    Set m_ColorCache = MakeDict()
    Set m_ImageCache = MakeDict()
    m_Ready = True
End Sub

Public Sub EnsureCache()
    If Not m_Ready Then InitCache
End Sub

' -- Xoa cache --------------------------------------------
Public Sub ClearAllCache()
    EnsureCache
    m_ThemeInfoCache.RemoveAll
    m_ColorCache.RemoveAll
    m_ImageCache.RemoveAll
    InvalidateManifestCache   ' modResourceManager -- xoa cache manifest cua TAT CA theme
End Sub

Public Sub ClearThemeCache(ByVal themeName As String)
    EnsureCache
    Dim k As String: k = LCase$(themeName)
    If m_ThemeInfoCache.Exists(k) Then m_ThemeInfoCache.Remove k
    If m_ColorCache.Exists(k) Then m_ColorCache.Remove k
    InvalidateManifestCache themeName   ' modResourceManager -- chi xoa cache cua theme nay
    ' ImageCache khong xoa theo theme (dung path lam key, co the chia se giua cac theme)
End Sub

' -- ThemeInfo cache ---------------------------------------
Public Function GetCachedThemeInfo(ByVal themeName As String) As Object
    EnsureCache
    Dim k As String: k = LCase$(themeName)
    If m_ThemeInfoCache.Exists(k) Then Set GetCachedThemeInfo = m_ThemeInfoCache(k)
End Function

Public Sub SetCachedThemeInfo(ByVal themeName As String, ByVal info As Object)
    EnsureCache
    Dim k As String: k = LCase$(themeName)
    If m_ThemeInfoCache.Exists(k) Then m_ThemeInfoCache.Remove k
    m_ThemeInfoCache.Add k, info
End Sub

' -- Color cache -------------------------------------------
Public Function GetCachedColors(ByVal themeName As String) As Object
    EnsureCache
    Dim k As String: k = LCase$(themeName)
    If m_ColorCache.Exists(k) Then Set GetCachedColors = m_ColorCache(k)
End Function

Public Sub SetCachedColors(ByVal themeName As String, ByVal Colors As Object)
    EnsureCache
    Dim k As String: k = LCase$(themeName)
    If m_ColorCache.Exists(k) Then m_ColorCache.Remove k
    m_ColorCache.Add k, Colors
End Sub

' -- Image cache -------------------------------------------
Public Function GetCachedImage(ByVal imagePath As String) As Object
    EnsureCache
    Dim k As String: k = LCase$(imagePath)
    If m_ImageCache.Exists(k) Then Set GetCachedImage = m_ImageCache(k)
End Function

Public Sub SetCachedImage(ByVal imagePath As String, ByVal img As Object)
    EnsureCache
    Dim k As String: k = LCase$(imagePath)
    If m_ImageCache.Exists(k) Then m_ImageCache.Remove k
    If Not img Is Nothing Then m_ImageCache.Add k, img
End Sub

Public Function IsCacheReady() As Boolean
    IsCacheReady = m_Ready
End Function

' -- Utility -----------------------------------------------
Private Function MakeDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1   ' case-insensitive
    Set MakeDict = d
End Function

