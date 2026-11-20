Attribute VB_Name = "frmLoading"
Attribute VB_Base = "0{E3B55652-FAFE-4920-A522-1B552FEC20A6}{9535AB0D-8CAF-4E0E-A133-ED8A15E16E60}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'Attribute VB_Name = "frmLoading"
'==========================================================
' frmLoading  — FIXED v3
'
' V?n d? còn l?i sau v2 (xác nh?n qua screenshot):
'
'   1. fraProgress.Caption = "Label1"
'      Frame control gi? caption m?c d?nh c?a VBA designer.
'      Caption này hi?n ? border c?a Frame, che m?t giao di?n.
'      Fix: set Caption = "" trong Initialize.
'
'   2. lblBar vô hình dù Width dã du?c gán dúng.
'      ApplyTheme Me ? ApplyColorsToControls recurses vào Frame
'      ? tìm lblBar (ki?u Label) ? set BackStyle = 0 (transparent).
'      Label transparent không có background fill ? bar bi?n m?t
'      hoàn toàn dù BackColor và Width dúng.
'      Fix: sau ApplyTheme, override lblBar.BackStyle = 1 (Opaque)
'           và set BackColor v? màu Primary c?a theme.
'
'   3. lblPercent n?m trong Frame cung b? BackStyle = 0.
'      Text v?n render du?c (transparent label hi?n text qua
'      ForeColor), nhung n?u ForeColor = Background thì m?t.
'      Fix: d?m b?o ForeColor set rõ ràng sau ApplyTheme.
'
' Nguyên t?c: ApplyTheme ch? du?c g?i M?T L?N trong Initialize.
' M?i property b? ApplyTheme ghi dè sai c?n du?c restore ngay
' sau ApplyTheme trong cùng Initialize — không ph?i trong Show/
' SetProgress vì Initialize ch? ch?y 1 l?n.
'==========================================================

Option Explicit

Private Const BAR_MAX_WIDTH  As Long = 234
Private Const BAR_BACK_COLOR As Long = &HF0F0F0   ' màu n?n bar (xám nh?t)
' Màu fill bar l?y t? theme Primary; n?u theme chua load thì dùng màu cam CPMS
Private Const BAR_FILL_COLOR As Long = &H2DB4F4F  ' #F4732D cam CPMS (BGR order)


'----------------------------------------------------------
' StartLoading
'----------------------------------------------------------
Public Sub StartLoading(ByVal title As String)

    Me.Show vbModeless          ' Initialize ch?y ? dây ? ApplyTheme ? fix restore

    lblTitle.Caption = title
    lblStep.Caption = "Starting..."
    lblPercent.Caption = "0%"
    lblBar.Width = 0

    ForceRepaint

End Sub


'----------------------------------------------------------
' SetProgress
'----------------------------------------------------------
Public Sub SetProgress(ByVal pct As Integer, ByVal StepText As String)

    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100

    lblStep.Caption = StepText
    lblPercent.Caption = pct & "%"
    lblBar.Width = CLng(BAR_MAX_WIDTH * pct / 100)

    ForceRepaint

End Sub



'----------------------------------------------------------
' ForceRepaint — toggle ScreenUpdating, Repaint, DoEvents
'----------------------------------------------------------
Private Sub ForceRepaint()

    Dim wasUpdating As Boolean
    wasUpdating = Application.ScreenUpdating
    If Not wasUpdating Then Application.ScreenUpdating = True
    Me.Repaint
    DoEvents
    If Not wasUpdating Then Application.ScreenUpdating = False

End Sub


'----------------------------------------------------------
' UserForm_Initialize
'
' QUAN TR?NG: ApplyTheme ghi dè BackStyle = 0 (transparent)
' lên lblBar vì nó là Label. C?n restore ngay sau ApplyTheme.
' Initialize ch? ch?y 1 l?n ? dây là noi dúng d? fix.
'----------------------------------------------------------
Private Sub UserForm_Initialize()

    ' 1. Apply theme màu s?c cho form
    ApplyTheme Me

    ' 2. Fix fraProgress: xóa caption "Label1" m?c d?nh c?a designer
    On Error Resume Next
    fraProgress.Caption = ""
    On Error GoTo 0

    ' 3. Fix lblBar: ApplyTheme set BackStyle=0 (transparent) vì lblBar
    '    là Label. Progress bar c?n BackStyle=1 (Opaque) m?i hi?n màu fill.
    On Error Resume Next
    lblBar.BackStyle = 1           ' fmBackStyleOpaque — B?T BU?C d? bar hi?n
    lblBar.BackColor = GetThemeColor("Primary")    ' màu fill = Primary c?a theme
    lblBar.Height = fraProgress.Height         ' kh?p chi?u cao frame (tr? border)
    lblBar.Top = 0
    lblBar.Left = 0
    On Error GoTo 0

    ' 4. Fix lblPercent: d?m b?o text color tuong ph?n v?i n?n
    On Error Resume Next
    lblPercent.ForeColor = GetThemeColor("Text")
    On Error GoTo 0

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    UnregisterForm Me
End Sub

