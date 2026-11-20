Attribute VB_Name = "clsProgressChartLabel"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'==========================================================
' CLASS MODULE: clsProgressChartLabel
' PURPOSE:
'   Catch DblClick events from labels created dynamically in
'   frmContractSearch.fraProgressChart.
'
' IMPORT:
'   Add this file as a Class Module named:
'       clsProgressChartLabel
'==========================================================

Private WithEvents mLabel As MSForms.Label
Attribute mLabel.VB_VarHelpID = -1
Private mOwner As Object
Private mContractID As String
Private mActionName As String

Public Sub Initialize(ByVal owner As Object, _
                      ByVal labelControl As MSForms.Label, _
                      ByVal contractID As String, _
                      ByVal actionName As String)

    Set mOwner = owner
    Set mLabel = labelControl
    mContractID = Trim$(contractID)
    mActionName = actionName

    ' Tag is useful for inspection/debugging and keeps the relationship
    ' visible on the dynamic control itself.
    mLabel.Tag = mContractID
    mLabel.ControlTipText = "Double-click to open " & mContractID

End Sub

Private Sub mLabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    On Error GoTo EXIT_HANDLER

    If Not mOwner Is Nothing Then
        mOwner.ProgressChartLabelDblClick mContractID, mActionName
    End If

EXIT_HANDLER:

End Sub


