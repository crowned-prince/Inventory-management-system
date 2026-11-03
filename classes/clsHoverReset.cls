VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHoverReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class module: removes the hover highlight when the mouse moves over a background area.

Private WithEvents mArea As MSForms.label
Attribute mArea.VB_VarHelpID = -1

Public Sub Attach(ByVal area As MSForms.label)
    Set mArea = area
End Sub

Private Sub mArea_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                            ByVal X As Single, ByVal Y As Single)
    modHover.ClearHover
End Sub

