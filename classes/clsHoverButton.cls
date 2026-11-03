VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHoverButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class module: gives one CommandButton a hover highlight.

Private WithEvents mButton As MSForms.CommandButton
Attribute mButton.VB_VarHelpID = -1
Private mNormalColor As Long
Private mHoverColor As Long

Public Sub Attach(ByVal btn As MSForms.CommandButton, ByVal normalColor As Long, _
                  ByVal hoverColor As Long)
    Set mButton = btn
    mNormalColor = normalColor
    mHoverColor = hoverColor
End Sub

Private Sub mButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                              ByVal X As Single, ByVal Y As Single)
    modHover.SetHover mButton, mNormalColor, mHoverColor
End Sub

