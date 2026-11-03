VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWelcome 
   Caption         =   "Welcome"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   OleObjectBlob   =   "frmWelcome.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmWelcome - the short welcome screen shown after login
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    BuildLayout
    Exit Sub

ErrorHandler:
    ' A cosmetic screen must never stop the program
End Sub

Private Sub BuildLayout()
    Dim lbl As MSForms.label

    Me.BackColor = ThemeNavy()

    Set lbl = modUI.AddLabel(Me, "lblLogo", 20, 24, 370, 56, "IMS", -1, ThemeCyan(), 36, True)
    lbl.TextAlign = fmTextAlignCenter

    Set lbl = modUI.AddLabel(Me, "lblShop", 20, 86, 370, 24, _
                             UCase$(GetSetting("ShopName", "Inventory System")), -1, ThemeWhite(), 13, True)
    lbl.TextAlign = fmTextAlignCenter

    Set lbl = modUI.AddLabel(Me, "lblWelcome", 20, 120, 370, 22, _
                             "Welcome, " & modMain.CurrentFullName, -1, ThemeWhite(), 12)
    lbl.TextAlign = fmTextAlignCenter

    Set lbl = modUI.AddLabel(Me, "lblRole", 20, 144, 370, 16, modMain.CurrentRole, -1, ThemeGrey(), 9)
    lbl.TextAlign = fmTextAlignCenter

    modUI.AddLabel Me, "lblTrack", 60, 182, 290, 6, "", ThemePanel()
    modUI.AddLabel Me, "lblBar", 60, 182, 1, 6, "", ThemeCyan()

    Set lbl = modUI.AddLabel(Me, "lblLoading", 20, 196, 370, 16, "Loading your workspace...", -1, ThemeGrey(), 8)
    lbl.TextAlign = fmTextAlignCenter
End Sub

' Called by modApp.ShowWelcome: fraction runs from 0 to 1.
Public Sub SetProgress(ByVal fraction As Double)
    Dim barWidth As Double

    On Error Resume Next
    If fraction < 0 Then fraction = 0
    If fraction > 1 Then fraction = 1

    barWidth = 290 * fraction
    If barWidth < 1 Then barWidth = 1
    Me.Controls("lblBar").Width = barWidth
End Sub

