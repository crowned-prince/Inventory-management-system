VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPassword 
   Caption         =   "Change Password"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   OleObjectBlob   =   "frmPassword.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmPassword - change my own password
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    BuildLayout
    StyleButtons
    Exit Sub

ErrorHandler:
    MsgBox "The password screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Password"
End Sub

Private Function Txt(ByVal ctlName As String) As MSForms.TextBox
    Set Txt = Me.Controls(ctlName)
End Function

Private Sub SetStatus(ByVal message As String)
    Me.Controls("lblStatusMsg").caption = message
End Sub

' ---------- Build the parts made by code ----------
Private Sub BuildLayout()
    Dim tb As MSForms.TextBox

    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 380, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 300, 26, "CHANGE PASSWORD", -1, ThemeCyan(), 14, True
    modUI.AddLabel Me, "lblAccount", 30, 46, 300, 14, "Account:  " & modMain.CurrentUsername, -1, ThemeGrey(), 9

    modUI.AddLabel Me, "lblCurrent", 30, 70, 300, 14, "Current password", -1, ThemeGrey(), 8
    Set tb = modUI.AddTextBox(Me, "txtCurrent", 30, 85, 300, 24)
    tb.PasswordChar = "*"

    modUI.AddLabel Me, "lblNew", 30, 120, 300, 14, "New password", -1, ThemeGrey(), 8
    Set tb = modUI.AddTextBox(Me, "txtNew", 30, 135, 300, 24)
    tb.PasswordChar = "*"

    modUI.AddLabel Me, "lblConfirm", 30, 170, 300, 14, "Confirm new password", -1, ThemeGrey(), 8
    Set tb = modUI.AddTextBox(Me, "txtConfirm", 30, 185, 300, 24)
    tb.PasswordChar = "*"

    With modUI.AddLabel(Me, "lblRules", 30, 220, 300, 30, _
        "At least 6 characters. It must not be your username, or your username followed by 123.", _
        -1, ThemeGrey(), 8)
        .WordWrap = True
    End With

    With modUI.AddLabel(Me, "lblStatusMsg", 30, 308, 300, 32, "", -1, ThemeRed(), 9)
        .WordWrap = True
    End With
End Sub
Private Sub StyleButtons()
    StyleButton btnSave, "SAVE", 30, 150, ThemeGreen(), ThemeNavy()
    StyleButton btnCancel, "CANCEL", 190, 140, ThemePanel(), ThemeWhite()
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal widthVal As Single, _
                        ByVal bgColor As Long, ByVal fgColor As Long)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = 260: .Width = widthVal: .Height = 36
        .BackColor = bgColor: .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = 10: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

' ---------- Save ----------
Private Sub btnSave_Click()
    Dim currentPw As String, newPw As String, confirmPw As String
    Dim msg As String

    On Error GoTo ErrorHandler

    currentPw = Txt("txtCurrent").text
    newPw = Txt("txtNew").text
    confirmPw = Txt("txtConfirm").text

    If Len(currentPw) = 0 Then
        SetStatus "Please enter your current password."
        Exit Sub
    End If
    If Len(newPw) = 0 Then
        SetStatus "Please enter a new password."
        Exit Sub
    End If
    If newPw <> confirmPw Then
        SetStatus "The new password and the confirmation do not match."
        Exit Sub
    End If

    msg = modUserAdmin.ChangeOwnPassword(currentPw, newPw)
    If Len(msg) > 0 Then
        SetStatus msg
        Exit Sub
    End If

    LogAction "CHANGE PASSWORD", modMain.CurrentUserID, modMain.CurrentUsername & " changed their password"
    MsgBox "Your password was changed.", vbInformation, "Password"
    Unload Me
    Exit Sub

ErrorHandler:
    SetStatus "The password could not be changed. " & Err.description
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


