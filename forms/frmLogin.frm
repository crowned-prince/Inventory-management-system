VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Inventory Management System"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17310
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmLogin - login screen
' ============================================================

Private failedAttempts As Long
Private loginSucceeded As Boolean

' ---------- Runs when the form opens ----------
Private Sub UserForm_Initialize()
    ApplyTheme
    failedAttempts = 0
    loginSucceeded = False
    lblMessage.caption = ""
    lblVersion.caption = "System Version: " & GetSetting("SystemVersion", "1.0")
    lblTitle.caption = UCase$(GetSetting("ShopName", "INVENTORY MANAGEMENT SYSTEM"))

    ' Fill in the remembered username if there is one
    txtUsername.text = GetSetting("RememberedUser", "")
    chkRemember.value = (Len(txtUsername.text) > 0)
End Sub

' ---------- Dark navy + cyan theme ----------
Private Sub ApplyTheme()
    Dim navy As Long, panel As Long, cyan As Long, white As Long, grey As Long

    navy = RGB(13, 21, 38)
    panel = RGB(26, 38, 64)
    cyan = RGB(0, 200, 255)
    white = RGB(255, 255, 255)
    grey = RGB(160, 174, 192)

    Me.BackColor = navy

    With lblLogo
        .BackColor = navy: .ForeColor = cyan
        .Font.name = "Segoe UI": .Font.Size = 34: .Font.Bold = True
    End With
    With lblTitle
        .BackColor = navy: .ForeColor = white
        .Font.name = "Segoe UI": .Font.Size = 11: .Font.Bold = True
        .WordWrap = True
    End With
    With lblUser
        .BackColor = navy: .ForeColor = grey
        .Font.name = "Segoe UI": .Font.Size = 9
    End With
    With lblPass
        .BackColor = navy: .ForeColor = grey
        .Font.name = "Segoe UI": .Font.Size = 9
    End With
    With txtUsername
        .BackColor = panel: .ForeColor = white
        .Font.name = "Segoe UI": .Font.Size = 12
    End With
    With txtPassword
        .BackColor = panel: .ForeColor = white
        .Font.name = "Segoe UI": .Font.Size = 12
    End With
    With btnShow
        .BackColor = panel: .ForeColor = cyan
        .Font.name = "Segoe UI": .Font.Size = 8: .Font.Bold = True
    End With
    With chkRemember
        .BackColor = navy: .ForeColor = grey
        .Font.name = "Segoe UI": .Font.Size = 9
    End With
    With lblMessage
        .BackColor = navy: .ForeColor = RGB(255, 90, 90)
        .Font.name = "Segoe UI": .Font.Size = 9
        .WordWrap = True
    End With
    With btnLogin
        .BackColor = cyan: .ForeColor = navy
        .Font.name = "Segoe UI": .Font.Size = 12: .Font.Bold = True
    End With
    With lblVersion
        .BackColor = navy: .ForeColor = grey
        .Font.name = "Segoe UI": .Font.Size = 8
    End With
End Sub

' ---------- Show / hide password ----------
Private Sub btnShow_Click()
    If txtPassword.PasswordChar = "*" Then
        txtPassword.PasswordChar = ""
        btnShow.caption = "HIDE"
    Else
        txtPassword.PasswordChar = "*"
        btnShow.caption = "SHOW"
    End If
    txtPassword.SetFocus
End Sub

' ---------- Login button ----------
Private Sub btnLogin_Click()
    Dim errorText As String

    On Error GoTo ErrorHandler

    errorText = ValidateLogin(txtUsername.text, txtPassword.text)

    If Len(errorText) = 0 Then
        ' Success
        If chkRemember.value Then
            SetSetting "RememberedUser", txtUsername.text
        Else
            SetSetting "RememberedUser", ""
        End If

        LogAction "LOGIN", modMain.CurrentUserID, modMain.CurrentUsername & " logged in"
                loginSucceeded = True
        Unload Me
    Else
        ' Failure
        failedAttempts = failedAttempts + 1
                If Len(Trim$(txtUsername.text)) > 0 And Len(txtPassword.text) > 0 Then
            LogActionAs Trim$(txtUsername.text), "LOGIN FAILED", "-", errorText
        End If
        lblMessage.caption = errorText
        txtPassword.text = ""
        txtPassword.SetFocus

        If failedAttempts >= 3 Then
            MsgBox "Too many failed attempts. The program will close the login screen.", _
                   vbExclamation, "Access Denied"
            Unload Me
        End If
    End If
    Exit Sub

ErrorHandler:
    lblMessage.caption = "Unexpected problem. Please try again."
End Sub

' ---------- Press Enter to log in ----------
Private Sub txtUsername_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then txtPassword.SetFocus
End Sub

Private Sub txtPassword_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then btnLogin_Click
End Sub

