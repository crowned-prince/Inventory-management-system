VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdmin 
   Caption         =   "Administration"
   ClientHeight    =   9015.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9165.001
   OleObjectBlob   =   "frmAdmin.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmAdmin - the administration centre (ADMIN only)
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    BuildLayout
    StyleButtons
    UpdateButtonState
    Exit Sub

ErrorHandler:
    MsgBox "The administration screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Administration"
End Sub

Private Sub BuildLayout()
    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 500, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 300, 26, "ADMINISTRATION", -1, ThemeCyan(), 14, True
    modUI.AddLabel Me, "lblWho", 30, 46, 400, 14, "Signed in as " & modMain.CurrentFullName & _
                   "  (" & modMain.CurrentRole & ")", -1, ThemeGrey(), 8
End Sub

Private Sub StyleButtons()
    StyleHubButton btnAuditLog, "AUDIT LOG", "See who did what, and when", 64
    StyleHubButton btnBackup, "BACKUP DATABASE", "Save a safe copy of all your data", 118
    StyleHubButton btnUsers, "USERS AND PASSWORDS", "Add staff, change roles, reset passwords", 172
    StyleHubButton btnSheets, "DATABASE SHEETS", "View or hide the data sheets", 226
    StyleHubButton btnShopSettings, "SHOP SETTINGS", "Shop name, address, tax and more", 280
    StyleHubButton btnHealth, "HEALTH CHECK", "Scan the whole system for problems", 334

    With btnClose
        .caption = "CLOSE"
        .Left = 30: .Top = 394: .Width = 400: .Height = 34
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub StyleHubButton(ByVal btn As MSForms.CommandButton, ByVal title As String, _
                           ByVal description As String, ByVal topPos As Single)
    With btn
        .caption = title & vbCrLf & description
        .Left = 30: .Top = topPos: .Width = 400: .Height = 48
        .BackColor = ThemePanel()
        .ForeColor = ThemeCyan()
        .Font.name = "Segoe UI": .Font.Size = 10: .Font.Bold = True
        .WordWrap = True
        .TakeFocusOnClick = False
    End With
End Sub

' Only an ADMIN may use these tools.
Private Sub UpdateButtonState()
    Dim allowed As Boolean
    allowed = modMain.IsAdmin()

    btnAuditLog.Enabled = allowed
    btnBackup.Enabled = allowed
    btnUsers.Enabled = allowed
    btnSheets.Enabled = allowed
    btnShopSettings.Enabled = allowed
    btnHealth.Enabled = allowed
End Sub

Private Sub btnAuditLog_Click()
    frmAuditLog.Show
End Sub

Private Sub btnBackup_Click()
    frmBackup.Show
End Sub

Private Sub btnUsers_Click()
    frmUsers.Show
End Sub

Private Sub btnSheets_Click()
    Dim hidingOn As Boolean
    Dim choice As VbMsgBoxResult
    Dim failed As Long

    hidingOn = modSecurity.HidingIsOn()

    choice = MsgBox("DATABASE SHEETS" & vbCrLf & vbCrLf & _
                    "Hiding the data sheets is currently " & IIf(hidingOn, "ON", "OFF") & "." & vbCrLf & vbCrLf & _
                    "Yes  =  show all sheets now (read-only)" & vbCrLf & _
                    "No  =  turn hiding " & IIf(hidingOn, "OFF", "ON") & vbCrLf & _
                    "Cancel  =  do nothing", _
                    vbYesNoCancel + vbQuestion, "Database sheets")

    If choice = vbYes Then
        failed = modSecurity.EnterViewMode()
        LogAction "VIEW SHEETS", "-", "Database sheets shown (read-only)"
        MsgBox "All sheets are now visible and protected (read-only)." & vbCrLf & vbCrLf & _
               "Close the Dashboard window (the X) to see them." & vbCrLf & _
               "To go back, click OPEN APPLICATION on the Dashboard sheet.", _
               vbInformation, "Database sheets"
        If failed > 0 Then
            MsgBox failed & " sheet(s) could not be protected (they may already have another password).", _
                   vbExclamation, "Database sheets"
        End If

    ElseIf choice = vbNo Then
        modSecurity.SetHiding Not hidingOn
        modSecurity.EnterAppMode
        LogAction "SHEET HIDING", "-", "Hiding the data sheets was switched " & IIf(hidingOn, "OFF", "ON")
        MsgBox "Hiding is now " & IIf(hidingOn, "OFF", "ON") & ".", vbInformation, "Database sheets"
    End If
End Sub

Private Sub btnShopSettings_Click()
    frmSettings.Show
End Sub

Private Sub btnHealth_Click()
    frmHealth.Show
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


