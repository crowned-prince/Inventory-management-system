VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFirstTimeSetup 
   Caption         =   "Inventory Management System"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFirstTimeSetup.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmFirstTimeSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' ============================================================
' frmFirstTimeSetup - first-time administrator account (IMS v1.1.0)
' Shown by modSetupAuth.EnsureAdministrator when no administrator exists.
' Like the other IMS screens, every control is built in code (see BuildLayout),
' using the shared theme colours and modUI helpers, so it matches frmLogin.
' All account rules live in modSetupAuth; this form only collects the input.
' ============================================================
Private Const FORM_W As Single = 380
Private Const FORM_H As Single = 410
Private Const MARGIN As Single = 30
Private Const FIELD_W As Single = 320
Private Const TOGGLE_W As Single = 64
Private Const GAP As Single = 6

Private WithEvents txtUsername As MSForms.TextBox
Attribute txtUsername.VB_VarHelpID = -1
Private WithEvents txtPassword As MSForms.TextBox
Attribute txtPassword.VB_VarHelpID = -1
Private WithEvents txtConfirm As MSForms.TextBox
Attribute txtConfirm.VB_VarHelpID = -1
Private WithEvents btnShowPw As MSForms.CommandButton
Attribute btnShowPw.VB_VarHelpID = -1
Private WithEvents btnShowConfirm As MSForms.CommandButton
Attribute btnShowConfirm.VB_VarHelpID = -1
Private WithEvents btnCreate As MSForms.CommandButton
Attribute btnCreate.VB_VarHelpID = -1
Private lblMessage As MSForms.label
Private lblStrength As MSForms.label
Private segBar(1 To 4) As MSForms.label
Private mHoverCreate As clsHoverButton
Private mHoverPw As clsHoverButton
Private mHoverConfirm As clsHoverButton
Private mResetHeader As clsHoverReset
Private mResetFooter As clsHoverReset
Private mBusy As Boolean
Private mCreated As Boolean
Private mInitFailed As Boolean

' ---------- Start-up ----------
Private Sub UserForm_Initialize()
 On Error GoTo ErrorHandler
 BuildLayout
 Exit Sub
ErrorHandler:
 mInitFailed = True
End Sub

Private Sub UserForm_Activate()
 If mInitFailed Then
  MsgBox "The setup screen could not be displayed.", vbExclamation, "IMS Setup"
  Unload Me
  Exit Sub
 End If
 On Error Resume Next
 txtUsername.SetFocus
End Sub

' ---------- Build the screen ----------
Private Sub BuildLayout()
 Dim lbl As MSForms.label
 Dim i As Long
 Dim segW As Single
 Dim boxW As Single
 Dim x2 As Single
 Me.caption = "Inventory Management System"
 Me.StartUpPosition = 2 ' centre on screen
 Me.BackColor = ThemeNavy()
 Me.Width = FORM_W + (Me.Width - Me.InsideWidth)
 Me.Height = FORM_H + (Me.Height - Me.InsideHeight)
 boxW = FIELD_W - TOGGLE_W - GAP
 x2 = MARGIN + FIELD_W - TOGGLE_W

 ' ----- Header -----
 modUI.AddLabel Me, "lblHeaderBar", 0, 0, FORM_W, 66, "", ThemeSidebar()
 modUI.AddLabel Me, "lblHeaderLine", 0, 66, FORM_W, 2, "", ThemeActive()
 modUI.AddLabel Me, "lblLogo", MARGIN, 12, 64, 42, "IMS", -1, ThemeCyan(), 26, True
 modUI.AddLabel Me, "lblDivider", 100, 15, 1, 36, "", ThemePanel()
 modUI.AddLabel Me, "lblBrand", 112, 15, 250, 18, "INVENTORY MANAGEMENT SYSTEM", -1, ThemeWhite(), 10, True
 modUI.AddLabel Me, "lblSubtitle", 112, 35, 250, 14, "FIRST-TIME SETUP", -1, ThemeCyan(), 8, True

 ' ----- Welcome -----
 modUI.AddLabel Me, "lblWelcome", MARGIN, 80, FIELD_W, 24, "Welcome to IMS", -1, ThemeWhite(), 16, True
 Set lbl = modUI.AddLabel(Me, "lblIntro", MARGIN, 106, FIELD_W, 30, _
  "Create the administrator account for this installation. You will use it to sign in " & _
  "and to manage other users. IMS has no default password, so choose one you will remember.", _
  -1, ThemeGrey(), 8)
 lbl.WordWrap = True

 ' ----- Username -----
 modUI.AddLabel Me, "lblUser", MARGIN, 144, FIELD_W, 12, "ADMINISTRATOR USERNAME", -1, ThemeGrey(), 7.5, True
 Set txtUsername = modUI.AddTextBox(Me, "txtUsername", MARGIN, 158, FIELD_W, 26)
 txtUsername.MaxLength = 30
 txtUsername.Font.Size = 11

 ' ----- Password + strength -----
 modUI.AddLabel Me, "lblPass", MARGIN, 192, FIELD_W, 12, "PASSWORD", -1, ThemeGrey(), 7.5, True
 Set txtPassword = modUI.AddTextBox(Me, "txtPassword", MARGIN, 206, boxW, 26)
 txtPassword.PasswordChar = "*"
 txtPassword.MaxLength = 128
 txtPassword.Font.Size = 11
 Set btnShowPw = AddToggle("btnShowPw", x2, 206)
 segW = (boxW - (3 * 4)) / 4
 For i = 1 To 4
  Set segBar(i) = modUI.AddLabel(Me, "lblSeg" & i, MARGIN + ((i - 1) * (segW + 4)), 240, segW, 4, "", ThemePanel())
 Next i
 Set lblStrength = modUI.AddLabel(Me, "lblStrength", x2, 236, TOGGLE_W, 12, "", -1, ThemeGrey(), 8, True)
 lblStrength.TextAlign = fmTextAlignRight

 ' ----- Confirm -----
 modUI.AddLabel Me, "lblConfirm", MARGIN, 256, FIELD_W, 12, "CONFIRM PASSWORD", -1, ThemeGrey(), 7.5, True
 Set txtConfirm = modUI.AddTextBox(Me, "txtConfirm", MARGIN, 270, boxW, 26)
 txtConfirm.PasswordChar = "*"
 txtConfirm.MaxLength = 128
 txtConfirm.Font.Size = 11
 Set btnShowConfirm = AddToggle("btnShowConfirm", x2, 270)

 ' ----- Rules and messages -----
 Set lbl = modUI.AddLabel(Me, "lblRules", MARGIN, 302, FIELD_W, 24, _
  "At least " & modSetupAuth.MIN_PASSWORD_LENGTH & " characters with three of: lowercase, uppercase, " & _
  "numbers, symbols (or a passphrase of 16+ characters). It must not contain the username.", _
  -1, ThemeGrey(), 7.5)
 lbl.WordWrap = True
 Set lblMessage = modUI.AddLabel(Me, "lblMessage", MARGIN, 328, FIELD_W, 26, "", -1, ThemeRed(), 8.5)
 lblMessage.WordWrap = True

 ' ----- Create button and footer -----
 Set btnCreate = Me.Controls.Add("Forms.CommandButton.1", "btnCreate", True)
 With btnCreate
  .caption = "CREATE ADMINISTRATOR"
  .Left = MARGIN: .Top = 358: .Width = FIELD_W: .Height = 34
  .BackColor = ThemeCyan(): .ForeColor = ThemeNavy()
  .Font.name = "Segoe UI": .Font.Size = 11: .Font.Bold = True
  .TakeFocusOnClick = False
 End With
 modUI.AddLabel Me, "lblVersion", MARGIN, 396, 120, 12, "IMS v" & modSetupAuth.APP_VERSION, -1, ThemeGrey(), 7.5
 Set lbl = modUI.AddLabel(Me, "lblStored", 150, 396, 200, 12, "Passwords are stored as salted hashes.", -1, ThemeGrey(), 7.5)
 lbl.TextAlign = fmTextAlignRight

 ' ----- Hover behaviour (same helpers as the dashboard) -----
 Set mHoverCreate = New clsHoverButton
 mHoverCreate.Attach btnCreate, ThemeCyan(), ThemeHover()
 Set mHoverPw = New clsHoverButton
 mHoverPw.Attach btnShowPw, ThemePanel(), ThemeActive()
 Set mHoverConfirm = New clsHoverButton
 mHoverConfirm.Attach btnShowConfirm, ThemePanel(), ThemeActive()
 Set mResetHeader = New clsHoverReset
 mResetHeader.Attach Me.Controls("lblHeaderBar")
 Set mResetFooter = New clsHoverReset
 mResetFooter.Attach Me.Controls("lblVersion")
End Sub

Private Function AddToggle(ByVal ctlName As String, ByVal leftPos As Single, _
 ByVal topPos As Single) As MSForms.CommandButton
 Dim btn As MSForms.CommandButton
 Set btn = Me.Controls.Add("Forms.CommandButton.1", ctlName, True)
 With btn
  .caption = "SHOW"
  .Left = leftPos: .Top = topPos: .Width = TOGGLE_W: .Height = 26
  .BackColor = ThemePanel(): .ForeColor = ThemeCyan()
  .Font.name = "Segoe UI": .Font.Size = 8: .Font.Bold = True
  .TakeFocusOnClick = False
  .TabStop = False
 End With
 Set AddToggle = btn
End Function

' ---------- Show / hide ----------
Private Sub btnShowPw_Click()
 ToggleVisible txtPassword, btnShowPw
End Sub

Private Sub btnShowConfirm_Click()
 ToggleVisible txtConfirm, btnShowConfirm
End Sub

Private Sub ToggleVisible(ByVal box As MSForms.TextBox, ByVal btn As MSForms.CommandButton)
 If box.PasswordChar = "*" Then
  box.PasswordChar = ""
  btn.caption = "HIDE"
 Else
  box.PasswordChar = "*"
  btn.caption = "SHOW"
 End If
 box.SetFocus
End Sub

' ---------- Live feedback ----------
Private Sub txtUsername_Change()
 lblMessage.caption = ""
 UpdateStrength
End Sub

Private Sub txtPassword_Change()
 lblMessage.caption = ""
 UpdateStrength
End Sub

Private Sub txtConfirm_Change()
 lblMessage.caption = ""
End Sub

Private Sub UpdateStrength()
 Dim level As Long
 Dim col As Long
 Dim i As Long
 On Error Resume Next
 level = modSetupAuth.PasswordStrength(txtPassword.text, Trim$(txtUsername.text))
 Select Case level
  Case 1: col = ThemeRed()
  Case 2: col = ThemeOrange()
  Case 3: col = ThemeCyan()
  Case 4: col = ThemeGreen()
  Case Else: col = ThemeGrey()
 End Select
 For i = 1 To 4
  If i <= level Then segBar(i).BackColor = col Else segBar(i).BackColor = ThemePanel()
 Next i
 lblStrength.caption = modSetupAuth.StrengthCaption(level)
 lblStrength.ForeColor = col
End Sub

' ---------- Focus highlight ----------
Private Sub txtUsername_Enter(): txtUsername.BorderColor = ThemeCyan(): End Sub
Private Sub txtUsername_Exit(ByVal Cancel As MSForms.ReturnBoolean): txtUsername.BorderColor = RGB(60, 80, 120): End Sub
Private Sub txtPassword_Enter(): txtPassword.BorderColor = ThemeCyan(): End Sub
Private Sub txtPassword_Exit(ByVal Cancel As MSForms.ReturnBoolean): txtPassword.BorderColor = RGB(60, 80, 120): End Sub
Private Sub txtConfirm_Enter(): txtConfirm.BorderColor = ThemeCyan(): End Sub
Private Sub txtConfirm_Exit(ByVal Cancel As MSForms.ReturnBoolean): txtConfirm.BorderColor = RGB(60, 80, 120): End Sub

' ---------- Keyboard: Enter moves on, and creates from the last box ----------
Private Sub txtUsername_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 If KeyCode = vbKeyReturn Then txtPassword.SetFocus
End Sub

Private Sub txtPassword_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 If KeyCode = vbKeyReturn Then txtConfirm.SetFocus
End Sub

Private Sub txtConfirm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
 If KeyCode = vbKeyReturn Then btnCreate_Click
End Sub

' ---------- Clear a stuck hover highlight ----------
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
 ByVal X As Single, ByVal Y As Single)
 modHover.ClearHover
End Sub

Private Sub txtUsername_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
 ByVal X As Single, ByVal Y As Single)
 modHover.ClearHover
End Sub

Private Sub txtPassword_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
 ByVal X As Single, ByVal Y As Single)
 modHover.ClearHover
End Sub

Private Sub txtConfirm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
 ByVal X As Single, ByVal Y As Single)
 modHover.ClearHover
End Sub

' ---------- Create ----------
Private Sub btnCreate_Click()
 Dim errText As String
 Dim errField As String
 If mBusy Then Exit Sub
 mBusy = True
 On Error GoTo ErrorHandler
 modHover.ClearHover
 lblMessage.caption = ""
 btnCreate.Enabled = False
 btnCreate.caption = "CREATING..."
 Me.Repaint
 If modSetupAuth.CreateInitialAdministrator(txtUsername.text, txtPassword.text, _
  txtConfirm.text, errText, errField) Then
  mCreated = True
  txtPassword.text = ""
  txtConfirm.text = ""
  MsgBox "Your administrator account was created." & vbCrLf & vbCrLf & _
   "Please sign in with the username and password you just chose.", _
   vbInformation, "IMS Setup"
  Unload Me
  Exit Sub
 End If
 lblMessage.caption = errText
 Select Case errField
  Case "username": txtUsername.SetFocus
  Case "password": txtPassword.SetFocus
  Case "confirm": txtConfirm.SetFocus
 End Select
 btnCreate.caption = "CREATE ADMINISTRATOR"
 btnCreate.Enabled = True
 mBusy = False
 Exit Sub
ErrorHandler:
 On Error Resume Next
 lblMessage.caption = "Something went wrong. Nothing was saved. Please try again."
 btnCreate.caption = "CREATE ADMINISTRATOR"
 btnCreate.Enabled = True
 mBusy = False
End Sub

' ---------- Closing with the X ----------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 If mCreated Or mInitFailed Then Exit Sub
 If CloseMode = vbFormControlMenu Then
  If MsgBox("Exit without creating the administrator account?" & vbCrLf & _
   "IMS cannot be used until one is created.", vbYesNo + vbQuestion, "IMS Setup") <> vbYes Then
   Cancel = True
  End If
 End If
End Sub


