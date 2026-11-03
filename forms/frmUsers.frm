VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUsers 
   Caption         =   "Users and Passwords"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16155
   OleObjectBlob   =   "frmUsers.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmUsers - add users, change roles, disable accounts, reset passwords
' (ADMIN only)
' ============================================================

Private mSelectedID As String      ' "" = new user, otherwise the user being edited
Private mLoading As Boolean

' ---------- Startup ----------
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    mLoading = True
    BuildLayout
    StyleControls
    LoadLookups
    mLoading = False

    RefreshList
    ClearForm
    Exit Sub

ErrorHandler:
    mLoading = False
    MsgBox "The users screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Users"
End Sub

Private Function Txt(ByVal ctlName As String) As MSForms.TextBox
    Set Txt = Me.Controls(ctlName)
End Function

Private Function Cmb(ByVal ctlName As String) As MSForms.ComboBox
    Set Cmb = Me.Controls(ctlName)
End Function

Private Sub SetStatus(ByVal message As String)
    Me.Controls("lblStatusMsg").caption = message
End Sub

' ---------- Build the parts made by code ----------
Private Sub BuildLayout()
    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 830, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "USERS AND PASSWORDS", -1, ThemeCyan(), 14, True

    ' Left column
    AddField "User ID", "txtUserID", "text", 1, 0
    AddField "Username *", "txtUsername", "text", 1, 1
    AddField "Full Name *", "txtFullName", "text", 1, 2
    AddField "Role *", "cmbRole", "combo", 1, 3
    AddField "Status *", "cmbStatus", "combo", 1, 4

    ' Right column
    AddField "Password (new user, or reset)", "txtPassword", "text", 2, 1
    AddField "Confirm Password", "txtConfirm", "text", 2, 2

    Txt("txtPassword").PasswordChar = "*"
    Txt("txtConfirm").PasswordChar = "*"
    Txt("txtUsername").MaxLength = 30
    Txt("txtFullName").MaxLength = 60

    LockBox "txtUserID", True

    With modUI.AddLabel(Me, "lblPwdNote", 270, 192, 230, 40, _
        "When editing, leave both password boxes empty to keep the current password.", _
        -1, ThemeGrey(), 8)
        .WordWrap = True
    End With

    modUI.AddLabel Me, "lblListTitle", 520, 56, 270, 14, "USERS  (click a row to edit)", -1, ThemeCyan(), 8, True

    With modUI.AddLabel(Me, "lblRules", 20, 322, 770, 46, _
        "SAFETY RULES:  You cannot change your own role or disable your own account.  " & _
        "There must always be at least one active ADMIN.  " & _
        "Users are disabled, never deleted, so old invoices and the audit log keep their names.  " & _
        "Use the PASSWORD button on the dashboard to change your own password.", _
        -1, ThemeGrey(), 8)
        .WordWrap = True
    End With

    modUI.AddLabel Me, "lblStatusMsg", 20, 378, 770, 18, "", -1, ThemeGrey(), 9
End Sub

Private Sub AddField(ByVal captionText As String, ByVal ctlName As String, _
                     ByVal kind As String, ByVal columnNo As Long, ByVal rowNo As Long)
    Dim X As Single, Y As Single

    If columnNo = 1 Then X = 20 Else X = 270
    Y = 56 + (rowNo * 44)

    modUI.AddLabel Me, "lbl" & Mid$(ctlName, 4), X, Y, 230, 14, captionText, -1, ThemeGrey(), 8

    If kind = "combo" Then
        modUI.AddCombo Me, ctlName, X, Y + 15, 230, 22
    Else
        modUI.AddTextBox Me, ctlName, X, Y + 15, 230, 22
    End If
End Sub

' Locks or unlocks a text box and changes its look to match.
Private Sub LockBox(ByVal ctlName As String, ByVal locked As Boolean)
    With Txt(ctlName)
        .locked = locked
        If locked Then
            .BackColor = ThemeSidebar()
            .ForeColor = ThemeCyan()
        Else
            .BackColor = ThemePanel()
            .ForeColor = ThemeWhite()
        End If
    End With
End Sub

Private Sub StyleControls()
    With lstUsers
        .Left = 520: .Top = 72: .Width = 270: .Height = 190
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    StyleButton btnSave, "SAVE NEW USER", 20, ThemeGreen(), ThemeNavy()
    StyleButton btnUpdate, "UPDATE", 172, ThemeCyan(), ThemeNavy()
    StyleButton btnClear, "NEW / CLEAR", 324, ThemePanel(), ThemeWhite()
    StyleButton btnClose, "CLOSE", 628, ThemePanel(), ThemeWhite()
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal bgColor As Long, ByVal fgColor As Long)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = 276: .Width = 140: .Height = 36
        .BackColor = bgColor: .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = 10: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub LoadLookups()
    Dim item As Variant

    For Each item In Array("ADMIN", "MANAGER", "STAFF")
        Cmb("cmbRole").AddItem CStr(item)
    Next item

    Cmb("cmbStatus").AddItem "Active"
    Cmb("cmbStatus").AddItem "Disabled"
End Sub

' Selects the combo row with this text (or none).
Private Sub SelectComboText(ByVal cb As MSForms.ComboBox, ByVal searchText As String)
    Dim i As Long

    cb.ListIndex = -1
    For i = 0 To cb.ListCount - 1
        If StrComp(CStr(cb.List(i)), searchText, vbTextCompare) = 0 Then
            cb.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Sub RefreshList()
    mLoading = True
    modUserAdmin.FillUserList lstUsers
    mLoading = False
End Sub

' ---------- New-user mode ----------
Private Sub ClearForm()
    mSelectedID = ""
    mLoading = True

    Txt("txtUserID").text = "AUTO"
    Txt("txtUsername").text = ""
    Txt("txtFullName").text = ""
    Txt("txtPassword").text = ""
    Txt("txtConfirm").text = ""
    SelectComboText Cmb("cmbRole"), "STAFF"
    SelectComboText Cmb("cmbStatus"), "Active"
    lstUsers.ListIndex = -1

    LockBox "txtUsername", False
    LockBox "txtPassword", False
    LockBox "txtConfirm", False
    Cmb("cmbRole").Enabled = True
    Cmb("cmbStatus").Enabled = True

    mLoading = False
    UpdateButtonState
    SetStatus "New user. Fill in the fields marked * and a password, then click SAVE NEW USER."
    Txt("txtUsername").SetFocus
End Sub

Private Sub UpdateButtonState()
    Dim allowed As Boolean, isNew As Boolean

    allowed = modMain.IsAdmin()
    isNew = (Len(mSelectedID) = 0)

    btnSave.Enabled = allowed And isNew
    btnUpdate.Enabled = allowed And Not isNew

    If Not allowed Then SetStatus "View only: only an administrator can manage users."
End Sub

' ---------- Selecting a user ----------
Private Sub lstUsers_Click()
    Dim p As modUserAdmin.TUser

    If mLoading Then Exit Sub
    If lstUsers.ListIndex < 0 Then Exit Sub

    If modUserAdmin.LoadUser(CStr(lstUsers.List(lstUsers.ListIndex, 0)), p) Then
        FillForm p
    End If
End Sub

Private Sub FillForm(ByRef p As modUserAdmin.TUser)
    Dim isSelf As Boolean

    isSelf = (p.userID = modMain.CurrentUserID)

    mLoading = True
    mSelectedID = p.userID

    Txt("txtUserID").text = p.userID
    Txt("txtUsername").text = p.username
    Txt("txtFullName").text = p.FullName
    Txt("txtPassword").text = ""
    Txt("txtConfirm").text = ""
    SelectComboText Cmb("cmbRole"), p.Role
    SelectComboText Cmb("cmbStatus"), p.status

    LockBox "txtUsername", True                       ' the username never changes

    ' You cannot change your own role, status or (here) password
    Cmb("cmbRole").Enabled = Not isSelf
    Cmb("cmbStatus").Enabled = Not isSelf
    LockBox "txtPassword", isSelf
    LockBox "txtConfirm", isSelf

    mLoading = False
    UpdateButtonState

    If isSelf Then
        SetStatus "This is you. You can change your name here. Use the PASSWORD button on the dashboard for your password."
    Else
        SetStatus "Editing " & p.username & ". Fill in the password boxes only if you want to reset the password."
    End If
End Sub

' ---------- Validation: "" = OK, otherwise the message ----------
Private Function ValidateForm() As String
    Dim msg As String
    Dim username As String, password As String, confirm As String
    Dim problem As String
    Dim i As Long
    Dim ch As String
    Dim isNew As Boolean

    isNew = (Len(mSelectedID) = 0)
    username = Trim$(Txt("txtUsername").text)
    password = Txt("txtPassword").text
    confirm = Txt("txtConfirm").text

    If isNew Then
        msg = CheckRequired(username, "username")
        If Len(msg) > 0 Then
            ValidateForm = msg
            Exit Function
        End If

        If Len(username) < 3 Then
            ValidateForm = "The username must have at least 3 characters."
            Exit Function
        End If

        For i = 1 To Len(username)
            ch = Mid$(username, i, 1)
            If Not (ch Like "[A-Za-z0-9_.]") Then
                ValidateForm = "The username can only contain letters, numbers, underscore and dot."
                Exit Function
            End If
        Next i

        If modUserAdmin.UsernameExists(username, "") Then
            ValidateForm = "That username is already taken."
            Exit Function
        End If
    End If

    msg = CheckRequired(Txt("txtFullName").text, "full name")
    If Len(msg) > 0 Then
        ValidateForm = msg
        Exit Function
    End If

    If Cmb("cmbRole").ListIndex < 0 Then
        ValidateForm = "Please select a role."
        Exit Function
    End If
    If Cmb("cmbStatus").ListIndex < 0 Then
        ValidateForm = "Please select a status."
        Exit Function
    End If

    If isNew And Len(password) = 0 Then
        ValidateForm = "Please enter a password for the new user."
        Exit Function
    End If

    If Len(password) > 0 Or Len(confirm) > 0 Then
        If password <> confirm Then
            ValidateForm = "The two passwords do not match."
            Exit Function
        End If
        problem = modUserAdmin.PasswordProblem(password, username)
        If Len(problem) > 0 Then
            ValidateForm = problem
            Exit Function
        End If
    End If
End Function

Private Sub ReadForm(ByRef p As modUserAdmin.TUser)
    p.userID = mSelectedID
    p.username = Trim$(Txt("txtUsername").text)
    p.FullName = Trim$(Txt("txtFullName").text)
    p.Role = UCase$(Cmb("cmbRole").text)
    p.status = Cmb("cmbStatus").text
End Sub

' ---------- Buttons ----------
Private Sub btnSave_Click()
    Dim p As modUserAdmin.TUser
    Dim msg As String, newID As String

    On Error GoTo ErrorHandler

    msg = ValidateForm()
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Please check your input"
        Exit Sub
    End If

    ReadForm p
    newID = modUserAdmin.AddNewUser(p, Txt("txtPassword").text)
    LogAction "ADD USER", newID, "Added " & p.username & " (" & p.Role & ")"

    RefreshList
    ClearForm
    SetStatus "Saved " & p.username & " as " & newID & ". Ready for the next user."
    Exit Sub

ErrorHandler:
    MsgBox "The user could not be saved." & vbCrLf & Err.description, vbExclamation, "Save failed"
End Sub

Private Sub btnUpdate_Click()
    Dim p As modUserAdmin.TUser
    Dim msg As String
    Dim passwordReset As Boolean

    On Error GoTo ErrorHandler

    If Len(mSelectedID) = 0 Then
        MsgBox "Please select a user from the list first.", vbInformation, "Update"
        Exit Sub
    End If

    msg = ValidateForm()
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Please check your input"
        Exit Sub
    End If

    ReadForm p
    passwordReset = (Len(Txt("txtPassword").text) > 0)

    If modUserAdmin.UpdateUser(p, Txt("txtPassword").text) Then
        LogAction "EDIT USER", p.userID, "Updated " & p.username & " (" & p.Role & ", " & p.status & ")"
        If passwordReset Then LogAction "RESET PASSWORD", p.userID, "Password reset for " & p.username

        RefreshList
        ClearForm
        SetStatus "Updated " & p.username & "."
    Else
        MsgBox "That user could not be found.", vbExclamation, "Update failed"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "The user could not be updated." & vbCrLf & Err.description, vbExclamation, "Update failed"
End Sub

Private Sub btnClear_Click()
    ClearForm
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub



