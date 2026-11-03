Attribute VB_Name = "modSetupAuth"
Option Explicit
' ============================================================
' modSetupAuth - first-time administrator setup (added in IMS v1.1.0)
'
' The set-up state is worked out from the real data in tblUsers, not from one
' flag cell:
'   usable administrator exists        -> Ready (log in as normal)
'   no administrator row, no marker    -> NeedsSetup (first launch)
'   administrator rows exist but none usable, or the "SetupCompleted"
'   marker is set                      -> Inconsistent (never offer set-up again
'                                         automatically; restore or run recovery)
'   Users/Settings table or a column is missing -> DataError
' Any unexpected error also ends in DataError, so the screen that creates an
' administrator is never shown by mistake.
'
' NOTE: this protects against mistakes and corrupted data. It is not a security
' boundary: anyone who can edit the workbook file can edit its sheets.
' ============================================================
Public Const APP_VERSION As String = "1.1.0"
Public Const MIN_PASSWORD_LENGTH As Long = 10
Private Const MAX_PASSWORD_LENGTH As Long = 128
Private Const KEY_COMPLETED As String = "SetupCompleted"
Private Const KEY_COMPLETED_ON As String = "SetupCompletedOn"
Private Const KEY_RECOVERY As String = "SetupRecoveryAuthorized"

Public Enum SetupState
 ssReady = 0
 ssNeedsSetup = 1
 ssInconsistent = 2
 ssDataError = 3
End Enum

Private mBusy As Boolean

' ============================================================
' Start-up gate (called once from modMain.StartApp)
' ============================================================
' Returns True when an administrator exists and the normal login can open.
' Shows the first-time setup screen when it is needed.
Public Function EnsureAdministrator() As Boolean
 Dim state As SetupState
 Dim detail As String
 state = GetSetupState(detail)
 Select Case state
  Case ssReady
   MarkSetupCompleted
   EnsureAdministrator = True
  Case ssNeedsSetup
   frmFirstTimeSetup.Show
   EnsureAdministrator = (GetSetupState(detail) = ssReady)
  Case ssInconsistent
   MsgBox "IMS found user data but no usable administrator account." & vbCrLf & vbCrLf & _
    "For safety, first-time setup is not offered again automatically." & vbCrLf & _
    "Please restore your latest backup, or ask your IMS installer to run the " & _
    "authorized setup recovery.", vbExclamation, "Administrator account missing"
  Case Else
   MsgBox "The IMS user database could not be read." & vbCrLf & vbCrLf & _
    "Please restore your latest backup or contact your IMS installer." & _
    IIf(Len(detail) > 0, vbCrLf & vbCrLf & "Details: " & detail, ""), _
    vbExclamation, "IMS cannot start"
 End Select
End Function

' ============================================================
' Set-up state
' ============================================================
Public Function GetSetupState(ByRef detail As String) As SetupState
 Dim tbl As ListObject
 Dim data As Variant
 Dim colID As Long, colRole As Long, colStatus As Long, colPass As Long
 Dim i As Long
 Dim usable As Long, adminRows As Long
 detail = ""
 On Error GoTo Fail
 If Not CheckDatabase(detail) Then
  GetSetupState = ssDataError
  Exit Function
 End If
 Set tbl = modDatabase.GetTable("Users", "tblUsers")
 If Not tbl.DataBodyRange Is Nothing Then
  data = tbl.DataBodyRange.value
  colID = tbl.ListColumns("UserID").index
  colRole = tbl.ListColumns("Role").index
  colStatus = tbl.ListColumns("Status").index
  colPass = tbl.ListColumns("Password").index
  For i = 1 To UBound(data, 1)
   If Len(CStr(data(i, colID))) > 0 Then
    If UCase$(Trim$(CStr(data(i, colRole)))) = "ADMIN" Then
     adminRows = adminRows + 1
     ' "Usable" mirrors what modUsers.ValidateLogin and modMain.IsAdmin require.
     If CStr(data(i, colRole)) = "ADMIN" And CStr(data(i, colStatus)) = "Active" _
      And Len(CStr(data(i, colPass))) > 0 Then usable = usable + 1
    End If
   End If
  Next i
 End If
 If usable > 0 Then
  GetSetupState = ssReady
 ElseIf SettingIsYes(KEY_RECOVERY) Then
  GetSetupState = ssNeedsSetup               ' authorized recovery
 ElseIf adminRows > 0 Or SettingIsYes(KEY_COMPLETED) Then
  GetSetupState = ssInconsistent
 Else
  GetSetupState = ssNeedsSetup
 End If
 Exit Function
Fail:
 detail = "unexpected data problem"
 GetSetupState = ssDataError
End Function

Private Function SettingIsYes(ByVal settingKey As String) As Boolean
 SettingIsYes = (UCase$(Trim$(GetSetting(settingKey, ""))) = "YES")
End Function

' Records that set-up has happened (also heals installs upgraded from v1.0).
Private Sub MarkSetupCompleted()
 If SettingIsYes(KEY_COMPLETED) Then Exit Sub
 On Error Resume Next
 SetSetting KEY_COMPLETED, "Yes"
 SetSetting KEY_COMPLETED_ON, Format$(Now, "yyyy-mm-dd hh:nn")
 On Error GoTo 0
End Sub

' Checks that everything the set-up needs exists. problem names what is missing.
Private Function CheckDatabase(ByRef problem As String) As Boolean
 Dim tbl As ListObject
 Dim missing As String
 Set tbl = FindTable("Users", "tblUsers")
 If tbl Is Nothing Then
  problem = "the Users table (tblUsers) is missing"
  Exit Function
 End If
 missing = MissingColumns(tbl, Array("UserID", "Username", "Password", "FullName", _
  "Role", "Status", "CreatedDate"))
 If Len(missing) > 0 Then
  problem = "tblUsers is missing column(s): " & missing
  Exit Function
 End If
 Set tbl = FindTable("Settings", "tblSettings")
 If tbl Is Nothing Then
  problem = "the Settings table (tblSettings) is missing"
  Exit Function
 End If
 missing = MissingColumns(tbl, Array("SettingKey", "SettingValue"))
 If Len(missing) > 0 Then
  problem = "tblSettings is missing column(s): " & missing
  Exit Function
 End If
 CheckDatabase = True
End Function

Private Function FindTable(ByVal sheetName As String, ByVal tableName As String) As ListObject
 Dim ws As Worksheet
 Dim lo As ListObject
 On Error Resume Next
 For Each ws In ThisWorkbook.Worksheets
  If StrComp(ws.name, sheetName, vbTextCompare) = 0 Then
   For Each lo In ws.ListObjects
    If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
     Set FindTable = lo
     Exit Function
    End If
   Next lo
  End If
 Next ws
 On Error GoTo 0
End Function

Private Function MissingColumns(ByVal tbl As ListObject, ByVal names As Variant) As String
 Dim i As Long
 Dim lc As ListColumn
 For i = LBound(names) To UBound(names)
  Set lc = Nothing
  On Error Resume Next
  Set lc = tbl.ListColumns(CStr(names(i)))
  On Error GoTo 0
  If lc Is Nothing Then MissingColumns = MissingColumns & CStr(names(i)) & " "
 Next i
 MissingColumns = Trim$(MissingColumns)
End Function

' ============================================================
' Creating the first administrator
' ============================================================
' Validates everything, creates the account, and checks it can really be used.
' Never raises. On failure errorText says why and errorField is
' "username", "password" or "confirm" (the box to put the cursor in).
' If anything goes wrong after the row is written, the row is removed again.
Public Function CreateInitialAdministrator(ByVal username As String, ByVal password As String, _
 ByVal confirm As String, ByRef errorText As String, ByRef errorField As String) As Boolean
 Dim state As SetupState
 Dim detail As String
 Dim p As modUserAdmin.TUser
 Dim errCode As Long
 errorText = ""
 errorField = ""
 If mBusy Then
  errorText = "Setup is already running. Please wait a moment."
  Exit Function
 End If
 mBusy = True
 On Error GoTo Fail
 username = Trim$(username)
 ' 1. The set-up must still be allowed (blocks duplicates and repeated attempts)
 state = GetSetupState(detail)
 If state <> ssNeedsSetup Then
  If state = ssReady Then
   errorText = "An administrator account already exists. First-time setup is no longer available."
  Else
   errorText = "First-time setup is not available. Please restore a backup or contact your IMS installer."
  End If
  GoTo Done
 End If
 ' 2. Validate every field
 errorText = ValidateSetupInput(username, password, confirm, errorField)
 If Len(errorText) > 0 Then GoTo Done
 ' 3. Create the account through the existing user-management code
 p.username = username
 p.FullName = "Administrator"
 p.Role = "ADMIN"
 p.status = "Active"
 modUserAdmin.AddNewUser p, password
 ' 4. Prove the account is usable before telling the user it worked
 If Not VerifyCreatedAdmin(p.userID, password) Then
  Err.Raise vbObjectError + 610, "modSetupAuth", "The new account failed verification."
 End If
 FinishSetup p.userID, username
 CreateInitialAdministrator = True
 GoTo Done
Fail:
 errCode = Err.Number
 On Error Resume Next
 If Len(p.userID) > 0 Then RemoveUserRow p.userID
 LogActionAs "system", "FIRST SETUP FAILED", "-", "Error " & CStr(errCode)
 On Error GoTo 0
 errorText = "The administrator account could not be created, and nothing was saved." & _
  vbCrLf & "Please try again. If it keeps happening, contact your IMS installer."
 errorField = ""
Done:
 mBusy = False
End Function

' Returns "" when the input is acceptable, otherwise the message to show.
Public Function ValidateSetupInput(ByVal username As String, ByVal password As String, _
 ByVal confirm As String, ByRef errorField As String) As String
 Dim msg As String
 errorField = ""
 msg = UsernameProblem(username)
 If Len(msg) > 0 Then
  errorField = "username"
  ValidateSetupInput = msg
  Exit Function
 End If
 If Len(password) = 0 Then
  errorField = "password"
  ValidateSetupInput = "Please enter a password."
  Exit Function
 End If
 msg = PasswordPolicyProblem(password, username)
 If Len(msg) > 0 Then
  errorField = "password"
  ValidateSetupInput = msg
  Exit Function
 End If
 If Len(confirm) = 0 Then
  errorField = "confirm"
  ValidateSetupInput = "Please confirm the password."
  Exit Function
 End If
 If StrComp(password, confirm, vbBinaryCompare) <> 0 Then
  errorField = "confirm"
  ValidateSetupInput = "The two passwords do not match."
 End If
End Function

Public Function UsernameProblem(ByVal username As String) As String
 Dim i As Long
 username = Trim$(username)
 If Len(username) = 0 Then
  UsernameProblem = "Please enter an administrator username."
  Exit Function
 End If
 If Len(username) < 3 Then
  UsernameProblem = "The username must have at least 3 characters."
  Exit Function
 End If
 If Len(username) > 30 Then
  UsernameProblem = "The username cannot be longer than 30 characters."
  Exit Function
 End If
 For i = 1 To Len(username)
  If Not (Mid$(username, i, 1) Like "[A-Za-z0-9_.]") Then
   UsernameProblem = "The username can only contain letters, numbers, underscore and dot."
   Exit Function
  End If
 Next i
 If modUserAdmin.UsernameExists(username, "") Then
  UsernameProblem = "That username is already taken. Please choose another."
 End If
End Function

' ---------- Password rules (stricter than the 6-character rule used for staff) ----------
' At least 10 characters, three of the four character types (or 16+ characters),
' no username inside it, not a very common password, not too repetitive.
Public Function PasswordPolicyProblem(ByVal password As String, ByVal username As String) As String
 username = Trim$(username)
 If Len(password) < MIN_PASSWORD_LENGTH Then
  PasswordPolicyProblem = "The password must have at least " & MIN_PASSWORD_LENGTH & " characters."
  Exit Function
 End If
 If Len(password) > MAX_PASSWORD_LENGTH Then
  PasswordPolicyProblem = "The password cannot be longer than " & MAX_PASSWORD_LENGTH & " characters."
  Exit Function
 End If
 If Len(username) >= 3 Then
  If InStr(1, password, username, vbTextCompare) > 0 Then
   PasswordPolicyProblem = "The password must not contain the username."
   Exit Function
  End If
 End If
 If IsCommonPassword(password) Then
  PasswordPolicyProblem = "That password is too common and easy to guess. Please choose another."
  Exit Function
 End If
 If DistinctCharCount(password) < 5 Then
  PasswordPolicyProblem = "The password is too repetitive. Please use a more varied password."
  Exit Function
 End If
 If CharClassCount(password) < 3 And Len(password) < 16 Then
  PasswordPolicyProblem = "Use at least three of: lowercase letters, uppercase letters, " & _
   "numbers and symbols (or a passphrase of 16 or more characters)."
 End If
End Function

' 0 = empty, 1 = Weak (does not meet the rules), 2 = Fair, 3 = Good, 4 = Strong
Public Function PasswordStrength(ByVal password As String, ByVal username As String) As Long
 Dim classes As Long
 If Len(password) = 0 Then Exit Function
 If Len(PasswordPolicyProblem(password, username)) > 0 Then
  PasswordStrength = 1
  Exit Function
 End If
 classes = CharClassCount(password)
 PasswordStrength = 2
 If Len(password) >= 12 Then PasswordStrength = 3
 If (Len(password) >= 14 And classes >= 3) Or Len(password) >= 20 Then PasswordStrength = 4
End Function

Public Function StrengthCaption(ByVal level As Long) As String
 Select Case level
  Case 1: StrengthCaption = "Weak"
  Case 2: StrengthCaption = "Fair"
  Case 3: StrengthCaption = "Good"
  Case 4: StrengthCaption = "Strong"
  Case Else: StrengthCaption = ""
 End Select
End Function

Private Function CharClassCount(ByVal s As String) As Long
 Dim i As Long
 Dim ch As String
 Dim hasLower As Boolean, hasUpper As Boolean, hasDigit As Boolean, hasOther As Boolean
 For i = 1 To Len(s)
  ch = Mid$(s, i, 1)
  If ch Like "[a-z]" Then
   hasLower = True
  ElseIf ch Like "[A-Z]" Then
   hasUpper = True
  ElseIf ch Like "[0-9]" Then
   hasDigit = True
  Else
   hasOther = True
  End If
 Next i
 CharClassCount = -(hasLower) - (hasUpper) - (hasDigit) - (hasOther)
End Function

Private Function DistinctCharCount(ByVal s As String) As Long
 Dim i As Long
 Dim seen As String
 For i = 1 To Len(s)
  If InStr(1, seen, Mid$(s, i, 1), vbBinaryCompare) = 0 Then seen = seen & Mid$(s, i, 1)
 Next i
 DistinctCharCount = Len(seen)
End Function

Private Function IsCommonPassword(ByVal password As String) As Boolean
 Dim common As Variant
 Dim item As Variant
 Dim p As String
 p = LCase$(password)
 common = Array("password", "password1", "password12", "password123", "password1234", _
  "passw0rd123", "p@ssw0rd123", "p@ssword123", "1234567890", "12345678910", "0123456789", _
  "123456789a", "qwertyuiop", "qwerty1234", "qwerty12345", "qwerty123456", "1q2w3e4r5t", _
  "qazwsxedc1", "zxcvbnm123", "abcdefghij", "abc1234567", "abcd123456", "iloveyou12", _
  "letmein123", "welcome123", "welcome1234", "changeme12", "changeme123", "administrator", _
  "administrator1", "admin12345", "admin123456", "inventory1", "inventory12", "inventory123", _
  "ims1234567", "ims12345678", "11111111111", "1111111111", "0000000000")
 For Each item In common
  If p = CStr(item) Then
   IsCommonPassword = True
   Exit Function
  End If
 Next item
End Function

' ============================================================
' Helpers for account creation
' ============================================================
' Reads the new account back and confirms it is an active ADMIN whose password
' verifies with the secure (non-legacy) hash.
Private Function VerifyCreatedAdmin(ByVal userID As String, ByVal password As String) As Boolean
 Dim tbl As ListObject
 Dim rowNo As Long
 Dim stored As String
 Set tbl = modDatabase.GetTable("Users", "tblUsers")
 rowNo = modUserAdmin.FindUserRow(userID)
 If rowNo = 0 Then Exit Function
 If CStr(modDatabase.GetCellValue(tbl, rowNo, "Role")) <> "ADMIN" Then Exit Function
 If CStr(modDatabase.GetCellValue(tbl, rowNo, "Status")) <> "Active" Then Exit Function
 stored = CStr(modDatabase.GetCellValue(tbl, rowNo, "Password"))
 If modPasswordHash.IsLegacyHash(stored) Then Exit Function
 VerifyCreatedAdmin = modPasswordHash.VerifyPassword(password, stored)
End Function

' Removes a half-created account (clears the row if it is the table's only row).
Private Sub RemoveUserRow(ByVal userID As String)
 Dim tbl As ListObject
 Dim rowNo As Long
 Set tbl = modDatabase.GetTable("Users", "tblUsers")
 rowNo = modUserAdmin.FindUserRow(userID)
 If rowNo = 0 Then Exit Sub
 If tbl.ListRows.count = 1 Then
  tbl.ListRows(1).Range.ClearContents
 Else
  tbl.ListRows(rowNo).Delete
 End If
End Sub

' Non-critical bookkeeping after a successful creation. Errors here must not undo the account.
Private Sub FinishSetup(ByVal userID As String, ByVal username As String)
 On Error Resume Next
 SetSetting KEY_COMPLETED, "Yes"
 SetSetting KEY_COMPLETED_ON, Format$(Now, "yyyy-mm-dd hh:nn")
 SetSetting KEY_RECOVERY, "No"
 SetSetting "RememberedUser", username      ' pre-fills the login screen
 LogActionAs "system", "FIRST SETUP", userID, "Initial administrator account created"
 modApp.SaveNow
 On Error GoTo 0
End Sub

' ============================================================
' Authorized recovery (run from the VBA editor: click inside, press F5)
' ============================================================
' For an installer or developer who has opened the VBA editor (lock the VBA
' project with a password to make that a real gate). It does NOT create or
' change any account. It only allows the first-time setup screen to appear once
' at the next start, and only if no usable administrator exists then.
' If old administrator rows exist, choose a username that is not already used.
Public Sub AuthorizedSetupRecovery()
 Dim answer As String
 Dim state As SetupState
 Dim detail As String
 On Error GoTo Fail
 answer = Trim$(InputBox("AUTHORIZED SETUP RECOVERY" & vbCrLf & vbCrLf & _
  "This lets the first-time setup screen appear once at the next start, " & _
  "but only if no usable administrator account exists." & vbCrLf & vbCrLf & _
  "Type RESET SETUP (capital letters) to continue:", "IMS setup recovery"))
 If answer <> "RESET SETUP" Then
  MsgBox "Nothing was changed.", vbInformation, "IMS setup recovery"
  Exit Sub
 End If
 modSecurity.EnterAppMode
 state = GetSetupState(detail)
 If state = ssReady Then
  MsgBox "A usable administrator account already exists. Nothing was changed." & vbCrLf & _
   "(To reset a forgotten password use EmergencyResetAdmin or restore a backup.)", _
   vbInformation, "IMS setup recovery"
  Exit Sub
 End If
 If state = ssDataError Then
  MsgBox "The user database is damaged, so setup cannot be enabled." & vbCrLf & _
   "Restore a backup first." & IIf(Len(detail) > 0, vbCrLf & detail, ""), _
   vbExclamation, "IMS setup recovery"
  Exit Sub
 End If
 SetSetting KEY_RECOVERY, "Yes"
 LogActionAs "system", "SETUP RECOVERY", "-", "First-time setup re-enabled by an authorized installer"
 MsgBox "Done. The first-time setup screen will appear the next time IMS starts." & vbCrLf & _
  "Save and reopen the workbook.", vbInformation, "IMS setup recovery"
 Exit Sub
Fail:
 MsgBox "The recovery could not be completed.", vbExclamation, "IMS setup recovery"
End Sub
