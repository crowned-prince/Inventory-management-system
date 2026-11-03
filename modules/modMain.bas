Attribute VB_Name = "modMain"
Option Explicit

' ============================================================
' modMain - start-up and "who is logged in" information
' ============================================================

Public CurrentUserID As String
Public CurrentUsername As String
Public CurrentFullName As String
Public CurrentRole As String          ' ADMIN, MANAGER or STAFF

Public LogoutRequested As Boolean     ' True when the user clicked Logout
Public PendingProductID As String     ' product chosen in Search, to be opened for editing

' Start-up: sheets -> login -> welcome -> password check -> backup -> alert -> dashboard.
' Repeats if the user logs out. Saves when the dashboard closes.
Public Sub StartApp()
 Dim previousState As Long
 previousState = Application.WindowState
 On Error GoTo ErrorHandler
 modSecurity.EnterAppMode ' unprotect the sheets, hide or show them as set
 modDash.SetUnlocked False ' cover the analytics until an admin or manager logs in
 If modApp.AppModeOn() Then Application.WindowState = xlMinimized
 If Not modSetupAuth.EnsureAdministrator() Then GoTo Cleanup ' <-- ADDED
 Do
 ClearSession
 LogoutRequested = False
 frmLogin.Show
 If Not IsLoggedIn() Then Exit Do
 modApp.ShowWelcome
 If modSecurity.HasDefaultPassword() Then modSecurity.PromptPasswordChange
 modBackup.AutoBackupIfDue
 modAlerts.ShowLoginAlert
 frmDashboard.Show
 modApp.SaveNow
 Loop While LogoutRequested
Cleanup:
 modDash.AfterSession ' fresh analytics for administrators and managers
 Application.WindowState = previousState
 Exit Sub
ErrorHandler:
 MsgBox "The application stopped unexpectedly." & vbCrLf & Err.description, _
 vbExclamation, "Inventory System"
 Resume Cleanup
End Sub

' Forgets the current user (used by Logout and at start-up).
Public Sub ClearSession()
    CurrentUserID = ""
    CurrentUsername = ""
    CurrentFullName = ""
    CurrentRole = ""
End Sub

Public Function IsLoggedIn() As Boolean
    IsLoggedIn = (Len(CurrentUserID) > 0)
End Function

' True for ADMIN. Used to block staff and managers from admin-only features.
Public Function IsAdmin() As Boolean
    IsAdmin = (CurrentRole = "ADMIN")
End Function

