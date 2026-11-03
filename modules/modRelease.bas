Attribute VB_Name = "modRelease"
Option Explicit

' ============================================================
' modRelease - the final steps before hand-in, and emergency tools.
' Run these from the VBA editor: click inside a macro and press F5.
' ============================================================

' Checks the health of the system, switches on the recommended settings,
' hides the data sheets, makes a backup and saves the file.
Public Sub PrepareFinalRelease()
    Dim results As Variant
    Dim okCount As Long, infoCount As Long, warnCount As Long, problemCount As Long
    Dim i As Long, shown As Long
    Dim problemText As String, warningText As String
    Dim backupPath As String

    On Error GoTo ErrorHandler

    If MsgBox("This prepares the workbook for hand-in:" & vbCrLf & vbCrLf & _
              "1.  Runs the health check" & vbCrLf & _
              "2.  Switches on: hide data sheets, auto-save and daily auto-backup" & vbCrLf & _
              "3.  Refreshes the Dashboard sheet" & vbCrLf & _
              "4.  Makes a backup and saves the file" & vbCrLf & vbCrLf & "Continue?", _
              vbYesNo + vbQuestion, "Prepare final release") <> vbYes Then Exit Sub

    modSecurity.EnterAppMode                       ' sheets must be unprotected while we work

    ' ---- 1. Health check ----
    results = modHealth.RunHealthCheck()
    modHealth.CountStatuses results, okCount, infoCount, warnCount, problemCount

    For i = 1 To UBound(results, 1)
        If CStr(results(i, 1)) = "PROBLEM" And shown < 8 Then
            problemText = problemText & "-  " & CStr(results(i, 3)) & ":  " & Left$(CStr(results(i, 4)), 110) & vbCrLf
            shown = shown + 1
        End If
    Next i

    If problemCount > 0 Then
        MsgBox problemCount & " problem(s) must be fixed first:" & vbCrLf & vbCrLf & problemText & vbCrLf & _
               "Open Settings, HEALTH CHECK for details and fixes, then run this again.", _
               vbExclamation, "Prepare final release"
        Exit Sub
    End If

    If warnCount > 0 Then
        shown = 0
        For i = 1 To UBound(results, 1)
            If CStr(results(i, 1)) = "WARNING" And shown < 8 Then
                warningText = warningText & "-  " & CStr(results(i, 3)) & ":  " & Left$(CStr(results(i, 4)), 110) & vbCrLf
                shown = shown + 1
            End If
        Next i

        If MsgBox(warnCount & " warning(s) were found:" & vbCrLf & vbCrLf & warningText & vbCrLf & _
                  "Continue anyway?", vbYesNo + vbExclamation, "Prepare final release") <> vbYes Then Exit Sub
    End If

    ' ---- 2. Recommended settings ----
    SetSetting "HideSheets", "Yes"
    SetSetting "AutoSave", "Yes"
    SetSetting "AutoBackup", "Yes"

    ' ---- 3. Dashboard sheet, then hide the data sheets ----
    modSecurity.SetupDashboardSheet
    modSecurity.EnterAppMode

    ' ---- 4. Log, back up, save ----
    LogActionAs "system", "FINAL RELEASE", "-", "Final release prepared"
    backupPath = modBackup.CreateBackup()

    If Len(ThisWorkbook.Path) > 0 Then
        If Not ThisWorkbook.ReadOnly Then ThisWorkbook.Save
    End If

    MsgBox "The final release is ready." & vbCrLf & vbCrLf & _
           "Health check:  " & problemCount & " problem(s), " & warnCount & " warning(s)" & vbCrLf & _
           "Backup:  " & backupPath & vbCrLf & vbCrLf & _
           "The data sheets are now hidden. Only the Dashboard sheet is visible." & vbCrLf & _
           "Still to do by hand: change the default passwords, and (optionally) lock the VBA project.", _
           vbInformation, "Prepare final release"
    Exit Sub

ErrorHandler:
    MsgBox "The final release could not be completed." & vbCrLf & Err.description, _
           vbExclamation, "Prepare final release"
End Sub

' Emergency: switches hiding OFF and shows and unprotects every sheet.
' Use it if you cannot open the application and need to see your data.
Public Sub ShowAllSheetsNow()
    Dim ws As Worksheet

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.ProtectContents Then ws.Unprotect ""      ' works when no sheet password was set
    Next ws
    On Error GoTo 0

    modSecurity.SetHiding False
    modSecurity.EnterAppMode                              ' shows every sheet when hiding is off

    MsgBox "All sheets are visible and unprotected. Hiding is switched OFF.", vbInformation, "Show all sheets"
End Sub

' Emergency: resets the user named admin to role ADMIN, status Active, password admin123.
' The program asks for a new password at the next login.
Public Sub EmergencyResetAdmin()
    Dim answer As String
    Dim tbl As ListObject
    Dim rowNo As Long
    Dim userID As String

    On Error GoTo ErrorHandler

    answer = Trim$(InputBox("EMERGENCY ADMIN RESET" & vbCrLf & vbCrLf & _
                            "This sets the user named admin back to:" & vbCrLf & _
                            "role ADMIN, status Active, password admin123." & vbCrLf & _
                            "You will be asked for a new password at the next login." & vbCrLf & vbCrLf & _
                            "Type RESET (in capital letters) to continue:", _
                            "Emergency admin reset"))

    If answer <> "RESET" Then
        MsgBox "Nothing was changed.", vbInformation, "Emergency admin reset"
        Exit Sub
    End If

    modSecurity.EnterAppMode

    Set tbl = modDatabase.GetTable("Users", "tblUsers")
    rowNo = modDatabase.FindRowByValue(tbl, "Username", "admin")
    If rowNo = 0 Then
        MsgBox "There is no user named admin.", vbExclamation, "Emergency admin reset"
        Exit Sub
    End If

    userID = CStr(modDatabase.GetCellValue(tbl, rowNo, "UserID"))
    modDatabase.SetCellValue tbl, rowNo, "Role", "ADMIN"
    modDatabase.SetCellValue tbl, rowNo, "Status", "Active"
    modUserAdmin.SetUserPassword userID, "admin123"

    LogActionAs "system", "EMERGENCY RESET", userID, "The admin account was reset"
    MsgBox "The admin account was reset." & vbCrLf & vbCrLf & _
           "Log in as admin with the password admin123. You will then be asked to choose a new one.", _
           vbInformation, "Emergency admin reset"
    Exit Sub

ErrorHandler:
    MsgBox "The reset could not be completed." & vbCrLf & Err.description, vbExclamation, "Emergency admin reset"
End Sub

