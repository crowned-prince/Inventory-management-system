Attribute VB_Name = "modSecurity"
Option Explicit

' ============================================================
' modSecurity - data-sheet hiding, read-only view, default-password
' check and the launcher button on the Dashboard sheet.
'
' How the sheets are handled:
'   APPLICATION MODE : sheets are unprotected (the program writes to them)
'                      and hidden if "HideSheets" is Yes on the Settings sheet.
'   VIEW MODE        : sheets are visible but protected (read-only).
' ============================================================

' Password used when protecting sheets for viewing.
' Keep it empty ("") so you can never lock yourself out. If you set one, write it down.
Private Const SHEET_PASSWORD As String = ""

' ---------- The "hide sheets" setting ----------
Public Function HidingIsOn() As Boolean
    HidingIsOn = (UCase$(Trim$(GetSetting("HideSheets", "No"))) = "YES")
End Function

Public Sub SetHiding(ByVal turnOn As Boolean)
    If turnOn Then SetSetting "HideSheets", "Yes" Else SetSetting "HideSheets", "No"
End Sub

' ---------- Application mode ----------
' Unprotects every sheet, then hides the data sheets (if hiding is on)
' or shows them all (if hiding is off). Safe to run at any time.
Public Sub EnterAppMode()
    Dim ws As Worksheet

    On Error Resume Next        ' a sheet protected with someone else's password must not stop the program

    For Each ws In ThisWorkbook.Worksheets
        If ws.ProtectContents Then ws.Unprotect SHEET_PASSWORD
    Next ws

    If HidingIsOn() Then
        ThisWorkbook.Worksheets("Dashboard").Visible = xlSheetVisible
        ThisWorkbook.Worksheets("Dashboard").Activate
        For Each ws In ThisWorkbook.Worksheets
            If ws.name <> "Dashboard" Then ws.Visible = xlSheetVeryHidden
        Next ws
    Else
        For Each ws In ThisWorkbook.Worksheets
            ws.Visible = xlSheetVisible
        Next ws
    End If

    On Error GoTo 0
End Sub

' ---------- View mode ----------
' Shows every sheet and protects them (read-only). Returns how many could NOT be protected.
Public Function EnterViewMode() As Long
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Visible = xlSheetVisible
        Err.Clear
        If ws.name <> "Dashboard" Then
            ws.Protect password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True
            If Err.Number <> 0 Then EnterViewMode = EnterViewMode + 1
        End If
        On Error GoTo 0
    Next ws
End Function

' Makes a sheet visible and active (used after an export so you can find the result).
Public Sub ShowSheet(ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then
        ws.Visible = xlSheetVisible
        ws.Activate
    End If
    On Error GoTo 0
End Sub

' ---------- Default passwords ----------
' True if the logged-in user still has the starting password (their username followed by 123).
Public Function HasDefaultPassword() As Boolean
    Dim tbl As ListObject
    Dim rowNo As Long
    Dim storedHash As String
    Dim loginName As String

    If Len(modMain.CurrentUserID) = 0 Then Exit Function

    Set tbl = modDatabase.GetTable("Users", "tblUsers")
    rowNo = modDatabase.FindRowByValue(tbl, "UserID", modMain.CurrentUserID)
    If rowNo = 0 Then Exit Function

    storedHash = CStr(modDatabase.GetCellValue(tbl, rowNo, "Password"))
    loginName = LCase$(CStr(modDatabase.GetCellValue(tbl, rowNo, "Username")))

    HasDefaultPassword = (storedHash = modUsers.HashPassword(loginName & "123"))
End Function

Public Sub PromptPasswordChange()
    MsgBox "You are still using the default password." & vbCrLf & vbCrLf & _
           "For safety, please choose a new password now.", vbExclamation, "Change your password"
    frmPassword.Show
End Sub

' Builds the analytics page on the Dashboard sheet (see modDashBuild). Run once (F5).
Public Sub SetupDashboardSheet()
    On Error GoTo ErrorHandler

    modDashBuild.BuildDashboardSheet
    MsgBox "The Dashboard sheet is ready.", vbInformation, "Setup"
    Exit Sub

ErrorHandler:
    MsgBox "Could not set up the Dashboard sheet." & vbCrLf & Err.description, vbExclamation, "Setup"
End Sub
