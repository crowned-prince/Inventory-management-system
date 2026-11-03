Attribute VB_Name = "modApp"
Option Explicit

' ============================================================
' modApp - application-wide behaviour: auto-save, window mode, welcome screen
' ============================================================

' ---------- Options (stored on the Settings sheet) ----------
Public Function AutoSaveOn() As Boolean
    AutoSaveOn = (UCase$(Trim$(GetSetting("AutoSave", "Yes"))) = "YES")
End Function

Public Function AppModeOn() As Boolean
    AppModeOn = (UCase$(Trim$(GetSetting("AppMode", "No"))) = "YES")
End Function

' ---------- Auto-save ----------
' Saves the workbook quietly, so a finished sale or purchase is never lost.
Public Sub SaveNow()
    If Not AutoSaveOn() Then Exit Sub
    If Len(ThisWorkbook.Path) = 0 Then Exit Sub        ' never saved yet: a Save As box would appear

    On Error Resume Next
    If Not ThisWorkbook.ReadOnly Then ThisWorkbook.Save
    On Error GoTo 0
End Sub

' ---------- Welcome screen ----------
' Shows the welcome screen with a loading bar for about a second and a half.
Public Sub ShowWelcome()
    Dim startTick As Double
    Dim elapsed As Double
    Const DURATION_SECONDS As Double = 1.4

    On Error GoTo ErrorHandler

    frmWelcome.Show vbModeless
    startTick = Timer

    Do
        elapsed = Timer - startTick
        If elapsed < 0 Then elapsed = elapsed + 86400   ' the clock passed midnight
        frmWelcome.SetProgress elapsed / DURATION_SECONDS
        DoEvents
    Loop While elapsed < DURATION_SECONDS

    Unload frmWelcome
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Unload frmWelcome                                  ' a cosmetic screen must never stop the program
End Sub
