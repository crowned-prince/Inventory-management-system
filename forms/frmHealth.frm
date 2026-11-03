VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHealth 
   Caption         =   "Health Check"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960.001
   OleObjectBlob   =   "frmHealth.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmHealth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmHealth - runs the health check and shows the result
' ============================================================

Private mResults As Variant
Private mOnlyIssues As Boolean

' ---------- Startup ----------
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    BuildLayout
    StyleControls
    UpdateButtonState
    RunNow
    Exit Sub

ErrorHandler:
    MsgBox "The health check screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Health check"
End Sub

' ---------- Build the parts made by code ----------
Private Sub BuildLayout()
    Dim headings As Variant, widths As Variant
    Dim i As Long
    Dim X As Single

    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 830, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "HEALTH CHECK", -1, ThemeCyan(), 14, True

    modUI.AddLabel Me, "lblBanner", 20, 48, 770, 28, "", ThemePanel(), ThemeWhite(), 12, True

    headings = Array("STATUS", "AREA", "CHECK", "DETAILS")
    widths = Array(70, 85, 215, 380)
    X = 23
    For i = 0 To UBound(headings)
        modUI.AddLabel Me, "lblHead" & i, X, 84, widths(i), 13, headings(i), -1, ThemeCyan(), 7, True
        X = X + widths(i)
    Next i

    With modUI.AddLabel(Me, "lblDetail", 20, 384, 770, 58, "", ThemePanel(), ThemeWhite(), 9)
        .WordWrap = True
    End With
End Sub

Private Sub StyleControls()
    With lstHealth
        .Left = 20: .Top = 98: .Width = 770: .Height = 280
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    StyleButton btnRun, "RUN AGAIN", 20, 120, ThemeGreen(), ThemeNavy()
    StyleButton btnFilter, "SHOW ISSUES ONLY", 150, 150, ThemePanel(), ThemeWhite()
    StyleButton btnFixCounters, "FIX COUNTERS", 310, 130, ThemeActive(), ThemeWhite()
    StyleButton btnExport, "EXPORT TO REPORTS SHEET", 450, 190, ThemeCyan(), ThemeNavy()
    StyleButton btnClose, "CLOSE", 650, 140, ThemePanel(), ThemeWhite()
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal widthVal As Single, _
                        ByVal bgColor As Long, ByVal fgColor As Long)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = 450: .Width = widthVal: .Height = 30
        .BackColor = bgColor: .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = 9: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub UpdateButtonState()
    btnFixCounters.Enabled = modMain.IsAdmin()
End Sub

Private Sub SetDetail(ByVal message As String)
    Me.Controls("lblDetail").caption = message
End Sub

' ---------- Running the check ----------
Private Sub RunNow()
    Dim errText As String

    On Error GoTo ErrorHandler

    Me.MousePointer = fmMousePointerHourGlass
    mResults = modHealth.RunHealthCheck()
    Me.MousePointer = fmMousePointerDefault

    ShowBanner
    FillList
    SetDetail "Click a line to see the details and how to fix it."
    Exit Sub

ErrorHandler:
    errText = Err.description
    Me.MousePointer = fmMousePointerDefault
    MsgBox "The health check could not finish." & vbCrLf & errText, vbExclamation, "Health check"
End Sub

Private Sub ShowBanner()
    Dim okCount As Long, infoCount As Long, warnCount As Long, problemCount As Long
    Dim banner As MSForms.label

    modHealth.CountStatuses mResults, okCount, infoCount, warnCount, problemCount
    Set banner = Me.Controls("lblBanner")

    If problemCount > 0 Then
        banner.ForeColor = ThemeRed()
        banner.caption = "  " & problemCount & " PROBLEM(S) and " & warnCount & " warning(s) found   (" & okCount & " checks passed)"
    ElseIf warnCount > 0 Then
        banner.ForeColor = ThemeOrange()
        banner.caption = "  " & warnCount & " warning(s) to look at   (" & okCount & " checks passed)"
    Else
        banner.ForeColor = ThemeGreen()
        banner.caption = "  HEALTHY: no problems or warnings   (" & okCount & " checks passed)"
    End If
End Sub

' ---------- The list ----------
Private Sub FillList()
    Dim i As Long, n As Long
    Dim status As String

    lstHealth.Clear
    lstHealth.ColumnCount = 5
    lstHealth.ColumnWidths = "70 pt;85 pt;215 pt;380 pt;0 pt"      ' the last column holds the result number

    If IsEmpty(mResults) Then Exit Sub

    For i = 1 To UBound(mResults, 1)
        status = CStr(mResults(i, 1))
        If (Not mOnlyIssues) Or status = "WARNING" Or status = "PROBLEM" Then
            lstHealth.AddItem status
            n = lstHealth.ListCount - 1
            lstHealth.List(n, 1) = CStr(mResults(i, 2))
            lstHealth.List(n, 2) = CStr(mResults(i, 3))
            lstHealth.List(n, 3) = CStr(mResults(i, 4))
            lstHealth.List(n, 4) = CStr(i)
        End If
    Next i
End Sub

Private Sub lstHealth_Click()
    Dim idx As Long
    Dim fixText As String
    Dim status As String
    Dim detail As MSForms.label

    If lstHealth.ListIndex < 0 Then Exit Sub

    idx = CLng(lstHealth.List(lstHealth.ListIndex, 4))
    status = CStr(mResults(idx, 1))
    fixText = CStr(mResults(idx, 5))
    If Len(fixText) = 0 Then fixText = "Nothing to do."

    Set detail = Me.Controls("lblDetail")
    detail.caption = status & ":  " & CStr(mResults(idx, 3)) & vbCrLf & _
                     "Details:  " & CStr(mResults(idx, 4)) & vbCrLf & _
                     "How to fix:  " & fixText

    Select Case status
        Case "PROBLEM": detail.ForeColor = ThemeRed()
        Case "WARNING": detail.ForeColor = ThemeOrange()
        Case "OK": detail.ForeColor = ThemeGreen()
        Case Else: detail.ForeColor = ThemeWhite()
    End Select
End Sub

' ---------- Buttons ----------
Private Sub btnRun_Click()
    RunNow
End Sub

Private Sub btnFilter_Click()
    mOnlyIssues = Not mOnlyIssues
    If mOnlyIssues Then
        btnFilter.caption = "SHOW ALL"
    Else
        btnFilter.caption = "SHOW ISSUES ONLY"
    End If
    FillList
End Sub

Private Sub btnFixCounters_Click()
    Dim fixedCount As Long

    On Error GoTo ErrorHandler

    If MsgBox("Raise any ID counter that is behind the highest ID in use?" & vbCrLf & _
              "This is safe: it only stops new IDs from repeating old ones.", _
              vbYesNo + vbQuestion, "Fix counters") <> vbYes Then Exit Sub

    fixedCount = modHealth.FixCounters()
    LogAction "FIX COUNTERS", "-", fixedCount & " counter(s) corrected"
    RunNow
    MsgBox fixedCount & " counter(s) were corrected.", vbInformation, "Fix counters"
    Exit Sub

ErrorHandler:
    MsgBox "The counters could not be fixed." & vbCrLf & Err.description, vbExclamation, "Fix counters"
End Sub

Private Sub btnExport_Click()
    Dim written As Long

    On Error GoTo ErrorHandler

    written = modHealth.WriteHealthToSheet(mResults)
    LogAction "EXPORT REPORT", "HEALTH CHECK", written & " line(s) written to the Reports sheet"

    modSecurity.ShowSheet "Reports"

    MsgBox "The health check was written to the Reports sheet (" & written & " lines)." & vbCrLf & _
           "Close the application windows to see it.", vbInformation, "Health check exported"
    Exit Sub

ErrorHandler:
    MsgBox "The health check could not be exported." & vbCrLf & Err.description, vbExclamation, "Export failed"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


