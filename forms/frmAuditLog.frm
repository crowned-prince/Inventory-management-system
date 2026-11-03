VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAuditLog 
   Caption         =   "Audit Log"
   ClientHeight    =   9615.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "frmAuditLog.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAuditLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmAuditLog - browse, filter and export the audit log
' ============================================================

Private mLoading As Boolean

' ---------- Startup ----------
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    mLoading = True
    BuildLayout
    StyleControls
    LoadLookups
    mLoading = False

    SetRange Date - 6, Date            ' start with the last 7 days
    Exit Sub

ErrorHandler:
    mLoading = False
    MsgBox "The audit log screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Audit log"
End Sub

Private Function Txt(ByVal ctlName As String) As MSForms.TextBox
    Set Txt = Me.Controls(ctlName)
End Function

Private Sub SetStatus(ByVal message As String)
    Me.Controls("lblStatusMsg").caption = message
End Sub

' ---------- Build the parts made by code ----------
Private Sub BuildLayout()
    Dim headings As Variant, widths As Variant
    Dim i As Long
    Dim X As Single

    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 830, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "AUDIT LOG", -1, ThemeCyan(), 14, True

    modUI.AddLabel Me, "lblFrom", 20, 48, 100, 14, "From (dd-mm-yyyy)", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblTo", 130, 48, 100, 14, "To (dd-mm-yyyy)", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblUser", 240, 48, 145, 14, "User", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblAction", 395, 48, 145, 14, "Action", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblSearch", 550, 48, 240, 14, "Search (record, description)", -1, ThemeGrey(), 8

    modUI.AddTextBox Me, "txtFromDate", 20, 63, 100, 22
    modUI.AddTextBox Me, "txtToDate", 130, 63, 100, 22

    ' Column headings (widths match the list columns set in RefreshList)
    headings = Array("DATE", "TIME", "USER", "ACTION", "RECORD", "DESCRIPTION")
    widths = Array(75, 45, 85, 120, 90, 330)
    X = 23
    For i = 0 To UBound(headings)
        modUI.AddLabel Me, "lblHead" & i, X, 128, widths(i), 13, headings(i), -1, ThemeCyan(), 7, True
        X = X + widths(i)
    Next i

    modUI.AddLabel Me, "lblSummary", 20, 428, 770, 18, "", -1, ThemeWhite(), 9, True
    modUI.AddLabel Me, "lblStatusMsg", 20, 448, 770, 18, "", -1, ThemeGrey(), 9
End Sub

Private Sub StyleControls()
    With cmbUser
        .Left = 240: .Top = 63: .Width = 145: .Height = 22
        .style = fmStyleDropDownList
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
    End With

    With cmbAction
        .Left = 395: .Top = 63: .Width = 145: .Height = 22
        .style = fmStyleDropDownList
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
    End With

    With txtSearch
        .Left = 550: .Top = 63: .Width = 240: .Height = 22
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = ThemeCyan()
    End With

    StyleButton btnToday, "Today", 20, 94, 70, 26, ThemePanel(), ThemeWhite()
    StyleButton btnWeek, "7 days", 96, 94, 70, 26, ThemePanel(), ThemeWhite()
    StyleButton btnMonth, "Month", 172, 94, 70, 26, ThemePanel(), ThemeWhite()
    StyleButton btnAll, "All time", 248, 94, 70, 26, ThemePanel(), ThemeWhite()
    StyleButton btnApply, "SHOW LOGS", 340, 94, 150, 26, ThemeGreen(), ThemeNavy()
    StyleButton btnExport, "EXPORT TO REPORTS SHEET", 500, 94, 190, 26, ThemeCyan(), ThemeNavy()
    StyleButton btnClose, "CLOSE", 700, 94, 90, 26, ThemePanel(), ThemeWhite()

    With lstAudit
        .Left = 20: .Top = 142: .Width = 770: .Height = 280
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal topPos As Single, _
                        ByVal widthVal As Single, ByVal heightVal As Single, _
                        ByVal bgColor As Long, ByVal fgColor As Long)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = topPos: .Width = widthVal: .Height = heightVal
        .BackColor = bgColor: .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = 9: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub LoadLookups()
    modAudit.FillDistinctCombo cmbUser, "Username", "All users"
    cmbUser.ListIndex = 0
    modAudit.FillDistinctCombo cmbAction, "Action", "All actions"
    cmbAction.ListIndex = 0
End Sub

' ---------- Quick ranges ----------
Private Sub SetRange(ByVal fromD As Date, ByVal toD As Date)
    Txt("txtFromDate").text = Format$(fromD, "dd-mm-yyyy")
    Txt("txtToDate").text = Format$(toD, "dd-mm-yyyy")
    RefreshList
End Sub

Private Sub btnToday_Click()
    SetRange Date, Date
End Sub

Private Sub btnWeek_Click()
    SetRange Date - 6, Date
End Sub

Private Sub btnMonth_Click()
    SetRange modProfit.MonthStart(), Date
End Sub

Private Sub btnAll_Click()
    SetRange DateSerial(2000, 1, 1), Date
End Sub

' ---------- Reading the inputs ----------
Private Function ReadInputs(ByRef fromD As Date, ByRef toD As Date) As Boolean
    If Not modValidation.TryParseDate(Txt("txtFromDate").text, fromD) Then
        MsgBox "Please enter the 'From' date as dd-mm-yyyy (for example 01-10-2026).", vbExclamation, "Audit log"
        Exit Function
    End If
    If Not modValidation.TryParseDate(Txt("txtToDate").text, toD) Then
        MsgBox "Please enter the 'To' date as dd-mm-yyyy (for example 31-10-2026).", vbExclamation, "Audit log"
        Exit Function
    End If
    If fromD > toD Then
        MsgBox "The 'From' date cannot be after the 'To' date.", vbExclamation, "Audit log"
        Exit Function
    End If

    ReadInputs = True
End Function

' "" when the first entry ("All ...") is selected, otherwise the selected text.
Private Function FilterValue(ByVal cb As MSForms.ComboBox) As String
    If cb.ListIndex > 0 Then FilterValue = cb.text
End Function

' ---------- Show the matching entries ----------
Private Sub RefreshList()
    Dim fromD As Date, toD As Date
    Dim data As Variant
    Dim formats As Variant
    Dim disp() As String
    Dim total As Long

    If mLoading Then Exit Sub

    On Error GoTo ErrorHandler

    If Not ReadInputs(fromD, toD) Then Exit Sub

    data = modAudit.BuildAuditRows(fromD, toD, FilterValue(cmbUser), FilterValue(cmbAction), _
                                   txtSearch.text, total, 3000)

    lstAudit.Clear
    lstAudit.ColumnCount = 6
    lstAudit.ColumnWidths = "75 pt;45 pt;85 pt;120 pt;90 pt;330 pt"

    If IsEmpty(data) Then
        Me.Controls("lblSummary").caption = "No log entries match these filters."
        SetStatus "Try a wider date range or clear the filters."
        Exit Sub
    End If

    formats = Array("dd-mm-yyyy", "@", "@", "@", "@", "@")
    disp = modReports.ToDisplayArray(data, formats)
    lstAudit.List = disp

    If total > UBound(data, 1) Then
        Me.Controls("lblSummary").caption = "Showing the newest " & UBound(data, 1) & " of " & total & _
                                            " entries. Narrow the filters, or export."
    Else
        Me.Controls("lblSummary").caption = total & " entries"
    End If
    SetStatus "Newest first. Click EXPORT TO REPORTS SHEET to keep a copy in Excel."
    Exit Sub

ErrorHandler:
    MsgBox "The audit log could not be shown." & vbCrLf & Err.description, vbExclamation, "Audit log"
End Sub

Private Sub cmbUser_Change()
    RefreshList
End Sub

Private Sub cmbAction_Change()
    RefreshList
End Sub

Private Sub txtSearch_Change()
    RefreshList
End Sub

Private Sub btnApply_Click()
    RefreshList
End Sub

' ---------- Export ----------
Private Sub btnExport_Click()
    Dim fromD As Date, toD As Date
    Dim exported As Long

    On Error GoTo ErrorHandler

    If Not ReadInputs(fromD, toD) Then Exit Sub

    exported = modAudit.WriteAuditToSheet(fromD, toD, FilterValue(cmbUser), FilterValue(cmbAction), txtSearch.text)
    LogAction "EXPORT REPORT", "AUDIT LOG", exported & " row(s) written to the Reports sheet"

        modSecurity.ShowSheet "Reports"
    MsgBox "The log was written to the Reports sheet (" & exported & " row(s))." & vbCrLf & _
           "Close the application windows to see it.", vbInformation, "Log exported"
    Exit Sub

ErrorHandler:
    MsgBox "The log could not be exported." & vbCrLf & Err.description, vbExclamation, "Export failed"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


