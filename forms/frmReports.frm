VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReports 
   Caption         =   "Reports"
   ClientHeight    =   9615.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "frmReports.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmReports - choose a report and a date range, generate and export
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

    ApplyDateState
    RunReport
    Exit Sub

ErrorHandler:
    mLoading = False
    MsgBox "The reports screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Reports"
End Sub

Private Function Txt(ByVal ctlName As String) As MSForms.TextBox
    Set Txt = Me.Controls(ctlName)
End Function

Private Sub SetStatus(ByVal message As String)
    Me.Controls("lblStatusMsg").caption = message
End Sub

' ---------- Build the parts made by code ----------
Private Sub BuildLayout()
    Dim i As Long

    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 830, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "REPORTS", -1, ThemeCyan(), 14, True

    modUI.AddLabel Me, "lblReport", 20, 48, 300, 14, "Report", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblFrom", 335, 48, 100, 14, "From (dd-mm-yyyy)", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblTo", 445, 48, 100, 14, "To (dd-mm-yyyy)", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblQuick", 555, 48, 230, 14, "Quick range", -1, ThemeGrey(), 8

    modUI.AddTextBox Me, "txtFromDate", 335, 63, 100, 22
    modUI.AddTextBox Me, "txtToDate", 445, 63, 100, 22
    Txt("txtFromDate").text = Format$(modProfit.MonthStart(), "dd-mm-yyyy")
    Txt("txtToDate").text = Format$(Date, "dd-mm-yyyy")

    modUI.AddLabel Me, "lblReportTitle", 20, 134, 770, 18, "", -1, ThemeCyan(), 11, True

    ' Nine heading labels, shown or hidden depending on the report
    For i = 0 To 8
        modUI.AddLabel Me, "lblHead" & i, 23, 156, 80, 13, "", -1, ThemeCyan(), 7, True
        Me.Controls("lblHead" & i).Visible = False
    Next i

    modUI.AddLabel Me, "lblSummary", 20, 424, 770, 18, "", -1, ThemeWhite(), 9, True
    modUI.AddLabel Me, "lblStatusMsg", 20, 446, 770, 18, "", -1, ThemeGrey(), 9
End Sub

Private Sub StyleControls()
    With cmbReport
        .Left = 20: .Top = 63: .Width = 300: .Height = 22
        .style = fmStyleDropDownList
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
    End With

    StyleButton btnToday, "Today", 555, 62, 54, 24, ThemePanel(), ThemeWhite(), 8
    StyleButton btnWeek, "Week", 615, 62, 54, 24, ThemePanel(), ThemeWhite(), 8
    StyleButton btnMonth, "Month", 675, 62, 54, 24, ThemePanel(), ThemeWhite(), 8
    StyleButton btnYear, "Year", 735, 62, 54, 24, ThemePanel(), ThemeWhite(), 8

    StyleButton btnGenerate, "GENERATE REPORT", 20, 96, 190, 30, ThemeGreen(), ThemeNavy(), 10
    StyleButton btnExport, "EXPORT TO REPORTS SHEET", 220, 96, 220, 30, ThemeCyan(), ThemeNavy(), 10
    StyleButton btnExpenses, "EXPENSES", 450, 96, 130, 30, ThemeActive(), ThemeWhite(), 10
    StyleButton btnClose, "CLOSE", 650, 96, 140, 30, ThemePanel(), ThemeWhite(), 10

    With lstReport
        .Left = 20: .Top = 170: .Width = 770: .Height = 250
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal topPos As Single, _
                        ByVal widthVal As Single, ByVal heightVal As Single, _
                        ByVal bgColor As Long, ByVal fgColor As Long, ByVal fontSize As Single)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = topPos: .Width = widthVal: .Height = heightVal
        .BackColor = bgColor: .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = fontSize: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub LoadLookups()
    Dim i As Long

    For i = 0 To modReports.ReportCount() - 1
        cmbReport.AddItem modReports.ReportNameAt(i)
    Next i
    cmbReport.ListIndex = 0
End Sub

' ---------- Date boxes only matter for some reports ----------
Private Sub ApplyDateState()
    Dim usesDates As Boolean

    If cmbReport.ListIndex < 0 Then Exit Sub
    usesDates = modReports.UsesDateRange(modReports.ReportKeyAt(cmbReport.ListIndex))

    Txt("txtFromDate").Enabled = usesDates
    Txt("txtToDate").Enabled = usesDates
    btnToday.Enabled = usesDates
    btnWeek.Enabled = usesDates
    btnMonth.Enabled = usesDates
    btnYear.Enabled = usesDates
End Sub

Private Sub cmbReport_Change()
    If mLoading Then Exit Sub
    ApplyDateState
    RunReport
End Sub

' ---------- Quick ranges ----------
Private Sub SetRange(ByVal fromD As Date, ByVal toD As Date)
    Txt("txtFromDate").text = Format$(fromD, "dd-mm-yyyy")
    Txt("txtToDate").text = Format$(toD, "dd-mm-yyyy")
    RunReport
End Sub

Private Sub btnToday_Click()
    SetRange Date, Date
End Sub

Private Sub btnWeek_Click()
    SetRange Date - (Weekday(Date, vbMonday) - 1), Date          ' Monday of this week until today
End Sub

Private Sub btnMonth_Click()
    SetRange modProfit.MonthStart(), Date
End Sub

Private Sub btnYear_Click()
    SetRange DateSerial(Year(Date), 1, 1), Date
End Sub

' ---------- Reading and checking the inputs ----------
Private Function ReadInputs(ByRef reportKey As String, ByRef fromD As Date, ByRef toD As Date) As Boolean
    If cmbReport.ListIndex < 0 Then
        MsgBox "Please choose a report.", vbExclamation, "Reports"
        Exit Function
    End If
    reportKey = modReports.ReportKeyAt(cmbReport.ListIndex)

    If modReports.UsesDateRange(reportKey) Then
        If Not modValidation.TryParseDate(Txt("txtFromDate").text, fromD) Then
            MsgBox "Please enter the 'From' date as dd-mm-yyyy (for example 01-10-2026).", vbExclamation, "Reports"
            Exit Function
        End If
        If Not modValidation.TryParseDate(Txt("txtToDate").text, toD) Then
            MsgBox "Please enter the 'To' date as dd-mm-yyyy (for example 31-10-2026).", vbExclamation, "Reports"
            Exit Function
        End If
        If fromD > toD Then
            MsgBox "The 'From' date cannot be after the 'To' date.", vbExclamation, "Reports"
            Exit Function
        End If
    Else
        fromD = Date
        toD = Date
    End If

    ReadInputs = True
End Function

' ---------- Generate and show the report ----------
Private Sub RunReport()
    Dim reportKey As String
    Dim fromD As Date, toD As Date
    Dim title As String, widths As String, summary As String
    Dim headers As Variant, formats As Variant, data As Variant
    Dim disp() As String
    Dim caption As String

    On Error GoTo ErrorHandler

    If Not ReadInputs(reportKey, fromD, toD) Then Exit Sub

    data = modReports.BuildReport(reportKey, fromD, toD, title, headers, formats, widths, summary)

    ShowHeadings headers, widths

    lstReport.Clear
    lstReport.ColumnCount = UBound(headers) + 1
    lstReport.ColumnWidths = widths

    caption = title
    If modReports.UsesDateRange(reportKey) Then
        caption = caption & "     (" & Format$(fromD, "dd-mm-yyyy") & " to " & Format$(toD, "dd-mm-yyyy") & ")"
    Else
        caption = caption & "     (right now)"
    End If
    Me.Controls("lblReportTitle").caption = caption

    If IsEmpty(data) Then
        Me.Controls("lblSummary").caption = "No data for this report and period."
        SetStatus "Nothing to show. Try a wider date range."
    Else
        disp = modReports.ToDisplayArray(data, formats)
        lstReport.List = disp
        Me.Controls("lblSummary").caption = summary
        SetStatus UBound(data, 1) & " row(s). Click EXPORT TO REPORTS SHEET to save a copy in Excel."
    End If
    Exit Sub

ErrorHandler:
    MsgBox "The report could not be generated." & vbCrLf & Err.description, vbExclamation, "Reports"
End Sub

' Positions and fills the column headings to match the list's column widths.
Private Sub ShowHeadings(ByVal headers As Variant, ByVal widths As String)
    Dim parts() As String
    Dim i As Long
    Dim X As Single
    Dim lbl As MSForms.label

    parts = Split(widths, ";")
    X = 23

    For i = 0 To 8
        Set lbl = Me.Controls("lblHead" & i)
        If i <= UBound(headers) And i <= UBound(parts) Then
            lbl.Left = X
            lbl.Width = val(parts(i))
            lbl.caption = UCase$(CStr(headers(i)))
            lbl.Visible = True
            X = X + val(parts(i))
        Else
            lbl.Visible = False
        End If
    Next i
End Sub

Private Sub btnGenerate_Click()
    RunReport
End Sub

' ---------- Export ----------
Private Sub btnExport_Click()
    Dim reportKey As String
    Dim fromD As Date, toD As Date
    Dim exported As Long

    On Error GoTo ErrorHandler

    If Not ReadInputs(reportKey, fromD, toD) Then Exit Sub

    exported = modReports.WriteReportToSheet(reportKey, fromD, toD)
    LogAction "EXPORT REPORT", reportKey, exported & " row(s) written to the Reports sheet"

        modSecurity.ShowSheet "Reports"
    MsgBox "The report was written to the Reports sheet (" & exported & " row(s))." & vbCrLf & _
           "Close the application windows to see it.", vbInformation, "Report exported"
    Exit Sub

ErrorHandler:
    MsgBox "The report could not be exported." & vbCrLf & Err.description, vbExclamation, "Export failed"
End Sub

' ---------- Other buttons ----------
Private Sub btnExpenses_Click()
    frmExpenses.Show
    RunReport                       ' expenses may have changed the profit figures
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub



