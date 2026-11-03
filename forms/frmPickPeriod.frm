VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPickPeriod 
   Caption         =   "Choose a Period"
   ClientHeight    =   8820.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8955.001
   OleObjectBlob   =   "frmPickPeriod.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPickPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmPickPeriod - choose the time period the Dashboard sheet should show
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    BuildLayout
    StyleControls
    LoadPeriods
    Exit Sub

ErrorHandler:
    MsgBox "The period list could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Choose a period"
End Sub

Private Function Txt(ByVal ctlName As String) As MSForms.TextBox
    Set Txt = Me.Controls(ctlName)
End Function

Private Sub SetStatus(ByVal message As String)
    Me.Controls("lblStatusMsg").caption = message
End Sub

Private Sub BuildLayout()
    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 480, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 300, 26, "CHOOSE A TIME PERIOD", -1, ThemeCyan(), 14, True

    modUI.AddLabel Me, "lblFrom", 20, 270, 140, 14, "Custom: From (dd-mm-yyyy)", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblTo", 180, 270, 140, 14, "Custom: To (dd-mm-yyyy)", -1, ThemeGrey(), 8
    modUI.AddTextBox Me, "txtFrom", 20, 285, 140, 22
    modUI.AddTextBox Me, "txtTo", 180, 285, 140, 22
    Txt("txtFrom").text = Format$(Date - 29, "dd-mm-yyyy")
    Txt("txtTo").text = Format$(Date, "dd-mm-yyyy")

    modUI.AddLabel Me, "lblHint", 20, 312, 400, 14, "The dates are used only when 'Custom range' is selected.", -1, ThemeGrey(), 8

    With modUI.AddLabel(Me, "lblStatusMsg", 20, 392, 400, 34, "", -1, ThemeRed(), 9)
        .WordWrap = True
    End With
End Sub

Private Sub StyleControls()
    With lstPeriods
        .Left = 20: .Top = 56: .Width = 400: .Height = 204
        .ColumnCount = 2
        .ColumnWidths = "0 pt;380 pt"                    ' the first column holds the hidden key
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    StyleButton btnApply, "SHOW THIS PERIOD", 20, 200, ThemeCyan(), ThemeNavy()
    StyleButton btnCancel, "CLOSE", 240, 180, ThemePanel(), ThemeWhite()
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal widthVal As Single, _
                        ByVal bgColor As Long, ByVal fgColor As Long)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = 340: .Width = widthVal: .Height = 34
        .BackColor = bgColor: .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = 10: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub LoadPeriods()
    Dim i As Long, n As Long
    Dim currentKey As String

    currentKey = modDash.CurrentPeriodKey()
    lstPeriods.Clear

    For i = 1 To modDash.PeriodCount()
        lstPeriods.AddItem modDash.PeriodKeyAt(i - 1)
        n = lstPeriods.ListCount - 1
        lstPeriods.List(n, 1) = modDash.PeriodLabelFor(modDash.PeriodKeyAt(i - 1))
        If modDash.PeriodKeyAt(i - 1) = currentKey Then lstPeriods.ListIndex = n
    Next i
End Sub

Private Sub ApplySelected()
    Dim key As String
    Dim fromD As Date, toD As Date

    If lstPeriods.ListIndex < 0 Then
        SetStatus "Please select a period first."
        Exit Sub
    End If

    key = CStr(lstPeriods.List(lstPeriods.ListIndex, 0))

    If key = "CUSTOM" Then
        If Not modValidation.TryParseDate(Txt("txtFrom").text, fromD) Then
            SetStatus "Please enter the 'From' date as dd-mm-yyyy (for example 01-10-2026)."
            Exit Sub
        End If
        If Not modValidation.TryParseDate(Txt("txtTo").text, toD) Then
            SetStatus "Please enter the 'To' date as dd-mm-yyyy (for example 31-10-2026)."
            Exit Sub
        End If
        If fromD > toD Then
            SetStatus "The 'From' date cannot be after the 'To' date."
            Exit Sub
        End If
    Else
        fromD = Date
        toD = Date
    End If

    modDash.SetPeriodFilter key, fromD, toD
    Unload Me
End Sub

Private Sub btnApply_Click()
    ApplySelected
End Sub

Private Sub lstPeriods_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ApplySelected
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

