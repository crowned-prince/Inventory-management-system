VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExpenses 
   Caption         =   "Expenses"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "frmExpenses.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmExpenses - record business expenses and see this month's profit
' ============================================================

Private mSelectedID As String      ' "" = new expense, otherwise the one being edited
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
    RefreshMonthPanel
    ClearForm
    Exit Sub

ErrorHandler:
    mLoading = False
    MsgBox "The expenses screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Expenses"
End Sub

Private Function Txt(ByVal ctlName As String) As MSForms.TextBox
    Set Txt = Me.Controls(ctlName)
End Function

Private Function Cmb(ByVal ctlName As String) As MSForms.ComboBox
    Set Cmb = Me.Controls(ctlName)
End Function

Private Function cur() As String
    cur = GetSetting("CurrencySymbol", "Rs.") & " "
End Function

Private Sub SetStatus(ByVal message As String)
    Me.Controls("lblStatusMsg").caption = message
End Sub

' ---------- Build the parts made by code ----------
Private Sub BuildLayout()
    Dim headings As Variant, widths As Variant
    Dim monthCaptions As Variant, monthNames As Variant
    Dim i As Long
    Dim X As Single, Y As Single

    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 830, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "EXPENSES", -1, ThemeCyan(), 14, True

    ' Left column
    AddField "Expense ID", "txtExpenseID", "text", 1, 0
    AddField "Date (dd-mm-yyyy) *", "txtExpenseDate", "text", 1, 1
    AddField "Category *", "cmbCategory", "combo", 1, 2
    AddField "Amount *", "txtAmount", "text", 1, 3

    ' Right column: description (tall box)
    AddField "Description", "txtDescription", "text", 2, 0
    With Txt("txtDescription")
        .Height = 66
        .Multiline = True
        .WordWrap = True
        .EnterKeyBehavior = True
        .MaxLength = 200
    End With

    With Txt("txtExpenseID")
        .locked = True
        .BackColor = ThemeSidebar()
        .ForeColor = ThemeCyan()
    End With

    ' "This month" profit panel on the right
    modUI.AddLabel Me, "lblMonthPanel", 520, 52, 270, 160, "", ThemePanel()
    modUI.AddLabel Me, "lblMonthTitle", 530, 58, 250, 14, "THIS MONTH  (profit so far)", -1, ThemeCyan(), 8, True

    monthCaptions = Array("Revenue", "Cost of goods", "Gross profit", "Expenses", "NET PROFIT")
    monthNames = Array("lblMonthRevenue", "lblMonthCost", "lblMonthGross", "lblMonthExpenses", "lblMonthNet")
    For i = 0 To 4
        Y = 80 + i * 26
        modUI.AddLabel Me, "lblMonthCap" & i, 530, Y, 130, 18, monthCaptions(i), -1, ThemeGrey(), 9
        With modUI.AddLabel(Me, CStr(monthNames(i)), 650, Y, 130, 18, "", -1, ThemeWhite(), 10, True)
            .TextAlign = fmTextAlignRight
        End With
    Next i

    ' Filter and list headings
    modUI.AddLabel Me, "lblFilterTitle", 20, 282, 400, 14, "EXPENSE LIST  (click a row to edit, or type to filter)", -1, ThemeCyan(), 8, True

    headings = Array("ID", "DATE", "CATEGORY", "DESCRIPTION", "AMOUNT")
    widths = Array(70, 80, 120, 330, 110)
    X = 23
    For i = 0 To UBound(headings)
        modUI.AddLabel Me, "lblHead" & i, X, 325, widths(i), 13, headings(i), -1, ThemeCyan(), 7, True
        X = X + widths(i)
    Next i

    modUI.AddLabel Me, "lblStatusMsg", 20, 462, 770, 18, "", -1, ThemeGrey(), 9
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

Private Sub StyleControls()
    With txtFilter
        .Left = 20: .Top = 297: .Width = 400: .Height = 22
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = ThemeCyan()
    End With

    With lstExpenses
        .Left = 20: .Top = 339: .Width = 770: .Height = 118
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    StyleButton btnSave, "SAVE", 20, ThemeGreen(), ThemeNavy()
    StyleButton btnUpdate, "UPDATE", 172, ThemeCyan(), ThemeNavy()
    StyleButton btnDelete, "DELETE", 324, ThemeRed(), ThemeNavy()
    StyleButton btnClear, "NEW / CLEAR", 476, ThemePanel(), ThemeWhite()
    StyleButton btnClose, "CLOSE", 628, ThemePanel(), ThemeWhite()
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal bgColor As Long, ByVal fgColor As Long)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = 240: .Width = 140: .Height = 34
        .BackColor = bgColor: .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = 10: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub LoadLookups()
    Dim item As Variant

    For Each item In Array("Rent", "Salaries", "Electricity", "Water", "Internet / phone", _
                           "Transport", "Maintenance", "Marketing", "Taxes / fees", "Other")
        Cmb("cmbCategory").AddItem CStr(item)
    Next item
End Sub

Private Sub RefreshList()
    mLoading = True
    modExpenses.FillExpenseList lstExpenses, txtFilter.text
    mLoading = False
End Sub

Private Sub txtFilter_Change()
    If mLoading Then Exit Sub
    RefreshList
End Sub

' ---------- This month's profit panel ----------
Private Sub RefreshMonthPanel()
    Dim r As modProfit.TProfit

    r = modProfit.CalculateProfit(modProfit.MonthStart(), Date)

    Me.Controls("lblMonthRevenue").caption = cur() & Format$(r.revenue, "#,##0.00")
    Me.Controls("lblMonthCost").caption = cur() & Format$(r.CostOfGoods, "#,##0.00")
    Me.Controls("lblMonthGross").caption = cur() & Format$(r.GrossProfit, "#,##0.00")
    Me.Controls("lblMonthExpenses").caption = cur() & Format$(r.Expenses, "#,##0.00")
    Me.Controls("lblMonthNet").caption = cur() & Format$(r.netProfit, "#,##0.00")

    If r.netProfit >= 0 Then
        Me.Controls("lblMonthNet").ForeColor = ThemeGreen()
    Else
        Me.Controls("lblMonthNet").ForeColor = ThemeRed()
    End If
End Sub

' ---------- New-expense mode ----------
Private Sub ClearForm()
    mSelectedID = ""
    mLoading = True

    Txt("txtExpenseID").text = "AUTO"
    Txt("txtExpenseDate").text = Format$(Date, "dd-mm-yyyy")
    Cmb("cmbCategory").ListIndex = -1
    Txt("txtAmount").text = ""
    Txt("txtDescription").text = ""
    lstExpenses.ListIndex = -1

    mLoading = False
    UpdateButtonState
    SetStatus "New expense. Fill in the fields marked * and click SAVE."
    Txt("txtAmount").SetFocus
End Sub

Private Sub UpdateButtonState()
    Dim canEdit As Boolean, isNew As Boolean

    canEdit = (modMain.CurrentRole <> "STAFF")
    isNew = (Len(mSelectedID) = 0)

    btnSave.Enabled = canEdit And isNew
    btnUpdate.Enabled = canEdit And Not isNew
    btnDelete.Enabled = modMain.IsAdmin() And Not isNew

    If Not canEdit Then SetStatus "View only: your role cannot change expenses."
End Sub

' ---------- Selecting an expense ----------
Private Sub lstExpenses_Click()
    Dim p As modExpenses.TExpense
    Dim i As Long

    If mLoading Then Exit Sub
    If lstExpenses.ListIndex < 0 Then Exit Sub

    If modExpenses.LoadExpense(CStr(lstExpenses.List(lstExpenses.ListIndex, 0)), p) Then
        mLoading = True
        mSelectedID = p.expenseID

        Txt("txtExpenseID").text = p.expenseID
        Txt("txtExpenseDate").text = p.ExpenseDate
        Txt("txtAmount").text = CStr(p.amount)
        Txt("txtDescription").text = p.description

        Cmb("cmbCategory").ListIndex = -1
        For i = 0 To Cmb("cmbCategory").ListCount - 1
            If StrComp(CStr(Cmb("cmbCategory").List(i)), p.category, vbTextCompare) = 0 Then
                Cmb("cmbCategory").ListIndex = i
                Exit For
            End If
        Next i

        mLoading = False
        UpdateButtonState
        SetStatus "Editing " & p.expenseID & "."
    End If
End Sub

' ---------- Validation: "" = OK, otherwise the message ----------
Private Function ValidateForm() As String
    Dim msg As String
    Dim dt As Date

    If Not modValidation.TryParseDate(Txt("txtExpenseDate").text, dt) Then
        ValidateForm = "Please enter the date as dd-mm-yyyy (for example 31-12-2026)."
        Exit Function
    End If
    If dt > Date Then
        ValidateForm = "The expense date cannot be in the future."
        Exit Function
    End If

    If Cmb("cmbCategory").ListIndex < 0 Then
        ValidateForm = "Please select a category."
        Exit Function
    End If

    msg = CheckNumber(Txt("txtAmount").text, "amount")
    If Len(msg) > 0 Then
        ValidateForm = msg
        Exit Function
    End If
    If CDbl(Txt("txtAmount").text) <= 0 Then
        ValidateForm = "The amount must be greater than zero."
        Exit Function
    End If
End Function

Private Sub ReadForm(ByRef p As modExpenses.TExpense)
    p.expenseID = mSelectedID
    p.ExpenseDate = Trim$(Txt("txtExpenseDate").text)
    p.category = Cmb("cmbCategory").text
    p.description = Trim$(Txt("txtDescription").text)
    p.amount = CDbl(Txt("txtAmount").text)
End Sub

' ---------- Buttons ----------
Private Sub btnSave_Click()
    Dim p As modExpenses.TExpense
    Dim msg As String, newID As String

    On Error GoTo ErrorHandler

    msg = ValidateForm()
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Please check your input"
        Exit Sub
    End If

    ReadForm p
    newID = modExpenses.AddExpense(p)
    LogAction "ADD EXPENSE", newID, p.category & " " & cur() & Format$(p.amount, "#,##0.00")

    RefreshList
    RefreshMonthPanel
    ClearForm
    SetStatus "Saved " & newID & ". Ready for the next expense."
    Exit Sub

ErrorHandler:
    MsgBox "The expense could not be saved." & vbCrLf & Err.description, vbExclamation, "Save failed"
End Sub

Private Sub btnUpdate_Click()
    Dim p As modExpenses.TExpense
    Dim msg As String

    On Error GoTo ErrorHandler

    If Len(mSelectedID) = 0 Then
        MsgBox "Please select an expense from the list first.", vbInformation, "Update"
        Exit Sub
    End If

    msg = ValidateForm()
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Please check your input"
        Exit Sub
    End If

    ReadForm p
    If modExpenses.UpdateExpense(p) Then
        LogAction "EDIT EXPENSE", p.expenseID, p.category & " " & cur() & Format$(p.amount, "#,##0.00")
        RefreshList
        RefreshMonthPanel
        ClearForm
        SetStatus "Updated " & p.expenseID & "."
    Else
        MsgBox "That expense could not be found.", vbExclamation, "Update failed"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "The expense could not be updated." & vbCrLf & Err.description, vbExclamation, "Update failed"
End Sub

Private Sub btnDelete_Click()
    Dim msg As String, deletedID As String

    On Error GoTo ErrorHandler

    If Len(mSelectedID) = 0 Then
        MsgBox "Please select an expense from the list first.", vbInformation, "Delete"
        Exit Sub
    End If

    deletedID = mSelectedID

    If MsgBox("Delete expense " & deletedID & "?" & vbCrLf & "This cannot be undone.", _
              vbYesNo + vbQuestion, "Confirm delete") <> vbYes Then Exit Sub

    msg = modExpenses.DeleteExpense(deletedID)
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Cannot delete"
    Else
        LogAction "DELETE EXPENSE", deletedID, "Expense deleted"
        RefreshList
        RefreshMonthPanel
        ClearForm
        SetStatus "Deleted " & deletedID & "."
    End If
    Exit Sub

ErrorHandler:
    MsgBox "The expense could not be deleted." & vbCrLf & Err.description, vbExclamation, "Delete failed"
End Sub

Private Sub btnClear_Click()
    ClearForm
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


