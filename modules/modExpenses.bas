Attribute VB_Name = "modExpenses"
Option Explicit

' ============================================================
' modExpenses - everything that reads or writes tblExpenses
' ============================================================

Public Type TExpense
    expenseID As String
    ExpenseDate As String        ' dd-mm-yyyy
    category As String
    description As String
    amount As Double
End Type

Private Function ExpensesTable() As ListObject
    Set ExpensesTable = modDatabase.GetTable("Expenses", "tblExpenses")
End Function

' ---------- One-time sheet formatting (run once) ----------
Public Sub PrepareExpensesSheet()
    Dim tbl As ListObject

    On Error GoTo ErrorHandler
    Set tbl = ExpensesTable()

    tbl.ListColumns("ExpenseDate").DataBodyRange.numberFormat = "dd-mm-yyyy"
    tbl.ListColumns("Amount").DataBodyRange.numberFormat = "#,##0.00"
    ThisWorkbook.Worksheets("Expenses").Columns.AutoFit

    MsgBox "Expenses sheet formatted.", vbInformation, "Setup"
    Exit Sub

ErrorHandler:
    MsgBox "Could not format the Expenses sheet." & vbCrLf & Err.description, vbExclamation, "Setup"
End Sub

Public Function FindExpenseRow(ByVal expenseID As String) As Long
    FindExpenseRow = modDatabase.FindRowByValue(ExpensesTable(), "ExpenseID", expenseID)
End Function

' ---------- Add / Update / Delete ----------
Public Function AddExpense(ByRef p As TExpense) As String
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim dt As Date

    If Not modValidation.TryParseDate(p.ExpenseDate, dt) Then
        Err.Raise vbObjectError + 540, , "The expense date is not valid."
    End If

    Set tbl = ExpensesTable()
    p.expenseID = GenerateID("EXP")

    Set newRow = NextTableRow(tbl)
    ' Order matches tblExpenses (6 columns)
    newRow.Range.value = Array(p.expenseID, dt, p.category, p.description, p.amount, modMain.CurrentUserID)

    AddExpense = p.expenseID
End Function

Public Function UpdateExpense(ByRef p As TExpense) As Boolean
    Dim tbl As ListObject
    Dim rowNo As Long
    Dim dt As Date

    If Not modValidation.TryParseDate(p.ExpenseDate, dt) Then
        Err.Raise vbObjectError + 541, , "The expense date is not valid."
    End If

    Set tbl = ExpensesTable()
    rowNo = FindExpenseRow(p.expenseID)
    If rowNo = 0 Then Exit Function

    modDatabase.SetCellValue tbl, rowNo, "ExpenseDate", dt
    modDatabase.SetCellValue tbl, rowNo, "Category", p.category
    modDatabase.SetCellValue tbl, rowNo, "Description", p.description
    modDatabase.SetCellValue tbl, rowNo, "Amount", p.amount

    UpdateExpense = True
End Function

' Returns "" if deleted, otherwise the reason it could not be deleted.
Public Function DeleteExpense(ByVal expenseID As String) As String
    Dim tbl As ListObject
    Dim rowNo As Long

    Set tbl = ExpensesTable()
    rowNo = FindExpenseRow(expenseID)
    If rowNo = 0 Then
        DeleteExpense = "That expense could not be found."
        Exit Function
    End If

    tbl.ListRows(rowNo).Delete
    DeleteExpense = ""
End Function

' ---------- Loading ----------
Public Function LoadExpense(ByVal expenseID As String, ByRef p As TExpense) As Boolean
    Dim tbl As ListObject
    Dim rowNo As Long
    Dim v As Variant

    rowNo = FindExpenseRow(expenseID)
    If rowNo = 0 Then Exit Function
    Set tbl = ExpensesTable()

    p.expenseID = CStr(modDatabase.GetCellValue(tbl, rowNo, "ExpenseID"))
    p.category = CStr(modDatabase.GetCellValue(tbl, rowNo, "Category"))
    p.description = CStr(modDatabase.GetCellValue(tbl, rowNo, "Description"))
    p.amount = ToNumber(modDatabase.GetCellValue(tbl, rowNo, "Amount"))

    v = modDatabase.GetCellValue(tbl, rowNo, "ExpenseDate")
    If IsDate(v) Then p.ExpenseDate = Format$(CDate(v), "dd-mm-yyyy") Else p.ExpenseDate = ""

    LoadExpense = True
End Function

' ---------- Fill the list, newest first (filter matches ID, category, description, date) ----------
Public Sub FillExpenseList(ByVal lst As MSForms.ListBox, ByVal filterText As String, _
                           Optional ByVal maxRows As Long = 1000)
    Dim tbl As ListObject
    Dim data As Variant
    Dim colID As Long, colDate As Long, colCategory As Long, colDesc As Long, colAmount As Long
    Dim i As Long, shown As Long, n As Long
    Dim dateText As String

    lst.Clear
    lst.ColumnCount = 5
    lst.ColumnWidths = "70 pt;80 pt;120 pt;330 pt;110 pt"

    Set tbl = ExpensesTable()
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    filterText = Trim$(filterText)
    data = tbl.DataBodyRange.value
    colID = tbl.ListColumns("ExpenseID").index
    colDate = tbl.ListColumns("ExpenseDate").index
    colCategory = tbl.ListColumns("Category").index
    colDesc = tbl.ListColumns("Description").index
    colAmount = tbl.ListColumns("Amount").index

    For i = UBound(data, 1) To 1 Step -1
        If Len(CStr(data(i, colID))) > 0 Then
            dateText = ""
            If IsDate(data(i, colDate)) Then dateText = Format$(CDate(data(i, colDate)), "dd-mm-yyyy")

            If Len(filterText) = 0 Or _
               InStr(1, CStr(data(i, colID)), filterText, vbTextCompare) > 0 Or _
               InStr(1, CStr(data(i, colCategory)), filterText, vbTextCompare) > 0 Or _
               InStr(1, CStr(data(i, colDesc)), filterText, vbTextCompare) > 0 Or _
               InStr(1, dateText, filterText, vbTextCompare) > 0 Then

                lst.AddItem CStr(data(i, colID))
                n = lst.ListCount - 1
                lst.List(n, 1) = dateText
                lst.List(n, 2) = CStr(data(i, colCategory))
                lst.List(n, 3) = CStr(data(i, colDesc))
                lst.List(n, 4) = Format$(ToNumber(data(i, colAmount)), "#,##0.00")

                shown = shown + 1
                If shown >= maxRows Then Exit For
            End If
        End If
    Next i
End Sub

