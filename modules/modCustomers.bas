Attribute VB_Name = "modCustomers"
Option Explicit

' ============================================================
' modCustomers - everything that reads or writes tblCustomers
' ============================================================

' The built-in customer used for sales without a name.
Public Const WALK_IN_ID As String = "CUS-00000"

Public Type TCustomer
    customerID As String
    customerName As String
    phone As String
    Email As String
    Address As String
    RegistrationDate As String        ' dd-mm-yyyy (display only)
End Type

Private Function CustomersTable() As ListObject
    Set CustomersTable = modDatabase.GetTable("Customers", "tblCustomers")
End Function

' ---------- One-time setup (run each once) ----------
Public Sub PrepareCustomersSheet()
    Dim tbl As ListObject

    On Error GoTo ErrorHandler
    Set tbl = CustomersTable()

    tbl.ListColumns("Phone").DataBodyRange.numberFormat = "@"          ' keep leading zeros
    tbl.ListColumns("TotalSpent").DataBodyRange.numberFormat = "#,##0.00"
    tbl.ListColumns("RegistrationDate").DataBodyRange.numberFormat = "dd-mm-yyyy"
    ThisWorkbook.Worksheets("Customers").Columns.AutoFit

    MsgBox "Customers sheet formatted.", vbInformation, "Setup"
    Exit Sub

ErrorHandler:
    MsgBox "Could not format the Customers sheet." & vbCrLf & Err.description, vbExclamation, "Setup"
End Sub

' Creates the Walk-in Customer if it does not exist yet. Safe to run again.
Public Sub EnsureWalkInCustomer()
    Dim tbl As ListObject
    Dim newRow As ListRow

    On Error GoTo ErrorHandler

    If FindCustomerRow(WALK_IN_ID) > 0 Then
        MsgBox "The Walk-in Customer already exists.", vbInformation, "Setup"
        Exit Sub
    End If

    Set tbl = CustomersTable()
    Set newRow = NextTableRow(tbl)
    ' Order must match the 7 columns of tblCustomers
    newRow.Range.value = Array(WALK_IN_ID, "Walk-in Customer", "", "", "", 0, Date)

    MsgBox "Walk-in Customer created.", vbInformation, "Setup"
    Exit Sub

ErrorHandler:
    MsgBox "Could not create the Walk-in Customer." & vbCrLf & Err.description, vbExclamation, "Setup"
End Sub

Public Function IsSystemCustomer(ByVal customerID As String) As Boolean
    IsSystemCustomer = (StrComp(customerID, WALK_IN_ID, vbTextCompare) = 0)
End Function

' ---------- Finding ----------
Public Function FindCustomerRow(ByVal customerID As String) As Long
    FindCustomerRow = modDatabase.FindRowByValue(CustomersTable(), "CustomerID", customerID)
End Function

' Keeps only the digits, so "0300-123 4567" and "03001234567" compare as equal.
Private Function DigitsOnly(ByVal inputText As String) As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(inputText)
        ch = Mid$(inputText, i, 1)
        If ch Like "[0-9]" Then DigitsOnly = DigitsOnly & ch
    Next i
End Function

' True if ANOTHER customer already has this phone number.
Public Function PhoneExists(ByVal phone As String, ByVal excludeID As String) As Boolean
    Dim tbl As ListObject
    Dim data As Variant
    Dim wanted As String
    Dim phoneCol As Long, idCol As Long, i As Long

    wanted = DigitsOnly(phone)
    If Len(wanted) = 0 Then Exit Function

    Set tbl = CustomersTable()
    If tbl.DataBodyRange Is Nothing Then Exit Function

    data = tbl.DataBodyRange.value
    phoneCol = tbl.ListColumns("Phone").index
    idCol = tbl.ListColumns("CustomerID").index

    For i = 1 To UBound(data, 1)
        If DigitsOnly(CStr(data(i, phoneCol))) = wanted And CStr(data(i, idCol)) <> excludeID Then
            PhoneExists = True
            Exit Function
        End If
    Next i
End Function

' ---------- Add / Update / Delete ----------
Public Function AddCustomer(ByRef p As TCustomer) As String
    Dim tbl As ListObject
    Dim newRow As ListRow

    Set tbl = CustomersTable()
    p.customerID = GenerateID("CUS")

    Set newRow = NextTableRow(tbl)
    newRow.Range.value = Array(p.customerID, p.customerName, p.phone, p.Email, p.Address, 0, Date)

    AddCustomer = p.customerID
End Function

' TotalSpent and RegistrationDate are never changed by editing.
Public Function UpdateCustomer(ByRef p As TCustomer) As Boolean
    Dim tbl As ListObject
    Dim rowNo As Long

    If IsSystemCustomer(p.customerID) Then Exit Function

    Set tbl = CustomersTable()
    rowNo = FindCustomerRow(p.customerID)
    If rowNo = 0 Then Exit Function

    modDatabase.SetCellValue tbl, rowNo, "CustomerName", p.customerName
    modDatabase.SetCellValue tbl, rowNo, "Phone", p.phone
    modDatabase.SetCellValue tbl, rowNo, "Email", p.Email
    modDatabase.SetCellValue tbl, rowNo, "Address", p.Address

    UpdateCustomer = True
End Function

' Returns "" if deleted, otherwise the reason it could not be deleted.
Public Function DeleteCustomer(ByVal customerID As String) As String
    Dim tbl As ListObject
    Dim rowNo As Long

    If IsSystemCustomer(customerID) Then
        DeleteCustomer = "The Walk-in Customer is built in and cannot be deleted."
        Exit Function
    End If

    If CountPurchases(customerID) > 0 Then
        DeleteCustomer = "This customer has purchase history, so it cannot be deleted."
        Exit Function
    End If

    Set tbl = CustomersTable()
    rowNo = FindCustomerRow(customerID)
    If rowNo = 0 Then
        DeleteCustomer = "That customer could not be found."
        Exit Function
    End If

    tbl.ListRows(rowNo).Delete
    DeleteCustomer = ""
End Function

' ---------- Loading and calculated values ----------
Public Function LoadCustomer(ByVal customerID As String, ByRef p As TCustomer) As Boolean
    Dim tbl As ListObject
    Dim rowNo As Long
    Dim v As Variant

    rowNo = FindCustomerRow(customerID)
    If rowNo = 0 Then Exit Function
    Set tbl = CustomersTable()

    p.customerID = CStr(modDatabase.GetCellValue(tbl, rowNo, "CustomerID"))
    p.customerName = CStr(modDatabase.GetCellValue(tbl, rowNo, "CustomerName"))
    p.phone = CStr(modDatabase.GetCellValue(tbl, rowNo, "Phone"))
    p.Email = CStr(modDatabase.GetCellValue(tbl, rowNo, "Email"))
    p.Address = CStr(modDatabase.GetCellValue(tbl, rowNo, "Address"))

    v = modDatabase.GetCellValue(tbl, rowNo, "RegistrationDate")
    If IsDate(v) Then p.RegistrationDate = Format$(CDate(v), "dd-mm-yyyy") Else p.RegistrationDate = ""

    LoadCustomer = True
End Function

' Number of invoices this customer has.
Public Function CountPurchases(ByVal customerID As String) As Long
    CountPurchases = modDatabase.CountMatches("Sales", "tblSales", "CustomerID", customerID)
End Function

' Total money this customer has spent (calculated from the Sales sheet).
Public Function TotalSpentBy(ByVal customerID As String) As Double
    TotalSpentBy = modDatabase.SumMatches("Sales", "tblSales", "CustomerID", customerID, "GrandTotal")
End Function

' Customer count for the dashboard card (does not count the Walk-in Customer).
Public Function RealCustomerCount() As Long
    RealCustomerCount = modDashboard.CountRecords("Customers", "tblCustomers")
    If FindCustomerRow(WALK_IN_ID) > 0 Then RealCustomerCount = RealCustomerCount - 1
End Function

' ---------- Fill the list (filter matches ID, name, phone or email) ----------
Public Sub FillCustomerList(ByVal lst As MSForms.ListBox, ByVal filterText As String)
    Dim tbl As ListObject
    Dim data As Variant
    Dim idCol As Long, nameCol As Long, phoneCol As Long, emailCol As Long
    Dim i As Long

    lst.Clear
    lst.ColumnCount = 3
    lst.ColumnWidths = "62 pt;130 pt;70 pt"

    Set tbl = CustomersTable()
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    filterText = Trim$(filterText)
    data = tbl.DataBodyRange.value
    idCol = tbl.ListColumns("CustomerID").index
    nameCol = tbl.ListColumns("CustomerName").index
    phoneCol = tbl.ListColumns("Phone").index
    emailCol = tbl.ListColumns("Email").index

    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, idCol))) > 0 Then
            If Len(filterText) = 0 Or _
               InStr(1, CStr(data(i, idCol)), filterText, vbTextCompare) > 0 Or _
               InStr(1, CStr(data(i, nameCol)), filterText, vbTextCompare) > 0 Or _
               InStr(1, CStr(data(i, phoneCol)), filterText, vbTextCompare) > 0 Or _
               InStr(1, CStr(data(i, emailCol)), filterText, vbTextCompare) > 0 Then

                lst.AddItem CStr(data(i, idCol))
                lst.List(lst.ListCount - 1, 1) = CStr(data(i, nameCol))
                lst.List(lst.ListCount - 1, 2) = CStr(data(i, phoneCol))
            End If
        End If
    Next i
End Sub
' Two columns: hidden CustomerID + visible "Name - phone".
Public Sub LoadCustomerCombo(ByVal Cmb As MSForms.ComboBox)
    Dim tbl As ListObject
    Dim data As Variant
    Dim idCol As Long, nameCol As Long, phoneCol As Long, i As Long
    Dim display As String

    Cmb.Clear
    Cmb.ColumnCount = 2
    Cmb.ColumnWidths = "0 pt;230 pt"
    Cmb.textColumn = 2

    Set tbl = CustomersTable()
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    data = tbl.DataBodyRange.value
    idCol = tbl.ListColumns("CustomerID").index
    nameCol = tbl.ListColumns("CustomerName").index
    phoneCol = tbl.ListColumns("Phone").index

    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, idCol))) > 0 Then
            display = CStr(data(i, nameCol))
            If Len(CStr(data(i, phoneCol))) > 0 Then display = display & "  -  " & CStr(data(i, phoneCol))
            Cmb.AddItem CStr(data(i, idCol))
            Cmb.List(Cmb.ListCount - 1, 1) = display
        End If
    Next i
End Sub

' Keeps the TotalSpent column on the Customers sheet up to date after a sale.
' (The Customer screen itself always recalculates from the Sales sheet.)
Public Sub AddToCustomerTotal(ByVal customerID As String, ByVal amount As Double)
    Dim tbl As ListObject
    Dim rowNo As Long

    rowNo = FindCustomerRow(customerID)
    If rowNo = 0 Then Exit Sub

    Set tbl = CustomersTable()
    modDatabase.SetCellValue tbl, rowNo, "TotalSpent", _
        ToNumber(modDatabase.GetCellValue(tbl, rowNo, "TotalSpent")) + amount
End Sub
