Attribute VB_Name = "modHealth"
Option Explicit

' ============================================================
' modHealth - scans the whole workbook for problems.
' Read-only, except FixCounters. Uses Scripting.Dictionary (Windows Excel).
' Statuses: OK, INFO, WARNING, PROBLEM
' ============================================================

Private Const MAX_RESULTS As Long = 300

Private mStatus(1 To 300) As String
Private mArea(1 To 300) As String
Private mCheck(1 To 300) As String
Private mDetails(1 To 300) As String
Private mFix(1 To 300) As String
Private mCount As Long

' Stores one result line.
Private Sub AddResult(ByVal status As String, ByVal area As String, ByVal checkName As String, _
                      ByVal details As String, ByVal fixTest As String)
    If mCount >= MAX_RESULTS Then Exit Sub
    mCount = mCount + 1
    mStatus(mCount) = status
    mArea(mCount) = area
    mCheck(mCount) = checkName
    mDetails(mCount) = details
    mFix(mCount) = fixTest
End Sub

' ============================================================
' Small helpers
' ============================================================

Private Function DictionaryAvailable() As Boolean
    Dim d As Object

    On Error Resume Next
    Set d = CreateObject("Scripting.Dictionary")
    DictionaryAvailable = Not (d Is Nothing)
    On Error GoTo 0
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

' The table with this name on this sheet, or Nothing.
Private Function FindTableOnSheet(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    If Not SheetExists(sheetName) Then Exit Function

    Set ws = ThisWorkbook.Worksheets(sheetName)
    For Each lo In ws.ListObjects
        If lo.name = tableName Then
            Set FindTableOnSheet = lo
            Exit Function
        End If
    Next lo
End Function

Private Function ColumnExists(ByVal tbl As ListObject, ByVal columnName As String) As Boolean
    Dim lc As ListColumn

    For Each lc In tbl.ListColumns
        If lc.name = columnName Then
            ColumnExists = True
            Exit Function
        End If
    Next lc
End Function

Private Function SettingExists(ByVal settingKey As String) As Boolean
    Dim tbl As ListObject

    Set tbl = modDatabase.GetTable("Settings", "tblSettings")
    If tbl.DataBodyRange Is Nothing Then Exit Function
    SettingExists = Not IsError(Application.Match(settingKey, tbl.ListColumns("SettingKey").DataBodyRange, 0))
End Function

' Reads a whole table into an array (Empty if there are no rows). tbl is returned too.
Private Function TableData(ByVal sheetName As String, ByVal tableName As String, _
                           ByRef tbl As ListObject) As Variant
    Set tbl = modDatabase.GetTable(sheetName, tableName)
    If tbl.DataBodyRange Is Nothing Then
        TableData = Empty
    Else
        TableData = tbl.DataBodyRange.value
    End If
End Function

' A dictionary holding every non-blank value of one column.
Private Function ColumnSet(ByVal sheetName As String, ByVal tableName As String, _
                           ByVal columnName As String) As Object
    Dim tbl As ListObject
    Dim data As Variant
    Dim d As Object
    Dim colNo As Long, i As Long
    Dim v As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    data = TableData(sheetName, tableName, tbl)
    If Not IsEmpty(data) Then
        colNo = tbl.ListColumns(columnName).index
        For i = 1 To UBound(data, 1)
            v = Trim$(CStr(data(i, colNo)))
            If Len(v) > 0 Then d(v) = True
        Next i
    End If

    Set ColumnSet = d
End Function

' Counts repeated values in one column and gives a few examples.
Private Function CountDuplicates(ByVal sheetName As String, ByVal tableName As String, _
                                 ByVal columnName As String, ByRef sample As String) As Long
    Dim tbl As ListObject
    Dim data As Variant
    Dim seen As Object
    Dim colNo As Long, i As Long, dupes As Long
    Dim v As String

    sample = ""
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    data = TableData(sheetName, tableName, tbl)
    If IsEmpty(data) Then Exit Function

    colNo = tbl.ListColumns(columnName).index
    For i = 1 To UBound(data, 1)
        v = Trim$(CStr(data(i, colNo)))
        If Len(v) > 0 Then
            If seen.Exists(v) Then
                dupes = dupes + 1
                If dupes <= 5 Then sample = sample & v & "  "
            Else
                seen(v) = True
            End If
        End If
    Next i

    CountDuplicates = dupes
End Function

' Counts rows whose ID is not in the given set.
Private Function CountNotIn(ByVal sheetName As String, ByVal tableName As String, _
                            ByVal idColumn As String, ByVal knownSet As Object, _
                            ByRef sample As String) As Long
    Dim tbl As ListObject
    Dim data As Variant
    Dim colNo As Long, i As Long, n As Long
    Dim v As String

    sample = ""
    data = TableData(sheetName, tableName, tbl)
    If IsEmpty(data) Then Exit Function

    colNo = tbl.ListColumns(idColumn).index
    For i = 1 To UBound(data, 1)
        v = Trim$(CStr(data(i, colNo)))
        If Len(v) > 0 Then
            If Not knownSet.Exists(v) Then
                n = n + 1
                If n <= 5 Then sample = sample & v & "  "
            End If
        End If
    Next i

    CountNotIn = n
End Function

' Counts values that start with a prefix such as "TST-".
Private Function CountPrefix(ByVal sheetName As String, ByVal tableName As String, _
                             ByVal columnName As String, ByVal prefix As String) As Long
    Dim tbl As ListObject
    Dim data As Variant
    Dim colNo As Long, i As Long, n As Long

    data = TableData(sheetName, tableName, tbl)
    If IsEmpty(data) Then Exit Function

    colNo = tbl.ListColumns(columnName).index
    For i = 1 To UBound(data, 1)
        If Left$(CStr(data(i, colNo)), Len(prefix)) = prefix Then n = n + 1
    Next i

    CountPrefix = n
End Function

' The highest number used in IDs such as PRD-00012 (returns 12).
Private Function MaxIdNumber(ByVal sheetName As String, ByVal tableName As String, _
                             ByVal columnName As String, ByVal prefix As String) As Long
    Dim tbl As ListObject
    Dim data As Variant
    Dim colNo As Long, i As Long
    Dim v As String, numText As String
    Dim num As Long, best As Long

    data = TableData(sheetName, tableName, tbl)
    If IsEmpty(data) Then Exit Function

    colNo = tbl.ListColumns(columnName).index
    For i = 1 To UBound(data, 1)
        v = CStr(data(i, colNo))
        If Left$(v, Len(prefix) + 1) = prefix & "-" Then
            numText = Mid$(v, Len(prefix) + 2)
            If IsNumeric(numText) And Len(numText) <= 9 Then
                num = CLng(numText)
                If num > best Then best = num
            End If
        End If
    Next i

    MaxIdNumber = best
End Function

' Counter name | sheet | table | ID column | prefix
Private Function CounterSpecs() As Variant
    CounterSpecs = Array( _
        "LastProductNo|Products|tblProducts|ProductID|PRD", _
        "LastCustomerNo|Customers|tblCustomers|CustomerID|CUS", _
        "LastSupplierNo|Suppliers|tblSuppliers|SupplierID|SUP", _
        "LastSaleNo|Sales|tblSales|SaleID|SAL", _
        "LastInvoiceNo|Sales|tblSales|InvoiceNo|INV", _
        "LastPurchaseNo|Purchases|tblPurchases|PurchaseID|PUR", _
        "LastPurchaseDetailNo|PurchaseDetails|tblPurchaseDetails|PurchaseDetailID|PDT", _
        "LastSaleDetailNo|SaleDetails|tblSaleDetails|SaleDetailID|SDT", _
        "LastMovementNo|Inventory|tblStockMovements|MovementID|MOV", _
        "LastExpenseNo|Expenses|tblExpenses|ExpenseID|EXP", _
        "LastUserNo|Users|tblUsers|UserID|USR", _
        "LastCategoryNo|Categories|tblCategories|CategoryID|CAT")
End Function

' ============================================================
' The checks
' ============================================================

Private Sub CheckEnvironment()
    Dim ws As Worksheet
    Dim protectedList As String

    On Error GoTo ErrorHandler

    If Len(ThisWorkbook.Path) = 0 Then
        AddResult "WARNING", "Program", "Workbook is saved", _
                  "The workbook has never been saved, so backups, PDF invoices and auto-save cannot work.", _
                  "Save it once with File, Save As (type .xlsm)."
    Else
        AddResult "OK", "Program", "Workbook is saved", "Saved in " & ThisWorkbook.Path, ""
    End If

    If ThisWorkbook.ReadOnly Then
        AddResult "WARNING", "Program", "Workbook is editable", _
                  "The workbook was opened read-only, so auto-save cannot save changes.", _
                  "Close it and open it normally."
    Else
        AddResult "OK", "Program", "Workbook is editable", "Opened normally.", ""
    End If

    For Each ws In ThisWorkbook.Worksheets
        If ws.ProtectContents Then protectedList = protectedList & ws.name & "  "
    Next ws
    If Len(protectedList) > 0 Then
        AddResult "WARNING", "Program", "Sheets are unprotected", _
                  "These sheets are protected: " & protectedList, _
                  "Click OPEN APPLICATION on the Dashboard sheet, or unprotect them (Review, Unprotect Sheet), so the program can save data."
    Else
        AddResult "OK", "Program", "Sheets are unprotected", "The program can write to every sheet.", ""
    End If

    If modSecurity.HidingIsOn() Then
        AddResult "INFO", "Program", "Sheet hiding", "Hiding the data sheets is ON.", ""
    Else
        AddResult "INFO", "Program", "Sheet hiding", _
                  "Hiding the data sheets is OFF (normal while building).", _
                  "Switch it ON before hand-in: Settings, DATABASE SHEETS, or run PrepareFinalRelease."
    End If
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Program", "Program checks", "This check could not run: " & Err.description, ""
End Sub

' Checks sheets, tables and columns. Returns True if the structure is complete.
Private Function CheckStructure() As Boolean
    Dim sheetNames As Variant
    Dim specs(1 To 13) As String
    Dim parts() As String, cols() As String
    Dim item As Variant
    Dim tbl As ListObject
    Dim i As Long, j As Long
    Dim missingSheets As String, missingTables As String, MissingColumns As String

    On Error GoTo ErrorHandler

    sheetNames = Array("Dashboard", "Users", "Products", "Categories", "Suppliers", "Customers", _
                       "Sales", "SaleDetails", "Purchases", "PurchaseDetails", "Inventory", _
                       "Expenses", "Settings", "AuditLog", "Reports")

    For Each item In sheetNames
        If Not SheetExists(CStr(item)) Then missingSheets = missingSheets & CStr(item) & "  "
    Next item

    specs(1) = "Users|tblUsers|UserID,Username,Password,FullName,Role,Status,CreatedDate"
    specs(2) = "Products|tblProducts|ProductID,Barcode,ProductName,Category,Brand,SupplierID,PurchasePrice,SellingPrice,Stock,MinStock,MaxStock,Unit,ExpiryDate,Description,Status,CreatedDate"
    specs(3) = "Categories|tblCategories|CategoryID,CategoryName,Description"
    specs(4) = "Suppliers|tblSuppliers|SupplierID,SupplierName,ContactPerson,Phone,Email,Address,TotalPurchases,CreatedDate"
    specs(5) = "Customers|tblCustomers|CustomerID,CustomerName,Phone,Email,Address,TotalSpent,RegistrationDate"
    specs(6) = "Sales|tblSales|SaleID,InvoiceNo,CustomerID,SaleDate,Subtotal,Discount,Tax,GrandTotal,UserID"
    specs(7) = "SaleDetails|tblSaleDetails|SaleDetailID,SaleID,ProductID,Quantity,Price,CostPrice,Discount,Total"
    specs(8) = "Purchases|tblPurchases|PurchaseID,SupplierID,PurchaseDate,TotalCost,UserID"
    specs(9) = "PurchaseDetails|tblPurchaseDetails|PurchaseDetailID,PurchaseID,ProductID,Quantity,PurchasePrice,Total"
    specs(10) = "Inventory|tblStockMovements|MovementID,MovementDate,ProductID,MovementType,Quantity,StockAfter,ReferenceID,UserID"
    specs(11) = "Expenses|tblExpenses|ExpenseID,ExpenseDate,Category,Description,Amount,UserID"
    specs(12) = "Settings|tblSettings|SettingKey,SettingValue"
    specs(13) = "AuditLog|tblAuditLog|LogDate,LogTime,Username,Action,RecordID,Description"

    For i = 1 To 13
        parts = Split(specs(i), "|")
        If SheetExists(parts(0)) Then
            Set tbl = FindTableOnSheet(parts(0), parts(1))
            If tbl Is Nothing Then
                missingTables = missingTables & parts(1) & " (on " & parts(0) & ")  "
            Else
                cols = Split(parts(2), ",")
                For j = 0 To UBound(cols)
                    If Not ColumnExists(tbl, cols(j)) Then
                        MissingColumns = MissingColumns & parts(1) & "." & cols(j) & "  "
                    End If
                Next j
            End If
        End If
    Next i

    If Len(missingSheets) > 0 Then
        AddResult "PROBLEM", "Structure", "Required sheets", "Missing sheets: " & missingSheets, _
                  "Run CreateAllSheets (module modSetup) to create them again."
    Else
        AddResult "OK", "Structure", "Required sheets", "All 15 required sheets exist.", ""
    End If

    If Len(missingTables) > 0 Then
        AddResult "PROBLEM", "Structure", "Required tables", "Missing tables: " & missingTables, _
                  "Run CreateAllTables (module modTables). It skips sheets that already have a table."
    Else
        AddResult "OK", "Structure", "Required tables", "All 13 data tables exist.", ""
    End If

    If Len(MissingColumns) > 0 Then
        AddResult "PROBLEM", "Structure", "Required columns", "Missing columns: " & MissingColumns, _
                  "Add the column back with its exact name (see the table layout in Phase 3)."
    Else
        AddResult "OK", "Structure", "Required columns", "Every table has all of its columns.", ""
    End If

    CheckStructure = (Len(missingSheets) = 0 And Len(missingTables) = 0 And Len(MissingColumns) = 0)
    Exit Function

ErrorHandler:
    AddResult "PROBLEM", "Structure", "Structure check", "This check could not run: " & Err.description, ""
    CheckStructure = False
End Function

Private Sub CheckSettings()
    Dim keys As Variant
    Dim item As Variant
    Dim missing As String
    Dim taxText As String
    Dim rate As Double

    On Error GoTo ErrorHandler

    keys = Array("ShopName", "ShopAddress", "ShopPhone", "CurrencySymbol", "TaxEnabled", "TaxRate", "SystemVersion")
    For Each item In keys
        If Not SettingExists(CStr(item)) Then missing = missing & CStr(item) & "  "
    Next item

    If Len(missing) > 0 Then
        AddResult "WARNING", "Settings", "Required settings exist", _
                  "Missing settings (default values are used): " & missing, _
                  "Open Shop Settings and click SAVE SETTINGS, or run LoadInitialData (module modInitialData)."
    Else
        AddResult "OK", "Settings", "Required settings exist", "All shop settings are present.", ""
    End If

    If UCase$(Trim$(GetSetting("TaxEnabled", "No"))) = "YES" Then
        taxText = Trim$(GetSetting("TaxRate", "0"))
        If IsNumeric(taxText) Then
            rate = CDbl(taxText)
            If rate < 0 Or rate > 100 Then
                AddResult "PROBLEM", "Settings", "Tax rate", _
                          "Tax is on but the rate (" & taxText & ") is not between 0 and 100.", _
                          "Open Shop Settings and enter a rate between 0 and 100."
            Else
                AddResult "OK", "Settings", "Tax rate", "Tax is on at " & taxText & " percent.", ""
            End If
        Else
            AddResult "PROBLEM", "Settings", "Tax rate", _
                      "Tax is on but the rate (" & taxText & ") is not a number.", _
                      "Open Shop Settings and enter a number between 0 and 100."
        End If
    Else
        AddResult "OK", "Settings", "Tax rate", "Tax is switched off.", ""
    End If
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Settings", "Settings checks", "This check could not run: " & Err.description, ""
End Sub

Private Sub CheckCounters()
    Dim specs As Variant
    Dim item As Variant
    Dim parts() As String
    Dim counterValue As Long, maxUsed As Long
    Dim behind As String
    Dim behindCount As Long

    On Error GoTo ErrorHandler

    specs = CounterSpecs()
    For Each item In specs
        parts = Split(CStr(item), "|")
        counterValue = CLng(val(GetSetting(parts(0), "0")))
        maxUsed = MaxIdNumber(parts(1), parts(2), parts(3), parts(4))
        If counterValue < maxUsed Then
            behindCount = behindCount + 1
            behind = behind & parts(0) & " (counter " & counterValue & ", highest used " & maxUsed & ")   "
        End If
    Next item

    If behindCount = 0 Then
        AddResult "OK", "Counters", "ID counters", "All 12 counters are at or above the highest ID in use.", ""
    Else
        AddResult "PROBLEM", "Counters", "ID counters", _
                  behindCount & " counter(s) are behind, so new IDs could repeat old ones: " & behind, _
                  "Click FIX COUNTERS on this screen."
    End If
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Counters", "ID counters", "This check could not run: " & Err.description, ""
End Sub

Private Sub CheckDuplicates()
    Dim specs(1 To 14) As String
    Dim parts() As String
    Dim i As Long
    Dim dupes As Long, totalDupes As Long
    Dim sample As String

    On Error GoTo ErrorHandler

    specs(1) = "Products|tblProducts|ProductID|Product IDs"
    specs(2) = "Products|tblProducts|Barcode|Barcodes"
    specs(3) = "Customers|tblCustomers|CustomerID|Customer IDs"
    specs(4) = "Suppliers|tblSuppliers|SupplierID|Supplier IDs"
    specs(5) = "Sales|tblSales|SaleID|Sale IDs"
    specs(6) = "Sales|tblSales|InvoiceNo|Invoice numbers"
    specs(7) = "SaleDetails|tblSaleDetails|SaleDetailID|Sale line IDs"
    specs(8) = "Purchases|tblPurchases|PurchaseID|Purchase IDs"
    specs(9) = "PurchaseDetails|tblPurchaseDetails|PurchaseDetailID|Purchase line IDs"
    specs(10) = "Inventory|tblStockMovements|MovementID|Stock movement IDs"
    specs(11) = "Expenses|tblExpenses|ExpenseID|Expense IDs"
    specs(12) = "Users|tblUsers|UserID|User IDs"
    specs(13) = "Users|tblUsers|Username|Usernames"
    specs(14) = "Categories|tblCategories|CategoryName|Category names"

    For i = 1 To 14
        parts = Split(specs(i), "|")
        dupes = CountDuplicates(parts(0), parts(1), parts(2), sample)
        If dupes > 0 Then
            totalDupes = totalDupes + dupes
            AddResult "PROBLEM", "Duplicates", "Duplicate " & LCase$(parts(3)), _
                      dupes & " duplicate(s). Examples: " & sample, _
                      "These values must be unique. Correct or delete the repeated rows, then click FIX COUNTERS if IDs are involved."
        End If
    Next i

    If totalDupes = 0 Then
        AddResult "OK", "Duplicates", "No duplicate IDs, barcodes or names", "Checked 14 columns.", ""
    End If
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Duplicates", "Duplicate checks", "This check could not run: " & Err.description, ""
End Sub

' One "does this value exist in the other table?" check.
Private Sub RelCheck(ByVal childSheet As String, ByVal childTable As String, ByVal childColumn As String, _
                     ByVal parentSet As Object, ByVal allowBlank As Boolean, ByVal level As String, _
                     ByVal title As String, ByVal fixText As String)
    Dim tbl As ListObject
    Dim data As Variant
    Dim colNo As Long, i As Long, missing As Long
    Dim v As String, sample As String

    data = TableData(childSheet, childTable, tbl)
    If Not IsEmpty(data) Then
        colNo = tbl.ListColumns(childColumn).index
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 Then
                v = Trim$(CStr(data(i, colNo)))
                If Len(v) = 0 Then
                    If Not allowBlank Then
                        missing = missing + 1
                        If missing <= 5 Then sample = sample & CStr(data(i, 1)) & "  "
                    End If
                ElseIf Not parentSet.Exists(v) Then
                    missing = missing + 1
                    If missing <= 5 Then sample = sample & CStr(data(i, 1)) & " -> " & v & "   "
                End If
            End If
        Next i
    End If

    If missing = 0 Then
        AddResult "OK", "Relationships", title, "No broken links.", ""
    Else
        AddResult level, "Relationships", title, _
                  missing & " record(s) point to something that does not exist. Examples: " & sample, fixText
    End If
End Sub

Private Sub CheckRelationships()
    Dim suppliers As Object, categories As Object, customers As Object, users As Object
    Dim products As Object, saleIDs As Object, purchaseIDs As Object

    On Error GoTo ErrorHandler

    Set suppliers = ColumnSet("Suppliers", "tblSuppliers", "SupplierID")
    Set categories = ColumnSet("Categories", "tblCategories", "CategoryName")
    Set customers = ColumnSet("Customers", "tblCustomers", "CustomerID")
    Set users = ColumnSet("Users", "tblUsers", "UserID")
    Set products = ColumnSet("Products", "tblProducts", "ProductID")
    Set saleIDs = ColumnSet("Sales", "tblSales", "SaleID")
    Set purchaseIDs = ColumnSet("Purchases", "tblPurchases", "PurchaseID")

    RelCheck "Products", "tblProducts", "SupplierID", suppliers, True, "WARNING", _
             "Products have a valid supplier", "Open the product and choose an existing supplier."
    RelCheck "Products", "tblProducts", "Category", categories, False, "WARNING", _
             "Products have a valid category", "Open the product and choose a category from the list."
    RelCheck "Sales", "tblSales", "CustomerID", customers, False, "PROBLEM", _
             "Sales have a valid customer", "Correct the CustomerID on the Sales sheet."
    RelCheck "Sales", "tblSales", "UserID", users, False, "WARNING", _
             "Sales have a valid user", "The sale was saved by an unknown user. Usually harmless."
    RelCheck "SaleDetails", "tblSaleDetails", "SaleID", saleIDs, False, "PROBLEM", _
             "Sale lines belong to a real sale", "Delete the orphan lines or restore the missing sale."
    RelCheck "SaleDetails", "tblSaleDetails", "ProductID", products, False, "PROBLEM", _
             "Sale lines use real products", "Restore the missing product or correct the ProductID."
    RelCheck "Purchases", "tblPurchases", "SupplierID", suppliers, False, "PROBLEM", _
             "Purchases have a valid supplier", "Correct the SupplierID on the Purchases sheet."
    RelCheck "PurchaseDetails", "tblPurchaseDetails", "PurchaseID", purchaseIDs, False, "PROBLEM", _
             "Purchase lines belong to a real purchase", "Delete the orphan lines or restore the missing purchase."
    RelCheck "PurchaseDetails", "tblPurchaseDetails", "ProductID", products, False, "PROBLEM", _
             "Purchase lines use real products", "Restore the missing product or correct the ProductID."
    RelCheck "Inventory", "tblStockMovements", "ProductID", products, False, "INFO", _
             "Stock movements belong to existing products", "Normal after deleting a product that was never sold or bought."
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Relationships", "Relationship checks", "This check could not run: " & Err.description, ""
End Sub

Private Sub CheckEmptyDocuments()
    Dim lineSet As Object
    Dim sample As String
    Dim n As Long

    On Error GoTo ErrorHandler

    Set lineSet = ColumnSet("SaleDetails", "tblSaleDetails", "SaleID")
    n = CountNotIn("Sales", "tblSales", "SaleID", lineSet, sample)
    If n = 0 Then
        AddResult "OK", "Relationships", "Every sale has lines", "No empty invoices.", ""
    Else
        AddResult "WARNING", "Relationships", "Every sale has lines", _
                  n & " sale(s) have no product lines. Examples: " & sample, _
                  "These invoices show nothing. Check them on the Sales and SaleDetails sheets."
    End If

    Set lineSet = ColumnSet("PurchaseDetails", "tblPurchaseDetails", "PurchaseID")
    n = CountNotIn("Purchases", "tblPurchases", "PurchaseID", lineSet, sample)
    If n = 0 Then
        AddResult "OK", "Relationships", "Every purchase has lines", "No empty purchases.", ""
    Else
        AddResult "WARNING", "Relationships", "Every purchase has lines", _
                  n & " purchase(s) have no product lines. Examples: " & sample, _
                  "Check them on the Purchases and PurchaseDetails sheets."
    End If
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Relationships", "Empty invoices", "This check could not run: " & Err.description, ""
End Sub

Private Sub CheckSalesMoney()
    Dim tblS As ListObject, tblD As ListObject
    Dim salesData As Variant, lineData As Variant
    Dim lineSum As Object
    Dim colSaleID As Long, colSub As Long, colDisc As Long, colTax As Long, colGrand As Long
    Dim colLineSale As Long, colQty As Long, colPrice As Long, colLineDisc As Long, colLineTotal As Long
    Dim i As Long
    Dim badLines As Long, badSubtotals As Long, badTotals As Long
    Dim sampleLines As String, sampleSubtotals As String, sampleTotals As String
    Dim expected As Double
    Dim saleID As String

    On Error GoTo ErrorHandler

    Set lineSum = CreateObject("Scripting.Dictionary")
    lineSum.CompareMode = vbTextCompare

    lineData = TableData("SaleDetails", "tblSaleDetails", tblD)
    If Not IsEmpty(lineData) Then
        colLineSale = tblD.ListColumns("SaleID").index
        colQty = tblD.ListColumns("Quantity").index
        colPrice = tblD.ListColumns("Price").index
        colLineDisc = tblD.ListColumns("Discount").index
        colLineTotal = tblD.ListColumns("Total").index

        For i = 1 To UBound(lineData, 1)
            If Len(CStr(lineData(i, 1))) > 0 Then
                saleID = CStr(lineData(i, colLineSale))
                lineSum(saleID) = lineSum(saleID) + ToNumber(lineData(i, colLineTotal))

                expected = ToNumber(lineData(i, colQty)) * ToNumber(lineData(i, colPrice)) - _
                           ToNumber(lineData(i, colLineDisc))
                If Abs(expected - ToNumber(lineData(i, colLineTotal))) > 0.011 Then
                    badLines = badLines + 1
                    If badLines <= 3 Then sampleLines = sampleLines & CStr(lineData(i, 1)) & "  "
                End If
            End If
        Next i
    End If

    salesData = TableData("Sales", "tblSales", tblS)
    If Not IsEmpty(salesData) Then
        colSaleID = tblS.ListColumns("SaleID").index
        colSub = tblS.ListColumns("Subtotal").index
        colDisc = tblS.ListColumns("Discount").index
        colTax = tblS.ListColumns("Tax").index
        colGrand = tblS.ListColumns("GrandTotal").index

        For i = 1 To UBound(salesData, 1)
            If Len(CStr(salesData(i, colSaleID))) > 0 Then
                saleID = CStr(salesData(i, colSaleID))

                expected = ToNumber(salesData(i, colSub)) - ToNumber(salesData(i, colDisc)) + _
                           ToNumber(salesData(i, colTax))
                If Abs(expected - ToNumber(salesData(i, colGrand))) > 0.011 Then
                    badTotals = badTotals + 1
                    If badTotals <= 3 Then sampleTotals = sampleTotals & saleID & "  "
                End If

                If lineSum.Exists(saleID) Then
                    If Abs(lineSum(saleID) - ToNumber(salesData(i, colSub))) > 0.011 Then
                        badSubtotals = badSubtotals + 1
                        If badSubtotals <= 3 Then sampleSubtotals = sampleSubtotals & saleID & "  "
                    End If
                End If
            End If
        Next i
    End If

    If badLines = 0 Then
        AddResult "OK", "Money", "Sale line totals", "Every line total equals quantity x price - discount.", ""
    Else
        AddResult "PROBLEM", "Money", "Sale line totals", badLines & " line(s) do not add up. Examples: " & sampleLines, _
                  "Check Quantity, Price, Discount and Total on the SaleDetails sheet."
    End If

    If badSubtotals = 0 Then
        AddResult "OK", "Money", "Sale subtotals", "Every sale subtotal equals the sum of its lines.", ""
    Else
        AddResult "PROBLEM", "Money", "Sale subtotals", badSubtotals & " sale(s) differ from their lines. Examples: " & sampleSubtotals, _
                  "Compare the Subtotal on the Sales sheet with the lines on SaleDetails."
    End If

    If badTotals = 0 Then
        AddResult "OK", "Money", "Sale grand totals", "Every grand total equals subtotal - discount + tax.", ""
    Else
        AddResult "PROBLEM", "Money", "Sale grand totals", badTotals & " sale(s) do not add up. Examples: " & sampleTotals, _
                  "Check Subtotal, Discount, Tax and GrandTotal on the Sales sheet."
    End If
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Money", "Sales totals", "This check could not run: " & Err.description, ""
End Sub

Private Sub CheckPurchaseMoney()
    Dim tblP As ListObject, tblD As ListObject
    Dim purchaseData As Variant, lineData As Variant
    Dim lineSum As Object
    Dim colPurchaseID As Long, colCost As Long
    Dim colLinePurchase As Long, colQty As Long, colPrice As Long, colLineTotal As Long
    Dim i As Long
    Dim badLines As Long, badTotals As Long
    Dim sampleLines As String, sampleTotals As String
    Dim expected As Double
    Dim purchaseID As String

    On Error GoTo ErrorHandler

    Set lineSum = CreateObject("Scripting.Dictionary")
    lineSum.CompareMode = vbTextCompare

    lineData = TableData("PurchaseDetails", "tblPurchaseDetails", tblD)
    If Not IsEmpty(lineData) Then
        colLinePurchase = tblD.ListColumns("PurchaseID").index
        colQty = tblD.ListColumns("Quantity").index
        colPrice = tblD.ListColumns("PurchasePrice").index
        colLineTotal = tblD.ListColumns("Total").index

        For i = 1 To UBound(lineData, 1)
            If Len(CStr(lineData(i, 1))) > 0 Then
                purchaseID = CStr(lineData(i, colLinePurchase))
                lineSum(purchaseID) = lineSum(purchaseID) + ToNumber(lineData(i, colLineTotal))

                expected = ToNumber(lineData(i, colQty)) * ToNumber(lineData(i, colPrice))
                If Abs(expected - ToNumber(lineData(i, colLineTotal))) > 0.011 Then
                    badLines = badLines + 1
                    If badLines <= 3 Then sampleLines = sampleLines & CStr(lineData(i, 1)) & "  "
                End If
            End If
        Next i
    End If

    purchaseData = TableData("Purchases", "tblPurchases", tblP)
    If Not IsEmpty(purchaseData) Then
        colPurchaseID = tblP.ListColumns("PurchaseID").index
        colCost = tblP.ListColumns("TotalCost").index

        For i = 1 To UBound(purchaseData, 1)
            If Len(CStr(purchaseData(i, colPurchaseID))) > 0 Then
                purchaseID = CStr(purchaseData(i, colPurchaseID))
                If lineSum.Exists(purchaseID) Then
                    If Abs(lineSum(purchaseID) - ToNumber(purchaseData(i, colCost))) > 0.011 Then
                        badTotals = badTotals + 1
                        If badTotals <= 3 Then sampleTotals = sampleTotals & purchaseID & "  "
                    End If
                End If
            End If
        Next i
    End If

    If badLines = 0 Then
        AddResult "OK", "Money", "Purchase line totals", "Every line total equals quantity x price.", ""
    Else
        AddResult "PROBLEM", "Money", "Purchase line totals", badLines & " line(s) do not add up. Examples: " & sampleLines, _
                  "Check Quantity, PurchasePrice and Total on the PurchaseDetails sheet."
    End If

    If badTotals = 0 Then
        AddResult "OK", "Money", "Purchase totals", "Every purchase total equals the sum of its lines.", ""
    Else
        AddResult "PROBLEM", "Money", "Purchase totals", badTotals & " purchase(s) differ from their lines. Examples: " & sampleTotals, _
                  "Compare TotalCost on the Purchases sheet with the lines on PurchaseDetails."
    End If
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Money", "Purchase totals", "This check could not run: " & Err.description, ""
End Sub

Private Sub CheckProducts()
    Dim tbl As ListObject
    Dim data As Variant
    Dim colID As Long, colCost As Long, colPrice As Long, colStock As Long
    Dim colMin As Long, colMax As Long, colStatus As Long
    Dim i As Long
    Dim negStock As Long, negPrice As Long, minMax As Long, belowCost As Long, badStatus As Long
    Dim sNeg As String, sPrice As String, sMinMax As String
    Dim stockQty As Double, costPrice As Double, sellPrice As Double
    Dim productStatus As String

    On Error GoTo ErrorHandler

    data = TableData("Products", "tblProducts", tbl)
    If Not IsEmpty(data) Then
        colID = tbl.ListColumns("ProductID").index
        colCost = tbl.ListColumns("PurchasePrice").index
        colPrice = tbl.ListColumns("SellingPrice").index
        colStock = tbl.ListColumns("Stock").index
        colMin = tbl.ListColumns("MinStock").index
        colMax = tbl.ListColumns("MaxStock").index
        colStatus = tbl.ListColumns("Status").index

        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, colID))) > 0 Then
                stockQty = ToNumber(data(i, colStock))
                costPrice = ToNumber(data(i, colCost))
                sellPrice = ToNumber(data(i, colPrice))
                productStatus = CStr(data(i, colStatus))

                If stockQty < 0 Then
                    negStock = negStock + 1
                    If negStock <= 5 Then sNeg = sNeg & CStr(data(i, colID)) & "  "
                End If
                If costPrice < 0 Or sellPrice < 0 Then
                    negPrice = negPrice + 1
                    If negPrice <= 5 Then sPrice = sPrice & CStr(data(i, colID)) & "  "
                End If
                If ToNumber(data(i, colMin)) > ToNumber(data(i, colMax)) Then
                    minMax = minMax + 1
                    If minMax <= 5 Then sMinMax = sMinMax & CStr(data(i, colID)) & "  "
                End If
                If sellPrice < costPrice And sellPrice >= 0 Then belowCost = belowCost + 1
                If productStatus <> "Active" And productStatus <> "Inactive" Then badStatus = badStatus + 1
            End If
        Next i
    End If

    If negStock = 0 Then
        AddResult "OK", "Products", "No negative stock", "No product has less than zero in stock.", ""
    Else
        AddResult "PROBLEM", "Products", "No negative stock", negStock & " product(s) have negative stock: " & sNeg, _
                  "Use Stock, ADJUST STOCK to set the real counted quantity."
    End If

    If negPrice = 0 Then
        AddResult "OK", "Products", "No negative prices", "All prices are zero or more.", ""
    Else
        AddResult "PROBLEM", "Products", "No negative prices", negPrice & " product(s) have a negative price: " & sPrice, _
                  "Open the product and correct the price."
    End If

    If minMax = 0 Then
        AddResult "OK", "Products", "Stock levels make sense", "No product has a minimum above its maximum.", ""
    Else
        AddResult "WARNING", "Products", "Stock levels make sense", minMax & " product(s) have a minimum above the maximum: " & sMinMax, _
                  "Open the product and correct the levels."
    End If

    If badStatus = 0 Then
        AddResult "OK", "Products", "Product status", "Every product is Active or Inactive.", ""
    Else
        AddResult "WARNING", "Products", "Product status", badStatus & " product(s) have an unknown status.", _
                  "Open the product and choose Active or Inactive."
    End If

    If belowCost > 0 Then
        AddResult "INFO", "Products", "Selling below cost", belowCost & " product(s) sell for less than they cost.", _
                  "Intended for clearance items. Otherwise raise the selling price."
    End If
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Products", "Product checks", "This check could not run: " & Err.description, ""
End Sub

Private Sub CheckStockLogMatch()
    Dim report As String
    Dim issues As Long

    On Error GoTo ErrorHandler

    issues = modInventory.CheckStockLog(False, report)
    If issues = 0 Then
        AddResult "OK", "Stock", "Stock agrees with the movement log", _
                  "Every product's stock equals the total of its stock movements.", ""
    Else
        AddResult "PROBLEM", "Stock", "Stock agrees with the movement log", _
                  issues & " product(s) differ. " & Replace(report, vbCrLf, "   "), _
                  "Open Stock and click VERIFY STOCK LOG. An administrator can add correction entries."
    End If
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Stock", "Stock vs movement log", "This check could not run: " & Err.description, ""
End Sub

Private Sub CheckUsers()
    Dim tbl As ListObject
    Dim data As Variant
    Dim colID As Long, colName As Long, colPass As Long, colRole As Long, colStatus As Long
    Dim i As Long
    Dim activeAdmins As Long
    Dim defaultNames As String, badRoles As String, badStatus As String
    Dim userRole As String, userStatus As String

    On Error GoTo ErrorHandler

    data = TableData("Users", "tblUsers", tbl)
    If Not IsEmpty(data) Then
        colID = tbl.ListColumns("UserID").index
        colName = tbl.ListColumns("Username").index
        colPass = tbl.ListColumns("Password").index
        colRole = tbl.ListColumns("Role").index
        colStatus = tbl.ListColumns("Status").index

        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, colID))) > 0 Then
                userRole = CStr(data(i, colRole))
                userStatus = CStr(data(i, colStatus))

                If userRole <> "ADMIN" And userRole <> "MANAGER" And userRole <> "STAFF" Then
                    badRoles = badRoles & CStr(data(i, colName)) & "  "
                End If
                If userStatus <> "Active" And userStatus <> "Disabled" Then
                    badStatus = badStatus & CStr(data(i, colName)) & "  "
                End If
                If userRole = "ADMIN" And userStatus = "Active" Then activeAdmins = activeAdmins + 1

                If CStr(data(i, colPass)) = modUsers.HashPassword(LCase$(CStr(data(i, colName))) & "123") Then
                    defaultNames = defaultNames & CStr(data(i, colName)) & "  "
                End If
            End If
        Next i
    End If

    If activeAdmins >= 1 Then
        AddResult "OK", "Security", "An active administrator exists", activeAdmins & " active ADMIN account(s).", ""
    Else
        AddResult "PROBLEM", "Security", "An active administrator exists", _
                  "There is no active ADMIN, so nobody can open the Administration screen.", _
                  "Run EmergencyResetAdmin (module modRelease) from the VBA editor."
    End If

    If Len(badRoles) = 0 And Len(badStatus) = 0 Then
        AddResult "OK", "Security", "User roles and status", "Every user has a valid role and status.", ""
    Else
        AddResult "PROBLEM", "Security", "User roles and status", _
                  "Invalid role: " & badRoles & "   Invalid status: " & badStatus, _
                  "Open Users and Passwords and correct them."
    End If

    If Len(defaultNames) = 0 Then
        AddResult "OK", "Security", "Default passwords", "No account uses its starting password.", ""
    Else
        AddResult "WARNING", "Security", "Default passwords", "Still using the starting password: " & defaultNames, _
                  "Each user should click PASSWORD on the dashboard, or an admin can reset it in Users and Passwords."
    End If
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Security", "User checks", "This check could not run: " & Err.description, ""
End Sub

Private Sub CheckWalkIn()
    On Error GoTo ErrorHandler

    If modCustomers.FindCustomerRow(modCustomers.WALK_IN_ID) > 0 Then
        AddResult "OK", "Customers", "Walk-in Customer exists", "The built-in customer for anonymous sales is present.", ""
    Else
        AddResult "WARNING", "Customers", "Walk-in Customer exists", _
                  "The Walk-in Customer is missing, so the Sales screen has no default customer.", _
                  "Run EnsureWalkInCustomer (module modCustomers)."
    End If
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Customers", "Walk-in Customer", "This check could not run: " & Err.description, ""
End Sub

' Compares the saved totals on the Customers and Suppliers sheets with the real totals.
Private Sub CheckCaches()
    Dim tblS As ListObject, tblC As ListObject, tblP As ListObject, tblU As ListObject
    Dim salesData As Variant, custData As Variant, purData As Variant, suppData As Variant
    Dim totals As Object
    Dim colKey As Long, colAmount As Long, colID As Long, colStored As Long
    Dim i As Long
    Dim id As String
    Dim computed As Double
    Dim badCustomers As Long, badSuppliers As Long

    On Error GoTo ErrorHandler

    ' --- Customers: saved TotalSpent vs the sum of their sales ---
    Set totals = CreateObject("Scripting.Dictionary")
    totals.CompareMode = vbTextCompare

    salesData = TableData("Sales", "tblSales", tblS)
    If Not IsEmpty(salesData) Then
        colKey = tblS.ListColumns("CustomerID").index
        colAmount = tblS.ListColumns("GrandTotal").index
        For i = 1 To UBound(salesData, 1)
            If Len(CStr(salesData(i, 1))) > 0 Then
                id = CStr(salesData(i, colKey))
                totals(id) = totals(id) + ToNumber(salesData(i, colAmount))
            End If
        Next i
    End If

    custData = TableData("Customers", "tblCustomers", tblC)
    If Not IsEmpty(custData) Then
        colID = tblC.ListColumns("CustomerID").index
        colStored = tblC.ListColumns("TotalSpent").index
        For i = 1 To UBound(custData, 1)
            If Len(CStr(custData(i, colID))) > 0 Then
                computed = 0
                If totals.Exists(CStr(custData(i, colID))) Then computed = totals(CStr(custData(i, colID)))
                If Abs(computed - ToNumber(custData(i, colStored))) > 0.5 Then badCustomers = badCustomers + 1
            End If
        Next i
    End If

    ' --- Suppliers: saved TotalPurchases vs the sum of their purchases ---
    Set totals = CreateObject("Scripting.Dictionary")
    totals.CompareMode = vbTextCompare

    purData = TableData("Purchases", "tblPurchases", tblP)
    If Not IsEmpty(purData) Then
        colKey = tblP.ListColumns("SupplierID").index
        colAmount = tblP.ListColumns("TotalCost").index
        For i = 1 To UBound(purData, 1)
            If Len(CStr(purData(i, 1))) > 0 Then
                id = CStr(purData(i, colKey))
                totals(id) = totals(id) + ToNumber(purData(i, colAmount))
            End If
        Next i
    End If

    suppData = TableData("Suppliers", "tblSuppliers", tblU)
    If Not IsEmpty(suppData) Then
        colID = tblU.ListColumns("SupplierID").index
        colStored = tblU.ListColumns("TotalPurchases").index
        For i = 1 To UBound(suppData, 1)
            If Len(CStr(suppData(i, colID))) > 0 Then
                computed = 0
                If totals.Exists(CStr(suppData(i, colID))) Then computed = totals(CStr(suppData(i, colID)))
                If Abs(computed - ToNumber(suppData(i, colStored))) > 0.5 Then badSuppliers = badSuppliers + 1
            End If
        Next i
    End If

    If badCustomers = 0 And badSuppliers = 0 Then
        AddResult "OK", "Money", "Saved customer and supplier totals", _
                  "The totals saved on the Customers and Suppliers sheets match the real totals.", ""
    Else
        AddResult "INFO", "Money", "Saved customer and supplier totals", _
                  badCustomers & " customer(s) and " & badSuppliers & " supplier(s) have a saved total that differs. " & _
                  "The screens always recalculate from the sales and purchases, so nothing is wrong.", _
                  "Usually caused by test rows typed in by hand. No action needed."
    End If
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Money", "Saved totals", "This check could not run: " & Err.description, ""
End Sub

Private Sub CheckLeftovers()
    Dim n As Long

    On Error GoTo ErrorHandler

    n = CountPrefix("Products", "tblProducts", "ProductID", "TST-") + _
        CountPrefix("Customers", "tblCustomers", "CustomerID", "TST-") + _
        CountPrefix("Suppliers", "tblSuppliers", "SupplierID", "TST-") + _
        CountPrefix("Sales", "tblSales", "SaleID", "TST-")

    If n = 0 Then
        AddResult "OK", "Housekeeping", "No leftover test rows", "No rows with a TST- ID were found.", ""
    Else
        AddResult "WARNING", "Housekeeping", "No leftover test rows", _
                  n & " test row(s) from the early phases (IDs starting with TST-) are still in the data.", _
                  "Run ClearTestData (module modTestData), or ClearDemoData to start completely clean."
    End If
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Housekeeping", "Leftover test rows", "This check could not run: " & Err.description, ""
End Sub

Private Sub CheckBackupAge()
    Dim t As String
    Dim lastDate As Date
    Dim ageDays As Long

    On Error GoTo ErrorHandler

    t = Trim$(GetSetting("LastBackupDate", ""))
    If Len(t) <> 8 Or Not IsNumeric(t) Then
        AddResult "WARNING", "Backup", "Recent backup", "No backup has been made yet.", _
                  "Settings, BACKUP DATABASE, BACKUP NOW."
        Exit Sub
    End If

    lastDate = DateSerial(CLng(Left$(t, 4)), CLng(Mid$(t, 5, 2)), CLng(Right$(t, 2)))
    ageDays = CLng(Date - lastDate)

    If ageDays > 7 Then
        AddResult "WARNING", "Backup", "Recent backup", _
                  "The last backup was " & Format$(lastDate, "dd-mm-yyyy") & " (" & ageDays & " days ago).", _
                  "Settings, BACKUP DATABASE, BACKUP NOW."
    Else
        AddResult "OK", "Backup", "Recent backup", _
                  "The last backup was " & Format$(lastDate, "dd-mm-yyyy") & " (" & ageDays & " day(s) ago).", ""
    End If

    If Not modBackup.IsAutoBackupOn() Then
        AddResult "INFO", "Backup", "Automatic backup", "Automatic daily backup is OFF.", _
                  "Switch it ON in Settings, BACKUP DATABASE."
    End If
    Exit Sub

ErrorHandler:
    AddResult "WARNING", "Backup", "Backup check", "This check could not run: " & Err.description, ""
End Sub

' ============================================================
' Running everything
' ============================================================

' Runs every check. Returns a 2D array (1 To rows, 1 To 5):
' Status, Area, Check, Details, How to fix.
Public Function RunHealthCheck() As Variant
    Dim results() As Variant
    Dim structureOK As Boolean
    Dim i As Long

    mCount = 0

    If Not DictionaryAvailable() Then
        AddResult "PROBLEM", "Environment", "Required Windows component", _
                  "This computer cannot create the Scripting.Dictionary object. Reports, charts and alerts need Windows Excel.", _
                  "Run the program in desktop Excel for Windows."
    Else
        CheckEnvironment
        structureOK = CheckStructure()

        If structureOK Then
            CheckSettings
            CheckCounters
            CheckDuplicates
            CheckRelationships
            CheckEmptyDocuments
            CheckSalesMoney
            CheckPurchaseMoney
            CheckProducts
            CheckStockLogMatch
            CheckUsers
            CheckWalkIn
            CheckCaches
            CheckLeftovers
            CheckBackupAge
        Else
            AddResult "INFO", "Structure", "Data checks", _
                      "The data checks were skipped because the structure above is not complete.", _
                      "Fix the structure problems first, then run the check again."
        End If
    End If

    If mCount = 0 Then AddResult "INFO", "Health check", "No checks ran", "Nothing was checked.", ""

    ReDim results(1 To mCount, 1 To 5)
    For i = 1 To mCount
        results(i, 1) = mStatus(i)
        results(i, 2) = mArea(i)
        results(i, 3) = mCheck(i)
        results(i, 4) = mDetails(i)
        results(i, 5) = mFix(i)
    Next i

    RunHealthCheck = results
End Function

' Counts the results by status.
Public Sub CountStatuses(ByVal results As Variant, ByRef okCount As Long, ByRef infoCount As Long, _
                         ByRef warnCount As Long, ByRef problemCount As Long)
    Dim i As Long

    okCount = 0
    infoCount = 0
    warnCount = 0
    problemCount = 0

    If IsEmpty(results) Then Exit Sub

    For i = 1 To UBound(results, 1)
        Select Case CStr(results(i, 1))
            Case "OK"
                okCount = okCount + 1
            Case "INFO"
                infoCount = infoCount + 1
            Case "WARNING"
                warnCount = warnCount + 1
            Case "PROBLEM"
                problemCount = problemCount + 1
        End Select
    Next i
End Sub

' Raises every ID counter that is behind the highest ID in use. Returns how many were fixed.
Public Function FixCounters() As Long
    Dim specs As Variant
    Dim item As Variant
    Dim parts() As String
    Dim counterValue As Long, maxUsed As Long
    Dim fixedCount As Long

    specs = CounterSpecs()
    For Each item In specs
        parts = Split(CStr(item), "|")
        counterValue = CLng(val(GetSetting(parts(0), "0")))
        maxUsed = MaxIdNumber(parts(1), parts(2), parts(3), parts(4))
        If counterValue < maxUsed Then
            SetSetting parts(0), CStr(maxUsed)
            fixedCount = fixedCount + 1
        End If
    Next item

    FixCounters = fixedCount
End Function

' Writes the results onto the Reports sheet. Returns how many lines were written.
Public Function WriteHealthToSheet(ByVal results As Variant) As Long
    Dim ws As Worksheet
    Dim n As Long, i As Long
    Dim okCount As Long, infoCount As Long, warnCount As Long, problemCount As Long

    If IsEmpty(results) Then Exit Function

    n = UBound(results, 1)
    CountStatuses results, okCount, infoCount, warnCount, problemCount

    Set ws = ThisWorkbook.Worksheets("Reports")
    ws.Cells.Clear

    With ws.Range("A1")
        .value = "HEALTH CHECK"
        .Font.Size = 16
        .Font.Bold = True
    End With
    ws.Range("A2").value = "Generated " & Format$(Now, "dd-mm-yyyy hh:nn") & " by " & modMain.CurrentUsername & _
                           "     " & problemCount & " problem(s), " & warnCount & " warning(s), " & _
                           infoCount & " note(s), " & okCount & " passed"

    ws.Range("A4:E4").value = Array("Status", "Area", "Check", "Details", "How to fix")
    With ws.Range("A4:E4")
        .Interior.Color = ThemeNavy()
        .Font.Color = ThemeWhite()
        .Font.Bold = True
    End With

    ws.Range(ws.Cells(5, 1), ws.Cells(4 + n, 5)).numberFormat = "@"
    ws.Range(ws.Cells(5, 1), ws.Cells(4 + n, 5)).value = results

    For i = 1 To n
        With ws.Cells(4 + i, 1)
            .Font.Bold = True
            Select Case CStr(results(i, 1))
                Case "PROBLEM": .Font.Color = RGB(192, 0, 0)
                Case "WARNING": .Font.Color = RGB(204, 102, 0)
                Case "OK": .Font.Color = RGB(0, 128, 0)
                Case Else: .Font.Color = RGB(100, 100, 100)
            End Select
        End With
    Next i

    ws.Columns("A").ColumnWidth = 11
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 38
    ws.Columns("D").ColumnWidth = 70
    ws.Columns("E").ColumnWidth = 55

    With ws.Range(ws.Cells(4, 1), ws.Cells(4 + n, 5))
        .WrapText = True
        .VerticalAlignment = xlTop
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range(ws.Cells(5, 1), ws.Cells(4 + n, 5)).Rows.AutoFit

    WriteHealthToSheet = n
End Function

