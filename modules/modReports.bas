Attribute VB_Name = "modReports"
Option Explicit

' ============================================================
' modReports - builds every report as a table of values.
' Uses Scripting.Dictionary (Windows Excel).
' ============================================================

' ---------- The list of reports (names and keys must stay in the same order) ----------
Private Function ReportNames() As Variant
    ReportNames = Array("Sales - Daily", "Sales - Weekly", "Sales - Monthly", "Sales - Yearly", _
                        "Sales - By product", "Sales - By customer", _
                        "Purchases - Daily", "Purchases - Monthly", "Purchases - By supplier", _
                        "Inventory - Current inventory", "Inventory - Low stock", _
                        "Inventory - Out of stock", "Inventory - Stock movement", _
                        "Financial - Profit and loss")
End Function

Private Function ReportKeys() As Variant
    ReportKeys = Array("SALES_DAY", "SALES_WEEK", "SALES_MONTH", "SALES_YEAR", _
                       "SALES_PRODUCT", "SALES_CUSTOMER", _
                       "PUR_DAY", "PUR_MONTH", "PUR_SUPPLIER", _
                       "INV_CURRENT", "INV_LOW", "INV_OUT", "INV_MOVES", "FIN_PL")
End Function

Public Function ReportCount() As Long
    Dim names As Variant
    names = ReportNames()
    ReportCount = UBound(names) + 1
End Function

Public Function ReportNameAt(ByVal index As Long) As String
    Dim names As Variant
    names = ReportNames()
    ReportNameAt = CStr(names(index))
End Function

Public Function ReportKeyAt(ByVal index As Long) As String
    Dim keys As Variant
    keys = ReportKeys()
    ReportKeyAt = CStr(keys(index))
End Function

' The inventory snapshot reports do not use the date range.
Public Function UsesDateRange(ByVal reportKey As String) As Boolean
    Select Case reportKey
        Case "INV_CURRENT", "INV_LOW", "INV_OUT"
            UsesDateRange = False
        Case Else
            UsesDateRange = True
    End Select
End Function

' ---------- Build any report ----------
' Returns a 2D array (1 To rows, 1 To columns), or Empty when there is no data.
' title, headers, formats, widths and summary are always filled in.
Public Function BuildReport(ByVal reportKey As String, ByVal fromDate As Date, ByVal toDate As Date, _
                            ByRef title As String, ByRef headers As Variant, ByRef formats As Variant, _
                            ByRef widths As String, ByRef summary As String) As Variant
    Dim rangeStart As Date, rangeEnd As Date

    rangeStart = DateValue(fromDate)
    rangeEnd = DateValue(toDate) + 1                 ' the end is exclusive: the whole "to" day counts
    summary = ""

    Select Case reportKey
        Case "SALES_DAY":      BuildReport = SalesByPeriod("DAY", rangeStart, rangeEnd, title, headers, formats, widths, summary)
        Case "SALES_WEEK":     BuildReport = SalesByPeriod("WEEK", rangeStart, rangeEnd, title, headers, formats, widths, summary)
        Case "SALES_MONTH":    BuildReport = SalesByPeriod("MONTH", rangeStart, rangeEnd, title, headers, formats, widths, summary)
        Case "SALES_YEAR":     BuildReport = SalesByPeriod("YEAR", rangeStart, rangeEnd, title, headers, formats, widths, summary)
        Case "SALES_PRODUCT":  BuildReport = SalesByProduct(rangeStart, rangeEnd, title, headers, formats, widths, summary)
        Case "SALES_CUSTOMER": BuildReport = SalesByCustomer(rangeStart, rangeEnd, title, headers, formats, widths, summary)
        Case "PUR_DAY":        BuildReport = PurchasesByPeriod("DAY", rangeStart, rangeEnd, title, headers, formats, widths, summary)
        Case "PUR_MONTH":      BuildReport = PurchasesByPeriod("MONTH", rangeStart, rangeEnd, title, headers, formats, widths, summary)
        Case "PUR_SUPPLIER":   BuildReport = PurchasesBySupplier(rangeStart, rangeEnd, title, headers, formats, widths, summary)
        Case "INV_CURRENT":    BuildReport = InventoryCurrent(title, headers, formats, widths, summary)
        Case "INV_LOW":        BuildReport = InventoryAlertReport("LOW", title, headers, formats, widths, summary)
        Case "INV_OUT":        BuildReport = InventoryAlertReport("OUT", title, headers, formats, widths, summary)
        Case "INV_MOVES":      BuildReport = StockMovements(rangeStart, rangeEnd, title, headers, formats, widths, summary)
        Case "FIN_PL":         BuildReport = ProfitAndLoss(fromDate, toDate, title, headers, formats, widths, summary)
        Case Else
            Err.Raise vbObjectError + 550, , "Unknown report: " & reportKey
    End Select
End Function

' ============================================================
' Small helpers
' ============================================================

' Reads a whole table into an array (Empty if the table has no rows). tbl is returned too.
Private Function TableValues(ByVal sheetName As String, ByVal tableName As String, _
                             ByRef tbl As ListObject) As Variant
    Set tbl = modDatabase.GetTable(sheetName, tableName)
    If tbl.DataBodyRange Is Nothing Then
        TableValues = Empty
    Else
        TableValues = tbl.DataBodyRange.value
    End If
End Function

' Dictionary of ID -> name for any table.
Private Function NameMap(ByVal sheetName As String, ByVal tableName As String, _
                         ByVal idColumn As String, ByVal nameColumn As String) As Object
    Dim map As Object
    Dim tbl As ListObject
    Dim data As Variant
    Dim idCol As Long, nameCol As Long, i As Long

    Set map = CreateObject("Scripting.Dictionary")
    Set tbl = modDatabase.GetTable(sheetName, tableName)

    If Not tbl.DataBodyRange Is Nothing Then
        data = tbl.DataBodyRange.value
        idCol = tbl.ListColumns(idColumn).index
        nameCol = tbl.ListColumns(nameColumn).index
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, idCol))) > 0 Then map(CStr(data(i, idCol))) = CStr(data(i, nameCol))
        Next i
    End If

    Set NameMap = map
End Function

' Name for an ID, or the ID itself if no name is known.
Private Function NameOf(ByVal map As Object, ByVal id As String) As String
    If map.Exists(id) Then
        NameOf = CStr(map(id))
    Else
        NameOf = id
    End If
End Function

Private Function InRange(ByVal v As Variant, ByVal rangeStart As Date, ByVal rangeEnd As Date) As Boolean
    Dim dt As Date
    If IsDate(v) Then
        dt = CDate(v)
        InRange = (dt >= rangeStart And dt < rangeEnd)
    End If
End Function

' Sortable key for a date: day, week (starting Monday), month or year.
Private Function PeriodKey(ByVal mode As String, ByVal dt As Date) As String
    Dim WeekStart As Date

    Select Case mode
        Case "DAY"
            PeriodKey = Format$(dt, "yyyy-mm-dd")
        Case "WEEK"
            WeekStart = DateValue(dt) - (Weekday(dt, vbMonday) - 1)
            PeriodKey = Format$(WeekStart, "yyyy-mm-dd")
        Case "MONTH"
            PeriodKey = Format$(dt, "yyyy-mm")
        Case Else
            PeriodKey = Format$(dt, "yyyy")
    End Select
End Function

Private Function KeyToDate(ByVal key As String) As Date
    KeyToDate = DateSerial(CLng(Left$(key, 4)), CLng(Mid$(key, 6, 2)), CLng(Mid$(key, 9, 2)))
End Function

' Text shown for a period key.
Private Function PeriodLabel(ByVal mode As String, ByVal key As String) As String
    Select Case mode
        Case "DAY"
            PeriodLabel = Format$(KeyToDate(key), "dd-mm-yyyy")
        Case "WEEK"
            PeriodLabel = "Week of " & Format$(KeyToDate(key), "dd-mm-yyyy")
        Case "MONTH"
            PeriodLabel = Format$(DateSerial(CLng(Left$(key, 4)), CLng(Mid$(key, 6, 2)), 1), "mmm yyyy")
        Case Else
            PeriodLabel = key
    End Select
End Function

' Sorts an array of text keys A-Z (in place).
Private Sub SortKeysAscending(ByRef keys As Variant)
    Dim i As Long, j As Long
    Dim tmp As Variant

    For i = LBound(keys) + 1 To UBound(keys)
        tmp = keys(i)
        j = i - 1
        Do While j >= LBound(keys)
            If CStr(keys(j)) <= CStr(tmp) Then Exit Do
            keys(j + 1) = keys(j)
            j = j - 1
        Loop
        keys(j + 1) = tmp
    Next i
End Sub

' Sorts keys so the biggest value (from the dictionary) comes first (in place).
Private Sub SortKeysByValueDesc(ByRef keys As Variant, ByVal values As Object)
    Dim i As Long, j As Long
    Dim tmp As Variant

    For i = LBound(keys) + 1 To UBound(keys)
        tmp = keys(i)
        j = i - 1
        Do While j >= LBound(keys)
            If values(CStr(keys(j))) >= values(CStr(tmp)) Then Exit Do
            keys(j + 1) = keys(j)
            j = j - 1
        Loop
        keys(j + 1) = tmp
    Next i
End Sub

' ============================================================
' Sales reports
' ============================================================

Private Function SalesByPeriod(ByVal mode As String, ByVal rangeStart As Date, ByVal rangeEnd As Date, _
                               ByRef title As String, ByRef headers As Variant, ByRef formats As Variant, _
                               ByRef widths As String, ByRef summary As String) As Variant
    Dim tblS As ListObject, tblD As ListObject
    Dim sales As Variant, details As Variant
    Dim saleCost As Object, cnt As Object, rev As Object, taxAmt As Object, tot As Object, costAmt As Object
    Dim colSaleID As Long, colDate As Long, colSub As Long, colDisc As Long, colTax As Long, colGrand As Long
    Dim colDetSale As Long, colDetQty As Long, colDetCost As Long
    Dim keys As Variant
    Dim result() As Variant
    Dim i As Long, k As Long
    Dim key As String, saleID As String
    Dim totInvoices As Long
    Dim totRevenue As Double, totCost As Double
    Dim cur As String

    Select Case mode
        Case "DAY":   title = "Sales - Daily"
        Case "WEEK":  title = "Sales - Weekly"
        Case "MONTH": title = "Sales - Monthly"
        Case Else:    title = "Sales - Yearly"
    End Select
    headers = Array("Period", "Invoices", "Revenue", "Tax", "Total", "Cost", "Gross Profit")
    formats = Array("@", "#,##0", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00")
    widths = "130 pt;70 pt;110 pt;90 pt;110 pt;110 pt;110 pt"
    cur = GetSetting("CurrencySymbol", "Rs.") & " "

    sales = TableValues("Sales", "tblSales", tblS)
    If IsEmpty(sales) Then Exit Function

    Set saleCost = CreateObject("Scripting.Dictionary")
    Set cnt = CreateObject("Scripting.Dictionary")
    Set rev = CreateObject("Scripting.Dictionary")
    Set taxAmt = CreateObject("Scripting.Dictionary")
    Set tot = CreateObject("Scripting.Dictionary")
    Set costAmt = CreateObject("Scripting.Dictionary")

    ' Cost of each sale (from its lines)
    details = TableValues("SaleDetails", "tblSaleDetails", tblD)
    If Not IsEmpty(details) Then
        colDetSale = tblD.ListColumns("SaleID").index
        colDetQty = tblD.ListColumns("Quantity").index
        colDetCost = tblD.ListColumns("CostPrice").index
        For i = 1 To UBound(details, 1)
            saleID = CStr(details(i, colDetSale))
            If Len(saleID) > 0 Then
                saleCost(saleID) = saleCost(saleID) + ToNumber(details(i, colDetQty)) * ToNumber(details(i, colDetCost))
            End If
        Next i
    End If

    colSaleID = tblS.ListColumns("SaleID").index
    colDate = tblS.ListColumns("SaleDate").index
    colSub = tblS.ListColumns("Subtotal").index
    colDisc = tblS.ListColumns("Discount").index
    colTax = tblS.ListColumns("Tax").index
    colGrand = tblS.ListColumns("GrandTotal").index

    For i = 1 To UBound(sales, 1)
        If Len(CStr(sales(i, colSaleID))) > 0 Then
            If InRange(sales(i, colDate), rangeStart, rangeEnd) Then
                key = PeriodKey(mode, CDate(sales(i, colDate)))
                saleID = CStr(sales(i, colSaleID))

                cnt(key) = cnt(key) + 1
                rev(key) = rev(key) + ToNumber(sales(i, colSub)) - ToNumber(sales(i, colDisc))
                taxAmt(key) = taxAmt(key) + ToNumber(sales(i, colTax))
                tot(key) = tot(key) + ToNumber(sales(i, colGrand))
                If Not costAmt.Exists(key) Then costAmt(key) = 0
                If saleCost.Exists(saleID) Then costAmt(key) = costAmt(key) + saleCost(saleID)
            End If
        End If
    Next i

    If cnt.count = 0 Then Exit Function

    keys = cnt.keys
    SortKeysAscending keys

    ReDim result(1 To cnt.count, 1 To 7)
    For k = 1 To cnt.count
        key = CStr(keys(k - 1))
        result(k, 1) = PeriodLabel(mode, key)
        result(k, 2) = cnt(key)
        result(k, 3) = Round(rev(key), 2)
        result(k, 4) = Round(taxAmt(key), 2)
        result(k, 5) = Round(tot(key), 2)
        result(k, 6) = Round(costAmt(key), 2)
        result(k, 7) = Round(rev(key) - costAmt(key), 2)

        totInvoices = totInvoices + cnt(key)
        totRevenue = totRevenue + rev(key)
        totCost = totCost + costAmt(key)
    Next k

    summary = totInvoices & " invoice(s)   |   Revenue " & cur & Format$(totRevenue, "#,##0.00") & _
              "   |   Gross profit " & cur & Format$(totRevenue - totCost, "#,##0.00")
    SalesByPeriod = result
End Function

Private Function SalesByProduct(ByVal rangeStart As Date, ByVal rangeEnd As Date, _
                                ByRef title As String, ByRef headers As Variant, ByRef formats As Variant, _
                                ByRef widths As String, ByRef summary As String) As Variant
    Dim tblS As ListObject, tblD As ListObject
    Dim sales As Variant, details As Variant
    Dim inRangeSales As Object, qty As Object, rev As Object, costAmt As Object, names As Object
    Dim colSaleID As Long, colDate As Long
    Dim colDetSale As Long, colDetProduct As Long, colDetQty As Long, colDetCost As Long, colDetTotal As Long
    Dim keys As Variant
    Dim result() As Variant
    Dim i As Long, k As Long
    Dim productID As String
    Dim totRevenue As Double, totCost As Double
    Dim cur As String

    title = "Sales - By product"
    headers = Array("Product ID", "Product", "Qty Sold", "Revenue", "Cost", "Gross Profit")
    formats = Array("@", "@", "General Number", "#,##0.00", "#,##0.00", "#,##0.00")
    widths = "75 pt;250 pt;80 pt;110 pt;110 pt;110 pt"
    cur = GetSetting("CurrencySymbol", "Rs.") & " "

    sales = TableValues("Sales", "tblSales", tblS)
    If IsEmpty(sales) Then Exit Function
    details = TableValues("SaleDetails", "tblSaleDetails", tblD)
    If IsEmpty(details) Then Exit Function

    Set inRangeSales = CreateObject("Scripting.Dictionary")
    Set qty = CreateObject("Scripting.Dictionary")
    Set rev = CreateObject("Scripting.Dictionary")
    Set costAmt = CreateObject("Scripting.Dictionary")

    colSaleID = tblS.ListColumns("SaleID").index
    colDate = tblS.ListColumns("SaleDate").index
    For i = 1 To UBound(sales, 1)
        If Len(CStr(sales(i, colSaleID))) > 0 Then
            If InRange(sales(i, colDate), rangeStart, rangeEnd) Then inRangeSales(CStr(sales(i, colSaleID))) = True
        End If
    Next i

    colDetSale = tblD.ListColumns("SaleID").index
    colDetProduct = tblD.ListColumns("ProductID").index
    colDetQty = tblD.ListColumns("Quantity").index
    colDetCost = tblD.ListColumns("CostPrice").index
    colDetTotal = tblD.ListColumns("Total").index

    For i = 1 To UBound(details, 1)
        If inRangeSales.Exists(CStr(details(i, colDetSale))) Then
            productID = CStr(details(i, colDetProduct))
            If Len(productID) > 0 Then
                qty(productID) = qty(productID) + ToNumber(details(i, colDetQty))
                rev(productID) = rev(productID) + ToNumber(details(i, colDetTotal))
                costAmt(productID) = costAmt(productID) + ToNumber(details(i, colDetQty)) * ToNumber(details(i, colDetCost))
            End If
        End If
    Next i

    If qty.count = 0 Then Exit Function

    Set names = NameMap("Products", "tblProducts", "ProductID", "ProductName")
    keys = qty.keys
    SortKeysByValueDesc keys, rev

    ReDim result(1 To qty.count, 1 To 6)
    For k = 1 To qty.count
        productID = CStr(keys(k - 1))
        result(k, 1) = productID
        result(k, 2) = NameOf(names, productID)
        result(k, 3) = qty(productID)
        result(k, 4) = Round(rev(productID), 2)
        result(k, 5) = Round(costAmt(productID), 2)
        result(k, 6) = Round(rev(productID) - costAmt(productID), 2)

        totRevenue = totRevenue + rev(productID)
        totCost = totCost + costAmt(productID)
    Next k

    summary = qty.count & " product(s)   |   Revenue " & cur & Format$(totRevenue, "#,##0.00") & _
              "   |   Gross profit " & cur & Format$(totRevenue - totCost, "#,##0.00") & _
              "   |   (invoice-level discounts are not included here)"
    SalesByProduct = result
End Function

Private Function SalesByCustomer(ByVal rangeStart As Date, ByVal rangeEnd As Date, _
                                 ByRef title As String, ByRef headers As Variant, ByRef formats As Variant, _
                                 ByRef widths As String, ByRef summary As String) As Variant
    Dim tbl As ListObject
    Dim sales As Variant
    Dim cnt As Object, rev As Object, tot As Object, names As Object
    Dim colDate As Long, colCustomer As Long, colSub As Long, colDisc As Long, colGrand As Long, colSaleID As Long
    Dim keys As Variant
    Dim result() As Variant
    Dim i As Long, k As Long
    Dim customerID As String
    Dim totInvoices As Long
    Dim totPaid As Double
    Dim cur As String

    title = "Sales - By customer"
    headers = Array("Customer ID", "Customer", "Invoices", "Revenue", "Total")
    formats = Array("@", "@", "#,##0", "#,##0.00", "#,##0.00")
    widths = "85 pt;250 pt;70 pt;130 pt;130 pt"
    cur = GetSetting("CurrencySymbol", "Rs.") & " "

    sales = TableValues("Sales", "tblSales", tbl)
    If IsEmpty(sales) Then Exit Function

    Set cnt = CreateObject("Scripting.Dictionary")
    Set rev = CreateObject("Scripting.Dictionary")
    Set tot = CreateObject("Scripting.Dictionary")

    colSaleID = tbl.ListColumns("SaleID").index
    colDate = tbl.ListColumns("SaleDate").index
    colCustomer = tbl.ListColumns("CustomerID").index
    colSub = tbl.ListColumns("Subtotal").index
    colDisc = tbl.ListColumns("Discount").index
    colGrand = tbl.ListColumns("GrandTotal").index

    For i = 1 To UBound(sales, 1)
        If Len(CStr(sales(i, colSaleID))) > 0 Then
            If InRange(sales(i, colDate), rangeStart, rangeEnd) Then
                customerID = CStr(sales(i, colCustomer))
                cnt(customerID) = cnt(customerID) + 1
                rev(customerID) = rev(customerID) + ToNumber(sales(i, colSub)) - ToNumber(sales(i, colDisc))
                tot(customerID) = tot(customerID) + ToNumber(sales(i, colGrand))
            End If
        End If
    Next i

    If cnt.count = 0 Then Exit Function

    Set names = NameMap("Customers", "tblCustomers", "CustomerID", "CustomerName")
    keys = cnt.keys
    SortKeysByValueDesc keys, tot

    ReDim result(1 To cnt.count, 1 To 5)
    For k = 1 To cnt.count
        customerID = CStr(keys(k - 1))
        result(k, 1) = customerID
        result(k, 2) = NameOf(names, customerID)
        result(k, 3) = cnt(customerID)
        result(k, 4) = Round(rev(customerID), 2)
        result(k, 5) = Round(tot(customerID), 2)

        totInvoices = totInvoices + cnt(customerID)
        totPaid = totPaid + tot(customerID)
    Next k

    summary = cnt.count & " customer(s)   |   " & totInvoices & " invoice(s)   |   Total " & _
              cur & Format$(totPaid, "#,##0.00")
    SalesByCustomer = result
End Function

' ============================================================
' Purchase reports
' ============================================================

Private Function PurchasesByPeriod(ByVal mode As String, ByVal rangeStart As Date, ByVal rangeEnd As Date, _
                                   ByRef title As String, ByRef headers As Variant, ByRef formats As Variant, _
                                   ByRef widths As String, ByRef summary As String) As Variant
    Dim tbl As ListObject
    Dim data As Variant
    Dim cnt As Object, tot As Object
    Dim colID As Long, colDate As Long, colCost As Long
    Dim keys As Variant
    Dim result() As Variant
    Dim i As Long, k As Long
    Dim key As String
    Dim totPurchases As Long
    Dim totCost As Double
    Dim cur As String

    If mode = "DAY" Then title = "Purchases - Daily" Else title = "Purchases - Monthly"
    headers = Array("Period", "Purchases", "Total Cost")
    formats = Array("@", "#,##0", "#,##0.00")
    widths = "160 pt;100 pt;140 pt"
    cur = GetSetting("CurrencySymbol", "Rs.") & " "

    data = TableValues("Purchases", "tblPurchases", tbl)
    If IsEmpty(data) Then Exit Function

    Set cnt = CreateObject("Scripting.Dictionary")
    Set tot = CreateObject("Scripting.Dictionary")

    colID = tbl.ListColumns("PurchaseID").index
    colDate = tbl.ListColumns("PurchaseDate").index
    colCost = tbl.ListColumns("TotalCost").index

    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, colID))) > 0 Then
            If InRange(data(i, colDate), rangeStart, rangeEnd) Then
                key = PeriodKey(mode, CDate(data(i, colDate)))
                cnt(key) = cnt(key) + 1
                tot(key) = tot(key) + ToNumber(data(i, colCost))
            End If
        End If
    Next i

    If cnt.count = 0 Then Exit Function

    keys = cnt.keys
    SortKeysAscending keys

    ReDim result(1 To cnt.count, 1 To 3)
    For k = 1 To cnt.count
        key = CStr(keys(k - 1))
        result(k, 1) = PeriodLabel(mode, key)
        result(k, 2) = cnt(key)
        result(k, 3) = Round(tot(key), 2)

        totPurchases = totPurchases + cnt(key)
        totCost = totCost + tot(key)
    Next k

    summary = totPurchases & " purchase(s)   |   Total cost " & cur & Format$(totCost, "#,##0.00")
    PurchasesByPeriod = result
End Function

Private Function PurchasesBySupplier(ByVal rangeStart As Date, ByVal rangeEnd As Date, _
                                     ByRef title As String, ByRef headers As Variant, ByRef formats As Variant, _
                                     ByRef widths As String, ByRef summary As String) As Variant
    Dim tbl As ListObject
    Dim data As Variant
    Dim cnt As Object, tot As Object, names As Object
    Dim colID As Long, colDate As Long, colSupplier As Long, colCost As Long
    Dim keys As Variant
    Dim result() As Variant
    Dim i As Long, k As Long
    Dim supplierID As String
    Dim totPurchases As Long
    Dim totCost As Double
    Dim cur As String

    title = "Purchases - By supplier"
    headers = Array("Supplier ID", "Supplier", "Purchases", "Total Cost")
    formats = Array("@", "@", "#,##0", "#,##0.00")
    widths = "90 pt;260 pt;90 pt;140 pt"
    cur = GetSetting("CurrencySymbol", "Rs.") & " "

    data = TableValues("Purchases", "tblPurchases", tbl)
    If IsEmpty(data) Then Exit Function

    Set cnt = CreateObject("Scripting.Dictionary")
    Set tot = CreateObject("Scripting.Dictionary")

    colID = tbl.ListColumns("PurchaseID").index
    colDate = tbl.ListColumns("PurchaseDate").index
    colSupplier = tbl.ListColumns("SupplierID").index
    colCost = tbl.ListColumns("TotalCost").index

    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, colID))) > 0 Then
            If InRange(data(i, colDate), rangeStart, rangeEnd) Then
                supplierID = CStr(data(i, colSupplier))
                cnt(supplierID) = cnt(supplierID) + 1
                tot(supplierID) = tot(supplierID) + ToNumber(data(i, colCost))
            End If
        End If
    Next i

    If cnt.count = 0 Then Exit Function

    Set names = NameMap("Suppliers", "tblSuppliers", "SupplierID", "SupplierName")
    keys = cnt.keys
    SortKeysByValueDesc keys, tot

    ReDim result(1 To cnt.count, 1 To 4)
    For k = 1 To cnt.count
        supplierID = CStr(keys(k - 1))
        result(k, 1) = supplierID
        result(k, 2) = NameOf(names, supplierID)
        result(k, 3) = cnt(supplierID)
        result(k, 4) = Round(tot(supplierID), 2)

        totPurchases = totPurchases + cnt(supplierID)
        totCost = totCost + tot(supplierID)
    Next k

    summary = cnt.count & " supplier(s)   |   " & totPurchases & " purchase(s)   |   Total cost " & _
              cur & Format$(totCost, "#,##0.00")
    PurchasesBySupplier = result
End Function

' ============================================================
' Inventory reports
' ============================================================

Private Function InventoryCurrent(ByRef title As String, ByRef headers As Variant, ByRef formats As Variant, _
                                  ByRef widths As String, ByRef summary As String) As Variant
    Dim tbl As ListObject
    Dim data As Variant
    Dim colID As Long, colName As Long, colCategory As Long, colStock As Long
    Dim colMin As Long, colCost As Long, colStatus As Long
    Dim result() As Variant
    Dim i As Long, n As Long, k As Long
    Dim stockQty As Double, costPrice As Double
    Dim totalUnits As Double, totalValue As Double
    Dim cur As String

    title = "Inventory - Current inventory"
    headers = Array("Product ID", "Product", "Category", "Stock", "Min", "Cost Price", "Stock Value", "Status")
    formats = Array("@", "@", "@", "General Number", "General Number", "#,##0.00", "#,##0.00", "@")
    widths = "70 pt;190 pt;90 pt;55 pt;45 pt;80 pt;90 pt;90 pt"
    cur = GetSetting("CurrencySymbol", "Rs.") & " "

    data = TableValues("Products", "tblProducts", tbl)
    If IsEmpty(data) Then Exit Function

    colID = tbl.ListColumns("ProductID").index
    colName = tbl.ListColumns("ProductName").index
    colCategory = tbl.ListColumns("Category").index
    colStock = tbl.ListColumns("Stock").index
    colMin = tbl.ListColumns("MinStock").index
    colCost = tbl.ListColumns("PurchasePrice").index
    colStatus = tbl.ListColumns("Status").index

    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, colID))) > 0 Then n = n + 1
    Next i
    If n = 0 Then Exit Function

    ReDim result(1 To n, 1 To 8)
    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, colID))) > 0 Then
            k = k + 1
            stockQty = ToNumber(data(i, colStock))
            costPrice = ToNumber(data(i, colCost))

            result(k, 1) = CStr(data(i, colID))
            result(k, 2) = CStr(data(i, colName))
            result(k, 3) = CStr(data(i, colCategory))
            result(k, 4) = stockQty
            result(k, 5) = ToNumber(data(i, colMin))
            result(k, 6) = costPrice
            result(k, 7) = Round(stockQty * costPrice, 2)
            result(k, 8) = modInventory.ProductStockState(stockQty, ToNumber(data(i, colMin)), CStr(data(i, colStatus)))

            totalUnits = totalUnits + stockQty
            totalValue = totalValue + stockQty * costPrice
        End If
    Next i

    summary = n & " product(s)   |   " & Format$(totalUnits, "#,##0") & " units   |   Stock value (cost) " & _
              cur & Format$(totalValue, "#,##0.00")
    InventoryCurrent = result
End Function

' Low-stock and out-of-stock reports reuse the Phase 14 alert logic. mode: "LOW" or "OUT".
Private Function InventoryAlertReport(ByVal mode As String, _
                                      ByRef title As String, ByRef headers As Variant, ByRef formats As Variant, _
                                      ByRef widths As String, ByRef summary As String) As Variant
    Dim alertRows As Variant

    If mode = "LOW" Then title = "Inventory - Low stock" Else title = "Inventory - Out of stock"
    headers = Array("Product ID", "Product", "Category", "Supplier", "Stock", "Min", "Max", "Suggested Order", "Status")
    formats = Array("@", "@", "@", "@", "General Number", "General Number", "General Number", "General Number", "@")
    widths = "70 pt;190 pt;85 pt;130 pt;45 pt;40 pt;40 pt;70 pt;90 pt"

    alertRows = modAlerts.BuildLowStockRows(mode, "")
    If IsEmpty(alertRows) Then Exit Function

    summary = UBound(alertRows, 1) & " product(s)"
    InventoryAlertReport = alertRows
End Function

Private Function StockMovements(ByVal rangeStart As Date, ByVal rangeEnd As Date, _
                                ByRef title As String, ByRef headers As Variant, ByRef formats As Variant, _
                                ByRef widths As String, ByRef summary As String) As Variant
    Dim tbl As ListObject
    Dim data As Variant
    Dim products As Object, users As Object
    Dim colDate As Long, colProduct As Long, colType As Long, colQty As Long
    Dim colAfter As Long, colRef As Long, colUser As Long
    Dim result() As Variant
    Dim i As Long, n As Long, k As Long
    Dim productID As String, userID As String

    title = "Inventory - Stock movement"
    headers = Array("Date", "Product ID", "Product", "Type", "Change", "Stock After", "Reference", "User")
    formats = Array("dd-mm-yyyy hh:mm", "@", "@", "@", "General Number", "General Number", "@", "@")
    widths = "100 pt;70 pt;170 pt;85 pt;50 pt;65 pt;150 pt;60 pt"

    data = TableValues("Inventory", "tblStockMovements", tbl)
    If IsEmpty(data) Then Exit Function

    colDate = tbl.ListColumns("MovementDate").index
    colProduct = tbl.ListColumns("ProductID").index
    colType = tbl.ListColumns("MovementType").index
    colQty = tbl.ListColumns("Quantity").index
    colAfter = tbl.ListColumns("StockAfter").index
    colRef = tbl.ListColumns("ReferenceID").index
    colUser = tbl.ListColumns("UserID").index

    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, colProduct))) > 0 Then
            If InRange(data(i, colDate), rangeStart, rangeEnd) Then n = n + 1
        End If
    Next i
    If n = 0 Then Exit Function

    Set products = NameMap("Products", "tblProducts", "ProductID", "ProductName")
    Set users = NameMap("Users", "tblUsers", "UserID", "Username")

    ReDim result(1 To n, 1 To 8)
    For i = UBound(data, 1) To 1 Step -1                 ' newest first
        If Len(CStr(data(i, colProduct))) > 0 Then
            If InRange(data(i, colDate), rangeStart, rangeEnd) Then
                k = k + 1
                productID = CStr(data(i, colProduct))
                userID = CStr(data(i, colUser))

                result(k, 1) = CDate(data(i, colDate))
                result(k, 2) = productID
                result(k, 3) = NameOf(products, productID)
                result(k, 4) = CStr(data(i, colType))
                result(k, 5) = ToNumber(data(i, colQty))
                result(k, 6) = ToNumber(data(i, colAfter))
                result(k, 7) = CStr(data(i, colRef))
                result(k, 8) = NameOf(users, userID)
            End If
        End If
    Next i

    summary = n & " movement(s), newest first"
    StockMovements = result
End Function

' ============================================================
' Financial report
' ============================================================

Private Function ProfitAndLoss(ByVal fromDate As Date, ByVal toDate As Date, _
                               ByRef title As String, ByRef headers As Variant, ByRef formats As Variant, _
                               ByRef widths As String, ByRef summary As String) As Variant
    Dim r As modProfit.TProfit
    Dim result(1 To 7, 1 To 2) As Variant
    Dim MARGIN As Double

    title = "Financial - Profit and loss"
    headers = Array("Item", "Amount")
    formats = Array("@", "#,##0.00")
    widths = "420 pt;160 pt"

    r = modProfit.CalculateProfit(fromDate, toDate)

    If r.revenue <> 0 Then MARGIN = Round(r.GrossProfit / r.revenue * 100, 2)

    result(1, 1) = "Revenue (sales after discounts, excluding tax)": result(1, 2) = r.revenue
    result(2, 1) = "Cost of goods sold":                             result(2, 2) = r.CostOfGoods
    result(3, 1) = "GROSS PROFIT":                                   result(3, 2) = r.GrossProfit
    result(4, 1) = "Gross margin (%)":                               result(4, 2) = MARGIN
    result(5, 1) = "Expenses":                                       result(5, 2) = r.Expenses
    result(6, 1) = "NET PROFIT":                                     result(6, 2) = r.netProfit
    result(7, 1) = "Tax collected (not part of profit)":             result(7, 2) = r.TaxCollected

    summary = r.SalesCount & " invoice(s) in this period"
    ProfitAndLoss = result
End Function

' ============================================================
' Display and export
' ============================================================

' Text for one value, using the column's format. Text stays text.
Public Function FormatReportValue(ByVal v As Variant, ByVal fmt As String) As String
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    If VarType(v) = vbString Then
        FormatReportValue = CStr(v)
    ElseIf Len(fmt) = 0 Or fmt = "@" Then
        FormatReportValue = CStr(v)
    ElseIf VarType(v) = vbDate Then
        FormatReportValue = Format$(v, fmt)
    ElseIf IsNumeric(v) Then
        FormatReportValue = Format$(v, fmt)
    Else
        FormatReportValue = CStr(v)
    End If
End Function

' Turns the report array into text, ready for a ListBox (0-based).
Public Function ToDisplayArray(ByVal data As Variant, ByVal formats As Variant) As String()
    Dim result() As String
    Dim RowCount As Long, colCount As Long
    Dim i As Long, j As Long

    RowCount = UBound(data, 1)
    colCount = UBound(data, 2)
    ReDim result(0 To RowCount - 1, 0 To colCount - 1)

    For i = 1 To RowCount
        For j = 1 To colCount
            result(i - 1, j - 1) = FormatReportValue(data(i, j), CStr(formats(j - 1)))
        Next j
    Next i

    ToDisplayArray = result
End Function

' Writes the report onto the Reports sheet. Returns the number of data rows written.
Public Function WriteReportToSheet(ByVal reportKey As String, ByVal fromDate As Date, _
                                   ByVal toDate As Date) As Long
    Dim ws As Worksheet
    Dim title As String, widths As String, summary As String
    Dim headers As Variant, formats As Variant, data As Variant
    Dim colCount As Long, n As Long, j As Long
    Dim numberFormat As String
    Dim subtitle As String

    data = BuildReport(reportKey, fromDate, toDate, title, headers, formats, widths, summary)
    colCount = UBound(headers) + 1

    Set ws = ThisWorkbook.Worksheets("Reports")
    ws.Cells.Clear

    With ws.Range("A1")
        .value = UCase$(title)
        .Font.Size = 16
        .Font.Bold = True
    End With

    subtitle = "Generated " & Format$(Now, "dd-mm-yyyy hh:nn") & " by " & modMain.CurrentUsername
    If UsesDateRange(reportKey) Then
        subtitle = "Period: " & Format$(fromDate, "dd-mm-yyyy") & " to " & Format$(toDate, "dd-mm-yyyy") & _
                   "     " & subtitle
    End If
    ws.Range("A2").value = subtitle

    ' Header row
    ws.Range(ws.Cells(4, 1), ws.Cells(4, colCount)).value = headers
    With ws.Range(ws.Cells(4, 1), ws.Cells(4, colCount))
        .Interior.Color = ThemeNavy()
        .Font.Color = ThemeWhite()
        .Font.Bold = True
    End With

    If IsEmpty(data) Then
        ws.Cells(5, 1).value = "No data for this report and period."
        ws.Range(ws.Cells(4, 1), ws.Cells(4, colCount)).Columns.AutoFit
        Exit Function
    End If

    n = UBound(data, 1)

    ' Set each column's format BEFORE writing so text stays text
    For j = 1 To colCount
        numberFormat = CStr(formats(j - 1))
        If numberFormat = "General Number" Then numberFormat = "General"
        ws.Range(ws.Cells(5, j), ws.Cells(4 + n, j)).numberFormat = numberFormat
    Next j

    ws.Range(ws.Cells(5, 1), ws.Cells(4 + n, colCount)).value = data

    With ws.Range(ws.Cells(4, 1), ws.Cells(4 + n, colCount))
        .Borders.LineStyle = xlContinuous
    End With

    With ws.Cells(6 + n, 1)
        .value = summary
        .Font.Bold = True
    End With

    ws.Range(ws.Cells(4, 1), ws.Cells(4 + n, colCount)).Columns.AutoFit

    WriteReportToSheet = n
End Function

