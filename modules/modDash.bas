Attribute VB_Name = "modDash"
Option Explicit

' ============================================================
' modDash - the analytics page on the Dashboard sheet:
' filters, calculations, KPI cards, tables and the chart data.
' Uses Scripting.Dictionary (Windows Excel).
' ============================================================

Private Type TStats
    revenue As Double
    CostOfGoods As Double
    GrossProfit As Double
    Units As Double
    Orders As Long
End Type

' ---------- The current filters (they reset to "all" whenever the workbook is reopened) ----------
Private mProductID As String            ' "" = all products
Private mProductName As String
Private mCategory As String             ' "" = all categories
Private mPeriodKey As String            ' "" or "ALL" = all time
Private mCustomFrom As Date
Private mCustomTo As Date
Private mUnlocked As Boolean            ' True after an administrator or manager has logged in

' ---------- Results of the latest analysis ----------
Private mFromDate As Date
Private mToDate As Date
Private mBucketMode As String
Private mBucketCount As Long
Private mBucketLabel() As String
Private mBucketRev() As Double
Private mBucketCost() As Double
Private mBucketBuy() As Double

Private mProdCat As Object              ' ProductID -> category
Private mProdName As Object             ' ProductID -> name
Private mCustName As Object             ' CustomerID -> name
Private mByProduct As Object            ' ProductID -> revenue
Private mProdUnits As Object            ' ProductID -> units
Private mProdCost As Object             ' ProductID -> cost
Private mByCategory As Object           ' category -> revenue
Private mByCustomer As Object           ' CustomerID -> revenue
Private mCustOrders As Object           ' CustomerID -> number of orders
Private mCustUnits As Object            ' CustomerID -> units
Private mWeekday(1 To 7) As Double      ' revenue by weekday, Monday = 1

Private mCur As TStats                  ' the chosen period
Private mPrev As TStats                 ' the previous period of the same length
Private mHavePrev As Boolean

' The sale lines of the chosen product/category (for the "latest sales" table)
Private mLineCount As Long
Private mLineDate() As Date
Private mLineInvoice() As String
Private mLineCustomer() As String
Private mLineQty() As Double
Private mLineTotal() As Double
Private mLineProfit() As Double

' ============================================================
' Filters
' ============================================================
Public Sub SetProductFilter(ByVal productID As String, ByVal productName As String)
    mProductID = productID
    mProductName = productName
    mCategory = ""
End Sub

Public Sub SetCategoryFilter(ByVal categoryName As String)
    mCategory = categoryName
    mProductID = ""
    mProductName = ""
End Sub

Public Sub SetPeriodFilter(ByVal key As String, ByVal customFrom As Date, ByVal customTo As Date)
    mPeriodKey = key
    mCustomFrom = customFrom
    mCustomTo = customTo
End Sub

Private Function PeriodKeyNow() As String
    If Len(mPeriodKey) = 0 Then
        PeriodKeyNow = "ALL"
    Else
        PeriodKeyNow = mPeriodKey
    End If
End Function

Public Function CurrentPeriodKey() As String
    CurrentPeriodKey = PeriodKeyNow()
End Function

' ---------- The list of periods ----------
Public Function PeriodCount() As Long
    PeriodCount = 10
End Function

Public Function PeriodKeyAt(ByVal index As Long) As String
    Dim keys As Variant
    keys = Array("ALL", "TODAY", "7D", "30D", "90D", "MTD", "LASTM", "12M", "YTD", "CUSTOM")
    PeriodKeyAt = CStr(keys(index))
End Function

Public Function PeriodLabelFor(ByVal key As String) As String
    Select Case key
        Case "ALL": PeriodLabelFor = "All time"
        Case "TODAY": PeriodLabelFor = "Today"
        Case "7D": PeriodLabelFor = "Last 7 days"
        Case "30D": PeriodLabelFor = "Last 30 days"
        Case "90D": PeriodLabelFor = "Last 90 days"
        Case "MTD": PeriodLabelFor = "This month"
        Case "LASTM": PeriodLabelFor = "Last month"
        Case "12M": PeriodLabelFor = "Last 12 months"
        Case "YTD": PeriodLabelFor = "This year"
        Case "CUSTOM": PeriodLabelFor = "Custom range"
        Case Else: PeriodLabelFor = "All time"
    End Select
End Function

' ============================================================
' Small helpers
' ============================================================
Private Function TableValues(ByVal sheetName As String, ByVal tableName As String, _
                             ByRef tbl As ListObject) As Variant
    Set tbl = modDatabase.GetTable(sheetName, tableName)
    If tbl.DataBodyRange Is Nothing Then
        TableValues = Empty
    Else
        TableValues = tbl.DataBodyRange.value
    End If
End Function

Private Function WeekStart(ByVal d As Date) As Date
    WeekStart = DateValue(d) - (Weekday(d, vbMonday) - 1)
End Function

' Works out the first and last day of the chosen period.
Private Sub ResolvePeriod(ByRef fromD As Date, ByRef toD As Date)
    toD = Date

    Select Case PeriodKeyNow()
        Case "TODAY"
            fromD = Date
        Case "7D"
            fromD = Date - 6
        Case "30D"
            fromD = Date - 29
        Case "90D"
            fromD = Date - 89
        Case "MTD"
            fromD = DateSerial(Year(Date), Month(Date), 1)
        Case "LASTM"
            fromD = DateSerial(Year(Date), Month(Date) - 1, 1)
            toD = DateSerial(Year(Date), Month(Date), 0)
        Case "12M"
            fromD = DateSerial(Year(Date), Month(Date) - 11, 1)
        Case "YTD"
            fromD = DateSerial(Year(Date), 1, 1)
        Case "CUSTOM"
            fromD = mCustomFrom
            toD = mCustomTo
        Case Else
            fromD = FirstSaleDate()
    End Select
End Sub

' The date of the oldest sale (or 30 days ago when there are no sales yet).
Private Function FirstSaleDate() As Date
    Dim tbl As ListObject
    Dim data As Variant
    Dim colDate As Long, i As Long
    Dim best As Date
    Dim found As Boolean

    data = TableValues("Sales", "tblSales", tbl)
    If Not IsEmpty(data) Then
        colDate = tbl.ListColumns("SaleDate").index
        For i = 1 To UBound(data, 1)
            If IsDate(data(i, colDate)) Then
                If Not found Then
                    best = CDate(data(i, colDate))
                    found = True
                ElseIf CDate(data(i, colDate)) < best Then
                    best = CDate(data(i, colDate))
                End If
            End If
        Next i
    End If

    If found Then
        FirstSaleDate = DateValue(best)
    Else
        FirstSaleDate = Date - 29
    End If
End Function

' Product, category and customer lookups.
Private Sub LoadMaps()
    Dim tbl As ListObject
    Dim data As Variant
    Dim i As Long
    Dim colID As Long, colName As Long, colCat As Long

    Set mProdCat = CreateObject("Scripting.Dictionary")
    Set mProdName = CreateObject("Scripting.Dictionary")
    Set mCustName = CreateObject("Scripting.Dictionary")

    data = TableValues("Products", "tblProducts", tbl)
    If Not IsEmpty(data) Then
        colID = tbl.ListColumns("ProductID").index
        colName = tbl.ListColumns("ProductName").index
        colCat = tbl.ListColumns("Category").index
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, colID))) > 0 Then
                mProdName(CStr(data(i, colID))) = CStr(data(i, colName))
                mProdCat(CStr(data(i, colID))) = CStr(data(i, colCat))
            End If
        Next i
    End If

    data = TableValues("Customers", "tblCustomers", tbl)
    If Not IsEmpty(data) Then
        colID = tbl.ListColumns("CustomerID").index
        colName = tbl.ListColumns("CustomerName").index
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, colID))) > 0 Then mCustName(CStr(data(i, colID))) = CStr(data(i, colName))
        Next i
    End If
End Sub

Private Function CategoryOf(ByVal pid As String) As String
    Dim cat As String

    If mProdCat.Exists(pid) Then cat = Trim$(CStr(mProdCat(pid)))
    If Len(cat) = 0 Then cat = "(none)"
    CategoryOf = cat
End Function

Private Function ProductLabel(ByVal pid As String) As String
    If mProdName.Exists(pid) Then
        ProductLabel = CStr(mProdName(pid))
    Else
        ProductLabel = pid
    End If
End Function

Private Function CustomerLabel(ByVal cid As String) As String
    If mCustName.Exists(cid) Then
        CustomerLabel = CStr(mCustName(cid))
    Else
        CustomerLabel = cid
    End If
End Function

Private Function ShortText(ByVal s As String, ByVal maxLen As Long) As String
    If Len(s) > maxLen Then
        ShortText = Left$(s, maxLen - 2) & ".."
    Else
        ShortText = s
    End If
End Function

' Does a sale line belong to the chosen product / category?
Private Function LineMatches(ByVal pid As String, ByVal productID As String, _
                             ByVal category As String) As Boolean
    If Len(productID) > 0 And pid <> productID Then Exit Function
    If Len(category) > 0 Then
        If StrComp(CategoryOf(pid), category, vbTextCompare) <> 0 Then Exit Function
    End If
    LineMatches = True
End Function

' ---------- Time buckets (days, weeks, months or years, depending on the length) ----------
Private Sub PrepareBuckets(ByVal fromD As Date, ByVal toD As Date)
    Dim span As Long
    Dim i As Long

    span = CLng(DateValue(toD) - DateValue(fromD)) + 1

    If span <= 31 Then
        mBucketMode = "DAY"
        mBucketCount = span
    ElseIf span <= 120 Then
        mBucketMode = "WEEK"
        mBucketCount = CLng((WeekStart(toD) - WeekStart(fromD)) / 7) + 1
    ElseIf span <= 1100 Then
        mBucketMode = "MONTH"
        mBucketCount = (Year(toD) - Year(fromD)) * 12 + Month(toD) - Month(fromD) + 1
    Else
        mBucketMode = "YEAR"
        mBucketCount = Year(toD) - Year(fromD) + 1
    End If
    If mBucketCount < 1 Then mBucketCount = 1

    ReDim mBucketLabel(1 To mBucketCount)
    ReDim mBucketRev(1 To mBucketCount)
    ReDim mBucketCost(1 To mBucketCount)
    ReDim mBucketBuy(1 To mBucketCount)

    For i = 1 To mBucketCount
        Select Case mBucketMode
            Case "DAY"
                mBucketLabel(i) = Format$(DateValue(fromD) + i - 1, "dd mmm")
            Case "WEEK"
                mBucketLabel(i) = Format$(WeekStart(fromD) + (i - 1) * 7, "dd mmm")
            Case "MONTH"
                mBucketLabel(i) = Format$(DateSerial(Year(fromD), Month(fromD) + i - 1, 1), "mmm yy")
            Case Else
                mBucketLabel(i) = CStr(Year(fromD) + i - 1)
        End Select
    Next i
End Sub

' Which bucket a date falls into (0 = outside the period).
Private Function BucketIndex(ByVal d As Date) As Long
    Dim idx As Long

    Select Case mBucketMode
        Case "DAY"
            idx = CLng(DateValue(d) - DateValue(mFromDate)) + 1
        Case "WEEK"
            idx = CLng((WeekStart(d) - WeekStart(mFromDate)) / 7) + 1
        Case "MONTH"
            idx = (Year(d) - Year(mFromDate)) * 12 + Month(d) - Month(mFromDate) + 1
        Case Else
            idx = Year(d) - Year(mFromDate) + 1
    End Select

    If idx >= 1 And idx <= mBucketCount Then BucketIndex = idx
End Function

' Gets ready to collect the details of the chosen period.
Private Sub InitDetail()
    Dim i As Long

    Set mByProduct = CreateObject("Scripting.Dictionary")
    Set mProdUnits = CreateObject("Scripting.Dictionary")
    Set mProdCost = CreateObject("Scripting.Dictionary")
    Set mByCategory = CreateObject("Scripting.Dictionary")
    Set mByCustomer = CreateObject("Scripting.Dictionary")
    Set mCustOrders = CreateObject("Scripting.Dictionary")
    Set mCustUnits = CreateObject("Scripting.Dictionary")

    For i = 1 To 7
        mWeekday(i) = 0
    Next i

    mLineCount = 0
    ReDim mLineDate(1 To 100)
    ReDim mLineInvoice(1 To 100)
    ReDim mLineCustomer(1 To 100)
    ReDim mLineQty(1 To 100)
    ReDim mLineTotal(1 To 100)
    ReDim mLineProfit(1 To 100)
End Sub

Private Sub AddLine(ByVal d As Date, ByVal invoice As String, ByVal customer As String, _
                    ByVal qty As Double, ByVal total As Double, ByVal profit As Double)
    mLineCount = mLineCount + 1

    If mLineCount > UBound(mLineDate) Then
        ReDim Preserve mLineDate(1 To mLineCount * 2)
        ReDim Preserve mLineInvoice(1 To mLineCount * 2)
        ReDim Preserve mLineCustomer(1 To mLineCount * 2)
        ReDim Preserve mLineQty(1 To mLineCount * 2)
        ReDim Preserve mLineTotal(1 To mLineCount * 2)
        ReDim Preserve mLineProfit(1 To mLineCount * 2)
    End If

    mLineDate(mLineCount) = d
    mLineInvoice(mLineCount) = invoice
    mLineCustomer(mLineCount) = customer
    mLineQty(mLineCount) = qty
    mLineTotal(mLineCount) = total
    mLineProfit(mLineCount) = profit
End Sub

' ============================================================
' The calculation
' ============================================================
' Analyses the sales between fromD and toD. productID / category "" = no filter.
' wantDetail False = only the totals (used for the previous period).
Private Sub Analyze(ByVal fromD As Date, ByVal toD As Date, ByVal productID As String, _
                    ByVal category As String, ByVal wantDetail As Boolean, ByRef stats As TStats)
    Dim zeroStats As TStats
    Dim tblS As ListObject, tblD As ListObject, tblP As ListObject, tblPD As ListObject
    Dim sales As Variant, saleLines As Variant, purchases As Variant, purLines As Variant
    Dim saleDate As Object, saleCust As Object, saleInvoice As Object, seenSale As Object, purDate As Object
    Dim colSaleID As Long, colCust As Long, colDate As Long, colSub As Long, colDisc As Long, colInvoice As Long
    Dim colLSale As Long, colLProd As Long, colLQty As Long, colLCost As Long, colLTotal As Long
    Dim colPID As Long, colPDate As Long, colPCost As Long
    Dim colPLPur As Long, colPLProd As Long, colPLTotal As Long
    Dim rangeStart As Date, rangeEnd As Date, dt As Date
    Dim i As Long, b As Long
    Dim sid As String, pid As String, cust As String, cat As String
    Dim qty As Double, lineRev As Double, lineCost As Double, saleRev As Double
    Dim lineMode As Boolean

    stats = zeroStats
    lineMode = (Len(productID) > 0 Or Len(category) > 0)
    rangeStart = DateValue(fromD)
    rangeEnd = DateValue(toD) + 1

    Set saleDate = CreateObject("Scripting.Dictionary")
    Set saleCust = CreateObject("Scripting.Dictionary")
    Set saleInvoice = CreateObject("Scripting.Dictionary")
    Set seenSale = CreateObject("Scripting.Dictionary")

    ' ---- 1. The sales in the period ----
    sales = TableValues("Sales", "tblSales", tblS)
    If Not IsEmpty(sales) Then
        colSaleID = tblS.ListColumns("SaleID").index
        colCust = tblS.ListColumns("CustomerID").index
        colDate = tblS.ListColumns("SaleDate").index
        colSub = tblS.ListColumns("Subtotal").index
        colDisc = tblS.ListColumns("Discount").index
        colInvoice = tblS.ListColumns("InvoiceNo").index

        For i = 1 To UBound(sales, 1)
            If Len(CStr(sales(i, colSaleID))) > 0 And IsDate(sales(i, colDate)) Then
                dt = CDate(sales(i, colDate))
                If dt >= rangeStart And dt < rangeEnd Then
                    sid = CStr(sales(i, colSaleID))
                    saleDate(sid) = dt
                    saleCust(sid) = CStr(sales(i, colCust))
                    saleInvoice(sid) = CStr(sales(i, colInvoice))

                    If Not lineMode Then
                        saleRev = ToNumber(sales(i, colSub)) - ToNumber(sales(i, colDisc))
                        stats.revenue = stats.revenue + saleRev
                        stats.Orders = stats.Orders + 1
                        If wantDetail Then
                            b = BucketIndex(dt)
                            If b > 0 Then mBucketRev(b) = mBucketRev(b) + saleRev
                        End If
                    End If
                End If
            End If
        Next i
    End If

    ' ---- 2. The lines of those sales ----
    saleLines = TableValues("SaleDetails", "tblSaleDetails", tblD)
    If Not IsEmpty(saleLines) And saleDate.count > 0 Then
        colLSale = tblD.ListColumns("SaleID").index
        colLProd = tblD.ListColumns("ProductID").index
        colLQty = tblD.ListColumns("Quantity").index
        colLCost = tblD.ListColumns("CostPrice").index
        colLTotal = tblD.ListColumns("Total").index

        For i = 1 To UBound(saleLines, 1)
            sid = CStr(saleLines(i, colLSale))
            If saleDate.Exists(sid) Then
                pid = CStr(saleLines(i, colLProd))
                If LineMatches(pid, productID, category) Then
                    qty = ToNumber(saleLines(i, colLQty))
                    lineRev = ToNumber(saleLines(i, colLTotal))
                    lineCost = qty * ToNumber(saleLines(i, colLCost))
                    dt = CDate(saleDate(sid))
                    cust = CStr(saleCust(sid))
                    cat = CategoryOf(pid)

                    stats.CostOfGoods = stats.CostOfGoods + lineCost
                    stats.Units = stats.Units + qty
                    If lineMode Then stats.revenue = stats.revenue + lineRev

                    ' Count each sale once
                    If Not seenSale.Exists(sid) Then
                        seenSale(sid) = True
                        If lineMode Then stats.Orders = stats.Orders + 1
                        If wantDetail Then mCustOrders(cust) = mCustOrders(cust) + 1
                    End If

                    If wantDetail Then
                        mByProduct(pid) = mByProduct(pid) + lineRev
                        mProdUnits(pid) = mProdUnits(pid) + qty
                        mProdCost(pid) = mProdCost(pid) + lineCost
                        mByCategory(cat) = mByCategory(cat) + lineRev
                        mByCustomer(cust) = mByCustomer(cust) + lineRev
                        mCustUnits(cust) = mCustUnits(cust) + qty
                        mWeekday(Weekday(dt, vbMonday)) = mWeekday(Weekday(dt, vbMonday)) + lineRev

                        b = BucketIndex(dt)
                        If b > 0 Then
                            mBucketCost(b) = mBucketCost(b) + lineCost
                            If lineMode Then mBucketRev(b) = mBucketRev(b) + lineRev
                        End If

                        If lineMode Then AddLine dt, CStr(saleInvoice(sid)), cust, qty, lineRev, lineRev - lineCost
                    End If
                End If
            End If
        Next i
    End If

    ' ---- 3. Purchases (only needed for the charts) ----
    If wantDetail Then
        purchases = TableValues("Purchases", "tblPurchases", tblP)
        If Not IsEmpty(purchases) Then
            Set purDate = CreateObject("Scripting.Dictionary")
            colPID = tblP.ListColumns("PurchaseID").index
            colPDate = tblP.ListColumns("PurchaseDate").index
            colPCost = tblP.ListColumns("TotalCost").index

            For i = 1 To UBound(purchases, 1)
                If Len(CStr(purchases(i, colPID))) > 0 And IsDate(purchases(i, colPDate)) Then
                    dt = CDate(purchases(i, colPDate))
                    If dt >= rangeStart And dt < rangeEnd Then
                        purDate(CStr(purchases(i, colPID))) = dt
                        If Not lineMode Then
                            b = BucketIndex(dt)
                            If b > 0 Then mBucketBuy(b) = mBucketBuy(b) + ToNumber(purchases(i, colPCost))
                        End If
                    End If
                End If
            Next i

            If lineMode And purDate.count > 0 Then
                purLines = TableValues("PurchaseDetails", "tblPurchaseDetails", tblPD)
                If Not IsEmpty(purLines) Then
                    colPLPur = tblPD.ListColumns("PurchaseID").index
                    colPLProd = tblPD.ListColumns("ProductID").index
                    colPLTotal = tblPD.ListColumns("Total").index

                    For i = 1 To UBound(purLines, 1)
                        sid = CStr(purLines(i, colPLPur))
                        If purDate.Exists(sid) Then
                            If LineMatches(CStr(purLines(i, colPLProd)), productID, category) Then
                                b = BucketIndex(CDate(purDate(sid)))
                                If b > 0 Then mBucketBuy(b) = mBucketBuy(b) + ToNumber(purLines(i, colPLTotal))
                            End If
                        End If
                    Next i
                End If
            End If
        End If
    End If

    stats.GrossProfit = stats.revenue - stats.CostOfGoods
End Sub

' ============================================================
' Sorting helpers
' ============================================================
' The dictionary's keys sorted by value, biggest first (at most maxItems).
Private Function TopKeys(ByVal d As Object, ByVal maxItems As Long) As Variant
    Dim keys As Variant
    Dim i As Long, j As Long, n As Long
    Dim tmp As Variant
    Dim result() As String

    If d.count = 0 Then
        TopKeys = Array()
        Exit Function
    End If

    keys = d.keys
    For i = LBound(keys) + 1 To UBound(keys)
        tmp = keys(i)
        j = i - 1
        Do While j >= LBound(keys)
            If d(CStr(keys(j))) >= d(CStr(tmp)) Then Exit Do
            keys(j + 1) = keys(j)
            j = j - 1
        Loop
        keys(j + 1) = tmp
    Next i

    n = d.count
    If n > maxItems Then n = maxItems

    ReDim result(0 To n - 1)
    For i = 0 To n - 1
        result(i) = CStr(keys(i))
    Next i
    TopKeys = result
End Function

' The positions of the newest sale lines, newest first.
Private Function LatestLineIndexes(ByVal maxItems As Long) As Variant
    Dim used() As Boolean
    Dim result() As Long
    Dim k As Long, i As Long, best As Long, n As Long

    n = maxItems
    If mLineCount < n Then n = mLineCount
    If n <= 0 Then
        LatestLineIndexes = Array()
        Exit Function
    End If

    ReDim used(1 To mLineCount)
    ReDim result(1 To n)

    For k = 1 To n
        best = 0
        For i = 1 To mLineCount
            If Not used(i) Then
                If best = 0 Then
                    best = i
                ElseIf mLineDate(i) > mLineDate(best) Then
                    best = i
                End If
            End If
        Next i
        used(best) = True
        result(k) = best
    Next k

    LatestLineIndexes = result
End Function

' ============================================================
' Writing the results
' ============================================================
' Writes one block (labels + values) onto the helper sheet.
Private Sub WriteBlock(ByVal wd As Worksheet, ByVal firstCol As Long, ByVal n As Long, _
                       ByRef labels() As String, ByRef vals() As Double)
    Dim out() As Variant
    Dim i As Long

    ReDim out(1 To n, 1 To 2)
    For i = 1 To n
        out(i, 1) = labels(i)
        out(i, 2) = vals(i)
    Next i

    wd.Cells(2, firstCol).Resize(n, 2).value = out
    wd.Cells(1, firstCol + 1).value = "Value"
End Sub

Private Sub WriteChartData(ByRef nTrend As Long, ByRef nDonut As Long, ByRef nBar As Long, _
                           ByRef nStock As Long, ByRef donutTitle As String, _
                           ByRef barTitle As String, ByRef stockTitle As String)
    Dim wd As Worksheet
    Dim trend() As Variant, buy() As Variant
    Dim labels(1 To 10) As String
    Dim vals(1 To 10) As Double
    Dim d As Object
    Dim keys As Variant
    Dim k As Variant
    Dim i As Long, n As Long
    Dim total As Double, shown As Double
    Dim stockData As Variant
    Dim tblP As ListObject
    Dim colID As Long, colCat As Long, colStock As Long, colMin As Long, colStatus As Long
    Dim p As modProducts.TProduct
    Dim cIn As Long, cLow As Long, cOut As Long, cInactive As Long
    Dim stockQty As Double
    Dim dayNames As Variant

    Set wd = modDashBuild.GetDataSheet()
    wd.Cells.ClearContents
    wd.Range("A:A,F:F,J:J,M:M,P:P").numberFormat = "@"

    ' ---- Trend (A:D) and sales vs purchases (F:H) ----
    nTrend = mBucketCount
    ReDim trend(1 To nTrend, 1 To 4)
    ReDim buy(1 To nTrend, 1 To 3)
    For i = 1 To nTrend
        trend(i, 1) = mBucketLabel(i)
        trend(i, 2) = Round(mBucketRev(i), 2)
        trend(i, 3) = Round(mBucketRev(i) - mBucketCost(i), 2)
        trend(i, 4) = Round(mBucketBuy(i), 2)
        buy(i, 1) = mBucketLabel(i)
        buy(i, 2) = Round(mBucketRev(i), 2)
        buy(i, 3) = Round(mBucketBuy(i), 2)
    Next i
    wd.Range("A2").Resize(nTrend, 4).value = trend
    wd.Range("B1").value = "Revenue"
    wd.Range("C1").value = "Gross profit"
    wd.Range("D1").value = "Purchases"
    wd.Range("F2").Resize(nTrend, 3).value = buy
    wd.Range("G1").value = "Sales"
    wd.Range("H1").value = "Purchases"

    ' ---- Donut (J:K): category share, product share or customer share ----
    If Len(mProductID) > 0 Then
        Set d = mByCustomer
        donutTitle = "WHO BUYS THIS PRODUCT"
    ElseIf Len(mCategory) > 0 Then
        Set d = mByProduct
        donutTitle = "PRODUCTS IN THIS CATEGORY"
    Else
        Set d = mByCategory
        donutTitle = "REVENUE BY CATEGORY"
    End If

    total = 0
    For Each k In d.keys
        total = total + d(k)
    Next k

    n = 0
    shown = 0
    If d.count > 0 Then
        keys = TopKeys(d, 6)
        For i = LBound(keys) To UBound(keys)
            n = n + 1
            If Len(mProductID) > 0 Then
                labels(n) = ShortText(CustomerLabel(CStr(keys(i))), 22)
            ElseIf Len(mCategory) > 0 Then
                labels(n) = ShortText(ProductLabel(CStr(keys(i))), 22)
            Else
                labels(n) = ShortText(CStr(keys(i)), 22)
            End If
            vals(n) = Round(d(CStr(keys(i))), 2)
            shown = shown + vals(n)
        Next i
    End If
    If total - shown > 0.5 Then
        n = n + 1
        labels(n) = "Other"
        vals(n) = Round(total - shown, 2)
    End If
    If n = 0 Then
        n = 1
        labels(1) = "No sales"
        vals(1) = 0
    End If
    nDonut = n
    WriteBlock wd, 10, n, labels, vals

    ' ---- Bars (M:N): top products, or the weekdays of one product ----
    n = 0
    If Len(mProductID) > 0 Then
        barTitle = "SALES BY WEEKDAY"
        dayNames = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
        For i = 1 To 7
            n = n + 1
            labels(n) = CStr(dayNames(i - 1))
            vals(n) = Round(mWeekday(i), 2)
        Next i
    Else
        If Len(mCategory) > 0 Then
            barTitle = "TOP PRODUCTS IN THIS CATEGORY"
        Else
            barTitle = "TOP PRODUCTS BY REVENUE"
        End If
        If mByProduct.count > 0 Then
            keys = TopKeys(mByProduct, 8)
            For i = LBound(keys) To UBound(keys)
                n = n + 1
                labels(n) = ShortText(ProductLabel(CStr(keys(i))), 20)
                vals(n) = Round(mByProduct(CStr(keys(i))), 2)
            Next i
        End If
    End If
    If n = 0 Then
        n = 1
        labels(1) = "No sales"
        vals(1) = 0
    End If
    nBar = n
    WriteBlock wd, 13, n, labels, vals

    ' ---- Stock pie (P:Q) ----
    n = 0
    If Len(mProductID) > 0 Then
        stockTitle = "STOCK LEVEL OF THIS PRODUCT"
        If modProducts.LoadProduct(mProductID, p) Then
            If p.stock > 0 Then
                n = n + 1
                labels(n) = "Stock (units)"
                vals(n) = p.stock
            End If
            If p.MaxStock > p.stock Then
                n = n + 1
                labels(n) = "Free space to max"
                vals(n) = p.MaxStock - p.stock
            End If
        End If
    Else
        If Len(mCategory) > 0 Then
            stockTitle = "STOCK HEALTH IN THIS CATEGORY"
        Else
            stockTitle = "STOCK HEALTH (PRODUCTS)"
        End If

        stockData = TableValues("Products", "tblProducts", tblP)
        If Not IsEmpty(stockData) Then
            colID = tblP.ListColumns("ProductID").index
            colCat = tblP.ListColumns("Category").index
            colStock = tblP.ListColumns("Stock").index
            colMin = tblP.ListColumns("MinStock").index
            colStatus = tblP.ListColumns("Status").index

            For i = 1 To UBound(stockData, 1)
                If Len(CStr(stockData(i, colID))) > 0 Then
                    If Len(mCategory) = 0 Or StrComp(Trim$(CStr(stockData(i, colCat))), mCategory, vbTextCompare) = 0 Then
                        If CStr(stockData(i, colStatus)) = "Active" Then
                            stockQty = ToNumber(stockData(i, colStock))
                            Select Case modInventory.StockStatus(stockQty, ToNumber(stockData(i, colMin)))
                                Case "IN STOCK"
                                    cIn = cIn + 1
                                Case "LOW STOCK"
                                    cLow = cLow + 1
                                Case Else
                                    cOut = cOut + 1
                            End Select
                        Else
                            cInactive = cInactive + 1
                        End If
                    End If
                End If
            Next i
        End If

        If cIn > 0 Then
            n = n + 1
            labels(n) = "In stock"
            vals(n) = cIn
        End If
        If cLow > 0 Then
            n = n + 1
            labels(n) = "Low stock"
            vals(n) = cLow
        End If
        If cOut > 0 Then
            n = n + 1
            labels(n) = "Out of stock"
            vals(n) = cOut
        End If
        If cInactive > 0 Then
            n = n + 1
            labels(n) = "Inactive"
            vals(n) = cInactive
        End If
    End If
    If n = 0 Then
        n = 1
        labels(1) = "No data"
        vals(1) = 0
    End If
    nStock = n
    WriteBlock wd, 16, n, labels, vals
End Sub

' ---------- Header texts ----------
Private Sub WriteHeader(ByVal ws As Worksheet)
    Dim focusText As String

    If Len(mProductID) > 0 Then
        focusText = mProductName
    ElseIf Len(mCategory) > 0 Then
        focusText = "Category: " & mCategory
    Else
        focusText = "All products"
    End If

    ws.Range("B2").value = "Showing:  " & focusText & "   |   " & PeriodLabelFor(PeriodKeyNow()) & _
                           "   (" & Format$(mFromDate, "dd-mm-yyyy") & " to " & Format$(mToDate, "dd-mm-yyyy") & ")"
    ws.Range("B60").value = "Updated " & Format$(Now, "dd-mm-yyyy hh:nn")
End Sub

' ---------- KPI cards ----------
Private Function DeltaText(ByVal currentValue As Double, ByVal previousValue As Double) As String
    Dim pct As Double

    If Not mHavePrev Then
        DeltaText = "whole history"
    ElseIf previousValue = 0 Then
        If currentValue > 0 Then
            DeltaText = ChrW(9650) & " new"
        Else
            DeltaText = "no change"
        End If
    Else
        pct = (currentValue - previousValue) / Abs(previousValue) * 100
        If pct > 0.05 Then
            DeltaText = ChrW(9650) & " " & Format$(pct, "0.0") & "% vs previous"
        ElseIf pct < -0.05 Then
            DeltaText = ChrW(9660) & " " & Format$(Abs(pct), "0.0") & "% vs previous"
        Else
            DeltaText = "no change"
        End If
    End If
End Function

Private Sub SetCard(ByVal ws As Worksheet, ByVal index As Long, ByVal captionText As String, _
                    ByVal valueText As String, ByVal deltaString As String)
    Dim col As Long

    col = 2 * index
    ws.Cells(6, col).value = captionText
    ws.Cells(7, col).numberFormat = "@"
    ws.Cells(7, col).value = valueText
    ws.Cells(8, col).numberFormat = "@"
    ws.Cells(8, col).value = deltaString

    If Left$(deltaString, 1) = ChrW(9650) Then
        ws.Cells(8, col).Font.Color = RGB(46, 230, 166)
    ElseIf Left$(deltaString, 1) = ChrW(9660) Then
        ws.Cells(8, col).Font.Color = RGB(255, 99, 132)
    Else
        ws.Cells(8, col).Font.Color = RGB(150, 160, 190)
    End If
End Sub

Private Sub WriteKpis(ByVal ws As Worksheet)
    Dim sym As String
    Dim MARGIN As Double, prevMargin As Double
    Dim aov As Double, prevAov As Double
    Dim marginText As String
    Dim diff As Double

    sym = Trim$(GetSetting("CurrencySymbol", "Rs.")) & " "

    If mCur.revenue <> 0 Then MARGIN = mCur.GrossProfit / mCur.revenue * 100
    If mCur.Orders > 0 Then aov = mCur.revenue / mCur.Orders

    If mHavePrev Then
        If mPrev.revenue <> 0 Then prevMargin = mPrev.GrossProfit / mPrev.revenue * 100
        If mPrev.Orders > 0 Then prevAov = mPrev.revenue / mPrev.Orders

        diff = MARGIN - prevMargin
        If diff > 0.05 Then
            marginText = ChrW(9650) & " " & Format$(diff, "0.0") & " pts vs previous"
        ElseIf diff < -0.05 Then
            marginText = ChrW(9660) & " " & Format$(Abs(diff), "0.0") & " pts vs previous"
        Else
            marginText = "no change"
        End If
    Else
        marginText = "whole history"
    End If

    SetCard ws, 1, "REVENUE", sym & Format$(mCur.revenue, "#,##0"), DeltaText(mCur.revenue, mPrev.revenue)
    SetCard ws, 2, "GROSS PROFIT", sym & Format$(mCur.GrossProfit, "#,##0"), DeltaText(mCur.GrossProfit, mPrev.GrossProfit)
    SetCard ws, 3, "GROSS MARGIN", Format$(MARGIN, "0.0") & "%", marginText
    SetCard ws, 4, "UNITS SOLD", Format$(mCur.Units, "#,##0.##"), DeltaText(mCur.Units, mPrev.Units)
    SetCard ws, 5, "ORDERS", Format$(mCur.Orders, "#,##0"), DeltaText(CDbl(mCur.Orders), CDbl(mPrev.Orders))
    SetCard ws, 6, "AVERAGE ORDER", sym & Format$(aov, "#,##0"), DeltaText(aov, prevAov)
End Sub

' ---------- Tables ----------
Private Sub PutText(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal text As String)
    ws.Cells(r, c).numberFormat = "@"
    ws.Cells(r, c).value = text
End Sub

Private Sub WriteTables(ByVal ws As Worksheet)
    Dim keys As Variant
    Dim idx As Variant
    Dim i As Long, k As Long, r As Long
    Dim pid As String, cid As String
    Dim revenue As Double, cost As Double
    Dim sectionTitle As String

    ws.Range("B47:M56").ClearContents

    ' ===== Left table =====
    If Len(mProductID) > 0 Then
        sectionTitle = "LATEST SALES OF " & UCase$(ShortText(mProductName, 30))
        ws.Range("B47:B56").numberFormat = "dd-mm-yy"
        ws.Range("H46:H56").HorizontalAlignment = xlRight

        PutText ws, 46, 2, "DATE"
        PutText ws, 46, 3, "CUSTOMER"
        PutText ws, 46, 5, "QTY"
        PutText ws, 46, 6, "TOTAL"
        PutText ws, 46, 7, "PROFIT"
        PutText ws, 46, 8, "INVOICE"

        If mLineCount > 0 Then
            idx = LatestLineIndexes(10)
            r = 46
            For k = LBound(idx) To UBound(idx)
                r = r + 1
                i = CLng(idx(k))
                ws.Cells(r, 2).value = mLineDate(i)
                PutText ws, r, 3, ShortText(CustomerLabel(mLineCustomer(i)), 24)
                ws.Cells(r, 5).value = mLineQty(i)
                ws.Cells(r, 6).value = Round(mLineTotal(i), 2)
                ws.Cells(r, 7).value = Round(mLineProfit(i), 2)
                PutText ws, r, 8, mLineInvoice(i)
            Next k
        End If
    Else
        If Len(mCategory) > 0 Then
            sectionTitle = "TOP PRODUCTS IN " & UCase$(ShortText(mCategory, 26))
        Else
            sectionTitle = "TOP PRODUCTS"
        End If
        ws.Range("B47:B56").numberFormat = "General"
        ws.Range("H47:H56").numberFormat = "0.0%"

        PutText ws, 46, 2, "#"
        PutText ws, 46, 3, "PRODUCT"
        PutText ws, 46, 5, "UNITS"
        PutText ws, 46, 6, "REVENUE"
        PutText ws, 46, 7, "PROFIT"
        PutText ws, 46, 8, "MARGIN"

        If mByProduct.count > 0 Then
            keys = TopKeys(mByProduct, 10)
            r = 46
            For k = LBound(keys) To UBound(keys)
                r = r + 1
                pid = CStr(keys(k))
                revenue = mByProduct(pid)
                cost = mProdCost(pid)

                ws.Cells(r, 2).value = k - LBound(keys) + 1
                PutText ws, r, 3, ShortText(ProductLabel(pid), 24)
                ws.Cells(r, 5).value = mProdUnits(pid)
                ws.Cells(r, 6).value = Round(revenue, 2)
                ws.Cells(r, 7).value = Round(revenue - cost, 2)
                If revenue <> 0 Then ws.Cells(r, 8).value = (revenue - cost) / revenue
            Next k
        End If
    End If
    ws.Range("B45").value = sectionTitle

    ' ===== Right table: customers =====
    If Len(mProductID) > 0 Then
        ws.Range("I45").value = "WHO BUYS THIS PRODUCT"
    Else
        ws.Range("I45").value = "TOP CUSTOMERS"
    End If

    PutText ws, 46, 9, "CUSTOMER"
    PutText ws, 46, 11, "ORDERS"
    PutText ws, 46, 12, "UNITS"
    PutText ws, 46, 13, "REVENUE"

    If mByCustomer.count > 0 Then
        keys = TopKeys(mByCustomer, 10)
        r = 46
        For k = LBound(keys) To UBound(keys)
            r = r + 1
            cid = CStr(keys(k))
            PutText ws, r, 9, ShortText(CustomerLabel(cid), 24)
            ws.Cells(r, 11).value = mCustOrders(cid)
            ws.Cells(r, 12).value = mCustUnits(cid)
            ws.Cells(r, 13).value = Round(mByCustomer(cid), 2)
        Next k
    End If
End Sub

' ---------- Button captions ----------
Private Sub UpdateButtons()
    Dim productText As String, categoryText As String

    If Len(mProductID) > 0 Then
        productText = ShortText(mProductName, 26)
    Else
        productText = "All products"
    End If

    If Len(mCategory) > 0 Then
        categoryText = ShortText(mCategory, 14)
    Else
        categoryText = "All"
    End If

    modDashBuild.SetButtonCaption "dashBtnProduct", "PRODUCT:  " & productText & "  " & ChrW(9662)
    modDashBuild.SetButtonCaption "dashBtnCategory", "CATEGORY:  " & categoryText & "  " & ChrW(9662)
    modDashBuild.SetButtonCaption "dashBtnPeriod", "PERIOD:  " & PeriodLabelFor(PeriodKeyNow()) & "  " & ChrW(9662)
End Sub
Private Sub DoRefresh(ByVal showErrors As Boolean)
    Dim ws As Worksheet, wd As Worksheet
    Dim zeroStats As TStats
    Dim prevFrom As Date, prevTo As Date
    Dim span As Long
    Dim nTrend As Long, nDonut As Long, nBar As Long, nStock As Long
    Dim donutTitle As String, barTitle As String, stockTitle As String
    Dim errText As String
    Dim dashWasLocked As Boolean, dataWasLocked As Boolean

    On Error GoTo ErrorHandler

    If Not modDashBuild.ChartsExist() Then Exit Sub

    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Set wd = modDashBuild.GetDataSheet()

    ' Both sheets must be writable (they are protected in the read-only view mode)
    dashWasLocked = MakeWritable(ws)
    dataWasLocked = MakeWritable(wd)

    ' 1. The period
    ResolvePeriod mFromDate, mToDate
    If mFromDate > mToDate Then mFromDate = mToDate
    LoadMaps

    ' 2. The previous period of the same length (not for "all time")
    mPrev = zeroStats
    mHavePrev = (PeriodKeyNow() <> "ALL")
    If mHavePrev Then
        span = CLng(DateValue(mToDate) - DateValue(mFromDate)) + 1
        prevTo = DateValue(mFromDate) - 1
        prevFrom = prevTo - (span - 1)
        Analyze prevFrom, prevTo, mProductID, mCategory, False, mPrev
    End If

    ' 3. The chosen period, in detail
    PrepareBuckets mFromDate, mToDate
    InitDetail
    Analyze mFromDate, mToDate, mProductID, mCategory, True, mCur

    ' 4. Show everything
    WriteChartData nTrend, nDonut, nBar, nStock, donutTitle, barTitle, stockTitle
    WriteHeader ws
    WriteKpis ws
    WriteTables ws
    modDashBuild.UpdateCharts nTrend, nDonut, nBar, nStock, donutTitle, barTitle, stockTitle
    UpdateButtons

    ' 5. Lock the sheets again if they were locked
    If dashWasLocked Then ProtectAgain ws
    If dataWasLocked Then ProtectAgain wd

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    errText = Err.description

    On Error Resume Next
    If dashWasLocked Then ProtectAgain ws
    If dataWasLocked Then ProtectAgain wd
    Application.ScreenUpdating = True
    On Error GoTo 0

    If showErrors Then
        If InStr(1, errText, "protected", vbTextCompare) > 0 Then
            errText = errText & vbCrLf & vbCrLf & _
                      "A sheet has its own password. Unprotect the Dashboard and DashData sheets " & _
                      "(Review, Unprotect Sheet), then try again."
        End If
        MsgBox "The dashboard could not be refreshed." & vbCrLf & errText, vbExclamation, "Dashboard"
    End If
End Sub

' Refreshes the analytics page (builds it first if it does not exist yet).
Public Sub RefreshDashboardSheet()
    If Not mUnlocked Then
        MsgBox "The analytics are locked." & vbCrLf & _
               "Click OPEN APPLICATION and log in as an administrator or manager first.", _
               vbInformation, "Dashboard"
        Exit Sub
    End If

    If Not modDashBuild.ChartsExist() Then
        modDashBuild.BuildDashboardSheet
    Else
        DoRefresh True
    End If
End Sub

' Shows or hides the privacy cover.
Public Sub SetUnlocked(ByVal value As Boolean)
    Dim ws As Worksheet
    Dim wasLocked As Boolean

    mUnlocked = value

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    wasLocked = MakeWritable(ws)

    If value Then
        ws.Shapes("dashCover").Visible = msoFalse
    Else
        ws.Shapes("dashCover").Visible = msoTrue
    End If

    If wasLocked Then ProtectAgain ws
    On Error GoTo 0
End Sub

' Called when the application stops: administrators and managers see fresh analytics,
' everybody else sees the cover.
Public Sub AfterSession()
    On Error Resume Next
    If modMain.IsLoggedIn() Then
        If modMain.CurrentRole = "ADMIN" Or modMain.CurrentRole = "MANAGER" Then
            SetUnlocked True
            DoRefresh False
            Exit Sub
        End If
    End If
    SetUnlocked False
    On Error GoTo 0
End Sub

' ============================================================
' The macros behind the sheet buttons
' ============================================================
Private Function CheckUnlocked() As Boolean
    If mUnlocked Then
        CheckUnlocked = True
    Else
        MsgBox "The analytics are locked." & vbCrLf & _
               "Click OPEN APPLICATION and log in as an administrator or manager first.", _
               vbInformation, "Dashboard"
    End If
End Function

Public Sub DashRefresh()
    RefreshDashboardSheet
End Sub

Public Sub DashReset()
    If Not CheckUnlocked() Then Exit Sub

    SetProductFilter "", ""
    SetCategoryFilter ""
    SetPeriodFilter "ALL", Date, Date
    RefreshDashboardSheet
End Sub

Public Sub DashPickProduct()
    If Not CheckUnlocked() Then Exit Sub

    On Error GoTo ErrorHandler
    frmPickProduct.Show
    RefreshDashboardSheet
    Exit Sub

ErrorHandler:
    MsgBox "The product list could not be opened." & vbCrLf & Err.description, vbExclamation, "Dashboard"
End Sub

Public Sub DashPickCategory()
    If Not CheckUnlocked() Then Exit Sub

    On Error GoTo ErrorHandler
    frmPickCategory.Show
    RefreshDashboardSheet
    Exit Sub

ErrorHandler:
    MsgBox "The category list could not be opened." & vbCrLf & Err.description, vbExclamation, "Dashboard"
End Sub

Public Sub DashPickPeriod()
    If Not CheckUnlocked() Then Exit Sub

    On Error GoTo ErrorHandler
    frmPickPeriod.Show
    RefreshDashboardSheet
    Exit Sub

ErrorHandler:
    MsgBox "The period list could not be opened." & vbCrLf & Err.description, vbExclamation, "Dashboard"
End Sub

' ============================================================
' Protected sheets
' ============================================================

' If a sheet is protected (for example in the read-only view mode), makes it writable.
' Returns True if it WAS protected, so it can be protected again afterwards.
Private Function MakeWritable(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    If ws.ProtectContents Or ws.ProtectDrawingObjects Then
        MakeWritable = True
        ws.Unprotect ""
    End If
    On Error GoTo 0
End Function

' Protects a sheet again (no password), the way it was before.
Private Sub ProtectAgain(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True
    On Error GoTo 0
End Sub

' Makes the two dashboard sheets writable (used when the page is rebuilt).
Public Sub UnlockDashboardSheets()
    Dim ignore As Boolean

    ignore = MakeWritable(ThisWorkbook.Worksheets("Dashboard"))
    ignore = MakeWritable(modDashBuild.GetDataSheet())
End Sub
