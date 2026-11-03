Attribute VB_Name = "modCharts"
Option Explicit

' ============================================================
' modCharts - the numbers behind the dashboard charts
' Uses Scripting.Dictionary (Windows Excel).
' ============================================================

' Revenue and net profit for each of the last monthCount months (oldest first).
Public Sub GetMonthlySeries(ByVal monthCount As Long, ByRef labels() As String, _
                            ByRef revenue() As Double, ByRef netProfit() As Double)
    Dim i As Long
    Dim firstDay As Date, lastDay As Date
    Dim r As modProfit.TProfit

    ReDim labels(0 To monthCount - 1)
    ReDim revenue(0 To monthCount - 1)
    ReDim netProfit(0 To monthCount - 1)

    For i = 0 To monthCount - 1
        ' i = 0 is the oldest month, the last one is the current month
        firstDay = DateSerial(Year(Date), Month(Date) - (monthCount - 1 - i), 1)
        lastDay = DateSerial(Year(firstDay), Month(firstDay) + 1, 0)       ' day 0 = last day of the month
        If lastDay > Date Then lastDay = Date

        r = modProfit.CalculateProfit(firstDay, lastDay)

        labels(i) = Format$(firstDay, "mmm yy")
        revenue(i) = r.revenue
        netProfit(i) = r.netProfit
    Next i
End Sub

' The best-selling products by revenue over the last "days" days.
' Fills labels and values (biggest first) and returns how many were found (0 = none).
Public Function GetTopProducts(ByVal maxItems As Long, ByVal days As Long, _
                               ByRef labels() As String, ByRef values() As Double) As Long
    Dim tblSales As ListObject, tblLines As ListObject
    Dim salesData As Variant, saleLines As Variant
    Dim saleIDs As Object, revenue As Object, names As Object
    Dim colSaleID As Long, colSaleDate As Long
    Dim colLineSale As Long, colLineProduct As Long, colLineTotal As Long
    Dim rangeStart As Date, rangeEnd As Date
    Dim keys As Variant
    Dim i As Long, n As Long
    Dim dt As Date
    Dim productID As String

    rangeStart = DateValue(Date) - (days - 1)
    rangeEnd = DateValue(Date) + 1

    Set tblSales = modDatabase.GetTable("Sales", "tblSales")
    Set tblLines = modDatabase.GetTable("SaleDetails", "tblSaleDetails")
    If tblSales.DataBodyRange Is Nothing Then Exit Function
    If tblLines.DataBodyRange Is Nothing Then Exit Function

    salesData = tblSales.DataBodyRange.value
    saleLines = tblLines.DataBodyRange.value

    colSaleID = tblSales.ListColumns("SaleID").index
    colSaleDate = tblSales.ListColumns("SaleDate").index
    colLineSale = tblLines.ListColumns("SaleID").index
    colLineProduct = tblLines.ListColumns("ProductID").index
    colLineTotal = tblLines.ListColumns("Total").index

    Set saleIDs = CreateObject("Scripting.Dictionary")
    Set revenue = CreateObject("Scripting.Dictionary")

    ' Which sales fall inside the period?
    For i = 1 To UBound(salesData, 1)
        If Len(CStr(salesData(i, colSaleID))) > 0 And IsDate(salesData(i, colSaleDate)) Then
            dt = CDate(salesData(i, colSaleDate))
            If dt >= rangeStart And dt < rangeEnd Then saleIDs(CStr(salesData(i, colSaleID))) = True
        End If
    Next i

    ' Revenue of each product on those sales
    For i = 1 To UBound(saleLines, 1)
        If saleIDs.Exists(CStr(saleLines(i, colLineSale))) Then
            productID = CStr(saleLines(i, colLineProduct))
            If Len(productID) > 0 Then
                revenue(productID) = revenue(productID) + ToNumber(saleLines(i, colLineTotal))
            End If
        End If
    Next i

    If revenue.count = 0 Then Exit Function

    Set names = NameMapOf("Products", "tblProducts", "ProductID", "ProductName")
    keys = revenue.keys
    SortKeysDesc keys, revenue

    n = revenue.count
    If n > maxItems Then n = maxItems

    ReDim labels(0 To n - 1)
    ReDim values(0 To n - 1)
    For i = 0 To n - 1
        productID = CStr(keys(i))
        If names.Exists(productID) Then
            labels(i) = CStr(names(productID))
        Else
            labels(i) = productID
        End If
        values(i) = Round(revenue(productID), 2)
    Next i

    GetTopProducts = n
End Function

' Units in stock per category (biggest first). Returns how many categories (0 = none).
Public Function GetStockByCategory(ByVal maxItems As Long, ByRef labels() As String, _
                                   ByRef values() As Double) As Long
    Dim tbl As ListObject
    Dim data As Variant
    Dim totals As Object
    Dim colID As Long, colCategory As Long, colStock As Long
    Dim keys As Variant
    Dim i As Long, n As Long
    Dim category As String

    Set tbl = modDatabase.GetTable("Products", "tblProducts")
    If tbl.DataBodyRange Is Nothing Then Exit Function

    data = tbl.DataBodyRange.value
    colID = tbl.ListColumns("ProductID").index
    colCategory = tbl.ListColumns("Category").index
    colStock = tbl.ListColumns("Stock").index

    Set totals = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, colID))) > 0 Then
            category = Trim$(CStr(data(i, colCategory)))
            If Len(category) = 0 Then category = "(none)"
            totals(category) = totals(category) + ToNumber(data(i, colStock))
        End If
    Next i

    If totals.count = 0 Then Exit Function

    keys = totals.keys
    SortKeysDesc keys, totals

    n = totals.count
    If n > maxItems Then n = maxItems

    ReDim labels(0 To n - 1)
    ReDim values(0 To n - 1)
    For i = 0 To n - 1
        labels(i) = CStr(keys(i))
        values(i) = totals(CStr(keys(i)))
    Next i

    GetStockByCategory = n
End Function

' ---------- Small helpers ----------

' Dictionary of ID -> name for any table.
Private Function NameMapOf(ByVal sheetName As String, ByVal tableName As String, _
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

    Set NameMapOf = map
End Function

' Sorts keys so the biggest value (from the dictionary) comes first (in place).
Private Sub SortKeysDesc(ByRef keys As Variant, ByVal values As Object)
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

