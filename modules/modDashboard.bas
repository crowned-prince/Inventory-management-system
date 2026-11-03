Attribute VB_Name = "modDashboard"
Option Explicit

' ============================================================
' modDashboard - calculates dashboard numbers and fills lists.
' Note: uses Scripting.Dictionary, which works on Windows Excel.
' ============================================================

' Returns the table's data as an array, or Empty if the table has no rows.
Private Function TableData(ByVal sheetName As String, ByVal tableName As String) As Variant
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    If tbl.DataBodyRange Is Nothing Then
        TableData = Empty
    Else
        TableData = tbl.DataBodyRange.value
    End If
End Function

' Safely turns a cell value into a number (blank or text becomes 0).
Private Function ToNum(ByVal v As Variant) As Double
    If IsNumeric(v) Then ToNum = CDbl(v) Else ToNum = 0
End Function

' True if the date value falls on today's date.
Private Function IsToday(ByVal v As Variant) As Boolean
    Dim dt As Date
    If IsDate(v) Then
        dt = CDate(v)
        IsToday = (dt >= Date And dt < Date + 1)
    End If
End Function

' ---------- Simple counts and totals ----------

Public Function CountRecords(ByVal sheetName As String, ByVal tableName As String) As Long
    Dim data As Variant
    Dim i As Long
    data = TableData(sheetName, tableName)
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, 1))) > 0 Then CountRecords = CountRecords + 1
    Next i
End Function

Public Function totalStock() As Double
    Dim data As Variant
    Dim i As Long
    data = TableData("Products", "tblProducts")
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, 1))) > 0 Then totalStock = totalStock + ToNum(data(i, 9))
    Next i
End Function

Public Function TodaySales() As Double
    Dim data As Variant
    Dim i As Long
    data = TableData("Sales", "tblSales")
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If IsToday(data(i, 4)) Then TodaySales = TodaySales + ToNum(data(i, 8))
    Next i
End Function

Public Function TodayPurchases() As Double
    Dim data As Variant
    Dim i As Long
    data = TableData("Purchases", "tblPurchases")
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If IsToday(data(i, 3)) Then TodayPurchases = TodayPurchases + ToNum(data(i, 4))
    Next i
End Function

' Profit of today's sales = line total minus (cost price x quantity).
' (Phase 15 will refine this for invoice-level discounts and expenses.)
' Gross profit of today's sales, from the shared profit calculator.
Public Function TodayProfit() As Double
    Dim r As modProfit.TProfit
    r = modProfit.CalculateProfit(Date, Date)
    TodayProfit = r.GrossProfit
End Function

Public Function LowStockCount() As Long
    Dim data As Variant
    Dim i As Long

    data = TableData("Products", "tblProducts")
    If IsEmpty(data) Then Exit Function

    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, 1))) > 0 And CStr(data(i, 15)) = "Active" Then
            If ToNum(data(i, 9)) <= ToNum(data(i, 10)) Then LowStockCount = LowStockCount + 1
        End If
    Next i
End Function

' ---------- Fill the three lists ----------

Public Sub FillRecentSales(ByVal lst As MSForms.ListBox, ByVal maxRows As Long)
    Dim data As Variant
    Dim i As Long, shown As Long

    lst.Clear
    lst.ColumnCount = 3
    lst.ColumnWidths = "95 pt;90 pt;95 pt"

    data = TableData("Sales", "tblSales")
    If Not IsEmpty(data) Then
        For i = UBound(data, 1) To 1 Step -1          ' newest first
            If Len(CStr(data(i, 1))) > 0 And shown < maxRows Then
                lst.AddItem CStr(data(i, 2))                              ' invoice
                If IsDate(data(i, 4)) Then
                    lst.List(lst.ListCount - 1, 1) = Format$(CDate(data(i, 4)), "dd-mm-yyyy")
                End If
                lst.List(lst.ListCount - 1, 2) = Format$(ToNum(data(i, 8)), "#,##0")
                shown = shown + 1
            End If
        Next i
    End If
    If shown = 0 Then lst.AddItem "No sales yet"
End Sub

Public Sub FillLowStock(ByVal lst As MSForms.ListBox, ByVal maxRows As Long)
    Dim data As Variant
    Dim i As Long, shown As Long

    lst.Clear
    lst.ColumnCount = 3
    lst.ColumnWidths = "170 pt;55 pt;55 pt"

    data = TableData("Products", "tblProducts")
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 And shown < maxRows Then
                If CStr(data(i, 15)) = "Active" Then
                    If ToNum(data(i, 9)) <= ToNum(data(i, 10)) Then
                        lst.AddItem CStr(data(i, 3))                           ' name
                        lst.List(lst.ListCount - 1, 1) = "Qty " & CStr(ToNum(data(i, 9)))
                        lst.List(lst.ListCount - 1, 2) = "Min " & CStr(ToNum(data(i, 10)))
                        shown = shown + 1
                    End If
                End If
            End If
        Next i
    End If
    If shown = 0 Then lst.AddItem "All products are well stocked"
End Sub
Public Sub FillTopSelling(ByVal lst As MSForms.ListBox, ByVal maxRows As Long)
    Dim details As Variant, products As Variant
    Dim qtyByProduct As Object
    Dim key As Variant, bestKey As String
    Dim bestQty As Double
    Dim i As Long, rank As Long

    lst.Clear
    lst.ColumnCount = 2
    lst.ColumnWidths = "210 pt;80 pt"

    Set qtyByProduct = CreateObject("Scripting.Dictionary")
    details = TableData("SaleDetails", "tblSaleDetails")

    If Not IsEmpty(details) Then
        For i = 1 To UBound(details, 1)
            If Len(CStr(details(i, 3))) > 0 Then
                qtyByProduct(CStr(details(i, 3))) = qtyByProduct(CStr(details(i, 3))) + ToNum(details(i, 4))
            End If
        Next i
    End If

    products = TableData("Products", "tblProducts")

    For rank = 1 To maxRows
        bestKey = ""
        bestQty = 0
        For Each key In qtyByProduct.keys
            If qtyByProduct(key) > bestQty Then
                bestQty = qtyByProduct(key)
                bestKey = CStr(key)
            End If
        Next key
        If Len(bestKey) = 0 Then Exit For

        lst.AddItem LookupProductName(products, bestKey)
        lst.List(lst.ListCount - 1, 1) = CStr(bestQty) & " sold"
        qtyByProduct(bestKey) = -1            ' exclude it from the next round
    Next rank

    If lst.ListCount = 0 Then lst.AddItem "No sales yet"
End Sub

Private Function LookupProductName(ByVal products As Variant, ByVal productID As String) As String
    Dim i As Long
    LookupProductName = productID               ' fallback if not found
    If IsEmpty(products) Then Exit Function
    For i = 1 To UBound(products, 1)
        If CStr(products(i, 1)) = productID Then
            LookupProductName = CStr(products(i, 3))
            Exit Function
        End If
    Next i
End Function


Public Sub RefreshDashboardCharts()
    Dim wsDash As Worksheet, wsInv As Worksheet
    Dim chartObj As ChartObject
    Dim lastRow As Long
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsInv = ThisWorkbook.Sheets("Inventory")
    
    lastRow = wsInv.Cells(wsInv.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub ' No data to plot
    
    ' Clear existing dashboard charts to prevent stacking
    For Each chartObj In wsDash.ChartObjects
        chartObj.Delete
    Next chartObj
    
    ' --- CHART 1: Stock Quantity Column Chart ---
    Dim stockChartObj As ChartObject
    Set stockChartObj = wsDash.ChartObjects.Add(Left:=30, Top:=120, Width:=480, Height:=280)
    
    With stockChartObj.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=wsInv.Range("B1:B" & lastRow & ",D1:D" & lastRow)
        .HasTitle = True
        .ChartTitle.text = "Current Stock Levels by Product"
        .HasLegend = False
        
        ' Dark Navy Theme Formatting
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(26, 38, 64)
        .ChartTitle.Font.Color = RGB(255, 255, 255)
        .Axes(xlCategory).TickLabels.Font.Color = RGB(200, 200, 200)
        .Axes(xlValue).TickLabels.Font.Color = RGB(200, 200, 200)
    End With
    
    ' --- CHART 2: Stock Value Pie Chart ---
    Dim valChartObj As ChartObject
    Set valChartObj = wsDash.ChartObjects.Add(Left:=530, Top:=120, Width:=450, Height:=280)
    
    With valChartObj.Chart
        .ChartType = xlPie
        .SetSourceData Source:=wsInv.Range("C1:C" & lastRow & ",E1:E" & lastRow)
        .HasTitle = True
        .ChartTitle.text = "Category Price Breakdown"
        
        ' Style formatting
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(26, 38, 64)
        .ChartTitle.Font.Color = RGB(255, 255, 255)
        .Legend.Font.Color = RGB(220, 220, 220)
    End With
    
    MsgBox "Dashboard charts updated!", vbInformation, "Dashboard"
End Sub
