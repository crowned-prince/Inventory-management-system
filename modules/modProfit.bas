Attribute VB_Name = "modProfit"
Option Explicit

' ============================================================
' modProfit - the ONE place where profit is calculated
' Uses Scripting.Dictionary (Windows Excel).
' ============================================================

Public Type TProfit
    revenue As Double          ' sales after all discounts, excluding tax
    CostOfGoods As Double      ' cost price x quantity of what was sold
    GrossProfit As Double
    Expenses As Double
    netProfit As Double
    TaxCollected As Double     ' shown for information only
    SalesCount As Long
End Type

Public Function MonthStart() As Date
    MonthStart = DateSerial(Year(Date), Month(Date), 1)
End Function

' Calculates everything for the dates fromDate to toDate (both days included).
Public Function CalculateProfit(ByVal fromDate As Date, ByVal toDate As Date) As TProfit
    Dim result As TProfit
    Dim rangeStart As Date, rangeEnd As Date
    Dim tbl As ListObject
    Dim data As Variant
    Dim saleIDs As Object
    Dim colSaleID As Long, colDate As Long, colSub As Long, colDisc As Long, colTax As Long
    Dim colDetSale As Long, colDetQty As Long, colDetCost As Long
    Dim colExpDate As Long, colExpAmount As Long
    Dim i As Long
    Dim dt As Date

    rangeStart = DateValue(fromDate)
    rangeEnd = DateValue(toDate) + 1          ' the end is exclusive, so the whole "to" day counts

    Set saleIDs = CreateObject("Scripting.Dictionary")

    ' ---- 1. Sales in the period: revenue and tax ----
    Set tbl = modDatabase.GetTable("Sales", "tblSales")
    If Not tbl.DataBodyRange Is Nothing Then
        data = tbl.DataBodyRange.value
        colSaleID = tbl.ListColumns("SaleID").index
        colDate = tbl.ListColumns("SaleDate").index
        colSub = tbl.ListColumns("Subtotal").index
        colDisc = tbl.ListColumns("Discount").index
        colTax = tbl.ListColumns("Tax").index

        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, colSaleID))) > 0 And IsDate(data(i, colDate)) Then
                dt = CDate(data(i, colDate))
                If dt >= rangeStart And dt < rangeEnd Then
                    saleIDs(CStr(data(i, colSaleID))) = True
                    result.revenue = result.revenue + ToNumber(data(i, colSub)) - ToNumber(data(i, colDisc))
                    result.TaxCollected = result.TaxCollected + ToNumber(data(i, colTax))
                    result.SalesCount = result.SalesCount + 1
                End If
            End If
        Next i
    End If

    ' ---- 2. Cost of the goods on those sales ----
    Set tbl = modDatabase.GetTable("SaleDetails", "tblSaleDetails")
    If Not tbl.DataBodyRange Is Nothing Then
        data = tbl.DataBodyRange.value
        colDetSale = tbl.ListColumns("SaleID").index
        colDetQty = tbl.ListColumns("Quantity").index
        colDetCost = tbl.ListColumns("CostPrice").index

        For i = 1 To UBound(data, 1)
            If saleIDs.Exists(CStr(data(i, colDetSale))) Then
                result.CostOfGoods = result.CostOfGoods + _
                    ToNumber(data(i, colDetQty)) * ToNumber(data(i, colDetCost))
            End If
        Next i
    End If

    ' ---- 3. Expenses in the period ----
    Set tbl = modDatabase.GetTable("Expenses", "tblExpenses")
    If Not tbl.DataBodyRange Is Nothing Then
        data = tbl.DataBodyRange.value
        colExpDate = tbl.ListColumns("ExpenseDate").index
        colExpAmount = tbl.ListColumns("Amount").index

        For i = 1 To UBound(data, 1)
            If IsDate(data(i, colExpDate)) Then
                dt = CDate(data(i, colExpDate))
                If dt >= rangeStart And dt < rangeEnd Then
                    result.Expenses = result.Expenses + ToNumber(data(i, colExpAmount))
                End If
            End If
        Next i
    End If

    ' ---- 4. Totals ----
    result.revenue = Round(result.revenue, 2)
    result.CostOfGoods = Round(result.CostOfGoods, 2)
    result.Expenses = Round(result.Expenses, 2)
    result.TaxCollected = Round(result.TaxCollected, 2)
    result.GrossProfit = Round(result.revenue - result.CostOfGoods, 2)
    result.netProfit = Round(result.GrossProfit - result.Expenses, 2)

    CalculateProfit = result
End Function

' Net profit from the 1st of this month until today (used on the dashboard).
Public Function MonthNetProfit() As Double
    Dim r As TProfit
    r = CalculateProfit(MonthStart(), Date)
    MonthNetProfit = r.netProfit
End Function

' Quick check you can run from the VBA editor (F5): shows today and this month.
Public Sub ShowProfitSummary()
    Dim t As TProfit, m As TProfit
    Dim cur As String

    cur = GetSetting("CurrencySymbol", "Rs.") & " "
    t = CalculateProfit(Date, Date)
    m = CalculateProfit(MonthStart(), Date)

    MsgBox "TODAY  (" & t.SalesCount & " invoice(s))" & vbCrLf & _
           "Revenue:  " & cur & Format$(t.revenue, "#,##0.00") & vbCrLf & _
           "Cost of goods:  " & cur & Format$(t.CostOfGoods, "#,##0.00") & vbCrLf & _
           "Gross profit:  " & cur & Format$(t.GrossProfit, "#,##0.00") & vbCrLf & _
           "Expenses:  " & cur & Format$(t.Expenses, "#,##0.00") & vbCrLf & _
           "Net profit:  " & cur & Format$(t.netProfit, "#,##0.00") & vbCrLf & vbCrLf & _
           "THIS MONTH  (" & m.SalesCount & " invoice(s))" & vbCrLf & _
           "Revenue:  " & cur & Format$(m.revenue, "#,##0.00") & vbCrLf & _
           "Cost of goods:  " & cur & Format$(m.CostOfGoods, "#,##0.00") & vbCrLf & _
           "Gross profit:  " & cur & Format$(m.GrossProfit, "#,##0.00") & vbCrLf & _
           "Expenses:  " & cur & Format$(m.Expenses, "#,##0.00") & vbCrLf & _
           "Net profit:  " & cur & Format$(m.netProfit, "#,##0.00"), _
           vbInformation, "Profit summary"
End Sub

