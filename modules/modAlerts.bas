Attribute VB_Name = "modAlerts"
Option Explicit

' ============================================================
' modAlerts - low-stock detection, login alert, report and sale warning
' Inactive products are never flagged. Uses Scripting.Dictionary (Windows).
' ============================================================

' Map of SupplierID -> SupplierName.
Private Function SupplierMap() As Object
    Dim map As Object
    Dim tbl As ListObject
    Dim data As Variant
    Dim i As Long

    Set map = CreateObject("Scripting.Dictionary")
    Set tbl = modDatabase.GetTable("Suppliers", "tblSuppliers")

    If Not tbl.DataBodyRange Is Nothing Then
        data = tbl.DataBodyRange.value
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 Then map(CStr(data(i, 1))) = CStr(data(i, 2))
        Next i
    End If

    Set SupplierMap = map
End Function

' Builds the list of products that need restocking, most urgent first.
' filterMode: "ALL" (low + out), "LOW" (low only), "OUT" (out of stock only)
' supplierID: "" = all suppliers
' Returns a 2D array (1 To n, 1 To 9) or Empty when nothing needs restocking.
' Columns: 1 ID, 2 Name, 3 Category, 4 Supplier, 5 Stock, 6 Min, 7 Max, 8 Suggested order, 9 Status
Public Function BuildLowStockRows(ByVal filterMode As String, ByVal supplierID As String) As Variant
    Dim tbl As ListObject
    Dim data As Variant
    Dim suppliers As Object
    Dim cid As Long, cName As Long, cCategory As Long, cSupplier As Long
    Dim cStock As Long, cMin As Long, cMax As Long, cStatus As Long
    Dim sortKeys() As Double
    Dim rowIndex() As Long
    Dim result() As Variant
    Dim i As Long, j As Long, k As Long, n As Long, r As Long
    Dim stockQty As Double, minQty As Double, maxQty As Double
    Dim state As String, supplierName As String
    Dim tmpKey As Double, tmpIndex As Long

    Set tbl = modDatabase.GetTable("Products", "tblProducts")
    If tbl.DataBodyRange Is Nothing Then Exit Function

    Set suppliers = SupplierMap()
    data = tbl.DataBodyRange.value

    cid = tbl.ListColumns("ProductID").index
    cName = tbl.ListColumns("ProductName").index
    cCategory = tbl.ListColumns("Category").index
    cSupplier = tbl.ListColumns("SupplierID").index
    cStock = tbl.ListColumns("Stock").index
    cMin = tbl.ListColumns("MinStock").index
    cMax = tbl.ListColumns("MaxStock").index
    cStatus = tbl.ListColumns("Status").index

    ReDim sortKeys(1 To UBound(data, 1))
    ReDim rowIndex(1 To UBound(data, 1))

    ' ---- 1. Find the products that need restocking ----
    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, cid))) > 0 Then
            If CStr(data(i, cStatus)) = "Active" Then
                stockQty = ToNumber(data(i, cStock))
                minQty = ToNumber(data(i, cMin))
                state = modInventory.StockStatus(stockQty, minQty)

                If state <> "IN STOCK" Then
                    If filterMode = "ALL" Or (filterMode = "LOW" And state = "LOW STOCK") Or _
                       (filterMode = "OUT" And state = "OUT OF STOCK") Then
                        If Len(supplierID) = 0 Or CStr(data(i, cSupplier)) = supplierID Then
                            n = n + 1
                            rowIndex(n) = i
                            sortKeys(n) = stockQty - minQty          ' more negative = more urgent
                            If state = "OUT OF STOCK" Then sortKeys(n) = sortKeys(n) - 1000000
                        End If
                    End If
                End If
            End If
        End If
    Next i

    If n = 0 Then Exit Function

    ' ---- 2. Sort by urgency (simple insertion sort) ----
    For i = 2 To n
        tmpKey = sortKeys(i)
        tmpIndex = rowIndex(i)
        j = i - 1
        Do While j >= 1
            If sortKeys(j) <= tmpKey Then Exit Do
            sortKeys(j + 1) = sortKeys(j)
            rowIndex(j + 1) = rowIndex(j)
            j = j - 1
        Loop
        sortKeys(j + 1) = tmpKey
        rowIndex(j + 1) = tmpIndex
    Next i

    ' ---- 3. Build the result table ----
    ReDim result(1 To n, 1 To 9)
    For k = 1 To n
        r = rowIndex(k)
        stockQty = ToNumber(data(r, cStock))
        minQty = ToNumber(data(r, cMin))
        maxQty = ToNumber(data(r, cMax))

        supplierName = "(None)"
        If suppliers.Exists(CStr(data(r, cSupplier))) Then
            supplierName = CStr(suppliers(CStr(data(r, cSupplier))))
        End If

        result(k, 1) = CStr(data(r, cid))
        result(k, 2) = CStr(data(r, cName))
        result(k, 3) = CStr(data(r, cCategory))
        result(k, 4) = supplierName
        result(k, 5) = stockQty
        result(k, 6) = minQty
        result(k, 7) = maxQty
        If maxQty > stockQty Then result(k, 8) = maxQty - stockQty Else result(k, 8) = 0
        result(k, 9) = modInventory.StockStatus(stockQty, minQty)
    Next k

    BuildLowStockRows = result
End Function

' How many products are low and how many are out of stock.
Public Sub CountAlerts(ByRef lowCount As Long, ByRef outCount As Long)
    Dim alertRows As Variant
    Dim i As Long

    lowCount = 0
    outCount = 0

    alertRows = BuildLowStockRows("ALL", "")
    If IsEmpty(alertRows) Then Exit Sub

    For i = 1 To UBound(alertRows, 1)
        If alertRows(i, 9) = "OUT OF STOCK" Then
            outCount = outCount + 1
        Else
            lowCount = lowCount + 1
        End If
    Next i
End Sub

' Fills the Low Stock list on the report screen. Returns how many rows.
Public Function FillLowStockList(ByVal lst As MSForms.ListBox, ByVal filterMode As String, _
                                 ByVal supplierID As String) As Long
    Dim alertRows As Variant
    Dim i As Long, c As Long

    lst.Clear
    lst.ColumnCount = 9
    lst.ColumnWidths = "62 pt;190 pt;85 pt;130 pt;45 pt;40 pt;40 pt;60 pt;90 pt"

    alertRows = BuildLowStockRows(filterMode, supplierID)
    If IsEmpty(alertRows) Then Exit Function

    For i = 1 To UBound(alertRows, 1)
        lst.AddItem CStr(alertRows(i, 1))
        For c = 2 To 9
            lst.List(lst.ListCount - 1, c - 1) = CStr(alertRows(i, c))
        Next c
    Next i

    FillLowStockList = UBound(alertRows, 1)
End Function

' Writes the report onto the Reports sheet. Returns how many products were written.
Public Function WriteLowStockReport(ByVal filterMode As String, ByVal supplierID As String) As Long
    Dim ws As Worksheet
    Dim alertRows As Variant
    Dim n As Long, i As Long

    Set ws = ThisWorkbook.Worksheets("Reports")
    ws.Cells.Clear

    With ws.Range("A1")
        .value = "LOW STOCK REPORT"
        .Font.Size = 16
        .Font.Bold = True
    End With
    ws.Range("A2").value = "Generated " & Format$(Now, "dd-mm-yyyy hh:nn") & " by " & modMain.CurrentUsername

    ws.Range("A4:I4").value = Array("Product ID", "Product", "Category", "Supplier", "Stock", _
                                    "Min Stock", "Max Stock", "Suggested Order", "Status")
    With ws.Range("A4:I4")
        .Interior.Color = ThemeNavy()
        .Font.Color = ThemeWhite()
        .Font.Bold = True
    End With

    alertRows = BuildLowStockRows(filterMode, supplierID)
    If IsEmpty(alertRows) Then
        ws.Range("A5").value = "No products need restocking."
          ws.Range("A4:I" & (4 + n)).Columns.AutoFit
        Exit Function
    End If

    n = UBound(alertRows, 1)
    ws.Range("A5").Resize(n, 9).value = alertRows

    For i = 1 To n
        If alertRows(i, 9) = "OUT OF STOCK" Then
            ws.Cells(4 + i, 9).Font.Color = RGB(192, 0, 0)
        Else
            ws.Cells(4 + i, 9).Font.Color = RGB(204, 102, 0)
        End If
        ws.Cells(4 + i, 9).Font.Bold = True
    Next i

    ws.Range("A4").Resize(n + 1, 9).Borders.LineStyle = xlContinuous
       ws.Range("A4:I" & (4 + n)).Columns.AutoFit

    WriteLowStockReport = n
End Function

' Shown once after login (ADMIN and MANAGER only).
Public Sub ShowLoginAlert()
    Dim lowCount As Long, outCount As Long
    Dim msg As String

    If modMain.CurrentRole = "STAFF" Then Exit Sub

    CountAlerts lowCount, outCount
    If lowCount + outCount = 0 Then Exit Sub

    msg = "STOCK ALERT" & vbCrLf & vbCrLf & _
          lowCount & " product(s) are low on stock" & vbCrLf & _
          outCount & " product(s) are out of stock" & vbCrLf & vbCrLf & _
          "Open the Low Stock report now?"

    If MsgBox(msg, vbYesNo + vbExclamation, "Low stock alert") = vbYes Then
        frmLowStock.Show
    End If
End Sub

' After a sale: names of the sold products that are now low or out of stock.
' Returns "" if none.
Public Function GetLowStockWarning(ByRef productIDs() As String) As String
    Dim p As modProducts.TProduct
    Dim seen As Object
    Dim i As Long
    Dim names As String

    Set seen = CreateObject("Scripting.Dictionary")

    For i = LBound(productIDs) To UBound(productIDs)
        If Not seen.Exists(productIDs(i)) Then
            seen(productIDs(i)) = True
            If modProducts.LoadProduct(productIDs(i), p) Then
                If p.status = "Active" Then
                    If modInventory.StockStatus(p.stock, p.minStock) <> "IN STOCK" Then
                        If Len(names) > 0 Then names = names & ",  "
                        names = names & p.productName & " (" & p.stock & " left)"
                    End If
                End If
            End If
        End If
    Next i

    If Len(names) > 0 Then GetLowStockWarning = "LOW STOCK: " & names
End Function

