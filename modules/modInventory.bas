Attribute VB_Name = "modInventory"
Option Explicit

' ============================================================
' modInventory - every stock change, count and check
'   Phase 7 : LogStockMovement
'   Phase 8 : StockStatus
'   Phase 11: IncreaseStock
'   Phase 12: DecreaseStock
'   Phase 13: overview, history, adjustments and stock-log check
' Uses Scripting.Dictionary (Windows Excel).
' ============================================================

' ------------------------------------------------------------
' Stock movement log
' ------------------------------------------------------------

' Records one stock change: OPENING, PURCHASE, SALE, ADJUSTMENT, CORRECTION ...
Public Sub LogStockMovement(ByVal productID As String, ByVal movementType As String, _
                            ByVal quantityChange As Double, ByVal stockAfter As Double, _
                            ByVal referenceID As String)
    Dim tbl As ListObject
    Dim newRow As ListRow

    Set tbl = ThisWorkbook.Worksheets("Inventory").ListObjects("tblStockMovements")
    Set newRow = NextTableRow(tbl)

    ' Order matches tblStockMovements (8 columns)
    newRow.Range.value = Array(GenerateID("MOV"), Now, productID, movementType, _
                               quantityChange, stockAfter, referenceID, modMain.CurrentUserID)
End Sub

' ------------------------------------------------------------
' Stock status
' ------------------------------------------------------------

' Returns the stock status text for a product.
Public Function StockStatus(ByVal stock As Double, ByVal minStock As Double) As String
    If stock <= 0 Then
        StockStatus = "OUT OF STOCK"
    ElseIf stock <= minStock Then
        StockStatus = "LOW STOCK"
    Else
        StockStatus = "IN STOCK"
    End If
End Function

' Stock state shown on screens. Inactive products are never flagged.
Public Function ProductStockState(ByVal stock As Double, ByVal minStock As Double, _
                                  ByVal status As String) As String
    If status <> "Active" Then
        ProductStockState = "INACTIVE"
    Else
        ProductStockState = StockStatus(stock, minStock)
    End If
End Function

' Current stock of one product (0 if the product does not exist).
Public Function CurrentStock(ByVal productID As String) As Double
    Dim tbl As ListObject
    Dim rowNo As Long

    rowNo = modProducts.FindProductRow(productID)
    If rowNo = 0 Then Exit Function

    Set tbl = modDatabase.GetTable("Products", "tblProducts")
    CurrentStock = ToNumber(modDatabase.GetCellValue(tbl, rowNo, "Stock"))
End Function

' ------------------------------------------------------------
' Changing stock (the only places where stock goes up or down)
' ------------------------------------------------------------

' Adds stock to a product and records the movement.
' movementType examples: "PURCHASE", "ADJUSTMENT". Raises an error if the product is missing.
Public Sub IncreaseStock(ByVal productID As String, ByVal quantity As Double, _
                         ByVal movementType As String, ByVal referenceID As String)
    Dim tbl As ListObject
    Dim rowNo As Long
    Dim newStock As Double

    Set tbl = modDatabase.GetTable("Products", "tblProducts")
    rowNo = modProducts.FindProductRow(productID)
    If rowNo = 0 Then Err.Raise vbObjectError + 520, , "Product " & productID & " was not found."

    newStock = ToNumber(modDatabase.GetCellValue(tbl, rowNo, "Stock")) + quantity
    modDatabase.SetCellValue tbl, rowNo, "Stock", newStock

    LogStockMovement productID, movementType, quantity, newStock, referenceID
End Sub

' Removes stock from a product and records the movement.
' Raises an error if there is not enough stock.
Public Sub DecreaseStock(ByVal productID As String, ByVal quantity As Double, _
                         ByVal movementType As String, ByVal referenceID As String)
    Dim tbl As ListObject
    Dim rowNo As Long
    Dim currentQty As Double, newStock As Double

    Set tbl = modDatabase.GetTable("Products", "tblProducts")
    rowNo = modProducts.FindProductRow(productID)
    If rowNo = 0 Then Err.Raise vbObjectError + 520, , "Product " & productID & " was not found."

    currentQty = ToNumber(modDatabase.GetCellValue(tbl, rowNo, "Stock"))
    If quantity > currentQty Then
        Err.Raise vbObjectError + 521, , "Insufficient stock for " & _
            CStr(modDatabase.GetCellValue(tbl, rowNo, "ProductName")) & _
            ". Only " & currentQty & " units are available."
    End If

    newStock = currentQty - quantity
    modDatabase.SetCellValue tbl, rowNo, "Stock", newStock

    LogStockMovement productID, movementType, -quantity, newStock, referenceID   ' negative = stock out
End Sub

' Sets a product's stock to the physically counted quantity.
' Logs the difference as an ADJUSTMENT. Returns the difference (+ or -).
Public Function AdjustStockToCount(ByVal productID As String, ByVal countedQty As Double, _
                                   ByVal reason As String, ByVal note As String) As Double
    Dim currentQty As Double, difference As Double
    Dim reference As String

    If modProducts.FindProductRow(productID) = 0 Then
        Err.Raise vbObjectError + 522, , "Product " & productID & " was not found."
    End If
    If countedQty < 0 Then
        Err.Raise vbObjectError + 523, , "The counted quantity cannot be negative."
    End If

    currentQty = CurrentStock(productID)
    difference = Round(countedQty - currentQty, 4)

    reference = reason
    If Len(Trim$(note)) > 0 Then reference = reference & ": " & Trim$(note)

    If difference > 0 Then
        IncreaseStock productID, difference, "ADJUSTMENT", reference
    ElseIf difference < 0 Then
        DecreaseStock productID, -difference, "ADJUSTMENT", reference
    End If

    AdjustStockToCount = difference
End Function

' ------------------------------------------------------------
' One-time sheet formatting (run once)
' ------------------------------------------------------------
Public Sub PrepareInventorySheet()
    Dim tbl As ListObject

    On Error GoTo ErrorHandler

    Set tbl = modDatabase.GetTable("Inventory", "tblStockMovements")
    tbl.ListColumns("MovementDate").DataBodyRange.numberFormat = "dd-mm-yyyy hh:mm"
    ThisWorkbook.Worksheets("Inventory").Columns.AutoFit

    MsgBox "Inventory sheet formatted.", vbInformation, "Setup"
    Exit Sub

ErrorHandler:
    MsgBox "Could not format the Inventory sheet." & vbCrLf & Err.description, vbExclamation, "Setup"
End Sub

' ------------------------------------------------------------
' Stock screen: summary cards, inventory list and movement history
' ------------------------------------------------------------

' Counts for the summary cards. Status counts include Active products only.
Public Sub InventorySummary(ByRef inStock As Long, ByRef lowStock As Long, _
                            ByRef outOfStock As Long, ByRef totalUnits As Double, _
                            ByRef costValue As Double)
    Dim tbl As ListObject
    Dim data As Variant
    Dim colID As Long, colCost As Long, colStock As Long, colMin As Long, colStatus As Long
    Dim i As Long
    Dim qty As Double

    inStock = 0
    lowStock = 0
    outOfStock = 0
    totalUnits = 0
    costValue = 0

    Set tbl = modDatabase.GetTable("Products", "tblProducts")
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    data = tbl.DataBodyRange.value
    colID = tbl.ListColumns("ProductID").index
    colCost = tbl.ListColumns("PurchasePrice").index
    colStock = tbl.ListColumns("Stock").index
    colMin = tbl.ListColumns("MinStock").index
    colStatus = tbl.ListColumns("Status").index

    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, colID))) > 0 Then
            qty = ToNumber(data(i, colStock))
            totalUnits = totalUnits + qty
            costValue = costValue + qty * ToNumber(data(i, colCost))

            If CStr(data(i, colStatus)) = "Active" Then
                Select Case StockStatus(qty, ToNumber(data(i, colMin)))
                    Case "IN STOCK"
                        inStock = inStock + 1
                    Case "LOW STOCK"
                        lowStock = lowStock + 1
                    Case Else
                        outOfStock = outOfStock + 1
                End Select
            End If
        End If
    Next i
End Sub

' Fills the inventory list. stateFilter: "" (all), "IN STOCK", "LOW STOCK",
' "OUT OF STOCK" or "INACTIVE". Returns how many rows were listed.
Public Function FillInventoryList(ByVal lst As MSForms.ListBox, ByVal stateFilter As String, _
                                  ByVal searchText As String, _
                                  Optional ByVal maxRows As Long = 1000) As Long
    Dim tbl As ListObject
    Dim data As Variant
    Dim colID As Long, colName As Long, colBarcode As Long, colCategory As Long
    Dim colStock As Long, colMin As Long, colStatus As Long
    Dim i As Long, found As Long
    Dim stockQty As Double, minQty As Double
    Dim state As String

    lst.Clear
    lst.ColumnCount = 6
    lst.ColumnWidths = "62 pt;150 pt;80 pt;45 pt;40 pt;85 pt"

    Set tbl = modDatabase.GetTable("Products", "tblProducts")
    If tbl.DataBodyRange Is Nothing Then Exit Function

    searchText = Trim$(searchText)
    data = tbl.DataBodyRange.value
    colID = tbl.ListColumns("ProductID").index
    colName = tbl.ListColumns("ProductName").index
    colBarcode = tbl.ListColumns("Barcode").index
    colCategory = tbl.ListColumns("Category").index
    colStock = tbl.ListColumns("Stock").index
    colMin = tbl.ListColumns("MinStock").index
    colStatus = tbl.ListColumns("Status").index

    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, colID))) > 0 Then
            stockQty = ToNumber(data(i, colStock))
            minQty = ToNumber(data(i, colMin))
            state = ProductStockState(stockQty, minQty, CStr(data(i, colStatus)))

            If Len(stateFilter) = 0 Or state = stateFilter Then
                If Len(searchText) = 0 Or _
                   InStr(1, CStr(data(i, colID)), searchText, vbTextCompare) > 0 Or _
                   InStr(1, CStr(data(i, colName)), searchText, vbTextCompare) > 0 Or _
                   InStr(1, CStr(data(i, colBarcode)), searchText, vbTextCompare) > 0 Or _
                   InStr(1, CStr(data(i, colCategory)), searchText, vbTextCompare) > 0 Then

                    lst.AddItem CStr(data(i, colID))
                    lst.List(lst.ListCount - 1, 1) = CStr(data(i, colName))
                    lst.List(lst.ListCount - 1, 2) = CStr(data(i, colCategory))
                    lst.List(lst.ListCount - 1, 3) = CStr(stockQty)
                    lst.List(lst.ListCount - 1, 4) = CStr(minQty)
                    lst.List(lst.ListCount - 1, 5) = state
                    found = found + 1
                    If found >= maxRows Then Exit For
                End If
            End If
        End If
    Next i

    FillInventoryList = found
End Function

' Fills the history list for one product, newest first.
Public Sub FillMovementHistory(ByVal lst As MSForms.ListBox, ByVal productID As String, _
                               Optional ByVal maxRows As Long = 40)
    Dim tbl As ListObject
    Dim data As Variant
    Dim colDate As Long, colProduct As Long, colType As Long, colQty As Long
    Dim colAfter As Long, colRef As Long
    Dim i As Long, shown As Long, n As Long
    Dim qty As Double

    lst.Clear
    lst.ColumnCount = 5
    lst.ColumnWidths = "78 pt;62 pt;38 pt;38 pt;50 pt"

    If Len(productID) = 0 Then Exit Sub

    Set tbl = modDatabase.GetTable("Inventory", "tblStockMovements")
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    data = tbl.DataBodyRange.value
    colDate = tbl.ListColumns("MovementDate").index
    colProduct = tbl.ListColumns("ProductID").index
    colType = tbl.ListColumns("MovementType").index
    colQty = tbl.ListColumns("Quantity").index
    colAfter = tbl.ListColumns("StockAfter").index
    colRef = tbl.ListColumns("ReferenceID").index

    For i = UBound(data, 1) To 1 Step -1
        If StrComp(CStr(data(i, colProduct)), productID, vbTextCompare) = 0 Then
            lst.AddItem ""
            n = lst.ListCount - 1

            If IsDate(data(i, colDate)) Then
                lst.List(n, 0) = Format$(CDate(data(i, colDate)), "dd-mm-yy hh:nn")
            End If
            lst.List(n, 1) = CStr(data(i, colType))

            qty = ToNumber(data(i, colQty))
            If qty > 0 Then
                lst.List(n, 2) = "+" & CStr(qty)
            Else
                lst.List(n, 2) = CStr(qty)
            End If

            lst.List(n, 3) = CStr(ToNumber(data(i, colAfter)))
            lst.List(n, 4) = CStr(data(i, colRef))

            shown = shown + 1
            If shown >= maxRows Then Exit For
        End If
    Next i
End Sub

' ------------------------------------------------------------
' Stock-log check
' ------------------------------------------------------------

' Adds up every product's movements and compares with its stored stock.
' repair = False : only builds a report.
' repair = True  : adds a CORRECTION movement so the log matches the stock
'                  (the Stock column itself is NOT changed).
' Returns the number of products that did not match.
Public Function CheckStockLog(ByVal repair As Boolean, ByRef report As String) As Long
    Dim totals As Object
    Dim tblMoves As ListObject, tblProducts As ListObject
    Dim moves As Variant, prods As Variant
    Dim colMoveProduct As Long, colMoveQty As Long
    Dim colID As Long, colName As Long, colStock As Long
    Dim i As Long, found As Long
    Dim stored As Double, expected As Double
    Dim id As String

    report = ""
    Set totals = CreateObject("Scripting.Dictionary")

    ' ---- 1. Add up the movements of each product ----
    Set tblMoves = modDatabase.GetTable("Inventory", "tblStockMovements")
    If Not tblMoves.DataBodyRange Is Nothing Then
        moves = tblMoves.DataBodyRange.value
        colMoveProduct = tblMoves.ListColumns("ProductID").index
        colMoveQty = tblMoves.ListColumns("Quantity").index

        For i = 1 To UBound(moves, 1)
            id = CStr(moves(i, colMoveProduct))
            If Len(id) > 0 Then totals(id) = totals(id) + ToNumber(moves(i, colMoveQty))
        Next i
    End If

    ' ---- 2. Compare with the stock stored on each product ----
    Set tblProducts = modDatabase.GetTable("Products", "tblProducts")
    If tblProducts.DataBodyRange Is Nothing Then Exit Function

    prods = tblProducts.DataBodyRange.value
    colID = tblProducts.ListColumns("ProductID").index
    colName = tblProducts.ListColumns("ProductName").index
    colStock = tblProducts.ListColumns("Stock").index

    For i = 1 To UBound(prods, 1)
        id = CStr(prods(i, colID))
        If Len(id) > 0 Then
            stored = ToNumber(prods(i, colStock))
            expected = 0
            If totals.Exists(id) Then expected = totals(id)

            If Abs(stored - expected) > 0.0001 Then
                found = found + 1
                If repair Then
                    LogStockMovement id, "CORRECTION", stored - expected, stored, "Log aligned to current stock"
                ElseIf found <= 15 Then
                    report = report & CStr(prods(i, colName)) & " (" & id & "):  stock " & stored & _
                             ",  movement log says " & expected & vbCrLf
                End If
            End If
        End If
    Next i

    If (Not repair) And found > 15 Then
        report = report & "... and " & (found - 15) & " more." & vbCrLf
    End If

    CheckStockLog = found
End Function

