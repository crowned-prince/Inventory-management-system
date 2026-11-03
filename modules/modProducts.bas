Attribute VB_Name = "modProducts"
Option Explicit

' ============================================================
' modProducts - everything that reads or writes tblProducts.
' Forms never touch the sheet directly; they call these.
' ============================================================

' A "Type" groups related values under one name, like a mini record.
Public Type TProduct
    productID As String
    barcode As String
    productName As String
    category As String
    brand As String
    supplierID As String
    PurchasePrice As Double
    sellingPrice As Double
    stock As Double
    minStock As Double
    MaxStock As Double
    Unit As String
    ExpiryDate As String          ' "" or dd-mm-yyyy
    description As String
    status As String
End Type

Private Function ProductsTable() As ListObject
    Set ProductsTable = ThisWorkbook.Worksheets("Products").ListObjects("tblProducts")
End Function

' ---------- One-time sheet formatting (run once) ----------
Public Sub PrepareProductsSheet()
    Dim tbl As ListObject

    On Error GoTo ErrorHandler
    Set tbl = ProductsTable()

    tbl.ListColumns("Barcode").DataBodyRange.numberFormat = "@"      ' keep leading zeros
    tbl.ListColumns("PurchasePrice").DataBodyRange.numberFormat = "#,##0.00"
    tbl.ListColumns("SellingPrice").DataBodyRange.numberFormat = "#,##0.00"
    tbl.ListColumns("ExpiryDate").DataBodyRange.numberFormat = "dd-mm-yyyy"
    tbl.ListColumns("CreatedDate").DataBodyRange.numberFormat = "dd-mm-yyyy"
    ThisWorkbook.Worksheets("Products").Columns.AutoFit

    MsgBox "Products sheet formatted.", vbInformation, "Setup"
    Exit Sub

ErrorHandler:
    MsgBox "Could not format the Products sheet." & vbCrLf & Err.description, vbExclamation, "Setup"
End Sub

' ---------- Finding records ----------

' Returns the row number inside the table (1 = first data row), or 0 if not found.
Public Function FindProductRow(ByVal productID As String) As Long
    Dim tbl As ListObject
    Dim data As Variant
    Dim idCol As Long, i As Long

    Set tbl = ProductsTable()
    If tbl.DataBodyRange Is Nothing Then Exit Function

    data = tbl.DataBodyRange.value
    idCol = tbl.ListColumns("ProductID").index

    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, idCol)), productID, vbTextCompare) = 0 And Len(productID) > 0 Then
            FindProductRow = i
            Exit Function
        End If
    Next i
End Function

' True if another product already uses this barcode.
Public Function BarcodeExists(ByVal barcode As String, ByVal excludeProductID As String) As Boolean
    Dim tbl As ListObject
    Dim data As Variant
    Dim codeCol As Long, idCol As Long, i As Long

    barcode = Trim$(barcode)
    If Len(barcode) = 0 Then Exit Function

    Set tbl = ProductsTable()
    If tbl.DataBodyRange Is Nothing Then Exit Function

    data = tbl.DataBodyRange.value
    codeCol = tbl.ListColumns("Barcode").index
    idCol = tbl.ListColumns("ProductID").index

    For i = 1 To UBound(data, 1)
        If CStr(data(i, codeCol)) = barcode And CStr(data(i, idCol)) <> excludeProductID Then
            BarcodeExists = True
            Exit Function
        End If
    Next i
End Function

' ---------- Add / Update / Delete ----------

' Saves a new product. Returns the new Product ID.
Public Function AddProduct(ByRef p As TProduct) As String
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim expiry As Variant
    Dim dt As Date

    Set tbl = ProductsTable()

    p.productID = GenerateID("PRD")

    expiry = ""
    If Len(p.ExpiryDate) > 0 Then
        If modValidation.TryParseDate(p.ExpiryDate, dt) Then expiry = dt
    End If

    Set newRow = NextTableRow(tbl)
    ' The order here MUST match the column order of tblProducts (16 columns)
    newRow.Range.value = Array(p.productID, p.barcode, p.productName, p.category, p.brand, _
                               p.supplierID, p.PurchasePrice, p.sellingPrice, p.stock, _
                               p.minStock, p.MaxStock, p.Unit, expiry, p.description, _
                               p.status, Date)

    If p.stock > 0 Then
        modInventory.LogStockMovement p.productID, "OPENING", p.stock, p.stock, p.productID
    End If

    AddProduct = p.productID
End Function

' Updates an existing product. Stock and CreatedDate are NOT changed here.
Public Function UpdateProduct(ByRef p As TProduct) As Boolean
    Dim tbl As ListObject
    Dim rowNo As Long
    Dim expiry As Variant
    Dim dt As Date

    Set tbl = ProductsTable()
    rowNo = FindProductRow(p.productID)
    If rowNo = 0 Then Exit Function

    expiry = ""
    If Len(p.ExpiryDate) > 0 Then
        If modValidation.TryParseDate(p.ExpiryDate, dt) Then expiry = dt
    End If

    SetField tbl, rowNo, "Barcode", p.barcode
    SetField tbl, rowNo, "ProductName", p.productName
    SetField tbl, rowNo, "Category", p.category
    SetField tbl, rowNo, "Brand", p.brand
    SetField tbl, rowNo, "SupplierID", p.supplierID
    SetField tbl, rowNo, "PurchasePrice", p.PurchasePrice
    SetField tbl, rowNo, "SellingPrice", p.sellingPrice
    SetField tbl, rowNo, "MinStock", p.minStock
    SetField tbl, rowNo, "MaxStock", p.MaxStock
    SetField tbl, rowNo, "Unit", p.Unit
    SetField tbl, rowNo, "ExpiryDate", expiry
    SetField tbl, rowNo, "Description", p.description
    SetField tbl, rowNo, "Status", p.status

    UpdateProduct = True
End Function

Private Sub SetField(ByVal tbl As ListObject, ByVal rowNo As Long, _
                     ByVal columnName As String, ByVal newValue As Variant)
    tbl.DataBodyRange.Cells(rowNo, tbl.ListColumns(columnName).index).value = newValue
End Sub

' Deletes a product. Returns "" if deleted, otherwise the reason it could not be.
Public Function DeleteProduct(ByVal productID As String) As String
    Dim tbl As ListObject
    Dim rowNo As Long

    If ProductIsUsed(productID) Then
        DeleteProduct = "This product has sales or purchase history, so it cannot be deleted." & _
                        vbCrLf & "Set its Status to Inactive instead."
        Exit Function
    End If

    Set tbl = ProductsTable()
    rowNo = FindProductRow(productID)
    If rowNo = 0 Then
        DeleteProduct = "That product could not be found."
        Exit Function
    End If

    tbl.ListRows(rowNo).Delete
    DeleteProduct = ""
End Function

' True if the product appears in any sale or purchase line.
Public Function ProductIsUsed(ByVal productID As String) As Boolean
    ProductIsUsed = ColumnContains("SaleDetails", "tblSaleDetails", "ProductID", productID) Or _
                    ColumnContains("PurchaseDetails", "tblPurchaseDetails", "ProductID", productID)
End Function

Private Function ColumnContains(ByVal sheetName As String, ByVal tableName As String, _
                                ByVal columnName As String, ByVal searchValue As String) As Boolean
    Dim tbl As ListObject
    Dim data As Variant
    Dim colNo As Long, i As Long

    Set tbl = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    If tbl.DataBodyRange Is Nothing Then Exit Function

    data = tbl.DataBodyRange.value
    colNo = tbl.ListColumns(columnName).index
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colNo)) = searchValue Then
            ColumnContains = True
            Exit Function
        End If
    Next i
End Function

' ---------- Loading a product into a TProduct ----------
Public Function LoadProduct(ByVal productID As String, ByRef p As TProduct) As Boolean
    Dim tbl As ListObject
    Dim rowNo As Long
    Dim v As Variant

    rowNo = FindProductRow(productID)
    If rowNo = 0 Then Exit Function
    Set tbl = ProductsTable()

    p.productID = CStr(CellValue(tbl, rowNo, "ProductID"))
    p.barcode = CStr(CellValue(tbl, rowNo, "Barcode"))
    p.productName = CStr(CellValue(tbl, rowNo, "ProductName"))
    p.category = CStr(CellValue(tbl, rowNo, "Category"))
    p.brand = CStr(CellValue(tbl, rowNo, "Brand"))
    p.supplierID = CStr(CellValue(tbl, rowNo, "SupplierID"))
    p.PurchasePrice = ToNumber(CellValue(tbl, rowNo, "PurchasePrice"))
    p.sellingPrice = ToNumber(CellValue(tbl, rowNo, "SellingPrice"))
    p.stock = ToNumber(CellValue(tbl, rowNo, "Stock"))
    p.minStock = ToNumber(CellValue(tbl, rowNo, "MinStock"))
    p.MaxStock = ToNumber(CellValue(tbl, rowNo, "MaxStock"))
    p.Unit = CStr(CellValue(tbl, rowNo, "Unit"))
    p.description = CStr(CellValue(tbl, rowNo, "Description"))
    p.status = CStr(CellValue(tbl, rowNo, "Status"))

    v = CellValue(tbl, rowNo, "ExpiryDate")
    If IsDate(v) Then p.ExpiryDate = Format$(CDate(v), "dd-mm-yyyy") Else p.ExpiryDate = ""

    LoadProduct = True
End Function

Private Function CellValue(ByVal tbl As ListObject, ByVal rowNo As Long, _
                           ByVal columnName As String) As Variant
    CellValue = tbl.DataBodyRange.Cells(rowNo, tbl.ListColumns(columnName).index).value
End Function

' ---------- Filling the screen controls ----------

Public Sub FillProductList(ByVal lst As MSForms.ListBox)
    Dim tbl As ListObject
    Dim data As Variant
    Dim arr() As String
    Dim idCol As Long, nameCol As Long, stockCol As Long
    Dim i As Long, k As Long, total As Long

    lst.Clear
    lst.ColumnCount = 3
    lst.ColumnWidths = "68 pt;135 pt;45 pt"

    Set tbl = ProductsTable()
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    data = tbl.DataBodyRange.value
    idCol = tbl.ListColumns("ProductID").index
    nameCol = tbl.ListColumns("ProductName").index
    stockCol = tbl.ListColumns("Stock").index

    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, idCol))) > 0 Then total = total + 1
    Next i
    If total = 0 Then Exit Sub

    ReDim arr(0 To total - 1, 0 To 2)
    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, idCol))) > 0 Then
            arr(k, 0) = CStr(data(i, idCol))
            arr(k, 1) = CStr(data(i, nameCol))
            arr(k, 2) = CStr(data(i, stockCol))
            k = k + 1
        End If
    Next i

    lst.List = arr            ' one assignment is much faster than adding items one by one
End Sub

Public Sub LoadCategoryCombo(ByVal Cmb As MSForms.ComboBox)
    Dim tbl As ListObject
    Dim data As Variant
    Dim i As Long

    Cmb.Clear
    Set tbl = ThisWorkbook.Worksheets("Categories").ListObjects("tblCategories")
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    data = tbl.DataBodyRange.value
    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, 2))) > 0 Then Cmb.AddItem CStr(data(i, 2))     ' column 2 = CategoryName
    Next i
End Sub

' Two columns: hidden SupplierID + visible SupplierName.
Public Sub LoadSupplierCombo(ByVal Cmb As MSForms.ComboBox)
    Dim tbl As ListObject
    Dim data As Variant
    Dim i As Long

    Cmb.Clear
    Cmb.ColumnCount = 2
    Cmb.ColumnWidths = "0 pt;200 pt"
    Cmb.textColumn = 2

    Cmb.AddItem ""                                   ' first entry = no supplier
    Cmb.List(0, 1) = "(None)"

    Set tbl = ThisWorkbook.Worksheets("Suppliers").ListObjects("tblSuppliers")
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    data = tbl.DataBodyRange.value
    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, 1))) > 0 Then
            Cmb.AddItem CStr(data(i, 1))
            Cmb.List(Cmb.ListCount - 1, 1) = CStr(data(i, 2))
        End If
    Next i
End Sub
' ---------- Search ----------

' Builds a dictionary SupplierID -> SupplierName (Windows Excel).
Private Function SupplierNameMap() As Object
    Dim map As Object
    Dim tbl As ListObject
    Dim data As Variant
    Dim i As Long

    Set map = CreateObject("Scripting.Dictionary")
    Set tbl = ThisWorkbook.Worksheets("Suppliers").ListObjects("tblSuppliers")

    If Not tbl.DataBodyRange Is Nothing Then
        data = tbl.DataBodyRange.value
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 Then map(CStr(data(i, 1))) = CStr(data(i, 2))
        Next i
    End If

    Set SupplierNameMap = map
End Function

' Looks up one supplier's name from its ID.
Public Function GetSupplierName(ByVal supplierID As String) As String
    Dim map As Object
    If Len(supplierID) = 0 Then
        GetSupplierName = "(None)"
        Exit Function
    End If
    Set map = SupplierNameMap()
    If map.Exists(supplierID) Then
        GetSupplierName = CStr(map(supplierID))
    Else
        GetSupplierName = supplierID
    End If
End Function

' Fills the list with matching products. Returns how many were found.
' searchBy: "All fields", "Product ID", "Product Name", "Barcode", "Category", "Brand", "Supplier"
Public Function SearchProducts(ByVal lst As MSForms.ListBox, ByVal searchBy As String, _
                               ByVal searchText As String, _
                               Optional ByVal maxResults As Long = 500) As Long
    Dim tbl As ListObject
    Dim data As Variant
    Dim suppliers As Object
    Dim cid As Long, cName As Long, cBarcode As Long, cCategory As Long
    Dim cBrand As Long, cSupplier As Long, cPrice As Long, cStock As Long
    Dim supplierName As String
    Dim i As Long, found As Long

    lst.Clear
    lst.ColumnCount = 6
    lst.ColumnWidths = "62 pt;150 pt;80 pt;70 pt;55 pt;40 pt"

    Set tbl = ProductsTable()
    If tbl.DataBodyRange Is Nothing Then Exit Function

    searchText = Trim$(searchText)
    Set suppliers = SupplierNameMap()
    data = tbl.DataBodyRange.value

    cid = tbl.ListColumns("ProductID").index
    cName = tbl.ListColumns("ProductName").index
    cBarcode = tbl.ListColumns("Barcode").index
    cCategory = tbl.ListColumns("Category").index
    cBrand = tbl.ListColumns("Brand").index
    cSupplier = tbl.ListColumns("SupplierID").index
    cPrice = tbl.ListColumns("SellingPrice").index
    cStock = tbl.ListColumns("Stock").index

    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, cid))) > 0 Then

            supplierName = ""
            If suppliers.Exists(CStr(data(i, cSupplier))) Then
                supplierName = CStr(suppliers(CStr(data(i, cSupplier))))
            End If

            If RowMatches(searchBy, searchText, CStr(data(i, cid)), CStr(data(i, cName)), _
                          CStr(data(i, cBarcode)), CStr(data(i, cCategory)), _
                          CStr(data(i, cBrand)), supplierName) Then

                lst.AddItem CStr(data(i, cid))
                lst.List(lst.ListCount - 1, 1) = CStr(data(i, cName))
                lst.List(lst.ListCount - 1, 2) = CStr(data(i, cCategory))
                lst.List(lst.ListCount - 1, 3) = CStr(data(i, cBrand))
                lst.List(lst.ListCount - 1, 4) = Format$(ToNumber(data(i, cPrice)), "#,##0.00")
                lst.List(lst.ListCount - 1, 5) = CStr(ToNumber(data(i, cStock)))
                found = found + 1
                If found >= maxResults Then Exit For
            End If
        End If
    Next i

    SearchProducts = found
End Function

Private Function RowMatches(ByVal searchBy As String, ByVal searchText As String, _
                            ByVal id As String, ByVal name As String, ByVal barcode As String, _
                            ByVal category As String, ByVal brand As String, _
                            ByVal supplierName As String) As Boolean
    If Len(searchText) = 0 Then
        RowMatches = True                      ' nothing typed = show everything
        Exit Function
    End If

    Select Case searchBy
        Case "Product ID":   RowMatches = Contains(id, searchText)
        Case "Product Name": RowMatches = Contains(name, searchText)
        Case "Barcode":      RowMatches = Contains(barcode, searchText)
        Case "Category":     RowMatches = Contains(category, searchText)
        Case "Brand":        RowMatches = Contains(brand, searchText)
        Case "Supplier":     RowMatches = Contains(supplierName, searchText)
        Case Else            ' "All fields"
            RowMatches = Contains(id, searchText) Or Contains(name, searchText) Or _
                         Contains(barcode, searchText) Or Contains(category, searchText) Or _
                         Contains(brand, searchText) Or Contains(supplierName, searchText)
    End Select
End Function

' Case-insensitive "contains" test.
Private Function Contains(ByVal haystack As String, ByVal needle As String) As Boolean
    Contains = (InStr(1, haystack, needle, vbTextCompare) > 0)
End Function
' Sets a product's purchase (cost) price. Used when a purchase is completed.
Public Sub SetPurchasePrice(ByVal productID As String, ByVal newPrice As Double)
    Dim tbl As ListObject
    Dim rowNo As Long

    rowNo = FindProductRow(productID)
    If rowNo = 0 Then Exit Sub
    Set tbl = ProductsTable()
    SetField tbl, rowNo, "PurchasePrice", newPrice
End Sub

' Returns the ProductID for a barcode, or "" if there is none.
Public Function FindProductByBarcode(ByVal barcode As String) As String
    Dim tbl As ListObject
    Dim data As Variant
    Dim codeCol As Long, idCol As Long, i As Long

    barcode = Trim$(barcode)
    If Len(barcode) = 0 Then Exit Function

    Set tbl = ProductsTable()
    If tbl.DataBodyRange Is Nothing Then Exit Function

    data = tbl.DataBodyRange.value
    codeCol = tbl.ListColumns("Barcode").index
    idCol = tbl.ListColumns("ProductID").index

    For i = 1 To UBound(data, 1)
        If CStr(data(i, codeCol)) = barcode Then
            FindProductByBarcode = CStr(data(i, idCol))
            Exit Function
        End If
    Next i
End Function

' Two columns: hidden ProductID + visible product name. Active products only by default.
Public Sub LoadProductCombo(ByVal Cmb As MSForms.ComboBox, Optional ByVal activeOnly As Boolean = True)
    Dim tbl As ListObject
    Dim data As Variant
    Dim idCol As Long, nameCol As Long, statusCol As Long, i As Long

    Cmb.Clear
    Cmb.ColumnCount = 2
    Cmb.ColumnWidths = "0 pt;310 pt"
    Cmb.textColumn = 2

    Set tbl = ProductsTable()
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    data = tbl.DataBodyRange.value
    idCol = tbl.ListColumns("ProductID").index
    nameCol = tbl.ListColumns("ProductName").index
    statusCol = tbl.ListColumns("Status").index

    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, idCol))) > 0 Then
            If (Not activeOnly) Or CStr(data(i, statusCol)) = "Active" Then
                Cmb.AddItem CStr(data(i, idCol))
                Cmb.List(Cmb.ListCount - 1, 1) = CStr(data(i, nameCol))
            End If
        End If
    Next i
End Sub
