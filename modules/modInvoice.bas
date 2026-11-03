Attribute VB_Name = "modInvoice"
Option Explicit

' ============================================================
' modInvoice - builds the invoice on the "Invoice" sheet,
' then prints it or saves it as a PDF.
' Uses Scripting.Dictionary (Windows Excel).
' ============================================================

' ---------- The Invoice sheet ----------

' Returns the Invoice sheet, creating it if it does not exist yet.
Private Function GetInvoiceSheet() As Worksheet
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "Invoice" Then
            Set GetInvoiceSheet = ws
            Exit Function
        End If
    Next ws

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.name = "Invoice"
    Set GetInvoiceSheet = ws
End Function

' Creates the Invoice sheet (safe to run again).
Public Sub PrepareInvoiceSheet()
    Dim ws As Worksheet

    On Error GoTo ErrorHandler
    Set ws = GetInvoiceSheet()
    MsgBox "The Invoice sheet is ready.", vbInformation, "Setup"
    Exit Sub

ErrorHandler:
    MsgBox "Could not create the Invoice sheet." & vbCrLf & Err.description, vbExclamation, "Setup"
End Sub

' ---------- Small helpers ----------

' Dictionary of ProductID -> ProductName.
Private Function ProductNameMap() As Object
    Dim map As Object
    Dim tbl As ListObject
    Dim data As Variant
    Dim colID As Long, colName As Long, i As Long

    Set map = CreateObject("Scripting.Dictionary")
    Set tbl = modDatabase.GetTable("Products", "tblProducts")

    If Not tbl.DataBodyRange Is Nothing Then
        data = tbl.DataBodyRange.value
        colID = tbl.ListColumns("ProductID").index
        colName = tbl.ListColumns("ProductName").index
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, colID))) > 0 Then map(CStr(data(i, colID))) = CStr(data(i, colName))
        Next i
    End If

    Set ProductNameMap = map
End Function

' Looks up one text value in a table by ID; returns the fallback if not found.
Private Function LookupText(ByVal sheetName As String, ByVal tableName As String, _
                            ByVal idColumn As String, ByVal idValue As String, _
                            ByVal textColumn As String, ByVal fallback As String) As String
    Dim tbl As ListObject
    Dim rowNo As Long

    LookupText = fallback
    Set tbl = modDatabase.GetTable(sheetName, tableName)
    rowNo = modDatabase.FindRowByValue(tbl, idColumn, idValue)
    If rowNo > 0 Then LookupText = CStr(modDatabase.GetCellValue(tbl, rowNo, textColumn))
End Function

' True if an invoice with this number exists.
Public Function InvoiceExists(ByVal invoiceNo As String) As Boolean
    Dim tbl As ListObject
    Set tbl = modDatabase.GetTable("Sales", "tblSales")
    InvoiceExists = (modDatabase.FindRowByValue(tbl, "InvoiceNo", Trim$(invoiceNo)) > 0)
End Function

' The folder where PDFs are saved (created if needed).
Private Function InvoiceFolder() As String
    Dim basePath As String
    Dim folder As String

    basePath = ThisWorkbook.Path

    ' A cloud (OneDrive) path starts with http and cannot be used for folders
    If Len(basePath) = 0 Or LCase$(Left$(basePath, 4)) = "http" Then
        basePath = Environ$("USERPROFILE") & "\Documents"
    End If

    folder = basePath & "\Invoices\"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    InvoiceFolder = folder
End Function

' Opens a file (for example a PDF) in its default program.
Public Sub OpenFile(ByVal filePath As String)
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:=filePath
    On Error GoTo 0
End Sub

' ---------- Building the invoice ----------

' Fills the Invoice sheet for one invoice number and returns the sheet.
' Raises an error with a readable message if the invoice does not exist.
Public Function BuildInvoice(ByVal invoiceNo As String) As Worksheet
    Dim ws As Worksheet
    Dim tblSales As ListObject, tblLines As ListObject
    Dim saleRow As Long
    Dim saleID As String, customerID As String, userID As String
    Dim customerName As String, cashier As String, phone As String
    Dim saleDate As Variant
    Dim subtotal As Double, discount As Double, tax As Double, grandTotal As Double
    Dim data As Variant
    Dim products As Object
    Dim colLineSale As Long, colLineProduct As Long, colLineQty As Long
    Dim colLinePrice As Long, colLineDisc As Long, colLineTotal As Long
    Dim i As Long, r As Long
    Dim lineCount As Long, lastLine As Long, totalsRow As Long, footerRow As Long
    Dim cur As String
    Dim productID As String, productName As String

    invoiceNo = Trim$(invoiceNo)

    ' ---- 1. Find the sale and read its figures ----
    Set tblSales = modDatabase.GetTable("Sales", "tblSales")
    saleRow = modDatabase.FindRowByValue(tblSales, "InvoiceNo", invoiceNo)
    If saleRow = 0 Then
        Err.Raise vbObjectError + 560, , "Invoice " & invoiceNo & " was not found."
    End If

    saleID = CStr(modDatabase.GetCellValue(tblSales, saleRow, "SaleID"))
    customerID = CStr(modDatabase.GetCellValue(tblSales, saleRow, "CustomerID"))
    userID = CStr(modDatabase.GetCellValue(tblSales, saleRow, "UserID"))
    saleDate = modDatabase.GetCellValue(tblSales, saleRow, "SaleDate")
    subtotal = ToNumber(modDatabase.GetCellValue(tblSales, saleRow, "Subtotal"))
    discount = ToNumber(modDatabase.GetCellValue(tblSales, saleRow, "Discount"))
    tax = ToNumber(modDatabase.GetCellValue(tblSales, saleRow, "Tax"))
    grandTotal = ToNumber(modDatabase.GetCellValue(tblSales, saleRow, "GrandTotal"))

    customerName = LookupText("Customers", "tblCustomers", "CustomerID", customerID, "CustomerName", customerID)
    cashier = LookupText("Users", "tblUsers", "UserID", userID, "FullName", userID)

    cur = Trim$(GetSetting("CurrencySymbol", "Rs."))
    phone = Trim$(GetSetting("ShopPhone", ""))

    ' ---- 2. Prepare the sheet ----
    Set ws = GetInvoiceSheet()
    ws.Cells.Clear
    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B").ColumnWidth = 8
    ws.Columns("C").ColumnWidth = 13
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 15
    ws.Cells.Font.name = "Calibri"
    ws.Cells.Font.Size = 10

    ' ---- 3. Shop header ----
    ws.Range("A1").value = GetSetting("ShopName", "My Retail Shop")
    ws.Range("A2").value = GetSetting("ShopAddress", "")
    If Len(phone) > 0 Then ws.Range("A3").value = "Phone: " & phone
    ws.Range("A1:E3").HorizontalAlignment = xlCenterAcrossSelection

    With ws.Range("A1")
        .Font.Size = 18
        .Font.Bold = True
    End With
    ws.Rows(1).RowHeight = 28
    ws.Range("A4:E4").Borders(xlEdgeBottom).LineStyle = xlContinuous

    With ws.Range("A5")
        .value = "INVOICE"
        .Font.Size = 14
        .Font.Bold = True
    End With
    ws.Range("A5:E5").HorizontalAlignment = xlCenterAcrossSelection
    ws.Rows(5).RowHeight = 24

    ' ---- 4. Invoice information ----
    ws.Range("A6").value = "Invoice No:  " & invoiceNo
    ws.Range("A6").Font.Bold = True
    If IsDate(saleDate) Then
        ws.Range("A7").value = "Date:  " & Format$(CDate(saleDate), "dd-mm-yyyy  hh:nn")
    Else
        ws.Range("A7").value = "Date:"
    End If
    ws.Range("A8").value = "Customer:  " & customerName
    ws.Range("A9").value = "Served by:  " & cashier

    With ws.Range("A10")
        .value = "Amounts are in " & cur
        .Font.Italic = True
        .Font.Size = 8
    End With

    ' ---- 5. Table header ----
    ws.Range("A11").value = "Product"
    ws.Range("B11").value = "Qty"
    ws.Range("C11").value = "Price"
    ws.Range("D11").value = "Discount"
    ws.Range("E11").value = "Total"
    With ws.Range("A11:E11")
        .Font.Bold = True
        .Interior.Color = RGB(230, 230, 230)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    ws.Range("B11:E11").HorizontalAlignment = xlRight

    ' ---- 6. The product lines ----
    Set tblLines = modDatabase.GetTable("SaleDetails", "tblSaleDetails")
    r = 12
    If Not tblLines.DataBodyRange Is Nothing Then
        Set products = ProductNameMap()
        data = tblLines.DataBodyRange.value
        colLineSale = tblLines.ListColumns("SaleID").index
        colLineProduct = tblLines.ListColumns("ProductID").index
        colLineQty = tblLines.ListColumns("Quantity").index
        colLinePrice = tblLines.ListColumns("Price").index
        colLineDisc = tblLines.ListColumns("Discount").index
        colLineTotal = tblLines.ListColumns("Total").index

        For i = 1 To UBound(data, 1)
            If CStr(data(i, colLineSale)) = saleID Then
                productID = CStr(data(i, colLineProduct))
                If products.Exists(productID) Then
                    productName = CStr(products(productID))
                Else
                    productName = productID
                End If

                ws.Cells(r, 1).numberFormat = "@"                  ' keep the name as plain text
                ws.Cells(r, 1).value = productName
                ws.Cells(r, 2).value = ToNumber(data(i, colLineQty))
                ws.Cells(r, 3).value = ToNumber(data(i, colLinePrice))
                ws.Cells(r, 4).value = ToNumber(data(i, colLineDisc))
                ws.Cells(r, 5).value = ToNumber(data(i, colLineTotal))

                r = r + 1
                lineCount = lineCount + 1
            End If
        Next i
    End If

    lastLine = 11 + lineCount
    If lineCount > 0 Then
        ws.Range(ws.Cells(12, 3), ws.Cells(lastLine, 5)).numberFormat = "#,##0.00"
    End If
    ws.Range(ws.Cells(lastLine, 1), ws.Cells(lastLine, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous

    ' ---- 7. Totals ----
    totalsRow = lastLine + 2
    ws.Cells(totalsRow, 4).value = "Subtotal"
    ws.Cells(totalsRow, 5).value = subtotal
    ws.Cells(totalsRow + 1, 4).value = "Discount"
    ws.Cells(totalsRow + 1, 5).value = discount
    ws.Cells(totalsRow + 2, 4).value = "Tax"
    ws.Cells(totalsRow + 2, 5).value = tax
    ws.Cells(totalsRow + 3, 4).value = "GRAND TOTAL"
    ws.Cells(totalsRow + 3, 5).value = grandTotal

    ws.Range(ws.Cells(totalsRow, 5), ws.Cells(totalsRow + 3, 5)).numberFormat = "#,##0.00"
    ws.Range(ws.Cells(totalsRow, 4), ws.Cells(totalsRow + 3, 4)).HorizontalAlignment = xlRight

    With ws.Range(ws.Cells(totalsRow + 3, 4), ws.Cells(totalsRow + 3, 5))
        .Font.Bold = True
        .Font.Size = 12
        .Borders(xlEdgeTop).LineStyle = xlContinuous
    End With

    ' ---- 8. Thank-you line ----
    footerRow = totalsRow + 6
    ws.Cells(footerRow, 1).value = GetSetting("InvoiceFooter", "Thank you for shopping with us!")
    ws.Range(ws.Cells(footerRow, 1), ws.Cells(footerRow, 5)).HorizontalAlignment = xlCenterAcrossSelection
    ws.Cells(footerRow, 1).Font.Italic = True

    ' ---- 9. Print area (skipped quietly if Excel has no printer to ask) ----
    On Error Resume Next
    ws.PageSetup.PrintArea = "$A$1:$E$" & footerRow
    On Error GoTo 0

    Set BuildInvoice = ws
End Function

' ---------- Printing and PDF ----------

' Prints the invoice on paper. Raises a readable error if printing fails.
Public Sub PrintInvoice(ByVal invoiceNo As String)
    Dim ws As Worksheet
    Dim previousState As Long
    Dim errText As String

    Set ws = BuildInvoice(invoiceNo)

    previousState = ws.Visible
    ws.Visible = xlSheetVisible

    On Error GoTo PrintError
    ws.PrintOut Copies:=1
    On Error GoTo 0

    ws.Visible = previousState
    Exit Sub

PrintError:
    errText = Err.description
    On Error Resume Next
    ws.Visible = previousState
    On Error GoTo 0
    Err.Raise vbObjectError + 562, , "The invoice could not be printed. " & errText & _
              " (Check that a printer is installed and set as the default.)"
End Sub

' Saves the invoice as a PDF and returns the file path.
Public Function SaveInvoicePdf(ByVal invoiceNo As String) As String
    Dim ws As Worksheet
    Dim filePath As String
    Dim previousState As Long
    Dim errText As String

    Set ws = BuildInvoice(invoiceNo)
    filePath = InvoiceFolder() & Trim$(invoiceNo) & ".pdf"

    previousState = ws.Visible
    ws.Visible = xlSheetVisible

    On Error GoTo ExportError
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=filePath, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    On Error GoTo 0

    ws.Visible = previousState
    SaveInvoicePdf = filePath
    Exit Function

ExportError:
    errText = Err.description
    On Error Resume Next
    ws.Visible = previousState
    On Error GoTo 0
    Err.Raise vbObjectError + 561, , "The PDF could not be created. " & errText
End Function

' ---------- Quick layout check ----------

' Builds the newest invoice and shows the Invoice sheet.
' Run it from the VBA editor (click inside it, press F5).
Public Sub PreviewLastInvoice()
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim invoiceNo As String
    Dim ws As Worksheet

    On Error GoTo ErrorHandler

    Set tbl = modDatabase.GetTable("Sales", "tblSales")
    lastRow = tbl.ListRows.count
    If lastRow > 0 Then invoiceNo = CStr(modDatabase.GetCellValue(tbl, lastRow, "InvoiceNo"))

    If Len(invoiceNo) = 0 Then
        MsgBox "There are no sales yet. Complete a sale first.", vbInformation, "Preview invoice"
        Exit Sub
    End If

    Set ws = BuildInvoice(invoiceNo)
    ws.Visible = xlSheetVisible
    ws.Activate
    Exit Sub

ErrorHandler:
    MsgBox "Could not build the invoice." & vbCrLf & Err.description, vbExclamation, "Preview invoice"
End Sub

