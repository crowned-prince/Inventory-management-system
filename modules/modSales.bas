Attribute VB_Name = "modSales"
Option Explicit

' ============================================================
' modSales - tax rules and saving a completed sale
' Uses Scripting.Dictionary (Windows Excel).
' ============================================================

' One-time number formatting for the two sales sheets (run once).
Public Sub PrepareSalesSheets()
    Dim tbl As ListObject

    On Error GoTo ErrorHandler

    Set tbl = modDatabase.GetTable("Sales", "tblSales")
    tbl.ListColumns("SaleDate").DataBodyRange.numberFormat = "dd-mm-yyyy hh:mm"
    tbl.ListColumns("Subtotal").DataBodyRange.numberFormat = "#,##0.00"
    tbl.ListColumns("Discount").DataBodyRange.numberFormat = "#,##0.00"
    tbl.ListColumns("Tax").DataBodyRange.numberFormat = "#,##0.00"
    tbl.ListColumns("GrandTotal").DataBodyRange.numberFormat = "#,##0.00"
    ThisWorkbook.Worksheets("Sales").Columns.AutoFit

    Set tbl = modDatabase.GetTable("SaleDetails", "tblSaleDetails")
    tbl.ListColumns("Price").DataBodyRange.numberFormat = "#,##0.00"
    tbl.ListColumns("CostPrice").DataBodyRange.numberFormat = "#,##0.00"
    tbl.ListColumns("Discount").DataBodyRange.numberFormat = "#,##0.00"
    tbl.ListColumns("Total").DataBodyRange.numberFormat = "#,##0.00"
    ThisWorkbook.Worksheets("SaleDetails").Columns.AutoFit

    MsgBox "Sales sheets formatted.", vbInformation, "Setup"
    Exit Sub

ErrorHandler:
    MsgBox "Could not format the sales sheets." & vbCrLf & Err.description, vbExclamation, "Setup"
End Sub

' ---------- Tax settings (stored on the Settings sheet) ----------
Public Function TaxEnabled() As Boolean
    TaxEnabled = (UCase$(Trim$(GetSetting("TaxEnabled", "No"))) = "YES")
End Function

Public Function TaxRatePercent() As Double
    TaxRatePercent = ToNumber(GetSetting("TaxRate", "0"))
End Function

' Tax on an amount (0 if tax is switched off).
Public Function CalculateTax(ByVal taxableAmount As Double) As Double
    If Not TaxEnabled() Then Exit Function
    If taxableAmount <= 0 Then Exit Function
    CalculateTax = Round(taxableAmount * TaxRatePercent() / 100, 2)
End Function

' ---------- Completing a sale ----------
' The four arrays hold one entry per cart line (same positions go together).
' Returns the new Sale ID, and gives back the invoice number and grand total.
' Raises an error with a readable message if anything is wrong, and
' checks EVERYTHING before writing anything.
Public Function CompleteSale(ByVal customerID As String, _
                             ByRef productIDs() As String, ByRef quantities() As Double, _
                             ByRef unitPrices() As Double, ByRef lineDiscounts() As Double, _
                             ByVal invoiceDiscount As Double, _
                             ByRef invoiceNo As String, ByRef grandTotal As Double) As String
    Dim required As Object
    Dim p As modProducts.TProduct
    Dim key As Variant
    Dim i As Long
    Dim lineTotal As Double, subtotal As Double, tax As Double
    Dim saleID As String
    Dim tblHeader As ListObject, tblLines As ListObject
    Dim newRow As ListRow

    ' ---- 1. Check everything first ----
    If modCustomers.FindCustomerRow(customerID) = 0 Then
        Err.Raise vbObjectError + 530, , "The selected customer no longer exists."
    End If

    Set required = CreateObject("Scripting.Dictionary")

    For i = LBound(productIDs) To UBound(productIDs)
        If Not modProducts.LoadProduct(productIDs(i), p) Then
            Err.Raise vbObjectError + 531, , "Product " & productIDs(i) & " no longer exists."
        End If
        If p.status <> "Active" Then
            Err.Raise vbObjectError + 532, , p.productName & " is not active, so it cannot be sold."
        End If
        If quantities(i) <= 0 Then
            Err.Raise vbObjectError + 533, , "Every quantity must be greater than zero."
        End If
        If unitPrices(i) < 0 Then
            Err.Raise vbObjectError + 534, , "A price cannot be negative."
        End If

        lineTotal = Round(quantities(i) * unitPrices(i), 2)
        If lineDiscounts(i) < 0 Or lineDiscounts(i) > lineTotal Then
            Err.Raise vbObjectError + 535, , "The discount for " & p.productName & " is not valid."
        End If

        subtotal = subtotal + Round(lineTotal - lineDiscounts(i), 2)
        required(productIDs(i)) = required(productIDs(i)) + quantities(i)   ' total needed per product
    Next i

    ' Enough stock for each product (even if it appears on several lines)?
    For Each key In required.keys
        modProducts.LoadProduct CStr(key), p
        If required(key) > p.stock Then
            Err.Raise vbObjectError + 536, , "Insufficient stock for " & p.productName & _
                ". Only " & p.stock & " units are available."
        End If
    Next key

    If invoiceDiscount < 0 Or invoiceDiscount > subtotal Then
        Err.Raise vbObjectError + 537, , "The discount cannot be more than the subtotal."
    End If

    tax = CalculateTax(subtotal - invoiceDiscount)
    grandTotal = Round(subtotal - invoiceDiscount + tax, 2)

    ' ---- 2. Write the sale header ----
    saleID = GenerateID("SAL")
    invoiceNo = GenerateID("INV")

    Set tblHeader = modDatabase.GetTable("Sales", "tblSales")
    Set newRow = NextTableRow(tblHeader)
    ' Order matches tblSales (9 columns)
    newRow.Range.value = Array(saleID, invoiceNo, customerID, Now, subtotal, _
                               invoiceDiscount, tax, grandTotal, modMain.CurrentUserID)

    ' ---- 3. Write each line and reduce the stock ----
    Set tblLines = modDatabase.GetTable("SaleDetails", "tblSaleDetails")
    For i = LBound(productIDs) To UBound(productIDs)
        modProducts.LoadProduct productIDs(i), p            ' read the CURRENT cost price
        lineTotal = Round(quantities(i) * unitPrices(i) - lineDiscounts(i), 2)

        Set newRow = NextTableRow(tblLines)
        ' Order matches tblSaleDetails (8 columns)
        newRow.Range.value = Array(GenerateID("SDT"), saleID, productIDs(i), quantities(i), _
                                   unitPrices(i), p.PurchasePrice, lineDiscounts(i), lineTotal)

        modInventory.DecreaseStock productIDs(i), quantities(i), "SALE", saleID
    Next i

    ' ---- 4. Update the customer's running total ----
    modCustomers.AddToCustomerTotal customerID, grandTotal

    CompleteSale = saleID
End Function

