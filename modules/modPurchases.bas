Attribute VB_Name = "modPurchases"
Option Explicit

' ============================================================
' modPurchases - saving a completed purchase
' ============================================================

' One-time number formatting for the two purchase sheets (run once).
Public Sub PreparePurchaseSheets()
    Dim tbl As ListObject

    On Error GoTo ErrorHandler

    Set tbl = modDatabase.GetTable("Purchases", "tblPurchases")
    tbl.ListColumns("PurchaseDate").DataBodyRange.numberFormat = "dd-mm-yyyy"
    tbl.ListColumns("TotalCost").DataBodyRange.numberFormat = "#,##0.00"
    ThisWorkbook.Worksheets("Purchases").Columns.AutoFit

    Set tbl = modDatabase.GetTable("PurchaseDetails", "tblPurchaseDetails")
    tbl.ListColumns("PurchasePrice").DataBodyRange.numberFormat = "#,##0.00"
    tbl.ListColumns("Total").DataBodyRange.numberFormat = "#,##0.00"
    ThisWorkbook.Worksheets("PurchaseDetails").Columns.AutoFit

    MsgBox "Purchase sheets formatted.", vbInformation, "Setup"
    Exit Sub

ErrorHandler:
    MsgBox "Could not format the purchase sheets." & vbCrLf & Err.description, vbExclamation, "Setup"
End Sub

' Saves the purchase and updates stock, supplier totals and logs.
' The three arrays hold one entry per product line (same positions go together).
' Returns the new Purchase ID. Raises an error (with a readable message) if
' anything is wrong - and it checks EVERYTHING before writing anything.
Public Function CompletePurchase(ByVal supplierID As String, ByVal purchaseDate As Date, _
                                 ByRef productIDs() As String, ByRef quantities() As Double, _
                                 ByRef unitCosts() As Double) As String
    Dim tblHeader As ListObject, tblLines As ListObject
    Dim newRow As ListRow
    Dim purchaseID As String
    Dim i As Long
    Dim lineTotal As Double, grandTotal As Double

    ' ---- 1. Check everything first ----
    If modSuppliers.FindSupplierRow(supplierID) = 0 Then
        Err.Raise vbObjectError + 510, , "The selected supplier no longer exists."
    End If

    For i = LBound(productIDs) To UBound(productIDs)
        If modProducts.FindProductRow(productIDs(i)) = 0 Then
            Err.Raise vbObjectError + 511, , "Product " & productIDs(i) & " no longer exists."
        End If
        If quantities(i) <= 0 Then
            Err.Raise vbObjectError + 512, , "Every quantity must be greater than zero."
        End If
        If unitCosts(i) < 0 Then
            Err.Raise vbObjectError + 513, , "A purchase price cannot be negative."
        End If
        grandTotal = grandTotal + quantities(i) * unitCosts(i)
    Next i

    ' ---- 2. Write the header ----
    purchaseID = GenerateID("PUR")
    Set tblHeader = modDatabase.GetTable("Purchases", "tblPurchases")
    Set newRow = NextTableRow(tblHeader)
    ' Order matches tblPurchases (5 columns)
    newRow.Range.value = Array(purchaseID, supplierID, purchaseDate, grandTotal, modMain.CurrentUserID)

    ' ---- 3. Write each line, raise the stock, update the cost price ----
    Set tblLines = modDatabase.GetTable("PurchaseDetails", "tblPurchaseDetails")
    For i = LBound(productIDs) To UBound(productIDs)
        lineTotal = quantities(i) * unitCosts(i)

        Set newRow = NextTableRow(tblLines)
        ' Order matches tblPurchaseDetails (6 columns)
        newRow.Range.value = Array(GenerateID("PDT"), purchaseID, productIDs(i), _
                                   quantities(i), unitCosts(i), lineTotal)

        modInventory.IncreaseStock productIDs(i), quantities(i), "PURCHASE", purchaseID
        modProducts.SetPurchasePrice productIDs(i), unitCosts(i)
    Next i

    ' ---- 4. Update the supplier's running total ----
    modSuppliers.AddToSupplierTotal supplierID, grandTotal

    CompletePurchase = purchaseID
End Function

