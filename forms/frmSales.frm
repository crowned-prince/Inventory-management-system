VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSales 
   Caption         =   "	Sales / Point of Sale"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "frmSales.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmSales - point-of-sale screen
' ============================================================

Private mLoading As Boolean
Private mLastInvoiceNo As String

' ---------- Startup ----------
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    mLoading = True
    BuildLayout
    StyleControls
    LoadLookups
    mLoading = False

    ResetEntry
    RefreshTotals
    
    SetStatus "Choose a customer, then scan or select products."
    Exit Sub

ErrorHandler:
    mLoading = False
    MsgBox "The sales screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Sales"
End Sub

Private Function Txt(ByVal ctlName As String) As MSForms.TextBox
    Set Txt = Me.Controls(ctlName)
End Function

Private Function Cmb(ByVal ctlName As String) As MSForms.ComboBox
    Set Cmb = Me.Controls(ctlName)
End Function

Private Sub SetStatus(ByVal message As String)
    Me.Controls("lblStatusMsg").caption = message
End Sub

Private Function cur() As String
    cur = GetSetting("CurrencySymbol", "Rs.") & " "
End Function

' ---------- Build the parts made by code ----------
Private Sub BuildLayout()
    Dim headings As Variant, widths As Variant
    Dim i As Long
    Dim X As Single

    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 830, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "SALES / POINT OF SALE", -1, ThemeCyan(), 14, True

    ' Row A: sale information
    modUI.AddLabel Me, "lblSaleID", 20, 48, 110, 14, "Sale ID", -1, ThemeGrey(), 8
    modUI.AddTextBox Me, "txtSaleID", 20, 63, 110, 22
    modUI.AddLabel Me, "lblSaleDate", 145, 48, 100, 14, "Date", -1, ThemeGrey(), 8
    modUI.AddTextBox Me, "txtSaleDate", 145, 63, 100, 22
    modUI.AddLabel Me, "lblCustomer", 260, 48, 250, 14, "Customer *", -1, ThemeGrey(), 8
    modUI.AddCombo Me, "cmbCustomer", 260, 63, 250, 22
    modUI.AddLabel Me, "lblScan", 530, 48, 260, 14, "Scan barcode (adds 1 to the cart)", -1, ThemeGrey(), 8

    LockBox "txtSaleID", "AUTO"
    LockBox "txtSaleDate", Format$(Date, "dd-mm-yyyy")

    ' Row B: product entry
    modUI.AddLabel Me, "lblProduct", 20, 94, 290, 14, "Product *", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblQuantity", 320, 94, 60, 14, "Qty *", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblPrice", 390, 94, 90, 14, "Price (automatic)", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblDiscount", 490, 94, 80, 14, "Discount", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblLineTotalCap", 580, 94, 100, 14, "Line total", -1, ThemeGrey(), 8
    modUI.AddTextBox Me, "txtPrice", 390, 109, 90, 22
    LockBox "txtPrice", ""
    modUI.AddLabel Me, "lblLineTotal", 580, 109, 100, 22, "-", ThemePanel(), ThemeCyan(), 10, True

    ' Column headings (visible columns start after the hidden ID column)
    headings = Array("PRODUCT", "QTY", "PRICE", "DISCOUNT", "TOTAL")
    widths = Array(300, 70, 110, 110, 130)
    X = 23
    For i = 0 To UBound(headings)
        modUI.AddLabel Me, "lblHead" & i, X, 140, widths(i), 13, headings(i), -1, ThemeCyan(), 7, True
        X = X + widths(i)
    Next i

    ' Totals panel (sent to the back so it does not cover the discount box)
    modUI.AddLabel(Me, "lblTotalsPanel", 470, 310, 320, 96, "", ThemePanel()).ZOrder 1

    modUI.AddLabel Me, "lblItemCount", 20, 316, 430, 18, "", -1, ThemeGrey(), 9

    modUI.AddLabel Me, "lblSubtotalCap", 482, 316, 120, 18, "Subtotal", -1, ThemeGrey(), 9
    With modUI.AddLabel(Me, "lblSubtotal", 600, 316, 180, 18, "", -1, ThemeWhite(), 10, True)
        .TextAlign = fmTextAlignRight
    End With

    modUI.AddLabel Me, "lblDiscountCap", 482, 340, 190, 18, "Discount (whole invoice)", -1, ThemeGrey(), 9

    modUI.AddLabel Me, "lblTaxCap", 482, 364, 120, 18, "Tax", -1, ThemeGrey(), 9
    With modUI.AddLabel(Me, "lblTax", 600, 364, 180, 18, "", -1, ThemeWhite(), 10, True)
        .TextAlign = fmTextAlignRight
    End With

    modUI.AddLabel Me, "lblGrandCap", 482, 384, 90, 20, "TOTAL", -1, ThemeCyan(), 10, True
    With modUI.AddLabel(Me, "lblGrandTotal", 540, 380, 240, 24, "", -1, ThemeCyan(), 15, True)
        .TextAlign = fmTextAlignRight
    End With

    modUI.AddLabel Me, "lblStatusMsg", 20, 456, 770, 18, "", -1, ThemeGrey(), 9
End Sub

Private Sub LockBox(ByVal ctlName As String, ByVal initialText As String)
    With Txt(ctlName)
        .text = initialText
        .locked = True
        .BackColor = ThemeSidebar()
        .ForeColor = ThemeCyan()
    End With
End Sub

Private Sub StyleControls()
    StyleInput cmbProduct, 20, 109, 290, 22
    cmbProduct.style = fmStyleDropDownList
    StyleInput txtQuantity, 320, 109, 60, 22
    StyleInput txtLineDiscount, 490, 109, 80, 22
    StyleInput txtInvoiceDiscount, 680, 338, 100, 20
    txtInvoiceDiscount.text = "0"

    With txtBarcode
        .Left = 530: .Top = 63: .Width = 260: .Height = 22
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = ThemeCyan()
    End With

    With lstCart
        .Left = 20: .Top = 154: .Width = 770: .Height = 150
        .ColumnCount = 8
        ' 0 ID(hidden) | 1 Product | 2 Qty | 3 Price | 4 Discount | 5 Total | 6 raw price(hidden) | 7 raw discount(hidden)
        .ColumnWidths = "0 pt;300 pt;70 pt;110 pt;110 pt;130 pt;0 pt;0 pt"
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    StyleButton btnAdd, "ADD TO CART", 690, 108, 100, 24, ThemeCyan(), ThemeNavy()
    StyleButton btnRemove, "REMOVE ITEM", 20, 414, 140, 36, ThemeRed(), ThemeNavy()
    StyleButton btnClearCart, "CLEAR CART", 172, 414, 140, 36, ThemePanel(), ThemeWhite()
    StyleButton btnPrint, "PRINT INVOICE", 324, 414, 140, 36, ThemeActive(), ThemeWhite()
    StyleButton btnComplete, "COMPLETE SALE", 476, 414, 140, 36, ThemeGreen(), ThemeNavy()
    StyleButton btnCancel, "CANCEL / CLOSE", 628, 414, 140, 36, ThemePanel(), ThemeWhite()
End Sub

Private Sub StyleInput(ByVal ctl As Object, ByVal leftPos As Single, ByVal topPos As Single, _
                       ByVal widthVal As Single, ByVal heightVal As Single)
    With ctl
        .Left = leftPos: .Top = topPos: .Width = widthVal: .Height = heightVal
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
    End With
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal topPos As Single, _
                        ByVal widthVal As Single, ByVal heightVal As Single, _
                        ByVal bgColor As Long, ByVal fgColor As Long)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = topPos: .Width = widthVal: .Height = heightVal
        .BackColor = bgColor: .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = 9: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub LoadLookups()
    Dim cb As MSForms.ComboBox

    Set cb = Cmb("cmbCustomer")
    modCustomers.LoadCustomerCombo cb
    SelectWalkIn

    modProducts.LoadProductCombo cmbProduct
    cmbProduct.ListIndex = -1
End Sub

' Selects the built-in Walk-in Customer (if it exists).
Private Sub SelectWalkIn()
    Dim cb As MSForms.ComboBox
    Dim i As Long

    Set cb = Cmb("cmbCustomer")
    cb.ListIndex = -1
    For i = 0 To cb.ListCount - 1
        If CStr(cb.List(i, 0)) = modCustomers.WALK_IN_ID Then
            cb.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

' ---------- Clearing the entry row ----------
Private Sub ResetEntry()
    mLoading = True
    cmbProduct.ListIndex = -1
    txtQuantity.text = ""
    txtLineDiscount.text = ""
    Txt("txtPrice").text = ""
    txtBarcode.text = ""
    Me.Controls("lblLineTotal").caption = "-"
    mLoading = False
End Sub

' ---------- Choosing a product ----------
Private Sub cmbProduct_Change()
    If mLoading Then Exit Sub
    If cmbProduct.ListIndex < 0 Then Exit Sub
    ShowProduct CStr(cmbProduct.List(cmbProduct.ListIndex, 0))
    txtQuantity.SetFocus
End Sub

' Fills in the selling price automatically and shows the stock.
Private Sub ShowProduct(ByVal productID As String)
    Dim p As modProducts.TProduct

    If Not modProducts.LoadProduct(productID, p) Then Exit Sub

    mLoading = True
    Txt("txtPrice").text = CStr(p.sellingPrice)
    If Len(Trim$(txtQuantity.text)) = 0 Then txtQuantity.text = "1"
    mLoading = False

    UpdateLineTotal
    SetStatus p.productName & ":  in stock " & p.stock & " " & p.Unit & _
              "   |   price " & cur() & Format$(p.sellingPrice, "#,##0.00") & _
              "   |   " & modInventory.StockStatus(p.stock, p.minStock)
End Sub

' ---------- Live line total ----------
Private Sub txtQuantity_Change()
    If mLoading Then Exit Sub
    UpdateLineTotal
End Sub

Private Sub txtLineDiscount_Change()
    If mLoading Then Exit Sub
    UpdateLineTotal
End Sub

Private Sub UpdateLineTotal()
    Dim disc As Double

    If IsNumeric(txtQuantity.text) And IsNumeric(Txt("txtPrice").text) Then
        If IsNumeric(txtLineDiscount.text) Then disc = CDbl(txtLineDiscount.text) Else disc = 0
        Me.Controls("lblLineTotal").caption = cur() & Format$( _
            Round(CDbl(txtQuantity.text) * CDbl(Txt("txtPrice").text) - disc, 2), "#,##0.00")
    Else
        Me.Controls("lblLineTotal").caption = "-"
    End If
End Sub

' ---------- Barcode scanner ----------
Private Sub txtBarcode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim productID As String

    If KeyCode <> vbKeyReturn Then Exit Sub

    productID = modProducts.FindProductByBarcode(txtBarcode.text)
    If Len(productID) = 0 Then
        SetStatus "No product found with barcode " & Trim$(txtBarcode.text) & "."
    Else
        TryAddToCart productID, 1, 0           ' a scan = 1 unit, no discount
    End If
    txtBarcode.text = ""
End Sub

' ---------- Cart helpers ----------
' How many units of a product are already in the cart.
Private Function QtyInCart(ByVal productID As String) As Double
    Dim i As Long
    For i = 0 To lstCart.ListCount - 1
        If CStr(lstCart.List(i, 0)) = productID Then QtyInCart = QtyInCart + CDbl(lstCart.List(i, 2))
    Next i
End Function

' Checks stock and adds the product to the cart. Returns True if added.
Private Function TryAddToCart(ByVal productID As String, ByVal qty As Double, _
                              ByVal lineDiscount As Double) As Boolean
    Dim p As modProducts.TProduct
    Dim inCart As Double, available As Double
    Dim price As Double, lineTotal As Double
    Dim i As Long, n As Long
    Dim msg As String
    Dim newQty As Double

    If Not modProducts.LoadProduct(productID, p) Then
        MsgBox "That product could not be found.", vbExclamation, "Add to cart"
        Exit Function
    End If
    If p.status <> "Active" Then
        MsgBox p.productName & " is not active, so it cannot be sold.", vbExclamation, "Add to cart"
        Exit Function
    End If

    ' ---- stock check ----
    inCart = QtyInCart(productID)
    available = p.stock - inCart
    If p.stock <= 0 Then
        MsgBox p.productName & " is out of stock.", vbExclamation, "Out of stock"
        Exit Function
    End If
    If qty > available Then
        If available < 0 Then available = 0
        msg = "Insufficient stock. Only " & available & " units are available."
        If inCart > 0 Then msg = msg & vbCrLf & "(" & inCart & " already in the cart.)"
        MsgBox msg, vbExclamation, "Insufficient stock"
        Exit Function
    End If

    ' ---- discount check ----
    price = p.sellingPrice
    lineTotal = Round(qty * price, 2)
    If lineDiscount < 0 Or lineDiscount > lineTotal Then
        MsgBox "The discount cannot be more than the line total.", vbExclamation, "Discount"
        Exit Function
    End If

    ' Same product, same price, no discounts? Merge into the existing line.
    If lineDiscount = 0 Then
        For i = 0 To lstCart.ListCount - 1
            If CStr(lstCart.List(i, 0)) = productID And CDbl(lstCart.List(i, 6)) = price _
               And CDbl(lstCart.List(i, 7)) = 0 Then
                newQty = CDbl(lstCart.List(i, 2)) + qty
                lstCart.List(i, 2) = CStr(newQty)
                lstCart.List(i, 5) = Format$(Round(newQty * price, 2), "#,##0.00")
                GoTo Finished
            End If
        Next i
    End If

    lstCart.AddItem productID
    n = lstCart.ListCount - 1
    lstCart.List(n, 1) = p.productName
    lstCart.List(n, 2) = CStr(qty)
    lstCart.List(n, 3) = Format$(price, "#,##0.00")
    lstCart.List(n, 4) = Format$(lineDiscount, "#,##0.00")
    lstCart.List(n, 5) = Format$(Round(lineTotal - lineDiscount, 2), "#,##0.00")
    lstCart.List(n, 6) = CStr(price)               ' hidden raw values used when saving
    lstCart.List(n, 7) = CStr(lineDiscount)

Finished:
    ResetEntry
    RefreshTotals
    SetStatus "Added " & p.productName & ".  Scan or choose the next product, or click COMPLETE SALE."
    TryAddToCart = True
End Function

' ---------- ADD TO CART button ----------
Private Sub btnAdd_Click()
    Dim msg As String
    Dim qty As Double, disc As Double

    If cmbProduct.ListIndex < 0 Then
        MsgBox "Please select a product.", vbExclamation, "Add to cart"
        Exit Sub
    End If

    msg = CheckNumber(txtQuantity.text, "quantity")
    If Len(msg) = 0 Then
        If CDbl(txtQuantity.text) <= 0 Then msg = "The quantity must be greater than zero."
    End If
    If Len(msg) = 0 And Len(Trim$(txtLineDiscount.text)) > 0 Then
        msg = CheckNumber(txtLineDiscount.text, "discount")
    End If
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Please check your input"
        Exit Sub
    End If

    qty = CDbl(txtQuantity.text)
    If Len(Trim$(txtLineDiscount.text)) > 0 Then disc = CDbl(txtLineDiscount.text) Else disc = 0

    TryAddToCart CStr(cmbProduct.List(cmbProduct.ListIndex, 0)), qty, disc
    txtBarcode.SetFocus
End Sub

' ---------- Totals ----------
' Sum of all line totals (qty x price - line discount), from the hidden raw values.
Private Function CartSubtotal() As Double
    Dim i As Long
    For i = 0 To lstCart.ListCount - 1
        CartSubtotal = CartSubtotal + Round(CDbl(lstCart.List(i, 2)) * CDbl(lstCart.List(i, 6)) - _
                                            CDbl(lstCart.List(i, 7)), 2)
    Next i
End Function

Private Sub txtInvoiceDiscount_Change()
    If mLoading Then Exit Sub
    RefreshTotals
End Sub

Private Sub RefreshTotals()
    Dim i As Long
    Dim Units As Double
    Dim subtotal As Double, disc As Double, tax As Double, grand As Double

    For i = 0 To lstCart.ListCount - 1
        Units = Units + CDbl(lstCart.List(i, 2))
    Next i

    subtotal = CartSubtotal()

    ' Show a sensible discount even while the user is still typing
    If IsNumeric(txtInvoiceDiscount.text) Then disc = CDbl(txtInvoiceDiscount.text) Else disc = 0
    If disc < 0 Then disc = 0
    If disc > subtotal Then disc = subtotal

    tax = modSales.CalculateTax(subtotal - disc)
    grand = Round(subtotal - disc + tax, 2)

    Me.Controls("lblItemCount").caption = "Items: " & lstCart.ListCount & "     Units: " & Units
    Me.Controls("lblSubtotal").caption = cur() & Format$(subtotal, "#,##0.00")
    Me.Controls("lblTax").caption = cur() & Format$(tax, "#,##0.00")
    Me.Controls("lblGrandTotal").caption = cur() & Format$(grand, "#,##0.00")

    If modSales.TaxEnabled() Then
        Me.Controls("lblTaxCap").caption = "Tax (" & modSales.TaxRatePercent() & "%)"
    Else
        Me.Controls("lblTaxCap").caption = "Tax (off)"
    End If
End Sub

' ---------- Remove / clear ----------
Private Sub btnRemove_Click()
    If lstCart.ListIndex < 0 Then
        MsgBox "Please select an item in the cart first.", vbInformation, "Remove item"
        Exit Sub
    End If
    lstCart.RemoveItem lstCart.ListIndex
    RefreshTotals
End Sub

Private Sub btnClearCart_Click()
    If lstCart.ListCount = 0 Then Exit Sub
    If MsgBox("Remove all items from the cart?", vbYesNo + vbQuestion, "Clear cart") = vbYes Then
        lstCart.Clear
        txtInvoiceDiscount.text = "0"
        RefreshTotals
    End If
End Sub

' ---------- Completing the sale ----------
Private Sub btnComplete_Click()
    Dim customerID As String
    Dim ids() As String, qtys() As Double, prices() As Double, discs() As Double
    Dim i As Long, n As Long
    Dim invDisc As Double
    Dim subtotal As Double
    Dim saleID As String, invoiceNo As String
    Dim grandTotal As Double
    Dim warning As String
    Dim msg As String

    On Error GoTo ErrorHandler

    ' --- check the form ---
    If Cmb("cmbCustomer").ListIndex < 0 Then
        MsgBox "Please select a customer." & vbCrLf & _
               "(If the Walk-in Customer is missing, run EnsureWalkInCustomer once.)", _
               vbExclamation, "Complete sale"
        Exit Sub
    End If
    customerID = CStr(Cmb("cmbCustomer").List(Cmb("cmbCustomer").ListIndex, 0))

    n = lstCart.ListCount
    If n = 0 Then
        MsgBox "The cart is empty. Add at least one product first.", vbExclamation, "Complete sale"
        Exit Sub
    End If

    If Len(Trim$(txtInvoiceDiscount.text)) = 0 Then txtInvoiceDiscount.text = "0"
    msg = CheckNumber(txtInvoiceDiscount.text, "discount")
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Complete sale"
        Exit Sub
    End If
    invDisc = CDbl(txtInvoiceDiscount.text)

    subtotal = CartSubtotal()
    If invDisc > subtotal Then
        MsgBox "The discount cannot be more than the subtotal.", vbExclamation, "Complete sale"
        Exit Sub
    End If

    ' --- copy the cart into arrays ---
    ReDim ids(1 To n)
    ReDim qtys(1 To n)
    ReDim prices(1 To n)
    ReDim discs(1 To n)
    For i = 1 To n
        ids(i) = CStr(lstCart.List(i - 1, 0))
        qtys(i) = CDbl(lstCart.List(i - 1, 2))
        prices(i) = CDbl(lstCart.List(i - 1, 6))
        discs(i) = CDbl(lstCart.List(i - 1, 7))
    Next i

    If MsgBox("Complete this sale?" & vbCrLf & vbCrLf & _
              "Customer:  " & Cmb("cmbCustomer").text & vbCrLf & _
              "Items:  " & n & vbCrLf & _
              "Total:  " & Me.Controls("lblGrandTotal").caption & vbCrLf & vbCrLf & _
              "Stock will be reduced automatically.", _
              vbYesNo + vbQuestion, "Confirm sale") <> vbYes Then Exit Sub

    ' --- save ---
    saleID = modSales.CompleteSale(customerID, ids, qtys, prices, discs, invDisc, invoiceNo, grandTotal)
    LogAction "SALE", saleID, "Completed sale " & invoiceNo & " for " & cur() & Format$(grandTotal, "#,##0.00")
    modApp.SaveNow
    mLastInvoiceNo = invoiceNo
    MsgBox "Sale " & saleID & " completed." & vbCrLf & _
           "Invoice:  " & invoiceNo & vbCrLf & _
           "Total:  " & cur() & Format$(grandTotal, "#,##0.00"), _
           vbInformation, "Sale complete"

    ' --- get ready for the next customer ---
    lstCart.Clear
    ResetEntry
    mLoading = True
    txtInvoiceDiscount.text = "0"
    mLoading = False
    SelectWalkIn
    modProducts.LoadProductCombo cmbProduct         ' (names unchanged, but keeps the list fresh)
    cmbProduct.ListIndex = -1
    RefreshTotals
    btnPrint.Enabled = True
        warning = modAlerts.GetLowStockWarning(ids)
    If Len(warning) > 0 Then
        SetStatus "Sale " & saleID & " saved (" & invoiceNo & ").   " & warning
    Else
        SetStatus "Sale " & saleID & " saved (" & invoiceNo & "). Ready for the next customer."
    End If
    txtBarcode.SetFocus
    Exit Sub

ErrorHandler:
    MsgBox "The sale could not be completed." & vbCrLf & Err.description & vbCrLf & vbCrLf & _
           "Please check the Sales and Products sheets before trying again.", _
           vbExclamation, "Sale failed"
End Sub

Private Sub btnPrint_Click()
    Dim invoiceNo As String
    Dim choice As VbMsgBoxResult
    Dim filePath As String

    On Error GoTo ErrorHandler

    ' Offers the last invoice by default; type another number to reprint an older one
    invoiceNo = Trim$(InputBox("Enter the invoice number to print (for example INV-00001):", _
                               "Print invoice", mLastInvoiceNo))
    If Len(invoiceNo) = 0 Then Exit Sub

    If Not modInvoice.InvoiceExists(invoiceNo) Then
        MsgBox "Invoice " & invoiceNo & " was not found.", vbExclamation, "Print invoice"
        Exit Sub
    End If

    choice = MsgBox("Invoice " & invoiceNo & vbCrLf & vbCrLf & _
                    "Yes  =  print on paper" & vbCrLf & _
                    "No  =  save as PDF" & vbCrLf & _
                    "Cancel  =  do nothing", _
                    vbYesNoCancel + vbQuestion, "Print invoice")

    If choice = vbYes Then
        modInvoice.PrintInvoice invoiceNo
        LogAction "PRINT INVOICE", invoiceNo, "Invoice printed"
        SetStatus "Invoice " & invoiceNo & " was sent to the printer."
    ElseIf choice = vbNo Then
        filePath = modInvoice.SaveInvoicePdf(invoiceNo)
        LogAction "PDF INVOICE", invoiceNo, "Saved " & filePath
        SetStatus "Invoice saved as PDF: " & filePath
        If MsgBox("The PDF was saved here:" & vbCrLf & filePath & vbCrLf & vbCrLf & "Open it now?", _
                  vbYesNo + vbQuestion, "Invoice saved") = vbYes Then
            modInvoice.OpenFile filePath
        End If
    End If
    Exit Sub

ErrorHandler:
    MsgBox "The invoice could not be produced." & vbCrLf & Err.description, vbExclamation, "Print invoice"
End Sub
' ---------- Closing ----------
Private Sub btnCancel_Click()
    If lstCart.ListCount > 0 Then
        If MsgBox("The cart has items that are not sold. Cancel this sale and close?", _
                  vbYesNo + vbQuestion, "Cancel sale") = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu And lstCart.ListCount > 0 Then
        If MsgBox("The cart has items that are not sold. Cancel this sale and close?", _
                  vbYesNo + vbQuestion, "Cancel sale") = vbNo Then Cancel = True
    End If
End Sub

