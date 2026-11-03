VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPurchase 
   Caption         =   "Purchase / Stock-In"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "frmPurchase.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmPurchase - record goods received from a supplier
' ============================================================

Private mLoading As Boolean

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
    SetStatus "Choose a supplier, then add products to the list."
    Exit Sub

ErrorHandler:
    mLoading = False
    MsgBox "The purchase screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Purchases"
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

' ---------- Build the screen ----------
Private Sub BuildLayout()
    Dim headings As Variant, widths As Variant
    Dim i As Long
    Dim X As Single

    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 830, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "PURCHASE / STOCK-IN", -1, ThemeCyan(), 14, True

    ' Row A: purchase information
    modUI.AddLabel Me, "lblPurchaseID", 20, 48, 130, 14, "Purchase ID", -1, ThemeGrey(), 8
    modUI.AddTextBox Me, "txtPurchaseID", 20, 63, 130, 22
    modUI.AddLabel Me, "lblPurchaseDate", 165, 48, 110, 14, "Date (dd-mm-yyyy)", -1, ThemeGrey(), 8
    modUI.AddTextBox Me, "txtPurchaseDate", 165, 63, 110, 22
    modUI.AddLabel Me, "lblSupplier", 290, 48, 250, 14, "Supplier *", -1, ThemeGrey(), 8
    modUI.AddCombo Me, "cmbSupplier", 290, 63, 250, 22
    modUI.AddLabel Me, "lblScan", 560, 48, 230, 14, "Scan barcode, then press Enter", -1, ThemeGrey(), 8

    With Txt("txtPurchaseID")
        .text = "AUTO": .locked = True
        .BackColor = ThemeSidebar(): .ForeColor = ThemeCyan()
    End With
    Txt("txtPurchaseDate").text = Format$(Date, "dd-mm-yyyy")

    ' Row B: product entry
    modUI.AddLabel Me, "lblProduct", 20, 94, 330, 14, "Product *", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblQuantity", 365, 94, 80, 14, "Quantity *", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblPrice", 460, 94, 100, 14, "Purchase price *", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblLineTotalCap", 575, 94, 110, 14, "Line total", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblLineTotal", 575, 109, 110, 22, "-", ThemePanel(), ThemeCyan(), 10, True

    ' Column headings above the cart (widths match StyleControls)
    headings = Array("ID", "PRODUCT", "QTY", "UNIT COST", "LINE TOTAL")
    widths = Array(70, 330, 80, 110, 130)
    X = 23
    For i = 0 To UBound(headings)
        modUI.AddLabel Me, "lblHead" & i, X, 140, widths(i), 13, headings(i), -1, ThemeCyan(), 7, True
        X = X + widths(i)
    Next i

    ' Totals
    modUI.AddLabel Me, "lblItemCount", 20, 380, 400, 18, "", -1, ThemeGrey(), 9
    With modUI.AddLabel(Me, "lblGrandTotal", 420, 376, 370, 26, "", -1, ThemeCyan(), 16, True)
        .TextAlign = fmTextAlignRight
    End With

    modUI.AddLabel Me, "lblStatusMsg", 20, 464, 770, 18, "", -1, ThemeGrey(), 9
End Sub

Private Sub StyleControls()
    StyleInput cmbProduct, 20, 109, 330, 22
    cmbProduct.style = fmStyleDropDownList
    StyleInput txtQuantity, 365, 109, 80, 22
    StyleInput txtPrice, 460, 109, 100, 22

    With txtBarcode
        .Left = 560: .Top = 63: .Width = 230: .Height = 22
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = ThemeCyan()
    End With

    With lstCart
        .Left = 20: .Top = 154: .Width = 770: .Height = 216
        .ColumnCount = 6
        .ColumnWidths = "70 pt;330 pt;80 pt;110 pt;130 pt;0 pt"      ' last column = hidden raw price
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    StyleButton btnAdd, "ADD ITEM", 695, 108, 95, 24, ThemeCyan(), ThemeNavy()
    StyleButton btnRemove, "REMOVE ITEM", 20, 414, 140, 40, ThemeRed(), ThemeNavy()
    StyleButton btnClearCart, "CLEAR LIST", 170, 414, 140, 40, ThemePanel(), ThemeWhite()
    StyleButton btnComplete, "COMPLETE PURCHASE", 460, 414, 220, 40, ThemeGreen(), ThemeNavy()
    StyleButton btnClose, "CLOSE", 690, 414, 100, 40, ThemePanel(), ThemeWhite()
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
        .Font.name = "Segoe UI": .Font.Size = 10: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub LoadLookups()
    Dim cb As MSForms.ComboBox

    Set cb = Cmb("cmbSupplier")
    modProducts.LoadSupplierCombo cb
    cb.ListIndex = -1

    modProducts.LoadProductCombo cmbProduct
    cmbProduct.ListIndex = -1
End Sub

' ---------- Clearing the entry row ----------
Private Sub ResetEntry()
    mLoading = True
    cmbProduct.ListIndex = -1
    txtQuantity.text = ""
    txtPrice.text = ""
    txtBarcode.text = ""
    Me.Controls("lblLineTotal").caption = "-"
    mLoading = False
End Sub

' ---------- Picking a product ----------
Private Sub cmbProduct_Change()
    If mLoading Then Exit Sub
    If cmbProduct.ListIndex < 0 Then Exit Sub
    ShowProduct CStr(cmbProduct.List(cmbProduct.ListIndex, 0))
    txtQuantity.SetFocus
End Sub

' Fills in the last purchase price and shows the stock level.
Private Sub ShowProduct(ByVal productID As String)
    Dim p As modProducts.TProduct

    If Not modProducts.LoadProduct(productID, p) Then Exit Sub

    mLoading = True
    txtPrice.text = CStr(p.PurchasePrice)
    If Len(Trim$(txtQuantity.text)) = 0 Then txtQuantity.text = "1"
    mLoading = False

    UpdateLineTotal
    SetStatus p.productName & ":  in stock " & p.stock & " " & p.Unit & _
              "   |   last purchase price " & Format$(p.PurchasePrice, "#,##0.00")
End Sub

' Barcode scanners type the code and press Enter.
Private Sub txtBarcode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim productID As String
    Dim i As Long

    If KeyCode <> vbKeyReturn Then Exit Sub

    productID = modProducts.FindProductByBarcode(txtBarcode.text)
    If Len(productID) = 0 Then
        SetStatus "No product found with barcode " & Trim$(txtBarcode.text) & "."
        txtBarcode.text = ""
        Exit Sub
    End If

    For i = 0 To cmbProduct.ListCount - 1
        If CStr(cmbProduct.List(i, 0)) = productID Then
            cmbProduct.ListIndex = -1          ' reset first so the Change event always fires
            cmbProduct.ListIndex = i
            txtBarcode.text = ""
            Exit Sub
        End If
    Next i

    SetStatus "That product is not active, so it cannot be purchased."
    txtBarcode.text = ""
End Sub

' ---------- Live line total ----------
Private Sub txtQuantity_Change()
    If mLoading Then Exit Sub
    UpdateLineTotal
End Sub

Private Sub txtPrice_Change()
    If mLoading Then Exit Sub
    UpdateLineTotal
End Sub

Private Sub UpdateLineTotal()
    Dim cur As String
    cur = GetSetting("CurrencySymbol", "Rs.") & " "

    If IsNumeric(txtQuantity.text) And IsNumeric(txtPrice.text) Then
        Me.Controls("lblLineTotal").caption = cur & Format$(CDbl(txtQuantity.text) * CDbl(txtPrice.text), "#,##0.00")
    Else
        Me.Controls("lblLineTotal").caption = "-"
    End If
End Sub

' ---------- Adding an item to the list ----------
Private Sub btnAdd_Click()
    Dim msg As String
    Dim productID As String, productName As String
    Dim qty As Double, cost As Double
    Dim i As Long, n As Long
    Dim newQty As Double

    If cmbProduct.ListIndex < 0 Then
        MsgBox "Please select a product.", vbExclamation, "Add item"
        Exit Sub
    End If

    msg = CheckNumber(txtQuantity.text, "quantity")
    If Len(msg) = 0 Then
        If CDbl(txtQuantity.text) <= 0 Then msg = "The quantity must be greater than zero."
    End If
    If Len(msg) = 0 Then msg = CheckNumber(txtPrice.text, "purchase price")
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Please check your input"
        Exit Sub
    End If

    productID = CStr(cmbProduct.List(cmbProduct.ListIndex, 0))
    productName = CStr(cmbProduct.List(cmbProduct.ListIndex, 1))
    qty = CDbl(txtQuantity.text)
    cost = CDbl(txtPrice.text)

    ' Same product at the same cost? Merge into the existing line.
    For i = 0 To lstCart.ListCount - 1
        If CStr(lstCart.List(i, 0)) = productID And CDbl(lstCart.List(i, 5)) = cost Then
            newQty = CDbl(lstCart.List(i, 2)) + qty
            lstCart.List(i, 2) = CStr(newQty)
            lstCart.List(i, 4) = Format$(newQty * cost, "#,##0.00")
            GoTo Finished
        End If
    Next i

    lstCart.AddItem productID
    n = lstCart.ListCount - 1
    lstCart.List(n, 1) = productName
    lstCart.List(n, 2) = CStr(qty)
    lstCart.List(n, 3) = Format$(cost, "#,##0.00")
    lstCart.List(n, 4) = Format$(qty * cost, "#,##0.00")
    lstCart.List(n, 5) = CStr(cost)             ' hidden raw value used when saving

Finished:
    ResetEntry
    RefreshTotals
    SetStatus "Added " & productName & ". Scan or choose the next product, or click COMPLETE PURCHASE."
    txtBarcode.SetFocus
End Sub

' ---------- Totals ----------
Private Sub RefreshTotals()
    Dim i As Long
    Dim Units As Double, total As Double
    Dim cur As String

    cur = GetSetting("CurrencySymbol", "Rs.") & " "

    For i = 0 To lstCart.ListCount - 1
        Units = Units + CDbl(lstCart.List(i, 2))
        total = total + CDbl(lstCart.List(i, 2)) * CDbl(lstCart.List(i, 5))
    Next i

    Me.Controls("lblItemCount").caption = "Items: " & lstCart.ListCount & "     Units: " & Units
    Me.Controls("lblGrandTotal").caption = "TOTAL COST:  " & cur & Format$(total, "#,##0.00")
End Sub

' ---------- Remove / clear ----------
Private Sub btnRemove_Click()
    If lstCart.ListIndex < 0 Then
        MsgBox "Please select an item in the list first.", vbInformation, "Remove item"
        Exit Sub
    End If
    lstCart.RemoveItem lstCart.ListIndex
    RefreshTotals
End Sub

Private Sub btnClearCart_Click()
    If lstCart.ListCount = 0 Then Exit Sub
    If MsgBox("Remove all items from the list?", vbYesNo + vbQuestion, "Clear list") = vbYes Then
        lstCart.Clear
        RefreshTotals
    End If
End Sub

' ---------- Completing the purchase ----------
Private Sub btnComplete_Click()
    Dim supplierID As String
    Dim purchaseDate As Date
    Dim ids() As String, qtys() As Double, costs() As Double
    Dim i As Long, n As Long
    Dim total As Double
    Dim purchaseID As String
    Dim cur As String

    On Error GoTo ErrorHandler

    cur = GetSetting("CurrencySymbol", "Rs.") & " "

    ' --- check the form ---
    If Cmb("cmbSupplier").ListIndex < 0 Then
        MsgBox "Please select a supplier.", vbExclamation, "Complete purchase"
        Exit Sub
    End If
    supplierID = CStr(Cmb("cmbSupplier").List(Cmb("cmbSupplier").ListIndex, 0))
    If Len(supplierID) = 0 Then
        MsgBox "Please select a supplier.", vbExclamation, "Complete purchase"
        Exit Sub
    End If

    If Not modValidation.TryParseDate(Txt("txtPurchaseDate").text, purchaseDate) Then
        MsgBox "Please enter the date as dd-mm-yyyy (for example 31-12-2026).", vbExclamation, "Complete purchase"
        Exit Sub
    End If
    If purchaseDate > Date Then
        MsgBox "The purchase date cannot be in the future.", vbExclamation, "Complete purchase"
        Exit Sub
    End If

    n = lstCart.ListCount
    If n = 0 Then
        MsgBox "The list is empty. Add at least one product first.", vbExclamation, "Complete purchase"
        Exit Sub
    End If

    ' --- copy the list into arrays ---
    ReDim ids(1 To n)
    ReDim qtys(1 To n)
    ReDim costs(1 To n)
    For i = 1 To n
        ids(i) = CStr(lstCart.List(i - 1, 0))
        qtys(i) = CDbl(lstCart.List(i - 1, 2))
        costs(i) = CDbl(lstCart.List(i - 1, 5))
        total = total + qtys(i) * costs(i)
    Next i

    If MsgBox("Complete this purchase?" & vbCrLf & vbCrLf & _
              "Supplier:  " & Cmb("cmbSupplier").text & vbCrLf & _
              "Items:  " & n & vbCrLf & _
              "Total:  " & cur & Format$(total, "#,##0.00") & vbCrLf & vbCrLf & _
              "Stock will be increased automatically.", _
              vbYesNo + vbQuestion, "Confirm purchase") <> vbYes Then Exit Sub

    ' --- save ---
    purchaseID = modPurchases.CompletePurchase(supplierID, purchaseDate, ids, qtys, costs)
    LogAction "PURCHASE", purchaseID, "Purchase of " & cur & Format$(total, "#,##0.00") & _
              " from " & Cmb("cmbSupplier").text
    modApp.SaveNow
    MsgBox "Purchase " & purchaseID & " saved." & vbCrLf & "Stock has been updated.", _
           vbInformation, "Purchase complete"

    lstCart.Clear
    ResetEntry
    RefreshTotals
    Cmb("cmbSupplier").ListIndex = -1
    Txt("txtPurchaseDate").text = Format$(Date, "dd-mm-yyyy")
    SetStatus "Saved " & purchaseID & ". Ready for the next purchase."
    Exit Sub

ErrorHandler:
    MsgBox "The purchase could not be completed." & vbCrLf & Err.description & vbCrLf & vbCrLf & _
           "Please check the Products and Purchases sheets before trying again.", _
           vbExclamation, "Purchase failed"
End Sub

' ---------- Closing ----------
Private Sub btnClose_Click()
    If lstCart.ListCount > 0 Then
        If MsgBox("The list has items that are not saved. Close anyway?", _
                  vbYesNo + vbQuestion, "Close") = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

' Also protect against the window's X button.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu And lstCart.ListCount > 0 Then
        If MsgBox("The list has items that are not saved. Close anyway?", _
                  vbYesNo + vbQuestion, "Close") = vbNo Then Cancel = True
    End If
End Sub

