VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProducts 
   Caption         =   "Product Management"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "frmProducts.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmProducts - add, update and delete products
' ============================================================

Private mSelectedID As String      ' "" = new product, otherwise the product being edited
Private mLoading As Boolean        ' stops list events from firing while we fill it

Private Sub btnSearchbtnSearch_Click()

End Sub

' ---------- Startup ----------
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    BuildLayout
    StyleButtonsAndList
    LoadCombos
    RefreshList
    ClearForm
    Exit Sub

ErrorHandler:
    MsgBox "The product screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Products"
End Sub

' ---------- Shortcuts to find controls created by code ----------
Private Function Txt(ByVal ctlName As String) As MSForms.TextBox
    Set Txt = Me.Controls(ctlName)
End Function

Private Function Cmb(ByVal ctlName As String) As MSForms.ComboBox
    Set Cmb = Me.Controls(ctlName)
End Function

' ---------- Build the screen ----------
Private Sub BuildLayout()
    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 830, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "PRODUCT MANAGEMENT", -1, ThemeCyan(), 14, True

    ' Left column
    AddField "Product ID", "txtProductID", "text", 1, 0
    AddField "Barcode", "txtBarcode", "text", 1, 1
    AddField "Product Name *", "txtProductName", "text", 1, 2
    AddField "Category *", "cmbCategory", "combo", 1, 3
    AddField "Brand", "txtBrand", "text", 1, 4
    AddField "Supplier", "cmbSupplier", "combo", 1, 5
    AddField "Unit *", "cmbUnit", "combo", 1, 6
    AddField "Status *", "cmbStatus", "combo", 1, 7

    ' Right column
    AddField "Purchase Price *", "txtPurchasePrice", "text", 2, 0
    AddField "Selling Price *", "txtSellingPrice", "text", 2, 1
    AddField "Quantity (opening stock)", "txtQuantity", "text", 2, 2
    AddField "Minimum Stock Level *", "txtMinStock", "text", 2, 3
    AddField "Maximum Stock Level *", "txtMaxStock", "text", 2, 4
    AddField "Expiry Date (dd-mm-yyyy)", "txtExpiryDate", "text", 2, 5
    AddField "Description", "txtDescription", "text", 2, 6

    With Txt("txtDescription")
        .Height = 66
        .Multiline = True
        .WordWrap = True
        .EnterKeyBehavior = True
    End With

    With Txt("txtProductID")
        .locked = True
        .ForeColor = ThemeCyan()
    End With

    Txt("txtProductName").MaxLength = 100

    modUI.AddLabel Me, "lblListTitle", 520, 56, 270, 14, "PRODUCT LIST  (click a row to edit)", -1, ThemeCyan(), 8, True
    modUI.AddLabel Me, "lblStatusMsg", 20, 468, 770, 18, "", -1, ThemeGrey(), 9
End Sub

' Adds a caption label + an input box/combo in a 2-column grid.
Private Sub AddField(ByVal captionText As String, ByVal ctlName As String, _
                     ByVal kind As String, ByVal columnNo As Long, ByVal rowNo As Long)
    Dim X As Single, Y As Single

    If columnNo = 1 Then X = 20 Else X = 270
    Y = 56 + (rowNo * 44)

    modUI.AddLabel Me, "lbl" & Mid$(ctlName, 4), X, Y, 230, 14, captionText, -1, ThemeGrey(), 8

    If kind = "combo" Then
        modUI.AddCombo Me, ctlName, X, Y + 15, 230, 22
    Else
        modUI.AddTextBox Me, ctlName, X, Y + 15, 230, 22
    End If
End Sub

Private Sub StyleButtonsAndList()
    StyleButton btnSave, "SAVE", 20, ThemeGreen(), ThemeNavy()
    StyleButton btnUpdate, "UPDATE", 150, ThemeCyan(), ThemeNavy()
    StyleButton btnDelete, "DELETE", 280, ThemeRed(), ThemeNavy()
    StyleButton btnSearch, "SEARCH", 410, ThemeActive(), ThemeWhite()
    StyleButton btnClear, "NEW / CLEAR", 540, ThemePanel(), ThemeWhite()
    StyleButton btnClose, "CLOSE", 670, ThemePanel(), ThemeWhite()

    With lstProducts
        .Left = 520: .Top = 72: .Width = 270: .Height = 329
        .BackColor = ThemePanel()
        .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal bgColor As Long, ByVal fgColor As Long)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = 416: .Width = 120: .Height = 40
        .BackColor = bgColor
        .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = 10: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

' ---------- Dropdown contents ----------
Private Sub LoadCombos()
    Dim cb As MSForms.ComboBox
    Dim item As Variant

    Set cb = Cmb("cmbCategory")
    modProducts.LoadCategoryCombo cb

    Set cb = Cmb("cmbSupplier")
    modProducts.LoadSupplierCombo cb

    Set cb = Cmb("cmbUnit")
    For Each item In Array("Pcs", "Kg", "Gram", "Litre", "Box", "Pack", "Dozen", "Bottle", "Carton")
        cb.AddItem CStr(item)
    Next item

    Set cb = Cmb("cmbStatus")
    cb.AddItem "Active"
    cb.AddItem "Inactive"
End Sub

' Selects the combo row whose text (or hidden ID) in the given column matches.
Private Sub SelectCombo(ByVal cb As MSForms.ComboBox, ByVal searchValue As String, ByVal columnNo As Long)
    Dim i As Long
    cb.ListIndex = -1
    For i = 0 To cb.ListCount - 1
        If StrComp(CStr(cb.List(i, columnNo)), searchValue, vbTextCompare) = 0 Then
            cb.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Sub RefreshList()
    mLoading = True
    modProducts.FillProductList lstProducts
    mLoading = False
End Sub

' ---------- Reset the form to "new product" mode ----------
Private Sub ClearForm()
    mSelectedID = ""
    mLoading = True

    Txt("txtProductID").text = "AUTO"
    Txt("txtBarcode").text = ""
    Txt("txtProductName").text = ""
    Txt("txtBrand").text = ""
    Txt("txtPurchasePrice").text = ""
    Txt("txtSellingPrice").text = ""
    Txt("txtQuantity").text = "0"
    Txt("txtMinStock").text = "10"
    Txt("txtMaxStock").text = "100"
    Txt("txtExpiryDate").text = ""
    Txt("txtDescription").text = ""

    Cmb("cmbCategory").ListIndex = -1
    Cmb("cmbSupplier").ListIndex = 0
    SelectCombo Cmb("cmbUnit"), "Pcs", 0
    SelectCombo Cmb("cmbStatus"), "Active", 0
    lstProducts.ListIndex = -1

    Txt("txtQuantity").locked = False
    mLoading = False

    UpdateButtonState
    SetStatus "New product. Fill in the fields marked * and click SAVE."
    Txt("txtBarcode").SetFocus
End Sub

Private Sub SetStatus(ByVal message As String)
    Me.Controls("lblStatusMsg").caption = message
End Sub

' ---------- Which buttons are allowed right now ----------
Private Sub UpdateButtonState()
    Dim canEdit As Boolean, isNew As Boolean

    canEdit = (modMain.CurrentRole <> "STAFF")
    isNew = (Len(mSelectedID) = 0)

    btnSave.Enabled = canEdit And isNew
    btnUpdate.Enabled = canEdit And Not isNew
    btnDelete.Enabled = modMain.IsAdmin() And Not isNew

    If Not canEdit Then SetStatus "View only: your role cannot change products."
End Sub

' ---------- Clicking a product in the list ----------
Private Sub lstProducts_Click()
    Dim p As modProducts.TProduct

    If mLoading Then Exit Sub
    If lstProducts.ListIndex < 0 Then Exit Sub

    If modProducts.LoadProduct(CStr(lstProducts.List(lstProducts.ListIndex, 0)), p) Then
        FillForm p
    End If
End Sub

Private Sub FillForm(ByRef p As modProducts.TProduct)
    mLoading = True
    mSelectedID = p.productID

    Txt("txtProductID").text = p.productID
    Txt("txtBarcode").text = p.barcode
    Txt("txtProductName").text = p.productName
    Txt("txtBrand").text = p.brand
    Txt("txtPurchasePrice").text = CStr(p.PurchasePrice)
    Txt("txtSellingPrice").text = CStr(p.sellingPrice)
    Txt("txtQuantity").text = CStr(p.stock)
    Txt("txtMinStock").text = CStr(p.minStock)
    Txt("txtMaxStock").text = CStr(p.MaxStock)
    Txt("txtExpiryDate").text = p.ExpiryDate
    Txt("txtDescription").text = p.description

    SelectCombo Cmb("cmbCategory"), p.category, 0
    SelectCombo Cmb("cmbSupplier"), p.supplierID, 0
    SelectCombo Cmb("cmbUnit"), p.Unit, 0
    SelectCombo Cmb("cmbStatus"), p.status, 0

    Txt("txtQuantity").locked = True          ' stock changes only via purchases/sales
    mLoading = False

    UpdateButtonState
    SetStatus "Editing " & p.productID & ". Stock is changed through Purchases and Sales."
End Sub

' ---------- Validation: returns "" if all OK, else the message ----------
Private Function ValidateForm() As String
    Dim msg As String

    msg = CheckRequired(Txt("txtProductName").text, "product name")
    If Len(msg) > 0 Then ValidateForm = msg: Exit Function

    If Cmb("cmbCategory").ListIndex < 0 Then
        ValidateForm = "Please select a category."
        Exit Function
    End If
    If Cmb("cmbUnit").ListIndex < 0 Then
        ValidateForm = "Please select a unit."
        Exit Function
    End If
    If Cmb("cmbStatus").ListIndex < 0 Then
        ValidateForm = "Please select a status."
        Exit Function
    End If

    msg = CheckNumber(Txt("txtPurchasePrice").text, "purchase price")
    If Len(msg) > 0 Then ValidateForm = msg: Exit Function

    msg = CheckNumber(Txt("txtSellingPrice").text, "selling price")
    If Len(msg) > 0 Then ValidateForm = msg: Exit Function

    msg = CheckNumber(Txt("txtQuantity").text, "quantity")
    If Len(msg) > 0 Then ValidateForm = msg: Exit Function

    msg = CheckNumber(Txt("txtMinStock").text, "minimum stock level")
    If Len(msg) > 0 Then ValidateForm = msg: Exit Function

    msg = CheckNumber(Txt("txtMaxStock").text, "maximum stock level")
    If Len(msg) > 0 Then ValidateForm = msg: Exit Function

    If CDbl(Txt("txtMaxStock").text) < CDbl(Txt("txtMinStock").text) Then
        ValidateForm = "The maximum stock level cannot be less than the minimum stock level."
        Exit Function
    End If

    msg = CheckOptionalDate(Txt("txtExpiryDate").text, "expiry date")
    If Len(msg) > 0 Then ValidateForm = msg: Exit Function

    If modProducts.BarcodeExists(Txt("txtBarcode").text, mSelectedID) Then
        ValidateForm = "Another product already uses this barcode."
        Exit Function
    End If
End Function

' Copies the form fields into a TProduct (call only after ValidateForm passed).
Private Sub ReadForm(ByRef p As modProducts.TProduct)
    p.productID = mSelectedID
    p.barcode = Trim$(Txt("txtBarcode").text)
    p.productName = Trim$(Txt("txtProductName").text)
    p.category = Cmb("cmbCategory").text
    p.brand = Trim$(Txt("txtBrand").text)
    p.supplierID = CStr(Cmb("cmbSupplier").List(Cmb("cmbSupplier").ListIndex, 0))
    p.PurchasePrice = CDbl(Txt("txtPurchasePrice").text)
    p.sellingPrice = CDbl(Txt("txtSellingPrice").text)
    p.stock = CDbl(Txt("txtQuantity").text)
    p.minStock = CDbl(Txt("txtMinStock").text)
    p.MaxStock = CDbl(Txt("txtMaxStock").text)
    p.Unit = Cmb("cmbUnit").text
    p.ExpiryDate = Trim$(Txt("txtExpiryDate").text)
    p.description = Trim$(Txt("txtDescription").text)
    p.status = Cmb("cmbStatus").text
End Sub

' Warns (but allows) selling below the purchase price.
Private Function ConfirmPrices() As Boolean
    ConfirmPrices = True
    If CDbl(Txt("txtSellingPrice").text) < CDbl(Txt("txtPurchasePrice").text) Then
        ConfirmPrices = (MsgBox("The selling price is lower than the purchase price." & vbCrLf & _
                         "You would lose money on every sale. Continue anyway?", _
                         vbYesNo + vbExclamation, "Check prices") = vbYes)
    End If
End Function

' ---------- Buttons ----------
Private Sub btnSave_Click()
    Dim p As modProducts.TProduct
    Dim msg As String, newID As String

    On Error GoTo ErrorHandler

    msg = ValidateForm()
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Please check your input"
        Exit Sub
    End If
    If Not ConfirmPrices() Then Exit Sub

    ReadForm p
    newID = modProducts.AddProduct(p)
    LogAction "ADD PRODUCT", newID, "Added " & p.productName

    RefreshList
    ClearForm
    SetStatus "Saved " & p.productName & " as " & newID & ". Ready for the next product."
    Exit Sub

ErrorHandler:
    MsgBox "The product could not be saved." & vbCrLf & Err.description, vbExclamation, "Save failed"
End Sub

Private Sub btnUpdate_Click()
    Dim p As modProducts.TProduct
    Dim msg As String

    On Error GoTo ErrorHandler

    If Len(mSelectedID) = 0 Then
        MsgBox "Please select a product from the list first.", vbInformation, "Update"
        Exit Sub
    End If

    msg = ValidateForm()
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Please check your input"
        Exit Sub
    End If
    If Not ConfirmPrices() Then Exit Sub

    ReadForm p
    If modProducts.UpdateProduct(p) Then
        LogAction "EDIT PRODUCT", p.productID, "Updated " & p.productName
        RefreshList
        ClearForm
        SetStatus "Updated " & p.productName & "."
    Else
        MsgBox "That product could not be found.", vbExclamation, "Update failed"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "The product could not be updated." & vbCrLf & Err.description, vbExclamation, "Update failed"
End Sub

Private Sub btnDelete_Click()
    Dim msg As String, deletedName As String, deletedID As String

    On Error GoTo ErrorHandler

    If Len(mSelectedID) = 0 Then
        MsgBox "Please select a product from the list first.", vbInformation, "Delete"
        Exit Sub
    End If

    deletedID = mSelectedID
    deletedName = Txt("txtProductName").text

    If MsgBox("Delete " & deletedName & " (" & deletedID & ")?" & vbCrLf & _
              "This cannot be undone.", vbYesNo + vbQuestion, "Confirm delete") <> vbYes Then Exit Sub

    msg = modProducts.DeleteProduct(deletedID)
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Cannot delete"
    Else
        LogAction "DELETE PRODUCT", deletedID, "Deleted " & deletedName
        RefreshList
        ClearForm
        SetStatus "Deleted " & deletedName & "."
    End If
    Exit Sub

ErrorHandler:
    MsgBox "The product could not be deleted." & vbCrLf & Err.description, vbExclamation, "Delete failed"
End Sub
Private Sub btnSearch_Click()
    Dim p As modProducts.TProduct

    frmSearch.Show                       ' waits here until the search screen closes

    RefreshList                          ' the search screen may have deleted something

    If Len(modMain.PendingProductID) > 0 Then
        ' The user clicked EDIT in the search screen: load that product
        If modProducts.LoadProduct(modMain.PendingProductID, p) Then FillForm p
        modMain.PendingProductID = ""
    ElseIf Len(mSelectedID) > 0 Then
        ' The product being edited may have been deleted in the search screen
        If modProducts.FindProductRow(mSelectedID) = 0 Then ClearForm
    End If
End Sub
Private Sub btnClear_Click()
    ClearForm
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


