VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearch 
   Caption         =   "	Product Search"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "frmSearch.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmSearch - find products by ID, name, barcode, category,
' brand or supplier, then edit / view / delete
' ============================================================

Private mSelectedID As String
Private mLoading As Boolean

' ---------- Startup ----------
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    mLoading = True
    BuildLayout
    StyleControls
    LoadSearchByCombo
    mLoading = False

    RunSearch                       ' show all products to start with
    UpdateButtonState
    txtSearch.SetFocus
    Exit Sub

ErrorHandler:
    mLoading = False
    MsgBox "The search screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Search"
End Sub

' ---------- Build the parts made by code ----------
Private Sub BuildLayout()
    Dim headings As Variant, widths As Variant
    Dim detailCaptions As Variant, detailNames As Variant
    Dim i As Long
    Dim X As Single, Y As Single

    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 830, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "PRODUCT SEARCH", -1, ThemeCyan(), 14, True

    modUI.AddLabel Me, "lblSearchBy", 20, 48, 150, 14, "Search by", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblSearchText", 185, 48, 255, 14, "Type to search (name, barcode, ID ...)", -1, ThemeGrey(), 8

    ' Column headings above the results list (widths match modProducts.SearchProducts)
    headings = Array("ID", "PRODUCT", "CATEGORY", "BRAND", "PRICE", "STOCK")
    widths = Array(62, 150, 80, 70, 55, 40)
    X = 23
    For i = 0 To UBound(headings)
        modUI.AddLabel Me, "lblHead" & i, X, 92, widths(i), 13, headings(i), -1, ThemeCyan(), 7, True
        X = X + widths(i)
    Next i

    ' Details panel on the right
    modUI.AddLabel Me, "lblDetailPanel", 510, 92, 280, 312, "", ThemePanel()
    modUI.AddLabel Me, "lblDetailTitle", 522, 98, 250, 16, "PRODUCT DETAILS", -1, ThemeCyan(), 9, True

    detailCaptions = Array("Product Name", "Selling Price", "Current Stock", "Stock Status", _
                           "Supplier", "Category", "Status")
    detailNames = Array("lblValName", "lblValPrice", "lblValStock", "lblValStockStatus", _
                        "lblValSupplier", "lblValCategory", "lblValStatus")

    For i = 0 To UBound(detailCaptions)
        Y = 122 + i * 40
        modUI.AddLabel Me, "lblCap" & i, 522, Y, 250, 14, detailCaptions(i), -1, ThemeGrey(), 8
        modUI.AddLabel Me, CStr(detailNames(i)), 522, Y + 14, 250, 20, "-", -1, ThemeWhite(), 11, True
    Next i

    modUI.AddLabel Me, "lblStatusMsg", 20, 468, 770, 18, "", -1, ThemeGrey(), 9
End Sub

Private Sub StyleControls()
    With cmbSearchBy
        .Left = 20: .Top = 62: .Width = 150: .Height = 24
        .style = fmStyleDropDownList
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
    End With

    With txtSearch
        .Left = 185: .Top = 62: .Width = 255: .Height = 24
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = ThemeCyan()
    End With

    StyleButton btnSearch, "SEARCH", 450, 62, 100, 24, ThemeCyan(), ThemeNavy()
    StyleButton btnClearSearch, "CLEAR", 560, 62, 100, 24, ThemePanel(), ThemeWhite()

    With lstResults
        .Left = 20: .Top = 106: .Width = 470: .Height = 298
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    StyleButton btnEdit, "EDIT", 20, 416, 140, 40, ThemeCyan(), ThemeNavy()
    StyleButton btnViewDetails, "VIEW DETAILS", 172, 416, 140, 40, ThemeActive(), ThemeWhite()
    StyleButton btnDelete, "DELETE", 324, 416, 140, 40, ThemeRed(), ThemeNavy()
    StyleButton btnClose, "CLOSE", 650, 416, 140, 40, ThemePanel(), ThemeWhite()
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

Private Sub LoadSearchByCombo()
    Dim item As Variant
    For Each item In Array("All fields", "Product ID", "Product Name", "Barcode", _
                           "Category", "Brand", "Supplier")
        cmbSearchBy.AddItem CStr(item)
    Next item
    cmbSearchBy.ListIndex = 0
End Sub

Private Sub SetStatus(ByVal message As String)
    Me.Controls("lblStatusMsg").caption = message
End Sub

' ---------- Searching ----------
Private Sub RunSearch()
    Dim foundCount As Long

    If mLoading Then Exit Sub

    foundCount = modProducts.SearchProducts(lstResults, cmbSearchBy.text, txtSearch.text)
    ClearDetails

    If foundCount = 0 Then
        SetStatus "No products found."
    ElseIf foundCount >= 500 Then
        SetStatus "Showing the first 500 matches. Type more to narrow the search."
    Else
        SetStatus foundCount & " product(s) found."
    End If

    ' Exactly one result (for example a scanned barcode): select it automatically
    If foundCount = 1 Then
        mLoading = True
        lstResults.ListIndex = 0
        mLoading = False
        ShowDetails CStr(lstResults.List(0, 0))
    End If
End Sub

Private Sub txtSearch_Change()
    RunSearch                       ' live search while typing
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then RunSearch
End Sub

Private Sub cmbSearchBy_Change()
    RunSearch
End Sub

Private Sub btnSearch_Click()
    RunSearch
End Sub

Private Sub btnClearSearch_Click()
    mLoading = True
    txtSearch.text = ""
    cmbSearchBy.ListIndex = 0
    mLoading = False
    RunSearch
    txtSearch.SetFocus
End Sub

' ---------- Showing the selected product ----------
Private Sub lstResults_Click()
    If mLoading Then Exit Sub
    If lstResults.ListIndex < 0 Then Exit Sub
    ShowDetails CStr(lstResults.List(lstResults.ListIndex, 0))
End Sub

Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If btnEdit.Enabled Then btnEdit_Click
End Sub

Private Sub ShowDetails(ByVal productID As String)
    Dim p As modProducts.TProduct
    Dim status As String

    If Not modProducts.LoadProduct(productID, p) Then
        ClearDetails
        Exit Sub
    End If

    mSelectedID = p.productID
    status = modInventory.StockStatus(p.stock, p.minStock)

    Me.Controls("lblValName").caption = p.productName
    Me.Controls("lblValPrice").caption = GetSetting("CurrencySymbol", "Rs.") & " " & _
                                         Format$(p.sellingPrice, "#,##0.00")
    Me.Controls("lblValStock").caption = CStr(p.stock) & " " & p.Unit
    Me.Controls("lblValStockStatus").caption = status
    Me.Controls("lblValSupplier").caption = modProducts.GetSupplierName(p.supplierID)
    Me.Controls("lblValCategory").caption = p.category
    Me.Controls("lblValStatus").caption = p.status

    Select Case status
        Case "IN STOCK":  Me.Controls("lblValStockStatus").ForeColor = ThemeGreen()
        Case "LOW STOCK": Me.Controls("lblValStockStatus").ForeColor = ThemeOrange()
        Case Else:        Me.Controls("lblValStockStatus").ForeColor = ThemeRed()
    End Select

    UpdateButtonState
End Sub

Private Sub ClearDetails()
    Dim names As Variant
    Dim item As Variant

    mSelectedID = ""
    names = Array("lblValName", "lblValPrice", "lblValStock", "lblValStockStatus", _
                  "lblValSupplier", "lblValCategory", "lblValStatus")
    For Each item In names
        Me.Controls(CStr(item)).caption = "-"
        Me.Controls(CStr(item)).ForeColor = ThemeWhite()
    Next item

    UpdateButtonState
End Sub

Private Sub UpdateButtonState()
    Dim hasSelection As Boolean
    hasSelection = (Len(mSelectedID) > 0)

    btnEdit.Enabled = hasSelection And (modMain.CurrentRole <> "STAFF")
    btnViewDetails.Enabled = hasSelection
    btnDelete.Enabled = hasSelection And modMain.IsAdmin()
End Sub

' ---------- Action buttons ----------
Private Sub btnEdit_Click()
    If Len(mSelectedID) = 0 Then
        MsgBox "Please select a product first.", vbInformation, "Edit"
        Exit Sub
    End If

    ' Tell the Product form which product to open, then close this screen
    modMain.PendingProductID = mSelectedID
    Unload Me
End Sub

Private Sub btnViewDetails_Click()
    Dim p As modProducts.TProduct
    Dim cur As String, text As String

    If Len(mSelectedID) = 0 Then
        MsgBox "Please select a product first.", vbInformation, "View details"
        Exit Sub
    End If
    If Not modProducts.LoadProduct(mSelectedID, p) Then Exit Sub

    cur = GetSetting("CurrencySymbol", "Rs.") & " "

    text = "Product ID:  " & p.productID & vbCrLf & _
           "Name:  " & p.productName & vbCrLf & _
           "Barcode:  " & p.barcode & vbCrLf & _
           "Category:  " & p.category & vbCrLf & _
           "Brand:  " & p.brand & vbCrLf & _
           "Supplier:  " & modProducts.GetSupplierName(p.supplierID) & vbCrLf & vbCrLf

    ' Cost price is hidden from STAFF
    If modMain.CurrentRole <> "STAFF" Then
        text = text & "Purchase price:  " & cur & Format$(p.PurchasePrice, "#,##0.00") & vbCrLf
    End If

    text = text & "Selling price:  " & cur & Format$(p.sellingPrice, "#,##0.00") & vbCrLf & _
           "Stock:  " & p.stock & " " & p.Unit & "  (" & modInventory.StockStatus(p.stock, p.minStock) & ")" & vbCrLf & _
           "Minimum / Maximum:  " & p.minStock & " / " & p.MaxStock & vbCrLf & _
           "Expiry date:  " & IIf(Len(p.ExpiryDate) > 0, p.ExpiryDate, "-") & vbCrLf & _
           "Status:  " & p.status & vbCrLf & vbCrLf & _
           "Description:  " & p.description

    MsgBox text, vbInformation, "Product details"
End Sub

Private Sub btnDelete_Click()
    Dim msg As String, deletedName As String, deletedID As String

    On Error GoTo ErrorHandler

    If Len(mSelectedID) = 0 Then
        MsgBox "Please select a product first.", vbInformation, "Delete"
        Exit Sub
    End If

    deletedID = mSelectedID
    deletedName = Me.Controls("lblValName").caption

    If MsgBox("Delete " & deletedName & " (" & deletedID & ")?" & vbCrLf & _
              "This cannot be undone.", vbYesNo + vbQuestion, "Confirm delete") <> vbYes Then Exit Sub

    msg = modProducts.DeleteProduct(deletedID)
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Cannot delete"
    Else
        LogAction "DELETE PRODUCT", deletedID, "Deleted " & deletedName
        RunSearch
        SetStatus "Deleted " & deletedName & "."
    End If
    Exit Sub

ErrorHandler:
    MsgBox "The product could not be deleted." & vbCrLf & Err.description, vbExclamation, "Delete failed"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


