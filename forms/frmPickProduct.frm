VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPickProduct 
   Caption         =   "Choose a Product"
   ClientHeight    =   8820.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10965
   OleObjectBlob   =   "frmPickProduct.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPickProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmPickProduct - choose the product the Dashboard sheet should focus on
' ============================================================

Private mLoading As Boolean

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    mLoading = True
    BuildLayout
    StyleControls
    mLoading = False

    RefreshList
    Exit Sub

ErrorHandler:
    mLoading = False
    MsgBox "The product list could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Choose a product"
End Sub

Private Sub BuildLayout()
    Dim headings As Variant, widths As Variant
    Dim i As Long
    Dim X As Single

    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 580, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "CHOOSE A PRODUCT", -1, ThemeCyan(), 14, True
    modUI.AddLabel Me, "lblHint", 20, 46, 520, 14, _
                   "Type to search by name, barcode, ID, category, brand or supplier", -1, ThemeGrey(), 8

    headings = Array("ID", "PRODUCT", "CATEGORY", "BRAND", "PRICE", "STOCK")
    widths = Array(62, 150, 80, 70, 55, 40)
    X = 23
    For i = 0 To UBound(headings)
        modUI.AddLabel Me, "lblHead" & i, X, 94, widths(i), 13, headings(i), -1, ThemeCyan(), 7, True
        X = X + widths(i)
    Next i

    modUI.AddLabel Me, "lblStatusMsg", 20, 408, 520, 18, "", -1, ThemeGrey(), 9
End Sub

Private Sub StyleControls()
    With txtSearch
        .Left = 20: .Top = 64: .Width = 520: .Height = 22
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = ThemeCyan()
    End With

    With lstProducts
        .Left = 20: .Top = 108: .Width = 520: .Height = 250
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    StyleButton btnAll, "ALL PRODUCTS", 20, 170, ThemePanel(), ThemeWhite()
    StyleButton btnSelect, "SHOW THIS PRODUCT", 200, 170, ThemeCyan(), ThemeNavy()
    StyleButton btnClose, "CLOSE", 380, 160, ThemePanel(), ThemeWhite()
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal widthVal As Single, _
                        ByVal bgColor As Long, ByVal fgColor As Long)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = 366: .Width = widthVal: .Height = 34
        .BackColor = bgColor: .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = 10: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub RefreshList()
    Dim n As Long

    If mLoading Then Exit Sub
    n = modProducts.SearchProducts(lstProducts, "All fields", txtSearch.text)
    Me.Controls("lblStatusMsg").caption = n & " product(s) found. Double-click one, or select it and click SHOW THIS PRODUCT."
End Sub

Private Sub txtSearch_Change()
    RefreshList
End Sub

Private Sub ChooseSelected()
    If lstProducts.ListIndex < 0 Then
        MsgBox "Please select a product first.", vbInformation, "Choose a product"
        Exit Sub
    End If

    modDash.SetProductFilter CStr(lstProducts.List(lstProducts.ListIndex, 0)), _
                             CStr(lstProducts.List(lstProducts.ListIndex, 1))
    Unload Me
End Sub

Private Sub btnSelect_Click()
    ChooseSelected
End Sub

Private Sub lstProducts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ChooseSelected
End Sub

Private Sub btnAll_Click()
    modDash.SetProductFilter "", ""
    Unload Me
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

