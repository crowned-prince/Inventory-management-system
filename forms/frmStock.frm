VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStock 
   Caption         =   "Stock"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "frmStock.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmStock - inventory overview, stock history and adjustments
' ============================================================

Private mSelectedID As String
Private mLoading As Boolean

' ---------- Startup ----------
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    mLoading = True
    BuildLayout
    StyleControls
    LoadLookups
    mLoading = False

    RefreshAll
    SetStatus "Select a product to see its stock history, or correct its stock below."
    Exit Sub

ErrorHandler:
    mLoading = False
    MsgBox "The stock screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Stock"
End Sub

Private Function Txt(ByVal ctlName As String) As MSForms.TextBox
    Set Txt = Me.Controls(ctlName)
End Function

Private Function Cmb(ByVal ctlName As String) As MSForms.ComboBox
    Set Cmb = Me.Controls(ctlName)
End Function

Private Function cur() As String
    cur = GetSetting("CurrencySymbol", "Rs.") & " "
End Function

Private Sub SetStatus(ByVal message As String)
    Me.Controls("lblStatusMsg").caption = message
End Sub

' ---------- Build the parts made by code ----------
Private Sub BuildLayout()
    Dim cardTitles As Variant, cardColors As Variant
    Dim headings As Variant, widths As Variant
    Dim moveHeadings As Variant, moveWidths As Variant
    Dim i As Long
    Dim X As Single

    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 830, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "STOCK", -1, ThemeCyan(), 14, True

    ' Five summary cards
    cardTitles = Array("IN STOCK", "LOW STOCK", "OUT OF STOCK", "TOTAL UNITS", "STOCK VALUE (COST)")
    cardColors = Array(ThemeGreen(), ThemeOrange(), ThemeRed(), ThemeCyan(), ThemeCyan())
    For i = 0 To 4
        X = 20 + i * 156
        modUI.AddLabel Me, "lblCard" & (i + 1) & "Back", X, 48, 144, 46, "", ThemePanel()
        modUI.AddLabel Me, "lblCard" & (i + 1) & "Title", X + 8, 52, 128, 14, cardTitles(i), -1, ThemeGrey(), 8, True
        modUI.AddLabel Me, "lblCard" & (i + 1) & "Value", X + 8, 66, 128, 24, "0", -1, cardColors(i), 13, True
    Next i

    ' Filter row
    modUI.AddLabel Me, "lblFilter", 20, 102, 150, 14, "Show", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblSearch", 185, 102, 290, 14, "Search (name, ID, barcode, category)", -1, ThemeGrey(), 8

    ' Inventory list headings (widths match modInventory.FillInventoryList)
    headings = Array("ID", "PRODUCT", "CATEGORY", "STOCK", "MIN", "STATUS")
    widths = Array(62, 150, 80, 45, 40, 85)
    X = 23
    For i = 0 To UBound(headings)
        modUI.AddLabel Me, "lblHead" & i, X, 146, widths(i), 13, headings(i), -1, ThemeCyan(), 7, True
        X = X + widths(i)
    Next i

    ' Movement history (right side)
    modUI.AddLabel Me, "lblMoveTitle", 515, 128, 195, 16, "MOVEMENT HISTORY", -1, ThemeCyan(), 8, True
    With modUI.AddLabel(Me, "lblSelStatus", 710, 128, 80, 16, "", -1, ThemeGrey(), 8, True)
        .TextAlign = fmTextAlignRight
    End With

    moveHeadings = Array("DATE", "TYPE", "CHG", "AFTER", "REF")
    moveWidths = Array(78, 62, 38, 38, 50)
    X = 518
    For i = 0 To UBound(moveHeadings)
        modUI.AddLabel Me, "lblMoveHead" & i, X, 146, moveWidths(i), 13, moveHeadings(i), -1, ThemeCyan(), 7, True
        X = X + moveWidths(i)
    Next i

    ' Adjust section
    modUI.AddLabel Me, "lblAdjustTitle", 20, 340, 500, 16, _
                   "ADJUST STOCK  (stock count, damaged, expired or lost goods)", -1, ThemeCyan(), 8, True
    modUI.AddLabel Me, "lblCounted", 20, 358, 110, 14, "Counted quantity *", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblReason", 145, 358, 200, 14, "Reason *", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblNote", 360, 358, 250, 14, "Note (optional)", -1, ThemeGrey(), 8
    modUI.AddTextBox Me, "txtCounted", 20, 373, 110, 22
    modUI.AddCombo Me, "cmbReason", 145, 373, 200, 22
    modUI.AddTextBox Me, "txtNote", 360, 373, 250, 22
    Txt("txtNote").MaxLength = 100

    modUI.AddLabel Me, "lblStatusMsg", 20, 456, 770, 18, "", -1, ThemeGrey(), 9
End Sub

Private Sub StyleControls()
    With cmbStatusFilter
        .Left = 20: .Top = 117: .Width = 150: .Height = 22
        .style = fmStyleDropDownList
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
    End With

    With txtSearch
        .Left = 185: .Top = 117: .Width = 290: .Height = 22
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = ThemeCyan()
    End With

    StyleList lstInventory, 20, 160, 480, 172
    StyleList lstMovements, 515, 160, 275, 172

    StyleButton btnAdjust, "ADJUST STOCK", 625, 372, 165, 24, ThemeCyan(), ThemeNavy()
    StyleButton btnVerify, "VERIFY STOCK LOG", 20, 410, 170, 38, ThemeActive(), ThemeWhite()
    StyleButton btnLowStock, "LOW STOCK REPORT", 200, 410, 170, 38, ThemeOrange(), ThemeNavy()
    StyleButton btnClose, "CLOSE", 650, 410, 140, 38, ThemePanel(), ThemeWhite()
End Sub

Private Sub StyleList(ByVal lst As MSForms.ListBox, ByVal leftPos As Single, ByVal topPos As Single, _
                      ByVal widthVal As Single, ByVal heightVal As Single)
    With lst
        .Left = leftPos: .Top = topPos: .Width = widthVal: .Height = heightVal
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
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
    Dim item As Variant

    For Each item In Array("All products", "In stock", "Low stock", "Out of stock", "Inactive")
        cmbStatusFilter.AddItem CStr(item)
    Next item
    cmbStatusFilter.ListIndex = 0

    For Each item In Array("Damaged", "Expired", "Lost / stolen", "Counting correction", _
                           "Returned to supplier", "Other")
        Cmb("cmbReason").AddItem CStr(item)
    Next item
    Cmb("cmbReason").ListIndex = -1
End Sub

' ---------- Filter text to the status word used in the data ----------
Private Function stateFilter() As String
    Select Case cmbStatusFilter.ListIndex
        Case 1: stateFilter = "IN STOCK"
        Case 2: stateFilter = "LOW STOCK"
        Case 3: stateFilter = "OUT OF STOCK"
        Case 4: stateFilter = "INACTIVE"
        Case Else: stateFilter = ""
    End Select
End Function

' ---------- Refreshing the screen ----------
Private Sub RefreshAll()
    RefreshSummary
    RefreshList
    If Len(mSelectedID) > 0 Then ShowSelected mSelectedID
End Sub

Private Sub RefreshSummary()
    Dim inStock As Long, lowStock As Long, outOfStock As Long
    Dim totalUnits As Double, costValue As Double

    modInventory.InventorySummary inStock, lowStock, outOfStock, totalUnits, costValue

    Me.Controls("lblCard1Value").caption = Format$(inStock, "#,##0")
    Me.Controls("lblCard2Value").caption = Format$(lowStock, "#,##0")
    Me.Controls("lblCard3Value").caption = Format$(outOfStock, "#,##0")
    Me.Controls("lblCard4Value").caption = Format$(totalUnits, "#,##0")
    Me.Controls("lblCard5Value").caption = cur() & Format$(costValue, "#,##0")
End Sub

Private Sub RefreshList()
    Dim foundCount As Long

    If mLoading Then Exit Sub

    foundCount = modInventory.FillInventoryList(lstInventory, stateFilter(), txtSearch.text)
    If foundCount >= 1000 Then
        SetStatus "Showing the first 1000 products. Type in the search box to narrow the list."
    End If
End Sub

Private Sub cmbStatusFilter_Change()
    RefreshList
End Sub

Private Sub txtSearch_Change()
    RefreshList
End Sub

' ---------- Selecting a product ----------
Private Sub lstInventory_Click()
    If mLoading Then Exit Sub
    If lstInventory.ListIndex < 0 Then Exit Sub
    ShowSelected CStr(lstInventory.List(lstInventory.ListIndex, 0))
End Sub

Private Sub ShowSelected(ByVal productID As String)
    Dim p As modProducts.TProduct
    Dim state As String
    Dim badge As MSForms.label

    If Not modProducts.LoadProduct(productID, p) Then
        ClearSelected
        Exit Sub
    End If

    mSelectedID = p.productID
    state = modInventory.ProductStockState(p.stock, p.minStock, p.status)

    Me.Controls("lblMoveTitle").caption = "HISTORY: " & Left$(p.productName, 22)

    Set badge = Me.Controls("lblSelStatus")
    badge.caption = state
    Select Case state
        Case "IN STOCK": badge.ForeColor = ThemeGreen()
        Case "LOW STOCK": badge.ForeColor = ThemeOrange()
        Case "OUT OF STOCK": badge.ForeColor = ThemeRed()
        Case Else: badge.ForeColor = ThemeGrey()
    End Select

    modInventory.FillMovementHistory lstMovements, p.productID, 40
    Txt("txtCounted").text = CStr(p.stock)
End Sub

Private Sub ClearSelected()
    mSelectedID = ""
    Me.Controls("lblMoveTitle").caption = "MOVEMENT HISTORY"
    Me.Controls("lblSelStatus").caption = ""
    lstMovements.Clear
    Txt("txtCounted").text = ""
End Sub

' ---------- ADJUST STOCK ----------
Private Sub btnAdjust_Click()
    Dim p As modProducts.TProduct
    Dim msg As String
    Dim counted As Double, difference As Double
    Dim reason As String, note As String

    On Error GoTo ErrorHandler

    If modMain.CurrentRole = "STAFF" Then
        MsgBox "Your role cannot adjust stock.", vbExclamation, "Adjust stock"
        Exit Sub
    End If

    If Len(mSelectedID) = 0 Then
        MsgBox "Please select a product from the list first.", vbInformation, "Adjust stock"
        Exit Sub
    End If

    msg = CheckNumber(Txt("txtCounted").text, "counted quantity")
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Please check your input"
        Exit Sub
    End If

    If Cmb("cmbReason").ListIndex < 0 Then
        MsgBox "Please choose a reason for the adjustment.", vbExclamation, "Adjust stock"
        Exit Sub
    End If

    If Not modProducts.LoadProduct(mSelectedID, p) Then
        MsgBox "That product could not be found.", vbExclamation, "Adjust stock"
        Exit Sub
    End If

    counted = CDbl(Txt("txtCounted").text)
    difference = Round(counted - p.stock, 4)
    If difference = 0 Then
        MsgBox "The counted quantity is the same as the current stock. Nothing to change.", _
               vbInformation, "Adjust stock"
        Exit Sub
    End If

    reason = Cmb("cmbReason").text
    note = Trim$(Txt("txtNote").text)

    If MsgBox("Change the stock of " & p.productName & "?" & vbCrLf & vbCrLf & _
              "Current stock:  " & p.stock & vbCrLf & _
              "Counted:  " & counted & vbCrLf & _
              "Difference:  " & IIf(difference > 0, "+", "") & difference & vbCrLf & _
              "Reason:  " & reason, _
              vbYesNo + vbQuestion, "Confirm adjustment") <> vbYes Then Exit Sub

    modInventory.AdjustStockToCount mSelectedID, counted, reason, note
    LogAction "ADJUST STOCK", mSelectedID, p.productName & ": " & p.stock & " to " & counted & " (" & reason & ")"

    Txt("txtNote").text = ""
    Cmb("cmbReason").ListIndex = -1
    RefreshAll
    SetStatus "Stock of " & p.productName & " is now " & counted & " (" & reason & ")."
    Exit Sub

ErrorHandler:
    MsgBox "The stock could not be adjusted." & vbCrLf & Err.description, vbExclamation, "Adjust stock"
End Sub

' ---------- VERIFY STOCK LOG ----------
Private Sub btnVerify_Click()
    Dim report As String
    Dim issues As Long, fixedCount As Long
    Dim msg As String

    On Error GoTo ErrorHandler

    issues = modInventory.CheckStockLog(False, report)

    If issues = 0 Then
        MsgBox "All stock figures match the movement log.", vbInformation, "Stock check"
        Exit Sub
    End If

    msg = issues & " product(s) do not match the movement log:" & vbCrLf & vbCrLf & report

    If modMain.IsAdmin() Then
        If MsgBox(msg & vbCrLf & _
                  "Add correction entries so the log matches the current stock?" & vbCrLf & _
                  "(The stock numbers themselves will not change.)", _
                  vbYesNo + vbQuestion, "Stock check") = vbYes Then

            fixedCount = modInventory.CheckStockLog(True, report)
            LogAction "STOCK LOG FIX", "-", fixedCount & " correction entries added"
            RefreshAll
            SetStatus fixedCount & " correction entries added to the movement log."
            MsgBox fixedCount & " correction entries were added.", vbInformation, "Stock check"
        End If
    Else
        MsgBox msg & vbCrLf & "Please ask an administrator to correct the movement log.", _
               vbExclamation, "Stock check"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "The stock check could not finish." & vbCrLf & Err.description, vbExclamation, "Stock check"
End Sub

Private Sub btnLowStock_Click()
    frmLowStock.Show
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


