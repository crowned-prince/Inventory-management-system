VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSuppliers 
   Caption         =   "Supplier Management"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "frmSuppliers.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmSuppliers - add, update and delete suppliers
' ============================================================

Private mSelectedID As String      ' "" = new supplier, otherwise the one being edited
Private mLoading As Boolean

' ---------- Startup ----------
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    BuildLayout
    StyleControls
    RefreshList
    ClearForm
    Exit Sub

ErrorHandler:
    MsgBox "The supplier screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Suppliers"
End Sub

Private Function Txt(ByVal ctlName As String) As MSForms.TextBox
    Set Txt = Me.Controls(ctlName)
End Function

' ---------- Build the screen ----------
Private Sub BuildLayout()
    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 830, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "SUPPLIER MANAGEMENT", -1, ThemeCyan(), 14, True

    ' Left column
    AddField "Supplier ID", "txtSupplierID", 1, 0
    AddField "Supplier Name *", "txtSupplierName", 1, 1
    AddField "Contact Person", "txtContactPerson", 1, 2
    AddField "Phone", "txtPhone", 1, 3
    AddField "Email", "txtEmail", 1, 4

    ' Right column
    AddField "Address", "txtAddress", 2, 0
    AddField "Products Supplied (automatic)", "txtProductsSupplied", 2, 2
    AddField "Total Purchases (automatic)", "txtTotalPurchases", 2, 3

    With Txt("txtAddress")
        .Height = 66
        .Multiline = True
        .WordWrap = True
        .EnterKeyBehavior = True
    End With

    ' Fields the user must not type in
    LockField "txtSupplierID"
    LockField "txtProductsSupplied"
    LockField "txtTotalPurchases"

    Txt("txtSupplierName").MaxLength = 100

    modUI.AddLabel Me, "lblFilterTitle", 520, 52, 270, 14, "SUPPLIER LIST  (type to filter)", -1, ThemeCyan(), 8, True
    modUI.AddLabel Me, "lblStatusMsg", 20, 372, 770, 18, "", -1, ThemeGrey(), 9
End Sub

Private Sub AddField(ByVal captionText As String, ByVal ctlName As String, _
                     ByVal columnNo As Long, ByVal rowNo As Long)
    Dim X As Single, Y As Single

    If columnNo = 1 Then X = 20 Else X = 270
    Y = 56 + (rowNo * 44)

    modUI.AddLabel Me, "lbl" & Mid$(ctlName, 4), X, Y, 230, 14, captionText, -1, ThemeGrey(), 8
    modUI.AddTextBox Me, ctlName, X, Y + 15, 230, 22
End Sub

Private Sub LockField(ByVal ctlName As String)
    With Txt(ctlName)
        .locked = True
        .BackColor = ThemeSidebar()
        .ForeColor = ThemeCyan()
    End With
End Sub

Private Sub StyleControls()
    With txtFilter
        .Left = 520: .Top = 68: .Width = 270: .Height = 22
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = ThemeCyan()
    End With

    With lstSuppliers
        .Left = 520: .Top = 96: .Width = 270: .Height = 226
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    StyleButton btnSave, "SAVE", 20, ThemeGreen(), ThemeNavy()
    StyleButton btnUpdate, "UPDATE", 172, ThemeCyan(), ThemeNavy()
    StyleButton btnDelete, "DELETE", 324, ThemeRed(), ThemeNavy()
    StyleButton btnClear, "NEW / CLEAR", 476, ThemePanel(), ThemeWhite()
    StyleButton btnClose, "CLOSE", 628, ThemePanel(), ThemeWhite()
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal bgColor As Long, ByVal fgColor As Long)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = 326: .Width = 140: .Height = 38
        .BackColor = bgColor: .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = 10: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub RefreshList()
    mLoading = True
    modSuppliers.FillSupplierList lstSuppliers, txtFilter.text
    mLoading = False
End Sub

Private Sub txtFilter_Change()
    If mLoading Then Exit Sub
    RefreshList
End Sub

Private Sub SetStatus(ByVal message As String)
    Me.Controls("lblStatusMsg").caption = message
End Sub

' ---------- New-supplier mode ----------
Private Sub ClearForm()
    mSelectedID = ""
    mLoading = True

    Txt("txtSupplierID").text = "AUTO"
    Txt("txtSupplierName").text = ""
    Txt("txtContactPerson").text = ""
    Txt("txtPhone").text = ""
    Txt("txtEmail").text = ""
    Txt("txtAddress").text = ""
    Txt("txtProductsSupplied").text = "-"
    Txt("txtTotalPurchases").text = "-"
    lstSuppliers.ListIndex = -1

    mLoading = False
    UpdateButtonState
    SetStatus "New supplier. Fill in the fields marked * and click SAVE."
    Txt("txtSupplierName").SetFocus
End Sub

Private Sub UpdateButtonState()
    Dim canEdit As Boolean, isNew As Boolean

    canEdit = (modMain.CurrentRole <> "STAFF")
    isNew = (Len(mSelectedID) = 0)

    btnSave.Enabled = canEdit And isNew
    btnUpdate.Enabled = canEdit And Not isNew
    btnDelete.Enabled = modMain.IsAdmin() And Not isNew

    If Not canEdit Then SetStatus "View only: your role cannot change suppliers."
End Sub

' ---------- Selecting a supplier from the list ----------
Private Sub lstSuppliers_Click()
    Dim p As modSuppliers.TSupplier

    If mLoading Then Exit Sub
    If lstSuppliers.ListIndex < 0 Then Exit Sub

    If modSuppliers.LoadSupplier(CStr(lstSuppliers.List(lstSuppliers.ListIndex, 0)), p) Then
        FillForm p
    End If
End Sub

Private Sub FillForm(ByRef p As modSuppliers.TSupplier)
    Dim cur As String
    Dim productCount As Long

    cur = GetSetting("CurrencySymbol", "Rs.") & " "
    productCount = modSuppliers.CountProductsBySupplier(p.supplierID)

    mLoading = True
    mSelectedID = p.supplierID

    Txt("txtSupplierID").text = p.supplierID
    Txt("txtSupplierName").text = p.supplierName
    Txt("txtContactPerson").text = p.ContactPerson
    Txt("txtPhone").text = p.phone
    Txt("txtEmail").text = p.Email
    Txt("txtAddress").text = p.Address
    Txt("txtProductsSupplied").text = productCount & " product(s)"
    Txt("txtTotalPurchases").text = cur & Format$(p.TotalPurchases, "#,##0.00")

    mLoading = False
    UpdateButtonState
    SetStatus "Editing " & p.supplierID & "."
End Sub

' ---------- Validation: "" = OK, otherwise the message ----------
Private Function ValidateForm() As String
    Dim msg As String

    msg = CheckRequired(Txt("txtSupplierName").text, "supplier name")
    If Len(msg) > 0 Then ValidateForm = msg: Exit Function

    If modSuppliers.SupplierNameExists(Txt("txtSupplierName").text, mSelectedID) Then
        ValidateForm = "A supplier with this name already exists."
        Exit Function
    End If

    msg = CheckPhone(Txt("txtPhone").text)
    If Len(msg) > 0 Then ValidateForm = msg: Exit Function

    msg = CheckEmail(Txt("txtEmail").text)
    If Len(msg) > 0 Then ValidateForm = msg: Exit Function
End Function

Private Sub ReadForm(ByRef p As modSuppliers.TSupplier)
    p.supplierID = mSelectedID
    p.supplierName = Trim$(Txt("txtSupplierName").text)
    p.ContactPerson = Trim$(Txt("txtContactPerson").text)
    p.phone = Trim$(Txt("txtPhone").text)
    p.Email = Trim$(Txt("txtEmail").text)
    p.Address = Trim$(Txt("txtAddress").text)
End Sub

' ---------- Buttons ----------
Private Sub btnSave_Click()
    Dim p As modSuppliers.TSupplier
    Dim msg As String, newID As String

    On Error GoTo ErrorHandler

    msg = ValidateForm()
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Please check your input"
        Exit Sub
    End If

    ReadForm p
    newID = modSuppliers.AddSupplier(p)
    LogAction "ADD SUPPLIER", newID, "Added " & p.supplierName

    RefreshList
    ClearForm
    SetStatus "Saved " & p.supplierName & " as " & newID & ". Ready for the next supplier."
    Exit Sub

ErrorHandler:
    MsgBox "The supplier could not be saved." & vbCrLf & Err.description, vbExclamation, "Save failed"
End Sub

Private Sub btnUpdate_Click()
    Dim p As modSuppliers.TSupplier
    Dim msg As String

    On Error GoTo ErrorHandler

    If Len(mSelectedID) = 0 Then
        MsgBox "Please select a supplier from the list first.", vbInformation, "Update"
        Exit Sub
    End If

    msg = ValidateForm()
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Please check your input"
        Exit Sub
    End If

    ReadForm p
    If modSuppliers.UpdateSupplier(p) Then
        LogAction "EDIT SUPPLIER", p.supplierID, "Updated " & p.supplierName
        RefreshList
        ClearForm
        SetStatus "Updated " & p.supplierName & "."
    Else
        MsgBox "That supplier could not be found.", vbExclamation, "Update failed"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "The supplier could not be updated." & vbCrLf & Err.description, vbExclamation, "Update failed"
End Sub

Private Sub btnDelete_Click()
    Dim msg As String, deletedName As String, deletedID As String

    On Error GoTo ErrorHandler

    If Len(mSelectedID) = 0 Then
        MsgBox "Please select a supplier from the list first.", vbInformation, "Delete"
        Exit Sub
    End If

    deletedID = mSelectedID
    deletedName = Txt("txtSupplierName").text

    If MsgBox("Delete " & deletedName & " (" & deletedID & ")?" & vbCrLf & _
              "This cannot be undone.", vbYesNo + vbQuestion, "Confirm delete") <> vbYes Then Exit Sub

    msg = modSuppliers.DeleteSupplier(deletedID)
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Cannot delete"
    Else
        LogAction "DELETE SUPPLIER", deletedID, "Deleted " & deletedName
        RefreshList
        ClearForm
        SetStatus "Deleted " & deletedName & "."
    End If
    Exit Sub

ErrorHandler:
    MsgBox "The supplier could not be deleted." & vbCrLf & Err.description, vbExclamation, "Delete failed"
End Sub

Private Sub btnClear_Click()
    ClearForm
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


