VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCustomers 
   Caption         =   "	Customer Management"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "frmCustomers.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmCustomers - add, update and delete customers
' ============================================================

Private mSelectedID As String      ' "" = new customer, otherwise the one being edited
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
    MsgBox "The customer screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Customers"
End Sub

Private Function Txt(ByVal ctlName As String) As MSForms.TextBox
    Set Txt = Me.Controls(ctlName)
End Function

' ---------- Build the screen ----------
Private Sub BuildLayout()
    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 830, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "CUSTOMER MANAGEMENT", -1, ThemeCyan(), 14, True

    ' Left column
    AddField "Customer ID", "txtCustomerID", 1, 0
    AddField "Customer Name *", "txtCustomerName", 1, 1
    AddField "Phone", "txtPhone", 1, 2
    AddField "Email", "txtEmail", 1, 3
    AddField "Registration Date (automatic)", "txtRegistrationDate", 1, 4

    ' Right column
    AddField "Address", "txtAddress", 2, 0
    AddField "Previous Purchases (automatic)", "txtPreviousPurchases", 2, 2
    AddField "Total Spent (automatic)", "txtTotalSpent", 2, 3

    With Txt("txtAddress")
        .Height = 66
        .Multiline = True
        .WordWrap = True
        .EnterKeyBehavior = True
    End With

    LockField "txtCustomerID"
    LockField "txtRegistrationDate"
    LockField "txtPreviousPurchases"
    LockField "txtTotalSpent"

    Txt("txtCustomerName").MaxLength = 100

    modUI.AddLabel Me, "lblFilterTitle", 520, 52, 270, 14, "CUSTOMER LIST  (type to search)", -1, ThemeCyan(), 8, True
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

    With lstCustomers
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
    modCustomers.FillCustomerList lstCustomers, txtFilter.text
    mLoading = False
End Sub

Private Sub txtFilter_Change()
    If mLoading Then Exit Sub
    RefreshList
End Sub

Private Sub SetStatus(ByVal message As String)
    Me.Controls("lblStatusMsg").caption = message
End Sub

' ---------- New-customer mode ----------
Private Sub ClearForm()
    mSelectedID = ""
    mLoading = True

    Txt("txtCustomerID").text = "AUTO"
    Txt("txtCustomerName").text = ""
    Txt("txtPhone").text = ""
    Txt("txtEmail").text = ""
    Txt("txtAddress").text = ""
    Txt("txtRegistrationDate").text = Format$(Date, "dd-mm-yyyy")
    Txt("txtPreviousPurchases").text = "-"
    Txt("txtTotalSpent").text = "-"
    lstCustomers.ListIndex = -1

    mLoading = False
    UpdateButtonState
    SetStatus "New customer. Fill in the fields marked * and click SAVE."
    Txt("txtCustomerName").SetFocus
End Sub

' All roles can add and edit customers. Only ADMIN can delete.
' The built-in Walk-in Customer can never be changed.
Private Sub UpdateButtonState()
    Dim isNew As Boolean, isSystem As Boolean

    isNew = (Len(mSelectedID) = 0)
    isSystem = modCustomers.IsSystemCustomer(mSelectedID)

    btnSave.Enabled = isNew
    btnUpdate.Enabled = (Not isNew) And (Not isSystem)
    btnDelete.Enabled = modMain.IsAdmin() And (Not isNew) And (Not isSystem)
End Sub

' ---------- Selecting a customer from the list ----------
Private Sub lstCustomers_Click()
    Dim p As modCustomers.TCustomer

    If mLoading Then Exit Sub
    If lstCustomers.ListIndex < 0 Then Exit Sub

    If modCustomers.LoadCustomer(CStr(lstCustomers.List(lstCustomers.ListIndex, 0)), p) Then
        FillForm p
    End If
End Sub

Private Sub FillForm(ByRef p As modCustomers.TCustomer)
    Dim cur As String

    cur = GetSetting("CurrencySymbol", "Rs.") & " "

    mLoading = True
    mSelectedID = p.customerID

    Txt("txtCustomerID").text = p.customerID
    Txt("txtCustomerName").text = p.customerName
    Txt("txtPhone").text = p.phone
    Txt("txtEmail").text = p.Email
    Txt("txtAddress").text = p.Address
    Txt("txtRegistrationDate").text = p.RegistrationDate
    Txt("txtPreviousPurchases").text = modCustomers.CountPurchases(p.customerID) & " invoice(s)"
    Txt("txtTotalSpent").text = cur & Format$(modCustomers.TotalSpentBy(p.customerID), "#,##0.00")

    mLoading = False
    UpdateButtonState

    If modCustomers.IsSystemCustomer(p.customerID) Then
        SetStatus "Built-in customer used for sales without a name. It cannot be changed."
    Else
        SetStatus "Editing " & p.customerID & "."
    End If
End Sub

' ---------- Validation: "" = OK, otherwise the message ----------
Private Function ValidateForm() As String
    Dim msg As String

    msg = CheckRequired(Txt("txtCustomerName").text, "customer name")
    If Len(msg) > 0 Then ValidateForm = msg: Exit Function

    msg = CheckPhone(Txt("txtPhone").text)
    If Len(msg) > 0 Then ValidateForm = msg: Exit Function

    If modCustomers.PhoneExists(Txt("txtPhone").text, mSelectedID) Then
        ValidateForm = "Another customer already uses this phone number."
        Exit Function
    End If

    msg = CheckEmail(Txt("txtEmail").text)
    If Len(msg) > 0 Then ValidateForm = msg: Exit Function
End Function

Private Sub ReadForm(ByRef p As modCustomers.TCustomer)
    p.customerID = mSelectedID
    p.customerName = Trim$(Txt("txtCustomerName").text)
    p.phone = Trim$(Txt("txtPhone").text)
    p.Email = Trim$(Txt("txtEmail").text)
    p.Address = Trim$(Txt("txtAddress").text)
End Sub

' ---------- Buttons ----------
Private Sub btnSave_Click()
    Dim p As modCustomers.TCustomer
    Dim msg As String, newID As String

    On Error GoTo ErrorHandler

    msg = ValidateForm()
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Please check your input"
        Exit Sub
    End If

    ReadForm p
    newID = modCustomers.AddCustomer(p)
    LogAction "ADD CUSTOMER", newID, "Added " & p.customerName

    RefreshList
    ClearForm
    SetStatus "Saved " & p.customerName & " as " & newID & ". Ready for the next customer."
    Exit Sub

ErrorHandler:
    MsgBox "The customer could not be saved." & vbCrLf & Err.description, vbExclamation, "Save failed"
End Sub

Private Sub btnUpdate_Click()
    Dim p As modCustomers.TCustomer
    Dim msg As String

    On Error GoTo ErrorHandler

    If Len(mSelectedID) = 0 Then
        MsgBox "Please select a customer from the list first.", vbInformation, "Update"
        Exit Sub
    End If

    msg = ValidateForm()
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Please check your input"
        Exit Sub
    End If

    ReadForm p
    If modCustomers.UpdateCustomer(p) Then
        LogAction "EDIT CUSTOMER", p.customerID, "Updated " & p.customerName
        RefreshList
        ClearForm
        SetStatus "Updated " & p.customerName & "."
    Else
        MsgBox "That customer could not be updated.", vbExclamation, "Update failed"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "The customer could not be updated." & vbCrLf & Err.description, vbExclamation, "Update failed"
End Sub

Private Sub btnDelete_Click()
    Dim msg As String, deletedName As String, deletedID As String

    On Error GoTo ErrorHandler

    If Len(mSelectedID) = 0 Then
        MsgBox "Please select a customer from the list first.", vbInformation, "Delete"
        Exit Sub
    End If

    deletedID = mSelectedID
    deletedName = Txt("txtCustomerName").text

    If MsgBox("Delete " & deletedName & " (" & deletedID & ")?" & vbCrLf & _
              "This cannot be undone.", vbYesNo + vbQuestion, "Confirm delete") <> vbYes Then Exit Sub

    msg = modCustomers.DeleteCustomer(deletedID)
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Cannot delete"
    Else
        LogAction "DELETE CUSTOMER", deletedID, "Deleted " & deletedName
        RefreshList
        ClearForm
        SetStatus "Deleted " & deletedName & "."
    End If
    Exit Sub

ErrorHandler:
    MsgBox "The customer could not be deleted." & vbCrLf & Err.description, vbExclamation, "Delete failed"
End Sub

Private Sub btnClear_Click()
    ClearForm
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
