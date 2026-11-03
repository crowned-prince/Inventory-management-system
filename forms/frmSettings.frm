VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Shop Settings"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10965
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmSettings - shop details and program options (ADMIN only)
' ============================================================

' ---------- Startup ----------
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    BuildLayout
    StyleButtons
    LoadLookups
    LoadValues
    UpdateButtonState
    Exit Sub

ErrorHandler:
    MsgBox "The settings screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Settings"
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

' ---------- Build the parts made by code ----------
Private Sub BuildLayout()
    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 580, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "SHOP SETTINGS", -1, ThemeCyan(), 14, True

    ' Left column
    AddField "Shop Name *", "txtShopName", "text", 1, 0
    AddField "Address", "txtShopAddress", "text", 1, 1
    AddField "Phone", "txtShopPhone", "text", 1, 2
    AddField "Invoice footer (thank-you line)", "txtInvoiceFooter", "text", 1, 3
    AddField "Backup folder (empty = a Backups folder next to this file)", "txtBackupFolder", "text", 1, 4

    ' Right column
    AddField "Currency symbol *", "txtCurrency", "text", 2, 0
    AddField "Tax on sales", "cmbTaxEnabled", "combo", 2, 1
    AddField "Tax rate (%)", "txtTaxRate", "text", 2, 2
    AddField "Auto-save after sales/purchases", "cmbAutoSave", "combo", 2, 3
    AddField "Minimize Excel while app runs", "cmbAppMode", "combo", 2, 4

    Txt("txtShopName").MaxLength = 80
    Txt("txtShopAddress").MaxLength = 120
    Txt("txtShopPhone").MaxLength = 40
    Txt("txtInvoiceFooter").MaxLength = 120
    Txt("txtBackupFolder").MaxLength = 200
    Txt("txtCurrency").MaxLength = 8

    With modUI.AddLabel(Me, "lblStatusMsg", 20, 358, 510, 36, "", -1, ThemeGrey(), 9)
        .WordWrap = True
    End With
End Sub

Private Sub AddField(ByVal captionText As String, ByVal ctlName As String, _
                     ByVal kind As String, ByVal columnNo As Long, ByVal rowNo As Long)
    Dim X As Single, Y As Single, w As Single

    If columnNo = 1 Then
        X = 20
        w = 300
    Else
        X = 340
        w = 190
    End If
    Y = 56 + (rowNo * 48)

    modUI.AddLabel Me, "lbl" & Mid$(ctlName, 4), X, Y, w, 14, captionText, -1, ThemeGrey(), 8

    If kind = "combo" Then
        modUI.AddCombo Me, ctlName, X, Y + 15, w, 22
    Else
        modUI.AddTextBox Me, ctlName, X, Y + 15, w, 22
    End If
End Sub

Private Sub StyleButtons()
    StyleButton btnSave, "SAVE SETTINGS", 20, 200, ThemeGreen(), ThemeNavy()
    StyleButton btnClose, "CLOSE", 380, 150, ThemePanel(), ThemeWhite()
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal widthVal As Single, _
                        ByVal bgColor As Long, ByVal fgColor As Long)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = 310: .Width = widthVal: .Height = 36
        .BackColor = bgColor: .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = 10: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub LoadLookups()
    Dim names As Variant
    Dim item As Variant

    names = Array("cmbTaxEnabled", "cmbAutoSave", "cmbAppMode")
    For Each item In names
        Cmb(CStr(item)).AddItem "Yes"
        Cmb(CStr(item)).AddItem "No"
    Next item
End Sub

' Selects Yes or No in a combo.
Private Sub SelectYesNo(ByVal cb As MSForms.ComboBox, ByVal value As String)
    If UCase$(Trim$(value)) = "YES" Then
        cb.ListIndex = 0
    Else
        cb.ListIndex = 1
    End If
End Sub

' Shows the current settings.
Private Sub LoadValues()
    Txt("txtShopName").text = GetSetting("ShopName", "My Retail Shop")
    Txt("txtShopAddress").text = GetSetting("ShopAddress", "")
    Txt("txtShopPhone").text = GetSetting("ShopPhone", "")
    Txt("txtInvoiceFooter").text = GetSetting("InvoiceFooter", "Thank you for shopping with us!")
    Txt("txtBackupFolder").text = GetSetting("BackupFolder", "")
    Txt("txtCurrency").text = GetSetting("CurrencySymbol", "Rs.")
    Txt("txtTaxRate").text = GetSetting("TaxRate", "0")

    SelectYesNo Cmb("cmbTaxEnabled"), GetSetting("TaxEnabled", "No")
    SelectYesNo Cmb("cmbAutoSave"), GetSetting("AutoSave", "Yes")
    SelectYesNo Cmb("cmbAppMode"), GetSetting("AppMode", "No")
End Sub

Private Sub UpdateButtonState()
    btnSave.Enabled = modMain.IsAdmin()
    If Not modMain.IsAdmin() Then SetStatus "View only: only an administrator can change settings."
End Sub

' True if the folder exists.
Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(folderPath) And vbDirectory) <> 0)
    If Err.Number <> 0 Then FolderExists = False
    On Error GoTo 0
End Function

' ---------- Validation: "" = OK, otherwise the message ----------
Private Function ValidateForm() As String
    Dim msg As String
    Dim folderText As String

    msg = CheckRequired(Txt("txtShopName").text, "shop name")
    If Len(msg) > 0 Then
        ValidateForm = msg
        Exit Function
    End If

    msg = CheckRequired(Txt("txtCurrency").text, "currency symbol")
    If Len(msg) > 0 Then
        ValidateForm = msg
        Exit Function
    End If

    msg = CheckNumber(Txt("txtTaxRate").text, "tax rate")
    If Len(msg) > 0 Then
        ValidateForm = msg
        Exit Function
    End If
    If CDbl(Txt("txtTaxRate").text) > 100 Then
        ValidateForm = "The tax rate cannot be more than 100."
        Exit Function
    End If

    folderText = Trim$(Txt("txtBackupFolder").text)
    If Len(folderText) > 0 Then
        If Not FolderExists(folderText) Then
            ValidateForm = "The backup folder does not exist. Create it first, or leave the box empty."
            Exit Function
        End If
    End If
End Function

' ---------- Save ----------
Private Sub btnSave_Click()
    Dim msg As String

    On Error GoTo ErrorHandler

    msg = ValidateForm()
    If Len(msg) > 0 Then
        MsgBox msg, vbExclamation, "Please check your input"
        Exit Sub
    End If

    SetSetting "ShopName", Trim$(Txt("txtShopName").text)
    SetSetting "ShopAddress", Trim$(Txt("txtShopAddress").text)
    SetSetting "ShopPhone", Trim$(Txt("txtShopPhone").text)
    SetSetting "InvoiceFooter", Trim$(Txt("txtInvoiceFooter").text)
    SetSetting "BackupFolder", Trim$(Txt("txtBackupFolder").text)
    SetSetting "CurrencySymbol", Trim$(Txt("txtCurrency").text)
    SetSetting "TaxRate", Trim$(Txt("txtTaxRate").text)
    SetSetting "TaxEnabled", Cmb("cmbTaxEnabled").text
    SetSetting "AutoSave", Cmb("cmbAutoSave").text
    SetSetting "AppMode", Cmb("cmbAppMode").text

    LogAction "EDIT SETTINGS", "-", "Shop settings updated"

    SetStatus "Settings saved. The shop name in the dashboard header and the window mode apply the next time the application starts."
    MsgBox "Settings saved.", vbInformation, "Shop settings"
    Exit Sub

ErrorHandler:
    MsgBox "The settings could not be saved." & vbCrLf & Err.description, vbExclamation, "Save failed"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

