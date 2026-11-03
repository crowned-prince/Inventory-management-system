VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPickCategory 
   Caption         =   "Choose a Category"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   OleObjectBlob   =   "frmPickCategory.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPickCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmPickCategory - choose the category the Dashboard sheet should focus on
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    BuildLayout
    StyleControls
    LoadCategories
    Exit Sub

ErrorHandler:
    MsgBox "The category list could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Choose a category"
End Sub

Private Sub BuildLayout()
    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 420, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 300, 26, "CHOOSE A CATEGORY", -1, ThemeCyan(), 14, True
    modUI.AddLabel Me, "lblHint", 20, 46, 340, 14, "Double-click a category, or select it and click SHOW CATEGORY", -1, ThemeGrey(), 8
End Sub

Private Sub StyleControls()
    With lstCategories
        .Left = 20: .Top = 66: .Width = 340: .Height = 236
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    StyleButton btnAll, "ALL CATEGORIES", 20, 110, ThemePanel(), ThemeWhite()
    StyleButton btnSelect, "SHOW CATEGORY", 138, 130, ThemeCyan(), ThemeNavy()
    StyleButton btnClose, "CLOSE", 276, 84, ThemePanel(), ThemeWhite()
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal widthVal As Single, _
                        ByVal bgColor As Long, ByVal fgColor As Long)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = 312: .Width = widthVal: .Height = 34
        .BackColor = bgColor: .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = 9: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

Private Sub LoadCategories()
    Dim tbl As ListObject
    Dim data As Variant
    Dim i As Long

    lstCategories.Clear

    Set tbl = modDatabase.GetTable("Categories", "tblCategories")
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    data = tbl.DataBodyRange.value
    For i = 1 To UBound(data, 1)
        If Len(CStr(data(i, 2))) > 0 Then lstCategories.AddItem CStr(data(i, 2))
    Next i
End Sub

Private Sub ChooseSelected()
    If lstCategories.ListIndex < 0 Then
        MsgBox "Please select a category first.", vbInformation, "Choose a category"
        Exit Sub
    End If

    modDash.SetCategoryFilter CStr(lstCategories.List(lstCategories.ListIndex))
    Unload Me
End Sub

Private Sub btnSelect_Click()
    ChooseSelected
End Sub

Private Sub lstCategories_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ChooseSelected
End Sub

Private Sub btnAll_Click()
    modDash.SetCategoryFilter ""
    Unload Me
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

