VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLowStock 
   Caption         =   "Low Stock Alerts"
   ClientHeight    =   9015.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "frmLowStock.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmLowStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmLowStock - products that need restocking
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

    RefreshAll
    Exit Sub

ErrorHandler:
    mLoading = False
    MsgBox "The low stock screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Low stock"
End Sub

Private Sub SetStatus(ByVal message As String)
    Me.Controls("lblStatusMsg").caption = message
End Sub

' ---------- Build the parts made by code ----------
Private Sub BuildLayout()
    Dim headings As Variant, widths As Variant
    Dim i As Long
    Dim X As Single

    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 830, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "LOW STOCK ALERTS", -1, ThemeCyan(), 14, True

    ' Warning banner (Segoe UI Symbol can draw the warning sign)
    With modUI.AddLabel(Me, "lblBanner", 20, 48, 770, 30, "", ThemePanel(), ThemeOrange(), 12, True)
        .Font.name = "Segoe UI Symbol"
    End With

    modUI.AddLabel Me, "lblShow", 20, 86, 170, 14, "Show", -1, ThemeGrey(), 8
    modUI.AddLabel Me, "lblSupplierFilter", 205, 86, 280, 14, "Supplier", -1, ThemeGrey(), 8

    ' Column headings (widths match modAlerts.FillLowStockList)
    headings = Array("ID", "PRODUCT", "CATEGORY", "SUPPLIER", "STOCK", "MIN", "MAX", "ORDER", "STATUS")
    widths = Array(62, 190, 85, 130, 45, 40, 40, 60, 90)
    X = 23
    For i = 0 To UBound(headings)
        modUI.AddLabel Me, "lblHead" & i, X, 132, widths(i), 13, headings(i), -1, ThemeCyan(), 7, True
        X = X + widths(i)
    Next i

    modUI.AddLabel Me, "lblStatusMsg", 20, 436, 770, 18, "", -1, ThemeGrey(), 9
End Sub

Private Sub StyleControls()
    With cmbShow
        .Left = 20: .Top = 101: .Width = 170: .Height = 22
        .style = fmStyleDropDownList
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
    End With

    With cmbSupplierFilter
        .Left = 205: .Top = 101: .Width = 280: .Height = 22
        .style = fmStyleDropDownList
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 10
    End With

    With lstLowStock
        .Left = 20: .Top = 146: .Width = 770: .Height = 236
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    StyleButton btnExport, "EXPORT TO REPORTS SHEET", 20, 392, 240, 36, ThemeCyan(), ThemeNavy()
    StyleButton btnClose, "CLOSE", 650, 392, 140, 36, ThemePanel(), ThemeWhite()
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
    Dim item As Variant

    For Each item In Array("Low + out of stock", "Low stock only", "Out of stock only")
        cmbShow.AddItem CStr(item)
    Next item
    cmbShow.ListIndex = 0

    modProducts.LoadSupplierCombo cmbSupplierFilter
    cmbSupplierFilter.List(0, 1) = "All suppliers"      ' the first entry has an empty ID = "no filter"
    cmbSupplierFilter.ListIndex = 0
End Sub

' ---------- What the two filters mean ----------
Private Function ModeCode() As String
    Select Case cmbShow.ListIndex
        Case 1: ModeCode = "LOW"
        Case 2: ModeCode = "OUT"
        Case Else: ModeCode = "ALL"
    End Select
End Function

Private Function SupplierFilterID() As String
    If cmbSupplierFilter.ListIndex >= 0 Then
        SupplierFilterID = CStr(cmbSupplierFilter.List(cmbSupplierFilter.ListIndex, 0))
    End If
End Function

' ---------- Refresh ----------
Private Sub RefreshAll()
    Dim shownCount As Long
    Dim lowCount As Long, outCount As Long
    Dim banner As MSForms.label

    If mLoading Then Exit Sub

    shownCount = modAlerts.FillLowStockList(lstLowStock, ModeCode(), SupplierFilterID())
    modAlerts.CountAlerts lowCount, outCount

    Set banner = Me.Controls("lblBanner")
    If lowCount + outCount > 0 Then
        banner.ForeColor = ThemeOrange()
        banner.caption = "  " & ChrW(9888) & "   " & lowCount & " low stock    |    " & outCount & " out of stock"
    Else
        banner.ForeColor = ThemeGreen()
        banner.caption = "  " & ChrW(10003) & "   All products are well stocked"
    End If

    If shownCount = 0 Then
        SetStatus "Nothing to show for this filter."
    Else
        SetStatus shownCount & " product(s) shown. Most urgent first."
    End If
End Sub

Private Sub cmbShow_Change()
    RefreshAll
End Sub

Private Sub cmbSupplierFilter_Change()
    RefreshAll
End Sub

' ---------- Export ----------
Private Sub btnExport_Click()
    Dim exported As Long

    On Error GoTo ErrorHandler

    exported = modAlerts.WriteLowStockReport(ModeCode(), SupplierFilterID())
    LogAction "EXPORT REPORT", "LOW STOCK", exported & " product(s) written to the Reports sheet"

        modSecurity.ShowSheet "Reports"

    MsgBox "The report was written to the Reports sheet (" & exported & " product(s))." & vbCrLf & _
           "Close the application windows to see and print it.", vbInformation, "Report exported"
    Exit Sub

ErrorHandler:
    MsgBox "The report could not be exported." & vbCrLf & Err.description, vbExclamation, "Export failed"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

