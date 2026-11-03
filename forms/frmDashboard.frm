VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDashboard 
   Caption         =   "Inventory Management System - Dashboard"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15765
   OleObjectBlob   =   "frmDashboard.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Hover-highlight objects (kept here so they stay alive while the form is open)
Private mHoverButtons As Collection
Private mHoverAreas As Collection
' ============================================================
' frmDashboard - main screen after login
' ============================================================

' ---------- Startup ----------
Private Sub UserForm_Initialize()
          On Error GoTo ErrorHandler

    BuildLayout
    StyleNavButtons
    AttachHover
    ApplyRolePermissions
    SetActiveNav "btnNavDashboard"
    RefreshData
    Exit Sub

ErrorHandler:
    MsgBox "The dashboard could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Dashboard"
End Sub

' ---------- Build the visual layout with code ----------
Private Sub BuildLayout()
    Dim cardTitles As Variant
    Dim i As Long
    Dim cardLeft As Single, cardTop As Single
    Dim sidebar As MSForms.label, headerBar As MSForms.label

    Me.BackColor = ThemeNavy()

    ' Header bar and sidebar background (sent behind the buttons)
    Set headerBar = AddLabel("lblHeaderBar", 0, 0, 820, 44, "", ThemeSidebar())
    headerBar.ZOrder 1
    Set sidebar = AddLabel("lblSidebar", 0, 44, 150, 480, "", ThemeSidebar())
    sidebar.ZOrder 1

       AddLabel "lblAppTitle", 12, 8, 300, 28, UCase$(GetSetting("ShopName", "Inventory System")), _
             -1, ThemeCyan(), 14, True
            With AddLabel("lblUserInfo", 300, 14, 340, 20, "", -1, ThemeGrey(), 9)
        .TextAlign = fmTextAlignRight
    End With

    ' Eight statistic cards: 4 per row
    cardTitles = Array("PRODUCTS", "TOTAL STOCK", "TODAY'S SALES", "TODAY'S PURCHASES", _
                       "TODAY'S GROSS PROFIT", "LOW STOCK ITEMS", "CUSTOMERS", "SUPPLIERS")
    For i = 0 To 7
        cardLeft = 162 + (i Mod 4) * 160
        cardTop = 56 + (i \ 4) * 76
        AddLabel "lblCard" & (i + 1) & "Back", cardLeft, cardTop, 148, 66, "", ThemePanel()
        AddLabel "lblCard" & (i + 1) & "Title", cardLeft + 10, cardTop + 8, 130, 16, _
                 cardTitles(i), -1, ThemeGrey(), 8, True
        AddLabel "lblCard" & (i + 1) & "Value", cardLeft + 10, cardTop + 26, 130, 32, _
                 "0", -1, ThemeCyan(), 16, True
    Next i

        ' Small extra line under the profit card (filled in by RefreshData)
    AddLabel "lblCard5Sub", 172, 182, 130, 14, "", -1, ThemeGrey(), 7, False
    AddLabel "lblRecentTitle", 162, 208, 306, 18, "RECENT SALES", -1, ThemeCyan(), 9, True
    AddList "lstRecentSales", 162, 228, 306, 120
    AddLabel "lblLowTitle", 484, 208, 306, 18, "LOW STOCK PRODUCTS", -1, ThemeOrange(), 9, True
    AddList "lstLowStock", 484, 228, 306, 120
    AddLabel "lblTopTitle", 162, 358, 306, 18, "TOP SELLING PRODUCTS", -1, ThemeCyan(), 9, True
    AddList "lstTopSelling", 162, 378, 306, 100
    ' The "Password" button in the header
    With btnPassword
        .Left = 650: .Top = 9: .Width = 90: .Height = 26
        .caption = "PASSWORD"
        .BackColor = ThemePanel()
        .ForeColor = ThemeCyan()
        .Font.name = "Segoe UI": .Font.Size = 8: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
       ' The "View charts" button fills the space where the placeholder was
       
    With btnCharts
        .Left = 484: .Top = 358: .Width = 306: .Height = 120
        .caption = "VIEW CHARTS" & vbCrLf & "Monthly sales, top products, stock and profit trend"
        .BackColor = ThemePanel()
        .ForeColor = ThemeCyan()
        .Font.name = "Segoe UI": .Font.Size = 11: .Font.Bold = True
        .WordWrap = True
        .TakeFocusOnClick = False
    End With
End Sub

' Creates a label with a consistent style. bgColor = -1 means transparent.
Private Function AddLabel(ByVal ctlName As String, ByVal leftPos As Single, ByVal topPos As Single, _
                          ByVal widthVal As Single, ByVal heightVal As Single, _
                          ByVal captionText As String, ByVal bgColor As Long, _
                          Optional ByVal fgColor As Long = &HFFFFFF, _
                          Optional ByVal fontSize As Single = 9, _
                          Optional ByVal isBold As Boolean = False) As MSForms.label
    Dim lbl As MSForms.label

    Set lbl = Me.Controls.Add("Forms.Label.1", ctlName, True)
    With lbl
        .Left = leftPos: .Top = topPos: .Width = widthVal: .Height = heightVal
        .caption = captionText
        .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = fontSize: .Font.Bold = isBold
        If bgColor = -1 Then
            .BackStyle = fmBackStyleTransparent
        Else
            .BackStyle = fmBackStyleOpaque
            .BackColor = bgColor
        End If
    End With
    Set AddLabel = lbl
End Function

' Creates a flat dark list box.
Private Sub AddList(ByVal ctlName As String, ByVal leftPos As Single, ByVal topPos As Single, _
                    ByVal widthVal As Single, ByVal heightVal As Single)
    Dim lst As MSForms.ListBox

    Set lst = Me.Controls.Add("Forms.ListBox.1", ctlName, True)
    With lst
        .Left = leftPos: .Top = topPos: .Width = widthVal: .Height = heightVal
        .BackColor = ThemePanel()
        .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

' ---------- Sidebar buttons ----------
Private Function NavNames() As Variant
    NavNames = Array("btnNavDashboard", "btnNavProducts", "btnNavSales", "btnNavPurchases", _
                     "btnNavStock", "btnNavCustomers", "btnNavSuppliers", "btnNavReports", _
                     "btnNavSettings", "btnNavLogout")
End Function

Private Sub StyleNavButtons()
    Dim names As Variant, captions As Variant
    Dim i As Long
    Dim btn As MSForms.CommandButton

    names = NavNames()
    captions = Array("Dashboard", "Products", "Sales", "Purchases", "Stock", _
                     "Customers", "Suppliers", "Reports", "Settings", "Logout")

    For i = 0 To UBound(names)
        Set btn = Me.Controls(names(i))
        With btn
            .caption = captions(i)
            .Left = 10: .Top = 58 + i * 42: .Width = 130: .Height = 34
            .BackColor = ThemeSidebar()
            .ForeColor = ThemeWhite()
            .Font.name = "Segoe UI": .Font.Size = 10
            .TakeFocusOnClick = False
        End With
    Next i

    btnNavLogout.ForeColor = ThemeRed()
End Sub

' Highlights the selected sidebar button.
Private Sub SetActiveNav(ByVal activeName As String)
    Dim names As Variant
    Dim i As Long

    names = NavNames()
    For i = 0 To UBound(names) - 1                 ' skip Logout
        If Me.Controls(names(i)).Enabled Then
            Me.Controls(names(i)).BackColor = ThemeSidebar()
            Me.Controls(names(i)).ForeColor = ThemeWhite()
        End If
    Next i

    Me.Controls(activeName).BackColor = ThemeActive()
    Me.Controls(activeName).ForeColor = ThemeCyan()
End Sub

' ---------- Role permissions ----------
Private Sub ApplyRolePermissions()
    Select Case modMain.CurrentRole
        Case "ADMIN"
            ' full access
        Case "MANAGER"
            DisableNav btnNavSettings
        Case Else                                   ' STAFF
            DisableNav btnNavPurchases
            DisableNav btnNavStock
            DisableNav btnNavSuppliers
            DisableNav btnNavReports
            DisableNav btnNavSettings
    End Select
End Sub

Private Sub DisableNav(ByVal btn As MSForms.CommandButton)
    btn.Enabled = False
    btn.ForeColor = RGB(90, 100, 120)
End Sub

' ---------- Refresh numbers and lists ----------
' Public so other forms can call frmDashboard.RefreshData after a sale.
Public Sub RefreshData()
    Dim cur As String
    Dim lst As MSForms.ListBox

    On Error GoTo ErrorHandler

    cur = GetSetting("CurrencySymbol", "Rs.") & " "

    Me.Controls("lblUserInfo").caption = modMain.CurrentFullName & "   |   " & _
        modMain.CurrentRole & "   |   " & Format$(Date, "dd mmm yyyy")

    Me.Controls("lblCard1Value").caption = Format$(modDashboard.CountRecords("Products", "tblProducts"), "#,##0")
    Me.Controls("lblCard2Value").caption = Format$(modDashboard.totalStock(), "#,##0")
    Me.Controls("lblCard3Value").caption = cur & Format$(modDashboard.TodaySales(), "#,##0")
    Me.Controls("lblCard4Value").caption = cur & Format$(modDashboard.TodayPurchases(), "#,##0")
    Me.Controls("lblCard5Value").caption = cur & Format$(modDashboard.TodayProfit(), "#,##0")
    Me.Controls("lblCard5Sub").caption = "Net this month: " & cur & Format$(modProfit.MonthNetProfit(), "#,##0")
    Me.Controls("lblCard6Value").caption = Format$(modDashboard.LowStockCount(), "#,##0")
    Me.Controls("lblCard7Value").caption = Format$(modCustomers.RealCustomerCount(), "#,##0")
    Me.Controls("lblCard8Value").caption = Format$(modDashboard.CountRecords("Suppliers", "tblSuppliers"), "#,##0")

    ' Orange warning colour when something is low
    If modDashboard.LowStockCount() > 0 Then
        Me.Controls("lblCard6Value").ForeColor = ThemeOrange()
    Else
        Me.Controls("lblCard6Value").ForeColor = ThemeGreen()
    End If

    Set lst = Me.Controls("lstRecentSales")
    modDashboard.FillRecentSales lst, 8
    Set lst = Me.Controls("lstLowStock")
    modDashboard.FillLowStock lst, 10
    Set lst = Me.Controls("lstTopSelling")
    modDashboard.FillTopSelling lst, 5
        HideFinancialsForStaff
    Exit Sub

ErrorHandler:
    MsgBox "Could not refresh the dashboard." & vbCrLf & Err.description, vbExclamation, "Dashboard"
End Sub

' ---------- Sidebar clicks ----------
Private Sub btnNavDashboard_Click()
    SetActiveNav "btnNavDashboard"
    RefreshData
End Sub

Private Sub btnNavProducts_Click()
    SetActiveNav "btnNavProducts"
    frmProducts.Show
    RefreshData                          ' update the cards after the form closes
    SetActiveNav "btnNavDashboard"
End Sub
Private Sub btnNavSales_Click()
    SetActiveNav "btnNavSales"
    frmSales.Show
    RefreshData
    SetActiveNav "btnNavDashboard"
End Sub

Private Sub btnNavPurchases_Click()
    SetActiveNav "btnNavPurchases"
    frmPurchase.Show
    RefreshData
    SetActiveNav "btnNavDashboard"
End Sub
Private Sub btnNavStock_Click()
    SetActiveNav "btnNavStock"
    frmStock.Show
    RefreshData
    SetActiveNav "btnNavDashboard"
End Sub
Private Sub btnNavCustomers_Click()
    SetActiveNav "btnNavCustomers"
    frmCustomers.Show
    RefreshData
    SetActiveNav "btnNavDashboard"
End Sub

Private Sub btnNavSuppliers_Click()
    SetActiveNav "btnNavSuppliers"
    frmSuppliers.Show
    RefreshData
    SetActiveNav "btnNavDashboard"
End Sub
Private Sub btnCharts_Click()
    frmCharts.Show
End Sub

Private Sub btnNavReports_Click()
    SetActiveNav "btnNavReports"
    frmReports.Show
    RefreshData
    SetActiveNav "btnNavDashboard"
End Sub


Private Sub ComingSoon(ByVal moduleName As String, ByVal phaseNumber As Long)
    MsgBox "The " & moduleName & " screen will be built in Phase " & phaseNumber & ".", _
           vbInformation, moduleName
End Sub
Private Sub btnNavSettings_Click()
    SetActiveNav "btnNavSettings"
    frmAdmin.Show
    RefreshData
    SetActiveNav "btnNavDashboard"
End Sub
' Shop staff do not see profit, purchase totals or the charts.
Private Sub HideFinancialsForStaff()
    If modMain.CurrentRole <> "STAFF" Then Exit Sub

    Me.Controls("lblCard4Value").caption = "-"
    Me.Controls("lblCard5Value").caption = "-"
    Me.Controls("lblCard5Sub").caption = ""
    btnCharts.Enabled = False
    btnCharts.ForeColor = RGB(90, 100, 120)
End Sub

Private Sub btnPassword_Click()
    frmPassword.Show
End Sub
' ---------- Hover highlight for the sidebar ----------
Private Sub AttachHover()
    Dim names As Variant
    Dim i As Long
    Dim btn As MSForms.CommandButton
    Dim area As MSForms.label
    Dim hb As clsHoverButton
    Dim ha As clsHoverReset

    Set mHoverButtons = New Collection
    Set mHoverAreas = New Collection

    names = NavNames()
    For i = 0 To UBound(names)
        Set btn = Me.Controls(CStr(names(i)))
        Set hb = New clsHoverButton
        hb.Attach btn, ThemeSidebar(), ThemeHover()
        mHoverButtons.Add hb
    Next i

    ' Moving over the sidebar background or the header removes the highlight
    Set area = Me.Controls("lblSidebar")
    Set ha = New clsHoverReset
    ha.Attach area
    mHoverAreas.Add ha

    Set area = Me.Controls("lblHeaderBar")
    Set ha = New clsHoverReset
    ha.Attach area
    mHoverAreas.Add ha
End Sub

' Moving over empty parts of the form also removes the highlight.
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    modHover.ClearHover
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    modHover.ClearHover
End Sub
Private Sub btnNavLogout_Click()
    If MsgBox("Do you want to log out?", vbYesNo + vbQuestion, "Logout") = vbYes Then
        LogAction "LOGOUT", modMain.CurrentUserID, modMain.CurrentUsername & " logged out"
        modMain.LogoutRequested = True
        Unload Me
    End If
End Sub

