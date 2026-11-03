VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCharts 
   Caption         =   "Dashboard Charts"
   ClientHeight    =   9615.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "frmCharts.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmCharts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmCharts - four dashboard charts drawn with labels
' ============================================================

' ---------- Startup ----------
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    BuildLayout
    StyleButtons
    RefreshCharts
    Exit Sub

ErrorHandler:
    MsgBox "The charts screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Charts"
End Sub

' ---------- Where each chart sits (1 = top left, 2 = top right, 3 = bottom left, 4 = bottom right) ----------
Private Function ChartLeft(ByVal ch As Long) As Single
    Select Case ch
        Case 1, 3
            ChartLeft = 20
        Case Else
            ChartLeft = 410
    End Select
End Function

Private Function ChartTop(ByVal ch As Long) As Single
    Select Case ch
        Case 1, 2
            ChartTop = 50
        Case Else
            ChartTop = 248
    End Select
End Function

' Charts 1 and 4 have vertical bars; charts 2 and 3 have horizontal bars.
Private Function IsColumnChart(ByVal ch As Long) As Boolean
    IsColumnChart = (ch = 1 Or ch = 4)
End Function

' ---------- Build the panels and the (hidden) bars ----------
Private Sub BuildLayout()
    Dim titles As Variant
    Dim ch As Long, i As Long
    Dim X As Single, Y As Single
    Dim lbl As MSForms.label

    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 830, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "DASHBOARD CHARTS", -1, ThemeCyan(), 14, True

    titles = Array("MONTHLY SALES  (revenue, last 6 months)", _
                   "TOP PRODUCTS  (revenue, last 30 days)", _
                   "STOCK BY CATEGORY  (units in stock)", _
                   "PROFIT TREND  (net profit, last 6 months)")

    For ch = 1 To 4
        X = ChartLeft(ch)
        Y = ChartTop(ch)

        modUI.AddLabel Me, "ch" & ch & "Panel", X, Y, 380, 190, "", ThemePanel()
        modUI.AddLabel Me, "ch" & ch & "Title", X + 10, Y + 6, 360, 14, CStr(titles(ch - 1)), -1, ThemeCyan(), 8, True

        Set lbl = modUI.AddLabel(Me, "ch" & ch & "Empty", X + 10, Y + 90, 360, 20, "No data yet", -1, ThemeGrey(), 9)
        lbl.TextAlign = fmTextAlignCenter

        If IsColumnChart(ch) Then
            ' A baseline plus 6 columns, each with a value on top and a name below
            modUI.AddLabel Me, "ch" & ch & "Base", X + 12, Y + 100, 356, 1, "", RGB(70, 85, 115)
            For i = 0 To 5
                modUI.AddLabel Me, "ch" & ch & "Bar" & i, X + 12, Y + 100, 10, 1, "", ThemeCyan()

                Set lbl = modUI.AddLabel(Me, "ch" & ch & "Val" & i, X + 12, Y + 100, 50, 11, "", -1, ThemeWhite(), 7, True)
                lbl.TextAlign = fmTextAlignCenter

                Set lbl = modUI.AddLabel(Me, "ch" & ch & "Cat" & i, X + 12, Y + 172, 50, 12, "", -1, ThemeGrey(), 7)
                lbl.TextAlign = fmTextAlignCenter
            Next i
        Else
            ' 7 rows, each with a name, a bar and a value
            For i = 0 To 6
                modUI.AddLabel Me, "ch" & ch & "Nam" & i, X + 12, Y + 30, 112, 14, "", -1, ThemeWhite(), 8
                modUI.AddLabel Me, "ch" & ch & "Bar" & i, X + 128, Y + 30, 10, 13, "", ThemeGreen()
                modUI.AddLabel Me, "ch" & ch & "Val" & i, X + 140, Y + 30, 62, 13, "", -1, ThemeWhite(), 8, True
            Next i
        End If

        HideChart ch
    Next ch
End Sub

Private Sub StyleButtons()
    StyleButton btnRefresh, "REFRESH", 20, ThemeCyan(), ThemeNavy()
    StyleButton btnClose, "CLOSE", 650, ThemePanel(), ThemeWhite()
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal bgColor As Long, ByVal fgColor As Long)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = 446: .Width = 140: .Height = 30
        .BackColor = bgColor: .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = 10: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

' ---------- Hiding and empty messages ----------
Private Sub HideChart(ByVal ch As Long)
    Dim i As Long

    Me.Controls("ch" & ch & "Empty").Visible = False

    If IsColumnChart(ch) Then
        Me.Controls("ch" & ch & "Base").Visible = False
        For i = 0 To 5
            Me.Controls("ch" & ch & "Bar" & i).Visible = False
            Me.Controls("ch" & ch & "Val" & i).Visible = False
            Me.Controls("ch" & ch & "Cat" & i).Visible = False
        Next i
    Else
        For i = 0 To 6
            Me.Controls("ch" & ch & "Nam" & i).Visible = False
            Me.Controls("ch" & ch & "Bar" & i).Visible = False
            Me.Controls("ch" & ch & "Val" & i).Visible = False
        Next i
    End If
End Sub

Private Sub ShowEmpty(ByVal ch As Long, ByVal message As String)
    With Me.Controls("ch" & ch & "Empty")
        .caption = message
        .Visible = True
    End With
End Sub

Private Function ShortText(ByVal s As String, ByVal maxLen As Long) As String
    If Len(s) > maxLen Then
        ShortText = Left$(s, maxLen - 2) & ".."
    Else
        ShortText = s
    End If
End Function

' ---------- Vertical bars (positive bars go up, negative bars go down) ----------
Private Sub DrawColumns(ByVal ch As Long, ByRef labels() As String, ByRef values() As Double, _
                        ByVal n As Long, ByVal positiveColor As Long, ByVal negativeColor As Long, _
                        ByVal valueFormat As String)
    Dim X As Single, Y As Single
    Dim i As Long
    Dim maxVal As Double, minVal As Double
    Dim unitHeight As Double, zeroY As Double
    Dim slotW As Double, barW As Double, barH As Double, barTop As Double
    Dim barLbl As MSForms.label, valLbl As MSForms.label, catLbl As MSForms.label
    Dim baseLbl As MSForms.label

    HideChart ch
    If n > 6 Then n = 6
    If n < 1 Then
        ShowEmpty ch, "No data yet"
        Exit Sub
    End If

    X = ChartLeft(ch)
    Y = ChartTop(ch)

    For i = 0 To n - 1
        If values(i) > maxVal Then maxVal = values(i)
        If values(i) < minVal Then minVal = values(i)
    Next i

    If maxVal = 0 And minVal = 0 Then
        ShowEmpty ch, "No data yet for this period"
        Exit Sub
    End If

    unitHeight = 110 / (maxVal - minVal)          ' how tall one unit of money is
    zeroY = Y + 46 + maxVal * unitHeight          ' where the zero line sits
    slotW = 356 / n
    barW = slotW * 0.6

    Set baseLbl = Me.Controls("ch" & ch & "Base")
    baseLbl.Left = X + 12
    baseLbl.Top = zeroY
    baseLbl.Width = 356
    baseLbl.Height = 1
    baseLbl.Visible = True

    For i = 0 To n - 1
        Set barLbl = Me.Controls("ch" & ch & "Bar" & i)
        Set valLbl = Me.Controls("ch" & ch & "Val" & i)
        Set catLbl = Me.Controls("ch" & ch & "Cat" & i)

        If values(i) >= 0 Then
            barH = values(i) * unitHeight
            barTop = zeroY - barH
            barLbl.BackColor = positiveColor
        Else
            barH = -values(i) * unitHeight
            barTop = zeroY
            barLbl.BackColor = negativeColor
        End If
        If values(i) <> 0 And barH < 1 Then barH = 1

        If values(i) = 0 Then
            barLbl.Visible = False
        Else
            barLbl.Left = X + 12 + i * slotW + (slotW - barW) / 2
            barLbl.Top = barTop
            barLbl.Width = barW
            barLbl.Height = barH
            barLbl.Visible = True
        End If

        valLbl.Left = X + 12 + i * slotW
        valLbl.Width = slotW
        If values(i) >= 0 Then
            valLbl.Top = zeroY - barH - 12
        Else
            valLbl.Top = zeroY + barH + 1
        End If
        valLbl.caption = Format$(values(i), valueFormat)
        valLbl.Visible = True

        catLbl.Left = X + 12 + i * slotW
        catLbl.Width = slotW
        catLbl.caption = labels(i)
        catLbl.Visible = True
    Next i
End Sub

' ---------- Horizontal bars ----------
Private Sub DrawRows(ByVal ch As Long, ByRef labels() As String, ByRef values() As Double, _
                     ByVal n As Long, ByVal barColor As Long, ByVal valueFormat As String)
    Dim X As Single, Y As Single
    Dim i As Long
    Dim maxVal As Double, barW As Double, rowTop As Double
    Dim nameLbl As MSForms.label, barLbl As MSForms.label, valLbl As MSForms.label

    HideChart ch
    If n > 7 Then n = 7
    If n < 1 Then
        ShowEmpty ch, "No data yet"
        Exit Sub
    End If

    X = ChartLeft(ch)
    Y = ChartTop(ch)

    For i = 0 To n - 1
        If values(i) > maxVal Then maxVal = values(i)
    Next i

    If maxVal <= 0 Then
        ShowEmpty ch, "No data yet for this period"
        Exit Sub
    End If

    For i = 0 To n - 1
        rowTop = Y + 30 + i * 21

        Set nameLbl = Me.Controls("ch" & ch & "Nam" & i)
        Set barLbl = Me.Controls("ch" & ch & "Bar" & i)
        Set valLbl = Me.Controls("ch" & ch & "Val" & i)

        nameLbl.Left = X + 12
        nameLbl.Top = rowTop + 2
        nameLbl.Width = 112
        nameLbl.Height = 14
        nameLbl.caption = ShortText(labels(i), 20)
        nameLbl.Visible = True

        barW = values(i) / maxVal * 180
        If barW < 2 Then barW = 2
        barLbl.BackColor = barColor
        barLbl.Left = X + 128
        barLbl.Top = rowTop + 3
        barLbl.Width = barW
        barLbl.Height = 13
        barLbl.Visible = True

        valLbl.Left = X + 128 + barW + 4
        valLbl.Top = rowTop + 3
        valLbl.Width = 62
        valLbl.Height = 13
        valLbl.caption = Format$(values(i), valueFormat)
        valLbl.Visible = True
    Next i
End Sub

' ---------- Get the numbers and draw everything ----------
Private Sub RefreshCharts()
    Dim labels() As String, revenue() As Double, netProfit() As Double
    Dim names() As String, values() As Double
    Dim n As Long

    On Error GoTo ErrorHandler

    ' Charts 1 and 4: the last 6 months
    modCharts.GetMonthlySeries 6, labels, revenue, netProfit
    DrawColumns 1, labels, revenue, 6, ThemeCyan(), ThemeRed(), "#,##0"
    DrawColumns 4, labels, netProfit, 6, ThemeGreen(), ThemeRed(), "#,##0"

    ' Chart 2: top 5 products in the last 30 days
    n = modCharts.GetTopProducts(5, 30, names, values)
    If n > 0 Then
        DrawRows 2, names, values, n, ThemeGreen(), "#,##0"
    Else
        HideChart 2
        ShowEmpty 2, "No sales in the last 30 days"
    End If

    ' Chart 3: stock by category
    n = modCharts.GetStockByCategory(7, names, values)
    If n > 0 Then
        DrawRows 3, names, values, n, ThemeOrange(), "#,##0"
    Else
        HideChart 3
        ShowEmpty 3, "No products yet"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "The charts could not be drawn." & vbCrLf & Err.description, vbExclamation, "Charts"
End Sub

Private Sub btnRefresh_Click()
    RefreshCharts
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

