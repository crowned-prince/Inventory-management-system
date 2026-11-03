Attribute VB_Name = "modDashBuild"
Option Explicit

' ============================================================
' modDashBuild - the look of the Dashboard sheet: layout, buttons and charts
' ============================================================

Private Const SHEET_NAME As String = "Dashboard"
Private Const DATA_NAME As String = "DashData"

' ---------- Colours ----------
Public Function PalColor(ByVal index As Long) As Long
    Select Case ((index - 1) Mod 8) + 1
        Case 1: PalColor = RGB(0, 194, 255)       ' cyan
        Case 2: PalColor = RGB(155, 89, 255)      ' purple
        Case 3: PalColor = RGB(255, 77, 125)      ' pink-red
        Case 4: PalColor = RGB(255, 168, 0)       ' amber
        Case 5: PalColor = RGB(46, 230, 166)      ' green
        Case 6: PalColor = RGB(90, 120, 255)      ' blue
        Case 7: PalColor = RGB(230, 70, 200)      ' magenta
        Case Else: PalColor = RGB(0, 160, 170)    ' teal
    End Select
End Function

Private Function CardColor(ByVal index As Long) As Long
    Select Case index
        Case 1: CardColor = RGB(0, 194, 255)
        Case 2: CardColor = RGB(46, 230, 166)
        Case 3: CardColor = RGB(155, 89, 255)
        Case 4: CardColor = RGB(255, 168, 0)
        Case 5: CardColor = RGB(255, 77, 125)
        Case Else: CardColor = RGB(0, 160, 170)
    End Select
End Function

' ---------- Sheets ----------

' The hidden helper sheet that feeds the charts (created if missing).
Public Function GetDataSheet() As Worksheet
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.name = DATA_NAME Then
            Set GetDataSheet = ws
            Exit Function
        End If
    Next ws

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.name = DATA_NAME
    Set GetDataSheet = ws
End Function

' True when the five charts are on the Dashboard sheet.
Public Function ChartsExist() As Boolean
    Dim ws As Worksheet
    Dim co As ChartObject
    Dim found As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    For Each co In ws.ChartObjects
        If Left$(co.name, 4) = "dash" Then found = found + 1
    Next co
    ChartsExist = (found >= 5)
End Function

' ---------- Building the sheet ----------
Public Sub BuildDashboardSheet()
    Dim ws As Worksheet
    Dim i As Long
    Dim r As Range
    Dim blockAddr As Variant
    Dim shp As Shape
    Dim errText As String

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    GetDataSheet
    modDash.UnlockDashboardSheets                                              ' make sure the helper sheet exists

    ' 1. Remove old charts and buttons
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

    ' 2. The dark canvas
    ws.Cells.Clear
    ws.Cells.Interior.Color = ThemeNavy()
    ws.Cells.Font.name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Cells.Font.Color = RGB(214, 220, 240)

    ws.Columns("A").ColumnWidth = 1.5
    ws.Range("B:M").ColumnWidth = 13
    ws.Columns("N").ColumnWidth = 1.5

    ws.Rows("1:60").RowHeight = 15
    ws.Rows(1).RowHeight = 34
    ws.Rows(2).RowHeight = 22
    ws.Rows(3).RowHeight = 8
    ws.Rows(4).RowHeight = 36
    ws.Rows(5).RowHeight = 8
    ws.Rows(6).RowHeight = 20
    ws.Rows(7).RowHeight = 34
    ws.Rows(8).RowHeight = 20
    ws.Rows(9).RowHeight = 8
    ws.Rows(10).RowHeight = 8
    ws.Rows(27).RowHeight = 8
    ws.Rows(44).RowHeight = 8
    ws.Rows(57).RowHeight = 8

    ' 3. Title
    With ws.Range("B1")
        .value = UCase$(GetSetting("ShopName", "Inventory Management System"))
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(0, 194, 255)
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B2").Font.Color = RGB(150, 160, 190)

    ' 4. The six KPI cards
    For i = 1 To 6
        Set r = ws.Range(ws.Cells(6, 2 * i), ws.Cells(8, 2 * i + 1))
        r.Interior.Color = ThemePanel()

        With r.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = CardColor(i)
        End With
        With r.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = ThemeNavy()
        End With

        With ws.Range(ws.Cells(6, 2 * i), ws.Cells(8, 2 * i))
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
        End With
        With ws.Cells(6, 2 * i)
            .Font.Size = 8
            .Font.Bold = True
            .Font.Color = RGB(150, 160, 190)
            .VerticalAlignment = xlBottom
        End With
        With ws.Cells(7, 2 * i)
            .Font.Size = 16
            .Font.Bold = True
            .Font.Color = CardColor(i)
            .VerticalAlignment = xlCenter
        End With
        With ws.Cells(8, 2 * i)
            .Font.Size = 8
            .Font.Color = RGB(150, 160, 190)
            .VerticalAlignment = xlTop
        End With
    Next i

    ' 5. The two tables: section titles, headers and banding
    ws.Range("B45").Font.Bold = True
    ws.Range("B45").Font.Color = RGB(0, 194, 255)
    ws.Range("I45").Font.Bold = True
    ws.Range("I45").Font.Color = RGB(0, 194, 255)

    For Each blockAddr In Array("B46:H46", "I46:M46")
        With ws.Range(CStr(blockAddr))
            .Interior.Color = RGB(40, 52, 92)
            .Font.Bold = True
            .Font.Size = 8
            .Font.Color = RGB(150, 200, 255)
        End With
    Next blockAddr

    For i = 47 To 56
        If i Mod 2 = 0 Then
            ws.Range("B" & i & ":H" & i).Interior.Color = ThemePanel()
            ws.Range("I" & i & ":M" & i).Interior.Color = ThemePanel()
        Else
            ws.Range("B" & i & ":H" & i).Interior.Color = RGB(34, 46, 82)
            ws.Range("I" & i & ":M" & i).Interior.Color = RGB(34, 46, 82)
        End If
    Next i

    ws.Range("E47:E56").numberFormat = "General"
    ws.Range("F47:G56").numberFormat = "#,##0"
    ws.Range("K47:M56").numberFormat = "#,##0"
    ws.Range("E46:H56").HorizontalAlignment = xlRight
    ws.Range("K46:M56").HorizontalAlignment = xlRight

    ' Data bars behind the revenue columns (cosmetic: skipped on old Excel versions)
    On Error Resume Next
    With ws.Range("F47:F56").FormatConditions.AddDatabar
        .barColor.Color = RGB(0, 150, 220)
        .BarFillType = xlDataBarFillGradient
        .BarBorder.Type = xlDataBarBorderNone
    End With
    With ws.Range("M47:M56").FormatConditions.AddDatabar
        .barColor.Color = RGB(155, 89, 255)
        .BarFillType = xlDataBarFillGradient
        .BarBorder.Type = xlDataBarBorderNone
    End With
    On Error GoTo ErrorHandler

    ' 6. Footnotes
    ws.Range("B58").value = "All products: revenue = sales after discounts, excluding tax (same as the Reports). " & _
                            "One product or category: revenue = the sum of its line totals."
    ws.Range("B59").value = "Cards compare with the previous period of the same length. Gross profit = revenue - cost of the goods sold."
    ws.Range("B58:B60").Font.Size = 8
    ws.Range("B58:B60").Font.Color = RGB(130, 140, 175)

    ' 7. The five charts
    NewChart ws, "dashTrend", ws.Range("B11:H26")
    NewChart ws, "dashCategory", ws.Range("I11:M26")
    NewChart ws, "dashTop", ws.Range("B28:E43")
    NewChart ws, "dashBuy", ws.Range("F28:I43")
    NewChart ws, "dashStock", ws.Range("J28:M43")

    ' 8. The buttons
    Set shp = AddButton(ws, "dashBtnProduct", ws.Range("B4:D4"), "PRODUCT:  All products", RGB(34, 46, 90), "DashPickProduct")
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(0, 194, 255)

    Set shp = AddButton(ws, "dashBtnCategory", ws.Range("E4:F4"), "CATEGORY:  All", RGB(34, 46, 90), "DashPickCategory")
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(155, 89, 255)

    Set shp = AddButton(ws, "dashBtnPeriod", ws.Range("G4:H4"), "PERIOD:  All time", RGB(34, 46, 90), "DashPickPeriod")
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(255, 168, 0)

    Set shp = AddButton(ws, "dashBtnReset", ws.Range("I4:I4"), "RESET", RGB(70, 80, 120), "DashReset")
    Set shp = AddButton(ws, "dashBtnRefresh", ws.Range("J4:J4"), "REFRESH", RGB(30, 150, 120), "DashRefresh")

    Set shp = AddButton(ws, "dashBtnStart", ws.Range("K4:M4"), "OPEN APPLICATION   (Ctrl+Shift+I)", RGB(0, 120, 255), "StartApp")
    On Error Resume Next
    shp.Fill.TwoColorGradient msoGradientHorizontal, 1
    shp.Fill.ForeColor.RGB = RGB(0, 190, 255)
    shp.Fill.BackColor.RGB = RGB(110, 70, 255)
    On Error GoTo ErrorHandler

    ' 9. The privacy cover (shown until an administrator or manager has logged in)
    Set shp = ws.Shapes.AddShape(1, ws.Range("B6").Left, ws.Range("B6").Top, _
                                 ws.Range("B6:M60").Width, ws.Range("B6:M60").Height)
    shp.name = "dashCover"
    shp.Fill.ForeColor.RGB = ThemeNavy()
    shp.Line.Visible = msoFalse
    With shp.TextFrame.Characters
        .text = "ANALYTICS LOCKED" & vbCrLf & vbCrLf & _
                "Click OPEN APPLICATION and log in as an administrator or manager."
        .Font.name = "Segoe UI"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(150, 160, 190)
    End With
    shp.Visible = msoFalse

    ' 10. Fill it with data
    SetStartShortcut
    modDash.SetUnlocked True
    modDash.RefreshDashboardSheet

    ' 11. Window options (cosmetic)
    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 100
    ws.Range("A1").Select
    On Error GoTo ErrorHandler

    modSecurity.EnterAppMode                              ' hides the helper sheet when hiding is on
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    errText = Err.description
    Application.ScreenUpdating = True
    Err.Raise vbObjectError + 700, , errText
End Sub

' Adds an empty chart over a block of cells.
Private Sub NewChart(ByVal ws As Worksheet, ByVal chartName As String, ByVal area As Range)
    Dim co As ChartObject

    Set co = ws.ChartObjects.Add(area.Left + 3, area.Top + 3, area.Width - 6, area.Height - 6)
    co.name = chartName
End Sub

' Adds a rounded button that runs a macro when clicked.
Private Function AddButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal area As Range, _
                           ByVal captionText As String, ByVal fillColor As Long, _
                           ByVal macroName As String) As Shape
    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(5, area.Left + 2, area.Top + 2, area.Width - 4, area.Height - 4)
    shp.name = shapeName
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.Visible = msoFalse
    shp.OnAction = macroName

    With shp.TextFrame
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        With .Characters
            .text = captionText
            .Font.name = "Segoe UI"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With
    End With

    Set AddButton = shp
End Function

' Ctrl+Shift+I starts the application (cosmetic: ignored if Excel refuses).
Private Sub SetStartShortcut()
    On Error Resume Next
    Application.MacroOptions Macro:="'" & ThisWorkbook.name & "'!StartApp", ShortcutKey:="I"
    On Error GoTo 0
End Sub

' Changes the text on one of the sheet buttons.
Public Sub SetButtonCaption(ByVal shapeName As String, ByVal captionText As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_NAME).Shapes(shapeName).TextFrame.Characters.text = captionText
    On Error GoTo 0
End Sub

' ---------- Chart helpers ----------
Private Sub ClearSeries(ByVal ch As Chart)
    Do While ch.SeriesCollection.count > 0
        ch.SeriesCollection(1).Delete
    Loop
End Sub

Private Function AddSeries(ByVal ch As Chart, ByVal seriesName As String, _
                           ByVal valuesRange As Range, ByVal labelsRange As Range) As Series
    Dim s As Series

    Set s = ch.SeriesCollection.NewSeries
    s.name = seriesName
    s.values = valuesRange
    s.XValues = labelsRange
    Set AddSeries = s
End Function

Private Sub StyleCommon(ByVal ch As Chart, ByVal titleText As String)
    On Error Resume Next
    ch.ChartArea.Format.Fill.Visible = msoTrue
    ch.ChartArea.Format.Fill.ForeColor.RGB = ThemePanel()
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.ChartArea.RoundedCorners = True
    ch.PlotArea.Format.Fill.Visible = msoFalse
    ch.PlotArea.Format.Line.Visible = msoFalse

    ch.HasTitle = True
    ch.ChartTitle.text = titleText
    ch.ChartTitle.Font.name = "Segoe UI"
    ch.ChartTitle.Font.Size = 11
    ch.ChartTitle.Font.Bold = True
    ch.ChartTitle.Font.Color = RGB(235, 238, 250)
    On Error GoTo 0
End Sub

Private Sub StyleAxes(ByVal ch As Chart, ByVal valueFormat As String)
    On Error Resume Next
    With ch.Axes(xlCategory)
        .TickLabels.Font.Color = RGB(190, 198, 222)
        .TickLabels.Font.Size = 8
        .Format.Line.ForeColor.RGB = RGB(70, 82, 120)
    End With
    With ch.Axes(xlValue)
        .TickLabels.Font.Color = RGB(190, 198, 222)
        .TickLabels.Font.Size = 8
        .TickLabels.numberFormat = valueFormat
        .HasMajorGridlines = True
        .MajorGridlines.Format.Line.ForeColor.RGB = RGB(52, 62, 100)
        .Format.Line.Visible = msoFalse
    End With
    On Error GoTo 0
End Sub

Private Sub StyleLegend(ByVal ch As Chart, ByVal position As Long)
    On Error Resume Next
    ch.HasLegend = True
    ch.Legend.position = position
    ch.Legend.Font.Color = RGB(200, 207, 230)
    ch.Legend.Font.Size = 8
    On Error GoTo 0
End Sub

Private Sub StyleTrend(ByVal ch As Chart)
    Dim s1 As Series, s2 As Series

    On Error Resume Next
    ch.ChartType = xlArea
    Set s1 = ch.SeriesCollection(1)
    Set s2 = ch.SeriesCollection(2)
    s2.ChartType = xlLineMarkers

    With s1.Format.Fill
        .Visible = msoTrue
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = RGB(0, 194, 255)
        .BackColor.RGB = RGB(30, 40, 78)
    End With
    s1.Format.Line.Visible = msoFalse

    With s2
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(255, 77, 125)
        .Format.Line.Weight = 2.5
        .MarkerStyle = xlMarkerStyleCircle
        .MarkerSize = 7
        .MarkerBackgroundColor = RGB(255, 77, 125)
        .MarkerForegroundColor = RGB(255, 255, 255)
        .Smooth = True
    End With

    StyleAxes ch, "#,##0"
    StyleLegend ch, xlLegendPositionBottom
    On Error GoTo 0
End Sub

Private Sub StyleDonut(ByVal ch As Chart, ByVal n As Long)
    Dim i As Long

    On Error Resume Next
    ch.ChartType = xlDoughnut
    ch.ChartGroups(1).DoughnutHoleSize = 62

    For i = 1 To n
        ch.SeriesCollection(1).Points(i).Format.Fill.ForeColor.RGB = PalColor(i)
        ch.SeriesCollection(1).Points(i).Format.Line.ForeColor.RGB = RGB(30, 40, 78)
    Next i

    ch.SeriesCollection(1).HasDataLabels = True
    With ch.SeriesCollection(1).DataLabels
        .ShowPercentage = True
        .ShowValue = False
        .ShowCategoryName = False
        .numberFormat = "0%"
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 8
        .Font.Bold = True
    End With

    StyleLegend ch, xlLegendPositionRight
    On Error GoTo 0
End Sub

Private Sub StyleTopBars(ByVal ch As Chart)
    Dim s As Series

    On Error Resume Next
    ch.ChartType = xlBarClustered
    Set s = ch.SeriesCollection(1)

    With s.Format.Fill
        .Visible = msoTrue
        .TwoColorGradient msoGradientVertical, 1
        .ForeColor.RGB = RGB(255, 168, 0)
        .BackColor.RGB = RGB(255, 77, 125)
    End With

    ch.ChartGroups(1).GapWidth = 55
    ch.Axes(xlCategory).ReversePlotOrder = True
    ch.Axes(xlCategory).TickLabels.Font.Color = RGB(190, 198, 222)
    ch.Axes(xlCategory).TickLabels.Font.Size = 8
    ch.Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(70, 82, 120)

    s.HasDataLabels = True
    s.DataLabels.numberFormat = "#,##0"
    s.DataLabels.Font.Color = RGB(235, 238, 250)
    s.DataLabels.Font.Size = 8

    ch.HasAxis(xlValue, xlPrimary) = False
    ch.HasLegend = False
    On Error GoTo 0
End Sub

Private Sub StyleBuy(ByVal ch As Chart)
    Dim s1 As Series, s2 As Series

    On Error Resume Next
    ch.ChartType = xlColumnClustered
    Set s1 = ch.SeriesCollection(1)
    Set s2 = ch.SeriesCollection(2)

    With s1.Format.Fill
        .Visible = msoTrue
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = RGB(0, 194, 255)
        .BackColor.RGB = RGB(155, 89, 255)
    End With
    With s2.Format.Fill
        .Visible = msoTrue
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = RGB(255, 168, 0)
        .BackColor.RGB = RGB(255, 77, 125)
    End With

    ch.ChartGroups(1).GapWidth = 70
    ch.ChartGroups(1).Overlap = -5

    StyleAxes ch, "#,##0"
    StyleLegend ch, xlLegendPositionBottom
    On Error GoTo 0
End Sub

Private Sub StylePie(ByVal ch As Chart, ByVal wd As Worksheet, ByVal n As Long)
    Dim i As Long
    Dim label As String
    Dim sliceColor As Long

    On Error Resume Next
    ch.ChartType = xlPie

    For i = 1 To n
        label = LCase$(CStr(wd.Cells(1 + i, 16).value))
        Select Case True
            Case InStr(label, "out of stock") > 0: sliceColor = RGB(255, 77, 125)
            Case InStr(label, "low") > 0: sliceColor = RGB(255, 168, 0)
            Case InStr(label, "in stock") > 0: sliceColor = RGB(46, 230, 166)
            Case InStr(label, "inactive") > 0: sliceColor = RGB(120, 130, 160)
            Case InStr(label, "free") > 0: sliceColor = RGB(155, 89, 255)
            Case Else: sliceColor = PalColor(i)
        End Select
        ch.SeriesCollection(1).Points(i).Format.Fill.ForeColor.RGB = sliceColor
        ch.SeriesCollection(1).Points(i).Format.Line.ForeColor.RGB = RGB(30, 40, 78)
    Next i

    ch.SeriesCollection(1).HasDataLabels = True
    With ch.SeriesCollection(1).DataLabels
        .ShowPercentage = True
        .ShowValue = False
        .ShowCategoryName = False
        .numberFormat = "0%"
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 8
        .Font.Bold = True
    End With

    StyleLegend ch, xlLegendPositionRight
    On Error GoTo 0
End Sub

' ---------- Updating the charts from the helper sheet ----------
' The helper sheet holds five blocks: A:D trend, F:H sales vs purchases,
' J:K donut, M:N bars, P:Q stock pie. Row 1 holds the headings.
Public Sub UpdateCharts(ByVal nTrend As Long, ByVal nDonut As Long, ByVal nBar As Long, _
                        ByVal nStock As Long, ByVal donutTitle As String, _
                        ByVal barTitle As String, ByVal stockTitle As String)
    Dim ws As Worksheet, wd As Worksheet
    Dim ch As Chart

    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Set wd = GetDataSheet()

    ' --- Revenue and gross profit over time ---
    Set ch = ws.ChartObjects("dashTrend").Chart
    ClearSeries ch
    AddSeries ch, "Revenue", wd.Range(wd.Cells(2, 2), wd.Cells(1 + nTrend, 2)), wd.Range(wd.Cells(2, 1), wd.Cells(1 + nTrend, 1))
    AddSeries ch, "Gross profit", wd.Range(wd.Cells(2, 3), wd.Cells(1 + nTrend, 3)), wd.Range(wd.Cells(2, 1), wd.Cells(1 + nTrend, 1))
    StyleCommon ch, "REVENUE AND GROSS PROFIT"
    StyleTrend ch

    ' --- Donut: category / product / customer share ---
    Set ch = ws.ChartObjects("dashCategory").Chart
    ClearSeries ch
    AddSeries ch, "Revenue", wd.Range(wd.Cells(2, 11), wd.Cells(1 + nDonut, 11)), wd.Range(wd.Cells(2, 10), wd.Cells(1 + nDonut, 10))
    StyleCommon ch, donutTitle
    StyleDonut ch, nDonut

    ' --- Bars: top products or weekdays ---
    Set ch = ws.ChartObjects("dashTop").Chart
    ClearSeries ch
    AddSeries ch, "Revenue", wd.Range(wd.Cells(2, 14), wd.Cells(1 + nBar, 14)), wd.Range(wd.Cells(2, 13), wd.Cells(1 + nBar, 13))
    StyleCommon ch, barTitle
    StyleTopBars ch

    ' --- Sales versus purchases ---
    Set ch = ws.ChartObjects("dashBuy").Chart
    ClearSeries ch
    AddSeries ch, "Sales", wd.Range(wd.Cells(2, 7), wd.Cells(1 + nTrend, 7)), wd.Range(wd.Cells(2, 6), wd.Cells(1 + nTrend, 6))
    AddSeries ch, "Purchases", wd.Range(wd.Cells(2, 8), wd.Cells(1 + nTrend, 8)), wd.Range(wd.Cells(2, 6), wd.Cells(1 + nTrend, 6))
    StyleCommon ch, "SALES VS PURCHASES"
    StyleBuy ch

    ' --- Stock pie ---
    Set ch = ws.ChartObjects("dashStock").Chart
    ClearSeries ch
    AddSeries ch, "Stock", wd.Range(wd.Cells(2, 17), wd.Cells(1 + nStock, 17)), wd.Range(wd.Cells(2, 16), wd.Cells(1 + nStock, 16))
    StyleCommon ch, stockTitle
    StylePie ch, wd, nStock
End Sub

