Attribute VB_Name = "modSampleData"
Option Explicit

' ============================================================
' modSampleData - realistic demo data, and a safe "clear everything" tool.
' Uses Scripting.Dictionary (Windows Excel).
' Run these from the VBA editor: click inside a macro and press F5.
' ============================================================

Private Const PRODUCT_COUNT As Long = 26
Private Const SUPPLIER_COUNT As Long = 5
Private Const CUSTOMER_COUNT As Long = 8

' What the loader remembers while it works
Private mProductIDs(0 To 25) As String
Private mPrices(0 To 25) As Double
Private mCosts(0 To 25) As Double
Private mSupplierNo(0 To 25) As Long
Private mBuyQty(0 To 25) As Double
Private mSupplierIDs(0 To 4) As String
Private mCustomerIDs(0 To 7) As String
Private mDateMap As Object           ' sale / purchase ID -> the date it should have

' ============================================================
' Loading the sample data
' ============================================================
Public Sub LoadSampleData()
    Dim oldUserID As String, oldUsername As String
    Dim dayOffset As Long, salesToday As Long, s As Long
    Dim report As String, issues As Long
    Dim pr As modProfit.TProfit
    Dim dummy As Single

    ' ---- Safety: only load into an empty system ----
    If modDashboard.CountRecords("Products", "tblProducts") > 0 Or _
       modDashboard.CountRecords("Sales", "tblSales") > 0 Then
        MsgBox "This workbook already contains products or sales." & vbCrLf & vbCrLf & _
               "Sample data can only be loaded into an empty system." & vbCrLf & _
               "To start clean, run ClearDemoData first (it makes a backup).", _
               vbExclamation, "Sample data"
        Exit Sub
    End If

    If MsgBox("This will add sample suppliers, customers, products, purchases, sales and expenses." & _
              vbCrLf & "It takes about half a minute." & vbCrLf & vbCrLf & "Continue?", _
              vbYesNo + vbQuestion, "Load sample data") <> vbYes Then Exit Sub

    ' Records need a user: borrow the first admin if nobody is logged in
    oldUserID = modMain.CurrentUserID
    oldUsername = modMain.CurrentUsername
    If Len(oldUserID) = 0 Then
        modMain.CurrentUserID = "USR-00001"
        modMain.CurrentUsername = "admin"
    End If

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    modSecurity.EnterAppMode                       ' sheets must be unprotected

    Set mDateMap = CreateObject("Scripting.Dictionary")
    dummy = Rnd(-1)                                ' same random numbers every time
    Randomize 42

    PrepareSampleFormats
    EnsureWalkIn
    CreateSuppliers
    CreateCustomers
    CreateProducts
    InitialPurchases

    ' 45 days of trading, oldest day first
    For dayOffset = 44 To 0 Step -1
        Application.StatusBar = "Loading sample data... day " & (45 - dayOffset) & " of 45"
        If dayOffset = 22 Then RestockPurchases
        salesToday = 2 + Int(Rnd * 4)              ' 2 to 5 sales a day
        For s = 1 To salesToday
            MakeRandomSale Date - dayOffset
        Next s
    Next dayOffset

    FinishStockStates
    CreateExpenses
    ApplyDates

    ' ---- Check the result ----
    issues = modInventory.CheckStockLog(False, report)
    pr = modProfit.CalculateProfit(Date - 29, Date)
    LogAction "SAMPLE DATA", "-", "Sample data loaded"

    Application.StatusBar = False
    Application.ScreenUpdating = True
    modMain.CurrentUserID = oldUserID
    modMain.CurrentUsername = oldUsername

    MsgBox "Sample data is ready." & vbCrLf & vbCrLf & _
           "Products:  " & modDashboard.CountRecords("Products", "tblProducts") & vbCrLf & _
           "Suppliers:  " & modDashboard.CountRecords("Suppliers", "tblSuppliers") & vbCrLf & _
           "Customers:  " & modCustomers.RealCustomerCount() & vbCrLf & _
           "Purchases:  " & modDashboard.CountRecords("Purchases", "tblPurchases") & vbCrLf & _
           "Sales:  " & modDashboard.CountRecords("Sales", "tblSales") & vbCrLf & _
           "Expenses:  " & modDashboard.CountRecords("Expenses", "tblExpenses") & vbCrLf & vbCrLf & _
           "Stock check:  " & IIf(issues = 0, "all stock figures match the movement log", _
                                  issues & " product(s) do not match") & vbCrLf & _
           "Last 30 days:  revenue " & Format$(pr.revenue, "#,##0") & _
           ",  net profit " & Format$(pr.netProfit, "#,##0"), _
           vbInformation, "Sample data"
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    modMain.CurrentUserID = oldUserID
    modMain.CurrentUsername = oldUsername
    MsgBox "Loading stopped: " & Err.description & vbCrLf & vbCrLf & _
           "Run ClearDemoData to remove what was added, then try again.", _
           vbExclamation, "Sample data"
End Sub

' ---------- Setting up ----------

' Text formats so barcodes and phone numbers keep their leading zeros (cosmetic, errors ignored).
Private Sub PrepareSampleFormats()
    Dim tbl As ListObject

    On Error Resume Next
    Set tbl = modDatabase.GetTable("Products", "tblProducts")
    tbl.ListColumns("Barcode").DataBodyRange.numberFormat = "@"
    Set tbl = modDatabase.GetTable("Suppliers", "tblSuppliers")
    tbl.ListColumns("Phone").DataBodyRange.numberFormat = "@"
    Set tbl = modDatabase.GetTable("Customers", "tblCustomers")
    tbl.ListColumns("Phone").DataBodyRange.numberFormat = "@"
    On Error GoTo 0
End Sub

' Makes sure the built-in Walk-in Customer exists (silently).
Private Sub EnsureWalkIn()
    Dim tbl As ListObject
    Dim newRow As ListRow

    If modCustomers.FindCustomerRow(modCustomers.WALK_IN_ID) > 0 Then Exit Sub

    Set tbl = modDatabase.GetTable("Customers", "tblCustomers")
    Set newRow = NextTableRow(tbl)
    ' Order matches tblCustomers (7 columns)
    newRow.Range.value = Array(modCustomers.WALK_IN_ID, "Walk-in Customer", "", "", "", 0, Date)
End Sub

' ---------- Suppliers, customers, products ----------
Private Sub CreateSuppliers()
    Dim defs(0 To 4) As String
    Dim parts() As String
    Dim s As modSuppliers.TSupplier
    Dim i As Long

    ' name | contact | phone | email | address
    defs(0) = "Metro Wholesale Foods|Bilal Ahmed|0300-1112233|bilal@metrofoods.example|Wholesale Market, Block A"
    defs(1) = "Techzone Distributors|Hamza Khan|0301-2223344|hamza@techzone.example|Electronics Plaza, Floor 2"
    defs(2) = "Glow Cosmetics Ltd|Ayesha Malik|0302-3334455|ayesha@glow.example|Beauty Street, Shop 14"
    defs(3) = "Fashion Hub Traders|Usman Raza|0303-4445566|usman@fashionhub.example|Garment Market, Lane 3"
    defs(4) = "HomeCare Supplies|Sana Iqbal|0304-5556677|sana@homecare.example|Industrial Area, Unit 7"

    For i = 0 To 4
        parts = Split(defs(i), "|")
        s.supplierName = parts(0)
        s.ContactPerson = parts(1)
        s.phone = parts(2)
        s.Email = parts(3)
        s.Address = parts(4)
        mSupplierIDs(i) = modSuppliers.AddSupplier(s)
    Next i
End Sub

Private Sub CreateCustomers()
    Dim defs(0 To 7) As String
    Dim parts() As String
    Dim c As modCustomers.TCustomer
    Dim i As Long

    ' name | phone | email | address
    defs(0) = "Ahmed Khan|0311-1000001|ahmed.khan@mail.example|House 12, Street 4"
    defs(1) = "Fatima Noor|0312-1000002|fatima.noor@mail.example|Flat 3, Green Towers"
    defs(2) = "Hassan Ali|0313-1000003|hassan.ali@mail.example|House 45, Model Town"
    defs(3) = "Zainab Shah|0314-1000004|zainab.shah@mail.example|Street 9, Garden Road"
    defs(4) = "Imran Qureshi|0315-1000005|imran.q@mail.example|Plot 21, Sector B"
    defs(5) = "Maryam Siddiqui|0316-1000006|maryam.s@mail.example|House 8, Park Lane"
    defs(6) = "Bilal Hussain|0317-1000007|bilal.h@mail.example|Block C, City Heights"
    defs(7) = "Sadia Rauf|0318-1000008|sadia.rauf@mail.example|House 30, Canal View"

    For i = 0 To 7
        parts = Split(defs(i), "|")
        c.customerName = parts(0)
        c.phone = parts(1)
        c.Email = parts(2)
        c.Address = parts(3)
        mCustomerIDs(i) = modCustomers.AddCustomer(c)
    Next i
End Sub

Private Sub CreateProducts()
    Dim defs(0 To 25) As String
    Dim parts() As String
    Dim p As modProducts.TProduct
    Dim i As Long

    ' name | category | brand | supplier no | cost | price | first purchase qty | min | max | unit
    defs(0) = "Basmati Rice 5kg|Grocery|Kisan|1|1450|1650|60|15|120|Bag"
    defs(1) = "Cooking Oil 3L|Grocery|Dalda|1|1380|1550|48|12|100|Bottle"
    defs(2) = "Sugar 1kg|Grocery|Local|1|135|155|120|30|300|Pack"
    defs(3) = "Tea Leaves 440g|Grocery|Tapal|1|520|600|50|12|100|Pack"
    defs(4) = "Wheat Flour 10kg|Grocery|Sunridge|1|1180|1320|40|10|80|Bag"
    defs(5) = "Red Lentils 1kg|Grocery|Local|1|280|330|45|10|90|Pack"
    defs(6) = "Orange Juice 1L|Beverages|Nestle|1|210|260|40|10|80|Bottle"
    defs(7) = "Mineral Water 1.5L|Beverages|Nestle|1|70|90|150|40|300|Bottle"
    defs(8) = "USB Cable 1m|Electronics|Anker|2|280|450|40|8|80|Pcs"
    defs(9) = "Wireless Mouse|Electronics|Logitech|2|1250|1650|15|5|40|Pcs"
    defs(10) = "Earphones|Electronics|Samsung|2|650|950|25|8|60|Pcs"
    defs(11) = "Power Bank 10000mAh|Electronics|Xiaomi|2|2600|3300|12|4|30|Pcs"
    defs(12) = "LED Bulb 12W|Electronics|Philips|2|220|320|60|15|120|Pcs"
    defs(13) = "Shampoo 200ml|Cosmetics|Sunsilk|3|330|420|35|10|80|Bottle"
    defs(14) = "Face Wash 100ml|Cosmetics|Garnier|3|410|520|30|8|60|Bottle"
    defs(15) = "Body Lotion 250ml|Cosmetics|Nivea|3|560|690|25|6|50|Bottle"
    defs(16) = "Toothpaste 150g|Cosmetics|Colgate|3|190|240|70|15|140|Pcs"
    defs(17) = "Cotton T-Shirt|Clothing|Gul Ahmed|4|650|990|30|8|60|Pcs"
    defs(18) = "Denim Jeans|Clothing|Outfitters|4|1800|2600|20|5|40|Pcs"
    defs(19) = "Socks (3 pairs)|Clothing|Bata|4|260|390|45|10|90|Pack"
    defs(20) = "Dishwashing Liquid 750ml|Household|Vim|5|250|320|40|10|80|Bottle"
    defs(21) = "Laundry Detergent 2kg|Household|Surf Excel|5|960|1150|30|8|60|Pack"
    defs(22) = "Floor Cleaner 1L|Household|Harpic|5|340|430|28|6|56|Bottle"
    defs(23) = "Kitchen Tissue Roll|Household|Local|5|120|170|60|15|120|Pcs"
    defs(24) = "Notebook A4|Stationery|Local|5|90|130|50|15|100|Pcs"
    defs(25) = "Ball Pen Pack (10)|Stationery|Dollar|5|110|160|40|10|80|Pack"

    For i = 0 To 25
        parts = Split(defs(i), "|")

        p.barcode = "89640" & Format$(i + 1, "00000000")
        p.productName = parts(0)
        p.category = parts(1)
        p.brand = parts(2)

        mSupplierNo(i) = CLng(parts(3))
        p.supplierID = mSupplierIDs(mSupplierNo(i) - 1)

        mCosts(i) = CDbl(parts(4))
        p.PurchasePrice = mCosts(i)
        mPrices(i) = CDbl(parts(5))
        p.sellingPrice = mPrices(i)
        mBuyQty(i) = CDbl(parts(6))

        p.stock = 0                                ' the stock arrives through the first purchase
        p.minStock = CDbl(parts(7))
        p.MaxStock = CDbl(parts(8))
        p.Unit = parts(9)
        p.ExpiryDate = ""
        p.description = "Sample product"
        p.status = "Active"

        mProductIDs(i) = modProducts.AddProduct(p)
    Next i
End Sub

' ---------- Purchases ----------

' Saves one purchase from a supplier. idx() holds product numbers (0-25), qty() the quantities.
Private Sub PostPurchase(ByVal supplierNo As Long, ByVal daysAgo As Long, _
                         ByRef idx() As Long, ByRef qty() As Double, ByVal n As Long)
    Dim ids() As String, quantities() As Double, costs() As Double
    Dim i As Long
    Dim purchaseDate As Date
    Dim purchaseID As String

    ReDim ids(1 To n)
    ReDim quantities(1 To n)
    ReDim costs(1 To n)

    For i = 1 To n
        ids(i) = mProductIDs(idx(i))
        quantities(i) = qty(i)
        costs(i) = mCosts(idx(i))
    Next i

    purchaseDate = Date - daysAgo
    purchaseID = modPurchases.CompletePurchase(mSupplierIDs(supplierNo - 1), purchaseDate, ids, quantities, costs)
    mDateMap(purchaseID) = purchaseDate + TimeSerial(9, 30, 0)
End Sub

' The first purchase from each supplier: everything it supplies, 45 days ago.
Private Sub InitialPurchases()
    Dim s As Long, i As Long, n As Long
    Dim idx() As Long, qty() As Double

    For s = 1 To SUPPLIER_COUNT
        n = 0
        ReDim idx(1 To PRODUCT_COUNT)
        ReDim qty(1 To PRODUCT_COUNT)

        For i = 0 To PRODUCT_COUNT - 1
            If mSupplierNo(i) = s Then
                n = n + 1
                idx(n) = i
                qty(n) = mBuyQty(i)
            End If
        Next i

        If n > 0 Then PostPurchase s, 45, idx, qty, n
    Next s
End Sub

' A restock of the fast sellers, 22 days ago.
Private Sub RestockPurchases()
    Dim idx() As Long, qty() As Double

    ReDim idx(1 To 5)
    ReDim qty(1 To 5)
    idx(1) = 0: qty(1) = 40        ' Basmati Rice
    idx(2) = 1: qty(2) = 40        ' Cooking Oil
    idx(3) = 2: qty(3) = 60        ' Sugar
    idx(4) = 3: qty(4) = 30        ' Tea
    idx(5) = 7: qty(5) = 100       ' Mineral Water
    PostPurchase 1, 22, idx, qty, 5

    ReDim idx(1 To 2)
    ReDim qty(1 To 2)
    idx(1) = 13: qty(1) = 30       ' Shampoo
    idx(2) = 16: qty(2) = 50       ' Toothpaste
    PostPurchase 3, 22, idx, qty, 2
End Sub

' ---------- Sales ----------

' One sale with 1 to 4 random lines, on the given day.
Private Sub MakeRandomSale(ByVal saleDay As Date)
    Dim lineCount As Long, k As Long, n As Long
    Dim idx As Long
    Dim qty As Double, available As Double, subtotal As Double, invDisc As Double
    Dim ids() As String, qtys() As Double, prices() As Double, discs() As Double
    Dim inCart As Object
    Dim customerID As String
    Dim saleDate As Date
    Dim saleID As String, invoiceNo As String
    Dim grandTotal As Double

    lineCount = 1 + Int(Rnd * 4)                    ' 1 to 4 lines
    ReDim ids(1 To lineCount)
    ReDim qtys(1 To lineCount)
    ReDim prices(1 To lineCount)
    ReDim discs(1 To lineCount)

    Set inCart = CreateObject("Scripting.Dictionary")

    For k = 1 To lineCount
        idx = Int(Rnd * PRODUCT_COUNT)
        qty = 1 + Int(Rnd * 3)                      ' 1 to 3 units

        ' Never sell more than is in stock (counting what is already in this cart)
        available = modInventory.CurrentStock(mProductIDs(idx)) - inCart(idx)
        If available >= qty Then
            n = n + 1
            ids(n) = mProductIDs(idx)
            qtys(n) = qty
            prices(n) = mPrices(idx)
            If Rnd < 0.08 Then
                discs(n) = 10 * (1 + Int(Rnd * 5))  ' 10 to 50 off this line
            Else
                discs(n) = 0
            End If
            inCart(idx) = inCart(idx) + qty
            subtotal = subtotal + qty * mPrices(idx) - discs(n)
        End If
    Next k

    If n = 0 Then Exit Sub                          ' nothing was available this time

    If n < lineCount Then
        ReDim Preserve ids(1 To n)
        ReDim Preserve qtys(1 To n)
        ReDim Preserve prices(1 To n)
        ReDim Preserve discs(1 To n)
    End If

    invDisc = 0
    If Rnd < 0.1 Then invDisc = 20 * (1 + Int(Rnd * 5))     ' 20 to 100 off the whole invoice
    If invDisc > subtotal Then invDisc = 0

    If Rnd < 0.6 Then
        customerID = modCustomers.WALK_IN_ID
    Else
        customerID = mCustomerIDs(Int(Rnd * CUSTOMER_COUNT))
    End If

    saleDate = saleDay + TimeSerial(9 + Int(Rnd * 11), Int(Rnd * 60), 0)
    If saleDate > Now Then saleDate = Now

    saleID = modSales.CompleteSale(customerID, ids, qtys, prices, discs, invDisc, invoiceNo, grandTotal)
    mDateMap(saleID) = saleDate
End Sub

' ---------- Finishing touches ----------

' Gives the demo some alert examples: one product out of stock, one low, one inactive.
Private Sub FinishStockStates()
    Dim tbl As ListObject
    Dim rowNo As Long

    modInventory.AdjustStockToCount mProductIDs(10), 0, "Counting correction", "Sample data: out-of-stock example"
    modInventory.AdjustStockToCount mProductIDs(14), 4, "Counting correction", "Sample data: low-stock example"

    Set tbl = modDatabase.GetTable("Products", "tblProducts")
    rowNo = modProducts.FindProductRow(mProductIDs(23))
    If rowNo > 0 Then modDatabase.SetCellValue tbl, rowNo, "Status", "Inactive"
End Sub

Private Sub CreateExpenses()
    Dim m As Long
    Dim firstDay As Date

    For m = 0 To 1                                  ' this month and last month
        firstDay = DateSerial(Year(Date), Month(Date) - m, 1)
        AddSampleExpense firstDay, "Rent", "Shop rent", 45000
        AddSampleExpense firstDay + 2, "Electricity", "Electricity bill", 12500
        AddSampleExpense firstDay + 4, "Internet / phone", "Internet and phone", 3500
        AddSampleExpense firstDay + 9, "Salaries", "Staff salaries", 85000
        AddSampleExpense firstDay + 14, "Transport", "Delivery and transport", 6000
    Next m
End Sub

Private Sub AddSampleExpense(ByVal expenseDay As Date, ByVal category As String, _
                             ByVal description As String, ByVal amount As Double)
    Dim p As modExpenses.TExpense

    If expenseDay > Date Then Exit Sub              ' expenses cannot be in the future
    If expenseDay < Date - 44 Then Exit Sub         ' nothing before the first sale

    p.ExpenseDate = Format$(expenseDay, "dd-mm-yyyy")
    p.category = category
    p.description = description
    p.amount = amount
    modExpenses.AddExpense p
End Sub

' Sales and stock movements are created "now" - this gives them the dates they should have.
Private Sub ApplyDates()
    Dim tbl As ListObject
    Dim data As Variant
    Dim values() As Variant
    Dim colID As Long, colDate As Long
    Dim i As Long, n As Long

    ' --- Sales: SaleID -> date ---
    Set tbl = modDatabase.GetTable("Sales", "tblSales")
    If Not tbl.DataBodyRange Is Nothing Then
        data = tbl.DataBodyRange.value
        colID = tbl.ListColumns("SaleID").index
        colDate = tbl.ListColumns("SaleDate").index
        n = UBound(data, 1)
        ReDim values(1 To n, 1 To 1)

        For i = 1 To n
            If mDateMap.Exists(CStr(data(i, colID))) Then
                values(i, 1) = mDateMap(CStr(data(i, colID)))
            Else
                values(i, 1) = data(i, colDate)
            End If
        Next i
        tbl.ListColumns("SaleDate").DataBodyRange.value = values
    End If

    ' --- Stock movements: ReferenceID (a sale or purchase) -> date ---
    Set tbl = modDatabase.GetTable("Inventory", "tblStockMovements")
    If Not tbl.DataBodyRange Is Nothing Then
        data = tbl.DataBodyRange.value
        colID = tbl.ListColumns("ReferenceID").index
        colDate = tbl.ListColumns("MovementDate").index
        n = UBound(data, 1)
        ReDim values(1 To n, 1 To 1)

        For i = 1 To n
            If mDateMap.Exists(CStr(data(i, colID))) Then
                values(i, 1) = mDateMap(CStr(data(i, colID)))
            Else
                values(i, 1) = data(i, colDate)
            End If
        Next i
        tbl.ListColumns("MovementDate").DataBodyRange.value = values
    End If
End Sub

' ============================================================
' Clearing everything (to start for real)
' ============================================================
Public Sub ClearDemoData()
    Dim answer As String
    Dim backupPath As String
    Dim key As Variant

    On Error GoTo ErrorHandler

    answer = Trim$(InputBox("This will DELETE all products, suppliers, customers, sales, purchases, " & _
                            "stock movements, expenses and the audit log." & vbCrLf & _
                            "Your users, categories and settings are kept." & vbCrLf & vbCrLf & _
                            "A backup is made first." & vbCrLf & vbCrLf & _
                            "Type DELETE (in capital letters) to continue:", _
                            "Clear all business data"))

    If answer <> "DELETE" Then
        MsgBox "Nothing was deleted.", vbInformation, "Clear all business data"
        Exit Sub
    End If

    modSecurity.EnterAppMode                        ' sheets must be unprotected
    backupPath = modBackup.CreateBackup()           ' safety copy first: if this fails, nothing is deleted

    Application.ScreenUpdating = False

    ClearTable "Products", "tblProducts"
    ClearTable "Suppliers", "tblSuppliers"
    ClearTable "Customers", "tblCustomers"
    ClearTable "Sales", "tblSales"
    ClearTable "SaleDetails", "tblSaleDetails"
    ClearTable "Purchases", "tblPurchases"
    ClearTable "PurchaseDetails", "tblPurchaseDetails"
    ClearTable "Inventory", "tblStockMovements"
    ClearTable "Expenses", "tblExpenses"
    ClearTable "AuditLog", "tblAuditLog"

    For Each key In Array("LastProductNo", "LastCustomerNo", "LastSupplierNo", "LastSaleNo", _
                          "LastPurchaseNo", "LastInvoiceNo", "LastMovementNo", _
                          "LastPurchaseDetailNo", "LastSaleDetailNo", "LastExpenseNo")
        SetSetting CStr(key), "0"
    Next key

    EnsureWalkIn
    LogActionAs "system", "CLEAR DATA", "-", "All business data cleared. Backup: " & _
                Mid$(backupPath, InStrRev(backupPath, "\") + 1)

    Application.ScreenUpdating = True
    MsgBox "All business data was cleared." & vbCrLf & vbCrLf & _
           "A backup was saved first:" & vbCrLf & backupPath, vbInformation, "Clear all business data"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Nothing more was deleted." & vbCrLf & Err.description, vbExclamation, "Clear all business data"
End Sub

Private Sub ClearTable(ByVal sheetName As String, ByVal tableName As String)
    Dim tbl As ListObject

    Set tbl = modDatabase.GetTable(sheetName, tableName)
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
End Sub

