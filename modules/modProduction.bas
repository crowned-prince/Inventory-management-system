Attribute VB_Name = "modProduction"
Option Explicit
' ============================================================
' modProduction - one-time clean-up before shipping IMS to a customer.
' Run from the VBA editor on a COPY of your workbook:
'   click inside ResetForProduction and press F5.
'
' It removes ALL users and ALL business data, resets the ID counters and the
' personal settings, and saves. It makes a full backup first.
' The next time the workbook opens, the first-time administrator setup appears.
'
' Kept on purpose: the table structures, sheet visibility, categories (set
' CLEAR_CATEGORIES to True to remove them), currency, tax, invoice footer and
' the AutoSave / AutoBackup / HideSheets / AppMode settings.
' ============================================================
Private Const CLEAR_CATEGORIES As Boolean = False
Private Const CONFIRM_PHRASE As String = "CLEAR ALL DATA"

Public Sub ResetForProduction()
 Dim answer As String
 Dim backupPath As String
 Dim plan As Variant
 Dim parts() As String
 Dim i As Long
 Dim report As String
 Dim leftover As Long
 Dim problems As String
 Dim n As Long
 On Error GoTo ErrorHandler
 plan = TablePlan()
 modSecurity.EnterAppMode ' unprotects sheets; hides/shows them as your settings say
 ' ---- Confirmation: show exactly what will be deleted ----
 For i = LBound(plan) To UBound(plan)
  parts = Split(CStr(plan(i)), "|")
  n = RowCount(parts(0), parts(1))
  If n < 0 Then
   MsgBox "The table " & parts(1) & " on sheet " & parts(0) & " was not found." & vbCrLf & _
    "Nothing was changed.", vbExclamation, "Reset for production"
   Exit Sub
  End If
  report = report & "  " & parts(1) & ": " & n & " row(s)" & vbCrLf
 Next i
 answer = Trim$(InputBox("THIS PERMANENTLY DELETES ALL USERS AND ALL BUSINESS DATA" & vbCrLf & _
  "in this file:" & vbCrLf & ThisWorkbook.FullName & vbCrLf & vbCrLf & report & vbCrLf & _
  "A backup is made first. Only run this on a COPY meant for shipping." & vbCrLf & vbCrLf & _
  "Type " & CONFIRM_PHRASE & " to continue:", "Reset for production"))
 If answer <> CONFIRM_PHRASE Then
  MsgBox "Nothing was changed.", vbInformation, "Reset for production"
  Exit Sub
 End If
 ' ---- 1. Backup (stop if it fails) ----
 backupPath = modBackup.CreateBackup()
 If Len(Dir(backupPath)) = 0 Then
  MsgBox "The backup could not be confirmed, so nothing was deleted.", vbExclamation, "Reset for production"
  Exit Sub
 End If
 Application.ScreenUpdating = False
 ' ---- 2. Clear the tables ----
 For i = LBound(plan) To UBound(plan)
  parts = Split(CStr(plan(i)), "|")
  ClearTable parts(0), parts(1)
 Next i
 If CLEAR_CATEGORIES Then ClearTable "Categories", "tblCategories"
 ' ---- 3. Exported report output ----
 On Error Resume Next
 ThisWorkbook.Worksheets("Reports").UsedRange.Clear
 On Error GoTo ErrorHandler
 ' ---- 4. Settings: counters, personal values, set-up markers ----
 ResetSettings
 ' ---- 5. Refresh the dashboard cache so no old figures remain ----
 On Error Resume Next
 modDash.AfterSession
 On Error GoTo ErrorHandler
 ' ---- 6. Verify ----
 For i = LBound(plan) To UBound(plan)
  parts = Split(CStr(plan(i)), "|")
  leftover = RowCount(parts(0), parts(1))
  If leftover <> 0 Then problems = problems & parts(1) & " still has " & leftover & " row(s)." & vbCrLf
 Next i
 Application.ScreenUpdating = True
 If Len(problems) > 0 Then
  MsgBox "The clean-up did not finish cleanly:" & vbCrLf & problems & vbCrLf & _
   "The file was NOT saved. Your backup is at:" & vbCrLf & backupPath, vbExclamation, "Reset for production"
  Exit Sub
 End If
 If Not ThisWorkbook.ReadOnly Then ThisWorkbook.Save
 MsgBox "Done. All users and data were removed and the file was saved." & vbCrLf & vbCrLf & _
  "NEXT STEPS" & vbCrLf & _
  "1. Close Excel now. Do not open the program on this copy: the first-time setup " & _
  "belongs to your customer." & vbCrLf & _
  "2. The backup you just made still holds ALL the old data:" & vbCrLf & backupPath & vbCrLf & _
  "   Keep it private. Do not ship it, and do not ship any older Backups folder." & vbCrLf & _
  "3. Optional: File > Info > Inspect Document, to remove author names.", _
  vbInformation, "Reset for production"
 Exit Sub
ErrorHandler:
 Application.ScreenUpdating = True
 MsgBox "The clean-up stopped part-way (error " & Err.Number & ")." & vbCrLf & _
  "The file was NOT saved. Close it WITHOUT saving, and use your backup if needed." & _
  IIf(Len(backupPath) > 0, vbCrLf & backupPath, ""), vbExclamation, "Reset for production"
End Sub

' Sheet|Table of everything that is removed (Categories is handled separately).
Private Function TablePlan() As Variant
 TablePlan = Array("Users|tblUsers", "Products|tblProducts", "Suppliers|tblSuppliers", _
  "Customers|tblCustomers", "Sales|tblSales", "SaleDetails|tblSaleDetails", _
  "Purchases|tblPurchases", "PurchaseDetails|tblPurchaseDetails", _
  "Inventory|tblStockMovements", "Expenses|tblExpenses", "AuditLog|tblAuditLog")
End Function

' Number of non-blank rows, or -1 if the table cannot be found.
Private Function RowCount(ByVal sheetName As String, ByVal tableName As String) As Long
 Dim tbl As ListObject
 Dim i As Long
 On Error GoTo NotFound
 Set tbl = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
 On Error GoTo 0
 If tbl.DataBodyRange Is Nothing Then Exit Function
 For i = 1 To tbl.ListRows.count
  If Application.WorksheetFunction.CountA(tbl.ListRows(i).Range) > 0 Then RowCount = RowCount + 1
 Next i
 Exit Function
NotFound:
 RowCount = -1
End Function

' Empties a table but keeps its header, formatting and ONE blank row
' (the rest of IMS expects a table to keep at least one row).
Private Sub ClearTable(ByVal sheetName As String, ByVal tableName As String)
 Dim tbl As ListObject
 Dim n As Long
 Set tbl = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
 If tbl.DataBodyRange Is Nothing Then
  tbl.ListRows.Add
  Exit Sub
 End If
 n = tbl.ListRows.count
 If n > 1 Then tbl.DataBodyRange.Offset(1, 0).Resize(n - 1).Delete xlShiftUp
 tbl.ListRows(1).Range.ClearContents
End Sub

Private Sub ResetSettings()
 Dim counters As Variant
 Dim item As Variant
 Dim tbl As ListObject
 Dim rowNo As Long
 counters = Array("LastProductNo", "LastCustomerNo", "LastSupplierNo", "LastSaleNo", _
  "LastPurchaseNo", "LastInvoiceNo", "LastUserNo", "LastExpenseNo", "LastSaleDetailNo", _
  "LastMovementNo", "LastPurchaseDetailNo")
 For Each item In counters
  If SettingExists(CStr(item)) Then SetSetting CStr(item), "0"
 Next item
 If CLEAR_CATEGORIES Then SetSetting "LastCategoryNo", "0"
 ' Personal / customer-specific values go back to the shipping defaults
 SetSetting "ShopName", "My Retail Shop"
 SetSetting "ShopAddress", "Main Street, Your City"
 SetSetting "ShopPhone", "0300-0000000"
 SetSetting "RememberedUser", ""
 SetSetting "LastBackupDate", ""
 SetSetting "BackupFolder", ""
 SetSetting "SystemVersion", modSetupAuth.APP_VERSION
 ' Remove any set-up markers so the first-time setup appears
 Set tbl = modDatabase.GetTable("Settings", "tblSettings")
 For Each item In Array("SetupCompleted", "SetupCompletedOn", "SetupRecoveryAuthorized")
  rowNo = modDatabase.FindRowByValue(tbl, "SettingKey", CStr(item))
  If rowNo > 0 Then
   If tbl.ListRows.count > 1 Then tbl.ListRows(rowNo).Delete
  End If
 Next item
End Sub

Private Function SettingExists(ByVal settingKey As String) As Boolean
 SettingExists = (GetSetting(settingKey, "#MISSING#") <> "#MISSING#")
End Function
