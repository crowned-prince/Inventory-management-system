Attribute VB_Name = "modInitialData"
Option Explicit

' Fills Settings and Categories with starting data.
' Safe to run more than once: it will not create duplicates.
Public Sub LoadInitialData()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    LoadDefaultSettings
    LoadDefaultCategories

    ThisWorkbook.Worksheets("Settings").Columns.AutoFit
    ThisWorkbook.Worksheets("Categories").Columns.AutoFit
    Application.ScreenUpdating = True

    MsgBox "Settings and categories are ready.", vbInformation, "Setup Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Could not load the initial data." & vbCrLf & Err.description, _
           vbExclamation, "Setup Error"
End Sub

Private Sub LoadDefaultSettings()
    ' Shop information
    AddSettingIfMissing "ShopName", "My Retail Shop"
    AddSettingIfMissing "ShopAddress", "Main Street, Your City"
    AddSettingIfMissing "ShopPhone", "0300-0000000"
    AddSettingIfMissing "CurrencySymbol", "Rs."
    AddSettingIfMissing "TaxEnabled", "No"
    AddSettingIfMissing "TaxRate", "0"
    AddSettingIfMissing "SystemVersion", "1.0"

    ' ID counters (all start at 0)
    AddSettingIfMissing "LastProductNo", "0"
    AddSettingIfMissing "LastCustomerNo", "0"
    AddSettingIfMissing "LastSupplierNo", "0"
    AddSettingIfMissing "LastSaleNo", "0"
    AddSettingIfMissing "LastPurchaseNo", "0"
    AddSettingIfMissing "LastInvoiceNo", "0"
    AddSettingIfMissing "LastUserNo", "0"
    AddSettingIfMissing "LastCategoryNo", "0"
    AddSettingIfMissing "LastExpenseNo", "0"
End Sub

Private Sub LoadDefaultCategories()
    AddCategoryIfMissing "Grocery", "Food and daily-use items"
    AddCategoryIfMissing "Electronics", "Devices and accessories"
    AddCategoryIfMissing "Clothing", "Garments and footwear"
    AddCategoryIfMissing "Cosmetics", "Beauty and personal care"
    AddCategoryIfMissing "Household", "Home and cleaning items"
    AddCategoryIfMissing "Stationery", "Office and school supplies"
    AddCategoryIfMissing "Beverages", "Drinks and juices"
End Sub

' Adds a setting only if the key does not already exist.
Private Sub AddSettingIfMissing(ByVal settingKey As String, ByVal settingValue As String)
    If GetSetting(settingKey, "#MISSING#") = "#MISSING#" Then
        SetSetting settingKey, settingValue
    End If
End Sub

' Adds a category only if the name does not already exist.
Private Sub AddCategoryIfMissing(ByVal categoryName As String, ByVal description As String)
    Dim tbl As ListObject
    Dim newRow As ListRow

    Set tbl = ThisWorkbook.Worksheets("Categories").ListObjects("tblCategories")

    If Not tbl.DataBodyRange Is Nothing Then
        If Not IsError(Application.Match(categoryName, _
                tbl.ListColumns("CategoryName").DataBodyRange, 0)) Then Exit Sub
    End If

    Set newRow = tbl.ListRows.Add
    newRow.Range.Cells(1, 1).value = GenerateID("CAT")
    newRow.Range.Cells(1, 2).value = categoryName
    newRow.Range.Cells(1, 3).value = description
End Sub
' Deletes completely empty rows from every table (keeps at least one row).
Public Sub CleanBlankTableRows()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long

    On Error GoTo ErrorHandler

    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            For i = tbl.ListRows.count To 1 Step -1
                If tbl.ListRows.count > 1 Then
                    If Application.WorksheetFunction.CountA(tbl.ListRows(i).Range) = 0 Then
                        tbl.ListRows(i).Delete
                    End If
                End If
            Next i
        Next tbl
    Next ws

    MsgBox "Blank table rows removed.", vbInformation, "Cleanup"
    Exit Sub

ErrorHandler:
    MsgBox "Cleanup failed: " & Err.description, vbExclamation, "Cleanup Error"
End Sub

' Returns the row to write a new record into.
' Uses the blank first row if the table is brand new, otherwise adds a row.
Public Function NextTableRow(ByVal tbl As ListObject) As ListRow
    If tbl.ListRows.count = 1 Then
        If Application.WorksheetFunction.CountA(tbl.ListRows(1).Range) = 0 Then
            Set NextTableRow = tbl.ListRows(1)
            Exit Function
        End If
    End If
    Set NextTableRow = tbl.ListRows.Add
End Function

' Writes one line to the AuditLog table.
Public Sub LogAction(ByVal actionName As String, _
                     ByVal recordID As String, _
                     ByVal description As String)
    Dim tbl As ListObject
    Dim newRow As ListRow

    On Error Resume Next   ' logging must never crash the program
    Set tbl = ThisWorkbook.Worksheets("AuditLog").ListObjects("tblAuditLog")
    Set newRow = NextTableRow(tbl)
    newRow.Range.Cells(1, 1).value = Date
    newRow.Range.Cells(1, 2).value = Format$(Time, "hh:mm")
    newRow.Range.Cells(1, 3).value = modMain.CurrentUsername
    newRow.Range.Cells(1, 4).value = actionName
    newRow.Range.Cells(1, 5).value = recordID
    newRow.Range.Cells(1, 6).value = description
    On Error GoTo 0
End Sub
