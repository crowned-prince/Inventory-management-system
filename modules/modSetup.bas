Attribute VB_Name = "modSetup"
Option Explicit

' Creates all database sheets for the Inventory System.
' Safe to run more than once: it skips sheets that already exist.
Public Sub CreateAllSheets()
    Dim sheetNames As Variant
    Dim i As Long
    Dim ws As Worksheet

    On Error GoTo ErrorHandler

    sheetNames = Array("Dashboard", "Users", "Products", "Categories", _
                       "Suppliers", "Customers", "Sales", "SaleDetails", _
                       "Purchases", "PurchaseDetails", "Inventory", _
                       "Expenses", "Settings", "AuditLog", "Reports")

    For i = LBound(sheetNames) To UBound(sheetNames)
        If Not SheetExists(CStr(sheetNames(i))) Then
            Set ws = ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
            ws.name = sheetNames(i)
        End If
    Next i

    ThisWorkbook.Worksheets("Dashboard").Activate
    MsgBox "All sheets are ready.", vbInformation, "Setup Complete"
    Exit Sub

ErrorHandler:
    MsgBox "Could not create the sheets." & vbCrLf & _
           "Details: " & Err.description, vbExclamation, "Setup Error"
End Sub

' Returns True if a worksheet with this name already exists.
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

