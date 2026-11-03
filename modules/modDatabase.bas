Attribute VB_Name = "modDatabase"
Option Explicit

' ============================================================
' modDatabase - generic helpers that work on any table
' ============================================================

Public Function GetTable(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Set GetTable = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
End Function

' Returns the row number (1 = first data row) where the column equals the value.
' Returns 0 if not found or the value is blank.
Public Function FindRowByValue(ByVal tbl As ListObject, ByVal columnName As String, _
                               ByVal searchValue As String) As Long
    Dim data As Variant
    Dim colNo As Long, i As Long

    If Len(searchValue) = 0 Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    data = tbl.DataBodyRange.value
    colNo = tbl.ListColumns(columnName).index

    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, colNo)), searchValue, vbTextCompare) = 0 Then
            FindRowByValue = i
            Exit Function
        End If
    Next i
End Function

' How many rows have this exact value in the column.
Public Function CountMatches(ByVal sheetName As String, ByVal tableName As String, _
                             ByVal columnName As String, ByVal searchValue As String) As Long
    Dim tbl As ListObject
    Dim data As Variant
    Dim colNo As Long, i As Long

    If Len(searchValue) = 0 Then Exit Function
    Set tbl = GetTable(sheetName, tableName)
    If tbl.DataBodyRange Is Nothing Then Exit Function

    data = tbl.DataBodyRange.value
    colNo = tbl.ListColumns(columnName).index

    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, colNo)), searchValue, vbTextCompare) = 0 Then
            CountMatches = CountMatches + 1
        End If
    Next i
End Function

Public Function ColumnContains(ByVal sheetName As String, ByVal tableName As String, _
                               ByVal columnName As String, ByVal searchValue As String) As Boolean
    ColumnContains = (CountMatches(sheetName, tableName, columnName, searchValue) > 0)
End Function

' Reads one cell using the column NAME (safe if columns are ever reordered).
Public Function GetCellValue(ByVal tbl As ListObject, ByVal rowNo As Long, _
                             ByVal columnName As String) As Variant
    GetCellValue = tbl.DataBodyRange.Cells(rowNo, tbl.ListColumns(columnName).index).value
End Function

Public Sub SetCellValue(ByVal tbl As ListObject, ByVal rowNo As Long, _
                        ByVal columnName As String, ByVal newValue As Variant)
    tbl.DataBodyRange.Cells(rowNo, tbl.ListColumns(columnName).index).value = newValue
End Sub
' Adds up one column for all rows where another column equals the value.
' Example: total of GrandTotal in Sales where CustomerID = "CUS-00001".
Public Function SumMatches(ByVal sheetName As String, ByVal tableName As String, _
                           ByVal matchColumn As String, ByVal searchValue As String, _
                           ByVal sumColumn As String) As Double
    Dim tbl As ListObject
    Dim data As Variant
    Dim matchCol As Long, sumCol As Long, i As Long

    If Len(searchValue) = 0 Then Exit Function
    Set tbl = GetTable(sheetName, tableName)
    If tbl.DataBodyRange Is Nothing Then Exit Function

    data = tbl.DataBodyRange.value
    matchCol = tbl.ListColumns(matchColumn).index
    sumCol = tbl.ListColumns(sumColumn).index

    For i = 1 To UBound(data, 1)
        If StrComp(CStr(data(i, matchCol)), searchValue, vbTextCompare) = 0 Then
            SumMatches = SumMatches + ToNumber(data(i, sumCol))
        End If
    Next i
End Function
