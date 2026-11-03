Attribute VB_Name = "modAudit"
Option Explicit

' ============================================================
' modAudit - reading, filtering and exporting the audit log
' Uses Scripting.Dictionary (Windows Excel).
' ============================================================

' One-time number formatting for the AuditLog sheet (run once).
Public Sub PrepareAuditSheet()
    Dim tbl As ListObject

    On Error GoTo ErrorHandler

    Set tbl = modDatabase.GetTable("AuditLog", "tblAuditLog")
    tbl.ListColumns("LogDate").DataBodyRange.numberFormat = "dd-mm-yyyy"
    tbl.ListColumns("LogTime").DataBodyRange.numberFormat = "hh:mm"
    ThisWorkbook.Worksheets("AuditLog").Columns.AutoFit

    MsgBox "AuditLog sheet formatted.", vbInformation, "Setup"
    Exit Sub

ErrorHandler:
    MsgBox "Could not format the AuditLog sheet." & vbCrLf & Err.description, vbExclamation, "Setup"
End Sub

' The time in the log may be stored as a real time or as plain text.
Private Function TimeText(ByVal v As Variant) As String
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    If VarType(v) = vbDate Then
        TimeText = Format$(v, "hh:nn")
    Else
        TimeText = CStr(v)
    End If
End Function

' Sorts an array of text values A-Z (in place).
Private Sub SortTextAscending(ByRef items As Variant)
    Dim i As Long, j As Long
    Dim tmp As Variant

    For i = LBound(items) + 1 To UBound(items)
        tmp = items(i)
        j = i - 1
        Do While j >= LBound(items)
            If StrComp(CStr(items(j)), CStr(tmp), vbTextCompare) <= 0 Then Exit Do
            items(j + 1) = items(j)
            j = j - 1
        Loop
        items(j + 1) = tmp
    Next i
End Sub

' Fills a combo with "All ..." followed by every different value in one log column.
Public Sub FillDistinctCombo(ByVal Cmb As MSForms.ComboBox, ByVal columnName As String, _
                             ByVal allText As String)
    Dim tbl As ListObject
    Dim data As Variant
    Dim seen As Object
    Dim keys As Variant
    Dim colNo As Long, i As Long
    Dim v As String

    Cmb.Clear
    Cmb.AddItem allText

    Set tbl = modDatabase.GetTable("AuditLog", "tblAuditLog")
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare             ' "Admin" and "admin" count as the same

    data = tbl.DataBodyRange.value
    colNo = tbl.ListColumns(columnName).index

    For i = 1 To UBound(data, 1)
        v = Trim$(CStr(data(i, colNo)))
        If Len(v) > 0 Then seen(v) = True
    Next i

    If seen.count = 0 Then Exit Sub

    keys = seen.keys
    SortTextAscending keys
    For i = LBound(keys) To UBound(keys)
        Cmb.AddItem CStr(keys(i))
    Next i
End Sub

' Finds the log entries that match the filters, newest first.
' Returns a 2D array (1 To rows, 1 To 6) with columns
' Date, Time, User, Action, Record ID, Description - or Empty when nothing matches.
' totalMatches gets the real number of matches even if only maxRows are returned.
Public Function BuildAuditRows(ByVal fromDate As Date, ByVal toDate As Date, _
                               ByVal userFilter As String, ByVal actionFilter As String, _
                               ByVal searchText As String, ByRef totalMatches As Long, _
                               Optional ByVal maxRows As Long = 3000) As Variant
    Dim tbl As ListObject
    Dim data As Variant
    Dim colDate As Long, colTime As Long, colUser As Long
    Dim colAction As Long, colRecord As Long, colDesc As Long
    Dim rangeStart As Date, rangeEnd As Date
    Dim matchIndex() As Long
    Dim result() As Variant
    Dim i As Long, k As Long, shown As Long
    Dim dt As Date
    Dim isMatch As Boolean

    totalMatches = 0

    Set tbl = modDatabase.GetTable("AuditLog", "tblAuditLog")
    If tbl.DataBodyRange Is Nothing Then Exit Function

    data = tbl.DataBodyRange.value
    colDate = tbl.ListColumns("LogDate").index
    colTime = tbl.ListColumns("LogTime").index
    colUser = tbl.ListColumns("Username").index
    colAction = tbl.ListColumns("Action").index
    colRecord = tbl.ListColumns("RecordID").index
    colDesc = tbl.ListColumns("Description").index

    rangeStart = DateValue(fromDate)
    rangeEnd = DateValue(toDate) + 1
    userFilter = Trim$(userFilter)
    actionFilter = Trim$(actionFilter)
    searchText = Trim$(searchText)

    ReDim matchIndex(1 To maxRows)

    For i = UBound(data, 1) To 1 Step -1                ' newest first
        isMatch = False

        If Len(CStr(data(i, colAction))) > 0 And IsDate(data(i, colDate)) Then
            dt = CDate(data(i, colDate))
            If dt >= rangeStart And dt < rangeEnd Then
                isMatch = True

                If Len(userFilter) > 0 Then
                    If StrComp(CStr(data(i, colUser)), userFilter, vbTextCompare) <> 0 Then isMatch = False
                End If

                If isMatch And Len(actionFilter) > 0 Then
                    If StrComp(CStr(data(i, colAction)), actionFilter, vbTextCompare) <> 0 Then isMatch = False
                End If

                If isMatch And Len(searchText) > 0 Then
                    If InStr(1, CStr(data(i, colRecord)), searchText, vbTextCompare) = 0 And _
                       InStr(1, CStr(data(i, colDesc)), searchText, vbTextCompare) = 0 And _
                       InStr(1, CStr(data(i, colAction)), searchText, vbTextCompare) = 0 And _
                       InStr(1, CStr(data(i, colUser)), searchText, vbTextCompare) = 0 Then
                        isMatch = False
                    End If
                End If
            End If
        End If

        If isMatch Then
            totalMatches = totalMatches + 1
            If shown < maxRows Then
                shown = shown + 1
                matchIndex(shown) = i
            End If
        End If
    Next i

    If shown = 0 Then Exit Function

    ReDim result(1 To shown, 1 To 6)
    For k = 1 To shown
        i = matchIndex(k)
        result(k, 1) = CDate(data(i, colDate))
        result(k, 2) = TimeText(data(i, colTime))
        result(k, 3) = CStr(data(i, colUser))
        result(k, 4) = CStr(data(i, colAction))
        result(k, 5) = CStr(data(i, colRecord))
        result(k, 6) = CStr(data(i, colDesc))
    Next k

    BuildAuditRows = result
End Function

' Writes the matching log entries onto the Reports sheet. Returns the number of rows written.
Public Function WriteAuditToSheet(ByVal fromDate As Date, ByVal toDate As Date, _
                                  ByVal userFilter As String, ByVal actionFilter As String, _
                                  ByVal searchText As String) As Long
    Dim ws As Worksheet
    Dim data As Variant
    Dim total As Long, n As Long
    Dim subtitle As String

    data = BuildAuditRows(fromDate, toDate, userFilter, actionFilter, searchText, total, 50000)

    Set ws = ThisWorkbook.Worksheets("Reports")
    ws.Cells.Clear

    With ws.Range("A1")
        .value = "AUDIT LOG"
        .Font.Size = 16
        .Font.Bold = True
    End With

    subtitle = "Period: " & Format$(fromDate, "dd-mm-yyyy") & " to " & Format$(toDate, "dd-mm-yyyy")
    If Len(Trim$(userFilter)) > 0 Then subtitle = subtitle & "     User: " & userFilter
    If Len(Trim$(actionFilter)) > 0 Then subtitle = subtitle & "     Action: " & actionFilter
    If Len(Trim$(searchText)) > 0 Then subtitle = subtitle & "     Search: " & searchText
    subtitle = subtitle & "     Generated " & Format$(Now, "dd-mm-yyyy hh:nn") & " by " & modMain.CurrentUsername
    ws.Range("A2").value = subtitle

    ws.Range("A4:F4").value = Array("Date", "Time", "User", "Action", "Record ID", "Description")
    With ws.Range("A4:F4")
        .Interior.Color = ThemeNavy()
        .Font.Color = ThemeWhite()
        .Font.Bold = True
    End With

    If IsEmpty(data) Then
        ws.Range("A5").value = "No log entries match these filters."
        ws.Range("A4:F4").Columns.AutoFit
        Exit Function
    End If

    n = UBound(data, 1)

    ' Formats first, so text such as "-" or "=1+1" always stays plain text
    ws.Range(ws.Cells(5, 1), ws.Cells(4 + n, 1)).numberFormat = "dd-mm-yyyy"
    ws.Range(ws.Cells(5, 2), ws.Cells(4 + n, 6)).numberFormat = "@"
    ws.Range(ws.Cells(5, 1), ws.Cells(4 + n, 6)).value = data

    ws.Range(ws.Cells(4, 1), ws.Cells(4 + n, 6)).Borders.LineStyle = xlContinuous

    With ws.Cells(6 + n, 1)
        .value = n & " entries"
        If total > n Then .value = n & " entries (the newest of " & total & ")"
        .Font.Bold = True
    End With

    ws.Range(ws.Cells(4, 1), ws.Cells(4 + n, 6)).Columns.AutoFit
    If ws.Columns(6).ColumnWidth > 80 Then ws.Columns(6).ColumnWidth = 80

    WriteAuditToSheet = n
End Function

