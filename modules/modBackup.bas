Attribute VB_Name = "modBackup"
Option Explicit

' ============================================================
' modBackup - safe copies of the whole workbook
' ============================================================

Private Const BACKUP_PREFIX As String = "InventoryBackup_"

' The file extension of this workbook (normally ".xlsm").
Private Function BackupExtension() As String
    Dim dotPos As Long

    dotPos = InStrRev(ThisWorkbook.name, ".")
    If dotPos > 0 Then
        BackupExtension = Mid$(ThisWorkbook.name, dotPos)
    Else
        BackupExtension = ".xlsm"
    End If
End Function

' The folder where backups are saved (created if needed).
' Uses the "BackupFolder" setting if it points to a real folder.
Public Function BackupFolder() As String
    Dim custom As String
    Dim basePath As String
    Dim folder As String

    custom = Trim$(GetSetting("BackupFolder", ""))
    If Len(custom) > 0 Then
        If Right$(custom, 1) <> "\" Then custom = custom & "\"
        If Len(Dir(custom, vbDirectory)) > 0 Then
            BackupFolder = custom
            Exit Function
        End If
    End If

    basePath = ThisWorkbook.Path
    ' A cloud (OneDrive) path starts with http and cannot be used for folders
    If Len(basePath) = 0 Or LCase$(Left$(basePath, 4)) = "http" Then
        basePath = Environ$("USERPROFILE") & "\Documents"
    End If

    folder = basePath & "\Backups\"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    BackupFolder = folder
End Function

' Makes a backup now and returns the full path of the new file.
' Raises an error (with the reason) if it cannot be made.
Public Function CreateBackup() As String
    Dim folder As String
    Dim filePath As String

    folder = BackupFolder()
    filePath = folder & BACKUP_PREFIX & Format$(Now, "yyyy-mm-dd_hhnn") & BackupExtension()

    ' Two backups in the same minute: add seconds so nothing is overwritten
    If Len(Dir(filePath)) > 0 Then
        filePath = folder & BACKUP_PREFIX & Format$(Now, "yyyy-mm-dd_hhnnss") & BackupExtension()
    End If

    ThisWorkbook.SaveCopyAs filePath

    SetSetting "LastBackupDate", Format$(Date, "yyyymmdd")
    CreateBackup = filePath
End Function

' ---------- Automatic daily backup ----------
Public Function IsAutoBackupOn() As Boolean
    IsAutoBackupOn = (UCase$(Trim$(GetSetting("AutoBackup", "Yes"))) = "YES")
End Function

Public Sub SetAutoBackup(ByVal turnOn As Boolean)
    If turnOn Then SetSetting "AutoBackup", "Yes" Else SetSetting "AutoBackup", "No"
End Sub

' Makes one backup per day, when an ADMIN or MANAGER logs in. Never stops the login.
Public Sub AutoBackupIfDue()
    Dim filePath As String

    If Not IsAutoBackupOn() Then Exit Sub
    If modMain.CurrentRole <> "ADMIN" And modMain.CurrentRole <> "MANAGER" Then Exit Sub
    If Trim$(GetSetting("LastBackupDate", "")) = Format$(Date, "yyyymmdd") Then Exit Sub

    On Error GoTo ErrorHandler
    filePath = CreateBackup()
    LogAction "AUTO BACKUP", Mid$(filePath, InStrRev(filePath, "\") + 1), "Daily backup created"
    Exit Sub

ErrorHandler:
    LogAction "AUTO BACKUP FAILED", "-", Err.description
End Sub

' "dd-mm-yyyy" of the last backup, or "never".
Public Function LastBackupText() As String
    Dim t As String

    t = Trim$(GetSetting("LastBackupDate", ""))
    If Len(t) = 8 Then
        LastBackupText = Right$(t, 2) & "-" & Mid$(t, 5, 2) & "-" & Left$(t, 4)
    Else
        LastBackupText = "never"
    End If
End Function

' ---------- Listing and cleaning ----------

' Fills names() with the backup file names, newest first, and returns how many.
Public Function ListBackups(ByRef names() As String) As Long
    Dim found As Collection
    Dim folder As String
    Dim fileName As String
    Dim i As Long, j As Long
    Dim tmp As String

    Set found = New Collection
    folder = BackupFolder()

    fileName = Dir(folder & BACKUP_PREFIX & "*" & BackupExtension())
    Do While Len(fileName) > 0
        found.Add fileName
        fileName = Dir
    Loop

    If found.count = 0 Then Exit Function

    ReDim names(1 To found.count)
    For i = 1 To found.count
        names(i) = CStr(found(i))
    Next i

    ' The date is inside the name, so sorting the names sorts by date (newest first)
    For i = 2 To found.count
        tmp = names(i)
        j = i - 1
        Do While j >= 1
            If StrComp(names(j), tmp, vbBinaryCompare) >= 0 Then Exit Do
            names(j + 1) = names(j)
            j = j - 1
        Loop
        names(j + 1) = tmp
    Next i

    ListBackups = found.count
End Function

Private Function FormatSize(ByVal bytes As Double) As String
    If bytes >= 1048576 Then
        FormatSize = Format$(bytes / 1048576, "0.0") & " MB"
    Else
        FormatSize = Format$(bytes / 1024, "#,##0") & " KB"
    End If
End Function

' Fills the list on the Backup screen. Returns how many backups there are.
Public Function FillBackupList(ByVal lst As MSForms.ListBox, ByRef totalBytes As Double) As Long
    Dim names() As String
    Dim folder As String, filePath As String
    Dim count As Long, i As Long, n As Long
    Dim fileSize As Double

    lst.Clear
    lst.ColumnCount = 3
    lst.ColumnWidths = "340 pt;140 pt;100 pt"
    totalBytes = 0

    count = ListBackups(names)
    If count = 0 Then Exit Function

    folder = BackupFolder()
    For i = 1 To count
        filePath = folder & names(i)

        lst.AddItem names(i)
        n = lst.ListCount - 1

        fileSize = 0
        On Error Resume Next
        lst.List(n, 1) = Format$(FileDateTime(filePath), "dd-mm-yyyy hh:nn")
        fileSize = FileLen(filePath)
        On Error GoTo 0

        lst.List(n, 2) = FormatSize(fileSize)
        totalBytes = totalBytes + fileSize
    Next i

    FillBackupList = count
End Function

' Deletes the oldest backups, keeping the newest keepCount. Returns how many were deleted.
Public Function DeleteOldBackups(ByVal keepCount As Long) As Long
    Dim names() As String
    Dim folder As String
    Dim count As Long, i As Long, removed As Long

    If keepCount < 1 Then keepCount = 1

    count = ListBackups(names)
    If count <= keepCount Then Exit Function

    folder = BackupFolder()
    For i = keepCount + 1 To count
        Err.Clear
        On Error Resume Next
        Kill folder & names(i)
        If Err.Number = 0 Then removed = removed + 1
        On Error GoTo 0
    Next i

    DeleteOldBackups = removed
End Function
