VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBackup 
   Caption         =   "Backup Database"
   ClientHeight    =   9210.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14955
   OleObjectBlob   =   "frmBackup.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmBackup - make, list and clean up backups (ADMIN only)
' ============================================================

' How many of the newest backups to keep when cleaning up
Private Const KEEP_COUNT As Long = 20

' ---------- Startup ----------
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    BuildLayout
    StyleControls
    UpdateButtonState
    RefreshAll
    Exit Sub

ErrorHandler:
    MsgBox "The backup screen could not load completely." & vbCrLf & Err.description, _
           vbExclamation, "Backup"
End Sub

Private Sub SetStatus(ByVal message As String)
    Me.Controls("lblStatusMsg").caption = message
End Sub

' ---------- Build the parts made by code ----------
Private Sub BuildLayout()
    Me.BackColor = ThemeNavy()

    modUI.AddLabel Me, "lblHeaderBar", 0, 0, 780, 40, "", ThemeSidebar()
    modUI.AddLabel Me, "lblTitle", 14, 8, 400, 26, "BACKUP DATABASE", -1, ThemeCyan(), 14, True

    modUI.AddLabel Me, "lblLast", 20, 50, 720, 18, "", -1, ThemeWhite(), 10, True
    modUI.AddLabel Me, "lblFolder", 20, 70, 720, 16, "", -1, ThemeGrey(), 8

    modUI.AddLabel Me, "lblHead0", 23, 140, 340, 13, "FILE NAME", -1, ThemeCyan(), 7, True
    modUI.AddLabel Me, "lblHead1", 363, 140, 140, 13, "DATE", -1, ThemeCyan(), 7, True
    modUI.AddLabel Me, "lblHead2", 503, 140, 100, 13, "SIZE", -1, ThemeCyan(), 7, True

    modUI.AddLabel Me, "lblSummary", 20, 420, 720, 18, "", -1, ThemeWhite(), 9, True
    modUI.AddLabel Me, "lblStatusMsg", 20, 440, 720, 18, "", -1, ThemeGrey(), 9
End Sub

Private Sub StyleControls()
    StyleButton btnBackupNow, "BACKUP NOW", 20, 96, 200, 34, ThemeGreen(), ThemeNavy()
    StyleButton btnOpenFolder, "OPEN FOLDER", 230, 96, 150, 34, ThemeCyan(), ThemeNavy()
    StyleButton btnAuto, "AUTO BACKUP", 390, 96, 170, 34, ThemePanel(), ThemeWhite()
    StyleButton btnRestoreHelp, "HOW TO RESTORE", 570, 96, 150, 34, ThemeActive(), ThemeWhite()
    StyleButton btnDeleteOld, "DELETE OLD BACKUPS", 20, 382, 220, 30, ThemeRed(), ThemeNavy()
    StyleButton btnClose, "CLOSE", 580, 382, 140, 30, ThemePanel(), ThemeWhite()

    With lstBackups
        .Left = 20: .Top = 154: .Width = 720: .Height = 220
        .BackColor = ThemePanel(): .ForeColor = ThemeWhite()
        .Font.name = "Segoe UI": .Font.Size = 9
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, ByVal captionText As String, _
                        ByVal leftPos As Single, ByVal topPos As Single, _
                        ByVal widthVal As Single, ByVal heightVal As Single, _
                        ByVal bgColor As Long, ByVal fgColor As Long)
    With btn
        .caption = captionText
        .Left = leftPos: .Top = topPos: .Width = widthVal: .Height = heightVal
        .BackColor = bgColor: .ForeColor = fgColor
        .Font.name = "Segoe UI": .Font.Size = 9: .Font.Bold = True
        .TakeFocusOnClick = False
    End With
End Sub

' Only an ADMIN may use the buttons that change anything.
Private Sub UpdateButtonState()
    Dim allowed As Boolean
    allowed = modMain.IsAdmin()

    btnBackupNow.Enabled = allowed
    btnAuto.Enabled = allowed
    btnDeleteOld.Enabled = allowed

    If Not allowed Then SetStatus "View only: only an administrator can change backups."
End Sub

' ---------- Refresh ----------
Private Sub RefreshAll()
    Dim count As Long
    Dim totalBytes As Double

    On Error GoTo ErrorHandler

    Me.Controls("lblLast").caption = "Last backup:  " & modBackup.LastBackupText()
    Me.Controls("lblFolder").caption = "Folder:  " & modBackup.BackupFolder()

    count = modBackup.FillBackupList(lstBackups, totalBytes)
    If count = 0 Then
        Me.Controls("lblSummary").caption = "No backups yet. Click BACKUP NOW to make the first one."
    Else
        Me.Controls("lblSummary").caption = count & " backup(s), " & _
                                            Format$(totalBytes / 1048576, "0.0") & " MB in total"
    End If

    ShowAutoState
    Exit Sub

ErrorHandler:
    MsgBox "The backup list could not be shown." & vbCrLf & Err.description, vbExclamation, "Backup"
End Sub

Private Sub ShowAutoState()
    If modBackup.IsAutoBackupOn() Then
        btnAuto.caption = "AUTO BACKUP: ON"
        btnAuto.ForeColor = ThemeGreen()
    Else
        btnAuto.caption = "AUTO BACKUP: OFF"
        btnAuto.ForeColor = ThemeOrange()
    End If
End Sub

' ---------- Buttons ----------
Private Sub btnBackupNow_Click()
    Dim filePath As String
    Dim errText As String

    On Error GoTo ErrorHandler

    Me.MousePointer = fmMousePointerHourGlass
    filePath = modBackup.CreateBackup()
    Me.MousePointer = fmMousePointerDefault

    LogAction "BACKUP", Mid$(filePath, InStrRev(filePath, "\") + 1), "Backup created"
    RefreshAll
    SetStatus "Backup saved."
    MsgBox "Backup saved:" & vbCrLf & filePath, vbInformation, "Backup complete"
    Exit Sub

ErrorHandler:
    errText = Err.description
    Me.MousePointer = fmMousePointerDefault
    MsgBox "The backup could not be created." & vbCrLf & errText & vbCrLf & vbCrLf & _
           "Check that the disk has free space and that you may write to the folder.", _
           vbExclamation, "Backup failed"
End Sub

Private Sub btnOpenFolder_Click()
    On Error GoTo ErrorHandler
    modInvoice.OpenFile modBackup.BackupFolder()
    Exit Sub

ErrorHandler:
    MsgBox "The folder could not be opened." & vbCrLf & Err.description, vbExclamation, "Backup"
End Sub

Private Sub btnAuto_Click()
    modBackup.SetAutoBackup Not modBackup.IsAutoBackupOn()
    LogAction "AUTO BACKUP SETTING", "-", "Automatic daily backup is now " & _
              IIf(modBackup.IsAutoBackupOn(), "ON", "OFF")
    ShowAutoState
End Sub

Private Sub btnRestoreHelp_Click()
    MsgBox "HOW TO RESTORE A BACKUP" & vbCrLf & vbCrLf & _
           "1.  Log out and close Excel completely." & vbCrLf & _
           "2.  Click OPEN FOLDER to find the backup you want. The date is in the file name." & vbCrLf & _
           "3.  Do NOT delete your current file. Rename it first, for example add _OLD to its name." & vbCrLf & _
           "4.  Copy the backup file next to it and rename the copy to your usual file name." & vbCrLf & _
           "5.  Open it, click Enable Content, log in and check your data." & vbCrLf & _
           "6.  Keep the _OLD file until you are sure everything is right.", _
           vbInformation, "Restore a backup"
End Sub

Private Sub btnDeleteOld_Click()
    Dim removed As Long

    On Error GoTo ErrorHandler

    If MsgBox("Delete all backups except the newest " & KEEP_COUNT & "?" & vbCrLf & _
              "This cannot be undone.", vbYesNo + vbExclamation, "Delete old backups") <> vbYes Then Exit Sub

    removed = modBackup.DeleteOldBackups(KEEP_COUNT)
    LogAction "DELETE OLD BACKUPS", "-", removed & " file(s) deleted"
    RefreshAll
    SetStatus removed & " old backup(s) deleted."
    Exit Sub

ErrorHandler:
    MsgBox "The old backups could not be deleted." & vbCrLf & Err.description, vbExclamation, "Backup"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


