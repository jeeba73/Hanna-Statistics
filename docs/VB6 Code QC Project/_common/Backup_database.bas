Attribute VB_Name = "Backup_database"
Option Explicit



Public Function BackupDatabase()
Dim bDatabaseCopy As Boolean
Dim DatabaseCopyDate As Date

On Error GoTo ERR_CHECK:

dbPath = GetSetting(App.Title, "ARCHIVIO", "PATH", APP_DATA_FOLDER)

If FileExists(dbPath & dbName) And FileExists(dbPath & dbCodeName) Then
    
    DatabaseCopyDate = FormatDateTime((GetSetting(App.Title, "Database Check", "DatabaseCopyDate", "0.0.00")), vbShortDate)
    bDatabaseCopy = IIf(FormatDateTime(Now(), vbShortDate) = DatabaseCopyDate, False, True)
    
    
    If bDatabaseCopy Then
        Debug.Print dbPath & dbName
        
        FileCopy dbPath & dbName, USER_DOCUMENTI & dbName
        FileCopy dbPath & dbCodeName, USER_DOCUMENTI & dbCodeName
        PopupMessage 2, "FileCopy OK", , , "Database Backup"
        SaveSetting App.Title, "Database Check", "DatabaseCopyDate", FormatDateTime(Now())
    End If

End If

ERR_END:
    On Error GoTo 0
    Exit Function
ERR_CHECK:
    MsgBox err.Description
    GoTo ERR_END:
End Function
