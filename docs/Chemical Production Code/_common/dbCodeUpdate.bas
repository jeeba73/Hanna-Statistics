Attribute VB_Name = "dbCodeUpdate"
Option Explicit

Public dbPathNEW As String

        
Public dbRelease As New ADODB.Recordset
Public dbCodeRelease As String
Public dbCodeDate As String
Public dbCodeOperator As String

Public bAddNewDatabaseRelease As Boolean



Public Function CheckDbCodeRelease()

With dbRelease
 
    .Open "SELECT *  FROM TabRelease", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
     If .EOF Then
     Else
        .MoveFirst
         dbCodeRelease = !Release
         dbCodeDate = !Date
         dbCodeOperator = !Operator
     End If
    .Close
End With

End Function


Public Function SetDbCodeRelease()

With dbRelease
 
    .Open "SELECT *  FROM TabRelease", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
     If .EOF Then
     Else
        .MoveFirst
         !Release = dbCodeRelease
         !Date = Now()
         !Operator = MyOperatore.Name
     End If
    .Update
    .Close
End With
        
        
End Function

Public Function AddReleaseNumber()
Dim Maj As Integer
Dim Med As Integer
Dim Rel As Integer

Dim var() As String
    
    var = Split(dbCodeRelease, ".")
    Maj = var(0)
    Med = var(1)
    Rel = var(2)
    
    If Rel + 1 = 1000 Then
        If Med + 1 = 100 Then
            Maj = Maj + 1
            Med = 0
            Rel = 0
        Else
            Med = Med + 1
            Rel = 0
        End If
    Else
        Rel = Rel + 1
    End If
    dbCodeRelease = CStr(Maj) & "." & CStr(Med) & "." + CStr(Rel)
    Call SetDbCodeRelease

End Function

Public Function CheckCodeDB() As Boolean
    CheckCodeDB = False
    
    If GetSetting(App.Title, "Opzioni", "Update dbCode", False) = False Then
        Call SetNewDatabase
       
    End If
    
    CheckDbCodeRelease
    
End Function

Public Function SetNewDatabase()

 Dim szFilename As String
Dim A_NAME As String
        PopupMessage 2, "Please select NEW dbCode.mdb..."
        szFilename = DialogFile(F_MAIN.hWnd, 1, "Open", "dbCode.mdb", "Access" & Chr(0) & "*.mdb" & Chr(0) & "Tutti i files" & Chr(0) & "*.*", "", "mdb")
        If InStr(1, szFilename, ".mdb") Then
            Call SplitPathFile(szFilename, dbPathNEW, A_NAME)
            'non serve per Chemical Production
            'Call UpdateNewdbCode
        End If
        
        dbCode.Close
        Kill APP_DATA_FOLDER & "dbCode.mdb"
        FileCopy szFilename, APP_DATA_FOLDER & "dbCode.mdb"
        SaveSetting App.Title, "Opzioni", "Update dbCode", True
        PopupMessage 2, "Restart " & PROGRAM_NAME
        
        End
End Function

