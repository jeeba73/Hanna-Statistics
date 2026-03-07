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
         dbCodeDate = !date
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
         !date = Now()
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

Dim Var() As String
    
    Var = Split(dbCodeRelease, ".")
    Maj = Var(0)
    Med = Var(1)
    Rel = Var(2)
    
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
    If dbCodeName = "dbCodeQC.mdb" Then
        Call SetNewDatabase
        CheckCodeDB = True
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
            Call UpdateNewdbCode
        End If
        
            dbCodeName = "dbCode.mdb"
            
            dbCodeNew.Close
            
            FileCopy szFilename, APP_DATA_FOLDER & "dbCode.mdb"
            
              SaveSetting App.Title, "Update", "dbName", dbCodeName

            PopupMessage 2, "Restart QC"
        
      
            End
        
        


    
End Function


Private Function UpdateNewdbCode() As Boolean
    Dim rc As Boolean
    Dim i As Integer
    Dim t As Integer
    Dim c As Integer
    Dim Code As String
    Dim RangeMin As String
    Dim RangeMax As String
    Dim Fields() As String
    Dim bCopy As Boolean
    
    On Error GoTo ERR_SAVE
    rc = True
    
    Dim Index As Integer
    
    
    If dbCodeNew.State = 1 Then
    Else
        If m_CreateArchivioNewDatabase(dbPathNEW, "dbCode.mdb") Then
        
        Else
           ' bVerPeriodica = False
            PopupMessage 2, "Unable to access dbCode.mdb Database"
            rc = False
            GoTo ERR_END
        End If
    End If
    
    Index = 0
    
    Do Until Index = 2
    
        With dbTabCode
             .filter = ""
            If .EOF Then
            Else
                .MoveFirst
                c = 0
                For i = 1 To .RecordCount
                    Code = !Code
                    RangeMin = IIf(IsNull(!RangeMin), "", Trim(!RangeMin))
                    RangeMax = IIf(IsNull(!RangeMax), "", Trim(!RangeMax))
start:
                    dbTabCodeNew.filter = ""
                    If Index = 1 Then
                        dbTabCodeNew.filter = "Code='" & Code & "' and RangeMin='" & RangeMin & "' and RangeMax='" & RangeMax & "'"
                    Else
                    
                        dbTabCodeNew.filter = "Code='" & Code & "'"
                    End If

                    If dbTabCodeNew.EOF Then
                         
                        Debug.Print Code ' è possibile???
                       
                        If Index = 1 Then
                            If bCopy = False Then
                                dbTabCodeNew.filter = ""
                                dbTabCodeNew.filter = "Code='" & Code & "'"
                                ReDim Fields(.Fields.Count - 1)
                                For t = 1 To .Fields.Count - 1
                                    Fields(t) = IIf(IsNull(dbTabCodeNew.Fields(t)), "", Trim(dbTabCodeNew.Fields(t)))
                                Next
                                bCopy = True
                                GoTo start:
                            Else
                                ' l'ho già copiato
                                
                                
                                
                                dbTabCodeNew.AddNew
                                For t = 1 To .Fields.Count - 1
                                    dbTabCodeNew.Fields(t) = Fields(t)
                                Next
                                dbTabCodeNew!RangeMin = RangeMin
                                dbTabCodeNew!RangeMax = RangeMax
                                dbTabCodeNew!Hide = True ' dovrebbe essere un codice doppio
                                bCopy = False
                                GoTo cont:
                            End If
                        Else ' codice non presente in Chemical Production
                         
                            dbTabCodeNew.AddNew
                            dbTabCodeNew!Hide = True
                            dbTabCodeNew!Code = IIf(IsNull(.Fields(2)), "", Trim(.Fields(2)))
                            dbTabCodeNew!ProductName = IIf(IsNull(.Fields(3)), "", Trim(.Fields(3)))
                            dbTabCodeNew!Line = IIf(IsNull(.Fields(4)), "", Trim(.Fields(4)))
                            dbTabCodeNew!Recipe = IIf(IsNull(.Fields(5)), "", Trim(.Fields(5)))
                            
                        End If
                    Else
                        If Index = 1 Then
                            GoTo cont:
                        End If
                    End If
                    
                    'Debug.Print dbTabCodeNew.Fields(23).Name
                    'Debug.Print .Fields(6).Name
                    For t = 6 To 49
                       
                        dbTabCodeNew.Fields(t + 17) = IIf(IsNull(.Fields(t)), "", Trim(.Fields(t)))
                    Next
                    'Debug.Print dbTabCodeNew.Fields(53 + 14).Name
                    'Debug.Print .Fields(53).Name
                    For t = 53 To 65
                        dbTabCodeNew.Fields(t + 14) = IIf(IsNull(.Fields(t)), "", Trim(.Fields(t)))
                    Next
                    
                    dbTabCodeNew!DateModified = Now()
                    
                    c = c + 1
cont:
                    .MoveNext
                Next
               If .RecordCount > 0 And c > 1 Then dbTabCodeNew.Update
            End If
        End With
        
        ' ripeto il giro per aggiungere gli hanna code DOPPI
        Index = Index + 1
        
    Loop

ERR_END:
    On Error GoTo 0
    UpdateNewdbCode = rc
    Exit Function
ERR_SAVE:
    rc = False
    MsgBox err.Description
    Resume Next

End Function




Public Function m_CreateArchivioNewDatabase(ByRef t_path As String, ByVal A_NAME As String, _
            Optional ByRef T_TEXT As String, Optional ByRef r_c As Boolean, _
            Optional ByRef maxRecord As Integer) As Boolean
    Dim rc As Boolean
    Dim stringOpen As String
    
    rc = True ' default
    
   On Error GoTo ERR_CREATE_OBJECT
    '   /////////////////////////////////////////////////////
            If dbCodeNew.State Then dbCodeNew.Close
    '   /////////////////////////////////////////////////////
    dbCodeNew.CursorLocation = adUseServer
    If InStr(Len(t_path) - 1, t_path, "\") Then
        stringOpen = t_path & A_NAME
    Else
        stringOpen = t_path & "\" & A_NAME
    End If
  '  MsgBox stringOpen
    dbCodeNew.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & stringOpen ' t_path & "\" & A_NAME '"\Archivio.mdb" '& ";Mode=ReadWrite;Persist Security Info=False"
    
 
      
    Call OpenTabNewDatabase

  
    
END_FN:
    On Error GoTo 0
    
    m_CreateArchivioNewDatabase = rc
    Exit Function
ERR_CREATE_OBJECT:
    rc = False
    'MsgBox err.Description
    Resume Next
End Function








Private Sub OpenTabNewDatabase()
    On Error GoTo ERR_SUB
    dbTabCodeNew.Open "SELECT *  FROM TabCode order by id", dbCodeNew, adOpenKeyset, adLockOptimistic, adCmdText
ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_SUB:
    'MsgBox err.Description
    Resume Next
End Sub


