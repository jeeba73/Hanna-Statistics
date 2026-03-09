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

    If GetSetting(App.Title, "Opzioni", "Update dbCode 03", False) = False Then
        Call SetNewDatabase
       
    End If
    
    CheckDbCodeRelease
    
End Function

Public Function SetNewDatabase()

 Dim szFilename As String
Dim A_NAME As String
Dim rc As Boolean
        
        
        PopupMessage 2, "Please select NEW dbCode.mdb..."

        szFilename = DialogFile(F_MAIN.hwnd, 1, "Open", "dbCode.mdb", "Access" & Chr(0) & "*.mdb" & Chr(0) & "Tutti i files" & Chr(0) & "*.*", "", "mdb")
        
        If InStr(1, szFilename, ".mdb") Then
            Call SplitPathFile(szFilename, dbPathNEW, A_NAME)
            'non serve per Chemical Production
            Call UpdateNewdbCode
        Else
            PopupMessage 2, "Please update to new Database", , True, "dbCode NEW"
            Exit Function
        End If
        
         
            SaveSetting App.Title, "Opzioni", "Update dbCode 03", True
            
            dbCodeNew.Close
            dbCode.Close
            
           
            
            FileCopy szFilename, APP_DATA_FOLDER & "dbCode.mdb"
            
            SaveSetting App.Title, "Update", "dbName", "dbCode.mdb"

            PopupMessage 2, "Restart MR"
        
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
    Dim fields() As String
    Dim bCopy As Boolean
    Dim FieldName As String
    
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
                    
                    dbTabCodeNew.filter = ""
                    dbTabCodeNew.filter = "Code='" & Code & "'"
                

                    If dbTabCodeNew.EOF Then
                         
                        Debug.Print Code ' è possibile???
                       

                                
                        dbTabCodeNew.AddNew
                        
                        For t = 1 To .fields.Count - 1
                        FieldName = .fields(t).Name
                        
                            If FieldName = "Certified" Then
                            ElseIf FieldName = "LastLot" Then
                            ElseIf FieldName = "HannaFGCode" Then
                            ElseIf FieldName = "Description" Then
                                dbTabCodeNew.fields("ProductName") = IIf(IsNull(.fields(t)), "", Trim(.fields(t)))
                            Else
                            dbTabCodeNew.fields(FieldName) = IIf(IsNull(.fields(t)), "", Trim(.fields(t)))
                            End If
                        Next


                    Else
                        '--------------------------------------------------------------------------------------------------------------------------
                        '  questi dati non sono presenti in CP o QC
                        '  FWParameterFormula  STDMR2  MS1val  MS1vol  MS2dil  MS2vol  MSEXP   STDMatrix   STDVolume   STDExp  STDNote STDStorage
                        '--------------------------------------------------------------------------------------------------------------------------
                    
                        
                        'For t = 54 To 65 '
                          'Debug.Print dbTabCodeNew.fields(t + 14).Name & " - " & .fields(t).Name
                      
                        dbTabCodeNew.fields("FWParameterFormula") = IIf(IsNull(.fields("FWParameterFormula")), "", Trim(.fields("FWParameterFormula")))
                        dbTabCodeNew.fields("STDMR2") = IIf(IsNull(.fields("STDMR2")), "", Trim(.fields("STDMR2")))
                        dbTabCodeNew.fields("MS1val") = IIf(IsNull(.fields("MS1val")), "", Trim(.fields("MS1val")))
                        dbTabCodeNew.fields("MS1vol") = IIf(IsNull(.fields("MS1vol")), "", Trim(.fields("MS1vol")))
                        dbTabCodeNew.fields("MS2dil") = IIf(IsNull(.fields("MS2dil")), "", Trim(.fields("MS2dil")))
                        dbTabCodeNew.fields("MS2vol") = IIf(IsNull(.fields("MS2vol")), "", Trim(.fields("MS2vol")))
                        dbTabCodeNew.fields("MSEXP") = IIf(IsNull(.fields("MSEXP")), "", Trim(.fields("MSEXP")))
                        dbTabCodeNew.fields("STDMatrix") = IIf(IsNull(.fields("STDMatrix")), "", Trim(.fields("STDMatrix")))
                        dbTabCodeNew.fields("STDVolume") = IIf(IsNull(.fields("STDVolume")), "", Trim(.fields("STDVolume")))
                        dbTabCodeNew.fields("STDExp") = IIf(IsNull(.fields("STDExp")), "", Trim(.fields("STDExp")))
                        dbTabCodeNew.fields("STDNote") = IIf(IsNull(.fields("STDNote")), "", Trim(.fields("STDNote")))
                        dbTabCodeNew.fields("STDStorage") = IIf(IsNull(.fields("STDStorage")), "", Trim(.fields("STDStorage")))
                        dbTabCodeNew.fields("ConcHannaParameter") = IIf(IsNull(.fields("ConcHannaParameter")), "", Trim(.fields("ConcHannaParameter")))
        
                        'Next
                    
                    
                        dbTabCodeNew!DateModified = Now()
                        c = c + 1
                    End If
                    
                   
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
    MsgBox Err.Description
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




