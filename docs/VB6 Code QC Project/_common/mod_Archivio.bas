Attribute VB_Name = "mod_Archivio"
Option Explicit

Public Const dbName = "dbChemicalQC.mdb"

Public dbCodeName As String


Public dbPath As String
Public bExistAccount As Boolean
Public bExistAdministrator As Boolean
Public bExistManager As Boolean
Public bExistTCO As Boolean
Public bLoginAvvio As Boolean
Public usa_pass As Boolean
Public MydbName As String



'\\\\\\\\\\\\\\\\\\\\\\\\ database \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public dbCode As New ADODB.Connection

Public dbCodeNew As New ADODB.Connection
Public dbTabCodeNew As New ADODB.Recordset

Public dbChemicalQC As New ADODB.Connection
'Public dbTabLotti As New ADODB.Recordset
Public dbTabCode As New ADODB.Recordset
Public dbTabQCType As New ADODB.Recordset
Public dbTabTestType As New ADODB.Recordset
Public dbTabUserAccount As New ADODB.Recordset
Public dbTabLaboratorio As New ADODB.Recordset
Public dbTabReport As New ADODB.Recordset
Public dbTabOperator As New ADODB.Recordset
Public dbTabMeter As New ADODB.Recordset
Public dbTabPHMeter As New ADODB.Recordset
Public dbTabTurbMeter As New ADODB.Recordset
Public dbTabSpectMeter As New ADODB.Recordset
Public dbTabMachineOperator As New ADODB.Recordset
Public dbTabDepartment As New ADODB.Recordset
Public dbTabRecipeRevisionHistory As New ADODB.Recordset
Public dbTabFinishGood As New ADODB.Recordset
Public dbTabFGrevisionHistory As New ADODB.Recordset

Public bDoTip As Boolean


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Private Sub OpenTab()


    On Error GoTo ERR_OPEN
    
    dbTabReport.Open "SELECT *  FROM TabReport  order by id -1 ", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabUserAccount.Open "SELECT *  FROM TabUserAccount order by USERID", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
    'dbTabCode.Open "SELECT *  FROM TabCode order by code ", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabCode.Open "SELECT *  FROM TabCode ORDER BY Code", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
     

    
    dbTabOperator.Open "SELECT *  FROM TabOperator order by Name", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabLaboratorio.Open "SELECT *  FROM TabLaboratorio", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabMeter.Open "SELECT *  FROM TabMeter order by code", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabPHMeter.Open "SELECT *  FROM TabPHMeter order by code", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabTurbMeter.Open "SELECT *  FROM TabTurbMeter order by code", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabSpectMeter.Open "SELECT *  FROM TabSpectMeter order by code", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabQCType.Open "SELECT *  FROM TabQCType order by Type", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabTestType.Open "SELECT *  FROM TabTestType order by Type", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabDepartment.Open "SELECT *  FROM TabDepartment order by Code", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabMachineOperator.Open "SELECT *  FROM TabMachineOperator order by Name", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
  
    dbTabRecipeRevisionHistory.Open "SELECT *  FROM TabRecipeRevisionHistory order by RevNumber", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
    
    dbTabFinishGood.Open "SELECT *  FROM TabFinishGood order by Code", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabFGrevisionHistory.Open "SELECT *  FROM TabFGrevisionHistory order by id", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
  
ERR_END:

    On Error GoTo 0
    Exit Sub
ERR_OPEN:
    Debug.Print Err.Description
    If InStr(Err.Description, "FGrevisionHistory") > 0 Then
        If AddTable("TabFGrevisionHistory") Then
    
            Call m_CreateArchivio(dbPath, MydbName)
            Resume ERR_END
        End If
    End If
    
     If InStr(Err.Description, "FinishGood") > 0 Then
        If AddTable("TabFinishGood") Then
    
            Call m_CreateArchivio(dbPath, MydbName)
            Resume ERR_END
        End If
    End If
    
       
    
    
    
    Resume ERR_END
   
    
End Sub

Public Function m_CreateArchivio(ByRef t_path As String, ByVal A_NAME As String, _
            Optional ByRef T_TEXT As String, Optional ByRef r_c As Boolean, _
            Optional ByRef maxRecord As Integer, Optional Index As Integer = 0, Optional ByVal OnlyCodeDb As Boolean) As Boolean
    Dim rc As Boolean
    rc = True ' default
   On Error GoTo ERR_CREATE_OBJECT
   
  ' MsgBox t_path & dbName
  ' MsgBox t_path & dbCodeName
   
     If OnlyCodeDb Then GoTo cont
     

        '   /////////////////////////////////////////////////////
                If dbChemicalQC.State Then dbChemicalQC.Close
        '   /////////////////////////////////////////////////////
        dbChemicalQC.CursorLocation = adUseServer
        dbChemicalQC.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & t_path & dbName


        
        
cont:
       
        If OnlyCodeDb Then
        
            If dbCodeName = A_NAME Then
            Else
                PopupMessage 2, "Please select dbCode.mdb file...."
                rc = False
                GoTo END_FN
            End If
        
        End If
        '   /////////////////////////////////////////////////////
                If dbCode.State Then dbCode.Close
        '   /////////////////////////////////////////////////////
        dbCode.CursorLocation = adUseServer
        dbCode.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & t_path & dbCodeName
        
        
        
        dbCode.Execute ("ALTER TABLE TabFinishGood ADD RangeFormula varchar(30) WITH COMPRESSION")
            
        Call OpenTab
        
       
END_FN:
    On Error GoTo 0
   ' If rc And Index = 0 Then SaveSetting App.Title, "ARCHIVIO", "PATH", t_path
    m_CreateArchivio = rc
    Exit Function
ERR_CREATE_OBJECT:
    'MsgBox Err.Description
    If InStr(1, Err.Description, "esistente") > 0 Or InStr(1, Err.Description, "exist") > 0 Then Resume Next
    rc = False
    r_c = False
    MsgBox Err.Description
    T_TEXT = Err.Description
    Select Case Err.NUMBER
        Case -2147467259
            'MsgBox err.NUMBER & vbCrLf & err.Description
            Call SearchArchivio(t_path, A_NAME)
            rc = True
    End Select
    Resume END_FN
End Function


Public Function CheckFilePath(FileName As String, Path As String, ByRef Name As String, Optional ByRef old_name As String, Optional USER_ESTENSIONE As String = ".mdb", Optional bAvvisa As Boolean = True) As Boolean
    Dim a
    Dim n_String As String
    Dim new_name As String
    CheckFilePath = False
    If InStr(1, FileName, USER_ESTENSIONE) Then
        CheckFilePath = True
            For a = Len(FileName) To 1 Step -1
                n_String = Mid(FileName, a, 1)
                new_name = Right(FileName, Len(FileName) - a)
              
                If n_String = "\" Then
                    
                    If UCase(Name) <> UCase(new_name) Then
                        If bAvvisa Then
                            If MsgBox("Attenzione si è scelto un file con nome differente, procedo ugualmente?", vbInformation + vbYesNo, "Cambio percorso archivio") = vbYes Then
                                old_name = Name
                                Name = new_name
                                Path = Left(FileName, a - 1)
                                Exit Function
                            Else
                                CheckFilePath = False
                                Exit Function
                            End If
                        Else
                            old_name = Name
                            Name = new_name
                            Path = Left(FileName, a - 1)
                            Exit Function
                        End If
                        
                    Else
                            old_name = Name
                            Name = new_name
                            Path = Left(FileName, a - 1)
                            Exit Function
                    End If
                End If
            Next
    End If
End Function
Private Function CreaDirReport(ByVal sPath As String)
    If sPath = "" Then Exit Function
    If DirExists(sPath) = False Then MakePath (sPath)
End Function


Public Function CheckAndModify(ByVal MyName As String) As String
    Dim rc As Boolean
    Dim sString As String
    Dim LeftString As String
    Dim RightString As String
    Dim AccPosition As Integer
    If (InStr(MyName, "'")) > 0 Then
        '------------------------------------------
        ' ho un accento modifico il nome
        '------------------------------------------
        AccPosition = (InStr(MyName, "'"))
        LeftString = Left(MyName, AccPosition - 1)
        MyName = LeftString
    End If
    
    
    CheckAndModify = MyName
    
End Function
Private Sub SearchArchivio(ByRef t_path As String, ByVal A_NAME As String)


SEARCH_PATH:
    
    
    If SearchPathArchivio(t_path, A_NAME) Then
    
           
            FileCopy t_path & A_NAME, APP_DATA_FOLDER & dbName
            SaveSetting App.Title, "ARCHIVIO", "PATH", APP_DATA_FOLDER

        
           If m_CreateArchivio(APP_DATA_FOLDER, dbName) Then
           
                 
            
           Else
           End If
    Else
            If F_MsgBox(("Invalid Database.Select another Path")) Then
                GoTo SEARCH_PATH
            Else
            End If
    End If
End Sub
Private Function SearchPathArchivio(ByRef t_path As String, ByRef A_NAME As String) As Boolean
    Dim szFilename As String
    SearchPathArchivio = False
    szFilename = DialogFile(F_MAIN.hWnd, 1, "Open", dbName, "Database Access" & Chr(0) & "*.mdb" & Chr(0) & ("Tutti i files") & Chr(0) & "*.*", "", "mdb")
    If InStr(1, szFilename, ".mdb") Then
        Call SplitPathFile(szFilename, t_path, A_NAME)
        
       ' t_path = Left(szFilename, Len(szFilename) - Len(A_NAME))
        
        SearchPathArchivio = True
    End If
    If szFilename = "" Then SearchPathArchivio = False
End Function



'-----------------------------------------------------------------------------------

'                       Database : Compact and Repair

'-----------------------------------------------------------------------------------

Public Function ExecuteCompactDB(pFileName As String, Optional bValue As Boolean) As Boolean
Dim rc As Boolean
On Error GoTo ErrH

Dim CONN As New JRO.JetEngine
Dim ConnstringSorg As String, ConnstringDest As String

    ' Ensure file is not read only
    SetAttr pFileName, vbNormal
    rc = True
    If bValue = False Then
        If dbChemicalQC.State = adStateOpen Then dbChemicalQC.Close
    End If
    
    ConnstringSorg = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
    pFileName & ";Jet OLEDB:Engine Type=5;"
    
    ConnstringDest = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
    pFileName & "Temp" & ";Jet OLEDB:Engine Type=5;"
    
    
    
    Screen.MousePointer = vbHourglass
  ' MsgBox "Attenzione l'operazione potrebbe richiedere qualche minuto." & vbCrLf & _
   "Dipende dalle dimensioni del file..."
    
    CONN.CompactDatabase ConnstringSorg, ConnstringDest
    
    
    Screen.MousePointer = vbDefault
    
    'Copia il file compattato.
    Kill pFileName
    FileCopy pFileName & "Temp", pFileName
    Kill pFileName & "Temp"
ERR_END:
    Set CONN = Nothing
    ExecuteCompactDB = rc
    Exit Function
ErrH:
    rc = False
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
    
   ' FRMMenu.sbMain.Panels(1).Text = err.Description
    Resume ERR_END
    
End Function





Public Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    ' test the directory attribute
    DirExists = GetAttr(DirName) And vbDirectory
    'MsgBox DirExists
ErrorHandler:
    ' if an error occurs, this function returns False
End Function
Public Function SettSavePath(ByVal Path As String) As Boolean
    Dim rc As Boolean
    On Error GoTo ERR_SETT
    rc = True
    If DirExists(Path) Then
    
    Else
        '----------------------------
        ' non esiste!!!
        '----------------------------
        MakePath (Path)
    End If
    
ERR_END:
    On Error GoTo 0
    SettSavePath = rc
    Exit Function
ERR_SETT:
    rc = False
    Resume ERR_END
End Function

Public Function MessaggioErrore(ByVal Index As Long, Optional ByVal strFunzione As String, Optional ByVal strErrore As String) As String
Dim strError As String
    Select Case Index
        Case 13
            strError = ("Attenzione impossibile procedere con l'operazione")
            
        Case Else
            strError = strErrore
            
    End Select
    MessaggioErrore = strError
End Function


Public Function SetNewdBasePath(ByVal MyPath As String, Optional sName As String, Optional Index As Integer = 0) As String
    Dim MydbName As String
        
        
        If Index = 0 Then
            MydbName = dbCodeName ' GetSetting(App.Title, "ARCHIVIO", "NOME", dbCodeName)
        Else
            MydbName = sName
        End If
        
        
        
        If F_SEARCHARCH.DoShow(MyPath, MydbName) Then
                
                PopupMessage 2, ("Archivio ripristinato correttamente...")
         
                SaveSetting App.Title, "ARCHIVIO", "PATH", MyPath
                SaveSetting App.Title, "ARCHIVIO", "NOME", MydbName
                dbPath = MyPath
              
        Else
                'PopupMessage 2, "Attenzione, l'archivio non è stato modificato : errore nel file", , True
        End If


End Function

