Attribute VB_Name = "mod_Archivio"
Option Explicit

Public Const dbName = "dbChemiMR.mdb"
Public Const dbCodeName = "dbCode.mdb"
Public dbPath As String
Public bExistAccount As Boolean
Public bExistAdministrator As Boolean
Public bExsistSTDPreparationManager As Boolean
Public bExsistLineLeader As Boolean
Public bLoginAvvio As Boolean
Public usa_pass As Boolean
Public MydbName As String



'\\\\\\\\\\\\\\\\\\\\\\\\ database \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public dbCode As New ADODB.Connection
Public dbChemicalMR As New ADODB.Connection

Public dbTabCode As New ADODB.Recordset
Public dbCodeNew As New ADODB.Connection
Public dbTabCodeNew As New ADODB.Recordset


Public dbTabPipette As New ADODB.Recordset

Public dbTabCodeClassification As New ADODB.Recordset
Public dbTabFrasiH As New ADODB.Recordset

Public dbTabMR As New ADODB.Recordset
Public dbTabMRrevisionHistory As New ADODB.Recordset
Public dbTabMRWarehouse As New ADODB.Recordset
Public dbTabSupplier As New ADODB.Recordset

Public dbTabPreparation As New ADODB.Recordset
Public dbTabStazioniNetwork As New ADODB.Recordset
Public dbTabUserAccount As New ADODB.Recordset
Public dbTabAcquisition As New ADODB.Recordset
Public dbTabPreparationNotes As New ADODB.Recordset
Public dbTabLaboratorio As New ADODB.Recordset
Public dbTabLocation As New ADODB.Recordset
Public dbTabMotherSolution As New ADODB.Recordset
Public bDoTip As Boolean

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Private Sub OpenTab()

    On Error GoTo ERR_OPEN

    dbTabLaboratorio.Open "SELECT *  FROM TabLaboratorio ", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabStazioniNetwork.Open "SELECT *  FROM TabStazioniNetwork order by Stazione ", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabUserAccount.Open "SELECT *  FROM TabUserAccount order by IndexPrivilege ", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabPreparation.Open "SELECT *  FROM TabPreparation order by id ", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabAcquisition.Open "SELECT *  FROM TabAcquisition order by id ", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabPreparationNotes.Open "SELECT *  FROM TabPreparationNotes order by  NoteDate", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
 
    dbTabMotherSolution.Open "SELECT *  FROM TabMotherSolution order by  ID", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
 
    dbTabCode.Open "SELECT *  FROM TabCode ORDER BY Code", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabCodeClassification.Open "SELECT *  FROM TabCodeClassification order by Code ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabFrasiH.Open "SELECT *  FROM TabFrasiH order by Code ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabMR.Open "SELECT *  FROM TabMR order by Code ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabMRrevisionHistory.Open "SELECT *  FROM TabMRrevisionHistory order by Code ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabMRWarehouse.Open "SELECT *  FROM TabMRWarehouse order by Code", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabSupplier.Open "SELECT *  FROM TabFrasiH order by TabSupplier ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText

    dbTabLocation.Open "SELECT *  FROM TabLocation order by Code ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText

    dbTabPipette.Open "SELECT *  FROM TabPipette order by ID ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText

 
ERR_END:

    On Error GoTo 0
    Exit Sub
ERR_OPEN:
    Debug.Print Err.Description
    If InStr(Err.Description, "RevisionHistory") > 0 Then
        If AddTable("TabRecipeRevisionHistory") Then
    
            Call m_CreateArchivio(dbPath, MydbName)
            Resume ERR_END
        End If
    End If
    If InStr(Err.Description, "TabPreparationNotes") > 0 Then
        
        If AddTable("TabPreparationNotes") Then
    
            Call m_CreateArchivio(dbPath, MydbName)
            Resume ERR_END
        End If
    End If
    If InStr(Err.Description, "TabSTDPreparationNotes") > 0 Then
        If AddTable("TabSTDPreparationNotes") Then
    
            Call m_CreateArchivio(dbPath, MydbName)
            Resume ERR_END
        End If
    End If
    
   ' MsgBox err.Description
    Resume Next
   
    
End Sub

Public Function m_CreateArchivio(ByRef t_path As String, ByVal A_NAME As String, _
            Optional ByRef T_TEXT As String, Optional ByRef r_c As Boolean, _
            Optional ByRef maxRecord As Integer, Optional Index As Integer = 0, Optional ByVal OnlyCodeDb As Boolean) As Boolean
    Dim rc As Boolean
    rc = True ' default
   On Error GoTo ERR_CREATE_OBJECT
   
   
  If OnlyCodeDb Then GoTo cont
    

        '   /////////////////////////////////////////////////////
                If dbChemicalMR.State Then dbChemicalMR.Close
        '   /////////////////////////////////////////////////////
        
        
        
        
        
        dbChemicalMR.CursorLocation = adUseServer
        dbChemicalMR.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & dbName
        
        
        
         dbChemicalMR.Execute ("ALTER TABLE TabAcquisition ADD MNP varchar(100)WITH COMPRESSION")
         dbChemicalMR.Execute ("ALTER TABLE TabAcquisition ADD ExpMR varchar(100)WITH COMPRESSION")
        
         dbChemicalMR.Execute ("ALTER TABLE TabAcquisition ADD STD_ID varchar(100)WITH COMPRESSION")
         
         dbChemicalMR.Execute ("ALTER TABLE TabAcquisition ADD ScaleID varchar(100)WITH COMPRESSION")
         dbChemicalMR.Execute ("ALTER TABLE TabAcquisition ADD GlassWareID varchar(100)WITH COMPRESSION")
         dbChemicalMR.Execute ("ALTER TABLE TabAcquisition ADD bManuale BIT")
         dbChemicalMR.Execute ("ALTER TABLE TabPreparation ADD bManuale BIT")
         
         
        
         dbChemicalMR.Execute ("ALTER TABLE TabMRWarehouse ADD U varchar(100)WITH COMPRESSION")

        
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
     

        dbCode.Execute ("ALTER TABLE TabCode ALTER STDNote varchar(255)WITH COMPRESSION")
      
        Call OpenTab


END_FN:
    On Error GoTo 0
   ' If rc And Index = 0 Then SaveSetting App.Title, "ARCHIVIO", "PATH", t_path
    m_CreateArchivio = rc
    Exit Function
ERR_CREATE_OBJECT:
    'MsgBox Err.Description
    Debug.Print Err.Description
    If InStr(Err.Description, "esistente") > 0 Then Resume Next
    If InStr(Err.Description, "exists") > 0 Then Resume Next
    If InStr(Err.Description, "sharing") > 0 Then Resume Next
    If InStr(Err.Description, UCase("condivisione")) > 0 Then Resume Next
    If InStr(UCase(Err.Description), UCase("Nessun")) > 0 Then Resume Next
    If InStr(UCase(Err.Description), UCase("field")) > 0 Then Resume Next
    If InStr(UCase(Err.Description), UCase("blocchi")) > 0 Then Resume Next
    
    
    rc = False
    r_c = False
    MsgBox Err.Description
    T_TEXT = Err.Description
    Select Case Err.NUMBER
        Case -2147467259
            'MsgBox err.NUMBER & vbCrLf & err.Description
           ' Call SearchArchivio(t_path, A_NAME)
            
           ' rc = True
    End Select
    Resume END_FN
End Function


Public Function CheckFilePath(FileName As String, PATH As String, ByRef Name As String, Optional ByRef old_name As String, Optional USER_ESTENSIONE As String = ".mdb", Optional bAvvisa As Boolean = True) As Boolean
    Dim a
    Dim n_String As String
    Dim new_name As String
    CheckFilePath = False
    If InStr(1, FileName, USER_ESTENSIONE) Then
        CheckFilePath = True
            For a = Len(FileName) To 1 Step -1
                n_String = Mid(FileName, a, 1)
                new_name = Right(FileName, Len(FileName) - a)
                Name = new_name
                If n_String = "\" Then
                    
                    If UCase(Name) <> UCase(new_name) Then
                        If bAvvisa Then
                            If MsgBox("Attenzione si è scelto un file con nome differente, procedo ugualmente?", vbInformation + vbYesNo, "Cambio percorso archivio") = vbYes Then
                                old_name = Name
                                Name = new_name
                                PATH = Left(FileName, a - 1)
                                Exit Function
                            Else
                                CheckFilePath = False
                                Exit Function
                            End If
                        Else
                            old_name = Name
                            Name = new_name
                            PATH = Left(FileName, a - 1)
                            Exit Function
                        End If
                        
                    Else
                            old_name = Name
                            Name = new_name
                            PATH = Left(FileName, a - 1)
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
    szFilename = DialogFile(F_MAIN.hwnd, 1, "Open", A_NAME, "Database Access" & Chr(0) & "*.mdb" & Chr(0) & ("Tutti i files") & Chr(0) & "*.*", "", "mdb")
    If InStr(1, szFilename, ".mdb") Then
        Call SplitPathFile(szFilename, t_path, A_NAME)
        
       ' t_path = Left(szFilename, Len(szFilename) - Len(A_NAME))
        
        SearchPathArchivio = True
    End If
    If szFilename = "" Then SearchPathArchivio = False
End Function



Public Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    ' test the directory attribute
    DirExists = GetAttr(DirName) And vbDirectory
    'MsgBox DirExists
ErrorHandler:
    ' if an error occurs, this function returns False
End Function
Public Function SettSavePath(ByVal PATH As String) As Boolean
    Dim rc As Boolean
    On Error GoTo ERR_SETT
    rc = True
    If DirExists(PATH) Then
    
    Else
        '----------------------------
        ' non esiste!!!
        '----------------------------
        MakePath (PATH)
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

