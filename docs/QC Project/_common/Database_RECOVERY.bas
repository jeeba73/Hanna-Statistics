Attribute VB_Name = "Database_RECOVERY_01_Rfp"
Option Explicit

Private SettingName As String
Private bIfDataPath As Boolean

Public Function RecoveryDatabaseFromFile()
Dim TotalsCount As Integer
TotalsCount = 0
If F_MsgBox.DoShow("Recovery database from files?") Then
    
    RecoveryAllReport
 
   
End If


End Function

Public Function RecoveryAllReport()
Dim TotalsCount As Integer
TotalsCount = 0
 
    Call RecoveryRfpFilesToDatabase(TotalsCount)


End Function


Private Function RecoveryRfpFilesToDatabase(ByRef TotalsCount As Integer)



        Dim rc As Boolean
        Dim Path As String
        rc = False
        Dim FSO As New Scripting.FileSystemObject
        
        Dim Cartella As Folder
        Dim FileGenerico As File
        
opened:
         
        USER_PATH = USER_TEMP_PATH
        bIfDataPath = IIf(USER_PATH = USER_DATA_PATH, True, False)
        GoTo cont
        
closed:

        USER_PATH = USER_DATA_PATH
        bIfDataPath = IIf(USER_PATH = USER_DATA_PATH, True, False)
                
        
cont:
      
        Path = USER_PATH
        Set Cartella = FSO.GetFolder(Path)
         
        For Each FileGenerico In Cartella.Files
        
            SettingName = FileGenerico.Name
            
            rc = ReportRecoveryGetSetting(SettingName, TotalsCount)
          
        Next
        
        If USER_PATH = USER_TEMP_PATH Then GoTo closed


    dbTabReport.Close
    dbTabReport.Open "SELECT *  FROM TabReport order by id ", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
 

End Function




Public Function ReportRecoveryGetSetting(ByVal SettName As String, ByRef TotalsCount As Integer) As Boolean
Dim rc As Boolean
Dim RecipeCount As Integer
Dim HannaCodesCount As Integer
Dim PackagingCount As Integer

On Error GoTo ERR_SAVE

    SettingName = SettName
    rc = True
   
 TotalsCount = 0
    If FileExists(USER_PATH & SettingName) = False Then
    
        rc = False
        GoTo ERR_END:
        
    End If
    
   
    CloseSettingDataFile
    
    
    With dbTabReport
        .filter = ""
        .filter = "NomeFile ='" & SettingName & "'"
        rc = .EOF
        If .EOF = False Then GoTo ERR_END:
        
        .AddNew
        
        TotalsCount = TotalsCount + 1
        !Lot = GetSettingData(SettingName, "Information QC", "Text10", "")
        !code = GetSettingData(SettingName, "Information QC", "Text11", "")
        !Description = GetSettingData(SettingName, "Information QC", "Text12", "")
        !Exp = GetSettingData(SettingName, "Information QC", "Text13", "")
        !PREPWK = GetSettingData(SettingName, "Information QC", "Text121", "")
        !Line = GetSettingData(SettingName, "Information QC", "Text14", "")
        !StartDate = GetSettingData(SettingName, "Information QC", "Modification Date", Now())
        
        !TestNumber = GetSettingData(SettingName, "Reading QC", "Grd2 Rows", 1) - 1
        !Recipe = GetSettingData(SettingName, "Information QC", "Text15", "")
        !RangeMin = GetSettingData(SettingName, "Information QC", "Text19", "")
        !RangeMax = GetSettingData(SettingName, "Information QC", "Text110", "")
        !Operator = MyOperatore.Name
        !Note = GetSettingData(SettingName, "Information QC", "Text130", "")
        !Department = GetSettingData(SettingName, "Information QC", "Text10031", "")
        !Visible = True
        !Nomefile = SettingName
        
        
        If GetSettingData(SettingName, "Close QC", "Date", "") <> "" Then
        
            !Finished = True
            !Evaluation = True


        End If


       
        .Update
        
        
    End With
    
    
    
    
 

ERR_END:
     On Error GoTo 0
     CloseSettingDataFile
     ReportRecoveryGetSetting = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function


Public Function ImportDataFromFile() As Boolean
Dim rc As Boolean
Dim SettingName As String
Dim szFilename As String
On Error GoTo ERR_SAVE

    rc = False
    If SearchFile(szFilename, SettingName) Then
    
        FileCopy szFilename, USER_PATH & SettingName
        
        Call ReportRecoveryGetSetting(SettingName, 0)
        
        rc = True
    End If

ERR_END:
     On Error GoTo 0
     CloseSettingDataFile
     If rc Then
        PopupMessage 2, "Data has been imported correctly", , , SettingName
     End If
     ImportDataFromFile = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function
Private Function SearchFile(ByRef szFilename As String, ByRef File As String) As Boolean

    
    Dim m_path As String
    
    SearchFile = True
    
    szFilename = DialogFile(F_MAIN.hWnd, 1, "Open", "", "QC Temp file" & Chr(0) & "*.qc" & Chr(0) & "Tutti i files" & Chr(0) & "*.*", "", "qc")
    If szFilename = "" Then
        SearchFile = False
        Exit Function
    
    End If
    
    DoEvents
    
    Call SplitPathFile(szFilename, m_path, File)
   
    
    
    
   
End Function




