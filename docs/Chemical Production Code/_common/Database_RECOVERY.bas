Attribute VB_Name = "Database_RECOVERY_01_Rfp"
Option Explicit

Private SettingName As String
Private bIfDataPath As Boolean
Private uRecipeForProduction As RecipeForProduction

Public Function RecoveryDatabaseFromFile()

If F_MsgBox.DoShow("Recovery database from files?") Then
    
    Call RecoveryRfpFilesToDatabase
    Call RecoveryPreparationFilesToDatabase
    Call RecoveryProductionFilesToDatabase
    
    
    PopupMessage 2, "Database recovery finished"
End If


End Function

Private Function RecoveryRfpFilesToDatabase()



        Dim rc As Boolean
        Dim Path As String
        rc = False
        Dim FSO As New Scripting.FileSystemObject
        
        Dim Cartella As Folder
        Dim FileGenerico As file
        
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
            
            rc = RfpRecoveryGetSetting(uRecipeForProduction, SettingName)
            If rc Then
                
                Call SaveRecipeForProductionInDatabase(uRecipeForProduction.Recipes)
            
            End If
        Next
        
        If USER_PATH = USER_TEMP_PATH Then GoTo closed


    dbTabReceiptForProduction.Close
    dbTabReceiptForProduction.Open "SELECT *  FROM TabReceiptForProduction order by id ", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
 

End Function


Private Function SaveRecipeForProductionInDatabase(uRecipe() As RecipeType) As Boolean
Dim rc As Boolean

' se sono in Data allora la ricerca è tra i Recipe chiusi!!
On Error GoTo SaveReceipt
    rc = True
    With dbTabReceiptForProduction
        .filter = ""
        .filter = "FileName ='" & SettingName & IIf(bIfDataPath, "' and bClosed=true", "' and bClosed=false")
        If .EOF Then
                
            .AddNew
            
        Else
        
        
        End If
        
        !Recipe = GetStrRecipe(uRecipe)
        !Description = GetStrDescriptionRecipe(uRecipe)
        !Line = GetStrLineRecipe(uRecipe)
        !PlanningReference = uRecipeForProduction.PlanningReference
        !DataRecipe = uRecipeForProduction.DateRecipe
        !RecipeWeek = uRecipeForProduction.PlannedPrepWeek
       ' !PlannedPreparation = uRecipeForProduction.pl
        !Operator = uRecipeForProduction.OperatorRfP
        !bClosed = bIfDataPath
        !Note = uRecipeForProduction.Note
        !FileName = SettingName
    
        .Update
    
    End With

ERR_END:
    On Error GoTo 0
    SaveRecipeForProductionInDatabase = rc
    Exit Function
SaveReceipt:
    rc = False
    MsgBox err.Description
    Resume Next
    
End Function


Public Function RfpRecoveryGetSetting(ByRef iRecipeForProduction As RecipeForProduction, ByVal SettName As String) As Boolean
Dim rc As Boolean
Dim RecipeCount As Integer
Dim HannaCodesCount As Integer
Dim TotalsCount As Integer
Dim PackagingCount As Integer

On Error GoTo ERR_SAVE

    SettingName = SettName
    rc = True
   
 
    If FileExists(USER_PATH & SettingName) = False Then
    
        rc = False
        GoTo ERR_END:
        
    End If
    
    
    With dbTabReceiptForProduction
        .filter = ""
        .filter = "FileName ='" & SettingName & "'"
        rc = .EOF
        If .EOF = False Then GoTo ERR_END:
    End With
    
    
    
    CloseSettingDataFile
  
  
    
    With iRecipeForProduction
       
        .bOpen = GetSettingData(SettingName, "iRecipeForProduction", "bOpen", .bOpen)
        .DateRecipe = GetSettingData(SettingName, "iRecipeForProduction", "DateRecipe", .DateRecipe)
        .Note = GetSettingData(SettingName, "iRecipeForProduction", "Note", .Note)
        .PlannedPrepWeek = GetSettingData(SettingName, "iRecipeForProduction", "PlannedPrepWeek", .PlannedPrepWeek)
        .bAllMixes = GetSettingData(SettingName, "iRecipeForProduction", "bAllMixes", .bAllMixes)
        .PlanningReference = GetSettingData(SettingName, "iRecipeForProduction", "PlanningReference", .PlanningReference)
        .numPrepWeek = GetSettingData(SettingName, "iRecipeForProduction", "NumPrepWeek", .numPrepWeek)
        .RecipeBy = GetSettingData(SettingName, "iRecipeForProduction", "RecipeBy", .RecipeBy)
        .fileNameRecForProd = GetSettingData(SettingName, "iRecipeForProduction", "fileNameRecForProd", .fileNameRecForProd)
       
    
        '-----------------------------------------------------------
        ' Recipes in Recipe for production
        '-----------------------------------------------------------
        
        .RecipeCount = GetSettingData(SettingName, "Recipes", "RecipeCount", 0)
        
        RecipeCount = .RecipeCount
        ReDim .Recipes(RecipeCount)
        If .RecipeCount > 0 Then
            Call GetRecipesFromFile(.Recipes, RecipeCount, SettingName)
        End If
        
    End With
    

ERR_END:
    On Error GoTo 0
     CloseSettingDataFile
     RfpRecoveryGetSetting = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function
