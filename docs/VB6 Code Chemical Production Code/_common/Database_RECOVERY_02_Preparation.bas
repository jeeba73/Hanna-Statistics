Attribute VB_Name = "Database_RECOVERY_02_Preparation"
Option Explicit
Private SettingName As String
Private bIfDataPath As Boolean
Private uPreparation As RecipeForProduction
Private Path As String

Public Function RecoveryPreparationFilesToDatabase()



        Dim rc As Boolean
        Dim Path As String
        rc = False
        Dim FSO As New Scripting.FileSystemObject
        
        Dim Cartella As Folder
        Dim FileGenerico As file
        
opened:
         
        USER_PATH = USER_PREPARATION_PATH
        
        GoTo cont
        
closed:

        USER_PATH = USER_PREPARATION_PATH & "data\"
      
        
cont:
        bIfDataPath = IIf(USER_PATH = USER_PREPARATION_PATH & "data\", True, False)
      
        Path = USER_PATH
        Set Cartella = FSO.GetFolder(Path)
         
        For Each FileGenerico In Cartella.Files
        
            SettingName = FileGenerico.Name
            
            rc = PreparationRecoveryGetSetting(uPreparation, SettingName)
            If rc Then
                Dim MyID As Long
                rc = SavePreparationInDatabase(MyID)
                
                If rc Then
                    Call AggiornaTabPreparation(MyID, uPreparation)
                    Call AggiornaTabAcquisition
                    
                End If
                
                
             
            
            End If
        Next
        
        If USER_PATH = USER_PREPARATION_PATH Then GoTo closed


    dbTabPreparation.Close
    dbTabPreparation.Open "SELECT *  FROM TabPreparation order by id ", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
    dbTabAcquisition.Close
    dbTabAcquisition.Open "SELECT *  FROM TabAcquisition order by index ", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
    

End Function

Private Function AggiornaTabAcquisition()
Dim i As Integer
Dim userAcquisition As PrepAcquisition
Dim userAcquisitionClean As PrepAcquisition

Dim AcquisitionCount As Integer
    
    AcquisitionCount = uPreparation.Recipes(1).AcquisitionCount
    
    
    For i = 1 To AcquisitionCount
        userAcquisition = userAcquisitionClean
        userAcquisition = uPreparation.Recipes(1).Acquisitions(i)
        
        With dbTabAcquisition
            .filter = "AcquisitionTime='" & userAcquisition.AcquisitionTime & "'"
            If .EOF Then
            
                .AddNew
                !AcquisitionTime = userAcquisition.AcquisitionTime
                !Code = userAcquisition.PrepBarcode.Code
                !ChemicalName = userAcquisition.PrepBarcode.ChemicalName
                !Cas = userAcquisition.PrepBarcode.Cas
                !Manufacturer = userAcquisition.PrepBarcode.Manufacturer
                !ManufacturerCode = userAcquisition.PrepBarcode.ManufacturerCode
                !ManufacturerLot = userAcquisition.PrepBarcode.ManufacturerLot
                !DeliveryDate = userAcquisition.PrepBarcode.DeliveryDate
                !QtyDelivered = userAcquisition.PrepBarcode.QtyDelivered
                !Package = userAcquisition.PrepBarcode.Package
                !WeekDelPackageNumber = userAcquisition.PrepBarcode.WeekDelPackageNumber
                !Index = userAcquisition.Index
                !ActualWeight = userAcquisition.ActualWeight
                !bRecalculation = userAcquisition.bRecalculation
                !bRecipeComponent = userAcquisition.bRecipeComponent
                !bFromBarcode = userAcquisition.bFromBarcode
                !Note = userAcquisition.Note
                !Operator = userAcquisition.Operator
                !RecipeCode = uPreparation.Recipes(1).Code
                !PrepWeek = uPreparation.PrepWeek
                !NumberPrepWeek = uPreparation.numPrepWeek
                !FileName = SettingName
                '!PreparationID = PreparationID
                !ExpDate = userAcquisition.ExpDate
                .Update
            
            End If
        
        End With
    
    Next


End Function
Private Function SavePreparationInDatabase(ByRef MyID As Long) As Boolean
Dim RecipeName As String
Dim rc As Boolean

SavePreparationInDatabase = True

RecipeName = uPreparation.Recipes(1).Code
    With dbTabPreparation
        .filter = ""
        .filter = "FileName ='" & SettingName & "'"
        If .EOF Then
            .AddNew
        Else
            ' esiste già!
            SavePreparationInDatabase = False
            Exit Function
        End If
        MyID = !ID
        !Line = uPreparation.Recipes(1).Line
        !Description = uPreparation.Recipes(1).Description
        !Recipe = Trim(RecipeName)
        !PlanningReference = uPreparation.PlanningReference
        !DataRecipe = uPreparation.DateRecipe
        !QtyToProduce = uPreparation.Recipes(1).TotalWeightKg
        !RecipeWeek = uPreparation.PrepWeek
        !PlannedPreparation = uPreparation.PlannedPrepWeek
        !Operator = uPreparation.OperatorPrep
        !bClosed = bIfDataPath
        !Note = uPreparation.Note
        !FileName = SettingName
        !RfpFileName = uPreparation.fileNameRecForProd
        !bIsMix = IfRecipeIsMixString(RecipeName)
        .Update
    End With
    
    
End Function




Public Function PreparationRecoveryGetSetting(ByRef iPreparation As RecipeForProduction, ByVal SettName As String) As Boolean
Dim rc As Boolean
Dim RecipeCount As Integer
Dim HannaCodesCount As Integer
Dim TotalsCount As Integer
Dim PackagingCount As Integer
Dim RecipeCode As String
Dim RfpFileName As String
On Error GoTo ERR_SAVE

    SettingName = SettName
    rc = True
   
 
    If FileExists(USER_PATH & SettingName) = False Then
    
        rc = False
        GoTo ERR_END:
        
    End If
    
    
    With dbTabPreparation
        .filter = ""
        .filter = "FileName ='" & SettingName & "'"
        rc = .EOF
        If .EOF = False Then GoTo ERR_END:
    End With
    
    
    
    CloseSettingDataFile
  
  
    
    With iPreparation
       
        .bOpen = GetSettingData(SettingName, "iRecipeForProduction", "bOpen", .bOpen)
        .DateRecipe = GetSettingData(SettingName, "iRecipeForProduction", "DateRecipe", .DateRecipe)
        .Note = GetSettingData(SettingName, "iRecipeForProduction", "Note", .Note)
        .PlannedPrepWeek = GetSettingData(SettingName, "iRecipeForProduction", "PlannedPrepWeek", .PlannedPrepWeek)
        
        .PreparationDate = GetSettingData(SettingName, "iRecipeForProduction", "PreparationDate", "")
        .PreparationLot = GetSettingData(SettingName, "iRecipeForProduction", "PreparationLot", "")
        .ExpDate = GetSettingData(SettingName, "iRecipeForProduction", "ExpDate", "")
        .PrepWeek = GetSettingData(SettingName, "iRecipeForProduction", "PrepWeek", .PrepWeek)
        .bAllMixes = GetSettingData(SettingName, "iRecipeForProduction", "bAllMixes", .bAllMixes)
        .PlanningReference = GetSettingData(SettingName, "iRecipeForProduction", "PlanningReference", .PlanningReference)
        .numPrepWeek = GetSettingData(SettingName, "iRecipeForProduction", "NumPrepWeek", .numPrepWeek)
        .RecipeBy = GetSettingData(SettingName, "iRecipeForProduction", "RecipeBy", .RecipeBy)
        .fileNameRecForProd = GetSettingData(SettingName, "iRecipeForProduction", "fileNameRecForProd", .fileNameRecForProd)
        .bCorrection = GetSettingData(SettingName, "iRecipeForProduction", "bCorrection", .bCorrection)
        
        .OperatorPrep = GetSettingData(SettingName, "iRecipeForProduction", "OperatorPrep", .OperatorPrep)
        .OperatorRfP = GetSettingData(SettingName, "iRecipeForProduction", "OperatorRfP", .OperatorRfP)
        

        RfpFileName = .fileNameRecForProd
        '-----------------------------------------------------------
        ' Recipes in Recipe for production
        '-----------------------------------------------------------
        
        .RecipeCount = GetSettingData(SettingName, "Recipes", "RecipeCount", 0)
        RecipeCode = GetSettingData(SettingName, "Recipes1", "Code", "")
        RecipeCount = .RecipeCount
        ReDim .Recipes(1)
        If .RecipeCount > 0 Then
            Call GetPreparationRecipesFromFile(.Recipes, RecipeCount, SettingName, RecipeCode)
        End If

        .QCCount = GetSettingData(SettingName, "QC", "Count", .QCCount)
        
        If .QCCount > 0 Then
            
            Call GetQCFromPreparationFile(.QCStatus(), .QCCount, SettingName, Path)
        
        End If
        


        If .Recipes(1).bIsMix Then GoTo ERR_END
            
        '-----------------------------------------------------------
        ' HANNA CODES in Recipe for production
        '-----------------------------------------------------------
        
        If FileExists(USER_PRODUCTION_PATH & RfpFileName) Then
            Path = USER_PRODUCTION_PATH
        ElseIf FileExists(USER_TEMP_PATH & RfpFileName) Then
          
            Path = USER_TEMP_PATH
        ElseIf FileExists(USER_DATA_PATH & RfpFileName) Then
        
            Path = USER_DATA_PATH
        End If
            
        CloseSettingDataFile
        .HannaCodesCount = GetSettingData(RfpFileName, "HannaCodes", "HannaCodesCount", 0, Path)
        
        
        If .HannaCodesCount > 0 Then
            HannaCodesCount = .HannaCodesCount
            ReDim .HannaCodes(0)
            Call GetPreparationHannaCodesFromFile(.HannaCodes, HannaCodesCount, RfpFileName, .Recipes(1).Code, Path)
            
            .HannaCodesCount = HannaCodesCount
        End If
        

        
    End With
    

ERR_END:
    On Error GoTo 0
     CloseSettingDataFile
     PreparationRecoveryGetSetting = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function
