Attribute VB_Name = "Formulation_03_SaveSettingData"
Option Explicit


Private SettingName As String

Public Function ReceiptSaveSetting(ByRef iRecipeForProduction As RecipeForProduction, ByVal SettName As String) As Boolean
Dim rc As Boolean
Dim RecipeCount As Integer
Dim HannaCodesCount As Integer
Dim TotalsCount As Integer
Dim PackagingCount As Integer

On Error GoTo ERR_SAVE

    SettingName = SettName
    
    rc = True


    'USER_PATH = USER_TEMP_PATH
    
    If FileExists(USER_PATH & SettingName) Then Kill USER_PATH & SettingName
    DoEvents
    
    CloseSettingDataFile
    
    
    SaveSettingData SettingName, "Program", "", ""
    SaveSettingData SettingName, App.EXEName, "", ""
    SaveSettingData SettingName, "Program", "Release", App.Major & "." & App.Minor & "." & App.Revision
    SaveSettingData SettingName, "Recipe For Production", "Create Recipe", ""
    SaveSettingData SettingName, "Recipe For Production", "Date", Now()
    SaveSettingData SettingName, "WorkStation", "Department", MyWorkStation.Department
    SaveSettingData SettingName, "WorkStation", "Description", MyWorkStation.Description
    SaveSettingData SettingName, "WorkStation", "LineLeader", MyWorkStation.LineLeader
    SaveSettingData SettingName, "WorkStation", "Workstation", MyWorkStation.Workstation

    
    
    With iRecipeForProduction
        SaveSettingData SettingName, "iRecipeForProduction", "bOpen", .bOpen
        SaveSettingData SettingName, "iRecipeForProduction", "DateRecipe", .DateRecipe
        SaveSettingData SettingName, "iRecipeForProduction", "Note", .Note
        SaveSettingData SettingName, "iRecipeForProduction", "PlannedPrepWeek", .PlannedPrepWeek
        SaveSettingData SettingName, "iRecipeForProduction", "bAllMixes", .bAllMixes
        SaveSettingData SettingName, "iRecipeForProduction", "PlanningReference", .PlanningReference
        SaveSettingData SettingName, "iRecipeForProduction", "NumPrepWeek", .numPrepWeek
        SaveSettingData SettingName, "iRecipeForProduction", "RecipeBy", .RecipeBy
        SaveSettingData SettingName, "iRecipeForProduction", "fileNameRecForProd", SettingName

        '-----------------------------------------------------------
        ' Packaging
        '-----------------------------------------------------------
        If .PackagingCount > 0 Then
            PackagingCount = .PackagingCount
            SaveSettingData SettingName, "Packaging", "PackagingCount", .PackagingCount
            Call SetPackagingInFile(.Packaging, PackagingCount)
        End If
                  
                  
        '-----------------------------------------------------------
        ' Totals
        '-----------------------------------------------------------
        If .TotalCount > 0 Then
            TotalsCount = .TotalCount
            SaveSettingData SettingName, "Totals Grid", "TotalCount", .TotalCount
            Call SetTotalsInFile(.TotalGrid, TotalsCount)
        End If
            
            
        
        
        '-----------------------------------------------------------
        ' HANNA CODES in Recipe for production
        '-----------------------------------------------------------
        If .HannaCodesCount > 0 Then
            HannaCodesCount = .HannaCodesCount
            SaveSettingData SettingName, "HannaCodes", "HannaCodesCount", .HannaCodesCount
            Call SetHannaCodesInFile(.HannaCodes, HannaCodesCount)
        End If
     
        '-----------------------------------------------------------
        ' Recipes
        '-----------------------------------------------------------
        
        .RecipeCount = UBound(.Recipes)
        
        SaveSettingData SettingName, "Recipes", "RecipeCount", .RecipeCount
        
        RecipeCount = .RecipeCount

        Call SetRecipesInFile(.Recipes, RecipeCount)


    End With

ERR_END:
    On Error GoTo 0
     CloseSettingDataFile
     ReceiptSaveSetting = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function


Public Function SetTotalsInFile(ByRef Totals() As Totals, ByVal TotalsCount As Integer) As Boolean

Dim i As Integer

CloseSettingDataFile

For i = 1 To TotalsCount
    With Totals(i)
        SaveSettingData SettingName, "Totals Grid" & i, "CkMax", .CkMax
        SaveSettingData SettingName, "Totals Grid" & i, "CkMin", .CkMin
        SaveSettingData SettingName, "Totals Grid" & i, "Description", .Description
        SaveSettingData SettingName, "Totals Grid" & i, "Max", .Max
        SaveSettingData SettingName, "Totals Grid" & i, "Min", .Min
        SaveSettingData SettingName, "Totals Grid" & i, "Minpcs", .Minpcs
        SaveSettingData SettingName, "Totals Grid" & i, "bMix", .bMix
        SaveSettingData SettingName, "Totals Grid" & i, "Multiple", .Multiple
        SaveSettingData SettingName, "Totals Grid" & i, "Recipe", .Recipe
        SaveSettingData SettingName, "Totals Grid" & i, "TotalMultiple", .TotalMultiple
        SaveSettingData SettingName, "Totals Grid" & i, "TotalWeighKg", .TotalWeighKg
        SaveSettingData SettingName, "Totals Grid" & i, "TotalWeighL", .TotalWeighL

        

    End With
Next

CloseSettingDataFile

End Function
Public Function SetRecipesInFile(ByRef Recipes() As RecipeType, ByVal RecipeCount As Integer) As Boolean

Dim i As Integer
Dim HannaCodesCount As Integer
Dim MaterialRequisitionCount As Integer
Dim RmxRecipeCount As Integer

For i = 1 To RecipeCount


    
    
    
    With Recipes(i)
        SaveSettingData SettingName, "RecipeIndex", .Code, i
        SaveSettingData SettingName, "Recipes" & i, "Code", .Code
        SaveSettingData SettingName, "Recipes" & i, "bHaveMixes", .bHaveMixes
        SaveSettingData SettingName, "Recipes" & i, "bUmMassa", .bUmMassa
        SaveSettingData SettingName, "Recipes" & i, "bUpdated", .bUpdated
        SaveSettingData SettingName, "Recipes" & i, "bIsMix", .bIsMix
        SaveSettingData SettingName, "Recipes" & i, "PreparationLotMix", .PreparationLotMix
        SaveSettingData SettingName, "Recipes" & i, "Machine", .Machine
        SaveSettingData SettingName, "Recipes" & i, "bTestLot", .bTestLot
     
        
        
        
        SaveSettingData SettingName, "Recipes" & i, "Density", .Density
        SaveSettingData SettingName, "Recipes" & i, "Description", .Description
        SaveSettingData SettingName, "Recipes" & i, "Exp", .Exp
        SaveSettingData SettingName, "Recipes" & i, "ID", i
        SaveSettingData SettingName, "Recipes" & i, "Line", .Line
        
        SaveSettingData SettingName, "Recipes" & i, "MaxQty", .MaxQty
        SaveSettingData SettingName, "Recipes" & i, "MinQty", .MinQty
        SaveSettingData SettingName, "Recipes" & i, "MinQty2", .MinQty2
        SaveSettingData SettingName, "Recipes" & i, "Mix", .Mix
        
        SaveSettingData SettingName, "Recipes" & i, "bNoPreparation", .bNoPreparation
        
        
        
        SaveSettingData SettingName, "Recipes" & i, "Multiple", .Multiple
        SaveSettingData SettingName, "Recipes" & i, "MultipleMassa", .MultipleMassa
        SaveSettingData SettingName, "Recipes" & i, "MultipleToProduce", .MultipleToProduce
        'SaveSettingData SettingName, "Recipes" & i, "NoteRev", .NoteRev
        'SaveSettingData SettingName, "Recipes" & i, "Procedure", .Procedure
        'SaveSettingData SettingName, "Recipes" & i, "ProductionWay.EsttimeD", .ProductionWay.EsttimeD
        'SaveSettingData SettingName, "Recipes" & i, "ProductionWay.EstTimeH", .ProductionWay.EstTimeH
        'SaveSettingData SettingName, "Recipes" & i, "ProductionWay.Line", .ProductionWay.Line
        'SaveSettingData SettingName, "Recipes" & i, "ProductionWay.Production", .ProductionWay.Production
        'SaveSettingData SettingName, "Recipes" & i, "ProductionWay.Speed", .ProductionWay.Speed
        
        'SaveSettingData SettingName, "Recipes" & i, "Rev", .Rev
        
        SaveSettingData SettingName, "Recipes" & i, "TotalMultiple", .TotalMultiple
        SaveSettingData SettingName, "Recipes" & i, "TotalRecipe", .TotalRecipe
        SaveSettingData SettingName, "Recipes" & i, "TotalWeightKg", .TotalWeightKg
        SaveSettingData SettingName, "Recipes" & i, "TotalWeightL", .TotalWeightL
        SaveSettingData SettingName, "Recipes" & i, "UmMax", .UmMax
        SaveSettingData SettingName, "Recipes" & i, "UmMinQty", .UmMinQty
        SaveSettingData SettingName, "Recipes" & i, "bHide", .bHide
        SaveSettingData SettingName, "Recipes" & i, "UmMultiple", .UmMultiple
        SaveSettingData SettingName, "Recipes" & i, "UmTotalWeightKg", .UmTotalWeightKg
        SaveSettingData SettingName, "Recipes" & i, "UmTotalWeightL", .UmTotalWeightL

        
        
        '-----------------------------------------------------------
        ' HANNA CODES x Recipe
        '-----------------------------------------------------------
        'If .HannaCodesCount > 0 Then
        '    HannaCodesCount = .HannaCodesCount
        '    SaveSettingData SettingName, "Recipes" & i & " - HannaCode", "HannaCodesCount", .HannaCodesCount
        '    Call SetHannaCodesPerRecipeInFile(i, .HannaCodes, HannaCodesCount)
        'End If

        '-----------------------------------------------------------
        ' RmxRecipe
        '-----------------------------------------------------------
        If .RmxRecipeCount >= 0 Then
            RmxRecipeCount = .RmxRecipeCount
            SaveSettingData SettingName, "Recipes" & i & " - RmxRecipe", "RmxRecipeCount", .RmxRecipeCount
            Call SetRmxRecipeInFile(i, .RmxRecipe, RmxRecipeCount)
        End If
    End With
Next

CloseSettingDataFile


End Function

Public Function SetHannaCodesInFile(ByRef HannaCodes() As HannaCode, ByVal HannaCodesCount As Integer)
Dim i As Integer
    For i = 1 To HannaCodesCount
        
        With HannaCodes(i)
            SaveSettingData SettingName, "HannaCode" & i, "bHide", .bHide
            SaveSettingData SettingName, "HannaCode" & i, "Code", .Code
            'SaveSettingData SettingName, "HannaCode" & i, "Density", .Density
            'SaveSettingData SettingName, "HannaCode" & i, "Exp", .Exp
            SaveSettingData SettingName, "HannaCode" & i, "ID", .ID
            SaveSettingData SettingName, "HannaCode" & i, "LastLot", .LastLot
            SaveSettingData SettingName, "HannaCode" & i, "Line", .Line
            'SaveSettingData SettingName, "HannaCode" & i, "LoadInPrint", '.LoadInPrint
            SaveSettingData SettingName, "HannaCode" & i, "MaxQty", .MaxQty
            SaveSettingData SettingName, "HannaCode" & i, "MinQty", .MinQty
            SaveSettingData SettingName, "HannaCode" & i, "Mix1", .Mix1
            SaveSettingData SettingName, "HannaCode" & i, "Mix2", .Mix2
           ' SaveSettingData SettingName, "HannaCode" & i, "Procedure", .Procedure
           ' SaveSettingData SettingName, "HannaCode" & i, "ProcedureRev", .ProcedureRev
            SaveSettingData SettingName, "HannaCode" & i, "ProductName", .ProductName
            SaveSettingData SettingName, "HannaCode" & i, "Qty", .Qty
            SaveSettingData SettingName, "HannaCode" & i, "QtyToProduce", .QtyToProduce
            SaveSettingData SettingName, "HannaCode" & i, "Recipe", .Recipe
            'SaveSettingData SettingName, "HannaCode" & i, "Std", .Std
            SaveSettingData SettingName, "HannaCode" & i, "Um", .Um
            'SaveSettingData SettingName, "HannaCode" & i, "UncertantlyFromCoA", .UncertantlyFromCoA
            SaveSettingData SettingName, "HannaCode" & i, "LotNumber", .LotNumber
        End With

    Next
    
    CloseSettingDataFile
    
End Function
Public Function SetHannaCodesPerRecipeInFile(ByVal t As Integer, ByRef HannaCodes() As HannaCode, ByVal HannaCodesCount As Integer)
Dim i As Integer
    For i = 1 To HannaCodesCount
        
        With HannaCodes(i)
        
            SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "Code", .Code
            'SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "Density", .Density
            'SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "Exp", .Exp
            SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "ID", .ID
            SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "LastLot", .LastLot
            SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "Line", .Line
            'SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "LoadInPrint", '.LoadInPrint
            SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "MaxQty", .MaxQty
            SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "MinQty", .MinQty
            SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "Mix1", .Mix1
            SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "Code", .Mix2
            'SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "Procedure", .Procedure
            'SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "ProcedureRev", .ProcedureRev
            'SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "ProductName", .ProductName
            SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "Qty", .Qty
            SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "Recipe", .Recipe
            'SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "Std", .Std
            SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "Um", .Um
            'SaveSettingData SettingName, "Recipes" & t & " - HannaCode" & i, "UncertantlyFromCoA", .UncertantlyFromCoA
        
        End With

    Next
End Function

Public Function SetPackagingInFile(ByRef Packaging() As ProdWay, ByVal PackagingCount As Integer)
Dim i As Integer
    For i = 1 To PackagingCount
        
        With Packaging(i)
        
            SaveSettingData SettingName, "Packaging" & i, "EsttimeD", .EsttimeD
            SaveSettingData SettingName, "Packaging" & i, "EstTimeH", .EstTimeH
            SaveSettingData SettingName, "Packaging" & i, "Line", .Line
            SaveSettingData SettingName, "Packaging" & i, "Head", .Head
            SaveSettingData SettingName, "Packaging" & i, "Production", .Production
            SaveSettingData SettingName, "Packaging" & i, "Recipe", .Recipe
            SaveSettingData SettingName, "Packaging" & i, "Speed", .Speed
        
        End With

    Next
End Function
Public Function SetRmxRecipeInFile(ByVal t As Integer, ByRef RmxRecipe() As RmxRecipe, ByVal RmxRecipeCount As Integer)
Dim i As Integer
    For i = 0 To RmxRecipeCount
        
        With RmxRecipe(i)
        
            
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "bMix", .bMix
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "bUmMassa", .bUmMassa
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Cas", .Cas
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "CHCode", .CHCode
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Density", .Density
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Description", .Description
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "ID", .ID
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "MaxQty", .MaxQty
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "MinQty", .MinQty
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "MinQty2", .MinQty2
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Multiple", .Multiple
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "MultipleInCell", .MultipleInCell
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "MultipleMassa", .MultipleMassa
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "MultipleToProduce", .MultipleToProduce
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Note", .Note
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Perc", .Perc
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Qty", .Qty
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "CriticalRM", .CriticalRM
            'SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.ChemicalReactionLiquid", .Specifications.ChemicalReactionLiquid
            'SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Classification", .Specifications.Classification
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Location", .Specifications.Location
            'SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.ManufacturerCode", .Specifications.ManufacturerCode
            'SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.ManufacturerName", .Specifications.ManufacturerName
            'SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Pictograms", .Specifications.Pictograms
            'SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.QtyToProduce", .Specifications.QtyToProduce
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.SpecifiedLocation", .Specifications.SpecifiedLocation
            'SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Tollerance", .Specifications.Tollerance
            'SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.UmQty", .Specifications.UmQty

            
            
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "RecipeCode", .RecipeCode
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "TolerancePerc", .TolerancePerc
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "TheoreticalWeight", .TheoreticalWeight
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "TotalMultiple", .TotalMultiple
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "TotalWeightKg", .TotalWeightKg
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "TotalWeightL", .TotalWeightL
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Um", .Um
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmMax", .UmMax
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmMinQty", .UmMinQty
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmMultiple", .UmMultiple
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmTheoreticalWeight", .UmTheoreticalWeight

        End With

    Next
    
    CloseSettingDataFile
End Function


