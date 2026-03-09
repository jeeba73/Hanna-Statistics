Attribute VB_Name = "Formulation_04_GettSettingdata"
Option Explicit


Private SettingName As String

Public Function ReceiptGetSetting(ByRef iRecipeForSTDPreparation As RecipeForSTDPreparation, ByVal SettName As String) As Boolean
Dim rc As Boolean
Dim RecipeCount As Integer
Dim HannaCodesCount As Integer
Dim TotalsCount As Integer
Dim PackagingCount As Integer

On Error GoTo ERR_SAVE

    SettingName = SettName
    rc = True
   
   If USER_PATH = "" Then USER_PATH = USER_TEMP_PATH
  
    
    If FileExists(USER_PATH & SettingName) = False Then
    
        rc = False
        GoTo ERR_END:
        
    End If
    
    
    CloseSettingDataFile
  
    
    With iRecipeForSTDPreparation
       
        .bOpen = GetSettingData(SettingName, "iRecipeForSTDPreparation", "bOpen", .bOpen)
        .DateRecipe = GetSettingData(SettingName, "iRecipeForSTDPreparation", "DateRecipe", .DateRecipe)
        .Note = GetSettingData(SettingName, "iRecipeForSTDPreparation", "Note", .Note)
        .PlannedPrepWeek = GetSettingData(SettingName, "iRecipeForSTDPreparation", "PlannedPrepWeek", .PlannedPrepWeek)
        .bAllMixes = GetSettingData(SettingName, "iRecipeForSTDPreparation", "bAllMixes", .bAllMixes)
        .PlanningReference = GetSettingData(SettingName, "iRecipeForSTDPreparation", "PlanningReference", .PlanningReference)
        .numPrepWeek = GetSettingData(SettingName, "iRecipeForSTDPreparation", "NumPrepWeek", .numPrepWeek)
        .RecipeBy = GetSettingData(SettingName, "iRecipeForSTDPreparation", "RecipeBy", .RecipeBy)
        .fileNameRecForProd = GetSettingData(SettingName, "iRecipeForSTDPreparation", "fileNameRecForProd", .fileNameRecForProd)


        
        
        
        '-----------------------------------------------------------
        ' Packaging
        '-----------------------------------------------------------
        .PackagingCount = GetSettingData(SettingName, "Packaging", "PackagingCount", 0)
        If .PackagingCount > 0 Then
            PackagingCount = .PackagingCount
            
            Call GetPackagingInFile(.Packaging, PackagingCount, SettingName)
        End If
                       
                       
        
                
                
        '-----------------------------------------------------------
        ' Totals
        '-----------------------------------------------------------
        .TotalCount = GetSettingData(SettingName, "Totals Grid", "TotalCount", 0)
        
        If .TotalCount > 0 Then
            ReDim .TotalGrid(.TotalCount)
            TotalsCount = .TotalCount
            
            Call GetTotalsFormFile(.TotalGrid, TotalsCount, SettingName)
        End If
                    
        
        
        '-----------------------------------------------------------
        ' HANNA CODES in Recipe for STDPreparation
        '-----------------------------------------------------------
        
        .HannaCodesCount = GetSettingData(SettingName, "HannaCodes", "HannaCodesCount", 0)
        ReDim .HannaCodes(.HannaCodesCount)
        If .HannaCodesCount > 0 Then
            HannaCodesCount = .HannaCodesCount
            
            Call GetHannaCodesFromFile(.HannaCodes, HannaCodesCount, SettingName)
        End If
        
        '-----------------------------------------------------------
        ' Recipes in Recipe for STDPreparation
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
     ReceiptGetSetting = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function



Public Function GetTotalsFormFile(ByRef Totals() As Totals, ByVal TotalsCount As Integer, ByVal SettingName As String) As Boolean

Dim i As Integer

CloseSettingDataFile

For i = 1 To TotalsCount
    With Totals(i)
        .CkMax = GetSettingData(SettingName, "Totals Grid" & i, "CkMax", .CkMax)
        .CkMin = GetSettingData(SettingName, "Totals Grid" & i, "CkMin", .CkMin)
        .Description = GetSettingData(SettingName, "Totals Grid" & i, "Description", .Description)
        .Max = GetSettingData(SettingName, "Totals Grid" & i, "Max", .Max)
        .Min = GetSettingData(SettingName, "Totals Grid" & i, "Min", .Min)
        .Minpcs = GetSettingData(SettingName, "Totals Grid" & i, "Minpcs", .Minpcs)
        .bMix = GetSettingData(SettingName, "Totals Grid" & i, "bMix", .bMix)
        .Multiple = GetSettingData(SettingName, "Totals Grid" & i, "Multiple", .Multiple)
        .Recipe = GetSettingData(SettingName, "Totals Grid" & i, "Recipe", .Recipe)
        .TotalMultiple = GetSettingData(SettingName, "Totals Grid" & i, "TotalMultiple", .TotalMultiple)
        .TotalWeighKg = GetSettingData(SettingName, "Totals Grid" & i, "TotalWeighKg", .TotalWeighKg)
        .TotalWeighL = GetSettingData(SettingName, "Totals Grid" & i, "TotalWeighL", .TotalWeighL)
    End With
Next

CloseSettingDataFile

End Function

Public Function GetRecipesFromFile(ByRef Recipes() As RecipeType, ByRef RecipeCount As Integer, ByVal SettingName As String, Optional ByVal bSTDPreparation As Boolean) As Boolean
    
    Dim i As Integer
    Dim rc As Boolean
    Dim Count As Integer
    Count = 1
    For i = 1 To RecipeCount


        rc = GetSingleRecipeFromFile(Recipes(Count), SettingName, i, bSTDPreparation)
        If bSTDPreparation Then
           If rc Then
            Count = Count + 1
            RecipeCount = Count - 1
           End If
        Else
            Count = Count + 1
        End If
        
    Next
  
End Function


Public Function GetSingleRecipeFromFile(ByRef Recipe As RecipeType, ByVal SettingName As String, ByVal i As Integer, Optional ByVal bSTDPreparation As Boolean) As Boolean


Dim HannaCodesCount As Integer
Dim MaterialRequisitionCount As Integer
Dim RmxRecipeCount As Integer
Dim rc As Boolean
On Error GoTo ERR_GET:

    rc = True

    With Recipe
        .bHide = GetSettingData(SettingName, "Recipes" & i, "bHide", False)
        If bSTDPreparation Then
            If .bHide Then
                rc = False
                GoTo ERR_END
            End If
        End If
        .Code = GetSettingData(SettingName, "Recipes" & i, "Code", .Code)
        
        .bHaveMixes = GetSettingData(SettingName, "Recipes" & i, "bHaveMixes", .bHaveMixes)
        
        .bUmMassa = GetSettingData(SettingName, "Recipes" & i, "bUmMassa", .bUmMassa)
        .bUpdated = GetSettingData(SettingName, "Recipes" & i, "bUpdated", False)
        .bIsMix = GetSettingData(SettingName, "Recipes" & i, "bIsMix", False)
        .Density = GetSettingData(SettingName, "Recipes" & i, "Density", .Density)
        .Description = GetSettingData(SettingName, "Recipes" & i, "Description", .Description)
        .Exp = GetSettingData(SettingName, "Recipes" & i, "Exp", .Exp)
        .ID = GetSettingData(SettingName, "Recipes" & i, "ID", 0)
        .Line = GetSettingData(SettingName, "Recipes" & i, "Line", .Line)
        
        .MaxQty = GetSettingData(SettingName, "Recipes" & i, "MaxQty", .MaxQty)
        .MinQty = GetSettingData(SettingName, "Recipes" & i, "MinQty", .MinQty)
        .MinQty2 = GetSettingData(SettingName, "Recipes" & i, "MinQty2", .MinQty2)
        .bNoPreparation = GetSettingData(SettingName, "Recipes" & i, "bNoPreparation", .bNoPreparation)
        
        .Mix = GetSettingData(SettingName, "Recipes" & i, "Mix", .Mix)
        
        .Multiple = GetSettingData(SettingName, "Recipes" & i, "Multiple", .Multiple)
        .MultipleMassa = GetSettingData(SettingName, "Recipes" & i, "MultipleMassa", .MultipleMassa)
        .MultipleToProduce = GetSettingData(SettingName, "Recipes" & i, "MultipleToProduce", .MultipleToProduce)
        '.NoteRev = GetSettingData(SettingName, "Recipes" & i, "NoteRev", .NoteRev)
        '.Procedure = GetSettingData(SettingName, "Recipes" & i, "Procedure", .Procedure)
        '.STDPreparationWay.EsttimeD = GetSettingData(SettingName, "Recipes" & i, "STDPreparationWay.EsttimeD", .STDPreparationWay.EsttimeD)
        '.STDPreparationWay.EstTimeH = GetSettingData(SettingName, "Recipes" & i, "STDPreparationWay.EstTimeH", .STDPreparationWay.EstTimeH)
        '.STDPreparationWay.Line = GetSettingData(SettingName, "Recipes" & i, "STDPreparationWay.Line", .STDPreparationWay.Line)
        '.STDPreparationWay.STDPreparation = GetSettingData(SettingName, "Recipes" & i, "STDPreparationWay.STDPreparation", .STDPreparationWay.STDPreparation)
        '.STDPreparationWay.Speed = GetSettingData(SettingName, "Recipes" & i, "STDPreparationWay.Speed", .STDPreparationWay.Speed)
        
        '.Rev = SaveSettingData(SettingName, "Recipes" & i, "Rev", .Rev)
        
        .TotalMultiple = GetSettingData(SettingName, "Recipes" & i, "TotalMultiple", .TotalMultiple)
        .TotalRecipe = GetSettingData(SettingName, "Recipes" & i, "TotalRecipe", .TotalRecipe)
        .TotalWeightKg = GetSettingData(SettingName, "Recipes" & i, "TotalWeightKg", .TotalWeightKg)
        .TotalWeightL = GetSettingData(SettingName, "Recipes" & i, "TotalWeightL", .TotalWeightL)
        .UmMax = GetSettingData(SettingName, "Recipes" & i, "UmMax", .UmMax)
        .UmMinQty = GetSettingData(SettingName, "Recipes" & i, "UmMinQty", .UmMinQty)
        .bHide = GetSettingData(SettingName, "Recipes" & i, "bHide", False)
        .UmMultiple = GetSettingData(SettingName, "Recipes" & i, "UmMultiple", .UmMultiple)
        .UmTotalWeightKg = GetSettingData(SettingName, "Recipes" & i, "UmTotalWeightKg", .UmTotalWeightKg)
        .UmTotalWeightL = GetSettingData(SettingName, "Recipes" & i, "UmTotalWeightL", .UmTotalWeightL)


        
        
        
        '-----------------------------------------------------------
        ' HANNA CODES x Recipe
        '-----------------------------------------------------------
       '  .HannaCodesCount = GetSettingData(SettingName, "Recipes" & i & " - HannaCode", "HannaCodesCount", 0)
        'If .HannaCodesCount > 0 Then
          '  HannaCodesCount = .HannaCodesCount
         '   ReDim .HannaCodes(HannaCodesCount)
         '   Call GetHannaCodesPerRecipeFromFile(i, .HannaCodes, HannaCodesCount, SettingName)
        'End If

        '-----------------------------------------------------------
        ' RmxRecipe
        '-----------------------------------------------------------
        .RmxRecipeCount = GetSettingData(SettingName, "Recipes" & i & " - RmxRecipe", "RmxRecipeCount", 0)
        If .RmxRecipeCount >= 0 Then
            RmxRecipeCount = .RmxRecipeCount
            ReDim .RmxRecipe(RmxRecipeCount)
            Call GetRmxRecipeFromFile(i, .RmxRecipe, RmxRecipeCount, SettingName)
        End If
    End With

    CloseSettingDataFile

ERR_END:
    On Error GoTo 0
    GetSingleRecipeFromFile = rc
    Exit Function
ERR_GET:
    MsgBox err.Description
    Resume Next


End Function

Public Function GetHannaCodesFromFile(ByRef HannaCodes() As HannaCode, ByVal HannaCodesCount As Integer, ByVal SettingName As String)
Dim i As Integer
    For i = 1 To HannaCodesCount
        
        With HannaCodes(i)
            .bHide = GetSettingData(SettingName, "HannaCode" & i, "bHide", .bHide)
            .Code = GetSettingData(SettingName, "HannaCode" & i, "Code", .Code)
            '.Density = GetSettingData(SettingName, "HannaCode" & i, "Density", .Density)
            .Exp = GetSettingData(SettingName, "HannaCode" & i, "Exp", .Exp)
            .ExpDate = GetSettingData(SettingName, "HannaCode" & i, "ExpDate", .ExpDate)
            .ID = GetSettingData(SettingName, "HannaCode" & i, "ID", .ID)
            '.LastLot = GetSettingData(SettingName, "HannaCode" & i, "LastLot", .LastLot)
            .Line = GetSettingData(SettingName, "HannaCode" & i, "Line", .Line)
            '.LoadInPrint = GetSettingData(SettingName, "HannaCode" & i, "LoadInPrint", .LoadInPrint)
            .MaxQty = GetSettingData(SettingName, "HannaCode" & i, "MaxQty", .MaxQty)
            .MinQty = GetSettingData(SettingName, "HannaCode" & i, "MinQty", .MinQty)
            .Mix1 = GetSettingData(SettingName, "HannaCode" & i, "Mix1", .Mix1)
            .Mix2 = GetSettingData(SettingName, "HannaCode" & i, "Mix2", .Mix2)
            '.Procedure = GetSettingData(SettingName, "HannaCode" & i, "Procedure", .Procedure)
            '.ProcedureRev = GetSettingData(SettingName, "HannaCode" & i, "ProcedureRev", .ProcedureRev)
            .ProductName = GetSettingData(SettingName, "HannaCode" & i, "ProductName", .ProductName)
            .Qty = GetSettingData(SettingName, "HannaCode" & i, "Qty", .Qty)
            .QtyToProduce = GetSettingData(SettingName, "HannaCode" & i, "QtyToProduce", .QtyToProduce)
            .Recipe = GetSettingData(SettingName, "HannaCode" & i, "Recipe", .Recipe)
            '.Std = GetSettingData(SettingName, "HannaCode" & i, "Std", .Std)
            .Um = GetSettingData(SettingName, "HannaCode" & i, "Um", .Um)
            '.UncertantlyFromCoA = GetSettingData(SettingName, "HannaCode" & i, "UncertantlyFromCoA", .UncertantlyFromCoA)
            .LotNumber = GetSettingData(SettingName, "HannaCode" & i, "LotNumber", .LotNumber)
        End With

    Next
End Function
Public Function GetHannaCodesPerRecipeFromFile(ByVal t As Integer, ByRef HannaCodes() As HannaCode, ByVal HannaCodesCount As Integer, ByVal SettingName As String)
Dim i As Integer
    For i = 1 To HannaCodesCount
        
        With HannaCodes(i)
        
            .Code = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "Code", .Code)
            '.Density = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "Density", .Density)
            .Exp = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "Exp", .Exp)
            .ExpDate = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "ExpDate", .ExpDate)
            .ID = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "ID", .ID)
            '.LastLot = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "LastLot", .LastLot)
            .Line = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "Line", .Line)
            '.LoadInPrint = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "LoadInPrint", .LoadInPrint)
            .MaxQty = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "MaxQty", .MaxQty)
            .MinQty = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "MinQty", .MinQty)
            .Mix1 = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "Mix1", .Mix1)
            .Mix2 = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "Code", .Mix2)
            '.Procedure = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "Procedure", .Procedure)
            '.ProcedureRev = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "ProcedureRev", .ProcedureRev)
            .ProductName = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "ProductName", .ProductName)
            .Qty = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "Qty", .Qty)
            .Recipe = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "Recipe", .Recipe)
            '.Std = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "Std", .Std)
            .Um = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "Um", .Um)
            '.UncertantlyFromCoA = GetSettingData(SettingName, "Recipes" & t & " - HannaCode" & i, "UncertantlyFromCoA", .UncertantlyFromCoA)
            
        End With

    Next
End Function


Public Function GetRmxRecipeFromFile(ByVal t As Integer, ByRef RmxRecipe() As RmxRecipe, ByVal RmxRecipeCount As Integer, ByVal SettingName As String)
Dim i As Integer
    For i = 0 To RmxRecipeCount
        
        With RmxRecipe(i)
        
            .bMix = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "bMix", .bMix)
            .bUmMassa = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "bUmMassa", .bUmMassa)
            .Cas = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Cas", .Cas)
            .CHCode = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "CHCode", .CHCode)
            .Density = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Density", .Density)
            .Description = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Description", .Description)
            .ID = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "ID", .ID)
            .MaxQty = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "MaxQty", .MaxQty)
            .MinQty = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "MinQty", .MinQty)
            .MinQty2 = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "MinQty2", .MinQty2)
            .Multiple = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Multiple", .Multiple)
            .MultipleInCell = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "MultipleInCell", .MultipleInCell)
            .MultipleMassa = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "MultipleMassa", .MultipleMassa)
            .MultipleToProduce = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "MultipleToProduce", .MultipleToProduce)
            .Note = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Note", .Note)
            .Perc = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Perc", .Perc)
            .Qty = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Qty", .Qty)
            .CriticalRM = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "CriticalRM", .CriticalRM)
            '.Specifications.ChemicalReactionLiquid = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.ChemicalReactionLiquid", "")
            '.Specifications.Classification = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Classification", "")
            .Specifications.Location = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Location", "")
            '.Specifications.ManufacturerCode = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.ManufacturerCode", "")
            '.Specifications.ManufacturerName = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.ManufacturerName", "")
            '.Specifications.Pictograms = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Pictograms", "")
            '.Specifications.QtyToProduce = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.QtyToProduce", "")
            .Specifications.SpecifiedLocation = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.SpecifiedLocation", "")
            '.Specifications.Tollerance = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Tollerance", "")
            '.Specifications.UmQty = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.UmQty", "")

            
            
            .RecipeCode = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "RecipeCode", .RecipeCode)
            .TolerancePerc = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "TolerancePerc", .TolerancePerc)
            .TheoreticalWeight = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "TheoreticalWeight", .TheoreticalWeight)
            .TotalMultiple = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "TotalMultiple", .TotalMultiple)
            .TotalWeightKg = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "TotalWeightKg", .TotalWeightKg)
            .TotalWeightL = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "TotalWeightL", .TotalWeightL)
            .Um = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Um", .Um)
            .UmMax = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmMax", .UmMax)
            .UmMinQty = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmMinQty", .UmMinQty)
            .UmMultiple = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmMultiple", .UmMultiple)
            .UmTheoreticalWeight = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmTheoreticalWeight", .UmTheoreticalWeight)

        End With

    Next
End Function



Private Function GetPackagingInFile(ByRef Packaging() As ProdWay, ByVal PackagingCount As Integer, ByVal SettingName As String)
Dim i As Integer

    ReDim Packaging(PackagingCount)
    
    For i = 1 To PackagingCount
        
        With Packaging(i)
            
            .EsttimeD = GetSettingData(SettingName, "Packaging" & i, "EsttimeD", .EsttimeD)
            .EstTimeH = GetSettingData(SettingName, "Packaging" & i, "EstTimeH", .EstTimeH)
            .Line = GetSettingData(SettingName, "Packaging" & i, "Line", .Line)
            .STDPreparation = GetSettingData(SettingName, "Packaging" & i, "STDPreparation", .STDPreparation)
            .Recipe = GetSettingData(SettingName, "Packaging" & i, "Recipe", .Recipe)
            .Speed = GetSettingData(SettingName, "Packaging" & i, "Speed", .Speed)
            
        End With

    Next
End Function
