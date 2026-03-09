Attribute VB_Name = "gPreparation_04_GettSettingdata"
Option Explicit


Private SettingName As String
Private ExpDate As String

Public Function PreparationGetSetting(ByRef iPreparation As RecipeForProduction, ByVal SettName As String, ByVal RecipeCode As String) As Boolean
Dim rc As Boolean
Dim RecipeCount As Integer
Dim HannaCodesCount As Integer
Dim TotalsCount As Integer
Dim PackagingCount As Integer
Dim RfpFileName As String
Dim Path As String

On Error GoTo ERR_SAVE

    SettingName = SettName
    rc = True
   
    ' USER_PATH = USER_PREPARATION_PATH
    
    If FileExists(USER_PATH & SettingName) = False Then
    
        rc = False
        GoTo ERR_END:
        
    End If
    
    
    CloseSettingDataFile
  
    
    
    With iPreparation
       
        .bOpen = GetSettingData(SettingName, "iRecipeForProduction", "bOpen", .bOpen)
        .DateRecipe = GetSettingData(SettingName, "iRecipeForProduction", "DateRecipe", .DateRecipe)
        .Note = GetSettingData(SettingName, "iRecipeForProduction", "Note", .Note)
        .PlannedPrepWeek = GetSettingData(SettingName, "iRecipeForProduction", "PlannedPrepWeek", .PlannedPrepWeek)
        
        .PreparationDate = GetSettingData(SettingName, "iRecipeForProduction", "PreparationDate", "")
        .PreparationLot = GetSettingData(SettingName, "iRecipeForProduction", "PreparationLot", "")
        .ExpDate = GetSettingData(SettingName, "iRecipeForProduction", "ExpDate", "")
        
        ExpDate = .ExpDate
        
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
     PreparationGetSetting = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function




Public Function GetPreparationRecipesFromFile(ByRef Recipes() As RecipeType, ByRef RecipeCount As Integer, ByVal SettingName As String, ByVal RecipeCode As String) As Boolean
    
    Dim RecipeIndex As Integer
    Dim rc As Boolean
    Dim Count As Integer
    Count = 1

        RecipeIndex = GetSettingData(SettingName, "RecipeIndex", RecipeCode, 0)

        rc = GetPreparationSingleRecipeFromFile(Recipes(1), SettingName, RecipeIndex)



End Function


Public Function GetPreparationSingleRecipeFromFile(ByRef Recipe As RecipeType, ByVal SettingName As String, ByVal i As Integer) As Boolean


Dim HannaCodesCount As Integer
Dim MaterialRequisitionCount As Integer
Dim RmxRecipeCount As Integer
Dim rc As Boolean
On Error GoTo ERR_GET:

    rc = True

    With Recipe
    
        .AcquisitionCount = GetSettingData(SettingName, "Recipes" & i, "AcquisitionCount", .AcquisitionCount)
        .bHide = GetSettingData(SettingName, "Recipes" & i, "bHide", False)
        .Code = GetSettingData(SettingName, "Recipes" & i, "Code", .Code)
        .bHaveMixes = GetSettingData(SettingName, "Recipes" & i, "bHaveMixes", .bHaveMixes)
        .bUmMassa = GetSettingData(SettingName, "Recipes" & i, "bUmMassa", .bUmMassa)
        .bUpdated = GetSettingData(SettingName, "Recipes" & i, "bUpdated", False)
        .bIsMix = GetSettingData(SettingName, "Recipes" & i, "bIsMix", False)
        .PreparationLotMix = GetSettingData(SettingName, "Recipes" & i, "PreparationLotMix", "")
        .Machine = GetSettingData(SettingName, "Recipes" & i, "Machine", "")
        .bTestLot = GetSettingData(SettingName, "Recipes" & i, "bTestLot", False)
        
        
        .Density = GetSettingData(SettingName, "Recipes" & i, "Density", .Density)
        .Description = GetSettingData(SettingName, "Recipes" & i, "Description", .Description)
        .Exp = GetSettingData(SettingName, "Recipes" & i, "Exp", .Exp)
        .ExpDate = GetSettingData(SettingName, "Recipes" & i, "ExpDate", .ExpDate)
        .ID = GetSettingData(SettingName, "Recipes" & i, "ID", 0)
        .Line = GetSettingData(SettingName, "Recipes" & i, "Line", .Line)
        
        .MaxQty = GetSettingData(SettingName, "Recipes" & i, "MaxQty", .MaxQty)
        .MinQty = GetSettingData(SettingName, "Recipes" & i, "MinQty", .MinQty)
        .MinQty2 = GetSettingData(SettingName, "Recipes" & i, "MinQty2", .MinQty2)
        .Classification = GetSettingData(SettingName, "Recipes" & i, "Classification", .Classification)
        
        .bNoPreparation = GetSettingData(SettingName, "Recipes" & i, "bNoPreparation", .bNoPreparation)
        
        .Mix = GetSettingData(SettingName, "Recipes" & i, "Mix", .Mix)
        
        .Multiple = GetSettingData(SettingName, "Recipes" & i, "Multiple", .Multiple)
        .MultipleMassa = GetSettingData(SettingName, "Recipes" & i, "MultipleMassa", .MultipleMassa)
        .MultipleToProduce = GetSettingData(SettingName, "Recipes" & i, "MultipleToProduce", .MultipleToProduce)
        .NoteRev = GetSettingData(SettingName, "Recipes" & i, "NoteRev", .NoteRev)
        .Procedure = GetSettingData(SettingName, "Recipes" & i, "Procedure", .Procedure)
        .ProcedureDate = GetSettingData(SettingName, "Recipes" & i, "ProcedureDate", .ProcedureDate)
        .ProductionWay.EsttimeD = GetSettingData(SettingName, "Recipes" & i, "ProductionWay.EsttimeD", .ProductionWay.EsttimeD)
        .ProductionWay.EstTimeH = GetSettingData(SettingName, "Recipes" & i, "ProductionWay.EstTimeH", .ProductionWay.EstTimeH)
        .ProductionWay.Line = GetSettingData(SettingName, "Recipes" & i, "ProductionWay.Line", .ProductionWay.Line)
        .ProductionWay.Head = GetSettingData(SettingName, "Recipes" & i, "ProductionWay.Head", .ProductionWay.Head)
        .ProductionWay.Production = GetSettingData(SettingName, "Recipes" & i, "ProductionWay.Production", .ProductionWay.Production)
        .ProductionWay.Speed = GetSettingData(SettingName, "Recipes" & i, "ProductionWay.Speed", .ProductionWay.Speed)
        
        .Rev = SaveSettingData(SettingName, "Recipes" & i, "Rev", .Rev)
        
        .TotalMultiple = GetSettingData(SettingName, "Recipes" & i, "TotalMultiple", .TotalMultiple)
        
        .ActualWeight = GetSettingData(SettingName, "Recipes" & i, "ActualWeight", .ActualWeight)
        
         .bRecalculation = GetSettingData(SettingName, "Recipes" & i, "bRecalculation", .bRecalculation)
        
        .TotalRecipe = GetSettingData(SettingName, "Recipes" & i, "TotalRecipe", .TotalRecipe)
        .TotalWeightKg = GetSettingData(SettingName, "Recipes" & i, "TotalWeightKg", 0)
        .TotalWeightL = GetSettingData(SettingName, "Recipes" & i, "TotalWeightL", 0)
        .UmMax = GetSettingData(SettingName, "Recipes" & i, "UmMax", .UmMax)
        .UmMinQty = GetSettingData(SettingName, "Recipes" & i, "UmMinQty", .UmMinQty)
        .bHide = GetSettingData(SettingName, "Recipes" & i, "bHide", True)
        .UmMultiple = GetSettingData(SettingName, "Recipes" & i, "UmMultiple", .UmMultiple)
        .UmTotalWeightKg = GetSettingData(SettingName, "Recipes" & i, "UmTotalWeightKg", .UmTotalWeightKg)
        .UmTotalWeightL = GetSettingData(SettingName, "Recipes" & i, "UmTotalWeightL", .UmTotalWeightL)



        If .TotalWeightKg = 0 Then
            
            .TotalWeightKg = GetSettingData(SettingName, "Totals Grid" & i, "TotalWeightKg", 0)
            .TotalWeightL = GetSettingData(SettingName, "Totals Grid" & i, "TotalWeightL", 0)
        
        End If
        '-----------------------------------------------------------
        ' RmxRecipe
        '-----------------------------------------------------------
        .RmxRecipeCount = GetSettingData(SettingName, "Recipes" & i & " - RmxRecipe", "RmxRecipeCount", 0)
        If .RmxRecipeCount >= 0 Then
            RmxRecipeCount = .RmxRecipeCount
            ReDim .RmxRecipe(0)
            Call GetPreparationRmxRecipeFromFile(i, .RmxRecipe, RmxRecipeCount, SettingName, .TotalWeightKg, .Code)
            .RmxRecipeCount = RmxRecipeCount
        Else
            
            ' carico i componenti da RabRMxRecipe
            RmxRecipeCount = .RmxRecipeCount
            ReDim .RmxRecipe(0)
           ' Call GetPreparationRmxRecipeFromFile(i, .RmxRecipe, RmxRecipeCount, SettingName, .TotalWeightKg, .Code)
            Call SetRmxRecipeByRecipeCode(.Code, .RmxRecipe, , , RmxRecipeCount)
            
            .RmxRecipeCount = RmxRecipeCount
            
        End If
        
        '-----------------------------------------------------------
        ' Acquisition
        '-----------------------------------------------------------
        
        
        If .AcquisitionCount > 0 Then
            ReDim .Acquisitions(.AcquisitionCount)
            Call GetAcquisitionformFile(i, .Acquisitions, .AcquisitionCount, SettingName)
        Else
            .AcquisitionCount = GetNumberAcquisitionFormDatabase(SettingName)
            If .AcquisitionCount > 0 Then
            
                SaveSettingData SettingName, "Recipes" & i, "AcquisitionCount", .AcquisitionCount
                ReDim .Acquisitions(.AcquisitionCount)
                Call GetAcquisitionformDatabase(i, .Acquisitions, .AcquisitionCount, SettingName)
            
            End If
        End If
            
            
    End With

    CloseSettingDataFile

ERR_END:
    On Error GoTo 0
    GetPreparationSingleRecipeFromFile = rc
    Exit Function
ERR_GET:
    MsgBox err.Description
    Resume Next


End Function




Public Function GetPreparationRmxRecipeFromFile(ByVal t As Integer, ByRef RmxRecipe() As RmxRecipe, ByRef RmxRecipeCount As Integer, ByVal SettingName As String, ByVal TotalWeightKg As Double, ByVal RecipeCode As String)
Dim i As Integer
Dim r As Integer
Dim RmxRecipeCode As String

On Error GoTo ERR_GET:

    r = 0
    For i = 0 To RmxRecipeCount
    
        RmxRecipeCode = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "RecipeCode", "")
        
        If RecipeCode <> RmxRecipeCode Then GoTo cont:
        
        ReDim Preserve RmxRecipe(r)
        
        With RmxRecipe(r)
        
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
            .RealPerc = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "RealPerc", .RealPerc)
            .Qty = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Qty", .Qty)
            .CriticalRM = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "CriticalRM", .CriticalRM)

            .RecipeCode = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "RecipeCode", .RecipeCode)
            .TolerancePerc = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "TolerancePerc", .TolerancePerc)
            .bAddedInPreparation = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "bAddedInPreparation", .bAddedInPreparation)
            
            .TheoreticalWeight = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "TheoreticalWeight", 0)
            .RealWeight = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "RealWeight", 0)
            
            .Variance = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Variance", 0)
            .VariancePerc = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "VariancePerc", 0)
            
            .TotalMultiple = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "TotalMultiple", 0)
            .TotalWeightKg = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "TotalWeightKg", 0)
            .TotalWeightL = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "TotalWeightL", 0)
            .Um = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Um", .Um)
            .UmMax = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmMax", .UmMax)
            .UmMinQty = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmMinQty", .UmMinQty)
            .UmMultiple = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmMultiple", .UmMultiple)
            .UmTheoreticalWeight = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmTheoreticalWeight", .UmTheoreticalWeight)

            
            .Specifications.ChemicalReactionLiquid = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.ChemicalReactionLiquid", .Specifications.ChemicalReactionLiquid)
            .Specifications.Classification = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Classification", .Specifications.Classification)
            .Specifications.Location = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Location", .Specifications.Location)
            .Specifications.ManufacturerCode = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.ManufacturerCode", .Specifications.ManufacturerCode)
            .Specifications.ManufacturerName = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.ManufacturerName", .Specifications.ManufacturerName)
            .Specifications.Pictograms = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Pictograms", .Specifications.Pictograms)
            .Specifications.QtyToProduce = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.QtyToProduce", .Specifications.QtyToProduce)
            .Specifications.SpecifiedLocation = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.SpecifiedLocation", .Specifications.SpecifiedLocation)
            .Specifications.Tollerance = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Tollerance", .Specifications.Tollerance)
            .Specifications.UmQty = GetSettingData(SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.UmQty", .Specifications.UmQty)

            
            
           ' If .Variance = 0 Then
            
                If (.RealWeight) > 0 And (.TheoreticalWeight) > 0 Then
                    .Variance = .RealWeight - .TheoreticalWeight
                End If
                
                If (.Variance) <> 0 And (.RealWeight) > 0 Then
                     .VariancePerc = (.Variance / .RealWeight) * 100
                End If
             
           ' End If
      
                




            If .bAddedInPreparation = False Then
                .TheoreticalWeight = TotalWeightKg * 1000 * .Perc / 100
                If .TheoreticalWeight = 0 And .bMix Then
                    
                End If
            End If

            
        End With
        r = r + 1
cont:
    Next
    
    RmxRecipeCount = r - 1
    
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_GET:
    MsgBox err.Description
    Resume Next


End Function


Private Function GetAcquisitionformFile(ByVal t As Integer, ByRef Acquisition() As PrepAcquisition, ByVal AcquisitionCount As Integer, ByVal SettingName As String)
Dim r As Integer
    
    CloseSettingDataFile
    
    
    For r = 1 To AcquisitionCount
        
        With Acquisition(r)
            .AcquisitionTime = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "AcquisitionTime", .AcquisitionTime)
            .ActualWeight = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "ActualWeight", .ActualWeight)
            .bFromBarcode = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "bFromBarcode", .bFromBarcode)
            .bRecalculation = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "bRecalculation", .bRecalculation)
            .bRecipeComponent = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "bRecipeComponent", .bRecipeComponent)
            .ID = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "ID", .ID)
            .Index = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "Index", .Index)
            .Note = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "Note", .Note)
            .Operator = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "Operator", .Operator)
            .PrepBarcode.Cas = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "Cas", .PrepBarcode.Cas)
            .PrepBarcode.ChemicalName = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "ChemicalName", .PrepBarcode.ChemicalName)
            .PrepBarcode.Code = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "Code", .PrepBarcode.Code)
            .PrepBarcode.DeliveryDate = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "DeliveryDate", .PrepBarcode.DeliveryDate)
            .PrepBarcode.Manufacturer = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "Manufacturer", .PrepBarcode.Manufacturer)
            .PrepBarcode.ManufacturerCode = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "ManufacturerCode", .PrepBarcode.ManufacturerCode)
            .PrepBarcode.ManufacturerLot = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "ManufacturerLot", .PrepBarcode.ManufacturerLot)
            .PrepBarcode.Package = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "Package", .PrepBarcode.Package)
            .PrepBarcode.QtyDelivered = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "QtyDelivered", .PrepBarcode.QtyDelivered)
            .PrepBarcode.WeekDelPackageNumber = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "WeekDelPackageNumber", .PrepBarcode.WeekDelPackageNumber)
            .ExpDate = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "ExpDate", .ExpDate)
        
        End With
    Next

   
    
    CloseSettingDataFile
    
End Function

Public Function GetPreparationHannaCodesFromFile(ByRef HannaCodes() As HannaCode, ByRef HannaCodesCount As Integer, ByVal SettingName As String, ByVal RecipeCode As String, ByVal Path As String)
Dim i As Integer
Dim t As Integer
Dim bHide As Boolean
Dim RecipeForHannaCode As String

    t = 1
    For i = 1 To HannaCodesCount
    
        
        bHide = GetSettingData(SettingName, "HannaCode" & i, "bHide", True, Path)
        
        RecipeForHannaCode = GetSettingData(SettingName, "HannaCode" & i, "Recipe", "", Path)
        
        If RecipeForHannaCode <> RecipeCode Then GoTo cont
        If bHide = True Then GoTo cont
        
        ReDim Preserve HannaCodes(t)
        
        With HannaCodes(t)
            .Index = i
            .bHide = GetSettingData(SettingName, "HannaCode" & i, "bHide", .bHide, Path)
            .Code = GetSettingData(SettingName, "HannaCode" & i, "Code", .Code, Path)
            .Density = GetSettingData(SettingName, "HannaCode" & i, "Density", .Density, Path)
            .Exp = GetSettingData(SettingName, "HannaCode" & i, "Exp", .Exp, Path)
            .ExpDate = ExpDate
            .ID = GetSettingData(SettingName, "HannaCode" & i, "ID", .ID, Path)
            .LastLot = GetSettingData(SettingName, "HannaCode" & i, "LastLot", .LastLot, Path)
            .Line = GetSettingData(SettingName, "HannaCode" & i, "Line", .Line, Path)
            .Procedure = GetSettingData(SettingName, "HannaCode" & i, "Procedure", .Procedure, Path)
            .ProcedureRev = GetSettingData(SettingName, "HannaCode" & i, "ProcedureRev", .ProcedureRev, Path)
            .ProductName = GetSettingData(SettingName, "HannaCode" & i, "ProductName", .ProductName, Path)
            .Qty = GetSettingData(SettingName, "HannaCode" & i, "Qty", .Qty, Path)
            .QtyToProduce = GetSettingData(SettingName, "HannaCode" & i, "QtyToProduce", .QtyToProduce, Path)
            .Recipe = GetSettingData(SettingName, "HannaCode" & i, "Recipe", .Recipe, Path)
            .STD = GetSettingData(SettingName, "HannaCode" & i, "Std", .STD, Path)
            .Um = GetSettingData(SettingName, "HannaCode" & i, "Um", .Um, Path)
            .LotNumber = GetSettingData(SettingName, "HannaCode" & i, "LotNumber", .LotNumber, Path)
            
        End With
        
        
        t = t + 1
cont:
    Next
    CloseSettingDataFile
    HannaCodesCount = t - 1
    
End Function






Public Function GetQCFromPreparationFile(ByRef QCStatus() As QCType, ByVal QCStatusCount As Integer, ByVal SettingName As String, ByVal Path As String)
Dim i As Integer

    CloseSettingDataFile

    ReDim QCStatus(QCStatusCount)
    
    For i = 1 To QCStatusCount
        
        
        With QCStatus(i)
            .Status = GetSettingData(SettingName, "QC", "Status" & i, .Status)
            .Operator = GetSettingData(SettingName, "QC", "Operator" & i, .Operator)
            .Date = GetSettingData(SettingName, "QC", "Date" & i, .Date)
            .Note = GetSettingData(SettingName, "QC", "Note" & i, .Note)
            .Registration = GetSettingData(SettingName, "QC", "Registration" & i, .Registration)
            .QCOperator = GetSettingData(SettingName, "QC", "QCOperator" & i, .QCOperator)
            .Correction = GetSettingData(SettingName, "QC", "Correction" & i, .Correction)
            .CorrectionDate = GetSettingData(SettingName, "QC", "CorrectionDate" & i, .CorrectionDate)
        End With
        
    Next
    
    CloseSettingDataFile

End Function


Private Function GetNumberAcquisitionFormDatabase(ByVal SettingName As String) As Integer
    GetNumberAcquisitionFormDatabase = 0
    With dbTabAcquisition
        .filter = ""
        .filter = "FileName='" & SettingName & "'"
        If .EOF Then
        Else
            GetNumberAcquisitionFormDatabase = .RecordCount
        End If
        
    
    
    End With
End Function

Private Function GetAcquisitionformDatabase(ByVal t As Integer, ByRef Acquisition() As PrepAcquisition, ByVal AcquisitionCount As Integer, ByVal SettingName As String)
Dim r As Integer
    
    CloseSettingDataFile
    
    dbTabAcquisition.MoveFirst
    
    For r = 1 To AcquisitionCount
    
        
        
        With Acquisition(r)
        
            ' get form database
            
            .AcquisitionTime = IIf(IsNull(Trim(dbTabAcquisition!AcquisitionTime)), "", Trim(dbTabAcquisition!AcquisitionTime))
            .ActualWeight = IIf(IsNull(Trim(dbTabAcquisition!ActualWeight)), 0, Trim(dbTabAcquisition!ActualWeight))
            .bFromBarcode = dbTabAcquisition!bFromBarcode
            .bRecalculation = dbTabAcquisition!bRecalculation
            .bRecipeComponent = dbTabAcquisition!bRecipeComponent
            .ID = dbTabAcquisition!ID
            .Index = IIf(IsNull(Trim(dbTabAcquisition!Index)), 0, Trim(dbTabAcquisition!Index))
            .Note = IIf(IsNull(Trim(dbTabAcquisition!Note)), "", Trim(dbTabAcquisition!Note))
            .Operator = IIf(IsNull(Trim(dbTabAcquisition!Operator)), "", Trim(dbTabAcquisition!Operator))
            .PrepBarcode.Cas = IIf(IsNull(Trim(dbTabAcquisition!Cas)), "", Trim(dbTabAcquisition!Cas))
            .PrepBarcode.ChemicalName = IIf(IsNull(Trim(dbTabAcquisition!ChemicalName)), "", Trim(dbTabAcquisition!ChemicalName))
            .PrepBarcode.Code = IIf(IsNull(Trim(dbTabAcquisition!Code)), "", Trim(dbTabAcquisition!Code))
            .PrepBarcode.DeliveryDate = IIf(IsNull(Trim(dbTabAcquisition!DeliveryDate)), "", Trim(dbTabAcquisition!DeliveryDate))
            .PrepBarcode.Manufacturer = IIf(IsNull(Trim(dbTabAcquisition!Manufacturer)), "", Trim(dbTabAcquisition!Manufacturer))
            .PrepBarcode.ManufacturerCode = IIf(IsNull(Trim(dbTabAcquisition!ManufacturerCode)), "", Trim(dbTabAcquisition!ManufacturerCode))
            .PrepBarcode.ManufacturerLot = IIf(IsNull(Trim(dbTabAcquisition!ManufacturerLot)), "", Trim(dbTabAcquisition!ManufacturerLot))
            .PrepBarcode.Package = IIf(IsNull(Trim(dbTabAcquisition!Package)), "", Trim(dbTabAcquisition!Package))
            .PrepBarcode.QtyDelivered = IIf(IsNull(Trim(dbTabAcquisition!QtyDelivered)), "", Trim(dbTabAcquisition!QtyDelivered))
            .PrepBarcode.WeekDelPackageNumber = IIf(IsNull(Trim(dbTabAcquisition!WeekDelPackageNumber)), "", Trim(dbTabAcquisition!WeekDelPackageNumber))
            .ExpDate = IIf(IsNull(Trim(dbTabAcquisition!ExpDate)), "", Trim(dbTabAcquisition!ExpDate))
            
            dbTabAcquisition.MoveNext
            
        End With
        
        With Acquisition(r)
            
            'save on file!!!
            
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "AcquisitionTime", .AcquisitionTime
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "ActualWeight", .ActualWeight
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "bFromBarcode", .bFromBarcode
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "bRecalculation", .bRecalculation
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "bRecipeComponent", .bRecipeComponent
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "ID", .ID
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "Index", .Index
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "Note", .Note
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "Operator", .Operator
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "Cas", .PrepBarcode.Cas
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "ChemicalName", .PrepBarcode.ChemicalName
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "Code", .PrepBarcode.Code
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "DeliveryDate", .PrepBarcode.DeliveryDate
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "Manufacturer", .PrepBarcode.Manufacturer
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "ManufacturerCode", .PrepBarcode.ManufacturerCode
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "ManufacturerLot", .PrepBarcode.ManufacturerLot
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "Package", .PrepBarcode.Package
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "QtyDelivered", .PrepBarcode.QtyDelivered
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "WeekDelPackageNumber", .PrepBarcode.WeekDelPackageNumber
            SaveSettingData SettingName, "Recipes" & t & " - Acquisition " & r, "ExpDate", .ExpDate
        
        End With
        
    Next

   
    
    CloseSettingDataFile
    
End Function
