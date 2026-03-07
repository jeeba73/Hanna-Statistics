Attribute VB_Name = "gPreparation_03_SaveSettingData"
Option Explicit


Private SettingName As String

Public Function PeparationSaveSetting(ByRef iPreparation As RecipeForProduction, ByVal SettName As String) As Boolean
Dim rc As Boolean
Dim RecipeCount As Integer
Dim HannaCodesCount As Integer
Dim TotalsCount As Integer
Dim PackagingCount As Integer
Dim FileName As String
Dim Path As String

On Error GoTo ERR_SAVE

    SettingName = SettName
    
    rc = True

    If USER_PATH = "" Then USER_PATH = USER_PREPARATION_PATH
    
    
    CloseSettingDataFile
    
    SaveSettingData SettingName, "iRecipeForProduction", "fileName", SettingName
    
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
    
    
    With iPreparation
        SaveSettingData SettingName, "iRecipeForProduction", "fileName", SettingName
        SaveSettingData SettingName, "iRecipeForProduction", "bOpen", .bOpen
        SaveSettingData SettingName, "iRecipeForProduction", "DateRecipe", .DateRecipe
        SaveSettingData SettingName, "iRecipeForProduction", "PreparationDate", .PreparationDate
        SaveSettingData SettingName, "iRecipeForProduction", "PreparationLot", .PreparationLot
        SaveSettingData SettingName, "iRecipeForProduction", "ExpDate", .ExpDate
        SaveSettingData SettingName, "iRecipeForProduction", "PrepWeek", .PrepWeek
        SaveSettingData SettingName, "iRecipeForProduction", "Note", .Note
        SaveSettingData SettingName, "iRecipeForProduction", "PlannedPrepWeek", .PlannedPrepWeek
        SaveSettingData SettingName, "iRecipeForProduction", "bAllMixes", .bAllMixes
        SaveSettingData SettingName, "iRecipeForProduction", "PlanningReference", .PlanningReference
        SaveSettingData SettingName, "iRecipeForProduction", "NumPrepWeek", .numPrepWeek
        SaveSettingData SettingName, "iRecipeForProduction", "RecipeBy", .RecipeBy
        SaveSettingData SettingName, "iRecipeForProduction", "fileNameRecForProd", .fileNameRecForProd
        
        SaveSettingData SettingName, "iRecipeForProduction", "bCorrection", .bCorrection
        SaveSettingData SettingName, "iRecipeForProduction", "OperatorPrep", .OperatorPrep
        SaveSettingData SettingName, "iRecipeForProduction", "OperatorRfp", .OperatorRfP
       
        CloseSettingDataFile

        If .Recipes(1).bIsMix = False Then
            '-----------------------------------------------------------
            '
            ' salvo le informazioni nel file originale rfp
            '
            '-----------------------------------------------------------
            Dim USER_FRP_PATH As String
        
            CloseSettingDataFile
    
            If FileExists(USER_TEMP_PATH & .fileNameRecForProd) Then
                USER_FRP_PATH = USER_TEMP_PATH
            ElseIf FileExists(USER_DATA_PATH & .fileNameRecForProd) Then
                USER_FRP_PATH = USER_DATA_PATH
            Else
                GoTo cont:
            End If
            Debug.Print
            SaveSettingData .fileNameRecForProd, "iRecipeForProduction", "PreparationDate", .PreparationDate, USER_FRP_PATH
            SaveSettingData .fileNameRecForProd, "iRecipeForProduction", "PreparationLot", .PreparationLot, USER_FRP_PATH
            SaveSettingData .fileNameRecForProd, "iRecipeForProduction", "ExpDate", .ExpDate, USER_FRP_PATH
            SaveSettingData .fileNameRecForProd, "iRecipeForProduction", "PrepWeek", .PrepWeek, USER_FRP_PATH
            SaveSettingData .fileNameRecForProd, "iRecipeForProduction", "Note", .Note, USER_FRP_PATH
           'SaveSettingData .fileNameRecForProd, "iRecipeForProduction", "PlannedPrepWeek", .PlannedPrepWeek, USER_FRP_PATH
            'SaveSettingData .fileNameRecForProd, "iRecipeForProduction", "bAllMixes", .bAllMixes, USER_FRP_PATH
            'SaveSettingData .fileNameRecForProd, "iRecipeForProduction", "PlanningReference", .PlanningReference, USER_FRP_PATH
            SaveSettingData .fileNameRecForProd, "iRecipeForProduction", "NumPrepWeek", .numPrepWeek, USER_FRP_PATH
           ' SaveSettingData .fileNameRecForProd, "iRecipeForProduction", "RecipeBy", .RecipeBy, USER_FRP_PATH
            'SaveSettingData .fileNameRecForProd, "iRecipeForProduction", "fileNameRecForProd", .fileNameRecForProd, USER_FRP_PATH
            SaveSettingData SettingName, "iRecipeForProduction", "OperatorPrep", .OperatorPrep, USER_FRP_PATH
            SaveSettingData SettingName, "iRecipeForProduction", "OperatorRfp", .OperatorRfP, USER_FRP_PATH
            
            
            CloseSettingDataFile
        End If
        
        '-----------------------------------------------------------
        ' Recipes
        '-----------------------------------------------------------
cont:
        .RecipeCount = UBound(.Recipes)
        
        SaveSettingData SettingName, "Recipes", "RecipeCount", .RecipeCount
        
        RecipeCount = .RecipeCount

        Call SetPreparationRecipesInFile(.Recipes, RecipeCount)




                
        '-----------------------------------------------------------
        ' HANNA CODES in Recipe for production
        '-----------------------------------------------------------
        
        
            If FileExists(USER_PRODUCTION_PATH & .fileNameRecForProd) Then
                Path = USER_PRODUCTION_PATH
            ElseIf FileExists(USER_TEMP_PATH & .fileNameRecForProd) Then
              
                Path = USER_TEMP_PATH
            ElseIf FileExists(USER_DATA_PATH & .fileNameRecForProd) Then
            
                Path = USER_DATA_PATH
            End If
                
            
        If .HannaCodesCount > 0 Then
            HannaCodesCount = .HannaCodesCount
          '  SaveSettingData SettingName, "HannaCodes", "HannaCodesCount", .HannaCodesCount
            Call SetHannaCodesInFile(.HannaCodes, HannaCodesCount, .fileNameRecForProd, Path)
        End If
        
    
        If .QCCount > 0 Then
        
        
            SaveSettingData SettName, "QC", "Count", .QCCount
            
            Call SetQCInPreparationFile(.QCStatus, .QCCount, SettingName, Path)
        
        End If
        
     
     
    End With

ERR_END:
    On Error GoTo 0
     CloseSettingDataFile
     PeparationSaveSetting = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function



Public Function SetPreparationRecipesInFile(ByRef Recipes() As RecipeType, ByVal RecipeCount As Integer) As Boolean

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
        SaveSettingData SettingName, "Recipes" & i, "ExpDate", .ExpDate
        SaveSettingData SettingName, "Recipes" & i, "ID", i
        SaveSettingData SettingName, "Recipes" & i, "Line", .Line
        
        SaveSettingData SettingName, "Recipes" & i, "MaxQty", .MaxQty
        SaveSettingData SettingName, "Recipes" & i, "MinQty", .MinQty
        SaveSettingData SettingName, "Recipes" & i, "MinQty2", .MinQty2
        SaveSettingData SettingName, "Recipes" & i, "Classification", .Classification
        SaveSettingData SettingName, "Recipes" & i, "Mix", .Mix
        SaveSettingData SettingName, "Recipes" & i, "bNoPreparation", .bNoPreparation
        
        
        SaveSettingData SettingName, "Recipes" & i, "Multiple", .Multiple
        SaveSettingData SettingName, "Recipes" & i, "MultipleMassa", .MultipleMassa
        SaveSettingData SettingName, "Recipes" & i, "MultipleToProduce", .MultipleToProduce
        SaveSettingData SettingName, "Recipes" & i, "NoteRev", .NoteRev
        SaveSettingData SettingName, "Recipes" & i, "Procedure", .Procedure
        SaveSettingData SettingName, "Recipes" & i, "ProcedureDate", .ProcedureDate
        SaveSettingData SettingName, "Recipes" & i, "ProductionWay.EsttimeD", .ProductionWay.EsttimeD
        SaveSettingData SettingName, "Recipes" & i, "ProductionWay.EstTimeH", .ProductionWay.EstTimeH
        SaveSettingData SettingName, "Recipes" & i, "ProductionWay.Line", .ProductionWay.Line
        SaveSettingData SettingName, "Recipes" & i, "ProductionWay.Head", .ProductionWay.Head
        SaveSettingData SettingName, "Recipes" & i, "ProductionWay.Production", .ProductionWay.Production
        SaveSettingData SettingName, "Recipes" & i, "ProductionWay.Speed", .ProductionWay.Speed
        
        SaveSettingData SettingName, "Recipes" & i, "Rev", .Rev
        
        SaveSettingData SettingName, "Recipes" & i, "ActualWeight", .ActualWeight
        
        SaveSettingData SettingName, "Recipes" & i, "bRecalculation", .bRecalculation
        
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

       ' SaveSettingData SettingName, "Recipes" & i, "AcquisitionCount", .AcquisitionCount

        '-----------------------------------------------------------
        ' RmxRecipe
        '-----------------------------------------------------------
        If .RmxRecipeCount >= 0 Then
            RmxRecipeCount = .RmxRecipeCount
            'SaveSettingData SettingName, "Recipes" & i & " - RmxRecipe", "RmxRecipeCount", .RmxRecipeCount
            Call SetPreparationRmxRecipeInFile(i, .RmxRecipe, RmxRecipeCount)
        End If
        
        SaveSettingData SettingName, "Recipes" & i & " - RmxRecipe", "RmxRecipeCount", .RmxRecipeCount
   
        If .AcquisitionCount > 0 Then
            Call SetAcquisitionInFile(i, .Acquisitions, .AcquisitionCount)
        End If
            
            
        SaveSettingData SettingName, "Recipes" & i, "AcquisitionCount", .AcquisitionCount
        
    End With
Next

CloseSettingDataFile


End Function



Public Function SetPreparationRmxRecipeInFile(ByVal t As Integer, ByRef RmxRecipe() As RmxRecipe, ByRef RmxRecipeCount As Integer)
Dim r As Integer
Dim i As Integer
    For r = 0 To RmxRecipeCount
    
    
        If RmxRecipe(r).bDeleted Then GoTo cont
        
        
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
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "RealPerc", .RealPerc
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Qty", .Qty
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "CriticalRM", .CriticalRM
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.ChemicalReactionLiquid", .Specifications.ChemicalReactionLiquid
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Classification", .Specifications.Classification
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Location", .Specifications.Location
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.ManufacturerCode", .Specifications.ManufacturerCode
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.ManufacturerName", .Specifications.ManufacturerName
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Pictograms", .Specifications.Pictograms
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.QtyToProduce", .Specifications.QtyToProduce
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.SpecifiedLocation", .Specifications.SpecifiedLocation
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.Tollerance", .Specifications.Tollerance
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Specifications.UmQty", .Specifications.UmQty

            
            
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "RecipeCode", .RecipeCode
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "TolerancePerc", .TolerancePerc
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "bAddedInPreparation", .bAddedInPreparation
            
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "TheoreticalWeight", .TheoreticalWeight
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "RealWeight", .RealWeight
            
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Variance", .Variance
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "VariancePerc", .VariancePerc
            
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "TotalMultiple", .TotalMultiple
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "TotalWeightKg", .TotalWeightKg
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "TotalWeightL", .TotalWeightL
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "Um", .Um
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmMax", .UmMax
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmMinQty", .UmMinQty
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmMultiple", .UmMultiple
            SaveSettingData SettingName, "Recipes" & t & " - RmxRecipe" & i, "UmTheoreticalWeight", .UmTheoreticalWeight


        End With
        
        i = i + 1
        
cont:

    Next
    
    RmxRecipeCount = i - 1
    
    CloseSettingDataFile
End Function


Private Function SetAcquisitionInFile(ByVal t As Integer, ByRef Acquisition() As PrepAcquisition, ByRef AcquisitionCount As Integer)
Dim r As Integer
Dim i As Integer
    CloseSettingDataFile
    r = 1
    
    For i = 1 To AcquisitionCount
    
        If Acquisition(i).bDeleted Then GoTo cont
    
        With Acquisition(r)
            
        
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
        
        r = r + 1
cont:
    Next

    AcquisitionCount = r - 1
    
    CloseSettingDataFile
    
End Function
Public Function SetHannaCodesInFile(ByRef HannaCodes() As HannaCode, ByVal HannaCodesCount As Integer, ByVal SettName As String, ByVal Path As String)
Dim i As Integer

    CloseSettingDataFile

    For i = 1 To HannaCodesCount
        
        With HannaCodes(i)
        
        
          '  SaveSettingData SettingName, "HannaCode" & .Index, "bHide", .bHide, PATH
            SaveSettingData SettName, "HannaCode" & .Index, "LotNumber", .LotNumber, Path
            
          '  SaveSettingData SettingName, "HannaCode" & .Index, "Code", .code, PATH
          '  SaveSettingData SettingName, "HannaCode" & .Index, "Density", .Density, PATH
            SaveSettingData SettingName, "HannaCode" & .Index, "Exp", .Exp, Path
            SaveSettingData SettingName, "HannaCode" & .Index, "ExpDate", .ExpDate, Path
          '  SaveSettingData SettingName, "HannaCode" & .Index, "ID", .ID, PATH
          '  SaveSettingData SettingName, "HannaCode" & .Index, "LastLot", .LastLot, PATH
          '  SaveSettingData SettingName, "HannaCode" & .Index, "Line", .Line, PATH
           
          '  SaveSettingData SettingName, "HannaCode" & .Index, "Procedure", .Procedure, PATH
          '  SaveSettingData SettingName, "HannaCode" & .Index, "ProcedureRev", .ProcedureRevv
          '  SaveSettingData SettingName, "HannaCode" & .Index, "ProductName", .ProductName, PATH
          '  SaveSettingData SettingName, "HannaCode" & .Index, "Qty", .Qty, PATH
          '  SaveSettingData SettingName, "HannaCode" & .Index, "QtyToProduce", .QtyToProduce, PATH
          '  SaveSettingData SettingName, "HannaCode" & .Index, "Recipe", .Recipe, PATH
          '  SaveSettingData SettingName, "HannaCode" & .Index, "Std", .Std, PATH
          '  SaveSettingData SettingName, "HannaCode" & .Index, "Um", .Um, PATH
            
        
        End With

    Next
    
    CloseSettingDataFile
    
End Function



Private Function SetQCInPreparationFile(ByRef QCStatus() As QCType, ByVal QCStatusCount As Integer, ByVal SettName As String, ByVal Path As String)
Dim i As Integer

    CloseSettingDataFile

    For i = 1 To QCStatusCount
        
        With QCStatus(i)
            SaveSettingData SettName, "QC", "Status" & i, .Status
            SaveSettingData SettName, "QC", "Operator" & i, .Operator
            SaveSettingData SettName, "QC", "Date" & i, .Date
            SaveSettingData SettName, "QC", "Note" & i, .Note
            SaveSettingData SettName, "QC", "Registration" & i, .Registration
            SaveSettingData SettName, "QC", "QCOperator" & i, .QCOperator
            SaveSettingData SettName, "QC", "Correction" & i, .Correction
            SaveSettingData SettName, "QC", "CorrectionDate" & i, .CorrectionDate
        End With
        
    Next
    
    CloseSettingDataFile

End Function
