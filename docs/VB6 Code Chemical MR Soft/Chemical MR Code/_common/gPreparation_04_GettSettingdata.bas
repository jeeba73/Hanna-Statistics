Attribute VB_Name = "gPreparation_04_GettSettingdata"
Option Explicit


Private SettingName As String
Private ExpDate As String

Public Function PreparationGetSetting(ByRef iPreparation As RecipeForProduction, ByVal SettName As String, ByVal HannaCode As String) As Boolean
Dim rc As Boolean
Dim RecipeCount As Integer
Dim HannaCodeCount As Integer
Dim TotalsCount As Integer
Dim PackagingCount As Integer
Dim RfPfileName As String
Dim PATH As String

On Error GoTo ERR_SAVE

    SettingName = SettName
    rc = True
   
    ' USER_PATH = USER_TEMP_PATH
    
    If FileExists(USER_PATH & SettingName) = False Then
    
        rc = False
        GoTo ERR_END:
        
    End If
    
    
    CloseSettingDataFile
  
    
    
    With iPreparation
       
       .ID = GetSettingData(SettingName, "iRecipeForProduction", "ID", .ID)
       
       
       
       
       .HannaCode.Code = GetSettingData(SettingName, "iRecipeForProduction", "HannaCode", .HannaCode.Code)
       .HannaCode.Description = GetSettingData(SettingName, "iRecipeForProduction", "Description", .HannaCode.Description)
       .MRCode = GetSettingData(SettingName, "iRecipeForProduction", "MRCode", .MRCode)
       .DataPrep = GetSettingData(SettingName, "iRecipeForProduction", "DataPrep", .DataPrep)
       .HourPrep = GetSettingData(SettingName, "iRecipeForProduction", "HourPrep", .HourPrep)
       .PrepWeek = GetSettingData(SettingName, "iRecipeForProduction", "PrepWeek", .PrepWeek)
       .Operator = GetSettingData(SettingName, "iRecipeForProduction", "Operator", .Operator)
       .QtyToProduce = GetSettingData(SettingName, "iRecipeForProduction", "QtyToProduce", .QtyToProduce)
       
       .ManualValue = GetSettingData(SettingName, "iRecipeForProduction", "ManualValue", .ManualValue)
       .ManualUnit = GetSettingData(SettingName, "iRecipeForProduction", "ManualUnit", .ManualUnit)
       
       .STD_Manual_ID = GetSettingData(SettingName, "iRecipeForProduction", "STD_Manual_ID", .STD_Manual_ID)
       
       .QtyProduced = GetSettingData(SettingName, "iRecipeForProduction", "QtyProduced", .QtyProduced)
       .Unit = GetSettingData(SettingName, "iRecipeForProduction", "Unit", .Unit)
       .HannaCode.STDMatrix = GetSettingData(SettingName, "iRecipeForProduction", "STDMatrix", .HannaCode.STDMatrix)
       .HannaCode.STDExp = GetSettingData(SettingName, "iRecipeForProduction", "STDExp", .HannaCode.STDExp)
       .HannaCode.STDStorage = GetSettingData(SettingName, "iRecipeForProduction", "STDStorage", .HannaCode.STDStorage)
       .bClosed = GetSettingData(SettingName, "iRecipeForProduction", "bClosed", .bClosed)
       .CloseDate = GetSettingData(SettingName, "iRecipeForProduction", "CloseDate", .CloseDate)
       .FileName = SettingName
       .MsType = GetSettingData(SettingName, "iRecipeForProduction", "MSType", .MsType)
       .MotherSol.ID = GetSettingData(SettingName, "iRecipeForProduction", "MotherSolID", .MotherSol.ID)
       .Note = GetSettingData(SettingName, "iRecipeForProduction", "Note", .Note)
       .bCorrection = GetSettingData(SettingName, "iRecipeForProduction", "bCorrection", .bCorrection)
       .ExpDate = GetSettingData(SettingName, "iRecipeForProduction", "ExpDate", .ExpDate)

        ExpDate = .ExpDate
        
        .bCorrection = GetSettingData(SettingName, "iRecipeForProduction", "bCorrection", .bCorrection)
        .Operator = GetSettingData(SettingName, "iRecipeForProduction", "Operator", .Operator)
        
        
        '-----------------------------------------------------------
        ' MotherSolution specifics in HannaCode
        '-----------------------------------------------------------
        
        .MS.DilConc = GetSettingData(SettingName, "iRecipeForProduction", ".MS.DilConc", .MS.DilConc)
        .MS.Exp = GetSettingData(SettingName, "iRecipeForProduction", ".MS.Exp", .MS.Exp)
        .MS.Qty = GetSettingData(SettingName, "iRecipeForProduction", ".MS.Qty", .MS.Qty)
        .MS.Unit = GetSettingData(SettingName, "iRecipeForProduction", ".MS.Unit", .MS.Unit)
        .MS.Value = GetSettingData(SettingName, "iRecipeForProduction", ".MS.Value", .MS.Value)
        .MS.Volume = GetSettingData(SettingName, "iRecipeForProduction", ".MS.Volume", .MS.Volume)
        
        
        '-----------------------------------------------------------
        ' MotherSolution Prepared/Used
        '-----------------------------------------------------------
        
         .MotherSol.bClosed = GetSettingData(SettingName, "MotherSolution", "bClosed", .MotherSol.bClosed)
         .MotherSol.Code = GetSettingData(SettingName, "MotherSolution", "Code", .MotherSol.Code)
         .MotherSol.DataPrep = GetSettingData(SettingName, "MotherSolution", "DataPrep", .MotherSol.DataPrep)
         .MotherSol.ExpDays = GetSettingData(SettingName, "MotherSolution", "ExpDays", .MotherSol.ExpDays)
         .MotherSol.HannaCode = GetSettingData(SettingName, "MotherSolution", "HannaCode", .MotherSol.HannaCode)
         .MotherSol.HourPrep = GetSettingData(SettingName, "MotherSolution", "HourPrep", .MotherSol.HourPrep)
         .MotherSol.ID = GetSettingData(SettingName, "MotherSolution", "ID", .MotherSol.ID)
         .MotherSol.MsType = GetSettingData(SettingName, "MotherSolution", "MSType", .MotherSol.MsType)
         .MotherSol.Note = GetSettingData(SettingName, "MotherSolution", "Note", .MotherSol.Note)
         .MotherSol.Operator = GetSettingData(SettingName, "MotherSolution", "Operator", .MotherSol.Operator)
         .MotherSol.QtyLeft = GetSettingData(SettingName, "MotherSolution", "QtyLeft", .MotherSol.QtyLeft)
         .MotherSol.PreparationID = GetSettingData(SettingName, "MotherSolution", "PreparationID", .MotherSol.PreparationID)
         .MotherSol.DataMS = GetSettingData(SettingName, "MotherSolution", "DataMS", .MotherSol.DataMS)
         .MotherSol.DataExp = GetSettingData(SettingName, "MotherSolution", "DataExp", .MotherSol.DataExp)
         .MotherSol.QtyProduced = GetSettingData(SettingName, "MotherSolution", "QtyProduced", .MotherSol.QtyProduced)
         .MotherSol.QtyUsed = GetSettingData(SettingName, "MotherSolution", "QtyUsed", .MotherSol.QtyUsed)
         .MotherSol.Unit = GetSettingData(SettingName, "MotherSolution", "Unit", .MotherSol.Unit)
         .MotherSol.WeekPrep = GetSettingData(SettingName, "MotherSolution", "WeekPrep", .MotherSol.WeekPrep)
         

         
         
         .MotherSol.Bottle.EntryBottle = GetSettingData(SettingName, "MotherSolution", "Bottle.EntryBottle", .MotherSol.Bottle.EntryBottle)
         .MotherSol.Bottle.Lot = GetSettingData(SettingName, "MotherSolution", "Bottle.Lot", .MotherSol.Bottle.Lot)
         .MotherSol.Bottle.StockQTY = GetSettingData(SettingName, "MotherSolution", "Bottle.StockQTY", .MotherSol.Bottle.StockQTY)
         .MotherSol.Bottle.ID = GetSettingData(SettingName, "MotherSolution", "Bottle.ID", .MotherSol.Bottle.ID)
         
         .MotherSol.Bottle.MNP = GetSettingData(SettingName, "MotherSolution", "Bottle.MNP", .MotherSol.Bottle.MNP)
         .MotherSol.Bottle.MREXP = GetSettingData(SettingName, "MotherSolution", "Bottle.MREXP", .MotherSol.Bottle.MREXP)
         
         
         
     
     
        If .ID = 0 Then
            
            .ID = GetPreparatonID(.HannaCode.Code, .DataPrep, .HourPrep, .PrepWeek, .bClosed)
        
        End If
        '-----------------------------------------------------------
        ' Recipe in Recipe for production
        '-----------------------------------------------------------
        rc = GetPreparationSingleRecipeFromFile(.Recipe, SettingName)

        CloseSettingDataFile
      
        Call GetPreparationHannaCodeFromFile(.HannaCode, SettingName, PATH)


        
    End With

ERR_END:
    On Error GoTo 0
     CloseSettingDataFile
     PreparationGetSetting = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox Err.Description
     Resume Next
End Function


Public Function GetPreparationSingleRecipeFromFile(ByRef Recipe As RecipeType, ByVal SettingName As String) As Boolean


Dim HannaCodeCount As Integer
Dim MaterialRequisitionCount As Integer
Dim STDcount As Integer
Dim rc As Boolean
On Error GoTo ERR_GET:

    rc = True

    With Recipe
    
      .ActualWeight = GetSettingData(SettingName, "Recipe", "ActualWeight", .ActualWeight)
      .AcquisitionCount = GetSettingData(SettingName, "Recipe", "AcquisitionCount", .AcquisitionCount)
      .bRecalculation = GetSettingData(SettingName, "Recipe", "bRecalculation", .bRecalculation)
      .bUmMassa = GetSettingData(SettingName, "Recipe", "bUmMassa", .bUmMassa)
      .Code = GetSettingData(SettingName, "Recipe", "Code", .Code)
      .Description = GetSettingData(SettingName, "Recipe", "Description", .Description)
      .ID = GetSettingData(SettingName, "Recipe", "ID", .ID)
      .STDcount = GetSettingData(SettingName, "Recipe", "STDcount", .STDcount)
      .STDUnit = GetSettingData(SettingName, "Recipe", "STDUnit", .STDUnit)
      .TotalWeight = GetSettingData(SettingName, "Recipe", "TotalWeight", .TotalWeight)
      .Type = GetSettingData(SettingName, "Recipe", "Type", .Type)
      
        '-----------------------------------------------------------
        ' STD
        '-----------------------------------------------------------
        .STDcount = GetSettingData(SettingName, " STD ", "STDCount", .STDcount)
        If .STDcount >= 0 Then
            STDcount = .STDcount
            ReDim .STD(0)
            Call GetPreparationSTDFromFile(.STD, STDcount, SettingName)
            .STDcount = STDcount
        Else
            
            ' carico i componenti da STD
            'STDCount = .STDCount
            'ReDim .STD(0)
            'Call SetSTDByHannaCode(.STD, .Code, , STDCount)
            '.STDCount = STDCount
            
        End If
        
        
        '-----------------------------------------------------------
        ' Acquisition
        '-----------------------------------------------------------
        
        
        If .AcquisitionCount > 0 Then
            ReDim .Acquisitions(.AcquisitionCount)
            Call GetAcquisitionformFile(.Acquisitions, .AcquisitionCount, SettingName)
        Else
            .AcquisitionCount = GetNumberAcquisitionFormDatabase(SettingName)
            If .AcquisitionCount > 0 Then
            
                SaveSettingData SettingName, "MRCode", "AcquisitionCount", .AcquisitionCount
                ReDim .Acquisitions(.AcquisitionCount)
                Call GetAcquisitionformDatabase(.Acquisitions, .AcquisitionCount, SettingName)
            
            End If
        End If
            
            
            

            
    End With

    CloseSettingDataFile

ERR_END:
    On Error GoTo 0
    GetPreparationSingleRecipeFromFile = rc
    Exit Function
ERR_GET:
    MsgBox Err.Description
    Resume Next


End Function



Private Function GetPreparationSTDFromFile(ByRef STD() As STD, ByRef Count As Integer, SettingName As String)
Dim r As Integer
Dim i As Integer
    CloseSettingDataFile
    
    On Error GoTo ERR_GET:
    r = 1
    
    ReDim STD(Count)
    For i = 1 To Count
    
        With STD(r)
            .NUMBER = GetSettingData(SettingName, " STD " & r, "Number", .NUMBER)
            .Value = GetSettingData(SettingName, " STD " & r, "Value", .Value)
            .TheoreticalWeight = GetSettingData(SettingName, " STD " & r, "TheoreticalWeight", .TheoreticalWeight)
            .MRCode = GetSettingData(SettingName, " STD " & r, "MRCode", .MRCode)
            .RealWeight = GetSettingData(SettingName, " STD " & r, "RealWeight", .RealWeight)
            .ActualWeight = GetSettingData(SettingName, " STD " & r, "ActualWeight", .ActualWeight)
            .Unit = GetSettingData(SettingName, " STD " & r, "Unit", .Unit)
            .Variance = GetSettingData(SettingName, " STD " & r, "Variance", .Variance)
            .VariancePerc = GetSettingData(SettingName, " STD " & r, "VariancePerc", .VariancePerc)
            .bChanged = GetSettingData(SettingName, " STD " & r, "bChanged", .bChanged)
            .bOK = GetSettingData(SettingName, " STD " & r, "bOK", .bOK)
            .Note = GetSettingData(SettingName, " STD " & r, "Note", .Note)
            .STD_ID = GetSettingData(SettingName, " STD " & r, "STD_ID", .STD_ID)

        End With
        
        r = r + 1
cont:
    Next

    Count = r - 1
ERR_END:
    On Error GoTo 0
    
    CloseSettingDataFile
    Exit Function
ERR_GET:
    MsgBox Err.Description
    Resume Next
    
End Function

Private Function GetAcquisitionformFile(ByRef Acquisition() As PrepAcquisition, ByVal AcquisitionCount As Integer, ByVal SettingName As String)
Dim r As Integer
    
    CloseSettingDataFile
    
    
    For r = 1 To AcquisitionCount
        
        With Acquisition(r)
            .Bottle = GetSettingData(SettingName, " Acquisition " & r, "Bottle", .Bottle)
            
            .CodicePipetta = GetSettingData(SettingName, " Acquisition " & r, "CodicePipetta", .CodicePipetta)
            .PipettaType = GetSettingData(SettingName, " Acquisition " & r, "PipettaType", .PipettaType)
            
            .ScaleID = GetSettingData(SettingName, " Acquisition " & r, "ScaleID", .ScaleID)
            .GlasswareID = GetSettingData(SettingName, " Acquisition " & r, "GlassWareID", .GlasswareID)
            .bManuale = GetSettingData(SettingName, " Acquisition " & r, "bManuale", .bManuale)
            
            
            .AcquisitionTime = GetSettingData(SettingName, " Acquisition " & r, "AcquisitionTime", .AcquisitionTime)
            .ActualWeight = GetSettingData(SettingName, " Acquisition " & r, "ActualWeight", .ActualWeight)
            .bFromBarcode = GetSettingData(SettingName, " Acquisition " & r, "bFromBarcode", .bFromBarcode)
            .bDeleted = GetSettingData(SettingName, " Acquisition " & r, "bDeleted", .bDeleted)
            .LeftInBottle = GetSettingData(SettingName, " Acquisition " & r, "LeftInBottle", .LeftInBottle)
            .ID = GetSettingData(SettingName, " Acquisition " & r, "ID", .ID)
            .Index = GetSettingData(SettingName, " Acquisition " & r, "Index", .Index)
            .Note = GetSettingData(SettingName, " Acquisition " & r, "Note", .Note)
            .Operator = GetSettingData(SettingName, " Acquisition " & r, "Operator", .Operator)
            
            .HourPrep = GetSettingData(SettingName, " Acquisition " & r, "HourPrep", .HourPrep)
            
            .Code = GetSettingData(SettingName, " Acquisition " & r, "Code", .Code)
            .DatePrep = GetSettingData(SettingName, " Acquisition " & r, "DatePrep", .DatePrep)
          
          
            .FileName = GetSettingData(SettingName, " Acquisition " & r, "FileName", .FileName)
            .HannaCode = GetSettingData(SettingName, " Acquisition " & r, "HannaCode", .HannaCode)
            .MotherSolutionDate = GetSettingData(SettingName, " Acquisition " & r, "MotherSolutionDate", .MotherSolutionDate)
            .MNP = GetSettingData(SettingName, " Acquisition " & r, "MNP", .MNP)
            .ExpMR = GetSettingData(SettingName, " Acquisition " & r, "ExpMR", .ExpMR)
            
            
            .CodicePipetta = GetSettingData(SettingName, " Acquisition " & r, "CodicePipetta", .CodicePipetta)
            .MRLot = GetSettingData(SettingName, " Acquisition " & r, "MRLot", .MRLot)
            .MsType = GetSettingData(SettingName, " Acquisition " & r, "MSType", .MsType)
            .PreparationID = GetSettingData(SettingName, " Acquisition " & r, "PreparationID", .PreparationID)
            .PrepBarcode.Bottle = GetSettingData(SettingName, " Acquisition " & r, "PrepBarcode.Bottle", .PrepBarcode.Bottle)
            .PrepBarcode.Code = GetSettingData(SettingName, " Acquisition " & r, "PrepBarcode.Code", .PrepBarcode.Code)
            .PrepBarcode.Date = GetSettingData(SettingName, " Acquisition " & r, "PrepBarcode.Date", .PrepBarcode.Date)
            .PrepBarcode.Lot = GetSettingData(SettingName, " Acquisition " & r, "PrepBarcode.Lot", .PrepBarcode.Lot)
            .WeekPrep = GetSettingData(SettingName, " Acquisition " & r, "WeekPrep", .WeekPrep)
            .STDNumber = GetSettingData(SettingName, " Acquisition " & r, "STDNumber", .STDNumber)
            .STDQty = GetSettingData(SettingName, " Acquisition " & r, "STDQty", .STDQty)
            .STDUnit = GetSettingData(SettingName, " Acquisition " & r, "STDUnit", .STDUnit)
            .STDValue = GetSettingData(SettingName, " Acquisition " & r, "STDValue", .STDValue)
          
        End With
    Next

   
    
    CloseSettingDataFile
    
End Function

Public Function GetPreparationHannaCodeFromFile(ByRef HannaCode As HannaCode, ByVal SettingName As String, ByVal PATH As String)
Dim i As Integer
Dim t As Integer
Dim bHide As Boolean
Dim RecipeForHannaCode As String

   
        

        With HannaCode

            .Code = GetSettingData(SettingName, "HannaCode", "Code", .Code, PATH)
            .ConcHannaParameter = GetSettingData(SettingName, "HannaCode", "ConcHannaParameter", .ConcHannaParameter, PATH)
            .Decimal = GetSettingData(SettingName, "HannaCode", "Decimal", .Decimal, PATH)
            .Description = GetSettingData(SettingName, "HannaCode", "Description", .Description, PATH)
            .FWHannaParameter = GetSettingData(SettingName, "HannaCode", "FWHannaParameter", .FWHannaParameter, PATH)
            .ID = GetSettingData(SettingName, "HannaCode", "ID", .ID, PATH)
            .MeasurementUnit = GetSettingData(SettingName, "HannaCode", "MeasurementUnit", .MeasurementUnit, PATH)
            .MR.Code = GetSettingData(SettingName, "HannaCode", ".MR.Code", .MR.Code, PATH)
            .MS1val = GetSettingData(SettingName, "HannaCode", "MS1val", .MS1val, PATH)
            .MS1vol = GetSettingData(SettingName, "HannaCode", "MS1vol", .MS1vol, PATH)
            .MS2Dil = GetSettingData(SettingName, "HannaCode", "MS2dil", .MS2Dil, PATH)
            .MS2vol = GetSettingData(SettingName, "HannaCode", "MS2vol", .MS2vol, PATH)
            .MSEXP = GetSettingData(SettingName, "HannaCode", "MSEXP", .MSEXP, PATH)
            .Hannaformula = GetSettingData(SettingName, "HannaCode", "Hannaformula", .Hannaformula, PATH)
            .ParameterMethod = GetSettingData(SettingName, "HannaCode", "ParameterMethod", .ParameterMethod, PATH)
            .QtyToProduce = GetSettingData(SettingName, "HannaCode", "QtyToProduce", .QtyToProduce, PATH)
            .RangeMax = GetSettingData(SettingName, "HannaCode", "RangeMax", .RangeMax, PATH)
            .RangeMin = GetSettingData(SettingName, "HannaCode", "RangeMin", .RangeMin, PATH)
            .Recipe = GetSettingData(SettingName, "HannaCode", "Recipe", .Recipe, PATH)
            .STDcount = GetSettingData(SettingName, "HannaCode", "STDcount", .STDcount, PATH)
            .STDExp = GetSettingData(SettingName, "HannaCode", "STDExp", .STDExp, PATH)
            .STDMatrix = GetSettingData(SettingName, "HannaCode", "STDMatrix", .STDMatrix, PATH)
            .STDMR2 = GetSettingData(SettingName, "HannaCode", "STDMR2", .STDMR2, PATH)
            .STDNote = GetSettingData(SettingName, "HannaCode", "STDNote", .STDNote, PATH)
            .STDStorage = GetSettingData(SettingName, "HannaCode", "STDStorage", .STDStorage, PATH)
            .STDType = GetSettingData(SettingName, "HannaCode", "STDType", .STDType, PATH)
            .STDUnit = GetSettingData(SettingName, "HannaCode", "STDUnit", .STDUnit, PATH)
            .STDVolume = GetSettingData(SettingName, "HannaCode", "STDVolume", .STDVolume, PATH)
            .UnitMR = GetSettingData(SettingName, "HannaCode", "UnitMR", .UnitMR, PATH)
         
        End With
        
        With HannaCode
        
            ReDim .STD(.STDcount)
            
         For i = 1 To .STDcount
            .STD(i).ActualWeight = GetSettingData(SettingName, "HannaCode STD " & i, "ActualWeight", .STD(i).ActualWeight, PATH)
            .STD(i).bChanged = GetSettingData(SettingName, "HannaCode STD " & i, "bChanged", .STD(i).bChanged, PATH)
            .STD(i).bOK = GetSettingData(SettingName, "HannaCode STD " & i, "bOK", .STD(i).bOK, PATH)
            .STD(i).Note = GetSettingData(SettingName, "HannaCode STD " & i, "Note", .STD(i).Note, PATH)
            .STD(i).NUMBER = GetSettingData(SettingName, "HannaCode STD " & i, "Number", .STD(i).NUMBER, PATH)
            .STD(i).RealWeight = GetSettingData(SettingName, "HannaCode STD " & i, "RealWeight", .STD(i).RealWeight, PATH)
            .STD(i).TheoreticalWeight = GetSettingData(SettingName, "HannaCode STD " & i, "TheoreticalWeight", .STD(i).TheoreticalWeight, PATH)
            .STD(i).MRCode = GetSettingData(SettingName, "HannaCode STD " & i, "MRCode", .STD(i).MRCode, PATH)
            .STD(i).Unit = GetSettingData(SettingName, "HannaCode STD " & i, "Unit", .STD(i).Unit, PATH)
            .STD(i).Value = GetSettingData(SettingName, "HannaCode STD " & i, "Value", .STD(i).Value, PATH)
            .STD(i).Variance = GetSettingData(SettingName, "HannaCode STD " & i, "Variance", .STD(i).Variance, PATH)
            .STD(i).VariancePerc = GetSettingData(SettingName, "HannaCode STD " & i, "VariancePerc", .STD(i).VariancePerc, PATH)
            .STD(i).STD_ID = GetSettingData(SettingName, "HannaCode STD " & i, "STD_ID", .STD(i).STD_ID, PATH)
        Next
    End With
    

  
        
        With HannaCode.MR
        
            
            .bMassa = GetSettingData(SettingName, "MR = " & .Code, "bMassa", .bMassa, PATH)
            .Classification = GetSettingData(SettingName, "MR = " & .Code, "Classification", .Classification, PATH)
            .Code = GetSettingData(SettingName, "MR = " & .Code, "Code", .Code, PATH)
            .Density = GetSettingData(SettingName, "MR = " & .Code, "Density", .Density, PATH)
            .Description = GetSettingData(SettingName, "MR = " & .Code, "Description", .Description, PATH)
            .FWParameter = GetSettingData(SettingName, "MR = " & .Code, "FWParameter", .FWParameter, PATH)
            .Location = GetSettingData(SettingName, "MR = " & .Code, "Location", .Location, PATH)
            .MinQty = GetSettingData(SettingName, "MR = " & .Code, "MinQty", .MinQty, PATH)
            .MNP = GetSettingData(SettingName, "MR = " & .Code, "MNP", .MNP, PATH)
            .Modified = GetSettingData(SettingName, "MR = " & .Code, "Modified", .Modified, PATH)
            .MRPurity = GetSettingData(SettingName, "MR = " & .Code, "MRPurity", .MRPurity, PATH)
         
            .MRValue = GetSettingData(SettingName, "MR = " & .Code, "MRValue", .MRValue, PATH)
            .Parameter = GetSettingData(SettingName, "MR = " & .Code, "Parameter", .Parameter, PATH)
            .PhysicalState = GetSettingData(SettingName, "MR = " & .Code, "PhysicalState", .PhysicalState, PATH)
          
            .ReductionExpDays = GetSettingData(SettingName, "MR = " & .Code, "ReductionExpDays", 120, PATH)
            If Trim(.ReductionExpDays) = "" Then .ReductionExpDays = 120
            .Rev = GetSettingData(SettingName, "MR = " & .Code, "Rev", .Rev, PATH)
            .STOCK_QTY = GetSettingData(SettingName, "MR = " & .Code, "STOCK_QTY", .STOCK_QTY, PATH)
            .STOCK_UNIT = GetSettingData(SettingName, "MR = " & .Code, "STOCK_UNIT", .STOCK_UNIT, PATH)
            .StorageT = GetSettingData(SettingName, "MR = " & .Code, "StorageT", .StorageT, PATH)
            .Supplier = GetSettingData(SettingName, "MR = " & .Code, "Supplier", .Supplier, PATH)
            .Unit = GetSettingData(SettingName, "MR = " & .Code, "Unit", .Unit, PATH)
            .Code = GetSettingData(SettingName, "MR = " & .Code, "Code", .Code, PATH)
           
            If .MRPurity = 0 And .MRValue = 0 And .Code <> "" Then
                
                Call SetMRFromDatabase(.Code, HannaCode.MR)
            
            End If
            
            
        End With
        
    
        

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

Private Function GetAcquisitionformDatabase(ByRef Acquisition() As PrepAcquisition, ByVal AcquisitionCount As Integer, ByVal SettingName As String)
Dim r As Integer
    
    CloseSettingDataFile
    
    dbTabAcquisition.MoveFirst
    
    For r = 1 To AcquisitionCount
    
        
        
        With Acquisition(r)
        
            ' get form database
            
            .AcquisitionTime = IIf(IsNull(Trim(dbTabAcquisition!AcquisitionTime)), "", Trim(dbTabAcquisition!AcquisitionTime))
            
            .CodicePipetta = IIf(IsNull(Trim(dbTabAcquisition!CodicePipetta)), "", Trim(dbTabAcquisition!CodicePipetta))
            .PipettaType = IIf(IsNull(Trim(dbTabAcquisition!PipettaType)), "", Trim(dbTabAcquisition!PipettaType))
            
            .ScaleID = IIf(IsNull(Trim(dbTabAcquisition!ScaleID)), "", Trim(dbTabAcquisition!ScaleID))
            .GlasswareID = IIf(IsNull(Trim(dbTabAcquisition!GlasswareID)), "", Trim(dbTabAcquisition!GlasswareID))
            .bManuale = IIf(IsNull(Trim(dbTabAcquisition!bManuale)), False, Trim(dbTabAcquisition!bManuale))
            
            
            .ActualWeight = IIf(IsNull(Trim(dbTabAcquisition!ActualWeight)), "", Trim(dbTabAcquisition!ActualWeight))
            .ID = dbTabAcquisition!ID
            .Index = IIf(IsNull(Trim(dbTabAcquisition!Index)), 0, Trim(dbTabAcquisition!Index))
            .Note = IIf(IsNull(Trim(dbTabAcquisition!Note)), "", Trim(dbTabAcquisition!Note))
            .Operator = IIf(IsNull(Trim(dbTabAcquisition!Operator)), "", Trim(dbTabAcquisition!Operator))
            .LeftInBottle = IIf(IsNull(Trim(dbTabAcquisition!LeftInBottle)), "", Trim(dbTabAcquisition!LeftInBottle))
        
            .Code = IIf(IsNull(Trim(dbTabAcquisition!Code)), "", Trim(dbTabAcquisition!Code))
            .DatePrep = IIf(IsNull(Trim(dbTabAcquisition!DatePrep)), "", Trim(dbTabAcquisition!DatePrep))
            .HourPrep = IIf(IsNull(Trim(dbTabAcquisition!HourPrep)), "", Trim(dbTabAcquisition!HourPrep))
        
            .FileName = IIf(IsNull(Trim(dbTabAcquisition!FileName)), "", Trim(dbTabAcquisition!FileName))
            .HannaCode = IIf(IsNull(Trim(dbTabAcquisition!HannaCode)), "", Trim(dbTabAcquisition!HannaCode))
            .MotherSolutionDate = IIf(IsNull(Trim(dbTabAcquisition!MotherSolutionDate)), "", Trim(dbTabAcquisition!MotherSolutionDate))
            
            .MNP = IIf(IsNull(Trim(dbTabAcquisition!MNP)), "", Trim(dbTabAcquisition!MNP))
            .ExpMR = IIf(IsNull(Trim(dbTabAcquisition!ExpMR)), "", Trim(dbTabAcquisition!ExpMR))
            
            .CodicePipetta = IIf(IsNull(Trim(dbTabAcquisition!CodicePipetta)), "", Trim(dbTabAcquisition!CodicePipetta))
            .MRLot = IIf(IsNull(Trim(dbTabAcquisition!MRLot)), "", Trim(dbTabAcquisition!MRLot))
            .MsType = IIf(IsNull(Trim(dbTabAcquisition!MsType)), "", Trim(dbTabAcquisition!MsType))
            .PreparationID = IIf(IsNull(Trim(dbTabAcquisition!PreparationID)), "", Trim(dbTabAcquisition!PreparationID))
            .WeekPrep = IIf(IsNull(Trim(dbTabAcquisition!WeekPrep)), "", Trim(dbTabAcquisition!WeekPrep))
            .STDNumber = IIf(IsNull(Trim(dbTabAcquisition!STDNumber)), "", Trim(dbTabAcquisition!STDNumber))
            .STDQty = IIf(IsNull(Trim(dbTabAcquisition!STDQty)), "", Trim(dbTabAcquisition!STDQty))
            .STDUnit = IIf(IsNull(Trim(dbTabAcquisition!STDUnit)), "", Trim(dbTabAcquisition!STDUnit))
            .STDValue = IIf(IsNull(Trim(dbTabAcquisition!STDValue)), "", Trim(dbTabAcquisition!STDValue))
            
          
            dbTabAcquisition.MoveNext
            
        End With
        
        With Acquisition(r)
            
            'save on file!!!
            
            SaveSettingData SettingName, " Acquisition " & r, "AcquisitionTime", .AcquisitionTime
            
            SaveSettingData SettingName, " Acquisition " & r, "CodicePipetta", .CodicePipetta
            SaveSettingData SettingName, " Acquisition " & r, "PipettaType", .PipettaType
            
            SaveSettingData SettingName, " Acquisition " & r, "ScaleID", .ScaleID
            SaveSettingData SettingName, " Acquisition " & r, "GlassWareID", .GlasswareID
            SaveSettingData SettingName, " Acquisition " & r, "bManuale", .bManuale
            
           
           
            SaveSettingData SettingName, " Acquisition " & r, "ActualWeight", .ActualWeight
            SaveSettingData SettingName, " Acquisition " & r, "bFromBarcode", .bFromBarcode
            SaveSettingData SettingName, " Acquisition " & r, "bDeleted", .bDeleted
            SaveSettingData SettingName, " Acquisition " & r, "LeftInBottle", .LeftInBottle
            SaveSettingData SettingName, " Acquisition " & r, "Bottle", .Bottle
            SaveSettingData SettingName, " Acquisition " & r, "Code", .Code
            
            SaveSettingData SettingName, " Acquisition " & r, "DatePrep", .DatePrep
            SaveSettingData SettingName, " Acquisition " & r, "HourPrep", .HourPrep
            SaveSettingData SettingName, " Acquisition " & r, "WeekPrep", .WeekPrep
            SaveSettingData SettingName, " Acquisition " & r, "FileName", .FileName
            SaveSettingData SettingName, " Acquisition " & r, "HannaCode", .HannaCode
            SaveSettingData SettingName, " Acquisition " & r, "ID", .ID
            SaveSettingData SettingName, " Acquisition " & r, "Index", .Index
            SaveSettingData SettingName, " Acquisition " & r, "MotherSolutionDate", .MotherSolutionDate
            
            SaveSettingData SettingName, " Acquisition " & r, "MNP", .MNP
            SaveSettingData SettingName, " Acquisition " & r, "ExpMR", .ExpMR
            
            
            SaveSettingData SettingName, " Acquisition " & r, "CodicePipetta", .CodicePipetta
            SaveSettingData SettingName, " Acquisition " & r, "MRLot", .MRLot
            SaveSettingData SettingName, " Acquisition " & r, "MSType", .MsType
            SaveSettingData SettingName, " Acquisition " & r, "Note", .Note
            SaveSettingData SettingName, " Acquisition " & r, "Operator", .Operator
            
            SaveSettingData SettingName, " Acquisition " & r, "PreparationID", .PreparationID
            SaveSettingData SettingName, " Acquisition " & r, "PrepBarcode.Bottle", .PrepBarcode.Bottle
            SaveSettingData SettingName, " Acquisition " & r, "PrepBarcode.Code", .PrepBarcode.Code
            SaveSettingData SettingName, " Acquisition " & r, "PrepBarcode.Date", .PrepBarcode.Date
            SaveSettingData SettingName, " Acquisition " & r, "PrepBarcode.Lot", .PrepBarcode.Lot
            
            SaveSettingData SettingName, " Acquisition " & r, "STDNumber", .STDNumber
            SaveSettingData SettingName, " Acquisition " & r, "STDQty", .STDQty
            SaveSettingData SettingName, " Acquisition " & r, "STDUnit", .STDUnit
            SaveSettingData SettingName, " Acquisition " & r, "STDValue", .STDValue
            
        End With
        
    Next

   
    
    CloseSettingDataFile
    
End Function
