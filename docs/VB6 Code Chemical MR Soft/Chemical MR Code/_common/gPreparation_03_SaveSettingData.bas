Attribute VB_Name = "gPreparation_03_SaveSettingData"
Option Explicit


Private SettingName As String

Public Function PeparationSaveSetting(ByRef iPreparation As RecipeForProduction, ByVal SettName As String) As Boolean
Dim rc As Boolean
Dim RecipeCount As Integer
Dim HannaCodeCount As Integer
Dim TotalsCount As Integer
Dim PackagingCount As Integer
Dim FileName As String
Dim PATH As String

On Error GoTo ERR_SAVE

    SettingName = SettName
    
    rc = True

    If USER_PATH = "" Then USER_PATH = USER_TEMP_PATH
    
    
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
        SaveSettingData SettingName, "iRecipeForProduction", "ID", .ID
        SaveSettingData SettingName, "iRecipeForProduction", "HannaCode", .HannaCode.Code
        SaveSettingData SettingName, "iRecipeForProduction", "Description", .HannaCode.Description
        SaveSettingData SettingName, "iRecipeForProduction", "MRCode", .MRCode
        SaveSettingData SettingName, "iRecipeForProduction", "DataPrep", .DataPrep
        SaveSettingData SettingName, "iRecipeForProduction", "HourPrep", .HourPrep
        SaveSettingData SettingName, "iRecipeForProduction", "PrepWeek", .PrepWeek
        SaveSettingData SettingName, "iRecipeForProduction", "Operator", .Operator
        
        SaveSettingData SettingName, "iRecipeForProduction", "QtyToProduce", .QtyToProduce
        SaveSettingData SettingName, "iRecipeForProduction", "ManualValue", .ManualValue
        SaveSettingData SettingName, "iRecipeForProduction", "ManualUnit", .ManualUnit
        SaveSettingData SettingName, "iRecipeForProduction", "STD_Manual_ID", .STD_Manual_ID
        
        
        
        SaveSettingData SettingName, "iRecipeForProduction", "QtyProduced", .QtyProduced
        SaveSettingData SettingName, "iRecipeForProduction", "Unit", .Unit
        SaveSettingData SettingName, "iRecipeForProduction", "STDMatrix", .HannaCode.STDMatrix
        SaveSettingData SettingName, "iRecipeForProduction", "STDExp", .HannaCode.STDExp
        
        SaveSettingData SettingName, "iRecipeForProduction", "STDStorage", .HannaCode.STDStorage
        SaveSettingData SettingName, "iRecipeForProduction", "bClosed", .bClosed
        SaveSettingData SettingName, "iRecipeForProduction", "CloseDate", .CloseDate
        
        SaveSettingData SettingName, "iRecipeForProduction", "FileName", .FileName
       
        SaveSettingData SettingName, "iRecipeForProduction", "MSType", .MsType
        SaveSettingData SettingName, "iRecipeForProduction", "MotherSolID", .MotherSol.ID
       
        SaveSettingData SettingName, "iRecipeForProduction", "Note", .Note
        SaveSettingData SettingName, "iRecipeForProduction", "bCorrection", .bCorrection
      
        SaveSettingData SettingName, "iRecipeForProduction", "ExpDate", .ExpDate
        
        SaveSettingData SettingName, "iRecipeForProduction", ".MS.DilConc", .MS.DilConc
        SaveSettingData SettingName, "iRecipeForProduction", ".MS.Exp", .MS.Exp
        SaveSettingData SettingName, "iRecipeForProduction", ".MS.Qty", .MS.Qty
        SaveSettingData SettingName, "iRecipeForProduction", ".MS.Unit", .MS.Unit
        SaveSettingData SettingName, "iRecipeForProduction", ".MS.Value", .MS.Value
        SaveSettingData SettingName, "iRecipeForProduction", ".MS.Volume", .MS.Volume
        
        
            
        '-----------------------------------------------------------
        ' MotherSolution
        '-----------------------------------------------------------
        
        SaveSettingData SettingName, "MotherSolution", "bClosed", .MotherSol.bClosed
        SaveSettingData SettingName, "MotherSolution", "Code", .MotherSol.Code
        SaveSettingData SettingName, "MotherSolution", "DataPrep", .MotherSol.DataPrep
        SaveSettingData SettingName, "MotherSolution", "ExpDays", .MotherSol.ExpDays
        SaveSettingData SettingName, "MotherSolution", "HannaCode", .MotherSol.HannaCode
        SaveSettingData SettingName, "MotherSolution", "HourPrep", .MotherSol.HourPrep
        SaveSettingData SettingName, "MotherSolution", "ID", .MotherSol.ID
        SaveSettingData SettingName, "MotherSolution", "MSType", .MotherSol.MsType
        SaveSettingData SettingName, "MotherSolution", "Note", .MotherSol.Note
        SaveSettingData SettingName, "MotherSolution", "Operator", .MotherSol.Operator
        SaveSettingData SettingName, "MotherSolution", "QtyLeft", .MotherSol.QtyLeft
        SaveSettingData SettingName, "MotherSolution", "PreparationID", .MotherSol.PreparationID
        SaveSettingData SettingName, "MotherSolution", "DataMS", .MotherSol.DataMS
        SaveSettingData SettingName, "MotherSolution", "DataExp", .MotherSol.DataExp
        SaveSettingData SettingName, "MotherSolution", "QtyProduced", .MotherSol.QtyProduced
        SaveSettingData SettingName, "MotherSolution", "QtyUsed", .MotherSol.QtyUsed
        SaveSettingData SettingName, "MotherSolution", "Unit", .MotherSol.Unit
        SaveSettingData SettingName, "MotherSolution", "WeekPrep", .MotherSol.WeekPrep
        
        SaveSettingData SettingName, "MotherSolution", "Bottle.EntryBottle", .MotherSol.Bottle.EntryBottle
        SaveSettingData SettingName, "MotherSolution", "Bottle.Lot", .MotherSol.Bottle.Lot
        SaveSettingData SettingName, "MotherSolution", "Bottle.StockQTY", .MotherSol.Bottle.StockQTY
        SaveSettingData SettingName, "MotherSolution", "Bottle.ID", .MotherSol.Bottle.ID
        
        SaveSettingData SettingName, "MotherSolution", "Bottle.MNP", .MotherSol.Bottle.MNP
        SaveSettingData SettingName, "MotherSolution", "Bottle.MREXP", .MotherSol.Bottle.MREXP
        
        

            
        CloseSettingDataFile
      
        
        '-----------------------------------------------------------
        ' Recipe
        '-----------------------------------------------------------
cont:

        Call SetPreparationRecipeInFile(.Recipe)

        '-----------------------------------------------------------
        ' HANNA CODE
        '-----------------------------------------------------------

            Call SetHannaCodeInFile(.HannaCode, .FileName, PATH)

       
     
    End With

ERR_END:
    On Error GoTo 0
     CloseSettingDataFile
     PeparationSaveSetting = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox Err.Description
     Resume Next
End Function



Public Function SetPreparationRecipeInFile(ByRef Recipe As RecipeType) As Boolean

Dim i As Integer
Dim HannaCodeCount As Integer
Dim MaterialRequisitionCount As Integer
Dim STDcount As Integer

    
    With Recipe
    
        SaveSettingData SettingName, "Recipe", "AcquisitionCount", .AcquisitionCount
        SaveSettingData SettingName, "Recipe", "ActualWeight", .ActualWeight
        SaveSettingData SettingName, "Recipe", "bRecalculation", .bRecalculation
        SaveSettingData SettingName, "Recipe", "bUmMassa", .bUmMassa
        SaveSettingData SettingName, "Recipe", "Code", .Code
        SaveSettingData SettingName, "Recipe", "Description", .Description
        SaveSettingData SettingName, "Recipe", "ID", .ID
        SaveSettingData SettingName, "Recipe", "STDcount", .STDcount
        SaveSettingData SettingName, "Recipe", "STDUnit", .STDUnit
        SaveSettingData SettingName, "Recipe", "TotalWeight", .TotalWeight
        SaveSettingData SettingName, "Recipe", "Type", .Type
        

        '-----------------------------------------------------------
        ' STD
        '-----------------------------------------------------------
        If .STDcount >= 0 Then
            STDcount = .STDcount
            SaveSettingData SettingName, " STD ", "STDCount", .STDcount
            Call SetPreparationSTDInFile(.STD, STDcount)
        End If
        
        SaveSettingData SettingName, " STD ", "STDCount", .STDcount
        
        '-----------------------------------------------------------
        ' Acquisition
        '-----------------------------------------------------------
        
        If .AcquisitionCount > 0 Then
            Call SetAcquisitionInFile(i, .Acquisitions, .AcquisitionCount)
        End If
            
            
        SaveSettingData SettingName, "Recipe", "AcquisitionCount", .AcquisitionCount
        
    

    End With

CloseSettingDataFile


End Function


Private Function SetPreparationSTDInFile(ByRef STD() As STD, ByRef Count As Integer)
Dim r As Integer
Dim i As Integer
    CloseSettingDataFile
    r = 1
    
    For i = 1 To Count
    
        With STD(r)
            SaveSettingData SettingName, " STD " & r, "Number", .NUMBER
            SaveSettingData SettingName, " STD " & r, "Value", .Value
            SaveSettingData SettingName, " STD " & r, "TheoreticalWeight", .TheoreticalWeight
            SaveSettingData SettingName, " STD " & r, "MRCode", .MRCode
            
            SaveSettingData SettingName, " STD " & r, "RealWeight", .RealWeight
            SaveSettingData SettingName, " STD " & r, "ActualWeight", .ActualWeight
            SaveSettingData SettingName, " STD " & r, "Unit", .Unit
            SaveSettingData SettingName, " STD " & r, "Variance", .Variance
            SaveSettingData SettingName, " STD " & r, "VariancePerc", .VariancePerc
            SaveSettingData SettingName, " STD " & r, "bChanged", .bChanged
            SaveSettingData SettingName, " STD " & r, "bOK", .bOK
            SaveSettingData SettingName, " STD " & r, "Note", .Note
            SaveSettingData SettingName, " STD " & r, "STD_ID", .STD_ID
            
            

        End With
        
        r = r + 1
cont:
    Next

    Count = r - 1
    
    CloseSettingDataFile
    
End Function



Private Function SetAcquisitionInFile(ByVal t As Integer, ByRef Acquisition() As PrepAcquisition, ByRef AcquisitionCount As Integer)
Dim r As Integer
Dim i As Integer

On Error GoTo ERR_SET:

    CloseSettingDataFile
    r = 1
    
    For i = 1 To AcquisitionCount
    
        If Acquisition(i).bDeleted Then GoTo cont
    
        With Acquisition(r)
            
            SaveSettingData SettingName, " Acquisition " & r, "CodicePipetta", .CodicePipetta
            SaveSettingData SettingName, " Acquisition " & r, "PipettaType", .PipettaType
            
            SaveSettingData SettingName, " Acquisition " & r, "ScaleID", .ScaleID
            SaveSettingData SettingName, " Acquisition " & r, "GlassWareID", .GlasswareID
            SaveSettingData SettingName, " Acquisition " & r, "bManuale", .bManuale
            
        
            SaveSettingData SettingName, " Acquisition " & r, "AcquisitionTime", .AcquisitionTime
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
        
        r = r + 1
cont:
    Next

    AcquisitionCount = r - 1
ERR_END:
    On Error GoTo 0
    CloseSettingDataFile
    Exit Function
ERR_SET:
    MsgBox Err.Description
    Resume Next
End Function
Public Function SetHannaCodeInFile(ByRef HannaCode As HannaCode, ByVal SettName As String, ByVal PATH As String)
Dim i As Integer



    CloseSettingDataFile

On Error GoTo ERR_SET:


        With HannaCode
        
            SaveSettingData SettName, "HannaCode", "Code", .Code, PATH
            SaveSettingData SettName, "HannaCode", "ConcHannaParameter", .ConcHannaParameter, PATH
            SaveSettingData SettName, "HannaCode", "Decimal", .Decimal, PATH
            SaveSettingData SettName, "HannaCode", "Description", .Description, PATH
            SaveSettingData SettName, "HannaCode", "FWHannaParameter", .FWHannaParameter, PATH
            SaveSettingData SettName, "HannaCode", "ID", .ID, PATH
            SaveSettingData SettName, "HannaCode", "MeasurementUnit", .MeasurementUnit, PATH
            SaveSettingData SettName, "HannaCode", ".MR.Code", .MR.Code, PATH
            SaveSettingData SettName, "HannaCode", "MS1val", .MS1val, PATH
            SaveSettingData SettName, "HannaCode", "MS1vol", .MS1vol, PATH
            SaveSettingData SettName, "HannaCode", "MS2dil", .MS2Dil, PATH
            SaveSettingData SettName, "HannaCode", "MS2vol", .MS2vol, PATH
            SaveSettingData SettName, "HannaCode", "MSEXP", .MSEXP, PATH
            SaveSettingData SettName, "HannaCode", "Hannaformula", .Hannaformula, PATH
            SaveSettingData SettName, "HannaCode", "ParameterMethod", .ParameterMethod, PATH
            SaveSettingData SettName, "HannaCode", "QtyToProduce", .QtyToProduce, PATH
            SaveSettingData SettName, "HannaCode", "RangeMax", .RangeMax, PATH
            SaveSettingData SettName, "HannaCode", "RangeMin", .RangeMin, PATH
            SaveSettingData SettName, "HannaCode", "STDcount", .STDcount, PATH
            SaveSettingData SettName, "HannaCode", "Recipe", .Recipe, PATH
            SaveSettingData SettName, "HannaCode", "STDExp", .STDExp, PATH
            SaveSettingData SettName, "HannaCode", "STDMatrix", .STDMatrix, PATH
            SaveSettingData SettName, "HannaCode", "STDMR2", .STDMR2, PATH
            SaveSettingData SettName, "HannaCode", "STDNote", .STDNote, PATH
            SaveSettingData SettName, "HannaCode", "STDStorage", .STDStorage, PATH
            SaveSettingData SettName, "HannaCode", "STDType", .STDType, PATH
            SaveSettingData SettName, "HannaCode", "STDUnit", .STDUnit, PATH
            SaveSettingData SettName, "HannaCode", "STDVolume", .STDVolume, PATH
            SaveSettingData SettName, "HannaCode", "UnitMR", .UnitMR, PATH
          
        End With
        
         
    With HannaCode
         For i = 1 To UBound(.STD)
            SaveSettingData SettName, "HannaCode STD " & i, "ActualWeight", .STD(i).ActualWeight, PATH
            SaveSettingData SettName, "HannaCode STD " & i, "bChanged", .STD(i).bChanged, PATH
            SaveSettingData SettName, "HannaCode STD " & i, "bOK", .STD(i).bOK, PATH
            SaveSettingData SettName, "HannaCode STD " & i, "Note", .STD(i).Note, PATH
            SaveSettingData SettName, "HannaCode STD " & i, "NUMBER", .STD(i).NUMBER, PATH
            SaveSettingData SettName, "HannaCode STD " & i, "RealWeight", .STD(i).RealWeight, PATH
            SaveSettingData SettName, "HannaCode STD " & i, "TheoreticalWeight", .STD(i).TheoreticalWeight, PATH
            
            SaveSettingData SettName, "HannaCode STD " & i, "MRCode", .STD(i).MRCode, PATH
            
            SaveSettingData SettName, "HannaCode STD " & i, "Unit", .STD(i).Unit, PATH
            SaveSettingData SettName, "HannaCode STD " & i, "Value", .STD(i).Value, PATH
            SaveSettingData SettName, "HannaCode STD " & i, "Variance", .STD(i).Variance, PATH
            SaveSettingData SettName, "HannaCode STD " & i, "VariancePerc", .STD(i).VariancePerc, PATH
            SaveSettingData SettName, "HannaCode STD " & i, "STD_ID", .STD(i).STD_ID, PATH
            
            
        Next
    End With
    
        With HannaCode.MR
            SaveSettingData SettName, "MR = " & .Code, "bMassa", .bMassa, PATH
            SaveSettingData SettName, "MR = " & .Code, "Classification", .Classification, PATH
            SaveSettingData SettName, "MR = " & .Code, "Code", .Code, PATH
            SaveSettingData SettName, "MR = " & .Code, "Density", .Density, PATH
            SaveSettingData SettName, "MR = " & .Code, "Description", .Description, PATH
            SaveSettingData SettName, "MR = " & .Code, "FWParameter", .FWParameter, PATH
            SaveSettingData SettName, "MR = " & .Code, "Location", .Location, PATH
            SaveSettingData SettName, "MR = " & .Code, "MinQty", .MinQty, PATH
            SaveSettingData SettName, "MR = " & .Code, "MNP", .MNP, PATH
            SaveSettingData SettName, "MR = " & .Code, "Modified", .Modified, PATH
            SaveSettingData SettName, "MR = " & .Code, "MRPurity", .MRPurity, PATH
        
            SaveSettingData SettName, "MR = " & .Code, "MRValue", .MRValue, PATH
            SaveSettingData SettName, "MR = " & .Code, "Parameter", .Parameter, PATH
            SaveSettingData SettName, "MR = " & .Code, "PhysicalState", .PhysicalState, PATH
          
            SaveSettingData SettName, "MR = " & .Code, "ReductionExpDays", .ReductionExpDays, PATH
            SaveSettingData SettName, "MR = " & .Code, "Rev", .Rev, PATH
            SaveSettingData SettName, "MR = " & .Code, "STOCK_QTY", .STOCK_QTY, PATH
            SaveSettingData SettName, "MR = " & .Code, "STOCK_UNIT", .STOCK_UNIT, PATH
            SaveSettingData SettName, "MR = " & .Code, "StorageT", .StorageT, PATH
            SaveSettingData SettName, "MR = " & .Code, "Supplier", .Supplier, PATH
            SaveSettingData SettName, "MR = " & .Code, "Unit", .Unit, PATH
        End With
        
ERR_END:
  
    CloseSettingDataFile
    Exit Function
ERR_SET:
    MsgBox Err.Description
    Resume Next
    
End Function


