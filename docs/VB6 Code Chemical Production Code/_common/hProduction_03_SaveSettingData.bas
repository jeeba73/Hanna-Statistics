Attribute VB_Name = "hProduction_03_SaveSettingData"
Option Explicit


Private SettingName As String

Public Function ProductionSaveSetting(ByRef iProduction As RecipeForProduction, ByVal SettName As String) As Boolean
Dim rc As Boolean
Dim RecipeCount As Integer
Dim HannaCodesCount As Integer
Dim TotalsCount As Integer
Dim PackagingCount As Integer

On Error GoTo ERR_SAVE

    SettingName = SettName
    
    rc = True


   ' USER_PATH = USER_PRODUCTION_PATH
    
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

    
    
    With iProduction
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
        
        SaveSettingData SettingName, "iRecipeForProduction", "WeekProd", .WeekProd

         SaveSettingData SettingName, "iRecipeForProduction", "ProductionID", .ProductionID
         
        '-----------------------------------------------------------
        ' HANNA CODES in Recipe for production
        '-----------------------------------------------------------
        If .HannaCodesCount > 0 Then
            HannaCodesCount = .HannaCodesCount
            SaveSettingData SettingName, "HannaCodes", "HannaCodesCount", .HannaCodesCount
            Call SetHannaCodesInFile(.HannaCodes, HannaCodesCount)
        End If
     

    End With

ERR_END:
    On Error GoTo 0
     CloseSettingDataFile
     ProductionSaveSetting = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function



Private Function SetHannaCodesInFile(ByRef HannaCodes() As HannaCode, ByVal HannaCodesCount As Integer)
Dim i As Integer
    For i = 1 To HannaCodesCount
        
        With HannaCodes(i)
            SaveSettingData SettingName, "HannaCode" & i, "bHide", .bHide
            SaveSettingData SettingName, "HannaCode" & i, "Code", .Code
            SaveSettingData SettingName, "HannaCode" & i, "Density", .Density
            SaveSettingData SettingName, "HannaCode" & i, "Exp", .Exp
            SaveSettingData SettingName, "HannaCode" & i, "ExpDate", .ExpDate
            SaveSettingData SettingName, "HannaCode" & i, "ID", .ID
            SaveSettingData SettingName, "HannaCode" & i, "LastLot", .LastLot
            SaveSettingData SettingName, "HannaCode" & i, "Line", .Line
            'SaveSettingData SettingName, "HannaCode" & i, "LoadInPrint", '.LoadInPrint
            SaveSettingData SettingName, "HannaCode" & i, "MaxQty", .MaxQty
            SaveSettingData SettingName, "HannaCode" & i, "MinQty", .MinQty
            SaveSettingData SettingName, "HannaCode" & i, "Mix1", .Mix1
            SaveSettingData SettingName, "HannaCode" & i, "Mix2", .Mix2
            SaveSettingData SettingName, "HannaCode" & i, "Procedure", .Procedure
            SaveSettingData SettingName, "HannaCode" & i, "ProcedureRev", .ProcedureRev
            SaveSettingData SettingName, "HannaCode" & i, "ProductName", .ProductName
            SaveSettingData SettingName, "HannaCode" & i, "Qty", .Qty
            SaveSettingData SettingName, "HannaCode" & i, "QtyToProduce", .QtyToProduce
            SaveSettingData SettingName, "HannaCode" & i, "QtyProduced", .QtyProduced
            
            
            SaveSettingData SettingName, "HannaCode" & i, "DateProd", .DateProd
            SaveSettingData SettingName, "HannaCode" & i, "LotNumber", .LotNumber
            SaveSettingData SettingName, "HannaCode" & i, "Machine", .Machine
            SaveSettingData SettingName, "HannaCode" & i, "WeekProd", .WeekProd
            
            
            SaveSettingData SettingName, "HannaCode" & i, "Recipe", .Recipe
            SaveSettingData SettingName, "HannaCode" & i, "Std", .STD
            SaveSettingData SettingName, "HannaCode" & i, "Um", .Um
            SaveSettingData SettingName, "HannaCode" & i, "UncertantlyFromCoA", .UncertantlyFromCoA
            
            
            
            If .AcquisitionCount > 0 Then
                Call SetAcquisitionInFile(i, .Acquisitions, .AcquisitionCount)
            End If
                
                
            SaveSettingData SettingName, "HannaCode" & i, "AcquisitionCount", .AcquisitionCount
        
        
        End With

    Next
    
    CloseSettingDataFile
    
End Function

Private Function SetAcquisitionInFile(ByVal t As Integer, ByRef Acquisition() As ProdAcquisition, ByRef AcquisitionCount As Integer)
Dim r As Integer
Dim i As Integer
    CloseSettingDataFile
    r = 1
    
    For i = 1 To AcquisitionCount
    
        If Acquisition(i).bDeleted Then GoTo cont
    
        With Acquisition(r)
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "AcquisitionTime", .AcquisitionTime
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "bDeleted", .bDeleted
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "Code", .Code
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "DateProd", .DateProd
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "ID", .ID
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "Index", .Index
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "Note", .Note
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "Operator", .Operator
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "LotNumber", .LotNumber
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "Machine", .Machine
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "QtyProduced", .QtyProduced
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "WeekProd", .WeekProd
            
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "Mix1Lot", .Mix1Lot
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "Mix2Lot", .Mix2Lot
            SaveSettingData SettingName, "HannaCode" & t & " - Acquisition " & r, "ExpDate", .ExpDate
            
            
        End With
        
        r = r + 1
cont:
    Next

    AcquisitionCount = r - 1
    
    CloseSettingDataFile
    
End Function
