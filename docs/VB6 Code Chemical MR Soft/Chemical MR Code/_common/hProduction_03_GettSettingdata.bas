Attribute VB_Name = "hProduction_04_GettSettingdata"
Option Explicit


Private SettingName As String

Public Function STDPreparationGetSetting(ByRef iSTDPreparation As RecipeForSTDPreparation, ByVal SettName As String) As Boolean
Dim rc As Boolean
Dim RecipeCount As Integer
Dim HannaCodesCount As Integer
Dim TotalsCount As Integer
Dim PackagingCount As Integer

On Error GoTo ERR_SAVE

    SettingName = SettName
    rc = True
   
    'USER_PATH = USER_STD_PREPARATION_PATH
    
        If FileExists(USER_STD_PREPARATION_PATH & SettingName) Then
            USER_PATH = USER_STD_PREPARATION_PATH
        ElseIf FileExists(USER_STD_PREPARATION_PATH & "data\" & SettingName) Then
            USER_PATH = USER_STD_PREPARATION_PATH & "data\"
        
        Else
            rc = False
            Exit Function
            
        End If
    
    CloseSettingDataFile
  
    
    With iSTDPreparation
       
        .bOpen = GetSettingData(SettingName, "iRecipeForSTDPreparation", "bOpen", .bOpen)
        .DateRecipe = GetSettingData(SettingName, "iRecipeForSTDPreparation", "DateRecipe", .DateRecipe)
        .Note = GetSettingData(SettingName, "iRecipeForSTDPreparation", "Note", .Note)
        .PlannedPrepWeek = GetSettingData(SettingName, "iRecipeForSTDPreparation", "PlannedPrepWeek", .PlannedPrepWeek)
        
        .PreparationDate = GetSettingData(SettingName, "iRecipeForSTDPreparation", "PreparationDate", "")
        .ExpDate = GetSettingData(SettingName, "iRecipeForSTDPreparation", "ExpDate", "")
        .PrepWeek = GetSettingData(SettingName, "iRecipeForSTDPreparation", "PrepWeek", .PrepWeek)
        .numPrepWeek = GetSettingData(SettingName, "iRecipeForSTDPreparation", "NumPrepWeek", .numPrepWeek)
        
        .PlanningReference = GetSettingData(SettingName, "iRecipeForSTDPreparation", "PlanningReference", .PlanningReference)
        
        .RecipeBy = GetSettingData(SettingName, "iRecipeForSTDPreparation", "RecipeBy", .RecipeBy)
        .fileNameRecForProd = GetSettingData(SettingName, "iRecipeForSTDPreparation", "fileNameRecForProd", .fileNameRecForProd)
        .WeekProd = GetSettingData(SettingName, "iRecipeForSTDPreparation", "WeekProd", .WeekProd)
        
        '-----------------------------------------------------------
        ' HANNA CODES in Recipe for STDPreparation
        '-----------------------------------------------------------
        
        .HannaCodesCount = GetSettingData(SettingName, "HannaCodes", "HannaCodesCount", 0)
        ReDim .HannaCodes(.HannaCodesCount)
        If .HannaCodesCount > 0 Then
            HannaCodesCount = .HannaCodesCount
            
            Call GetSTDPreparationHannaCodesFromFile(.HannaCodes, HannaCodesCount, SettingName)
        End If
        
    End With

ERR_END:
    On Error GoTo 0
     CloseSettingDataFile
     STDPreparationGetSetting = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function



Public Function GetSTDPreparationHannaCodesFromFile(ByRef HannaCodes() As HannaCode, ByVal HannaCodesCount As Integer, ByVal SettingName As String)
Dim i As Integer

On Error GoTo ERR_GET:

    For i = 1 To HannaCodesCount
        
        With HannaCodes(i)
            .bHide = GetSettingData(SettingName, "HannaCode" & i, "bHide", .bHide)
            .Code = GetSettingData(SettingName, "HannaCode" & i, "Code", .Code)
            .Density = GetSettingData(SettingName, "HannaCode" & i, "Density", .Density)
            .Exp = GetSettingData(SettingName, "HannaCode" & i, "Exp", .Exp)
            .ExpDate = GetSettingData(SettingName, "HannaCode" & i, "ExpDate", .ExpDate)
            .ID = GetSettingData(SettingName, "HannaCode" & i, "ID", .ID)
            .LastLot = GetSettingData(SettingName, "HannaCode" & i, "LastLot", .LastLot)
            .Line = GetSettingData(SettingName, "HannaCode" & i, "Line", .Line)
            .LoadInPrint = GetSettingData(SettingName, "HannaCode" & i, "LoadInPrint", .LoadInPrint)
            .MaxQty = GetSettingData(SettingName, "HannaCode" & i, "MaxQty", .MaxQty)
            .MinQty = GetSettingData(SettingName, "HannaCode" & i, "MinQty", .MinQty)
            .Mix1 = GetSettingData(SettingName, "HannaCode" & i, "Mix1", .Mix1)
            .Mix2 = GetSettingData(SettingName, "HannaCode" & i, "Mix2", .Mix2)
            .Procedure = GetSettingData(SettingName, "HannaCode" & i, "Procedure", .Procedure)
            .ProcedureRev = GetSettingData(SettingName, "HannaCode" & i, "ProcedureRev", .ProcedureRev)
            .ProductName = GetSettingData(SettingName, "HannaCode" & i, "ProductName", .ProductName)
            .Qty = GetSettingData(SettingName, "HannaCode" & i, "Qty", .Qty)
            .QtyToProduce = GetSettingData(SettingName, "HannaCode" & i, "QtyToProduce", "0")
            
            .Recipe = GetSettingData(SettingName, "HannaCode" & i, "Recipe", .Recipe)
            .STD = GetSettingData(SettingName, "HannaCode" & i, "Std", .STD)
            .Um = GetSettingData(SettingName, "HannaCode" & i, "Um", .Um)
            .UncertantlyFromCoA = GetSettingData(SettingName, "HannaCode" & i, "UncertantlyFromCoA", .UncertantlyFromCoA)
            
            
            .QtyProduced = GetSettingData(SettingName, "HannaCode" & i, "QtyProduced", "0")
            .LotNumber = GetSettingData(SettingName, "HannaCode" & i, "LotNumber", .LotNumber)
            .DateProd = GetSettingData(SettingName, "HannaCode" & i, "DateProd", .DateProd)
            .Machine = GetSettingData(SettingName, "HannaCode" & i, "Machine", .Machine)
            .WeekProd = GetSettingData(SettingName, "HannaCode" & i, "WeekProd", .WeekProd)
        
        
            .AcquisitionCount = GetSettingData(SettingName, "HannaCode" & i, "AcquisitionCount", .AcquisitionCount)
            '-----------------------------------------------------------
            ' Acquisition
            '-----------------------------------------------------------
            If .bHide = False Then
            
                If .AcquisitionCount > 0 Then
                    ReDim .Acquisitions(.AcquisitionCount)
                    Call GetAcquisitionformFile(i, .Acquisitions, .AcquisitionCount, SettingName)
                End If
        
            End If

        End With

    Next
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_GET:
    MsgBox err.Description
    Resume Next
    
End Function
Public Function GetAcquisitionformFile(ByVal t As Integer, ByRef Acquisition() As ProdAcquisition, ByVal AcquisitionCount As Integer, ByVal SettingName As String)
Dim r As Integer
    
    CloseSettingDataFile
 On Error GoTo ERR_GET:
    
    For r = 1 To AcquisitionCount
        
        With Acquisition(r)
            .AcquisitionTime = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "AcquisitionTime", .AcquisitionTime)
            
            .bDeleted = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "bDeleted", .bDeleted)
            .Code = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "Code", .Code)
            .DateProd = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "DateProd", .DateProd)
            .ID = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "ID", .ID)
            .Index = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "Index", .Index)
            .LotNumber = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "LotNumber", .LotNumber)
            .Machine = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "Machine", .Machine)
            .Note = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "Note", .Note)
            .Operator = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "Operator", .Operator)
            .QtyProduced = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "QtyProduced", .QtyProduced)
            .WeekProd = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "WeekProd", .WeekProd)
            
            .Mix1Lot = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "Mix1Lot", .Mix1Lot)
            .Mix2Lot = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "Mix2Lot", .Mix2Lot)
            .ExpDate = GetSettingData(SettingName, "HannaCode" & t & " - Acquisition " & r, "ExpDate", .ExpDate)
            
            
        End With
    Next

   
    
    
ERR_END:
    On Error GoTo 0
    CloseSettingDataFile
    Exit Function
ERR_GET:
    MsgBox err.Description
    Resume Next
End Function
