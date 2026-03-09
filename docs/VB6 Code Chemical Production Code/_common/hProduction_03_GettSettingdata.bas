Attribute VB_Name = "hProduction_04_GettSettingdata"
Option Explicit


Private SettingName As String

Public Function ProductionGetSetting(ByRef iProduction As RecipeForProduction, ByVal SettName As String) As Boolean
Dim rc As Boolean
Dim RecipeCount As Integer
Dim HannaCodesCount As Integer
Dim TotalsCount As Integer
Dim PackagingCount As Integer
Dim ProductionID As Long

On Error GoTo ERR_SAVE

    SettingName = SettName
    rc = True
   
    'USER_PATH = USER_PRODUCTION_PATH
    
        If FileExists(USER_PRODUCTION_PATH & SettingName) Then
            USER_PATH = USER_PRODUCTION_PATH
        ElseIf FileExists(USER_PRODUCTION_PATH & "data\" & SettingName) Then
            USER_PATH = USER_PRODUCTION_PATH & "data\"
        
        Else
            rc = False
            Exit Function
            
        End If
    
    CloseSettingDataFile
  
    
    With iProduction
       
        .bOpen = GetSettingData(SettingName, "iRecipeForProduction", "bOpen", .bOpen)
        .DateRecipe = GetSettingData(SettingName, "iRecipeForProduction", "DateRecipe", .DateRecipe)
        .Note = GetSettingData(SettingName, "iRecipeForProduction", "Note", .Note)
        .PlannedPrepWeek = GetSettingData(SettingName, "iRecipeForProduction", "PlannedPrepWeek", .PlannedPrepWeek)
        
        .PreparationDate = GetSettingData(SettingName, "iRecipeForProduction", "PreparationDate", "")
        .PreparationLot = GetSettingData(SettingName, "iRecipeForProduction", "PreparationLot", "")
        .ExpDate = GetSettingData(SettingName, "iRecipeForProduction", "ExpDate", "")
        .PrepWeek = GetSettingData(SettingName, "iRecipeForProduction", "PrepWeek", .PrepWeek)
        .numPrepWeek = GetSettingData(SettingName, "iRecipeForProduction", "NumPrepWeek", .numPrepWeek)
        
        .PlanningReference = GetSettingData(SettingName, "iRecipeForProduction", "PlanningReference", .PlanningReference)
        
        .RecipeBy = GetSettingData(SettingName, "iRecipeForProduction", "RecipeBy", .RecipeBy)
        .fileNameRecForProd = GetSettingData(SettingName, "iRecipeForProduction", "fileNameRecForProd", .fileNameRecForProd)
        .WeekProd = GetSettingData(SettingName, "iRecipeForProduction", "WeekProd", .WeekProd)
        .ProductionID = GetSettingData(SettingName, "iRecipeForProduction", "ProductionID", ProductionID)
        ProductionID = .ProductionID '= GetSettingData(SettingName, "iRecipeForProduction", "ProductionID", ProductionID)
        '-----------------------------------------------------------
        ' HANNA CODES in Recipe for production
        '-----------------------------------------------------------
        
        .HannaCodesCount = GetSettingData(SettingName, "HannaCodes", "HannaCodesCount", 0)
        ReDim .HannaCodes(.HannaCodesCount)
        If .HannaCodesCount > 0 Then
            HannaCodesCount = .HannaCodesCount
            
            Call GetProductionHannaCodesFromFile(.HannaCodes, HannaCodesCount, SettingName, ProductionID)
        End If
        
    End With

ERR_END:
    On Error GoTo 0
     CloseSettingDataFile
     ProductionGetSetting = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function



Public Function GetProductionHannaCodesFromFile(ByRef HannaCodes() As HannaCode, ByVal HannaCodesCount As Integer, ByVal SettingName As String, ByVal ProductionID As Long)
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
           ' '.LoadInPrint = GetSettingData(SettingName, "HannaCode" & i, "LoadInPrint", '.LoadInPrint)
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
                    If ProductionID = 0 Then
                    
                        Call GetAcquisitionformFile(i, .Acquisitions, .AcquisitionCount, SettingName)
                    Else
                        Call GetAcquisitionfromDB(i, .Acquisitions, .AcquisitionCount, ProductionID, .Code)
                    End If
                    
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


Public Function GetAcquisitionfromDB(ByVal t As Integer, ByRef Acquisition() As ProdAcquisition, ByRef AcquisitionCount As Integer, ByVal ProductionID As Long, ByVal Code As String)
Dim r As Integer
    
    CloseSettingDataFile
 On Error GoTo ERR_GET:
    
    With dbTabProdHistory
            .filter = ""
            .filter = "ProductionID='" & ProductionID & "' and Code='" & Code & "'"
            If .EOF Then
            
            Else
                .MoveFirst
                AcquisitionCount = .RecordCount
                ReDim Preserve Acquisition(.RecordCount)
                For r = 1 To AcquisitionCount
                    
                    With Acquisition(r)
                        .AcquisitionTime = IIf(IsNull(Trim(dbTabProdHistory!AcquisitionTime)), "", Trim(dbTabProdHistory!AcquisitionTime))
                        .bDeleted = False
                        .Code = IIf(IsNull(Trim(dbTabProdHistory!Code)), "", Trim(dbTabProdHistory!Code))
                        .DateProd = IIf(IsNull(Trim(dbTabProdHistory!DateProd)), "", Trim(dbTabProdHistory!DateProd))
                        .ID = dbTabProdHistory!ID
                        .Index = IIf(IsNull(Trim(dbTabProdHistory!Index)), "", Trim(dbTabProdHistory!Index))
                        .LotNumber = IIf(IsNull(Trim(dbTabProdHistory!LotNumber)), "", Trim(dbTabProdHistory!LotNumber))
                        .Machine = IIf(IsNull(Trim(dbTabProdHistory!Machine)), "", Trim(dbTabProdHistory!Machine))
                        .Note = IIf(IsNull(Trim(dbTabProdHistory!Note)), "", Trim(dbTabProdHistory!Note))
                        .Operator = IIf(IsNull(Trim(dbTabProdHistory!Operator)), "", Trim(dbTabProdHistory!Operator))
                        .QtyProduced = IIf(IsNull(Trim(dbTabProdHistory!QtyProduced)), "", Trim(dbTabProdHistory!QtyProduced))
                        .WeekProd = IIf(IsNull(Trim(dbTabProdHistory!WeekProd)), "", Trim(dbTabProdHistory!WeekProd))
                        .Mix1Lot = IIf(IsNull(Trim(dbTabProdHistory!Mix1Lot)), "", Trim(dbTabProdHistory!Mix1Lot))
                        .Mix2Lot = IIf(IsNull(Trim(dbTabProdHistory!Mix2Lot)), "", Trim(dbTabProdHistory!Mix2Lot))
                        .ExpDate = IIf(IsNull(Trim(dbTabProdHistory!ExpDate)), "", Trim(dbTabProdHistory!ExpDate))
                        
                    End With
                    
                    .MoveNext
                Next
                
                
            End If

    End With
   
    
    
ERR_END:
    On Error GoTo 0
    CloseSettingDataFile
    Exit Function
ERR_GET:
    MsgBox err.Description
    Resume Next
End Function
