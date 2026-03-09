Attribute VB_Name = "Database_RECOVERY_03_Production"
Option Explicit
Private SettingName As String
Private bIfDataPath As Boolean
Private uProduction As RecipeForProduction
Private Path As String
Private ProductionID As Long

Public Function RecoveryProductionFilesToDatabase()



        Dim rc As Boolean
        Dim Path As String
        rc = False
        Dim FSO As New Scripting.FileSystemObject
        
        Dim Cartella As Folder
        Dim FileGenerico As file
        
opened:
         
        USER_PATH = USER_PRODUCTION_PATH
        
        GoTo cont
        
closed:

        USER_PATH = USER_PRODUCTION_PATH & "data\"
      
        
cont:
        bIfDataPath = IIf(USER_PATH = USER_PRODUCTION_PATH & "data\", True, False)
      
        Path = USER_PATH
        Set Cartella = FSO.GetFolder(Path)
         
        For Each FileGenerico In Cartella.Files
        
            SettingName = FileGenerico.Name
            
            rc = ProductionRecoveryGetSetting(uProduction, SettingName)
            If rc Then
                Dim MyID As Long
                rc = AggiornaTabProduction
                
                If rc Then
                    Call SaveAcquisitionInTabProdHistory
                    
                End If
                
                
             
            
            End If
        Next
        
        If USER_PATH = USER_PRODUCTION_PATH Then GoTo closed


    dbTabProduction.Close
    dbTabProduction.Open "SELECT *  FROM TabProduction order by id -1", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
    
    dbTabProdHistory.Close
    dbTabProdHistory.Open "SELECT *  FROM TabProdHistory ", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
   

End Function

Private Function SaveAcquisitionInTabProdHistory()
Dim i As Integer
Dim t As Integer
Dim HannaCodes() As HannaCode
Dim HannaCodesCount As Integer
Dim userAcquisition As ProdAcquisition
Dim userAcquisitionClean As ProdAcquisition

HannaCodesCount = uProduction.HannaCodesCount
If HannaCodesCount > 0 Then

    HannaCodes = uProduction.HannaCodes


     For i = 1 To HannaCodesCount
        
        With HannaCodes(i)
            If .bHide Then
            Else
                If .AcquisitionCount > 0 Then
                
                    For t = 1 To .AcquisitionCount
                        userAcquisition = userAcquisitionClean
                        userAcquisition = .Acquisitions(t)
                        
                        With dbTabProdHistory
                            .AddNew
                            !AcquisitionTime = userAcquisition.AcquisitionTime
                            !Code = userAcquisition.Code
                            !Index = userAcquisition.Index
                            
                            !DateProd = userAcquisition.DateProd
                            !LotNumber = userAcquisition.LotNumber
                            !Machine = userAcquisition.Machine
                        
                            !QtyProduced = userAcquisition.QtyProduced
                            !Note = userAcquisition.Note
                            !Operator = userAcquisition.Operator
                            !WeekProd = userAcquisition.WeekProd
                            !FileName = SettingName
                            !ProductionID = ProductionID
                            !Mix1Lot = userAcquisition.Mix1Lot
                            !Mix2Lot = userAcquisition.Mix2Lot
                            !ExpDate = userAcquisition.ExpDate
                            .Update
                        End With
                    Next
                End If
            End If
        End With
    Next
End If

End Function

Private Function AggiornaTabProduction()

Dim rc As Boolean
rc = True
On Error GoTo ERR_AGG

    With dbTabProduction
        .filter = ""
        .filter = "FileName='" & SettingName & "'"
        If .EOF Then
            .AddNew
        End If
        
        !Line = uProduction.HannaCodes(1).Line
        !PlanningReference = uProduction.PlanningReference
        !DataRecipe = uProduction.DateRecipe
        !Recipe = uProduction.HannaCodes(1).Recipe

        If !startDate = "" Or IsNull(!startDate) Then
            !startDate = FormatDateTime(Now(), vbShortDate)
        End If
        
        !OperatorRfP = uProduction.OperatorRfP
        !FileName = SettingName
        !RfpID = GetRfpID(uProduction.fileNameRecForProd)
        !bClosed = bIfDataPath
        !HannaCode = GetHannCodePerRfp(uProduction.fileNameRecForProd)
        
        If uProduction.PreparationDate <> "" Then
            !PrepDate = uProduction.PreparationDate
            !ExpDate = uProduction.ExpDate
            !PrepWeek = uProduction.PrepWeek
            !numPrepWeek = uProduction.numPrepWeek
        End If
        
        .Update
        ProductionID = !ID
    End With
ERR_END:
    On Error GoTo 0
    
    AggiornaTabProduction = rc
    Exit Function
ERR_AGG:
    MsgBox err.Description
    rc = False
    GoTo ERR_END
End Function
Private Function GetHannCodePerRfp(ByVal RfpFileName As String) As String
Dim i As Integer
Dim MaxHannaCode As Integer
Dim strHannaCode As String
Dim HannaCode As String
Dim Path As String
Dim bHide As Boolean
On Error GoTo ERR_GET:
    If FileExists(USER_PRODUCTION_PATH & RfpFileName) Then
        Path = USER_PRODUCTION_PATH
    
    Else
      
        GoTo ERR_END
    End If
    
    MaxHannaCode = GetSettingData(RfpFileName, "HannaCodes", "HannaCodesCount", 0, Path)
    strHannaCode = ""
    
    For i = 1 To MaxHannaCode
        
        bHide = GetSettingData(RfpFileName, "HannaCode" & i, "bHide", True, Path)
        HannaCode = GetSettingData(RfpFileName, "HannaCode" & i, "Code", "", Path)
        
        If Not (bHide) Then
        
            If strHannaCode = "" Then
            Else
                strHannaCode = strHannaCode & " ; "
            End If
            
            strHannaCode = strHannaCode & HannaCode
    
        End If
    Next
    
ERR_END:
    On Error GoTo 0
    GetHannCodePerRfp = strHannaCode
    
    Debug.Print Len(strHannaCode)
    
    strHannaCode = Left(strHannaCode, 250)
    CloseSettingDataFile
    
    GetHannCodePerRfp = strHannaCode
    Exit Function
ERR_GET:
    MsgBox err.Description
    Resume Next
    
End Function
Private Function GetRfpID(ByVal fileNameRecForProd As String) As Long

    GetRfpID = 0

    If fileNameRecForProd <> "" Then
        With dbTabReceiptForProduction
            .filter = ""
            .filter = "FileName='" & fileNameRecForProd & "'"
            If .EOF Then
            Else
                GetRfpID = !ID
            End If
        End With
    End If
    
End Function



Public Function ProductionRecoveryGetSetting(ByRef iProduction As RecipeForProduction, ByVal SettName As String) As Boolean
Dim rc As Boolean
Dim RecipeCount As Integer
Dim HannaCodesCount As Integer
Dim TotalsCount As Integer
Dim PackagingCount As Integer

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
  
   ' With dbTabProduction
    '    .filter = ""
     '   .filter = "FileName='" & SettingName & "'"
     '   rc = .EOF
     '   If .EOF = False Then GoTo ERR_END:
   ' End With
    
            
            
    
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
        '-----------------------------------------------------------
        ' HANNA CODES in Recipe for production
        '-----------------------------------------------------------
        
        ProductionID = .ProductionID
        
        .HannaCodesCount = GetSettingData(SettingName, "HannaCodes", "HannaCodesCount", 0)
        ReDim .HannaCodes(.HannaCodesCount)
        If .HannaCodesCount > 0 Then
            HannaCodesCount = .HannaCodesCount
            
            Call GetProductionHannaCodesFromFile(.HannaCodes, HannaCodesCount, SettingName, ProductionID)
        End If
        
    End With
CloseSettingDataFile
ERR_END:
    On Error GoTo 0
     
     ProductionRecoveryGetSetting = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function

