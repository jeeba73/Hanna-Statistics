Attribute VB_Name = "Database_RECOVERY_03_Production"
Option Explicit
Private SettingName As String
Private bIfDataPath As Boolean
Private uSTDPreparation As RecipeForSTDPreparation
Private Path As String
Private STDPreparationID As Long

Public Function RecoverySTDPreparationFilesToDatabase()



        Dim rc As Boolean
        Dim Path As String
        rc = False
        Dim FSO As New Scripting.FileSystemObject
        
        Dim Cartella As Folder
        Dim FileGenerico As file
        
opened:
         
        USER_PATH = USER_STD_PREPARATION_PATH
        
        GoTo cont
        
closed:

        USER_PATH = USER_STD_PREPARATION_PATH & "data\"
      
        
cont:
        bIfDataPath = IIf(USER_PATH = USER_STD_PREPARATION_PATH & "data\", True, False)
      
        Path = USER_PATH
        Set Cartella = FSO.GetFolder(Path)
         
        For Each FileGenerico In Cartella.Files
        
            SettingName = FileGenerico.Name
            
            rc = STDPreparationRecoveryGetSetting(uSTDPreparation, SettingName)
            If rc Then
                Dim MyID As Long
                rc = AggiornaTabSTDPreparation
                
                If rc Then
                    Call SaveAcquisitionInTabProdHistory
                    
                End If
                
                
             
            
            End If
        Next
        
        If USER_PATH = USER_STD_PREPARATION_PATH Then GoTo closed


    dbTabSTDPreparation.Close
    dbTabSTDPreparation.Open "SELECT *  FROM TabSTDPreparation order by id ", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
    
    dbTabProdHistory.Close
    dbTabProdHistory.Open "SELECT *  FROM TabProdHistory ", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
   

End Function

Private Function SaveAcquisitionInTabProdHistory()
Dim i As Integer
Dim t As Integer
Dim HannaCodes() As HannaCode
Dim HannaCodesCount As Integer
Dim userAcquisition As ProdAcquisition
Dim userAcquisitionClean As ProdAcquisition

HannaCodesCount = uSTDPreparation.HannaCodesCount
If HannaCodesCount > 0 Then

    HannaCodes = uSTDPreparation.HannaCodes


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
                            !STDPreparationID = STDPreparationID
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

Private Function AggiornaTabSTDPreparation()

Dim rc As Boolean
rc = True
On Error GoTo ERR_AGG

    With dbTabSTDPreparation
        .filter = ""
        .filter = "FileName='" & SettingName & "'"
        If .EOF Then
            .AddNew
        End If
        
        !Line = uSTDPreparation.HannaCodes(1).Line
        !PlanningReference = uSTDPreparation.PlanningReference
        !DataRecipe = uSTDPreparation.DateRecipe
        !Recipe = uSTDPreparation.HannaCodes(1).Recipe

        If !startDate = "" Or IsNull(!startDate) Then
            !startDate = FormatDateTime(Now(), vbShortDate)
        End If
        
        !OperatorRfP = uSTDPreparation.OperatorRfP
        !FileName = SettingName
        !RfpID = GetRfpID(uSTDPreparation.fileNameRecForProd)
        !bClosed = bIfDataPath
        !HannaCode = GetHannCodePerRfp(uSTDPreparation.fileNameRecForProd)
        
        If uSTDPreparation.PreparationDate <> "" Then
            !PrepDate = uSTDPreparation.PreparationDate
            !ExpDate = uSTDPreparation.ExpDate
            !PrepWeek = uSTDPreparation.PrepWeek
            !numPrepWeek = uSTDPreparation.numPrepWeek
        End If
        
        .Update
        STDPreparationID = !ID
    End With
ERR_END:
    On Error GoTo 0
    
    AggiornaTabSTDPreparation = rc
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
    If FileExists(USER_STD_PREPARATION_PATH & RfpFileName) Then
        Path = USER_STD_PREPARATION_PATH
    
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
        With dbTabReceiptForSTDPreparation
            .filter = ""
            .filter = "FileName='" & fileNameRecForProd & "'"
            If .EOF Then
            Else
                GetRfpID = !ID
            End If
        End With
    End If
    
End Function



Public Function STDPreparationRecoveryGetSetting(ByRef iSTDPreparation As RecipeForSTDPreparation, ByVal SettName As String) As Boolean
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
  
   ' With dbTabSTDPreparation
    '    .filter = ""
     '   .filter = "FileName='" & SettingName & "'"
     '   rc = .EOF
     '   If .EOF = False Then GoTo ERR_END:
   ' End With
    
            
            
    
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
CloseSettingDataFile
ERR_END:
    On Error GoTo 0
     
     STDPreparationRecoveryGetSetting = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function

