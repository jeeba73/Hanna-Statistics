Attribute VB_Name = "Main_04_Production"
Option Explicit

Public Type RfpDetails
    ID                  As Long
    PlanningReference   As String
    PlannedPreparation  As String
    DataRecipe          As String
    RecipeWeek          As String
    Operator            As String
    bClosed             As Boolean
    Recipe              As String
    Note                As String
    FileName            As String
    Line                As String
    
    PrepDate            As String
    PrepWeek            As String
    numPrepWeek         As String
    ExpDate             As String
    Lot                 As String
  

End Type

Private userRfp As RfpDetails



Public Function SetDatabaseTabProduction(ByRef RecipeQC As QCType) As Boolean

Dim rc As Boolean
Dim Path As String
Dim ID_PREPARATION As Long
Dim RfpFileName As String
On Error GoTo ERR_SET

    RfpFileName = RecipeQC.SettingName
    ID_PREPARATION = RecipeQC.ID
    
    rc = True
    CloseSettingDataFile
    
    If FileExists(USER_DATA_PATH & RfpFileName) Then
        Path = USER_DATA_PATH
    ElseIf FileExists(USER_TEMP_PATH & RfpFileName) Then
        Path = USER_TEMP_PATH
    Else
        rc = False
        GoTo ERR_END
    End If
    

    FileCopy Path & RfpFileName, USER_PRODUCTION_PATH & RfpFileName
    
    
    With dbTabReceiptForProduction
        .filter = ""
        .filter = "FileName ='" & RfpFileName & "'"
        If .EOF Then
            'rc = False
            GoTo prep
        End If
        
        userRfp.bClosed = !bClosed
        userRfp.DataRecipe = IIf(IsNull(Trim(!DataRecipe)), "", Trim(!DataRecipe))
        userRfp.FileName = IIf(IsNull(Trim(!FileName)), "", Trim(!FileName))
        userRfp.ID = !ID
        userRfp.Recipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
        userRfp.Line = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
        userRfp.Note = IIf(IsNull(Trim(!Note)), "", Trim(!Note))
        userRfp.Operator = IIf(IsNull(Trim(!Operator)), "", Trim(!Operator))
        userRfp.PlanningReference = IIf(IsNull(Trim(!PlanningReference)), "", Trim(!PlanningReference))
        userRfp.RecipeWeek = IIf(IsNull(Trim(!RecipeWeek)), "", Trim(!RecipeWeek))
        userRfp.PlannedPreparation = IIf(IsNull(Trim(!PlannedPreparation)), "", Trim(!PlannedPreparation))
        
        
        If !bClosed Then
        Else
            Call CloseRecipeForProduction(RfpFileName, Path)
            userRfp.bClosed = True
        End If
    
    End With
    
prep:
    
    With dbTabPreparation
        .filter = ""
        .filter = "ID='" & ID_PREPARATION & "'"
        
        If .EOF Then
        
        Else
            userRfp.PrepDate = IIf(IsNull(Trim(!PrepDate)), "", Trim(!PrepDate))
            userRfp.PrepWeek = IIf(IsNull(Trim(!PrepWeek)), "", Trim(!PrepWeek))
            userRfp.numPrepWeek = IIf(IsNull(Trim(!numPrepWeek)), "", Trim(!numPrepWeek))
            userRfp.ExpDate = IIf(IsNull(Trim(!ExpDate)), "", Trim(!ExpDate))
            userRfp.Lot = IIf(IsNull(Trim(!Lot)), "", Trim(!Lot))
            
            If !bClosed Then
            
            Else
            
            End If
            
            ' se non ho i valori di preparation allora cerco in RfpFileName se li ha scritti qualche MIX
            If userRfp.PrepDate = "" And userRfp.PrepWeek = "" Then
                CloseSettingDataFile
                userRfp.PrepDate = GetSettingData(RfpFileName, "iRecipeForProduction", "PreparationDate", "", USER_PRODUCTION_PATH)
                userRfp.Lot = GetSettingData(RfpFileName, "iRecipeForProduction", "PreparationLot", "", USER_PRODUCTION_PATH)
                userRfp.PrepWeek = GetSettingData(RfpFileName, "iRecipeForProduction", "PrepWeek", "", USER_PRODUCTION_PATH)
                userRfp.numPrepWeek = GetSettingData(RfpFileName, "iRecipeForProduction", "numPrepWeek", "", USER_PRODUCTION_PATH)
                userRfp.ExpDate = GetSettingData(RfpFileName, "iRecipeForProduction", "ExpDate", "", USER_PRODUCTION_PATH)
                CloseSettingDataFile
                GoTo cont:
            End If
            
        End If
    
    End With
    
    CloseSettingDataFile
    
    With userRfp
        SaveSettingData RfpFileName, "iRecipeForProduction", "PreparationDate", .PrepDate, USER_PRODUCTION_PATH
        SaveSettingData RfpFileName, "iRecipeForProduction", "PreparationLot", .Lot, USER_PRODUCTION_PATH
        SaveSettingData RfpFileName, "iRecipeForProduction", "PrepWeek", .PrepWeek, USER_PRODUCTION_PATH
        SaveSettingData RfpFileName, "iRecipeForProduction", "numPrepWeek", .numPrepWeek, USER_PRODUCTION_PATH
        SaveSettingData RfpFileName, "iRecipeForProduction", "ExpDate", .ExpDate, USER_PRODUCTION_PATH
    End With
    
    CloseSettingDataFile
    
cont:
    
    With dbTabProduction
        .filter = ""
        .filter = "RfpID='" & userRfp.ID & "'"
        If .EOF Then
            .AddNew
        End If
        !Line = userRfp.Line
        !PlanningReference = userRfp.PlanningReference
        !DataRecipe = IIf(IsNull(userRfp.DataRecipe), FormatTimeLAT(Now()), userRfp.DataRecipe)
        !Recipe = userRfp.Recipe

        If !startDate = "" Or IsNull(!startDate) Then
            !startDate = FormatDateTime(Now(), vbShortDate)
        End If
        
        !OperatorRfP = userRfp.Operator
        !FileName = RfpFileName
        !Lot = userRfp.Lot
        !RfpID = userRfp.ID
        !bClosed = False
        !HannaCode = GetHannCodePerRfp(RfpFileName, USER_PRODUCTION_PATH)
        
        !RecipeWeek = userRfp.RecipeWeek
        !PlannedPreparation = userRfp.PlannedPreparation

        
        
        
        
        If userRfp.PrepDate <> "" Then
            !PrepDate = userRfp.PrepDate
            !ExpDate = userRfp.ExpDate
            !PrepWeek = userRfp.PrepWeek
            !numPrepWeek = userRfp.numPrepWeek
        Else
        End If
        
        
        
        .Update
        .Close
        .Open "SELECT *  FROM TabProduction order by id ", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
    End With
    
    
    
    
    
ERR_END:
    On Error GoTo 0
    SetDatabaseTabProduction = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox err.Description
    Resume Next

End Function

Public Function GetHannCodePerRfp(ByVal RfpFileName As String, ByVal USER_Rfp_PATH As String) As String
Dim i As Integer
Dim MaxHannaCode As Integer
Dim strHannaCode As String
Dim HannaCode As String
Dim Path As String
Dim bHide As Boolean
On Error GoTo ERR_GET:
    If FileExists(USER_Rfp_PATH & RfpFileName) Then
        Path = USER_Rfp_PATH
    
    Else
      
        GoTo ERR_END
    End If
    
    CloseSettingDataFile
    
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

Public Function GetSFGLotPerRfp(ByVal RfpFileName As String, Optional ByVal USER_Rfp_PATH As String) As String
Dim i As Integer
Dim MaxSFGLot As String

Dim strSFGLot As String
Dim SFGLot As String
Dim Path As String
Dim bHide As Boolean
On Error GoTo ERR_GET:

    If FileExists(USER_Rfp_PATH & RfpFileName) Then
        Path = USER_Rfp_PATH
    ElseIf FileExists(USER_TEMP_PATH & RfpFileName) Then
        Path = USER_TEMP_PATH
    ElseIf FileExists(USER_PRODUCTION_PATH & RfpFileName) Then
        Path = USER_PRODUCTION_PATH
    ElseIf FileExists(USER_PREPARATION_PATH & RfpFileName) Then
        Path = USER_PREPARATION_PATH
    Else
        GoTo ERR_END
    End If
    
    CloseSettingDataFile
    
    MaxSFGLot = GetSettingData(RfpFileName, "HannaCodes", "HannaCodesCount", 0, Path)
    strSFGLot = ""
    Debug.Print GetSettingData(RfpFileName, "iRecipeForProduction", "fileName", 0, Path)
   ' Debug.Print Path
    
    For i = 1 To MaxSFGLot
        
        bHide = GetSettingData(RfpFileName, "HannaCode" & i, "bHide", True, Path)
        SFGLot = GetSettingData(RfpFileName, "HannaCode" & i, "LotNumber", "", Path)
        
        If Not (bHide) Then
        
            If strSFGLot = "" Then
            Else
                strSFGLot = strSFGLot & " ; "
            End If
            
            strSFGLot = strSFGLot & SFGLot
    
        End If
    Next
    
ERR_END:
    On Error GoTo 0
    GetSFGLotPerRfp = strSFGLot
    
    Debug.Print Len(strSFGLot)
    CloseSettingDataFile
    strSFGLot = Left(strSFGLot, 250)
    
    
    GetSFGLotPerRfp = strSFGLot
    Exit Function
ERR_GET:
    MsgBox err.Description
    Resume Next
    
End Function

Public Function GetProductionInGrid4(ByRef Grid4 As Grid, ByVal bValue As Boolean, ByVal HannaCode As String, ByRef Frame5 As Frame, ByVal strLine As String) As Boolean
Dim i As Integer
Dim t As Integer
Dim sString As String
  
    sString = "bClosed='" & bValue & "'"
    
    If LCase(HannaCode) <> "search" Then
    
        If Trim(HannaCode) <> "" Then sString = sString & " and HannaCode like '*" & Replace(Trim(HannaCode), "'", "''") & "*'"
    
    End If
    
    If LCase(strLine) = "all lines" Or strLine = "" Then
       ' sString = ""
        Grid4.Column(2).Width = 150
    Else
    
        sString = sString & " and line like '*" & Replace(Trim(strLine), "'", "''") & "*'"
        
        
        
        Grid4.Column(2).Width = 0
    End If

With Grid4
    .Rows = 1
    .AutoRedraw = False
    
    With dbTabProduction
        .Close
        .Open "SELECT *  FROM TabProduction order by id -1", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
        .filter = ""
        .filter = sString
        If .EOF Then
        Else
            .MoveFirst
            For i = 1 To .RecordCount
            
                Grid4.AddItem "", False
                Grid4.Cell(i, 1).Text = IIf(IsNull(Trim(!HannaCode)), "", Trim(!HannaCode))
                Grid4.Cell(i, 2).Text = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
                Grid4.Cell(i, 3).Text = IIf(IsNull(Trim(!DataRecipe)), "", Trim(!DataRecipe))
                
                '------------------
                ' Recipes
                ' Mixes
                '------------------
                Grid4.Cell(i, 4).Text = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
               ' Grid4.Cell(i, 4).Text = IIf(IsNull(Trim(!RecipeWeek)), "", Trim(!RecipeWeek))
               ' Grid4.Cell(i, 5).Text = IIf(IsNull(Trim(!PlannedPreparation)), "", Trim(!PlannedPreparation))
              ' Grid4.Cell(i, 6).Text = IIf(IsNull(Trim(!PlanningReference)), "", Trim(!PlanningReference))
                
                Grid4.Cell(i, 7).Text = IIf(IsNull(Trim(!PrepDate)), "", Trim(!PrepDate))
                Grid4.Cell(i, 8).Text = IIf(IsNull(Trim(!PrepWeek)), "", Trim(!PrepWeek))
                Grid4.Cell(i, 9).Text = IIf(IsNull(Trim(!numPrepWeek)), "", Trim(!numPrepWeek))
                
                
                
                Grid4.Cell(i, 10).Text = IIf(IsNull(Trim(!FileName)), "", Trim(!FileName))
                Grid4.Cell(i, 11).Text = !ID

                If IsProductionStarted(!ID) Then
                 
                    For t = 1 To Grid4.Cols - 1
                        Grid4.Cell(Grid4.Rows - 1, t).FontBold = True
                        Grid4.Cell(Grid4.Rows - 1, t).ForeColor = &H4D3B37   '&H644603
                        
                    Next
                    Grid4.Cell(i, 12).BackColor = vbColorOrange
                Else
                    
                End If
                Grid4.Cell(i, 1).ForeColor = &H4D3B37
                Grid4.Cell(i, 1).FontBold = True
                
                Grid4.Cell(i, 13).Text = IIf(IsNull(Trim(!Lot)), "", Trim(!Lot))
                    
                .MoveNext

            Next
        End If
    End With
   
    .Refresh
    .AutoRedraw = True
End With

Frame5.Visible = IIf(Grid4.Rows > 1, False, True)

End Function



Public Function IsProductionStarted(ByVal ID As Long) As Boolean
Dim rc As Boolean
rc = False
With dbTabProdHistory
    .filter = ""
    .filter = "ProductionID='" & ID & "'"
    rc = Not (.EOF)
End With
IsProductionStarted = rc
End Function


Public Function SetProductionCheckOut(ByVal ProductionFileName As String, ByVal ProductionID As Long) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim fileNamePreparation As String
rc = True

On Error GoTo ERR_SET:



    '---------------------------------------------------------------
    ' check out : production
    ' 1 close TabProduction
    '---------------------------------------------------------------
    With dbTabProduction
        .filter = ""
        .filter = "ID='" & ProductionID & "'"
        If .EOF Then
        Else
            !bClosed = True
            !CloseDate = FormatDataLAT(Now())
        End If
        .Update
    End With
    '---------------------------------------------------------------
    ' 2 move ProductionFileName to DATA
    '---------------------------------------------------------------
    If FileExists(USER_PRODUCTION_PATH & ProductionFileName) Then
        FileCopy USER_PRODUCTION_PATH & ProductionFileName, USER_PRODUCTION_PATH & "Data\" & ProductionFileName
        Kill USER_PRODUCTION_PATH & ProductionFileName
    End If
    
    '---------------------------------------------------------------
    ' 2.1 check and move recipe for production file from temp to data
    '---------------------------------------------------------------
    
      If FileExists(USER_TEMP_PATH & ProductionFileName) Then
        FileCopy USER_TEMP_PATH & ProductionFileName, USER_DATA_PATH & ProductionFileName
        Kill USER_TEMP_PATH & ProductionFileName
    End If
    '---------------------------------------------------------------
    ' 3 verify and close preparation / recipefor production....
    '---------------------------------------------------------------
    With dbTabPreparation
        .filter = ""
        .filter = "RfPFileName ='" & ProductionFileName & "'"
        If .EOF Then
        Else
            .MoveFirst
            fileNamePreparation = ""
            For i = 1 To .RecordCount
                !bClosed = True
                fileNamePreparation = IIf(IsNull(Trim(!FileName)), "", Trim(!FileName))
                '-----------------------------------------------------------------------
                ' 3.1  check and move recipe Preparations file to USER_PREPARATION\Data
                '-----------------------------------------------------------------------
                            
                If FileExists(USER_PREPARATION_PATH & fileNamePreparation) Then
                    FileCopy USER_PREPARATION_PATH & fileNamePreparation, USER_PREPARATION_PATH & "Data\" & fileNamePreparation
                    Kill USER_PREPARATION_PATH & fileNamePreparation
                End If
                
                
                fileNamePreparation = ""
                .MoveNext
                
            Next
          
        End If
    End With
    
ERR_END:
    On Error GoTo 0
    SetProductionCheckOut = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox err.Description
    Resume Next
    
    
End Function

Public Function SetProductionDelete(ByVal ProductionFileName As String, ByVal ProductionID As Long) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim prepCount As Integer
Dim fileNamePreparation As String
rc = True

On Error GoTo ERR_SET:



    '---------------------------------------------------------------
    ' check out : production
    ' 1 close TabProduction
    '---------------------------------------------------------------
    With dbTabProduction
        .filter = ""
        .filter = "ID='" & ProductionID & "'"
        If .EOF Then
        Else
            .Delete
            .Update
        End If
    End With
    '---------------------------------------------------------------
    ' 2 move ProductionFileName to DATA
    '---------------------------------------------------------------
    If FileExists(USER_PRODUCTION_PATH & ProductionFileName) Then
        'FileCopy USER_PRODUCTION_PATH & ProductionFileName, USER_PRODUCTION_PATH & "Data\" & ProductionFileName
        Kill USER_PRODUCTION_PATH & ProductionFileName
    End If
    
    '---------------------------------------------------------------
    ' 2.1 check and move recipe for production file from temp to data
    '---------------------------------------------------------------
    
      If FileExists(USER_TEMP_PATH & ProductionFileName) Then
        'FileCopy USER_TEMP_PATH & ProductionFileName, USER_DATA_PATH & ProductionFileName
        Kill USER_TEMP_PATH & ProductionFileName
    End If
    '---------------------------------------------------------------
    ' 3 verify and close preparation / recipefor production....
    '---------------------------------------------------------------
    With dbTabPreparation
        .filter = ""
        .filter = "RfPFileName ='" & ProductionFileName & "'"
        If .EOF Then
        Else
            .MoveFirst
            fileNamePreparation = ""
            
            
            For i = 1 To .RecordCount
                
                fileNamePreparation = IIf(IsNull(Trim(!FileName)), "", Trim(!FileName))
                '-----------------------------------------------------------------------
                ' 3.1  check and move recipe Preparations file to USER_PREPARATION\Data
                '-----------------------------------------------------------------------
                
                If fileNamePreparation = "" Then
                Else
                    
                                
                    If FileExists(USER_PREPARATION_PATH & fileNamePreparation) Then
                        Kill USER_PREPARATION_PATH & fileNamePreparation
                    End If
                
                End If
                
                fileNamePreparation = ""
                .MoveNext
            Next
            .MoveFirst
            prepCount = .RecordCount
            For i = 1 To .RecordCount
                .Delete
                .MoveNext
            Next
            'If prepCount > 0 Then .Update
           .Close
           .Open "SELECT *  FROM TabPreparation order by id ", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
          
        End If
    End With
    
ERR_END:
    On Error GoTo 0
    SetProductionDelete = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox err.Description
    Resume Next
    
    
End Function

