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
  

End Type

Private userRfp As RfpDetails



Public Function SetDatabaseTabSTDPreparation(ByRef RecipeQC As QCType) As Boolean

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
    

    FileCopy Path & RfpFileName, USER_STD_PREPARATION_PATH & RfpFileName
    
    
    With dbTabReceiptForSTDPreparation
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
            Call CloseRecipeForSTDPreparation(RfpFileName, Path)
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
            If !bClosed Then
            
            Else
            
            End If
            
            ' se non ho i valori di preparation allora cerco in RfpFileName se li ha scritti qualche MIX
            If userRfp.PrepDate = "" And userRfp.PrepWeek = "" Then
                CloseSettingDataFile
                userRfp.PrepDate = GetSettingData(RfpFileName, "iRecipeForSTDPreparation", "PreparationDate", "", USER_STD_PREPARATION_PATH)
                userRfp.PrepWeek = GetSettingData(RfpFileName, "iRecipeForSTDPreparation", "PrepWeek", "", USER_STD_PREPARATION_PATH)
                userRfp.numPrepWeek = GetSettingData(RfpFileName, "iRecipeForSTDPreparation", "numPrepWeek", "", USER_STD_PREPARATION_PATH)
                userRfp.ExpDate = GetSettingData(RfpFileName, "iRecipeForSTDPreparation", "ExpDate", "", USER_STD_PREPARATION_PATH)
                CloseSettingDataFile
                GoTo cont:
            End If
            
        End If
    
    End With
    
    CloseSettingDataFile
    
    With userRfp
        SaveSettingData RfpFileName, "iRecipeForSTDPreparation", "PreparationDate", .PrepDate, USER_STD_PREPARATION_PATH
        SaveSettingData RfpFileName, "iRecipeForSTDPreparation", "PrepWeek", .PrepWeek, USER_STD_PREPARATION_PATH
        SaveSettingData RfpFileName, "iRecipeForSTDPreparation", "numPrepWeek", .numPrepWeek, USER_STD_PREPARATION_PATH
        SaveSettingData RfpFileName, "iRecipeForSTDPreparation", "ExpDate", .ExpDate, USER_STD_PREPARATION_PATH
    End With
    
    CloseSettingDataFile
    
cont:
    
    With dbTabSTDPreparation
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
        !RfpID = userRfp.ID
        !bClosed = False
        !HannaCode = GetHannCodePerRfp(RfpFileName)
        
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
        .Open "SELECT *  FROM TabSTDPreparation order by id ", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
    End With
    
    
    
    
    
ERR_END:
    On Error GoTo 0
    SetDatabaseTabSTDPreparation = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox err.Description
    Resume Next

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
    
    
    GetHannCodePerRfp = strHannaCode
    Exit Function
ERR_GET:
    MsgBox err.Description
    Resume Next
    
End Function


Public Function GetSTDPreparationInGrid4(ByRef Grid4 As Grid, ByVal bValue As Boolean, ByVal HannaCode As String, ByRef Frame5 As Frame, ByVal strLine As String) As Boolean
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
    
    With dbTabSTDPreparation
        .Close
        .Open "SELECT *  FROM TabSTDPreparation order by id -1", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
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

                If IsSTDPreparationStarted(!ID) Then
                 
                    For t = 1 To Grid4.Cols - 1
                        Grid4.Cell(Grid4.Rows - 1, t).FontBold = True
                        Grid4.Cell(Grid4.Rows - 1, t).ForeColor = &H4D3B37   '&H644603
                        
                    Next
                    Grid4.Cell(i, 12).BackColor = vbColorOrange
                Else
                    
                End If
                Grid4.Cell(i, 1).ForeColor = &H4D3B37
                Grid4.Cell(i, 1).FontBold = True
                    
                .MoveNext

            Next
        End If
    End With
   
    .Refresh
    .AutoRedraw = True
End With

Frame5.Visible = IIf(Grid4.Rows > 1, False, True)

End Function



Public Function IsSTDPreparationStarted(ByVal ID As Long) As Boolean
Dim rc As Boolean
rc = False
With dbTabProdHistory
    .filter = ""
    .filter = "STDPreparationID='" & ID & "'"
    rc = Not (.EOF)
End With
IsSTDPreparationStarted = rc
End Function


Public Function SetSTDPreparationCheckOut(ByVal STDPreparationFileName As String, ByVal STDPreparationID As Long) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim fileNamePreparation As String
rc = True

On Error GoTo ERR_SET:



    '---------------------------------------------------------------
    ' check out : STDPreparation
    ' 1 close TabSTDPreparation
    '---------------------------------------------------------------
    With dbTabSTDPreparation
        .filter = ""
        .filter = "ID='" & STDPreparationID & "'"
        If .EOF Then
        Else
            !bClosed = True
            !CloseDate = FormatDataLAT(Now())
        End If
        .Update
    End With
    '---------------------------------------------------------------
    ' 2 move STDPreparationFileName to DATA
    '---------------------------------------------------------------
    If FileExists(USER_STD_PREPARATION_PATH & STDPreparationFileName) Then
        FileCopy USER_STD_PREPARATION_PATH & STDPreparationFileName, USER_STD_PREPARATION_PATH & "Data\" & STDPreparationFileName
        Kill USER_STD_PREPARATION_PATH & STDPreparationFileName
    End If
    
    '---------------------------------------------------------------
    ' 2.1 check and move recipe for STDPreparation file from temp to data
    '---------------------------------------------------------------
    
      If FileExists(USER_TEMP_PATH & STDPreparationFileName) Then
        FileCopy USER_TEMP_PATH & STDPreparationFileName, USER_DATA_PATH & STDPreparationFileName
        Kill USER_TEMP_PATH & STDPreparationFileName
    End If
    '---------------------------------------------------------------
    ' 3 verify and close preparation / recipefor STDPreparation....
    '---------------------------------------------------------------
    With dbTabPreparation
        .filter = ""
        .filter = "RfPFileName ='" & STDPreparationFileName & "'"
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
                            
                If FileExists(USER_SCHEDULED_STD_PATH & fileNamePreparation) Then
                    FileCopy USER_SCHEDULED_STD_PATH & fileNamePreparation, USER_SCHEDULED_STD_PATH & "Data\" & fileNamePreparation
                    Kill USER_SCHEDULED_STD_PATH & fileNamePreparation
                End If
                
                
                fileNamePreparation = ""
                .MoveNext
                
            Next
          
        End If
    End With
    
ERR_END:
    On Error GoTo 0
    SetSTDPreparationCheckOut = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox err.Description
    Resume Next
    
    
End Function

Public Function SetSTDPreparationDelete(ByVal STDPreparationFileName As String, ByVal STDPreparationID As Long) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim prepCount As Integer
Dim fileNamePreparation As String
rc = True

On Error GoTo ERR_SET:



    '---------------------------------------------------------------
    ' check out : STDPreparation
    ' 1 close TabSTDPreparation
    '---------------------------------------------------------------
    With dbTabSTDPreparation
        .filter = ""
        .filter = "ID='" & STDPreparationID & "'"
        If .EOF Then
        Else
            .Delete
            .Update
        End If
    End With
    '---------------------------------------------------------------
    ' 2 move STDPreparationFileName to DATA
    '---------------------------------------------------------------
    If FileExists(USER_STD_PREPARATION_PATH & STDPreparationFileName) Then
        'FileCopy USER_STD_PREPARATION_PATH & STDPreparationFileName, USER_STD_PREPARATION_PATH & "Data\" & STDPreparationFileName
        Kill USER_STD_PREPARATION_PATH & STDPreparationFileName
    End If
    
    '---------------------------------------------------------------
    ' 2.1 check and move recipe for STDPreparation file from temp to data
    '---------------------------------------------------------------
    
      If FileExists(USER_TEMP_PATH & STDPreparationFileName) Then
        'FileCopy USER_TEMP_PATH & STDPreparationFileName, USER_DATA_PATH & STDPreparationFileName
        Kill USER_TEMP_PATH & STDPreparationFileName
    End If
    '---------------------------------------------------------------
    ' 3 verify and close preparation / recipefor STDPreparation....
    '---------------------------------------------------------------
    With dbTabPreparation
        .filter = ""
        .filter = "RfPFileName ='" & STDPreparationFileName & "'"
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
                    
                                
                    If FileExists(USER_SCHEDULED_STD_PATH & fileNamePreparation) Then
                        Kill USER_SCHEDULED_STD_PATH & fileNamePreparation
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
           .Open "SELECT *  FROM TabPreparation order by id ", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
          
        End If
    End With
    
ERR_END:
    On Error GoTo 0
    SetSTDPreparationDelete = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox err.Description
    Resume Next
    
    
End Function

