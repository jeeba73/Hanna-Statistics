Attribute VB_Name = "Main_01_ReceiptForProduction"
Option Explicit


Public Function GetReceiptFromDatabase(ByRef Grid As Grid, ByVal bClosed As Boolean, Optional ByVal bTuttiRecord As Boolean, Optional strLine As String) As Boolean
Dim i As Integer
Dim Count As Integer
Dim sString As String

CloseSettingDataFile


    sString = "bClosed='" & IIf(bClosed, "True", "False") & "'"
    
    If LCase(strLine) = "all lines" Or strLine = "" Then
       ' sString = ""
        'Grid.Column(2).Width = 150
    Else
    
    
        sString = sString & " and line like '*" & Replace(Trim(strLine), "'", "''") & "*'"
        
        
        
        
    End If

    Dim dDate As Date
    
    
    dDate = DateAdd("m", -6, Date)
    
    If bClosed Then sString = sString & " and DataRecipe>=#" & dDate & "# "
    


With Grid

    .Rows = 1
    .AutoRedraw = False
    With dbTabReceiptForProduction
        .filter = ""
        .filter = sString
        
        If .EOF Then
            Count = 0
        Else
            Count = .RecordCount
            .MoveLast
        End If
        
        
        If bTuttiRecord = False And bClosed Then
            Count = IIf(Count > 30, 30, Count)
        End If

               
        '.Cell(0, 1).Text = "Line"
        '.Cell(0, 2).Text = "Date Recipe"
        '.Cell(0, 3).Text = "Prep. Week"
        '.Cell(0, 4).Text = "Pl. Prep Week"
        '.Cell(0, 5).Text = "Pl. Reference"
        '.Cell(0, 6).Text = "Operator"
    
        '.Cell(0, 7).Text = "Recipes"

        '.Cell(0, 8).Text = "Description"
        '.Cell(0, 9).Text = "MR Printed"
        '.Cell(0, 10).Text = "MR Number"
        '.Cell(0, 11).Text = "Note"
        
        
       ' .Cell(0, 12).Text = "FileName"
       ' .Cell(0, 13).Text = "ID"
        
        
        For i = 1 To Count
        
            Grid.AddItem "", False
            Grid.Cell(Grid.Rows - 1, 1).Text = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
            Grid.Cell(Grid.Rows - 1, 2).Text = IIf(IsNull(Trim(!DataRecipe)), "", Trim(!DataRecipe))
            Grid.Cell(Grid.Rows - 1, 3).Text = IIf(IsNull(Trim(!RecipeWeek)), "", Trim(!RecipeWeek))
            Grid.Cell(Grid.Rows - 1, 4).Text = IIf(IsNull(Trim(!PlannedPreparation)), "", Trim(!PlannedPreparation))
            Grid.Cell(Grid.Rows - 1, 5).Text = IIf(IsNull(Trim(!PlanningReference)), "", Trim(!PlanningReference))
            Grid.Cell(Grid.Rows - 1, 6).Text = IIf(IsNull(Trim(!Operator)), "", Trim(!Operator))
            Grid.Cell(Grid.Rows - 1, 7).Text = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
            Grid.Cell(Grid.Rows - 1, 8).Text = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
            Grid.Cell(Grid.Rows - 1, 9).Text = IIf(IsNull(Trim(!bMaterialRequisitionPrinted)), False, !bMaterialRequisitionPrinted)
            Grid.Cell(Grid.Rows - 1, 10).Text = IIf(IsNull(Trim(!MaterialRequisitionNumber)), "", Trim(!MaterialRequisitionNumber))
            Grid.Cell(Grid.Rows - 1, 11).Text = IIf(IsNull(Trim(!Note)), "", Trim(!Note))
            Grid.Cell(Grid.Rows - 1, 12).Text = IIf(IsNull(Trim(!FileName)), "", Trim(!FileName))
            Grid.Cell(Grid.Rows - 1, 13).Text = !ID
            
            Grid.Cell(Grid.Rows - 1, 9).Alignment = cellCenterCenter
            Grid.Cell(Grid.Rows - 1, 7).FontBold = True
            Grid.Cell(Grid.Rows - 1, 7).ForeColor = &H473733
            Grid.Cell(Grid.Rows - 1, 9).FontBold = True
            Grid.Cell(Grid.Rows - 1, 9).ForeColor = &H473733
            
            .MovePrevious
        Next
    
    End With
    .Column(7).AutoFit
    .Column(8).AutoFit
    .Column(10).AutoFit
    .Column(11).AutoFit
    .Refresh
    .AutoRedraw = True
    .ReadOnly = True


End With

End Function


Public Function DeleteRecipeForProduction(ByVal ID As Long, ByVal FileName As String) As Boolean
Dim rc As Boolean
On Error GoTo ERR_DELETE
rc = False


    
    With dbTabReceiptForProduction
        .filter = ""
        If ID = 0 Then
            .filter = "FileName='" & FileName & "'"
        Else
            .filter = "ID='" & ID & "'"
        End If
        If .EOF Then
            rc = False
            GoTo ERR_END
        Else
            rc = True
            .Delete
            .Update
            .Close
            .Open "SELECT *  FROM TabReceiptForProduction order by id ", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
        End If

    End With
    
    

    If FileExists(USER_PATH & FileName) Then
        If USER_PATH = USER_DATA_PATH Then
        Else
            FileCopy USER_PATH & FileName, USER_DATA_PATH & FileName
            
            Kill USER_PATH & FileName
        End If
    End If
    


ERR_END:
    On Error GoTo 0
    DeleteRecipeForProduction = rc
    Exit Function
ERR_DELETE:
    rc = False
    MsgBox err.Description
    Resume Next
End Function


Public Function SetRecipeGrid5(ByVal Grid5 As Grid, ByVal ReceiptFileName As String) As Boolean
Dim rc As Boolean

Dim i As Integer
Dim iRecipe() As RecipeType
Dim RecipeCount As Integer
Dim MRNumber As String

On Error GoTo ErrSetRecipe

    rc = True
    
    
    If ReceiptFileName = "" Then Exit Function
    
        '-----------------------------------------------------------
        ' Recipes in Recipe for production
        '-----------------------------------------------------------
        Debug.Print USER_PATH
        RecipeCount = GetSettingData(ReceiptFileName, "Recipes", "RecipeCount", 0)
        
        
        ReDim iRecipe(RecipeCount)
        If RecipeCount > 0 Then
            Call GetRecipesFromFile(iRecipe, RecipeCount, ReceiptFileName, True) ' true perchè voglio visualizzare solo le ricette visibili
              With Grid5
                .Rows = 1
                .AutoRedraw = False
                .Column(5).CellType = cellCheckBox
                    For i = 1 To RecipeCount
        
        '.Cell(0, 1).Text = "Recipe"
        '.Cell(0, 2).Text = "Description"
        '.Cell(0, 3).Text = "Line"
        '.Cell(0, 4).Text = "Mix"
        '.Cell(0, 5).Text = "MR Print"
        '.Cell(0, 6).Text = "MR Number"
        '.Cell(0, 7).Text = "Q.ty to produce"
                       
                        
                        CloseSettingDataFile
                        
                        MRNumber = GetSettingData(ReceiptFileName, "Material Requisition" & iRecipe(i).ID, "txDocument(0)", "")
                      
                        .AddItem "", True
                        .Cell(.Rows - 1, 1).Text = iRecipe(i).Code
                        .Cell(.Rows - 1, 2).Text = iRecipe(i).Description
                        .Cell(.Rows - 1, 3).Text = iRecipe(i).Line
                        .Cell(.Rows - 1, 4).Text = iRecipe(i).Mix
                        .Cell(.Rows - 1, 5).Text = IIf(MRNumber <> "", True, False)
                        .Cell(.Rows - 1, 6).Text = MRNumber
                        .Cell(.Rows - 1, 7).Text = GetSettingData(ReceiptFileName, "Material Requisition" & iRecipe(i).ID, "CheckOut", False)
                        .Cell(.Rows - 1, 8).Text = GetSettingData(ReceiptFileName, "Material Requisition" & iRecipe(i).ID, "DateCheckOut", "")
                        .Cell(.Rows - 1, 9).Text = PadString(iRecipe(i).TotalWeightKg) & " kg"
                        .Cell(.Rows - 1, 10).Text = iRecipe(i).ID


                        If iRecipe(i).bHide Then .RowHeight(.Rows - 1) = 0
cont:
                    Next
                    .SelectionMode = cellSelectionByRow
                    .Column(4).Width = 0
                     
                    .Refresh
                    .AutoRedraw = True
             End With
        End If

        
    
    Call CheckOutMaterialRequisitionGrid5(Grid5, ReceiptFileName, RecipeCount, iRecipe)
    
    
    
    

ERR_END:
    On Error GoTo 0
    SetRecipeGrid5 = rc
    Exit Function
ErrSetRecipe:
    rc = False
    MsgBox err.Description
    Resume Next
End Function


Public Sub AddMaterialRequisitionFromFile(ByVal Grd As Grid, ByVal Index As Integer, ByRef xDocument() As String, ByVal SettingName As String)
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim X As Integer
Dim RowsCount As Integer


    If Index = 0 Then Index = 1
    
    rc = True
    

    If SettingName = "" Then
        Exit Sub
     
    End If

    CloseSettingDataFile

    
    For i = 0 To UBound(xDocument)
        xDocument(i) = GetSettingData(SettingName, "Material Requisition" & Index, "txDocument(" & i & ")", "")
    Next
    
    RowsCount = GetSettingData(SettingName, "Material Requisition" & Index, "Rows", 0)
    With Grd
        .AutoRedraw = False
        For i = 1 To RowsCount
            .AddItem "", False
            For t = 1 To .Cols - 1
                .Cell(i, t).Text = GetSettingData(SettingName, "Material Requisition" & Index, "Grd(" & i & "," & t & ")", "")
                .Column(t).Alignment = cellLeftCenter
                .Column(t).Width = 150
                .Cell(0, t).FontBold = True
            
            Next
            .Cell(i, 5).Text = "        " & .Cell(i, 5).Text
          '  .Cell(i, 6).Text = "    " & .Cell(i, 6).Text
            
        Next
        .Cell(0, 5).Alignment = cellCenterCenter
        .Column(2).Width = 250
        .Column(3).Width = 100
        .Column(5).Width = 150
        .Column(6).AutoFit
        .Column(2).AutoFit
        .Column(4).Alignment = cellRightCenter
        .Column(5).Alignment = cellLeftCenter
        .Column(6).Alignment = cellLeftCenter
        .Refresh
        .AutoRedraw = True
    End With

    


End Sub


Public Function CheckOutMaterialRequisitionInFile(ByVal Index As Integer, ByVal SettingName As String, ByRef RecipeName As String) As Boolean

Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim X As Integer
Dim RowsCount As Integer


    If Index = 0 Then Index = 1
    
    rc = True
    

    If SettingName = "" Then
        Exit Function
     
    End If

    CloseSettingDataFile

    RecipeName = GetSettingData(SettingName, "Recipes" & Index, "Code", "")

    CloseSettingDataFile
    SaveSettingData SettingName, "Material Requisition" & Index, "CheckOut", True
    SaveSettingData SettingName, "Material Requisition" & Index, "DateCheckOut", FormatDataLAT(Now())


    SaveSettingData SettingName, "Preparation", "Recipe", Trim(RecipeName)
    SaveSettingData SettingName, "Preparation", "RecipeIndex", Index

    CloseSettingDataFile
    
    CheckOutMaterialRequisitionInFile = rc
    
End Function

Public Function CheckOutMaterialRequisitionGrid5(ByVal Grid5 As Grid, ByVal ReceiptFileName As String, ByVal RecipeCount As Integer, ByRef uRecipe() As RecipeType) As Boolean
Dim rc As Boolean

Dim i As Integer
Dim t As Integer

'Dim RecipeCount As Integer
Dim bCheckOut As Boolean
Dim Count As Integer

On Error GoTo ErrSetRecipe

    rc = True
    
    CloseSettingDataFile
    
    If ReceiptFileName = "" Then Exit Function
    
        '-----------------------------------------------------------
        ' Recipes in Recipe for production
        '-----------------------------------------------------------
        
       ' RecipeCount = GetSettingData(ReceiptFileName, "Recipes", "RecipeCount", 0)
        
        
        ReDim iRecipe(RecipeCount)
        If RecipeCount > 0 Then
            
              With Grid5
                .AutoRedraw = False
                    Count = 1
                    For i = 1 To RecipeCount
                    
                     '   If GetSettingData(ReceiptFileName, "Recipes" & i, "bHide", True) Then GoTo cont:
                        
                       
                        bCheckOut = GetSettingData(ReceiptFileName, "Material Requisition" & uRecipe(i).ID, "CheckOut", False)
                        
                        .Cell(Count, 7).Text = bCheckOut
                        .Cell(Count, 8).Text = GetSettingData(ReceiptFileName, "Material Requisition" & uRecipe(i).ID, "DateCheckOut", "")
                        
                        If bCheckOut Then
                            For t = 0 To .Cols - 1
                                .Cell(i, t).BackColor = IIf(bCheckOut, vbColorAzzurrino, &HF0F0F0)
                            Next
                        End If
                        Count = Count + 1
cont:
                    Next
                    .Refresh
                    .AutoRedraw = True
             End With
        End If


ERR_END:
    On Error GoTo 0
    
    CloseSettingDataFile
    
    CheckOutMaterialRequisitionGrid5 = rc
    Exit Function
ErrSetRecipe:
    rc = False
    MsgBox err.Description
    Resume Next
End Function

Public Function CheckOutAllRequisitionInFile(ByVal ReceiptFileName As String) As Boolean
Dim rc As Boolean

Dim i As Integer
Dim t As Integer

Dim RecipeCount As Integer
Dim bCheckOut As String

On Error GoTo ErrSetRecipe

    rc = True
    
    
    CloseSettingDataFile
    
    If ReceiptFileName = "" Then Exit Function
    
        '-----------------------------------------------------------
        ' Recipes in Recipe for production
        '-----------------------------------------------------------
        
        RecipeCount = GetSettingData(ReceiptFileName, "Recipes", "RecipeCount", 0)
        
        
        ReDim iRecipe(RecipeCount)
        
        If RecipeCount > 0 Then
        
            For i = 1 To RecipeCount
            
                If CBool(GetSettingData(ReceiptFileName, "Recipes" & i, "bHide", True)) Then GoTo cont
                
                bCheckOut = GetSettingData(ReceiptFileName, "Material Requisition" & i, "CheckOut", "")
                If UCase(bCheckOut) = UCase("false") Or UCase(bCheckOut) = "" Then
                    rc = False
                End If
cont:
            Next
            
        End If


ERR_END:
    On Error GoTo 0
    
    CloseSettingDataFile
    
    CheckOutAllRequisitionInFile = rc
    Exit Function
ErrSetRecipe:
    rc = False
    MsgBox err.Description
    Resume Next
End Function

Public Function SetTabPreparationPerRecipe(ByRef FileName As String, ByVal RecipeName As String) As Boolean
Dim rc                  As Boolean
Dim mrc                 As Boolean
Dim NewFileName         As String
Dim numPrepWeek         As String
Dim DataRecipe          As String
Dim PlanningReference   As String
Dim Operator            As String
Dim PlannedPreparation  As String
Dim Line                As String
Dim Description         As String
Dim strPrepWeek         As String
Dim bIsMix              As Boolean


On Error GoTo ErrSetPrepRec



    rc = True

    CloseSettingDataFile

    With dbTabReceiptForProduction
        .filter = ""
        .filter = "FileName ='" & FileName & "' and bClosed=false"
        If .EOF Then
            rc = False
            GoTo ERR_END
        Else
            numPrepWeek = IIf(IsNull(!RecipeWeek), "", Trim(!RecipeWeek))
            DataRecipe = IIf(IsNull(!DataRecipe), "", Trim(!DataRecipe))
            PlannedPreparation = IIf(IsNull(!PlannedPreparation), "", Trim(!PlannedPreparation))
            Operator = IIf(IsNull(!Operator), "", Trim(!Operator))
            PlanningReference = IIf(IsNull(!PlanningReference), "", Trim(!PlanningReference))
            Line = IIf(IsNull(!Line), "", Trim(!Line))
            Description = IIf(IsNull(!Description), "", Trim(!Description))
            
            
            SaveSettingData FileName, "dbTabReceiptForProduction", "ID", !ID
            
        End If

    End With
    
TryAgaing:
    
    NewFileName = FormatNomeFile(RecipeName & "." & numPrepWeek & "." & PlannedPreparation & "." & year(Now())) & USER_ESTENSIONE_PREPARATION

    CloseSettingDataFile

    Dim IndexRecipe As Integer
    Dim strQuantity As String
    
    IndexRecipe = GetSettingData(FileName, "RecipeIndex", RecipeName, 0)
    
    If IndexRecipe > 0 Then
        
        strQuantity = GetSettingData(FileName, "Recipes" & IndexRecipe, "TotalWeightKg", "0")
        
        
    End If
    
    
    With dbTabPreparation
        .filter = ""
        .filter = "FileName ='" & NewFileName & "' and bClosed=false"
        If .EOF Then
            .AddNew
        Else
            ' esiste già!
            'rc = False
            'GoTo ERR_END
            If F_MsgBox.DoShow("A recipe with the same #PrepWeek = " & numPrepWeek & " is already in Preparation." & vbCrLf & "Overwrite or change #PrepWeek?", RecipeName, True, "Overwrite", "Change") Then
                
            Else
                
                strPrepWeek = numPrepWeek
                If F_InputBox.DoShow("Enter new #PrepWeek", RecipeName, , , , strPrepWeek, , True) Then
                    
                    numPrepWeek = strPrepWeek
                    GoTo TryAgaing
                    
                Else
                    rc = False
                    GoTo ERR_END
                    
                End If
            
            End If
        End If
    
        !Line = Line
        !Description = GetRecipeDescription(RecipeName)
        !Recipe = Trim(RecipeName)
        !PlanningReference = Trim(PlanningReference)
        !DataRecipe = FormatDataLAT(Trim(DataRecipe))
        !QtyToProduce = strQuantity
        !RecipeWeek = Trim(numPrepWeek)
        !PlannedPreparation = Trim(PlannedPreparation)
        !Operator = Trim(Operator)
        !bClosed = False
        !Note = ""
        !FileName = NewFileName
        !RfpFileName = FileName
        !bIsMix = IfRecipeIsMixString(RecipeName)
        !HannaCode = GetHannCodePerRfp(FileName, USER_TEMP_PATH)
        
         bIsMix = !bIsMix
        .Update
    End With
    
   
    CloseSettingDataFile

    SaveSettingData FileName, "Preparation", "Recipe", Trim(RecipeName)
    SaveSettingData FileName, "iRecipeForProduction", "fileName", Trim(FileName)
    CloseSettingDataFile
    
    If FileExists(USER_PATH & FileName) Then
        FileCopy USER_PATH & FileName, USER_PREPARATION_PATH & NewFileName
    End If
    
    
     CloseSettingDataFile
    
    '-----------------------------------------------------------------
    '
    '
    '       controllo se in GridTotals ci sono Mix e
    '       li metto in Preparation
    '
    '
    '-----------------------------------------------------------------
    
    Dim GridCount As Integer
    Dim RecipeCode As String
    Dim QtyKg As Double
    Dim isMix As Boolean
    Dim Mix As Integer
    Dim i As Integer
    
    
    GridCount = GetSettingData(FileName, "Totals Grid", "TotalCount", 1)
    
    '--------------------------------------------------------------
    ' se ho 1 sola ricetta è inutile che faccio un check delle mix
    '--------------------------------------------------------------
    If GridCount <= 1 Then GoTo ERR_END
    If bIsMix Then GoTo ERR_END
    '--------------------------------------------------------------
    
    For i = 1 To GridCount
    
        QtyKg = CDbl(GetSettingData(FileName, "Totals Grid" & i, "TotalWeighKg", 0))
        isMix = GetSettingData(FileName, "Totals Grid" & i, "bMix", False)
        
        If QtyKg > 0 And isMix Then
        
            
            RecipeCode = GetSettingData(FileName, "Totals Grid" & i, "Recipe", "")
            IndexRecipe = GetSettingData(FileName, "RecipeIndex", RecipeCode, 0)
            
TryAgaingMix:
            NewFileName = FormatNomeFile(RecipeCode & "." & numPrepWeek & "." & PlannedPreparation & "." & year(Now())) & USER_ESTENSIONE_PREPARATION
            
            
            With dbTabPreparation
                .filter = ""
                .filter = "FileName ='" & NewFileName & "' and bClosed=false"
                If .EOF Then
                    If F_MsgBox.DoShow("Mix in Recipe for production." & vbCrLf & "Set Preparation for this Mix with Quantity : " & QtyKg & " kg", RecipeCode) = False Then GoTo cont
          
                    .AddNew
                Else
                    ' esiste già!
                    If F_MsgBox.DoShow("A recipe with the same #PrepWeek = " & numPrepWeek & " is already in Preparation." & vbCrLf & "Overwrite or change #PrepWeek?", RecipeName, True, "Overwrite", "Change") Then
                        
                    Else
                        
                        strPrepWeek = numPrepWeek
                        If F_InputBox.DoShow("Enter new #PrepWeek", RecipeName, , , , strPrepWeek, , True) Then
                            
                            numPrepWeek = strPrepWeek
                            GoTo TryAgaingMix
                            
                        Else
                            rc = False
                            GoTo ERR_END
                            
                            
                        End If
                    
                    End If
                End If
            
                !Line = Line
                !Description = GetRecipeDescription(RecipeCode)
                !Recipe = Trim(RecipeCode)
                !PlanningReference = Trim(PlanningReference)
                !DataRecipe = Trim(DataRecipe)
                !QtyToProduce = QtyKg
                !RecipeWeek = Trim(numPrepWeek)
                !PlannedPreparation = Trim(PlannedPreparation)
                !Operator = Trim(Operator)
                !bClosed = False
                !Note = ""
                !FileName = NewFileName
                !RfpFileName = FileName
                !bIsMix = IfRecipeIsMixString(RecipeCode)
                .Update
            
            
               Call CheckOutMaterialRequisitionInFile(IndexRecipe, FileName, RecipeCode)
              

            
    
            End With
            
            
            If FileExists(USER_PATH & FileName) Then
                FileCopy USER_PATH & FileName, USER_PREPARATION_PATH & NewFileName
            End If
            
            
             CloseSettingDataFile
     
        End If
cont:

    Next

   

    
    
ERR_END:
    On Error GoTo 0
    CloseSettingDataFile
    SetTabPreparationPerRecipe = rc
    Exit Function
ErrSetPrepRec:
    rc = False
    MsgBox err.Description
    Resume Next
End Function


Public Function CloseRecipeForProduction(ByVal FileName As String, ByVal Path As String) As Boolean
Dim rc As Boolean

On Error GoTo ErrSetPrepRec

    USER_PATH = IIf(Path <> "", Path, "")

    rc = True

    CloseSettingDataFile
    

    With dbTabReceiptForProduction
        .filter = ""
        .filter = "FileName ='" & FileName & "' and bClosed=false"
        If .EOF Then
            rc = False
            GoTo ERR_END
        Else
            !bClosed = True
            .Update
        End If
    
    
    End With
    
    
    
    
    SaveSettingData FileName, "iRecipeForProduction", "bOpen", False, Path
    
    CloseSettingDataFile
 
    If FileExists(USER_PATH & FileName) Then
    
        If USER_PATH = USER_DATA_PATH Then
        Else
            FileCopy USER_PATH & FileName, USER_DATA_PATH & FileName
            Kill USER_PATH & FileName
        End If
        
    End If

    
ERR_END:
    On Error GoTo 0
    CloseRecipeForProduction = rc
    Exit Function
ErrSetPrepRec:
    rc = False
    MsgBox err.Description
    Resume Next
End Function



Public Function IfRecipeForProductionHasAllMixes(ByRef uRecipe() As RecipeType)
Dim rc As Boolean
Dim mrc As Boolean
Dim RecipesCount As Integer
Dim RecipeCode As String
Dim MixesCount As Integer
Dim i As Integer

    RecipesCount = UBound(uRecipe)
    rc = False
    For i = 1 To RecipesCount
        RecipeCode = uRecipe(i).Code
        mrc = IfAllMixes(RecipeCode)
        If mrc = False Then Exit For
    Next
    rc = mrc
    IfRecipeForProductionHasAllMixes = rc
End Function
