Attribute VB_Name = "Main_02_Preparation"
Option Explicit


Public Function GetDataPreparationInGrid(ByVal Grid As Grid, ByVal strLine As String, ByVal strRecipe As String, ByVal bPreparationDetails As Boolean)

Dim i As Integer
Dim t As Integer
Dim MaxCount As Integer
Dim DiffQty As Double
Dim sString As String
Dim QtyProduced As Double
Dim QtyToProduce As Double
Dim bAllMixes As Boolean
Dim strFileName As String
Dim strStatus As String
Dim MixesCount As Integer
Dim bNoPreparation As Boolean
Dim bHasMixes       As Boolean
On Error GoTo ERR_GET:

    If LCase(strRecipe) = "search" Then
    
        sString = ""
    Else
        sString = " and Recipe like '*" & Replace(Trim(strRecipe), "'", "''") & "*'"
    End If


    If LCase(strLine) = "all lines" Or strLine = "" Then
       ' sString = ""
        Grid.Column(1).Width = 150
    Else
        sString = " and line like '*" & strLine & "*'"
        Grid.Column(1).Width = 0
    End If
    With Grid
        .Rows = 1
        .AutoRedraw = False
       
        With dbTabPreparation
            .filter = ""
            .filter = "bClosed=false" & sString
            If .EOF Then
                GoTo exitMe
            End If
            .MoveFirst
            
            MaxCount = .RecordCount
            
        
            For i = 1 To .RecordCount
            


                 strStatus = (IIf(IsNull(Trim(!QCStatus)), "", Trim(!QCStatus)))
                
                If strStatus = "Passed" Then GoTo cont
                
        
                Grid.AddItem "", False
                
                
                strRecipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
                bAllMixes = IfAllMixes(strRecipe, MixesCount)
                strFileName = IIf(IsNull(Trim(!FileName)), "", Trim(!FileName))
                
                
               
                Grid.Cell(Grid.Rows - 1, 1).Text = IIf(IsNull(Trim(!Line)), "", Trim(!Line)) ' iPreparation(i).Recipe.Line
                Grid.Cell(Grid.Rows - 1, 2).Text = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe)) ' iPreparation(i).Recipe.Code
                Grid.Cell(Grid.Rows - 1, 3).Text = IIf(IsNull(Trim(!HannaCode)), "", Trim(!HannaCode)) ' iPreparation(i).Recipe.Description
         
                Grid.Cell(Grid.Rows - 1, 17).Text = IIf(IsNull(Trim(!Description)), "", Trim(!Description)) ' iPreparation(i).Recipe.Description
         

        
                If bPreparationDetails Then
                
                    Grid.Cell(Grid.Rows - 1, 4).Text = IIf(IsNull(Trim(!numPrepWeek)), "", Trim(!numPrepWeek)) ' iPreparation(i).NumPrepWeek
                    Grid.Cell(Grid.Rows - 1, 5).Text = IIf(IsNull(Trim(!PrepWeek)), "", Trim(!PrepWeek)) ' iPreparation(i).PlannedNumPrepWeek
                    Grid.Cell(Grid.Rows - 1, 6).Text = IIf(IsNull(Trim(!PrepDate)), "", Trim(!PrepDate)) ' iPreparation(i).DateRecipe
                    
                    
                Else
                    Grid.Cell(Grid.Rows - 1, 4).Text = IIf(IsNull(Trim(!RecipeWeek)), "", Trim(!RecipeWeek)) ' iPreparation(i).NumPrepWeek
                    Grid.Cell(Grid.Rows - 1, 5).Text = IIf(IsNull(Trim(!PlannedPreparation)), "", Trim(!PlannedPreparation)) ' iPreparation(i).PlannedNumPrepWeek
                    Grid.Cell(Grid.Rows - 1, 6).Text = IIf(IsNull(Trim(!DataRecipe)), "", Trim(!DataRecipe)) ' iPreparation(i).DateRecipe
                End If
                
                
                
                Grid.Cell(Grid.Rows - 1, 7).Text = PadString(IIf(IsNull(Trim(!QtyToProduce)), "0", Trim(!QtyToProduce))) & " kg" ' iPreparation(i).QtyToProduce & " kg"
                Grid.Cell(Grid.Rows - 1, 8).Text = PadString(IIf(IsNull(Trim(!QtyProduced)), "0", Trim(!QtyProduced))) & " kg" ' iPreparation(i).QtyProduced & " kg"
                Grid.Cell(Grid.Rows - 1, 9).Text = IIf(IsNull(Trim(!FileName)), "", Trim(!FileName))  ' iPreparation(i).FileName
                Grid.Cell(Grid.Rows - 1, 13).Text = !ID ' iPreparation(i).ID
                
                
                
                Grid.Cell(Grid.Rows - 1, 14).Text = "  " & IIf(IsNull(Trim(!QCStatus)), "", Trim(!QCStatus)) ' iPreparation(i).Note
                

                
                
                Grid.Cell(Grid.Rows - 1, 15).Text = "     " & IIf(IsNull(Trim(!Note)), "", Trim(!Note)) ' iPreparation(i).Note
                
                bNoPreparation = IfRecipeNoPreparation(strRecipe)
                bHasMixes = IfRecipeHasMixes(strRecipe)
                
                Grid.Cell(Grid.Rows - 1, 16).Text = bNoPreparation
                
                
                 
                QtyProduced = CDbl(IIf(IsNull(Trim(!QtyProduced)), 0, Trim(!QtyProduced)))
                QtyToProduce = CDbl(IIf(IsNull(Trim(!QtyToProduce)), 0, Trim(!QtyToProduce)))
                

                If !bPesatoTuttiComponenti Then
                
                    Select Case !bCorrection
                        Case True
                            Grid.Cell(Grid.Rows - 1, 11).BackColor = &HC0&
                        Case False
                            Grid.Cell(Grid.Rows - 1, 11).BackColor = vbColorGreen
                    End Select
                Else
                
                    Grid.Cell(Grid.Rows - 1, 11).BackColor = &HC0&
                
                End If
                
                
                Grid.Cell(Grid.Rows - 1, 18).Text = IIf(IsNull(Trim(!Lot)), "", Trim(!Lot))
                Grid.Cell(Grid.Rows - 1, 18).Alignment = cellCenterCenter
                Grid.Cell(Grid.Rows - 1, 18).FontBold = True
                Grid.Cell(Grid.Rows - 1, 18).ForeColor = vbColorDarkUnabled
                  
                 
                
              '  If bHasMixes Then
                If QtyProduced > 0 Then
                    For t = 1 To Grid.Cols - 1
                        Grid.Cell(Grid.Rows - 1, t).FontBold = True
                        Grid.Cell(Grid.Rows - 1, t).ForeColor = vbColorDarkUnabled  '&H644603
                    Next
                End If
                              
                If bNoPreparation Then
                    For t = 1 To Grid.Cols - 1
                        Grid.Cell(Grid.Rows - 1, t).FontBold = True
                        Grid.Cell(Grid.Rows - 1, t).ForeColor = vbColorOrange
                    Next
                End If
                               
                               
                If IfRecipeIsMixString(Trim(!Recipe)) Then
                
                    For t = 1 To Grid.Cols - 1
                        Grid.Cell(Grid.Rows - 1, t).FontBold = True
                        Grid.Cell(Grid.Rows - 1, t).ForeColor = &H644603
                    Next
                
                End If
                Grid.Cell(Grid.Rows - 1, 7).Alignment = cellRightCenter
                Grid.Cell(Grid.Rows - 1, 8).Alignment = cellRightCenter
                
                    
                Select Case Trim(!QCStatus)
                    Case "Passed"
                        Grid.Cell(Grid.Rows - 1, 14).BackColor = vbColorGreen
                         Grid.Cell(Grid.Rows - 1, 14).ForeColor = vbWhite
                    Case "Waiting"
                        Grid.Cell(Grid.Rows - 1, 14).BackColor = vbColorAzzurrino
                    Case "Failed"
                        Grid.Cell(Grid.Rows - 1, 14).BackColor = vbColorRed
                        Grid.Cell(Grid.Rows - 1, 14).ForeColor = vbWhite
                        
                        
                End Select
                If Trim(!QCStatus) <> "" Then Grid.Cell(Grid.Rows - 1, 11).BackColor = Grid.Cell(Grid.Rows - 1, 14).BackColor
               
                Grid.Cell(Grid.Rows - 1, 7).BackColor = vbColorResults
                Grid.Cell(Grid.Rows - 1, 8).BackColor = vbColorResults
cont:
                CloseSettingDataFile
                
                
                .MoveNext
            Next
        End With
        
        If bPreparationDetails Then
        
            .Cell(0, 4).Text = "# Prep. Week"
            .Cell(0, 5).Text = "Prep. Week"
            .Cell(0, 6).Text = "Preparation Date"
        Else
            .Cell(0, 4).Text = "# Prep. Week"
            .Cell(0, 5).Text = "Planned Prep."
            .Cell(0, 6).Text = "Date Recipe"
        
        End If
        
        .Column(0).Width = 0
        .Column(16).Width = 0
        .Column(15).AutoFit
        .Column(17).AutoFit
        
        .Column(1).AutoFit
        .Column(2).AutoFit
        .Column(4).AutoFit
        .Column(5).AutoFit
        .Column(6).AutoFit
        
exitMe:
        .SelectionMode = cellSelectionByRow
        .Refresh
        .AutoRedraw = True
    End With

ERR_END:
    On Error GoTo 0
    Exit Function
ERR_GET:
    Resume Next

End Function



Public Function GetPreparationDataMixes(ByRef Grid As Grid, ByVal FileName As String, ByVal lRow As Long, ByVal MixesCount As Integer, ByVal RecipeName As String) As Boolean

Dim rc As Boolean
Dim i As Integer
Dim IndexRecipe As Integer
Dim RecipesCode As String
Dim QtyProduced As String
Dim QtyToProduce As String
Dim strMixes As String

rc = True
On Error GoTo ERR_GET:

    CloseSettingDataFile
    
    USER_PATH = USER_PREPARATION_PATH
    
    
    If FileName = "" Then
        rc = False
        Exit Function
    End If
    
    If FileExists(USER_PREPARATION_PATH & FileName) Then
    
        With Grid
            IndexRecipe = GetSettingData(FileName, "RecipeIndex", RecipeName, 0)
            If IndexRecipe = 0 Then IndexRecipe = 1
            
            For i = 0 To MixesCount - 2
             
                strMixes = strMixes & "Mix (" & i + 1 & ") " & vbCrLf
                RecipesCode = RecipesCode & "      " & GetSettingData(FileName, "Recipes" & IndexRecipe & " - RmxRecipe" & i, "CHCode", "") & vbCrLf
                QtyToProduce = QtyToProduce & PadString(GetSettingData(FileName, "Recipes" & IndexRecipe & " - RmxRecipe" & i, "TotalWeightKg", "0")) & " kg" & vbCrLf
                QtyProduced = QtyProduced & PadString(GetSettingData(FileName, "Recipes" & IndexRecipe & " - RmxRecipe" & i, "TotalWeightProduced", "0")) & " kg" & vbCrLf
        
            Next
    
            
            .Cell(lRow, 5).Text = strMixes
            .Cell(lRow, 5).Alignment = cellLeftCenter
   
            .Cell(lRow, 6).Text = RecipesCode
            .Cell(lRow, 6).Alignment = cellLeftCenter
      
            .Cell(lRow, 7).Text = QtyToProduce
            .Cell(lRow, 7).Alignment = cellRightCenter
         
            .Cell(lRow, 8).Text = QtyProduced
            .Cell(lRow, 8).Alignment = cellRightCenter
            
            For i = 1 To .Cols - 1
                .Cell(lRow, i).WrapText = True
               .Cell(lRow, i).FontSize = 9
               .Cell(lRow, i).Font = "Segoe UI"
               .Cell(lRow, i).ForeColor = &H644603
               '.Cell(lRow, i).FontBold = True
            Next
            
        End With
    End If
    
    
ERR_END:
    On Error GoTo 0
    CloseSettingDataFile
    GetPreparationDataMixes = rc
    Exit Function
ERR_GET:
    rc = False
    MsgBox err.Description
    Resume Next

End Function


Public Function GetDataPreparationInGrid_ByFile(ByVal Grid As Grid, ByVal strLine As String)

Dim i As Integer
Dim MaxCount As Integer
Dim iPreparation() As PreparationType
Dim DiffQty As Double
Dim sString As String

On Error GoTo ERR_GET:


    If LCase(strLine) = "all lines" Or strLine = "" Then
        sString = ""
    Else
        sString = " and line like '*" & strLine & "*'"
    End If
    With Grid
        .Rows = 1
        .AutoRedraw = False
    
        With dbTabPreparation
            .filter = ""
            .filter = "bClosed=false" & sString
            .MoveFirst
            
            MaxCount = .RecordCount
            
            ReDim iPreparation(.RecordCount)
        
            For i = 1 To .RecordCount
                
                Call GetPreparations(iPreparation(i))
                iPreparation(i).ID = !ID
                DoEvents
                .MoveNext
            Next
            
        End With
            
        For i = 1 To MaxCount
        
            .AddItem "", False
            .Cell(.Rows - 1, 1).Text = iPreparation(i).Recipe.Line
            .Cell(.Rows - 1, 2).Text = iPreparation(i).Recipe.Code
            .Cell(.Rows - 1, 3).Text = iPreparation(i).Recipe.Description
            .Cell(.Rows - 1, 4).Text = iPreparation(i).numPrepWeek
            .Cell(.Rows - 1, 5).Text = iPreparation(i).PlannedPrepWeek
            .Cell(.Rows - 1, 6).Text = iPreparation(i).DateRecipe
            .Cell(.Rows - 1, 7).Text = iPreparation(i).QtyToProduce & " kg"
            .Cell(.Rows - 1, 8).Text = iPreparation(i).QtyProduced & " kg"
            .Cell(.Rows - 1, 9).Text = iPreparation(i).FileName
            .Cell(.Rows - 1, 12).Text = iPreparation(i).ID
            .Cell(.Rows - 1, 13).Text = "     " & iPreparation(i).Note
            
            DiffQty = iPreparation(i).QtyProduced - iPreparation(i).QtyToProduce
            
            Select Case DiffQty
                Case Is > 0
                    .Cell(.Rows - 1, 11).BackColor = vbColorGreen
                Case -iPreparation(i).QtyToProduce * 0.1 To 0
                    .Cell(.Rows - 1, 11).BackColor = vbColorOrange
                Case Is < -iPreparation(i).QtyToProduce * 0.1
                    .Cell(.Rows - 1, 11).BackColor = &HC0&
            
            End Select
            
        
            .Cell(.Rows - 1, 7).Alignment = cellRightCenter
            .Cell(.Rows - 1, 8).Alignment = cellRightCenter
            
                 '.Cell(0, 1).Text = "Recipe"
                 '.Cell(0, 2).Text = "Description"
                 '.Cell(0, 3).Text = "Prep Week"
                 '.Cell(0, 4).Text = "Planned Prep."
                 '.Cell(0, 5).Text = "Data Recipe"
                 '.Cell(0, 6).Text = "Qty To Produce"
                ' .Cell(0, 7).Text = "Qty Produced"
                ' .Cell(0, 8).Text = "FileName"
                
        Next
       .Column(13).AutoFit
        .SelectionMode = cellSelectionByRow
       .Refresh
       .AutoRedraw = True
    End With

ERR_END:
    On Error GoTo 0
    Exit Function
ERR_GET:
    GoTo ERR_END

End Function


Public Function GetPreparations(ByRef uPrep As PreparationType)
Dim FileName As String
Dim i As Integer
Dim IndexRecipe As Integer

CloseSettingDataFile

    With dbTabPreparation
        
        uPrep.bOpen = True
        uPrep.FileName = IIf(IsNull(Trim(!FileName)), "", Trim(!FileName))
        uPrep.PlanningReference = IIf(IsNull(Trim(!PlanningReference)), "", Trim(!PlanningReference))
        uPrep.DateRecipe = IIf(IsNull(Trim(!DataRecipe)), "", Trim(!DataRecipe))
        uPrep.numPrepWeek = IIf(IsNull(Trim(!RecipeWeek)), "", Trim(!RecipeWeek))
        uPrep.PlannedPrepWeek = IIf(IsNull(Trim(!PlannedPreparation)), "", Trim(!PlannedPreparation))
        
        
        
        FileName = uPrep.FileName
        If FileName <> "" Then
            If FileExists(USER_PREPARATION_PATH & FileName) Then
            
    'SaveSettingData SettingName, "Preparation", "Recipe", Trim(RecipeName)
    'SaveSettingData SettingName, "Preparation", "RecipeIndex", Index
    
        If USER_PATH <> USER_PREPARATION_PATH Then USER_PATH = USER_PREPARATION_PATH
     
                IndexRecipe = GetSettingData(FileName, "Preparation", "RecipeIndex", 0)
                
               
                uPrep.Note = GetSettingData(FileName, "iRecipeForProduction", "Note", "")
                
                Call GetSingleRecipeFromFile(uPrep.Recipe, FileName, IndexRecipe)

            End If
            
            
            uPrep.QtyToProduce = uPrep.Recipe.TotalWeightKg
            uPrep.QtyProduced = GetSettingData(FileName, "Preparation", "QtyProduced", 0)
        End If
       
    End With
    
 CloseSettingDataFile
 
End Function



Public Function CheckIfPreparationCorrection(ByVal MyID As Long) As Boolean
Dim rc As Boolean
    With dbTabPreparation
        .filter = ""
        .filter = "ID='" & MyID & "'"
        If .EOF Then
        Else
            rc = !bCorrection
        End If
    End With
    
    CheckIfPreparationCorrection = rc
    
End Function
