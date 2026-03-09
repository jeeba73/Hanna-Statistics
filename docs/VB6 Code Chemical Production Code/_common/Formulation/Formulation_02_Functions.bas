Attribute VB_Name = "Formulation_02_Functions"
Option Explicit




Public Function AddCodeInRecipeForProductionGrid(ByVal Grid1 As Grid, ByVal HannaCode As String, ByVal bSalta As Boolean) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim Recipe As String
Dim Mix1 As String
Dim Mix2 As String
Dim t As Integer
Dim uHannaCode() As HannaCode
            
On Error GoTo ERR_ADD:

    rc = True
        HannaCode = Trim(HannaCode)
     
        With Grid1
            .AutoRedraw = False
            If .Rows > 1 Then
                For i = 1 To .Rows - 1
                    
                        If Trim(LCase(.Cell(i, 1).Text)) = Trim(LCase(HannaCode)) Then
                            If bSalta Then GoTo ERR_END
                            If F_MsgBox.DoShow("Warning : Hanna Code already in Table!", HannaCode, , "Add Again", "Don't") Then
                                Exit For
                            Else
                                rc = False
                                GoTo ERR_END
                            End If
                  
                    Else
                    
                    End If
                Next
            End If
            
            With dbTabCode
                .filter = ""
                .filter = "Code='" & HannaCode & "'"
                If .EOF Then
            
                Else
                    
                    Recipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
                    Mix1 = IIf(IsNull(Trim(!Mix1)), "", Trim(!Mix1))
                    Mix2 = IIf(IsNull(Trim(!Mix2)), "", Trim(!Mix2))
                    
                End If
                .filter = ""
                .filter = ""
                .MoveFirst
                t = 0
                For i = 1 To .RecordCount
                
                    
                    'If (InStr(!Recipe, Recipe) And Recipe <> "") Or ((UCase(!Mix1) = UCase(Mix1) And Mix1 <> "") And (UCase(!Mix2) = UCase(Mix2) And Mix2 <> "")) Then
                     
                    If (InStr(!Recipe, Recipe) And Recipe <> "") Then
                    
                        
                        ReDim Preserve uHannaCode(t)
                    
                        uHannaCode(t).Code = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
                        uHannaCode(t).Line = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
                        uHannaCode(t).STD = IIf(IsNull(Trim(!STD)), "", Trim(!STD))
                        uHannaCode(t).ProductName = IIf(IsNull(Trim(!ProductName)), "", Trim(!ProductName))
                        uHannaCode(t).Recipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
                        uHannaCode(t).Mix1 = IIf(IsNull(Trim(!Mix1)), "", Trim(!Mix1))
                        uHannaCode(t).Mix2 = IIf(IsNull(Trim(!Mix2)), "", Trim(!Mix2))
                        uHannaCode(t).Um = IIf(IsNull(Trim(!Um)), "", Trim(!Um))
                        uHannaCode(t).Qty = CheckDot(IIf(IsNull(Trim(!Qty)), "0", Trim(!Qty)))
                        uHannaCode(t).MinQty = CheckDot(IIf(IsNull(Trim(!MinQty)), "0", Trim(!MinQty)))
                        uHannaCode(t).MaxQty = CheckDot(IIf(IsNull(Trim(!MaxQty)), "0", Trim(!MaxQty)))

                        ' in attesa di ok da vianello
                        'uHannaCode(t).LotNumber = IIf(IsNull(Trim(!PrintedInformation)), "", Trim(!PrintedInformation))
                        
                        t = t + 1
                    End If
cont:
                    .MoveNext
                Next
                
            End With
            
            For i = LBound(uHannaCode) To UBound(uHannaCode)
                .AddItem "", False
                .Cell(.Rows - 1, 1).Text = uHannaCode(i).Code
                .Cell(.Rows - 1, 2).Text = uHannaCode(i).ProductName
                .Cell(.Rows - 1, 3).Text = uHannaCode(i).Line
                .Cell(.Rows - 1, 4).Text = uHannaCode(i).Qty
                .Cell(.Rows - 1, 5).Text = uHannaCode(i).Um
                .Cell(.Rows - 1, 6).Text = ""
                .Cell(.Rows - 1, 7).Text = uHannaCode(i).Recipe
                .Cell(.Rows - 1, 8).Text = uHannaCode(i).Mix1 & IIf(Len(uHannaCode(i).Mix2) > 0, ";" & uHannaCode(i).Mix2, "")
                .Cell(.Rows - 1, 9).Text = ""
                .Cell(.Rows - 1, 10).Text = ""
            Next
            
                    
        '.Cell(0, 1).Text = "Code"
        '.Cell(0, 2).Text = "Product Name"
        '.Cell(0, 3).Text = "Line"
        '.Cell(0, 4).Text = "Volume/Weight"
        '.Cell(0, 5).Text = "(um)"
        '.Cell(0, 6).Text = "Q.ty to produce"
        '.Cell(0, 7).Text = "Recipe"
        '.Cell(0, 8).Text = "Mix"
        
            Call SetHannaGridSpecific(Grid1)
            
            .ReadOnly = False
            .Range(0, 4, 0, 5).Merge
            .Cell(0, 4).Alignment = cellCenterCenter
            .Refresh
            .AutoRedraw = True
            
            .Column(3).Width = 0
            .Column(8).AutoFit
            .Column(9).AutoFit
            .Column(4).AutoFit
            .Column(7).AutoFit
        End With
ERR_END:

    On Error GoTo 0
  
    AddCodeInRecipeForProductionGrid = rc
    Exit Function
ERR_ADD:
    rc = False
    MsgBox err.Description
    Resume Next

End Function

Public Function SetHannaGridSpecific(ByVal Grd As Grid, Optional ByVal bProduction As Boolean)

Dim t As Integer
Dim i As Integer

With Grd
    For t = 1 To .Rows - 1
        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = IIf(i > 6, cellCenterCenter, cellLeftCenter)
            If i = 4 Then .Column(i).Alignment = cellRightCenter
            .Cell(t, i).Locked = True

        Next
        
        
        .Cell(t, 1).FontSize = 11
        .Cell(t, 2).FontSize = 9
        .Cell(t, 1).FontBold = True
        .Cell(t, 1).ForeColor = &H404040
        .Cell(t, 6).Alignment = cellCenterCenter
        If .Cols > 10 Then .Cell(t, 10).Alignment = cellCenterCenter
        
        If bProduction Then
            .Cell(t, 7).BackColor = &HF0F0F0  ' &HC0C0C0  ' vbColorResults
            .Cell(t, 7).Alignment = cellCenterCenter
            .Cell(t, 7).Locked = False

        Else
            .Cell(t, 6).BackColor = &HE0E0E0  ' vbColorResults
            .Cell(t, 6).Locked = False
            If .Cols > 10 Then .Cell(t, 10).BackColor = &HE0E0E0 ' vbColorResults
            If .Cols > 10 Then .Cell(t, 10).Locked = False
        End If
        
        
        
    Next
    .Column(1).AutoFit
    
    If bProduction Then
        .Column(7).Width = 200
    Else
        .Column(8).AutoFit
    End If
    
    
    
    .ReadOnly = False
    .Refresh
    .AutoRedraw = True
End With
End Function

Public Function AddRecipeInRecipeForProductionGrid(ByVal Grid1 As Grid, ByVal Grid2 As Grid, ByVal Grid4 As Grid, ByRef iRecipe() As RecipeType) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim AllCount As Integer
Dim strRecipe As String
Dim Recipes() As String
Dim AllRecipes() As String
Dim Quanti As Integer
Dim uRFProduction As RecipeType
Dim uRFProductionClean As RecipeType
On Error GoTo ERR_ADD:

    rc = True
    
    With Grid1
       
        .AutoRedraw = False
        AllCount = 0
        For i = 1 To .Rows - 1
        
            Quanti = 0
            strRecipe = Trim(.Cell(i, 7).Text)
            If Len(strRecipe) > 0 Then
                Call SplitTextStringClassification("", strRecipe, Recipes(), Quanti)
                If Quanti > 0 Then
            
                    For t = 0 To Quanti - 1
                        If AllCount > 0 Then
                        
                            
                            If GetIndexArStrOneDim(AllRecipes(), Recipes(t)) = -1 Then
Aggiungi:
                                If IfRecipeExsists(Recipes(t)) And IfRecipeNotInGrid2(Recipes(t), Grid2) Then
                                    ReDim Preserve AllRecipes(AllCount)
                                    AllRecipes(AllCount) = Recipes(t)
                                    AllCount = AllCount + 1
                                End If
                    
                               
                            End If
                        Else
                            GoTo Aggiungi
                        End If
                    Next
                End If
            End If
    
        Next
        .Refresh
        .AutoRedraw = True
    End With
    If AllCount > 0 Then
        Call AddRecipeInRecipeGrid2(Grid2, Grid4, AllRecipes(), iRecipe(), AllCount)
    End If
    
ERR_END:
    On Error GoTo 0
    AddRecipeInRecipeForProductionGrid = rc
    Exit Function
ERR_ADD:
    rc = False
    MsgBox err.Description
    Resume Next
End Function

Public Function AddRecipeInRecipeGrid2(ByVal Grid2 As Grid, ByVal Grid4 As Grid, ByRef AllRecipes() As String, ByRef sRecipe() As RecipeType, ByVal AllCount As Integer) As Boolean

Dim rc As Boolean
Dim i As Integer
Dim t As Integer

Dim strRecipe As String
Dim iRecipe() As RecipeType
Dim CleanRecipe() As RecipeType

Dim uRFProduction As RecipeType
Dim uRFProductionClean As RecipeType

On Error GoTo ERR_ADD:

rc = True

iRecipe = sRecipe

 sRecipe = CleanRecipe

If AllCount > 0 Then



    
        Dim MaxRecipes As Integer
        MaxRecipes = UBound(iRecipe)
   
        
       With Grid2
                
            .Rows = MaxRecipes + 1
            .AutoRedraw = False

            For i = 0 To AllCount - 1
            
                uRFProduction = uRFProductionClean
                

 
                If MaxRecipes - 1 >= i Then

                    If iRecipe(i + 1).Code = AllRecipes(i) Then
                    
                        uRFProduction = iRecipe(i + 1)

                    Else
                    

                        Call SetMyRecipeByCode(AllRecipes(i), uRFProduction)
                         ReDim Preserve iRecipe(MaxRecipes + 1 + i)
                         iRecipe(MaxRecipes + 1 + i) = uRFProduction

                    End If
                Else

                     Call SetMyRecipeByCode(AllRecipes(i), uRFProduction)
                     ReDim Preserve iRecipe(MaxRecipes + 1 + i)
                     iRecipe(MaxRecipes + 1 + i) = uRFProduction

                End If

                iRecipe(MaxRecipes + 1 + i).bHaveMixes = IfAllMixes(iRecipe(MaxRecipes + 1 + i).Code)
                
                Debug.Print iRecipe(MaxRecipes + 1 + i).Code & "----bHaveMixes " & iRecipe(MaxRecipes + 1 + i).bHaveMixes
                
                If uRFProduction.Code = "" Then GoTo cont2
 
                .AddItem "", False
                
                .Cell(.Rows - 1, 1).Text = uRFProduction.Code
                .Cell(.Rows - 1, 2).Text = uRFProduction.Description
                .Cell(.Rows - 1, 3).Text = uRFProduction.Line
                .Cell(.Rows - 1, 4).Text = uRFProduction.MultipleToProduce
                .Cell(.Rows - 1, 5).Text = uRFProduction.UmMultiple
                .Cell(.Rows - 1, 6).Text = uRFProduction.TotalRecipe
                .Cell(.Rows - 1, 7).Text = uRFProduction.Mix
                .Cell(.Rows - 1, 8).Text = uRFProduction.Density
                .Cell(.Rows - 1, 9).Text = uRFProduction.MinQty & " " & uRFProduction.UmMax
                .Cell(.Rows - 1, 10).Text = uRFProduction.MaxQty & " " & uRFProduction.UmMax
                .Cell(.Rows - 1, 11).Text = uRFProduction.MinQty2 & " " & uRFProduction.UmMinQty
                .Cell(.Rows - 1, 12).Text = uRFProduction.Multiple
                .Cell(.Rows - 1, 13).Text = uRFProduction.UmMultiple
                .Cell(.Rows - 1, 14).Text = uRFProduction.Exp
                .Cell(.Rows - 1, 15).Text = uRFProduction.Procedure
                .Cell(.Rows - 1, 16).Text = uRFProduction.Rev
                .Cell(.Rows - 1, 17).Text = uRFProduction.NoteRev
                .Cell(.Rows - 1, 18).Text = uRFProduction.bIsMix
             
                .Cell(.Rows - 1, 2).FontSize = 9
                
        '.Cell(0, 1).Text = "Recipe"
        '.Cell(0, 2).Text = "Description"
        '.Cell(0, 3).Text = "Cas"
        '.Cell(0, 4).Text = "Q.ty/multiple"
        '.Cell(0, 5).Text = "(um)"
        '.Cell(0, 6).Text = "Theorethical weight"
        '.Cell(0, 7).Text = "Mix"

        
        '.Cell(0, 8).Text = "Density"
        '.Cell(0, 9).Text = "Min Q.ty"
        '.Cell(0, 10).Text = "Max Q.ty"
        '.Cell(0, 11).Text = "Min Q.ty (pcs)"
        '.Cell(0, 12).Text = "Multiple"
        '.Cell(0, 13).Text = "(um)"
        '.Cell(0, 14).Text = "Exp (years)"
        '.Cell(0, 15).Text = "Procedure"
        '.Cell(0, 16).Text = "Revision"
        '.Cell(0, 17).Text = "Note Revision"
        '
cont2:
            
            Next
            
            Call SetRecipesGridSpecifics(Grid2)
            .ReadOnly = False
            .Column(2).AutoFit
            .Range(0, 4, 0, 5).Merge
            .Cell(0, 4).Alignment = cellCenterCenter
            .Range(0, 12, 0, 13).Merge
            .Cell(0, 12).Alignment = cellCenterCenter
                        
             .ReadOnly = True
        End With
    
        '-----------------------------------
        ' griglia TOTALI
        '-----------------------------------
         With Grid4
            .Rows = 1
            .AutoRedraw = False
            
            For i = 1 To Grid2.Rows - 1
                .AddItem "", False
                .Cell(i, 1).Text = Grid2.Cell(i, 1).Text
                .Cell(i, 2).Text = Grid2.Cell(i, 2).Text
                .Cell(i, 3).BackColor = vbColorResults
                .Cell(i, 4).BackColor = vbColorResults
                .Cell(i, 5).BackColor = vbColorResults
                .Cell(i, 3).Alignment = cellRightCenter
                .Cell(i, 4).Alignment = cellRightCenter
                .Cell(i, 5).Alignment = cellRightCenter
                
                
                .Cell(i, 7).Text = Grid2.Cell(i, 9).Text
                .Cell(i, 10).Text = Grid2.Cell(i, 10).Text
                
                .Cell(i, 12).Text = Grid2.Cell(i, 11).Text
                .Cell(i, 13).Text = Grid2.Cell(i, 12).Text
                .Cell(i, 14).Text = Grid2.Cell(i, 13).Text
                .Cell(i, 15).Text = Grid2.Cell(i, 18).Text
            Next
            .Column(3).Alignment = cellRightCenter
            .Column(4).Alignment = cellRightCenter
            .Column(5).Alignment = cellRightCenter
            .Column(12).Alignment = cellRightCenter
            .Column(1).AutoFit
            .Refresh
            .ReadOnly = True
            .AutoRedraw = True
        End With
        
    
    End If
    
ERR_END:
    On Error GoTo 0
   
    sRecipe = iRecipe
   
    AddRecipeInRecipeGrid2 = rc
    Exit Function
ERR_ADD:
    MsgBox err.Description
    rc = False
    Resume Next
End Function

Public Function SetRecipesGridSpecifics(ByVal Grd As Grid)
Dim i As Integer
Dim t As Integer
With Grd
    For t = 1 To .Rows - 1

        For i = 1 To .Cols - 1
            .Column(i).Alignment = IIf(i > 3, cellCenterCenter, cellLeftCenter)
           ' If i = 4 Then .Column(i).Alignment = cellRightCenter
            .Cell(t, i).Locked = True
            
        Next
        .Cell(t, 1).FontSize = 11
        '.Cell(t, 2).FontSize = 9
        .Cell(t, 1).FontBold = True
        .Cell(t, 1).ForeColor = &H404040
        
        .Cell(t, 4).BackColor = &HE0E0E0  ' vbColorResults
        .Cell(t, 6).BackColor = vbColorResults
        .Cell(t, 12).BackColor = vbColorResults
        .Cell(t, 13).BackColor = vbColorResults
        .Cell(t, 4).Alignment = cellCenterCenter
        .Cell(t, 4).Locked = False
        
    Next
    .Column(12).Alignment = cellRightCenter
    .Column(13).Alignment = cellLeftCenter
    
    .Column(1).Width = 100
    .Column(2).Width = 250
    .Column(3).Width = 150
    .Column(5).Width = 0 ' UM multiple!!
    .Refresh
    .ReadOnly = False
    .AutoRedraw = True
End With
End Function

Public Function ResetQuantityHannaCode(ByVal Grid1 As Grid) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
On Error GoTo ERR_ADD:

    rc = True
    
    If F_MsgBox.DoShow("Warning : Delete all quantity in Hanna Codes Table?", "RecipeForProduction") Then
    
        With Grid1
            For i = 1 To .Rows - 1
                .Cell(i, 6).Text = ""
                
            Next
            
            .Refresh
            
        
        End With
    Else
        rc = False
    End If
    

ERR_END:
    On Error GoTo 0
    ResetQuantityHannaCode = rc
    Exit Function
ERR_ADD:
    rc = False
    MsgBox err.Description
    Resume Next
End Function


Public Function AddComponentInGrid7(ByVal Grid As Grid, ByVal Code As String, ByVal Grid4 As Grid, ByRef MixRecipe As RecipeType) As Boolean
Dim rc As Boolean
Dim mrc As Boolean
Dim i As Integer
Dim t As Integer
Dim strMixes As String
Dim Quanti As Integer
Dim Recipes() As String
Dim uRmxRecipeMixes() As RmxRecipe
Dim uRmxRecipeMixesClean As RmxRecipe
On Error GoTo ERR_ADD:

    rc = True
With Grid
    .AutoRedraw = False
    

 
        If Code = "" Then Exit Function
        
        If MixRecipe.Code = Code Then

            uRmxRecipeMixes = MixRecipe.RmxRecipe
        
        Else
              

            mrc = SetRmxRecipeByRecipeCode(Code, uRmxRecipeMixes(), False)
            MixRecipe.bUpdated = mrc
        
        End If
        

 
        For i = LBound(uRmxRecipeMixes) To UBound(uRmxRecipeMixes)
            If uRmxRecipeMixes(i).CHCode <> "" And uRmxRecipeMixes(i).RecipeCode = Code Then
                
        
                .AddItem "", False
                .Cell(.Rows - 1, 1).Text = uRmxRecipeMixes(i).CHCode
                .Cell(.Rows - 1, 2).Text = uRmxRecipeMixes(i).Description
                .Cell(.Rows - 1, 3).Text = uRmxRecipeMixes(i).Cas
                .Cell(.Rows - 1, 4).Text = uRmxRecipeMixes(i).MultipleInCell
                .Cell(.Rows - 1, 5).Text = uRmxRecipeMixes(i).Um
                .Cell(.Rows - 1, 6).Text = PadString(uRmxRecipeMixes(i).Perc)
                .Cell(.Rows - 1, 7).Text = uRmxRecipeMixes(i).TheoreticalWeight
                .Cell(.Rows - 1, 8).Text = uRmxRecipeMixes(i).UmTheoreticalWeight
                .Cell(.Rows - 1, 9).Text = uRmxRecipeMixes(i).Density
                .Cell(.Rows - 1, 10).Text = uRmxRecipeMixes(i).MinQty & " " & uRmxRecipeMixes(i).UmMax
                .Cell(.Rows - 1, 11).Text = uRmxRecipeMixes(i).MaxQty & " " & uRmxRecipeMixes(i).UmMax
                .Cell(.Rows - 1, 12).Text = uRmxRecipeMixes(i).MinQty2 & " " & uRmxRecipeMixes(i).UmMinQty
                .Cell(.Rows - 1, 13).Text = uRmxRecipeMixes(i).Multiple
                .Cell(.Rows - 1, 14).Text = uRmxRecipeMixes(i).UmMultiple

        

       ' .Cell(0, 9).Text = "Density"
       ' .Cell(0, 10).Text = "Min Q.ty"
       ' .Cell(0, 11).Text = "Max Q.ty"
       ' .Cell(0, 12).Text = "Q.ty Recipe (pcs)"
       ' .Cell(0, 13).Text = "Multiple"
       ' .Cell(0, 14).Text = "(um)"
        
             
                    For t = 1 To .Cols - 1
                        .Cell(.Rows - 1, t).Alignment = IIf(t > 8, cellCenterCenter, cellLeftCenter)
                        If t = 4 Then .Cell(.Rows - 1, t).Alignment = cellCenterCenter
                        .Cell(.Rows - 1, t).Locked = True
                        
                    Next
                    .Cell(.Rows - 1, 4).BackColor = vbColorIns
                    .Cell(.Rows - 1, 7).BackColor = vbColorResults
                    .Cell(.Rows - 1, 13).BackColor = vbColorResults
                    .Cell(.Rows - 1, 14).BackColor = vbColorResults
                    .Cell(.Rows - 1, 7).Alignment = cellCenterCenter
                    .Cell(.Rows - 1, 4).Alignment = cellCenterCenter
                    .Cell(.Rows - 1, 6).Alignment = cellCenterCenter
                    .Cell(.Rows - 1, 13).Alignment = cellRightCenter
                    .Cell(.Rows - 1, 14).Alignment = cellLeftCenter
                    
                    
                    

                    With Grid4
                        Dim r As Integer
                        For r = 1 To .Rows - 1
                            If InStr(.Cell(r, 1).Text, uRmxRecipeMixes(i).CHCode) Then
                                GoTo cont:
                            End If
                            
                        Next
                        .AddItem "", False
                        .Cell(.Rows - 1, 1).Text = uRmxRecipeMixes(i).CHCode
                        
                        .Cell(.Rows - 1, 7).Text = uRmxRecipeMixes(i).MinQty & " " & uRmxRecipeMixes(i).UmMax
                        .Cell(.Rows - 1, 10).Text = uRmxRecipeMixes(i).MaxQty & " " & uRmxRecipeMixes(i).UmMax
                        .Cell(.Rows - 1, 12).Text = uRmxRecipeMixes(i).MinQty2 & " " & uRmxRecipeMixes(i).UmMinQty
                        .Cell(.Rows - 1, 13).Text = Int(uRmxRecipeMixes(i).Multiple)
                        .Cell(.Rows - 1, 14).Text = uRmxRecipeMixes(i).UmMultiple
                        .Cell(.Rows - 1, 15).Text = uRmxRecipeMixes(i).bMix
                        .Cell(.Rows - 1, 3).BackColor = vbColorResults
                        .Cell(.Rows - 1, 4).BackColor = vbColorResults
                        .Cell(.Rows - 1, 5).BackColor = vbColorResults
                        .Cell(.Rows - 1, 3).Alignment = cellRightCenter
                        .Cell(.Rows - 1, 4).Alignment = cellRightCenter
                        .Cell(.Rows - 1, 5).Alignment = cellRightCenter
                        
                        
                        For t = 1 To .Cols - 1

                            If uRmxRecipeMixes(i).bMix Then
                                .Cell(.Rows - 1, t).FontBold = True
                                .Cell(.Rows - 1, t).ForeColor = &H644603
                            
                            End If
                            
                        Next
                        
                        
                        
                        .Column(1).AutoFit
                         .ReadOnly = False
                      
                        .Range(0, 13, 0, 14).Merge
                        .Cell(0, 13).Alignment = cellCenterCenter
            
                        .Refresh
                        .ReadOnly = True
                        .AutoRedraw = True
                        
cont:
                    End With
            End If
            
        Next
        .ReadOnly = False
        .Range(0, 4, 0, 5).Merge
        .Cell(0, 4).Alignment = cellCenterCenter
        .Range(0, 7, 0, 8).Merge
        .Cell(0, 7).Alignment = cellCenterCenter
         .Range(0, 13, 0, 14).Merge
        .Cell(0, 13).Alignment = cellCenterCenter
        .Column(5).Width = 0
        .Column(1).AutoFit
        .Cell(0, 7).Alignment = cellLeftCenter
        .ReadOnly = True
        .Refresh
        .AutoRedraw = True
   End With

ERR_END:
    On Error GoTo 0
     
   ' MixRecipe.RmxRecipe = uRmxRecipeMixes

    AddComponentInGrid7 = rc
  
    Exit Function
    
    
ERR_ADD:
    rc = False
    MsgBox err.Description
    Resume Next
End Function


Public Function AddComponentInGrid3(ByVal Grid As Grid, ByVal Code As String) As Boolean
Dim rc As Boolean
Dim mrc As Boolean
Dim i As Integer
Dim t As Integer
Dim strMixes As String
Dim Quanti As Integer
Dim Recipes() As String
Dim uRmxRecipeMixes() As RmxRecipe
Dim uRmxRecipeMixesClean As RmxRecipe
On Error GoTo ERR_ADD:

    rc = True
With Grid
    .AutoRedraw = False
              
        mrc = SetRmxRecipeByRecipeCode(Code, uRmxRecipeMixes(), False)
        
        
        For i = LBound(uRmxRecipeMixes) To UBound(uRmxRecipeMixes)
            If uRmxRecipeMixes(i).CHCode <> "" And uRmxRecipeMixes(i).RecipeCode = Code Then
                
        
                .AddItem "", False
                .Cell(.Rows - 1, 1).Text = uRmxRecipeMixes(i).CHCode
                .Cell(.Rows - 1, 2).Text = uRmxRecipeMixes(i).Description
                .Cell(.Rows - 1, 3).Text = uRmxRecipeMixes(i).Cas
                .Cell(.Rows - 1, 4).Text = PadString(uRmxRecipeMixes(i).Qty)
                .Cell(.Rows - 1, 5).Text = uRmxRecipeMixes(i).Um
                .Cell(.Rows - 1, 6).Text = FormatNumber(uRmxRecipeMixes(i).Perc, 4)
                .Cell(.Rows - 1, 7).Text = ""
                .Cell(.Rows - 1, 8).Text = ""
                .Cell(.Rows - 1, 9).Text = uRmxRecipeMixes(i).Note
                .Cell(.Rows - 1, 10).Text = uRmxRecipeMixes(i).bMix
                .Cell(.Rows - 1, 11).Text = uRmxRecipeMixes(i).CriticalRM
                '.Cell(0, 1).Text = "CH Code"
                '.Cell(0, 2).Text = "Description"
                '.Cell(0, 3).Text = "CAS"
                '.Cell(0, 4).Text = "Q.ty/multiple"
                '.Cell(0, 5).Text = "(um)"
                '.Cell(0, 6).Text = "%"
                '.Cell(0, 7).Text = "Theorethical weight"
                '.Cell(0, 8).Text = "(um)"
                '.Cell(0, 9).Text = "Note"
             
                    For t = 1 To .Cols - 1
                        '.Column(i).Alignment = IIf(t > 6, cellCenterCenter, cellLeftCenter)
                        If t = 4 Then .Column(t).Alignment = cellRightCenter
                        .Cell(.Rows - 1, t).Locked = True
                        
                        If uRmxRecipeMixes(i).bMix Then
                            .Cell(.Rows - 1, t).FontBold = True
                            .Cell(.Rows - 1, t).ForeColor = &H644603
                        
                        End If
                        
                        
                        If Len(uRmxRecipeMixes(i).CriticalRM) > 0 Then
                            .Cell(.Rows - 1, t).FontBold = True
                            .Cell(.Rows - 1, t).ForeColor = &H40C0&
                            
                        End If
                    Next
                    
                    .Cell(.Rows - 1, 7).BackColor = vbColorResults
                    .Cell(.Rows - 1, 7).Alignment = cellCenterCenter
                   ' .Cell(t, 4).Locked = False
                '
                
            End If
            
        Next
        .Column(0).Width = 0
        .ReadOnly = False
        .Range(0, 7, 0, 8).Merge
        .Cell(0, 7).Alignment = cellLeftCenter
            
        .Column(1).AutoFit
        .Column(9).AutoFit
        .ReadOnly = True
        .Refresh
        .AutoRedraw = True
   End With

ERR_END:
    On Error GoTo 0
    AddComponentInGrid3 = rc
    Exit Function
ERR_ADD:
    rc = False
    MsgBox err.Description
    Resume Next
End Function


Public Function ViewRecipesRFP(ByRef iRecipes() As RecipeType, ByVal Grid1 As Grid, ByVal Grid2 As Grid, ByVal Grid4 As Grid, ByVal Grid5 As Grid, ByVal bView As Boolean) As Boolean
Dim i As Integer
Dim bIsMix As Boolean

    If bView Then
    
        With Grid1
            If .Rows < 1 Then Exit Function
            .AutoRedraw = False
            For i = 1 To .Rows - 1
                .RowHeight(i) = 25
              
            Next
            .Refresh
            .AutoRedraw = True
        End With
        
        With Grid2
            If .Rows < 1 Then Exit Function
            .AutoRedraw = False
            For i = 1 To .Rows - 1
                .RowHeight(i) = 25
                iRecipes(i).bHide = False
            Next
            .Refresh
            .AutoRedraw = True
        End With
        
        With Grid4
            If .Rows < 1 Then Exit Function
            .AutoRedraw = False
            For i = 1 To .Rows - 1
                .RowHeight(i) = 25
            Next
            .Refresh
            .AutoRedraw = True
        End With
        
        With Grid5
            If .Rows < 1 Then Exit Function
            .AutoRedraw = False
            For i = 1 To .Rows - 1
                .RowHeight(i) = 25
            Next
            .Refresh
            .AutoRedraw = True
        End With
        
        
    Else
        
        Dim strValue As String
        Dim dValue As Double
        
        With Grid1
            If .Rows < 1 Then Exit Function
            
            .AutoRedraw = False
                For i = 1 To .Rows - 1
                    strValue = .Cell(i, 6).Text
                    strValue = Replace(LCase(strValue), "kg", "")
                    strValue = Replace(LCase(strValue), "l", "")
                    
                    
                    
                    If strValue <> "" Then
                        dValue = CDbl(strValue)
                        If dValue > 0 Then
                        Else
                            .RowHeight(i) = 0
                          
                        End If
                    Else
                        .RowHeight(i) = 0
                    
                    End If
                Next
            .Refresh
            .AutoRedraw = True
        End With
        
        With Grid2
            If .Rows < 1 Then Exit Function
            
            .AutoRedraw = False
            For i = 1 To .Rows - 1
                bIsMix = IIf(.Cell(i, 18).Text <> "", .Cell(i, 18).Text, False)
                strValue = .Cell(i, 6).Text
                
                strValue = Replace(LCase(strValue), "kg", "")
                strValue = Replace(LCase(strValue), "l", "")
                If strValue = "" Then GoTo cont
                If bIsMix Then
                
                    strValue = Grid4.Cell(i, 3).Text
                    strValue = Replace(LCase(strValue), "kg", "")
                    strValue = Replace(LCase(strValue), "l", "")
                   
                
                End If
                
                If strValue <> "" Then
                    dValue = CDbl(strValue)
                    If dValue > 0 Then
                    Else
                        .RowHeight(i) = 0
                        If Grid4.Rows > i Then Grid4.RowHeight(i) = 0
                        If Grid5.Rows > i Then Grid5.RowHeight(i) = 0
                        iRecipes(i).bHide = True
                    End If
                Else
                    .RowHeight(i) = 0
                    If Grid4.Rows > i Then Grid4.RowHeight(i) = 0
                    If Grid5.Rows > i Then Grid5.RowHeight(i) = 0
                    iRecipes(i).bHide = True
                End If
cont:
                
            Next
            .Refresh
            .AutoRedraw = True
        End With
      
    End If
End Function





Public Function CheckRfPBeforeSave(ByRef uRecipeForProduction As RecipeForProduction) As Boolean
    
    Dim rc As Boolean
    Dim filterString As String
    Dim strRecipe As String
    
    
  
    rc = True
    
    With uRecipeForProduction

        
        strRecipe = GetStrRecipe(.Recipes)
    
    filterString = "RecipeWeek='" & Trim(.numPrepWeek) & "' and PlannedPreparation='" & Trim(.PlannedPrepWeek) & "' and Recipe='" & Trim(strRecipe) & "'"

    End With
    With dbTabReceiptForProduction
        .filter = ""
        .filter = filterString
        If .EOF Then
        Else
           ' If uRecipeForProduction.fileNameRecForProd <> !FileName Then
                rc = False
          '  End If
        End If
    End With

    CheckRfPBeforeSave = rc


End Function
