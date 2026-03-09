Attribute VB_Name = "Formulation_05_FillFromFile"
Option Explicit



    
Public Function FillGridRfPFromFile(ByVal Grid As Grid, ByRef uRecipeForProduction As RecipeForProduction, ByVal Index As Integer)

Dim i As Integer

    Select Case Index
        Case 1
            ' Hanna Codes
            Call GetCodeGrid(Grid, uRecipeForProduction)
        Case 2
            ' Recipes
            Call GetRecipeGrid(Grid, uRecipeForProduction)
        Case 3
            ' Components
        Case 4
            ' totals
            Call GetTotalsGrid(Grid, uRecipeForProduction)
        Case 5
            ' Bottling / Packaging
            Call GetPackagingGrid(Grid, uRecipeForProduction)

    
    
    End Select

    
End Function
    


Private Sub GetCodeGrid(ByVal Grid As Grid, ByRef uRecipeForProduction As RecipeForProduction)
Dim i As Integer
Dim HannaCount As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grid
      
      
      .Rows = 1

        .AutoRedraw = False
    
        
        '.Cell(0, 1).Text = "Code"
        ''.Cell(0, 2).Text = "Product Name"
        '.Cell(0, 3).Text = "Line"
        '.Cell(0, 4).Text = "Volume/Weight"
        '.Cell(0, 5).Text = "(um)"
        '.Cell(0, 6).Text = "Q.ty to produce"
        '.Cell(0, 7).Text = "Recipe"
        '.Cell(0, 8).Text = "Mix"
        
        
        With uRecipeForProduction
            HannaCount = .HannaCodesCount
            
            For i = 1 To HannaCount
                
                Grid.AddItem "", False
                Grid.Cell(Grid.Rows - 1, 1).Text = .HannaCodes(i).Code
                Grid.Cell(Grid.Rows - 1, 2).Text = .HannaCodes(i).ProductName
                Grid.Cell(Grid.Rows - 1, 3).Text = .HannaCodes(i).Line
                Grid.Cell(Grid.Rows - 1, 4).Text = .HannaCodes(i).Qty
                Grid.Cell(Grid.Rows - 1, 5).Text = .HannaCodes(i).Um
                Grid.Cell(Grid.Rows - 1, 6).Text = .HannaCodes(i).QtyToProduce
                Grid.Cell(Grid.Rows - 1, 7).Text = .HannaCodes(i).Recipe
                Grid.Cell(Grid.Rows - 1, 8).Text = .HannaCodes(i).Mix1 & IIf(Len(.HannaCodes(i).Mix2) > 0, ";" & .HannaCodes(i).Mix2, "")
                Grid.Cell(Grid.Rows - 1, 10).Text = .HannaCodes(i).LotNumber
            Next
        
        
        End With

         Call SetHannaGridSpecific(Grid)
         .Column(3).Width = 0
         .Column(4).AutoFit
         
        .Refresh
        .AutoRedraw = True
    End With
   

End Sub



Private Sub GetRecipeGrid(ByVal Grid As Grid, ByRef uRecipeForProduction As RecipeForProduction)
Dim i As Integer
Dim RecipeCount As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grid
      
      
      .Rows = 1

        .AutoRedraw = False
        
       
        '.Cell(0, 1).Text = "Recipe"
        '.Cell(0, 2).Text = "Description"
        '.Cell(0, 3).Text = "Line"
        '.Cell(0, 4).Text = "Multiple to Prod."
        '.Cell(0, 5).Text = "(um)"
        '.Cell(0, 6).Text = "Total"
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

        With uRecipeForProduction
            RecipeCount = .RecipeCount
            
            For i = 1 To RecipeCount
                
                Grid.AddItem "", False
                Grid.Cell(Grid.Rows - 1, 1).Text = .Recipes(i).Code
                Grid.Cell(Grid.Rows - 1, 2).Text = .Recipes(i).Description
                Grid.Cell(Grid.Rows - 1, 3).Text = .Recipes(i).Line
                Grid.Cell(Grid.Rows - 1, 4).Text = .Recipes(i).MultipleToProduce
                Grid.Cell(Grid.Rows - 1, 5).Text = .Recipes(i).UmMultiple
                Grid.Cell(Grid.Rows - 1, 6).Text = PadString(.Recipes(i).TotalRecipe)
                Grid.Cell(Grid.Rows - 1, 7).Text = .Recipes(i).Mix
                Grid.Cell(Grid.Rows - 1, 8).Text = .Recipes(i).Density
                Grid.Cell(Grid.Rows - 1, 9).Text = .Recipes(i).MinQty & " " & .Recipes(i).UmMax
                Grid.Cell(Grid.Rows - 1, 10).Text = .Recipes(i).MaxQty & " " & .Recipes(i).UmMax
                Grid.Cell(Grid.Rows - 1, 11).Text = .Recipes(i).MinQty2 & " " & .Recipes(i).UmMinQty
                Grid.Cell(Grid.Rows - 1, 12).Text = .Recipes(i).Multiple
                Grid.Cell(Grid.Rows - 1, 13).Text = .Recipes(i).UmMultiple
                Grid.Cell(Grid.Rows - 1, 14).Text = .Recipes(i).Exp
                Grid.Cell(Grid.Rows - 1, 15).Text = .Recipes(i).Procedure
                Grid.Cell(Grid.Rows - 1, 16).Text = .Recipes(i).Rev
                Grid.Cell(Grid.Rows - 1, 17).Text = .Recipes(i).NoteRev
                Grid.Cell(Grid.Rows - 1, 18).Text = .Recipes(i).bIsMix
                Grid.Cell(Grid.Rows - 1, 2).FontSize = 9
                Grid.RowHeight(Grid.Rows - 1) = IIf(.Recipes(i).bHide, 0, Grid.RowHeight(Grid.Rows - 1))
            Next
        
        
        End With

         Call SetRecipesGridSpecifics(Grid)
        
        .Column(2).AutoFit
       .Refresh
        .AutoRedraw = True
    End With
   


End Sub

Private Sub GetTotalsGrid(ByVal Grid As Grid, ByRef uRecipeForProduction As RecipeForProduction)
Dim i As Integer
Dim TotalCount As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grid

      .Rows = 1

      
        
        '.Cell(0, 1).Text = "Recipe"
        '.Cell(0, 2).Text = "Description"
        '.Cell(0, 3).Text = "Total Weight (Kg)"
        '.Cell(0, 4).Text = "Total Volume (L)"
        '.Cell(0, 5).Text = "Multiple"
        '.Cell(0, 7).Text = "Min"
        '.Cell(0, 8).Text = "Max"
        '.Cell(0, 10).Text = "Min"
        '.Cell(0, 11).Text = "Max"
        '.Cell(0, 12).Text = "Min pcs"
        '.Cell(0, 13).Text = "Multiple"
        
        

        With uRecipeForProduction
        
            For i = 1 To .TotalCount
                Grid.AddItem "", False
                Grid.Cell(i, 1).Text = .TotalGrid(i).Recipe
                Grid.Cell(i, 2).Text = .TotalGrid(i).Description
                Grid.Cell(i, 3).Text = PadString(.TotalGrid(i).TotalWeighKg)
                Grid.Cell(i, 4).Text = PadString(.TotalGrid(i).TotalWeighL)
                Grid.Cell(i, 5).Text = FormatNumber(.TotalGrid(i).TotalMultiple, 1)
                
                Grid.Cell(i, 8).BackColor = IIf(.TotalGrid(i).CkMin, &H8000&, &HC0&)
                Grid.Cell(i, 11).BackColor = IIf(.TotalGrid(i).CkMax, &H8000&, &HC0&)
                
                
                Grid.Cell(i, 7).Text = .TotalGrid(i).Min
                Grid.Cell(i, 10).Text = .TotalGrid(i).Max
                Grid.Cell(i, 12).Text = .TotalGrid(i).Minpcs
                Grid.Cell(i, 13).Text = Int(.TotalGrid(i).Multiple)
                
                Grid.Cell(i, 15).Text = .TotalGrid(i).bMix
                
                
                
                Grid.Cell(i, 3).BackColor = vbColorResults
                Grid.Cell(i, 4).BackColor = vbColorResults
                Grid.Cell(i, 5).BackColor = vbColorResults
                Grid.Cell(i, 3).Alignment = cellRightCenter
                Grid.Cell(i, 4).Alignment = cellRightCenter
                Grid.Cell(i, 5).Alignment = cellRightCenter
                Grid.RowHeight(i) = IIf(.Recipes(i).bHide, 0, Grid.RowHeight(i))
            Next
            
        End With
        
        .Column(3).Alignment = cellRightCenter
        .Column(4).Alignment = cellRightCenter
        .Column(5).Alignment = cellRightCenter
        .Column(12).Alignment = cellRightCenter
        .Column(1).AutoFit
        .Refresh
        .ReadOnly = True
        .AutoRedraw = True
    End With
   
 

End Sub


Private Sub GetPackagingGrid(ByVal Grid As Grid, ByRef uRecipeForProduction As RecipeForProduction)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForProduction  Bottiling
        '------------------------------------------------
    With Grid
      
      
      .Rows = 1

        .AutoRedraw = False

        
        With uRecipeForProduction
        
            For i = 1 To .PackagingCount
                Grid.AddItem "", False
                Grid.Cell(i, 1).Text = .Packaging(i).Recipe
                Grid.Cell(i, 2).Text = .Packaging(i).Line
                Grid.Cell(i, 3).Text = .Packaging(i).Production
                Grid.Cell(i, 4).Text = .Packaging(i).Head
                Grid.Cell(i, 5).Text = .Packaging(i).Speed
                Grid.Cell(i, 6).Text = .Packaging(i).EstTimeH
                Grid.Cell(i, 7).Text = .Packaging(i).EsttimeD
                Grid.Cell(i, 3).BackColor = vbColorIns
                Grid.Cell(i, 4).BackColor = vbColorIns
                
                Grid.RowHeight(i) = IIf(.Recipes(i).bHide, 0, Grid.RowHeight(i))
           
            Next
                Grid.Column(1).Locked = True
                Grid.Column(2).Locked = True
                Grid.Column(3).Locked = False
                Grid.Column(4).Locked = False
                Grid.Column(5).Locked = True
                Grid.Column(6).Locked = True
                Grid.Column(7).Locked = True
        End With
        .ReadOnly = False
        .Refresh
        .AutoRedraw = True
    End With
   
  

End Sub


