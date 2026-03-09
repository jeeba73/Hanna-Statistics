Attribute VB_Name = "Main_01_Grid"
Option Explicit


Private Grid() As Grid

Public Function SetAllMainGrid(ByVal Grid As Variant) As Boolean



' 0 RfP


    
    Call SetRecipeForProductionGrid(Grid(0))

' 1 preparation

' 4  material requisition list

    
    Call SetRecipeGrid(Grid(4))

' 5  material requisition specifics


    
   Call SetMaterialReqGrid(Grid(5), False)


    
    Call SetPreparationRecipeGrid(Grid(1))


' 2 QC : recipe list

    
    Call SetRecipesQCGrid(Grid(2))

' 3 QC : details

    
    Call SetQCperRecipeGrid(Grid(3))
    
    
    Call SetProductionTable(Grid(6))
  
    




End Function


Private Sub SetRecipeForProductionGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
        .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        
        
        .Cols = 14
        
        
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        .RowHeight(0) = 50
        .Cell(0, 1).Text = "Line"
        .Cell(0, 2).Text = "Date Recipe"
        .Cell(0, 3).Text = "# Prep. Week"
        .Cell(0, 4).Text = "Pl. Prep Week"
        .Cell(0, 5).Text = "Pl. Reference"
        .Cell(0, 6).Text = "Operator"
    
        .Cell(0, 7).Text = "Recipes"

        .Cell(0, 8).Text = "Description"
        .Cell(0, 9).Text = "MR Printed"
        .Cell(0, 10).Text = "MR Number"
        .Cell(0, 11).Text = "Note"
        
        
        .Cell(0, 12).Text = "FileName"
        .Cell(0, 13).Text = "ID"
        
        

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 110
            .Cell(0, i).FontBold = True
            
        Next
        
        
        
        .Column(3).Width = 140
        .Column(4).Width = 140
        .Column(5).Width = 140
        .Column(8).Width = 250
        .Column(10).Width = 150
        .Column(11).Width = 500
        
        .Column(12).Width = 0
        .Column(13).Width = 0
        .Column(9).CellType = cellCheckBox
        
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
  

End Sub



Private Sub SetRecipeGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 11
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Recipe"
        .Cell(0, 2).Text = "Description"
        .Cell(0, 3).Text = "Line"
        .Cell(0, 4).Text = "Mix"
        .Cell(0, 5).Text = "MR Print"
        .Cell(0, 6).Text = "MR Number"
        
        .Cell(0, 7).Text = "MR CheckOut"
        .Cell(0, 8).Text = "Date CheckOut"
        
        .Cell(0, 9).Text = "Q.ty to produce"
        .Cell(0, 10).Text = "Recipe#"


        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 120
            .Cell(0, i).FontBold = True
            

            
        Next
        .Column(5).CellType = cellCheckBox
        .Column(7).CellType = cellCheckBox
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellCenterCenter
        .Column(8).Alignment = cellCenterCenter
        .Column(9).Alignment = cellCenterCenter
        .Column(10).Width = 0
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   


End Sub




Private Sub SetPreparationRecipeGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      Main Preparation
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 19
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        '!Recipe = Trim(RecipeName)
        '!PlanningReference = Trim(PlanningReference)
        '!DataRecipe = Trim(DataRecipe)
        '!RecipeWeek = Trim(NumPrepWeek)
        '!PlannedPreparation = Trim(PlannedPreparation)
        '!Operator = Trim(Operator)
        '!bClosed = False
        '!Note = ""
        '!FileName = NewFileName
        .Cell(0, 1).Text = "Line"
        .Cell(0, 2).Text = "Recipe"
        .Cell(0, 3).Text = "Hanna Code"
        .Cell(0, 4).Text = "# Prep. Week"
        .Cell(0, 5).Text = "Planned Prep."
        .Cell(0, 6).Text = "Data Recipe"
        .Cell(0, 7).Text = "Qty To Produce"
        
        .Cell(0, 8).Text = "Qty Produced"
        .Cell(0, 9).Text = "FileName"
        
        .Cell(0, 13).Text = "ID"
        .Cell(0, 14).Text = "QC Status"
        .Cell(0, 15).Text = "Note"
        .Cell(0, 16).Text = "NoPreparation"
        
         .Cell(0, 17).Text = "Description"
         .Cell(0, 18).Text = "Lot Number"
        
        
        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 130
            .Cell(0, i).FontBold = True
            

            
        Next
        .Column(14).Alignment = cellCenterCenter
        .Column(4).Alignment = cellCenterCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellRightCenter
        .Column(8).Alignment = cellRightCenter
        
        .Column(3).Width = 200
        .Column(4).Width = 100
      
        .Column(9).Width = 0
        
        
        .Column(10).Width = 0
        .Column(11).Width = 7
        .Column(12).Width = 20
        .Column(13).Width = 0
        '.Column(14).Width = 150
        .Column(16).Width = 0
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   


End Sub




Private Sub SetRecipesQCGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      Main Preparation
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 17
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11
        
        .DefaultFont.Name = "Calibri"
        

        .Cell(0, 1).Text = "Line"
        .Cell(0, 2).Text = "Recipe"
        .Cell(0, 3).Text = "Hanna Code"
        .Cell(0, 4).Text = "# Prep. Week"
        .Cell(0, 5).Text = "Planned Prep."
        .Cell(0, 6).Text = "Data Recipe"
        .Cell(0, 7).Text = "QC Status"
        .Cell(0, 8).Text = "QC Date"
        .Cell(0, 9).Text = "FileName"
        .Cell(0, 12).Text = "ID"
        .Cell(0, 13).Text = "Operator"
        .Cell(0, 14).Text = "Note"
        
        .Cell(0, 15).Text = "Description"
        .Cell(0, 16).Text = "Lot"
        
        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 130
            .Cell(0, i).FontBold = True
        Next

        .Column(4).Alignment = cellCenterCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellCenterCenter
        .Column(8).Alignment = cellCenterCenter
        
        .Column(3).Width = 200
        .Column(4).Width = 100
        .Column(1).Width = 0
        .Column(9).Width = 0
        .Column(10).Width = 0
        .Column(11).Width = 0
        .Column(12).Width = 0
        

        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   


End Sub



Public Sub SetQCperRecipeGrid(ByVal Grd As Grid)
Dim i As Integer
        
        '------------------------------------------------
        '      List of QC per Recipe
        '------------------------------------------------
        
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 9
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11
        .DefaultFont.Name = "Calibri"

        .Cell(0, 1).Text = "QC status"
        .Cell(0, 2).Text = "QC Date"
        .Cell(0, 3).Text = "Operator"
        .Cell(0, 4).Text = "Note"
        
        .Cell(0, 5).Text = "Registration"
        .Cell(0, 6).Text = "QC Operator"
        .Cell(0, 7).Text = "Correction"
        .Cell(0, 8).Text = "Correction Date"
        
    
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 150
            .Cell(0, i).FontBold = True
            

            
        Next
        .Column(1).Width = 100
        .Column(1).Alignment = cellCenterCenter
        .Column(2).Alignment = cellCenterCenter
        .Column(3).Alignment = cellCenterCenter
        .Column(4).Alignment = cellRightCenter

      
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   


End Sub

Private Sub SetProductionTable(ByVal Grd As Grid)

Dim i As Integer
        '------------------------------------------------
        '
        '      Set Production Table
        '
        '------------------------------------------------
    With Grd
      
      
        
        .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        
        
        .Cols = 14
        
        
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        .RowHeight(0) = 50
        .Cell(0, 1).Text = "Hanna Code"
        .Cell(0, 2).Text = "Line"
        .Cell(0, 3).Text = "Date Recipe"
        .Cell(0, 4).Text = "Recipe"
        .Cell(0, 5).Text = "Mix"
        .Cell(0, 6).Text = "Pl. Ref."
        
        .Cell(0, 7).Text = "Preparation Date"
        .Cell(0, 8).Text = "Preparation Week"
        .Cell(0, 9).Text = "# Prep. Week"
        
        
        .Cell(0, 10).Text = "FileName"
        .Cell(0, 11).Text = "ID"
        .Cell(0, 12).Text = ""
        .Cell(0, 13).Text = "Preparation Lot"

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellRightCenter
            .Column(i).Width = 150
            .Cell(0, i).FontBold = True
            
        Next
        
        .Column(3).Alignment = cellCenterCenter
        .Column(1).Alignment = cellLeftCenter
        .Column(2).Alignment = cellLeftCenter
        .Column(4).Alignment = cellLeftCenter
      
      
        
        .Column(7).Width = 130
        .Column(8).Width = 130
        .Column(9).Width = 130
        .Column(5).Width = 0
        .Column(6).Width = 0
        .Column(10).Width = 0
        .Column(11).Width = 0
        .Column(12).Width = 8
        .Column(13).Width = 200
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
  

End Sub
