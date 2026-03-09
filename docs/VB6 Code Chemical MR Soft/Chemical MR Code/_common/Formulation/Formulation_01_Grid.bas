Attribute VB_Name = "Formulation_01_Grid"
Option Explicit

Private Grid() As Grid

Public Function SetAllRecipeForSTDPreparationGrid(ByVal Grid As Variant) As Boolean



' 0 code


    
    Call SetCodeGrid(Grid(0))

' 1 recipe


    
    Call SetRecipeGrid(Grid(1), True)
    
    Call SetMixGrid(Grid(6))
  

' 2 component

    
    Call SetComponentGrid(Grid(2))

' 3 totals

    
    Call SetTotalsGrid(Grid(3))

' 4 STDPreparation

    
    Call SetPackagingGrid(Grid(4))

' 5 Material Requisition

   Call SetMaterialReqGrid(Grid(5))



End Function


Public Sub SetCodeGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForSTDPreparation  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click

        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 9
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Code"
        .Cell(0, 2).Text = "Product Name"
        .Cell(0, 3).Text = "Line"
        .Cell(0, 4).Text = "Volume/Weight"
        .Cell(0, 5).Text = "(um)"
        .Cell(0, 6).Text = "Q.ty to produce"
        .Cell(0, 7).Text = "Recipe"
        .Cell(0, 8).Text = "Mix"

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 150
            .Cell(0, i).FontBold = True
            If i > 7 Then .Column(i).Width = 100
        Next
        .Column(1).Width = 100
        .Column(3).Width = 100
        .Column(2).Width = 250
        .Column(4).Width = 120
        .Column(5).Width = 80
        .Column(7).Width = 200
        .Column(8).Width = 400
        .Column(4).Alignment = cellRightCenter
        .Column(8).Alignment = cellCenterCenter
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   

End Sub


Private Sub SetRecipeGrid(ByVal Grd As Grid, ByVal bValue As Boolean)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForSTDPreparation  TABELLA Codici
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
        
        
        .Cell(0, 1).Text = "Recipe"
        .Cell(0, 2).Text = "Description"
        
        
        .Cell(0, 3).Text = "Line"
        .Cell(0, 4).Text = "Multiple to Prod."
        .Cell(0, 5).Text = "(um)"
        
        .Cell(0, 6).Text = "Total"
        
        .Cell(0, 7).Text = "Mix"
        
        .Cell(0, 8).Text = "Density"
        .Cell(0, 9).Text = "Min Q.ty"
        .Cell(0, 10).Text = "Max Q.ty"
        .Cell(0, 11).Text = "Min Q.ty (pcs)"
        .Cell(0, 12).Text = "Multiple"
        .Cell(0, 13).Text = "(um)"
        .Cell(0, 14).Text = "Exp (years)"
        .Cell(0, 15).Text = "Procedure"
        .Cell(0, 16).Text = "Revision"
        .Cell(0, 17).Text = "Note Revision"
        .Cell(0, 18).Text = "bIsMix"
        
        

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 120
            .Cell(0, i).FontBold = True
            

            
        Next
        
        .Column(3).Width = 80
        .Column(4).Width = 150
       ' .Column(4).Alignment = cellRightCenter
        .Column(5).Width = 40
        .Column(6).Width = 150 'IIf(bValue, 0, 150)
        .Column(8).Width = 80
        
        .Column(13).Width = 40
        .Column(14).Width = 80
        .Column(15).Width = 400
        .Column(6).Alignment = cellCenterCenter
         .Column(7).Width = 0
         .Column(18).Width = 0
        '.Column(4).Alignment = cellRightCenter
        
        
         For i = 7 To .Cols - 1

                .Column(i).Width = IIf(bValue, .Column(i).Width, 0)

            
          Next
          
        .Column(9).Width = 0
        .Column(10).Width = 0
        .Column(11).Width = 0
        .Column(14).Width = 0
        .Column(15).Width = 0
        .Column(16).Width = 0
        .Column(17).Width = 0
          
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   


End Sub
Public Sub SetMixGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForSTDPreparation  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 15
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Code"
        .Cell(0, 2).Text = "Description"
        .Cell(0, 3).Text = "CAS"
        .Cell(0, 4).Text = "Multiple To Prod."
        .Cell(0, 5).Text = "(um)"
        .Cell(0, 6).Text = "%"
        .Cell(0, 7).Text = "Theorethical Weight"
        .Cell(0, 8).Text = "(um)"
        
        .Cell(0, 9).Text = "Density"
        .Cell(0, 10).Text = "Min Q.ty"
        .Cell(0, 11).Text = "Max Q.ty"
        .Cell(0, 12).Text = "Min Q.ty (pcs)"
        .Cell(0, 13).Text = "Multiple"
        .Cell(0, 14).Text = "(um)"
        
        
        
        

        For i = 1 To .Cols - 1
            .Column(i).Alignment = IIf(i > 8, cellCenterCenter, cellLeftCenter)
            .Column(i).Width = 100
            .Cell(0, i).FontBold = True
            
        Next
        
        .Column(2).Width = 250
        .Column(3).Width = 80
        .Column(4).Width = 150
        .Column(5).Width = 40
        .Column(6).Width = 80
        .Column(7).Width = 150
        .Column(8).Width = 40
        .Column(14).Width = 40
        .Column(6).Alignment = cellCenterCenter
        .Column(13).Alignment = cellRightCenter
        .Column(14).Alignment = cellLeftCenter
        .Column(7).Alignment = cellLeftCenter
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With

End Sub

Public Sub SetComponentGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForSTDPreparation  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 12
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Code"
        .Cell(0, 2).Text = "Description"
        .Cell(0, 3).Text = "CAS"
        .Cell(0, 4).Text = "Q.ty/multiple"
        .Cell(0, 5).Text = "(um)"
        .Cell(0, 6).Text = "%"
        .Cell(0, 7).Text = "Theorethical weight"
        .Cell(0, 8).Text = "(um)"
        .Cell(0, 9).Text = "Note"
        .Cell(0, 10).Text = "bMix"
        .Cell(0, 11).Text = "Critical RM"
        

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 100
            .Cell(0, i).FontBold = True
            
        Next
        
        .Column(2).Width = 250
        .Column(3).Width = 80
        .Column(5).Width = 80
        .Column(6).Width = 80
        .Column(7).Width = 150
        .Column(8).Width = 80
        .Column(9).Width = 200
        .Column(10).Width = 0
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellLeftCenter
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With

End Sub

Private Sub SetTotalsGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForSTDPreparation  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 16
        
        
        .RowHeight(0) = 35
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Recipe"
        .Cell(0, 2).Text = "Description"
        
        
        .Cell(0, 3).Text = "Total Weight (Kg)"
        .Cell(0, 4).Text = "Total Volume (L)"
        .Cell(0, 5).Text = "Multiple"
        
        .Cell(0, 7).Text = "Min"
        '.Cell(0, 8).Text = "Max"
        
        
        
        .Cell(0, 10).Text = "Max"
       ' .Cell(0, 11).Text = "Max"
        .Cell(0, 12).Text = "Min pcs"
        .Cell(0, 13).Text = "Multiple"
        .Cell(0, 15).Text = "bMix"

        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            If i > 6 Then .Column(i).Alignment = cellCenterCenter
            .Column(i).Width = 130
            .Cell(0, i).FontBold = True
            
        Next
        
       ' .Column(0).Width = 0
        .Column(2).Width = 0

        .Column(6).Width = 20
        
        .Column(7).Width = 70
        .Column(8).Width = 10
        .Column(9).Width = 0
        .Column(10).Width = 70
       .Column(11).Width = 10
       
        .Column(12).Width = 100
        .Column(13).Width = 0
        .Column(14).Width = 0
        .Column(15).Width = 0

        
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
 

End Sub

Private Sub SetPackagingGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForSTDPreparation  Bottiling
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 7
        
        
      
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        .Cell(0, 1).Text = "Recipe"
        .Cell(0, 2).Text = "Line"
        .Cell(0, 3).Text = "STDPreparation Way"
        .Cell(0, 4).Text = "Prod. speed ( pcs/min )"
        .Cell(0, 5).Text = "Est. time machine ( h )"
        .Cell(0, 6).Text = "Est. Shift machine ( S )"
        
        .Column(3).CellType = cellComboBox
        .ComboBox(3).Locked = True
        
        For i = 2 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
            .Column(i).Width = 160
        Next
        .Column(0).Width = 0
        .Column(1).Alignment = cellLeftCenter
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
  

End Sub

