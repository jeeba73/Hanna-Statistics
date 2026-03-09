Attribute VB_Name = "Main_01_Grid"
Option Explicit


Private Grid() As Grid

Public Function SetAllMainGrid(ByRef Grid As Variant) As Boolean



' 0 RfP


    
    Call SetStockTable(Grid(0))


    
    Call SetPreparationGrid(Grid(1))


' 2 QC : recipe list

    
    Call SetRecipeQCGrid(Grid(2))





End Function


Public Sub SetStockTable(ByVal Grd As Grid)
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
        
        
        .Cols = 22
        
        .FrozenCols = 2
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        .RowHeight(0) = 50
        .Cell(0, 1).Text = "Code"
        
        .Cell(0, 2).Text = "Supplier"
        .Cell(0, 3).Text = "NMP"
        
        .Cell(0, 4).Text = "Bottle"
        .Cell(0, 5).Text = "Description"
        .Cell(0, 6).Text = "MR Lot"
        .Cell(0, 7).Text = "Purity  %"
        .Cell(0, 8).Text = "MR Value"
    
        .Cell(0, 9).Text = "MR Unit"

        .Cell(0, 10).Text = "Location"
        .Cell(0, 11).Text = "QTY"
        .Cell(0, 12).Text = "Unit"
        .Cell(0, 13).Text = "Arrived"
        
        .Cell(0, 14).Text = "Open"
        .Cell(0, 15).Text = "Finished"
        .Cell(0, 16).Text = "Supplier EXP"
        .Cell(0, 17).Text = "MR EXP"
        .Cell(0, 18).Text = "Status"
        .Cell(0, 19).Text = "Note"

        .Cell(0, 20).Text = "ID"
        
        .Cell(0, 21).Text = "U"

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 110
            .Cell(0, i).FontBold = True
            
        Next
        
        
      
        .Column(20).Width = 0
        
        
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
  

End Sub


Public Sub SetPreparationGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      Main Preparation
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
         .SelectionMode = cellSelectionByRow
        .DefaultRowHeight = 25
        
        
        .Cols = 21
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        

        .Cell(0, 1).Text = "HannaCode"
        .Cell(0, 2).Text = "Description"
        .Cell(0, 3).Text = "MRCode"
        .Cell(0, 4).Text = "Data Prep."
        .Cell(0, 5).Text = "Hour Prep."
        
        .Cell(0, 6).Text = "PrepWeek"
        .Cell(0, 7).Text = "Operator"
        
        .Cell(0, 8).Text = "QtyToProduce"
        .Cell(0, 9).Text = "Unit"
        
        .Cell(0, 10).Text = "STD Matrix"
        .Cell(0, 11).Text = "STD Exp (days)"
        .Cell(0, 12).Text = "STD Storage"
        
     
        .Cell(0, 13).Text = "Note"
        
        
        .Cell(0, 14).Text = "MS Type"
        .Cell(0, 15).Text = "Closed"
        .Cell(0, 16).Text = "ID"
        .Cell(0, 17).Text = "Filename"
        
        
        
        .Cell(0, 18).Text = "Close Date"
        .Cell(0, 19).Text = "Excel Done"
        
        .Cell(0, 20).Text = "Manual Preparation"
        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 130
            .Cell(0, i).FontBold = True
            

            
        Next

        .Column(4).Alignment = cellCenterCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellCenterCenter
        .Column(7).Alignment = cellCenterCenter

        
         .Column(16).Width = 0
         .Column(17).Width = 0
        .Column(18).Width = 0
        .Column(19).Width = 0
    
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   


End Sub




Private Sub SetRecipeQCGrid(ByVal Grd As Grid)
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
        
        
        .Cols = 15
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11
        
        .DefaultFont.Name = "Calibri"
        

        .Cell(0, 1).Text = "Line"
        .Cell(0, 2).Text = "MRCode"
        .Cell(0, 3).Text = "Description"
        .Cell(0, 4).Text = "# Prep. Week"
        .Cell(0, 5).Text = "Planned Prep."
        .Cell(0, 6).Text = "Data Recipe"
        .Cell(0, 7).Text = "QC Status"
        .Cell(0, 8).Text = "QC Date"
        .Cell(0, 9).Text = "FileName"
        .Cell(0, 12).Text = "ID"
        .Cell(0, 13).Text = "Operator"
        .Cell(0, 14).Text = "Note"
        
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







Public Sub SetHannaCodeGrid(ByVal Grd As Grid)
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
        
        
        .Cols = 12
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        

        .Cell(0, 1).Text = "HannaCode"
        .Cell(0, 2).Text = "Description"
        .Cell(0, 3).Text = "MRCode"
        .Cell(0, 4).Text = "Hanna Parameter"
        
        .Cell(0, 5).Text = "FW Hanna Parameter"
        .Cell(0, 6).Text = "STD Volume (ml)"
        .Cell(0, 7).Text = "STD Matrix"
        .Cell(0, 8).Text = "STD Exp (days)"
        .Cell(0, 9).Text = "STD Note"
        .Cell(0, 10).Text = "STD Storage"
        .Cell(0, 11).Text = "ID"
      
        
        
        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 130
            .Cell(0, i).FontBold = True
            

            
        Next
       
        .Column(4).Alignment = cellCenterCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellCenterCenter
        .Column(7).Alignment = cellCenterCenter

        
         .Column(11).Width = 0
        
    
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   


End Sub



