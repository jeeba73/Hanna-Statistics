Attribute VB_Name = "gPreparation_01_Grid"

Option Explicit

Private Grid() As Grid

Public Function SetAllPreparationGrid(ByVal Grid As Variant) As Boolean



' 0 Component


    
    Call SetPreparationComponentGrid(Grid(0))

' 1 Acquisition


    
    Call SetPreparationAcquisition(Grid(1))
    
    Call SetPreparationHannaCode(Grid(2))
    
    Call SetMaterialReqGrid(Grid(3), True)


End Function


Public Sub SetPreparationComponentGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 2

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
        
        
        .Cell(0, 2).Text = "Code"
        .Cell(0, 3).Text = "Description"
        .Cell(0, 4).Text = "CAS"
        .Cell(0, 5).Text = "%"
        .Cell(0, 6).Text = "Theor. Weight (g)"
        
        
        .Cell(0, 7).Text = "Real Weight (g)"
        .Cell(0, 8).Text = "Variance (g)"
        .Cell(0, 9).Text = "Variance %"
        .Cell(0, 10).Text = ""
        
        .Cell(0, 11).Text = "%"
        
        .Cell(0, 12).Text = "Note"
        .Cell(0, 13).Text = "bMix"
        .Cell(0, 14).Text = "Critical RM"

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 130
            .Cell(0, i).FontBold = True
            
        Next
        .Column(1).Width = 20
        .Column(3).Width = 250
        .Column(4).Width = 80
        .Column(5).Width = 80
       
        .Column(10).Width = 10
        .Column(12).Width = 200
        .Column(13).Width = 0
        .Column(5).Alignment = cellCenterCenter
        .Column(11).Alignment = cellCenterCenter
        .Column(6).Alignment = cellRightCenter
        .Column(7).Alignment = cellRightCenter
        .Column(8).Alignment = cellRightCenter
        .Column(9).Alignment = cellRightCenter
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With

End Sub
Public Sub SetPreparationAcquisition(ByVal Grd As Grid)
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
        
        
        .Cols = 17
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Code"
        .Cell(0, 2).Text = "Description"
        .Cell(0, 3).Text = "CAS"
        .Cell(0, 4).Text = "Real Weight (g)"
        .Cell(0, 5).Text = "Manufacturer"
        .Cell(0, 6).Text = "Manufacturer Code"
        .Cell(0, 7).Text = "Manufacturer Lot"
        .Cell(0, 8).Text = "Delivery Date"
        .Cell(0, 9).Text = "Qty Delivered"
        .Cell(0, 10).Text = "Week Delivery"
        .Cell(0, 11).Text = "Package"
        
        .Cell(0, 12).Text = "Note"
        .Cell(0, 13).Text = "Operator"
        .Cell(0, 14).Text = "Acquisition Time"
        .Cell(0, 15).Text = "ID"
        .Cell(0, 16).Text = "ExpDate"

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 100
            .Cell(0, i).FontBold = True
            
        Next
        
        .Column(2).Width = 250
        .Column(3).Width = 80
       
        .Column(4).Width = 150
        .Column(5).Width = 150
        .Column(6).Width = 150
        .Column(7).Width = 150
        .Column(8).Width = 150
        .Column(9).Width = 150
        
        .Column(12).Width = 200
        .Column(13).Width = 150
        .Column(14).Width = 150
        .Column(15).Width = 0
        .Column(16).Width = 150
        
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellCenterCenter
        .Column(8).Alignment = cellCenterCenter
        .Column(9).Alignment = cellCenterCenter
        .Column(10).Alignment = cellCenterCenter
        .Column(11).Alignment = cellCenterCenter
        
        
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With

End Sub


Public Sub SetPreparationHannaCode(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click

        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 10
        
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
        .Cell(0, 7).Text = "Lot Number"
        .Cell(0, 8).Text = "ID"

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
        .Column(8).Width = 0
        .Column(4).Alignment = cellRightCenter
        '.Column(8).Alignment = cellCenterCenter
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   

End Sub
