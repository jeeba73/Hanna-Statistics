Attribute VB_Name = "CLP_Main_01_Grid"
Option Explicit

Private Grid() As Grid

Public Function SetAllMainCLPGrid(ByVal Grid As Variant) As Boolean



' 0 Chemical

    Call SetCLPGridChemicalRM(Grid(0))


' 0 Recipe
    Call SetCLPCodeGrid(Grid(1))

' 0 Hanna Code
    
   Call SetCLPHannaCodeGrid(Grid(2))



End Function

Private Function SetCLPGridChemicalRM(ByVal Grd1 As Grid) As Boolean
Dim i As Integer

       '------------------------------------------------
        '       SET TABELLA CHEMICAL RM
        '------------------------------------------------
    With Grd1
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        
        
        .Cols = 12
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
       
        .Cell(0, 1).Text = "Code"
        .Column(1).Width = 120
        .Cell(0, 2).Text = "Description"
        .Column(2).Width = 250
        
        .Cell(0, 3).Text = "Cas"
        .Column(3).Width = 100
        .Cell(0, 4).Text = "Chemical Reaction Liquid"
        .Column(4).Width = 100
        .Cell(0, 5).Text = "Manufacturer Name"
        .Column(5).Width = 100
        .Cell(0, 6).Text = "Manufacturer Code"
        .Column(6).Width = 100
        .Cell(0, 7).Text = "Location"
        .Column(7).Width = 100
        .Cell(0, 8).Text = "Specified Location"
        .Column(8).Width = 100
        .Cell(0, 9).Text = "bMix"
        .Column(9).Width = 0
        .Cell(0, 10).Text = ""
        .Column(10).Width = 10
        .Cell(0, 11).Text = "ID"
        .Column(11).Width = 0

       For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 150
            .Cell(0, i).FontBold = True
        Next
        
          .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
   
 
End Function
Private Sub SetCLPCodeGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '       TABELLA RECIPE
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        
        
        .Cols = 7
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Code"
        .Cell(0, 2).Text = "Description"
        .Cell(0, 3).Text = "Line"
        .Cell(0, 4).Text = "Mix"
        .Cell(0, 5).Text = "ID"
       .Cell(0, 6).Text = ""
     

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 150
            .Cell(0, i).FontBold = True
            
        Next
        
        .Column(5).Width = 0
        .Column(6).Width = 10
        .Column(2).Width = 250
     
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
 
End Sub

Private Sub SetCLPHannaCodeGrid(ByVal Grd As Grid)


Dim i As Integer
        '------------------------------------------------
        '      Hanna Code CLP Grid
        '------------------------------------------------
    With Grd
      
      
      .Rows = 3

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        
        
        .Cols = 8
        .RowHeight(0) = 35
        
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Hanna Code"
        .Cell(0, 2).Text = "Product Name"
        .Cell(0, 3).Text = "Line"
        .Cell(0, 4).Text = "Recipe"
        .Cell(0, 5).Text = "Mix #1"
        .Cell(0, 6).Text = "Mix #2"
        .Cell(0, 7).Text = "ID"
     
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 130
            .Cell(0, i).FontBold = True
        Next
        
        .Column(7).Width = 0
        .Column(2).Width = 250

        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
End Sub
