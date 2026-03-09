Attribute VB_Name = "hProduction_01_Grid"
Option Explicit

Private Grid() As Grid

Public Function SetAllSTDPreparationGrid(ByVal Grid As Variant) As Boolean


' 0 SetSTDPreparationCodeGrid
     Call SetSTDPreparationCodeGrid(Grid(0))
' 1 Acquisition
    Call SetSTDPreparationHistory(Grid(1))
    
  ' Call SetPreparationHannaCode(Grid(2))

End Function


Public Sub SetSTDPreparationCodeGrid(ByVal Grd As Grid)
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
        
        
        .Cols = 12
        
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
        .Cell(0, 7).Text = "Q.ty produced"
        .Cell(0, 8).Text = ""
        
        .Cell(0, 9).Text = "%"
        .Cell(0, 10).Text = "Recipe"
        .Cell(0, 11).Text = "Mix"

        

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 150
            .Cell(0, i).FontBold = True
            If i > 7 Then .Column(i).Width = 100
        Next
        .Column(1).Width = 150
        .Column(2).Width = 200
        .Column(3).Width = 120
        .Column(4).Width = 120
        .Column(5).Width = 40
        .Column(7).Width = 150
        .Column(8).Width = 8
        .Column(9).Width = 90

        
        .Column(5).Alignment = cellRightCenter
        .Column(6).Alignment = cellRightCenter
        .Column(7).Alignment = cellRightCenter
        .Column(9).Alignment = cellCenterCenter
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   

End Sub



Public Sub SetSTDPreparationHistory(ByVal Grd As Grid)
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
        
        
        .Cols = 15
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Code"
        .Cell(0, 2).Text = "Q.ty Produced"
        .Cell(0, 3).Text = "Lot Number"
        .Cell(0, 4).Text = "Operator"
        .Cell(0, 5).Text = "DateProd"
        .Cell(0, 6).Text = "WeekProd"
        .Cell(0, 7).Text = "Machine"
        .Cell(0, 8).Text = "Note"
        .Cell(0, 9).Text = "Acquisition Time"
        .Cell(0, 10).Text = "ID"
        .Cell(0, 11).Text = "Index"
        
        
        .Cell(0, 12).Text = "Mix 1 Lot"
        .Cell(0, 13).Text = "Mix 2 Lot"
        
        .Cell(0, 14).Text = "Exp Date"
        

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 150
            .Cell(0, i).FontBold = True
        Next
        .Column(10).Width = 0
        .Column(11).Width = 0
        .Column(2).Alignment = cellRightCenter
        .Column(3).Alignment = cellCenterCenter
        .Column(14).Alignment = cellCenterCenter
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   

End Sub

