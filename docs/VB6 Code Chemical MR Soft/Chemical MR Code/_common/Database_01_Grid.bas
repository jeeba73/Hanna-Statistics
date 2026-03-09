Attribute VB_Name = "Database_01_Grid"
Option Explicit



Public Sub SetDatabaseSTDPreparationTable(ByVal Grd As Grid)

Dim i As Integer
        '------------------------------------------------
        '
        '      Set STDPreparation Table
        '
        '------------------------------------------------
    With Grd
      
      
        
        .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        
        
        .Cols = 18
        
        
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        .RowHeight(0) = 50
        
        
        .Cell(0, 1).Text = "Line"
        .Cell(0, 2).Text = "Prod. Date"
        .Cell(0, 3).Text = "Prod. Week"
        .Cell(0, 4).Text = "Hanna Code"
        .Cell(0, 5).Text = "Date Recipe"
        .Cell(0, 6).Text = "Recipe"
        .Cell(0, 7).Text = "Mix"
        .Cell(0, 8).Text = "Pl. Ref."
        
        .Cell(0, 9).Text = "Preparation Date"
        .Cell(0, 10).Text = "Preparation Week"
        .Cell(0, 11).Text = "# Prep. Week"
        .Cell(0, 12).Text = "Note"
        .Cell(0, 13).Text = "FileName"
        .Cell(0, 14).Text = "ID"
        .Cell(0, 15).Text = ""
        .Cell(0, 16).Text = "Close Date"
        .Cell(0, 17).Text = "Excel Done"
        

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellRightCenter
            .Column(i).Width = 120
            .Cell(0, i).FontBold = True
            
        Next
        .Column(2).Alignment = cellCenterCenter
        .Column(3).Alignment = cellCenterCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(1).Alignment = cellLeftCenter
        .Column(4).Alignment = cellLeftCenter
        .Column(6).Alignment = cellLeftCenter
        .Column(12).Alignment = cellLeftCenter
        .Column(11).Alignment = cellCenterCenter
         .Column(16).Alignment = cellCenterCenter
        
        .Column(4).Width = 200
        .Column(6).Width = 200
        .Column(9).Width = 130
        .Column(10).Width = 130
        .Column(11).Width = 130
        .Column(7).Width = 0
        .Column(8).Width = 0
        .Column(13).Width = 0
        .Column(14).Width = 0
        .Column(15).Width = 0
        
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
  

End Sub


Public Sub SetDatabaseSTDPreparationHistory(ByVal Grd As Grid)
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

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 150
            .Cell(0, i).FontBold = True
        Next
        .Column(10).Width = 0
        .Column(11).Width = 0
        .Column(2).Alignment = cellRightCenter
        .Column(3).Alignment = cellCenterCenter
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   

End Sub


Public Sub SetDatabasePreparationTable(ByVal Grd As Grid)

Dim i As Integer
        '------------------------------------------------
        '
        '      Set STDPreparation Table
        '
        '------------------------------------------------
    With Grd
      
      
        
        .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        
        
        .Cols = 20
        
        
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        .RowHeight(0) = 50
        
        .Cell(0, 1).Text = "Line"
        .Cell(0, 2).Text = "Hanna Code"
        .Cell(0, 3).Text = "Recipe"
        .Cell(0, 4).Text = "# Lot"
        .Cell(0, 5).Text = "Prep. Operator"
        .Cell(0, 6).Text = "Prep. Date"
        .Cell(0, 7).Text = "Prep. Week"
        .Cell(0, 8).Text = "# Prep. Week"
        
        .Cell(0, 9).Text = "QC Operator"
        .Cell(0, 10).Text = "Correction"
        .Cell(0, 11).Text = "Correction Date"
        .Cell(0, 12).Text = "Note"
        .Cell(0, 13).Text = "FileName"
        .Cell(0, 14).Text = "ID"
        .Cell(0, 15).Text = ""
        .Cell(0, 16).Text = "IsMix"
        .Cell(0, 17).Text = "Exp Date"
        .Cell(0, 18).Text = "Close Date"
        .Cell(0, 19).Text = "Excel Done"
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 120
            .Cell(0, i).FontBold = True
            
        Next

        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellCenterCenter
        .Column(8).Alignment = cellCenterCenter
        .Column(17).Alignment = cellCenterCenter
        .Column(18).Alignment = cellCenterCenter
        .Column(19).Alignment = cellCenterCenter
       
        .Column(13).Width = 0
        .Column(14).Width = 0
        .Column(15).Width = 0
        .Column(16).Width = 0
        
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
  

End Sub

