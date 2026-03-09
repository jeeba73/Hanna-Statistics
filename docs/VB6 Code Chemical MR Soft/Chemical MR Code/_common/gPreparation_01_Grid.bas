Attribute VB_Name = "gPreparation_01_Grid"

Option Explicit

Private Grid() As Grid

Public Function SetAllPreparationGrid(ByVal Grid As Variant, Optional bManual As Boolean) As Boolean


    If bManual Then
        Call SetManualPreparationComponentGrid(Grid(0))
    Else
    
        Call SetPreparationComponentGrid(Grid(0))
    End If
    
    
    Call SetPreparationAcquisition(Grid(1), bManual)
    
    Call SetStockTable(Grid(2))
    
    Call SetMotherSolutionTable(Grid(4))
    
    Call SetGridPipette(Grid(5))
    
    
    
    With Grid(5)
        .Column(5).Width = 150
        .Column(6).Width = 150
        .Column(7).Width = 50
    End With
    
    
   
    

End Function


Public Sub SetPreparationComponentGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForProduction  STD
        '------------------------------------------------
    With Grd
      
      
      .Rows = 2

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
    
        
        .Cols = 11
        
        .RowHeight(0) = 45
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
      '  .DefaultFont.Size = 16
        
        .Cell(0, 1).Text = "STD Number"
        .Cell(0, 2).Text = "STD Value"
        .Cell(0, 4).Text = "MR Qty"
        .Cell(0, 3).Text = "(Unit)"
        
        
        .Cell(0, 5).Text = "MR Acquired"
        .Cell(0, 6).Text = "Variance"
        .Cell(0, 7).Text = "Variance %"
        .Cell(0, 8).Text = ""
       
        .Cell(0, 9).Text = "Note"
        
        .Cell(0, 10).Text = "STD_ID"


        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 130
            .Cell(0, i).FontBold = True
            
        Next

        .Column(1).Alignment = cellCenterCenter
        .Column(2).Alignment = cellCenterCenter
        .Column(3).Alignment = cellLeftCenter
        .Column(4).Alignment = cellCenterCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellCenterCenter
        .Column(8).Alignment = cellCenterCenter
        .Column(10).Alignment = cellCenterCenter
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With

End Sub
Public Sub SetManualPreparationComponentGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForProduction  STD
        '------------------------------------------------
    With Grd
      
      
      .Rows = 2

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
    
        
        .Cols = 11
        
        .RowHeight(0) = 45
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
      '  .DefaultFont.Size = 16
        
        .Cell(0, 1).Text = "MR Code"
        .Cell(0, 2).Text = "STD Value"
        .Cell(0, 4).Text = "MR Qty"
        .Cell(0, 3).Text = "(Unit)"
        
        
        .Cell(0, 5).Text = "MR Acquired"
        .Cell(0, 6).Text = "Variance"
        .Cell(0, 7).Text = "Variance %"
        .Cell(0, 8).Text = ""
       
        .Cell(0, 9).Text = "Note"
        
        .Cell(0, 10).Text = ""


        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 130
            .Cell(0, i).FontBold = True
            
        Next
        
        .Column(2).Width = 0
        .Column(10).Width = 0

        .Column(1).Alignment = cellCenterCenter
        .Column(2).Alignment = cellCenterCenter
        .Column(3).Alignment = cellLeftCenter
        .Column(4).Alignment = cellCenterCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellCenterCenter
        .Column(8).Alignment = cellCenterCenter
        .Column(10).Alignment = cellCenterCenter
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With

End Sub


Public Sub GetInitPreparationSTDGrid(ByVal Grid As Grid, ByRef uPreparation As RecipeForProduction)
Dim i As Integer
Dim t As Integer
Dim x As Integer
Dim RecipeCount As Integer
Dim Variance As Double
Dim VariancePerc As Double
Dim TotalRealWeight As Double
Dim TheorWeight As Double
Dim bUmMassa As Boolean
Dim Density As Double
Dim bRecalculate As Boolean
Dim PesoIntolleranza As Double
Dim MyColor As OLE_COLOR
Dim InputWeight As String
Dim bCorrection As Boolean
On Error GoTo ERR_GET:


        '------------------------------------------------
        '      Preparation Tella STDs
        '------------------------------------------------
        
    With Grid
      
      
      .Rows = 1

        .AutoRedraw = False
        .ReadOnly = False
        

    
       
        Call SetSTDTheoreticalWeight(uPreparation.MsType, uPreparation)
        
        
        With uPreparation.Recipe
          
            For i = 1 To .STDcount
         
                     
        '.Cell(0, 1).Text = "Number"
        '.Cell(0, 2).Text = "Value"
        '.Cell(0, 3).Text = "MR Qty"
        '.Cell(0, 4).Text = "Unit"
        '.Cell(0, 5).Text = "Real Weight"
        '.Cell(0, 6).Text = "Variance"
        '.Cell(0, 7).Text = "Variance %"
        '.Cell(0, 8).Text = ""
        '.Cell(0, 9).Text = "Note"
        
            
        
        
        
        

                Grid.AddItem "", False
                Grid.Cell(Grid.Rows - 1, 1).Text = .STD(i).NUMBER
                Grid.Cell(Grid.Rows - 1, 2).Text = .STD(i).Value
                Grid.Cell(Grid.Rows - 1, 4).Text = PadString(.STD(i).TheoreticalWeight)
                
                
                TotalRealWeight = TotalRealWeight + .STD(i).TheoreticalWeight
                     
                Grid.Cell(Grid.Rows - 1, 3).Text = uPreparation.Recipe.STDUnit
                

                Grid.Cell(Grid.Rows - 1, 5).Text = PadString(.STD(i).ActualWeight)

                Grid.Cell(Grid.Rows - 1, 4).FontBold = True

                
                Grid.Cell(Grid.Rows - 1, 4).BackColor = vbColorResults
                
                Grid.Cell(Grid.Rows - 1, 3).Alignment = cellRightCenter
                Grid.Cell(Grid.Rows - 1, 5).Alignment = cellCenterCenter
                Grid.Cell(Grid.Rows - 1, 6).Alignment = cellCenterCenter
                Grid.Cell(Grid.Rows - 1, 7).Alignment = cellCenterCenter
                
                Grid.Cell(Grid.Rows - 1, 5).BackColor = vbColorResults
                Grid.Cell(Grid.Rows - 1, 6).BackColor = vbColorResults
                Grid.Cell(Grid.Rows - 1, 7).BackColor = vbColorResults
            
    
               
cont:
            Next

            .TotalWeight = TotalRealWeight

            .STDcount = Grid.Rows - 1
    
                
        End With
        .ReadOnly = True
       .Column(4).AutoFit
       
ERR_END:
        .Refresh
        .AutoRedraw = True
    End With
   On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox Err.Description
    Resume Next

End Sub





Public Sub SetPreparationAcquisition(ByVal Grd As Grid, Optional ByVal bManual As Boolean)
Dim i As Integer
        '------------------------------------------------
        '      Grid2 Acquisition
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 20
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Bottle"
        .Cell(0, 2).Text = "Lot"

        .Cell(0, 3).Text = "STDNumber"
        .Cell(0, 4).Text = "STDValue"
        .Cell(0, 5).Text = "STDQty"
        .Cell(0, 6).Text = "STDUnit"
        .Cell(0, 7).Text = "MR Acquired"
        
        .Cell(0, 8).Text = "Note"
        .Cell(0, 9).Text = "Operator"
        .Cell(0, 10).Text = "Acquisition Time"
        .Cell(0, 11).Text = "ID"
        .Cell(0, 12).Text = "Left In bottle"
        .Cell(0, 13).Text = "Pipette Code"
        .Cell(0, 14).Text = "Pipette Type"
        
        
        .Cell(0, 15).Text = "Scale Code"
        .Cell(0, 16).Text = "Glassware Code"
        
        
        .Cell(0, 17).Text = "MotherSolution Date"
        
        
        .Cell(0, 18).Text = "MNP"
        .Cell(0, 19).Text = "Exp.MR"
        

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 100
            .Cell(0, i).FontBold = True
            
        Next
        
        
        .Column(4).Width = 150
        .Column(5).Width = 150
        .Column(6).Width = 150
        .Column(7).Width = 150
        .Column(8).Width = 150
        .Column(9).Width = 150
        
        
        .Column(2).Alignment = cellCenterCenter
        .Column(3).Alignment = cellCenterCenter
        .Column(4).Alignment = cellCenterCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellCenterCenter
        
        If bManual Then
            .Column(17).Width = 0
            .Cell(0, 3).Text = "MR Code"
        End If
        
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With

End Sub


Public Sub SetMotherSolutionTable(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '     Mother Solution
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
        
        
        .Cell(0, 1).Text = "DataMS"
        .Cell(0, 2).Text = "QtyProduced"

        .Cell(0, 3).Text = "DataExp"
        .Cell(0, 4).Text = "MRCode"
        .Cell(0, 5).Text = "Operator"
        .Cell(0, 6).Text = "QtyLeft"
        
        .Cell(0, 7).Text = "Note"
        
        .Cell(0, 8).Text = "Bottle Number"
        .Cell(0, 9).Text = "BottleLot"
        .Cell(0, 10).Text = "Bottle Qty"
        .Cell(0, 11).Text = "ID"
        

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
            .Column(i).Width = 100
            .Cell(0, i).FontBold = True
            
        Next
      
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With

End Sub
