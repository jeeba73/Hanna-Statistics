Attribute VB_Name = "mod_Grid"
Option Explicit

Public Function SetGridtest(ByRef Grd2 As Grid)

       '------------------------------------------------
        '       SET TABELLA Test
        '------------------------------------------------
    With Grd2
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        .DefaultFont.Size = 11 '* m_ControlGridFontSize
        .DefaultRowHeight = 40
        .Cols = 36
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Standard"
        .Column(1).Width = 120
        .Cell(0, 2).Text = "STD Value"
        .Column(2).Width = 120
        .Cell(0, 3).Text = "#"
       
        .Column(3).Width = 50
        .Cell(0, 4).Text = "TEST"
        .Column(4).Width = 150
        .Cell(0, 5).Text = "QC DATE"
        .Column(5).Width = 120
        .Cell(0, 6).Text = "QC TIME"
        .Column(6).Width = 100
        .Cell(0, 7).Text = "PROD. DATE"
        .Column(7).Width = 120
        .Cell(0, 8).Text = "PROD. TIME"
        .Column(8).Width = 120
               
        .Cell(0, 9).Text = "PROD. OPERATOR"
        .Column(9).Width = 200
        
        
        .Cell(0, 10).Text = "HEAD"
        .Column(10).Width = 80
        .Cell(0, 11).Text = "METER 1 [ppm]"
        .Column(11).Width = 170
        
        .Cell(0, 12).Text = "METER 2 [ppm]"
        .Column(12).Width = 170
        .Cell(0, 13).Text = "METER 3 [ppm]"
        .Column(13).Width = 170
         .Cell(0, 14).Text = "METER 4 [ppm]"
        .Column(14).Width = 170
        .Cell(0, 15).Text = "SPECTR. [ABS]"
        .Column(15).Width = 150
        
        .Cell(0, 16).Text = "pH 1"
        .Column(16).Width = 80
        .Cell(0, 17).Text = "pH 2"
        .Column(17).Width = 80
        .Cell(0, 18).Text = "pH 3"
        .Column(18).Width = 80
               
        
        
        .Cell(0, 19).Text = "TURB."
        .Column(19).Width = 120
        .Cell(0, 20).Text = "WEIGHT [mg]"
        .Column(20).Width = 150
        .Cell(0, 21).Text = "REAGENT SET"
        .Column(21).Width = 150
        .Cell(0, 22).Text = "QC OPERATOR"
        .Column(22).Width = 200
        
        .Cell(0, 23).Text = "NOTE"
        .Column(23).Width = 300
        
        .Cell(0, 1).BackColor = vbColorTextLightBlue
        .Cell(0, 2).BackColor = vbColorTextLightBlue
        .Cell(0, 11).BackColor = vbColorTextLightBlue
        .Cell(0, 12).BackColor = vbColorTextLightBlue
        .Cell(0, 13).BackColor = vbColorTextLightBlue
        .Cell(0, 14).BackColor = vbColorTextLightBlue
        .Cell(0, 16).BackColor = vbColorTextLightBlue
        .Cell(0, 17).BackColor = vbColorTextLightBlue
        .Cell(0, 18).BackColor = vbColorTextLightBlue
        .Cell(0, 20).BackColor = vbColorTextLightBlue
        
        .Cell(0, 24).Text = "phNumber"
        .Column(24).Width = 0
              
        .Cell(0, 25).Text = "STD"
        .Column(25).Width = 0
                          
        .Cell(0, 26).Text = "STD Value"
        .Column(26).Width = 0
        .Cell(0, 27).Text = "STD Min"
        .Column(27).Width = 0
        .Cell(0, 28).Text = "STD Max"
        .Column(28).Width = 0
        
        .Cell(0, 29).Text = "Weight Value"
        .Column(29).Width = 0
        .Cell(0, 30).Text = "Weight Min"
        .Column(30).Width = 0
        .Cell(0, 31).Text = "Weight Max"
        .Column(31).Width = 0
        
        .Cell(0, 32).Text = "Range STD Min"
        .Column(32).Width = 0
        .Cell(0, 33).Text = "Range STD Max"
        .Column(33).Width = 0
        
        .Cell(0, 34).Text = "OTHER CODE SFG"
        .Column(34).Width = 200
        .Cell(0, 35).Text = "LOT"
        .Column(35).Width = 200
              
                                           
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
End Function

Public Function SetGridStandardTolerance(ByRef Grd As Grid) As Boolean


       '------------------------------------------------
        '       SET TABELLA Codici 1
        '------------------------------------------------
    With Grd
      .Rows = 3

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        .Cols = 36
        .FixedRows = 2
        .ReadOnly = False
        .DefaultFont.Size = 12 ' * m_ControlGridFontSize
        .DefaultFont.Bold = False
        
        
        
        
        .Range(0, 1, 0, 4).Merge
        
        .Cell(0, 1).Text = "Tolerance"
        .Range(0, 5, 1, 5).Merge
        
        .Cell(1, 1).Text = "Fixed"
        .Cell(1, 2).Text = "And / Or"
        .Cell(1, 3).Text = "%"
        .Cell(1, 4).Text = "Qc Restriction"
        
        
        .Cell(0, 5).Text = "STD MR"
        .Range(0, 6, 0, 8).Merge
        .Cell(0, 6).Text = "STD1"
        

         .Range(0, 9, 0, 11).Merge
        .Cell(0, 9).Text = "STD2"
        
        .Range(0, 12, 0, 14).Merge
        .Cell(0, 12).Text = "STD3"
        
        .Range(0, 15, 0, 17).Merge
        .Cell(0, 15).Text = "STD4"
        
        .Range(0, 18, 0, 20).Merge
        .Cell(0, 18).Text = "STD5"
        
        .Range(0, 21, 0, 23).Merge
        .Cell(0, 21).Text = "STD6"
        
        .Range(0, 24, 0, 26).Merge
        .Cell(0, 24).Text = "pH 1"
        
        .Range(0, 27, 0, 29).Merge
        .Cell(0, 27).Text = "pH 2"
        
        .Range(0, 30, 0, 32).Merge
        .Cell(0, 30).Text = "pH 3"
             
             
        .Range(0, 33, 0, 35).Merge
        .Cell(0, 33).Text = "Weight"
        
        Dim i As Integer
        For i = 6 To 35 Step 3
            .Cell(1, i).Text = "Value"
            .Cell(1, i + 1).Text = "Min"
            .Cell(1, i + 2).Text = "Max"
        Next
      


        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
            .Column(i).Width = 150
        Next
        .Column(0).Width = 0
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
   
 
End Function




Public Function SetGrid(ByRef Grd1 As Grid, Optional ByRef Grd2 As Grid) As Boolean

       '------------------------------------------------
        '       SET TABELLA SEDI
        '------------------------------------------------
    With Grd1
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        .DefaultFont.Size = 11
        
        .DefaultFont.Bold = False
        .DefaultRowHeight = 25
        .Cols = 20
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Lot Number"
        .Column(1).Width = 150
        .Cell(0, 2).Text = "Code SFG"
        .Column(2).Width = 200
        .Cell(0, 3).Text = "Description"
        .Column(3).Width = 250
        .Cell(0, 4).Text = "Recipe"
        .Column(4).Width = 100
        .Cell(0, 5).Text = "Prep. Week"
        .Column(5).Width = 100
        .Cell(0, 6).Text = "Range Min"
        .Column(6).Width = 250
        .Cell(0, 7).Text = "Range Max"
        .Column(7).Width = 250
        .Cell(0, 8).Text = "Date"
        .Column(8).Width = 150
        .Cell(0, 9).Text = "Exp.Date"
        .Column(9).Width = 150
        .Cell(0, 10).Text = "# Test" ' quanti test ho fatto
        .Column(10).Width = 100
        .Cell(0, 11).Text = "Mean Value" ' se ho fatto calcolo medie
        .Column(11).Width = 0
        .Column(11).CellType = cellCheckBox
        .Cell(0, 12).Text = "Finalise" ' se ho finalizzat ( solo Laboratory Manager )
        .Column(12).Width = 120
        .Column(12).CellType = cellCheckBox
        .Cell(0, 13).Text = "User"
        .Column(13).Width = 170
        .Cell(0, 14).Text = "QC Note"
        .Column(14).Width = 300
         .Cell(0, 15).Text = "ID"
        .Column(15).Width = 0
         .Cell(0, 16).Text = "FileName"
        .Column(16).Width = 0
        .Cell(0, 17).Text = "NomeFileReport"
        .Column(17).Width = 0
        .Cell(0, 18).Text = "NomeFileExcel"
        .Column(18).Width = 0
        .Cell(0, 19).Text = "CODE_ID"
        .Column(19).Width = 0
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Cell(0, i).FontBold = True
        Next

        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
    

End Function
