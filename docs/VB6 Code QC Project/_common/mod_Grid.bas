Attribute VB_Name = "mod_Grid"
Option Explicit

Public Function SetGridtest(ByRef Grd2 As Grid)

       '------------------------------------------------
        '       SET TABELLA Test
        '------------------------------------------------
    With Grd2
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = True ' False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        .DefaultFont.Size = 12 '* m_ControlGridFontSize
        .DefaultRowHeight = 40
        
        .Cols = 38
        
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
        .Cell(0, 11).Text = "METER 1"
        .Column(11).Width = 170
        .Cell(0, 12).Text = "METER 2"
        .Column(12).Width = 170
        .Cell(0, 13).Text = "METER 3"
        .Column(13).Width = 170
         .Cell(0, 14).Text = "METER 4"
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
        
        .Cell(0, 23).Text = "CORRECTION"
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
              
        .Cell(0, 36).Text = "STD_ID"
        .Column(36).Width = 200
                                           
        .Cell(0, 37).Text = "NOTE"
        .Column(37).Width = 300
                                           
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .DefaultFont.Bold = False
        
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
            .Column(i).Width = 50
        Next
        .Column(0).Width = 0
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
   
 
End Function

Public Function SetGridCode(ByRef Grd1 As Grid) As Boolean


       '------------------------------------------------
        '       SET TABELLA Codici 1
        '------------------------------------------------
    With Grd1
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        .Cols = 8
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Code SFG"
        .Column(1).Width = 150
        .Cell(0, 2).Text = "Description"
        .Column(2).Width = 200
        
        .Cell(0, 3).Text = "Line"
        .Column(3).Width = 100
        
        .Cell(0, 4).Text = "Recipe"
        .Column(4).Width = 100
        
               
        .Cell(0, 5).Text = "Range Min"
        .Column(5).Width = 100
                     
        .Cell(0, 6).Text = "Range Max"
        .Column(6).Width = 100
                     
        .Cell(0, 7).Text = "ID"
        .Column(7).Width = 0

        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .DefaultFont.Size = 12 ' * m_ControlGridFontSize
        .DefaultFont.Bold = False
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
   
 
End Function

Public Function SetGridEditCode(ByRef Grd As Grid) As Boolean
   
    With Grd
        .Rows = 1
        .AutoRedraw = False
        .ReadOnly = False
        .DisplayFocusRect = True 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionFree
        .DefaultFont.Size = 12 '* m_ControlGridFontSize
        
        .DefaultFont.Bold = True
        .DefaultRowHeight = 35
        .ExtendLastCol = True
        .Cols = 3
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Column(1).Width = 250
        .Column(2).Width = 250
        .Rows = 76

        .Cell(1, 1).Text = "  " & "Hanna SFG Code"
        .Cell(2, 1).Text = "  " & "SFG Description"
        .Cell(3, 1).Text = "  " & "Line"
        .Cell(4, 1).Text = "  " & "Recipe"
        
        .Cell(5, 1).Text = "  " & "QC Method"
        .Cell(6, 1).Text = "  " & "Meter Family 1"
        .Cell(7, 1).Text = "  " & "Meter Family 2"
        
        .Cell(8, 1).Text = "  " & "Parameter Method"
        .Cell(9, 1).Text = "  " & "Parameter Formula"
        .Cell(10, 1).Text = "  " & "Measurement Unit"
        
        
        .Range(11, 1, 11, 2).Merge
        .Cell(11, 1).Text = "User manual parameter data"
        .Cell(12, 1).Text = "  " & "Range Min"
        .Cell(13, 1).Text = "  " & "Range Max"
        .Cell(14, 1).Text = "  " & "Decimal"
        
        .Range(15, 1, 15, 2).Merge
        .Cell(15, 1).Text = "Tolerance"
        .Cell(16, 1).Text = "  " & "Fixed"
        .Cell(17, 1).Text = "  " & "And / Or"
        .Cell(18, 1).Text = "  " & "Percentage (%)"
        .Cell(19, 1).Text = "  " & "QC Restriction (%)"
       ' .Cell(20, 1).Text = "  " & "STD MR"
        
        .Range(21, 1, 21, 2).Merge
        .Cell(21, 1).Text = "STD1"
        .Cell(22, 1).Text = "  " & "Value"
        .Cell(23, 1).Text = "  " & "Min"
        .Cell(24, 1).Text = "  " & "Max"
        
        .Range(25, 1, 25, 2).Merge
        .Cell(25, 1).Text = "STD2"
        .Cell(26, 1).Text = "  " & "Value"
        .Cell(27, 1).Text = "  " & "Min"
        .Cell(28, 1).Text = "  " & "Max"
        
        .Range(29, 1, 29, 2).Merge
        .Cell(29, 1).Text = "STD3"
        .Cell(30, 1).Text = "  " & "Value"
        .Cell(31, 1).Text = "  " & "Min"
        .Cell(32, 1).Text = "  " & "Max"
        
        .Range(33, 1, 33, 2).Merge
        .Cell(33, 1).Text = "STD4"
        .Cell(34, 1).Text = "  " & "Value"
        .Cell(35, 1).Text = "  " & "Min"
        .Cell(36, 1).Text = "  " & "Max"
        
         .Range(37, 1, 37, 2).Merge
        .Cell(37, 1).Text = "STD5"
        .Cell(38, 1).Text = "  " & "Value"
        .Cell(39, 1).Text = "  " & "Min"
        .Cell(40, 1).Text = "  " & "Max"
        
         .Range(41, 1, 41, 2).Merge
        .Cell(41, 1).Text = "STD6"
        .Cell(42, 1).Text = "  " & "Value"
        .Cell(43, 1).Text = "  " & "Min"
        .Cell(44, 1).Text = "  " & "Max"
                
                
        .Range(45, 1, 45, 2).Merge
        .Cell(45, 1).Text = "pH 1"
        .Cell(46, 1).Text = "  " & "Value"
        .Cell(47, 1).Text = "  " & "Min"
        .Cell(48, 1).Text = "  " & "Max"
                
                
        .Range(49, 1, 49, 2).Merge
        .Cell(49, 1).Text = "pH 2"
        .Cell(50, 1).Text = "  " & "Value"
        .Cell(51, 1).Text = "  " & "Min"
        .Cell(52, 1).Text = "  " & "Max"
                
        .Range(53, 1, 53, 2).Merge
        .Cell(53, 1).Text = "pH 3"
        .Cell(54, 1).Text = "  " & "Value"
        .Cell(55, 1).Text = "  " & "Min"
        .Cell(56, 1).Text = "  " & "Max"
                             
                             
        .Range(57, 1, 57, 2).Merge
        .Cell(57, 1).Text = "Weight (mg)"
        .Cell(58, 1).Text = "  " & "Value"
        .Cell(59, 1).Text = "  " & "Min"
        .Cell(60, 1).Text = "  " & "Max"
                
       ' .Cell(61, 1).Text = "  " & "Certified"
        .Cell(62, 1).Text = "  " & "Revision Date"
        
        
        .Cell(63, 1).Text = "  " & "MR1"
        .Cell(64, 1).Text = "  " & "MR2"
        '.Cell(65, 1).Text = "  " & "MS1 Value"
        '.Cell(66, 1).Text = "  " & "MS1 Volume (ml)"
        '.Cell(67, 1).Text = "  " & "MS2 Conc"
        '.Cell(68, 1).Text = "  " & "MS2 Volume (ml)"
           
        '.Cell(69, 1).Text = "  " & "MS EXP (days)"
        '.Cell(70, 1).Text = "  " & "STD Matrix"
        '.Cell(71, 1).Text = "  " & "STD Volume (ml)"
        '.Cell(72, 1).Text = "  " & "STD EXP (days)"
        '.Cell(73, 1).Text = "  " & "STD Note"
        '.Cell(74, 1).Text = "  " & "FW Parameter Formula"

        '.Cell(75, 1).Text = "  " & "STD Storage"




        
        .RowHeight(61) = 0
        Dim i As Integer
        For i = 65 To .Rows - 1
            .RowHeight(i) = 0
            
        Next
        
        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            
        Next
        
        
        For i = 1 To .Rows - 1
            .Cell(i, 1).BackColor = vbColorUnabled
            .Cell(i, 1).ForeColor = vbTimBlue 'vbColorDarkFont 'vbColorForeFixed  ' vbColorTextDarkBlue
            .Cell(i, 1).FontBold = True
            .Cell(i, 1).Locked = True
            .Cell(i, 2).ForeColor = vbColorDarkFont
            If i = 11 Or i = 15 Or i = 21 Or i = 25 Or i = 29 Or i = 33 Or i = 37 Or i = 41 Or i = 45 Or i = 49 Or i = 53 Or i = 57 Then
                .Cell(i, 1).Alignment = cellCenterCenter
                .Cell(i, 1).BackColor = vbColorUnabled ' vbColorTextBlue ' &HF0F0F0
                .Cell(i, 1).ForeColor = vbTimBlue ' vbColorDarkFont 'vbWhite  ' &HF0F0F0
            End If
            
        Next


        
        
        .ReadOnly = False
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
        
        .DefaultRowHeight = 25
        .Cols = 26
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
        .Cell(0, 20).Text = ""
        .Column(20).Width = 10
        
         .Cell(0, 21).Text = "Excel Done"
         .Column(21).Width = 100
         
         
         .Cell(0, 22).Text = "REAGENT SET 1"
         .Column(22).Width = 100
         .Cell(0, 23).Text = "REAGENT SET 1 CODE "
         .Column(23).Width = 100
         
         .Cell(0, 24).Text = "REAGENT SET 2"
         .Column(24).Width = 100
         .Cell(0, 25).Text = "REAGENT SET 2 CODE"
         .Column(25).Width = 100
         
        ' .Cell(0, 22).Text = "Excel Scheduled"
        ' .Column(22).Width = 100
        ' .Column(22).CellType = cellCheckBox
         
         
                  
        
        
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
           
        Next
        .DefaultFont.Size = 12 ' * m_ControlGridFontSize
        .DefaultFont.Bold = True
        
        .ReadOnly = True
        
        .AutoRedraw = True
        .Refresh
        
    End With
    
    If Grd2 Is Nothing Then Exit Function
    
    With Grd2
      .Rows = 1

        .AutoRedraw = False
        .ReadOnly = True
        .DisplayFocusRect = True 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        .DefaultFont.Size = 12 * m_ControlGridFontSize
        .DefaultFont.Bold = False
        .DefaultRowHeight = 35
        
        .Cols = 9
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Hanna Code SFG"
        .Column(1).Width = 190
        .Cell(0, 2).Text = "SFG Description"
        .Column(2).Width = 250
        
        .Cell(0, 3).Text = "Line"
        .Column(3).Width = 150
        .Cell(0, 4).Text = "Recipe"
        .Column(4).Width = 150
        
        .Cell(0, 5).Text = "Reference Weight"
        .Column(5).Width = 170
       
        .Cell(0, 6).Text = "Parameter / Method" ' quanti test ho fatto
        .Column(6).Width = 250
       
     
         .Cell(0, 7).Text = "ID"
        .Column(7).Width = 0
           
         .Cell(0, 8).Text = "FileName"
        .Column(8).Width = 0
                           


        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
            
        Next
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With


    Dim t As Integer
    With Grd1
        .Rows = 4
        .Cell(1, 1).Text = "#2"
        .Cell(1, 2).Text = "0.2"
        .Cell(2, 1).Text = "#3"
        .Cell(2, 2).Text = "0.55"
        .Cell(3, 1).Text = "#4"
        .Cell(3, 2).Text = "0.75"
        For i = 0 To .Rows - 1
            For t = 0 To .Cols - 1
                .Cell(i, t).ForeColor = vbColorDarkFont
                .Cell(i, t).Alignment = cellCenterCenter
                
            Next
        Next
    End With
    
    With Grd2
        .Rows = 4
        .Cell(1, 1).Text = "HI93709B-0"
        .Cell(1, 2).Text = "Manganese HR Reagent B"
        .Cell(1, 3).Text = "L57 Powder"
        .Cell(1, 4).Text = "R034"
        
        .Cell(1, 5).Text = "500"
        .Cell(1, 6).Text = "Manganese High Range (Mn)"

        .Cell(2, 1).Text = "HI93709B-0"
        .Cell(2, 2).Text = "Manganese HR Reagent B"
        .Cell(2, 3).Text = "L57 Powder"
        .Cell(2, 4).Text = "R034"
        
        .Cell(2, 5).Text = "500"
        .Cell(2, 6).Text = "Manganese High Range (Mn)"
        
        .Cell(3, 1).Text = "HI93709B-0"
        .Cell(3, 2).Text = "Manganese HR Reagent B"
        .Cell(3, 3).Text = "L57 Powder"
        .Cell(3, 4).Text = "R034"
        
        .Cell(3, 5).Text = "500"
        .Cell(3, 6).Text = "Manganese High Range (Mn)"


        For i = 0 To .Rows - 1
            For t = 0 To .Cols - 1
                .Cell(i, t).ForeColor = vbColorDarkFont
                .Cell(i, t).Alignment = cellCenterCenter
                
            Next
        Next
    End With

End Function




Public Sub GrdRisultati(ByVal Grd As Grid, ByVal UNIT_PP As String)
    With Grd
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
       ' .DefaultFont.Size = 12 ' * m_ControlGridFontSize
       ' .DefaultFont.Bold = True
        .DefaultRowHeight = 35
        .Cols = 8
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Standard Value [" & UNIT_PP & "]"
        .Column(1).Width = 200
        .Cell(0, 2).Text = "Target Value " & Chr$(177) & " U [" & UNIT_PP & "]"
        .Column(2).Width = 200
        .Cell(0, 3).Text = "Mean Value [" & UNIT_PP & "]"
        .Column(3).Width = 200
        .Cell(0, 4).Text = "Tot Average [" & UNIT_PP & "]"
        .Column(4).Width = 200
        .Cell(0, 5).Text = "STDNumber"
        .Column(5).Width = 0
        .Cell(0, 6).Text = "STDValue"
        .Column(6).Width = 0
        .Cell(0, 7).Text = "Passed"
        .Column(7).Width = 150
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
            .Cell(0, i).BackColor = vbColorTextLightBlue
        Next
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
    
    
    Dim t As Integer
    With Grd
        .Rows = 4
        .Cell(1, 1).Text = "0.00 " & UNIT_PP & ""
        .Cell(1, 2).Text = "0.00 " & Chr$(247) & " 0.05"
        .Cell(1, 3).Text = "0.00"
        
        .Cell(2, 1).Text = "0.25 " & UNIT_PP & ""
        .Cell(2, 2).Text = "0.25 " & Chr$(177) & " 0.05"
        .Cell(2, 3).Text = "0.25"
        
        .Cell(3, 1).Text = "1.00 " & UNIT_PP & ""
        .Cell(3, 2).Text = "1.00 " & Chr$(177) & " 0.05"
        .Cell(3, 3).Text = "1.00"
        
        For i = 0 To .Rows - 1
            For t = 0 To .Cols - 1
                .Cell(i, t).ForeColor = vbColorDarkFont
                .Cell(i, t).Alignment = cellCenterCenter
                
            Next
        Next
    End With
    
    
End Sub



