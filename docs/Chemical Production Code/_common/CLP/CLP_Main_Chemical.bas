Attribute VB_Name = "CLP_Main_Chemical"
Option Explicit



Public Function AddChemicalRMinGrid(ByVal Grd As Grid, Optional ByVal RecipeCode As String)

Dim i As Integer
Dim t As Integer
  
        '.Cell(0, 0).Text = "n."
        '.Cell(0, 1).Text = "Code"
        '.Cell(0, 2).Text = "Description"
        '.Cell(0, 3).Text = "Cas"
        '.Cell(0, 4).Text = "Chemical Reaction Liquid"
        '.Cell(0, 5).Text = "Manufacturer Name"
        '.Cell(0, 6).Text = "Manufacturer Code"
        '.Cell(0, 7).Text = "Location"
        '.Cell(0, 8).Text = "Specified Location"
        '.Cell(0, 9).Text = "bMix"
        '.Cell(0, 10).Text = ""
        '.Cell(0, 11).Text = "ID"
        
        With Grd
            .Rows = 1
            .AutoRedraw = False
            '.DefaultFont.Size = 12
            '.DefaultRowHeight = 25
            With dbTabRawMaterial
                .filter = ""
                
                'If UserLine <> "" And UCase(UserLine) <> UCase("All Lines") Then
                '        .filter = ""
                'End If
                
                If .EOF Then
                
                Else
                    .MoveFirst
                    For i = 1 To .RecordCount
                        Grd.AddItem "", False
                        Grd.Cell(Grd.Rows - 1, 0).Text = i
                        Grd.Cell(Grd.Rows - 1, 1).Text = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
                        Grd.Cell(Grd.Rows - 1, 2).Text = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
                        Grd.Cell(Grd.Rows - 1, 3).Text = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
                        Grd.Cell(Grd.Rows - 1, 4).Text = IIf(IsNull(Trim(!ChemicalReactionLiquid)), "", Trim(!ChemicalReactionLiquid))
                        Grd.Cell(Grd.Rows - 1, 5).Text = IIf(IsNull(Trim(!ManufacturerName)), "", Trim(!ManufacturerName))
                        Grd.Cell(Grd.Rows - 1, 6).Text = IIf(IsNull(Trim(!ManufacturerCode)), "", Trim(!ManufacturerCode))
                        Grd.Cell(Grd.Rows - 1, 7).Text = IIf(IsNull(Trim(!Location)), "", Trim(!Location))
                        Grd.Cell(Grd.Rows - 1, 8).Text = IIf(IsNull(Trim(!SpecifiedLocation)), "", Trim(!SpecifiedLocation))
                        Grd.Cell(Grd.Rows - 1, 9).Text = !bMix
                        Grd.Cell(Grd.Rows - 1, 11).Text = !Id
                        
                       
                       If IsNull(Trim(!Classification)) Or Trim(!Classification) = "" Then
                                
                       Else

                            If IsNull(Trim(!Pictograms)) Or Trim(!Pictograms) = "" Then
                            
                                Grd.Cell(Grd.Rows - 1, 10).BackColor = vbColorOrange
                            Else
                            
                                Grd.Cell(Grd.Rows - 1, 10).BackColor = vbColorRed
                            End If

                       End If
                       
                       
                       If !bMix Then
                           ' For t = 1 To Grd.Cols - 1
                               ' Grd.Cell(Grd.Rows - 1, t).ForeColor = &H644603
                                'Grd.Cell(Grd.Rows - 1, t).FontBold = True
                                
                            'Next
                       
                       End If
                        
                        .MoveNext
                    Next
                End If
                
            End With
            
           
            For i = 0 To .Cols - 1
                .Column(i).AutoFit
            Next
            
            
            .Column(2).Width = 200
            .Column(3).Width = 150
            .Column(4).Width = 200
            
            
            .Column(5).Width = 0
            .Column(6).Width = 0
            
            .Column(9).Width = 0
            .Column(10).Width = 10
            .Column(11).Width = 0
            
            .Refresh
            .ReadOnly = True
            .AutoRedraw = True
        
        End With
        
        

End Function

