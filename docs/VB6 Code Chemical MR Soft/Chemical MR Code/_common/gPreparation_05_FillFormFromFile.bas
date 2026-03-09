Attribute VB_Name = "gPreparation_05_FillFromFile"
Option Explicit
Private PreparationWeight As Double


    
Public Function FillGridPreparationFromFile(ByVal Grid As Grid, ByRef uPreparation As RecipeForProduction, ByVal Index As Integer, ByVal PreparationID As Long, Optional bAcquisitiTutti As Boolean, Optional bManual As Boolean)

Dim i As Integer
PreparationWeight = 0
    If PreparationID > 0 Then
        With dbTabPreparation
            .filter = ""
            .filter = "ID='" & PreparationID & "'"
            If .EOF Then
                PreparationWeight = 0
            Else
                PreparationWeight = IIf(IsNull(Trim(!QtyToProduce)), 0, Trim(!QtyToProduce))
            End If
        End With
    End If

    Select Case Index
        Case 1
            ' Component
            Call GetPreparationSTDGrid(Grid, uPreparation, bAcquisitiTutti, bManual)
        Case 2
            ' Acquisitions
            Call GetAcquisitionGrid(Grid, uPreparation, bManual)
        Case 3
          
    End Select

    
End Function



Public Sub GetPreparationSTDGrid(ByVal Grid As Grid, ByRef uPreparation As RecipeForProduction, Optional bAcquisitiTutti As Boolean, Optional bManual As Boolean)
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
        
        
        If bManual Then
        Else
        
            Call SetSTDTheoreticalWeight(uPreparation.MsType, uPreparation)
        
        End If
            
            

        With uPreparation.Recipe

            bUmMassa = .bUmMassa
            
           
            
           
          '  uPreparation.Recipe.STD(i).TheoreticalWeight
            
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
                Grid.Cell(Grid.Rows - 1, 1).Text = IIf(bManual, .STD(i).MRCode, .STD(i).NUMBER)
                Grid.Cell(Grid.Rows - 1, 2).Text = .STD(i).Value
                Grid.Cell(Grid.Rows - 1, 4).Text = PadString(.STD(i).TheoreticalWeight)
                
                     
                Grid.Cell(Grid.Rows - 1, 3).Text = uPreparation.Recipe.STDUnit
                

                Grid.Cell(Grid.Rows - 1, 5).Text = PadString(.STD(i).RealWeight)
                Grid.Cell(Grid.Rows - 1, 5).BackColor = vbColorResults
                If i > 1 Then
                Grid.Cell(Grid.Rows - 1, 5).BackColor = IIf(.STD(i).RealWeight > 0, Grid.Cell(Grid.Rows - 1, 5).BackColor, vbColorRosaTabella)
                ' Grid.Cell(Grid.Rows - 1, 1).BackColor = IIf(.STD(i).RealWeight > 0, Grid.Cell(Grid.Rows - 1, 1).BackColor, vbColorRosaTabella)
                 Grid.Cell(Grid.Rows - 1, 2).BackColor = IIf(.STD(i).RealWeight > 0, Grid.Cell(Grid.Rows - 1, 2).BackColor, vbColorRosaTabella)
                End If
        
               TotalRealWeight = TotalRealWeight + .STD(i).TheoreticalWeight

                PesoIntolleranza = .STD(i).RealWeight * 0.1
                
                If .STD(i).TheoreticalWeight > 0 Then
                
                Variance = .STD(i).TheoreticalWeight - .STD(i).RealWeight
                VariancePerc = (Variance / .STD(i).TheoreticalWeight) * 100
                   
                End If
                
                PesoIntolleranza = .STD(i).TheoreticalWeight * TolerancePerc
                
                
                MyColor = ColorTolerance(Variance, PesoIntolleranza, bAcquisitiTutti, bCorrection)
                Grid.Cell(Grid.Rows - 1, 8).BackColor = MyColor
                  
                   If bCorrection Then
                        uPreparation.bCorrection = True
                   End If
                   .STD(i).Variance = Variance
                   .STD(i).VariancePerc = VariancePerc

                    Grid.Cell(Grid.Rows - 1, 6).Text = -PadString(Variance)
                    Grid.Cell(Grid.Rows - 1, 7).Text = -FormatNumber(VariancePerc, 2) & " %"
               
                
                Grid.Cell(Grid.Rows - 1, 9).Text = .STD(i).Note
                Grid.Cell(Grid.Rows - 1, 10).Text = .STD(i).STD_ID
                
                  
                
              
                Grid.Cell(Grid.Rows - 1, 4).FontBold = True

                
                Grid.Cell(Grid.Rows - 1, 4).BackColor = vbColorResults
                
                Grid.Cell(Grid.Rows - 1, 3).Alignment = cellLeftCenter
                
                
                Grid.Cell(Grid.Rows - 1, 4).Alignment = cellCenterCenter
                Grid.Cell(Grid.Rows - 1, 5).Alignment = cellCenterCenter
                Grid.Cell(Grid.Rows - 1, 6).Alignment = cellCenterCenter
                Grid.Cell(Grid.Rows - 1, 7).Alignment = cellCenterCenter
                
              
                Grid.Cell(Grid.Rows - 1, 6).BackColor = vbColorResults
                Grid.Cell(Grid.Rows - 1, 7).BackColor = vbColorResults
                
                Grid.Cell(Grid.Rows - 1, 10).BackColor = vbColorResults
            
    
               
cont:
            Next

            Select Case .Type
                Case 0
                    Grid.Cell(0, 4).Text = "MR Qty"
                    Grid.Cell(0, 5).Text = "MR Acquired"
                Case 1, 2
                    Grid.Cell(0, 4).Text = "MS Qty"
                    Grid.Cell(0, 5).Text = "MS Acquired"
            End Select
            
            .STDcount = Grid.Rows - 1
    
    
            .TotalWeight = TotalRealWeight
            
        End With
        
        .ReadOnly = True
        
        If bManual Then
            .Column(2).Width = 0
            .Column(3).Width = 0
        Else
            .Column(2).AutoFit
            .Column(3).AutoFit
        End If
        
        .Column(3).Alignment = cellLeftCenter
        .Column(3).Width = .Column(3).Width * 2
        .Column(8).Width = 10
        .Column(10).AutoFit
        
        
        
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


Private Function GetAcquisitionGrid(ByVal Grid As Grid, ByRef uPreparation As RecipeForProduction, Optional ByVal bManual As Boolean)
Dim i As Integer
Dim t As Integer
Dim RecipeCount As Integer
Dim Variance As Double
Dim VariancePerc As Double
Dim TotalRealWeight As Double
Dim bUmMassa As Boolean
Dim Density As Double
Dim bRecalculate As Boolean
Dim PesoIntolleranza As Double
Dim MyColor As OLE_COLOR
On Error GoTo ERR_GET:
        '------------------------------------------------
        '      Acquisition Grid
        '------------------------------------------------
    With Grid
      
      
      .Rows = 1

        .AutoRedraw = False
        .ReadOnly = False


        With uPreparation.Recipe
            bUmMassa = .bUmMassa
           ' Density = .Density
            If .AcquisitionCount > 0 Then
                For i = 1 To .AcquisitionCount
                    Call AddNewRowInAcquisition(Grid, uPreparation.Recipe.Acquisitions(i))
                Next
            End If
                    
         For i = 1 To Grid.Cols - 1
         Grid.Column(i).AutoFit
         Next
         Grid.Column(11).Width = 0
         If bManual Then
            Grid.Cell(0, 3).Text = "MR Code"
            Grid.Column(4).Width = 0
            Grid.Column(17).Width = 0
           
         End If
            
             Select Case .Type
              Case 0
                 ' Grid.Cell(0, 4).Text = "MR Qty"
                  Grid.Cell(0, 7).Text = "MR Acquired"
                  Grid.Column(13).Width = 0
              Case 1, 2
                 ' Grid.Cell(0, 4).Text = "MS Qty"
                  Grid.Cell(0, 7).Text = "MS Acquired"
                 
          End Select
            
            
            
        End With
        
        
        

        
        
        
        .Column(2).AutoFit
        .Column(3).AutoFit
        .Column(5).AutoFit
        .Column(6).AutoFit
        .Column(7).AutoFit
        .Column(8).AutoFit
        .Column(9).AutoFit
        
        
        .ReadOnly = True
        .Refresh
        .AutoRedraw = True
    End With
ERR_END:
   
   On Error GoTo 0
    Exit Function
ERR_GET:
    MsgBox Err.Description
    Resume Next
End Function


