Attribute VB_Name = "gPreparation_05_FillFromFile"
Option Explicit
Private PreparationWeight As Double


    
Public Function FillGridPreparationFromFile(ByVal Grid As Grid, ByRef uPreparation As RecipeForProduction, ByVal Index As Integer, ByVal PreparationID As Long)

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
            Call GetComponentGrid(Grid, uPreparation)
        Case 2
            ' Acquisitions
            Call GetAcquisitionGrid(Grid, uPreparation)
        Case 3
            ' hanna code
            Call GetHannaCodeGrid(Grid, uPreparation)
    End Select

    
End Function
    


Private Sub GetComponentGrid(ByVal Grid As Grid, ByRef uPreparation As RecipeForProduction)
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
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grid
      
      
      .Rows = 1

        .AutoRedraw = False
        .ReadOnly = False


        With uPreparation.Recipes(1)
            uPreparation.bCorrection = False
            bUmMassa = .bUmMassa
            Density = .Density
            'If .RmxRecipeCount = 0 Then GoTo ERR_END:
            For i = 0 To .RmxRecipeCount
         
                     
                If uPreparation.Recipes(1).Code <> .RmxRecipe(i).RecipeCode Then GoTo cont:
                If .RmxRecipe(i).bDeleted Then GoTo cont:
                
                
            ' .Cell(0, 1).Text = "Code"
            ' .Cell(0, 2).Text = "Description"
            ' .Cell(0, 3).Text = "CAS"
            ' .Cell(0, 4).Text = "Theorethical weight"
            ' .Cell(0, 5).Text = "Real Weight"
            ' .Cell(0, 6).Text = "Variance"
            ' .Cell(0, 7).Text = "Variance %"
            ' .Cell(0, 8).Text = ""
            ' .Cell(0, 9).Text = "Note"
            ' .Cell(0, 10).Text = "bMix"
        


                Grid.AddItem "", False
                Grid.Cell(Grid.Rows - 1, 2).Text = .RmxRecipe(i).CHCode
                Grid.Cell(Grid.Rows - 1, 3).Text = .RmxRecipe(i).Description
                Grid.Cell(Grid.Rows - 1, 4).Text = .RmxRecipe(i).Cas
                
                If PreparationWeight > 0 And PreparationWeight <> uPreparation.Recipes(1).TotalWeightKg Then
                        uPreparation.Recipes(1).TotalWeightKg = PreparationWeight
                        GoTo SetComponent
                
                End If
                
                If uPreparation.Recipes(1).TotalWeightKg = 0 Then
                    ' devo ricalcolare altrimenti errori ovunque!!
                        If PreparationWeight > 0 Then
                            uPreparation.Recipes(1).TotalWeightKg = PreparationWeight
                            GoTo SetComponent
                        End If
                        
                        InputWeight = uPreparation.Recipes(1).TotalWeightKg
                        PopupMessage 2, "Recipe without Total Weight....", , True
                        If F_InputBox.DoShow("Enter Total Weight ( kg ) ", .RmxRecipe(i).CHCode, , , , InputWeight, , True) Then
                            If InputWeight <= 0 Then
                                Exit Sub
                            Else
                                uPreparation.Recipes(1).TotalWeightKg = InputWeight
                            End If
                        End If
SetComponent:
                        For x = 0 To UBound(uPreparation.Recipes(1).RmxRecipe)
                            .RmxRecipe(x).TheoreticalWeight = (uPreparation.Recipes(1).TotalWeightKg * 1000 * .RmxRecipe(x).Perc) / 100
                        
                        Next
                
                End If
                
                
                
                If .RmxRecipe(i).bAddedInPreparation Then
                    If .RmxRecipe(i).TheoreticalWeight <> .RmxRecipe(i).RealWeight Then

                        Grid.Cell(Grid.Rows - 1, 6).Text = PadString(.RmxRecipe(i).TheoreticalWeight)
                        TheorWeight = .RmxRecipe(i).TheoreticalWeight
                    Else
                        
                        Grid.Cell(Grid.Rows - 1, 6).Text = "-"
                        TheorWeight = 0
                    End If
                    'Grid.Cell(Grid.Rows - 1, 6).Text = "-"
                    Grid.Cell(Grid.Rows - 1, 5).Text = "-"
                    
                Else
                    Grid.Cell(Grid.Rows - 1, 5).Text = FormatNumber(.RmxRecipe(i).Perc, 4)
                    
                    TheorWeight = .RmxRecipe(i).TheoreticalWeight
                    If TheorWeight = 0 Then
                        InputWeight = TheorWeight
                        If F_InputBox.DoShow("Theoretical Weight?", .RmxRecipe(i).CHCode, , , , InputWeight, , True) Then
                            TheorWeight = InputWeight
                            .RmxRecipe(i).TheoreticalWeight = TheorWeight
                            
                        End If
                    End If
                    Grid.Cell(Grid.Rows - 1, 6).Text = PadString(.RmxRecipe(i).TheoreticalWeight)
                End If
                
                
               
                
                Grid.Cell(Grid.Rows - 1, 7).Text = PadString(.RmxRecipe(i).RealWeight)
                
                
                PesoIntolleranza = .RmxRecipe(i).RealWeight * .RmxRecipe(i).TolerancePerc / 100
                
                Variance = .RmxRecipe(i).RealWeight - TheorWeight
                
                If .RmxRecipe(i).bAddedInPreparation Then
                    If .RmxRecipe(i).TheoreticalWeight <> .RmxRecipe(i).RealWeight Then
                        GoTo calcVariance
                    End If
                Else
calcVariance:
                If TheorWeight = 0 Then
                    TheorWeight = .RmxRecipe(i).RealWeight
                End If
                
                
                   VariancePerc = (Variance / TheorWeight) * 100
                   
                   
                '.RmxRecipe(i).Variance = .RmxRecipe(i).RealWeight - .RmxRecipe(i).TheoreticalWeight
               ' .RmxRecipe(i).VariancePerc = (.RmxRecipe(i).Variance / .RmxRecipe(i).RealWeight) * 100
                
                
                   
                   
                   MyColor = ColorTolerance(Variance, PesoIntolleranza, bRecalculate, bCorrection)
                   If bCorrection Then
                        uPreparation.bCorrection = True
                   End If
                   .RmxRecipe(i).Variance = Variance
                   .RmxRecipe(i).VariancePerc = VariancePerc
                   
                End If
                
                TotalRealWeight = TotalRealWeight + .RmxRecipe(i).RealWeight
                
                If .RmxRecipe(i).bAddedInPreparation Then
                    If .RmxRecipe(i).TheoreticalWeight <> .RmxRecipe(i).RealWeight Then
                        GoTo PrintVariance
                    Else
                        Grid.Cell(Grid.Rows - 1, 8).Text = PadString(.RmxRecipe(i).RealWeight)
                        Grid.Cell(Grid.Rows - 1, 9).Text = "-"
          
                    End If
                Else
PrintVariance:
                    Grid.Cell(Grid.Rows - 1, 8).Text = PadString(Variance)
                    Grid.Cell(Grid.Rows - 1, 9).Text = FormatNumber(VariancePerc, 2) & "%"
                End If
                
                Grid.Cell(Grid.Rows - 1, 11).Text = .RmxRecipe(i).RealPerc
                Grid.Cell(Grid.Rows - 1, 12).Text = .RmxRecipe(i).Note
                Grid.Cell(Grid.Rows - 1, 13).Text = .RmxRecipe(i).bMix
                
                
                
                Grid.Cell(Grid.Rows - 1, 14).Text = .RmxRecipe(i).CriticalRM
                
                
              
                
                Grid.Cell(Grid.Rows - 1, 3).FontSize = 9
                Grid.Cell(Grid.Rows - 1, 5).Alignment = cellCenterCenter
                Grid.Cell(Grid.Rows - 1, 6).Alignment = cellRightCenter
                Grid.Cell(Grid.Rows - 1, 7).Alignment = cellRightCenter
                Grid.Cell(Grid.Rows - 1, 8).Alignment = cellRightCenter
                Grid.Cell(Grid.Rows - 1, 9).Alignment = cellRightCenter
                
                
                Grid.Cell(Grid.Rows - 1, 6).BackColor = vbColorResults
                Grid.Cell(Grid.Rows - 1, 7).BackColor = vbColorResults
                Grid.Cell(Grid.Rows - 1, 8).BackColor = vbColorResults
                Grid.Cell(Grid.Rows - 1, 9).BackColor = vbColorResults
                
                
                If .RmxRecipe(i).RealWeight > 0 And Not (.RmxRecipe(i).bAddedInPreparation) Then
                    Grid.Cell(Grid.Rows - 1, 10).BackColor = MyColor
                End If
                
                
                
                If .RmxRecipe(i).bAddedInPreparation Then
                
                    Grid.Cell(Grid.Rows - 1, 10).BackColor = &HFFFF&
                    
                    
                End If
                
                If .RmxRecipe(i).bCorrection Then
                    Grid.Cell(Grid.Rows - 1, 10).BackColor = &HFFFF&
                End If
                
               
               
               For t = 1 To Grid.Cols - 1

                    If .RmxRecipe(i).bMix Then
                        Grid.Cell(Grid.Rows - 1, t).FontBold = True
                        Grid.Cell(Grid.Rows - 1, t).ForeColor = &H644603
                    End If
                    
                    If Len(.RmxRecipe(i).CriticalRM) > 0 Then
                        Grid.Cell(Grid.Rows - 1, t).FontBold = True
                        Grid.Cell(Grid.Rows - 1, t).ForeColor = &H40C0&
                    End If
                    
                Next
                        
               
cont:
            Next
            
            TotalRealWeight = TotalRealWeight / 1000
            .ActualWeight = TotalRealWeight
            .ActualWeightUm = "kg"
            
            
            For i = 1 To Grid.Rows - 1
            
                If .ActualWeight > 0 Then
                    .RmxRecipe(i - 1).RealPerc = FormatNumber((.RmxRecipe(i - 1).RealWeight / (.ActualWeight * 1000)) * 100, 4)
                Else
                    .RmxRecipe(i - 1).RealPerc = 0
                End If
                Grid.Cell(i, 11).Text = .RmxRecipe(i - 1).RealPerc
            Next
            
            
            
            .RecipeComponentCount = Grid.Rows - 1
        
            ' totals Kili
        
        
            Grid.AddItem "", False
            Grid.AddItem "", False
            
            
            Grid.Range(Grid.Rows - 1, 1, Grid.Rows - 1, 4).Merge
            Grid.Cell(Grid.Rows - 1, 1).Text = "TotalWeight (Kg)"
            Grid.Cell(Grid.Rows - 1, 1).Alignment = cellRightCenter
            
            Grid.Cell(Grid.Rows - 1, 6).Text = PadString(uPreparation.Recipes(1).TotalWeightKg)
            Grid.Cell(Grid.Rows - 1, 7).Text = PadString(TotalRealWeight)
            
            Variance = TotalRealWeight - uPreparation.Recipes(1).TotalWeightKg
            VariancePerc = (Variance / uPreparation.Recipes(1).TotalWeightKg) * 100
            Grid.Cell(Grid.Rows - 1, 8).Text = PadString(Variance)
            Grid.Cell(Grid.Rows - 1, 9).Text = FormatNumber(VariancePerc, 2) & "%"
            
            Grid.Cell(Grid.Rows - 1, 1).FontBold = True
            Grid.Cell(Grid.Rows - 1, 6).FontBold = True
            Grid.Cell(Grid.Rows - 1, 7).FontBold = True
            Grid.Cell(Grid.Rows - 1, 8).FontBold = True
            Grid.Cell(Grid.Rows - 1, 9).FontBold = True
            Grid.Cell(Grid.Rows - 1, 1).ForeColor = &H473733  ' &H644603
            Grid.Cell(Grid.Rows - 1, 6).ForeColor = &H473733   ' &H644603
            Grid.Cell(Grid.Rows - 1, 7).ForeColor = &H473733   ' &H644603
            Grid.Cell(Grid.Rows - 1, 8).ForeColor = &H473733   ' &H644603
            Grid.Cell(Grid.Rows - 1, 9).ForeColor = &H473733   ' &H644603
            
            If .bRecalculation Then
                Grid.Cell(Grid.Rows - 1, 1).ForeColor = vbColorRed   ' &H644603
                Grid.Cell(Grid.Rows - 1, 6).ForeColor = vbColorRed   ' &H644603
                Grid.Cell(Grid.Rows - 1, 6).Text = Grid.Cell(Grid.Rows - 1, 6).Text & " (R)"
            End If
            
            '&H00886010&
            
            If Not (bUmMassa) Then
            
                ' tot Litri
                
                Grid.AddItem "", False
                
                
                Grid.Range(Grid.Rows - 1, 1, Grid.Rows - 1, 4).Merge
                Grid.Cell(Grid.Rows - 1, 1).Text = "TotalWeight (L)"
                Grid.Cell(Grid.Rows - 1, 1).Alignment = cellRightCenter
                
                Grid.Cell(Grid.Rows - 1, 6).Text = PadString(uPreparation.Recipes(1).TotalWeightKg / Density)
                Grid.Cell(Grid.Rows - 1, 7).Text = PadString(TotalRealWeight / Density)
                
                Variance = (TotalRealWeight / Density - uPreparation.Recipes(1).TotalWeightKg / Density)
                VariancePerc = (Variance / (uPreparation.Recipes(1).TotalWeightKg / Density)) * 100
                Grid.Cell(Grid.Rows - 1, 8).Text = PadString(Variance)
                Grid.Cell(Grid.Rows - 1, 9).Text = FormatNumber(VariancePerc, 2) & "%"
                
                Grid.Cell(Grid.Rows - 1, 1).FontBold = True
                Grid.Cell(Grid.Rows - 1, 6).FontBold = True
                Grid.Cell(Grid.Rows - 1, 7).FontBold = True
                Grid.Cell(Grid.Rows - 1, 8).FontBold = True
                Grid.Cell(Grid.Rows - 1, 9).FontBold = True
                
                
                Grid.Cell(Grid.Rows - 1, 1).ForeColor = &H574743  ' &H886010
                Grid.Cell(Grid.Rows - 1, 6).ForeColor = &H574743  '&H886010
                Grid.Cell(Grid.Rows - 1, 7).ForeColor = &H574743  '&H886010
                Grid.Cell(Grid.Rows - 1, 8).ForeColor = &H574743  '&H886010
                Grid.Cell(Grid.Rows - 1, 9).ForeColor = &H574743  '&H886010
                
                If .bRecalculation Then
                    Grid.Cell(Grid.Rows - 1, 1).ForeColor = vbColorRed   ' &H644603
                    Grid.Cell(Grid.Rows - 1, 6).ForeColor = vbColorRed   ' &H644603
                    Grid.Cell(Grid.Rows - 1, 6).Text = Grid.Cell(Grid.Rows - 1, 6).Text & " (R)"
                End If
            
            End If
            
            
                
                
        End With
        .Column(12).AutoFit
        .Column(14).AutoFit
        .Column(0).Width = 0
        .ReadOnly = True
        .Column(2).AutoFit

ERR_END:
        .Refresh
        .AutoRedraw = True
    End With
   On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox err.Description
    Resume Next

End Sub


Private Sub GetAcquisitionGrid(ByVal Grid As Grid, ByRef uPreparation As RecipeForProduction)
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


        'With Acquisition(r)
        '    .AcquisitionTime = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "AcquisitionTime", .AcquisitionTime)
        '    .ActualWeight = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "ActualWeight", .ActualWeight)
        '    .bFromBarcode = SettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "bFromBarcode", .bFromBarcode)
        '    .bRecalculation = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "bRecalculation", .bRecalculation)
        '    .bRecipeComponent = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "bRecipeComponent", .bRecipeComponent)
        '    .ID = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "ID", .ID)
        '    .Index = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "Index", .Index)
        '    .Note = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "Note", .Note)
        '    .Operator = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "Operator", .Operator)
        '    .PrepBarcode.Cas = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "Cas", .PrepBarcode.Cas)
        '    .PrepBarcode.ChemicalName = SettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "ChemicalName", .PrepBarcode.ChemicalName)
        '    .PrepBarcode.Code = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "Code", .PrepBarcode.Code)
        '    .PrepBarcode.DeliveryDate = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "DeliveryDate", .PrepBarcode.DeliveryDate)
        '    .PrepBarcode.Manufacturer = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "Manufacturer", .PrepBarcode.Manufacturer)
        '    .PrepBarcode.ManufacturerCode = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "ManufacturerCode", .PrepBarcode.ManufacturerCode)
        '    .PrepBarcode.ManufacturerLot = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "ManufacturerLot", .PrepBarcode.ManufacturerLot)
        '    .PrepBarcode.Package = SettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "Package", .PrepBarcode.Package)
        '    .PrepBarcode.QtyDelivered = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "QtyDelivered", .PrepBarcode.QtyDelivered)
        '    .PrepBarcode.WeekDelPackageNumber = GetSettingData(SettingName, "Recipes" & t & " - Acquisition " & r, "WeekDelPackageNumber", .PrepBarcode.WeekDelPackageNumber)
       '
        'End With
        
        
     '.Cell(0, 1).Text = "Code"
    '.Cell(0, 2).Text = "Description"
    '.Cell(0, 3).Text = "CAS"
    '.Cell(0, 4).Text = "Real Weight (g)"
    '.Cell(0, 5).Text = "Manufacturer"
    '.Cell(0, 6).Text = "Manufacturer Code"
    '.Cell(0, 7).Text = "Manufacturer Lot"
    '.Cell(0, 8).Text = "Delivery Date"
    '.Cell(0, 9).Text = "Qty Delivered"
    '.Cell(0, 10).Text = "Week Delivery"
    '.Cell(0, 11).Text = "Package"
    
    '.Cell(0, 12).Text = "Note"
    '.Cell(0, 13).Text = "Operator"
    '.Cell(0, 14).Text = "Acquisition Time"
    '.Cell(0, 15).Text = "ID"




        With uPreparation.Recipes(1)
            bUmMassa = .bUmMassa
            Density = .Density
            If .AcquisitionCount > 0 Then
                For i = 1 To .AcquisitionCount
                    Call AddNewRowInAcquisition(Grid, uPreparation.Recipes(1).Acquisitions(i))
                Next
            End If
        End With
        
        .ReadOnly = True
        .Refresh
        .AutoRedraw = True
    End With
ERR_END:
   
   On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox err.Description
    Resume Next
End Sub

Private Sub GetHannaCodeGrid(ByVal Grid As Grid, ByRef uPreparation As RecipeForProduction)
Dim i As Integer
Dim HannaCount As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    
On Error GoTo ERR_GET:
    
    With Grid
      
      
      .Rows = 1

        .AutoRedraw = False
    
        
        '.Cell(0, 1).Text = "Code"
        ''.Cell(0, 2).Text = "Product Name"
        '.Cell(0, 3).Text = "Line"
        '.Cell(0, 4).Text = "Volume/Weight"
        '.Cell(0, 5).Text = "(um)"
        '.Cell(0, 6).Text = "Q.ty to produce"
        '.Cell(0, 7).Text = "Recipe"
        '.Cell(0, 8).Text = "Mix"
        
        
        With uPreparation
            HannaCount = .HannaCodesCount
            
            For i = 1 To HannaCount
                
               
                    Grid.AddItem "", False
                    Grid.Cell(Grid.Rows - 1, 1).Text = .HannaCodes(i).Code
                    Grid.Cell(Grid.Rows - 1, 2).Text = .HannaCodes(i).ProductName
                    Grid.Cell(Grid.Rows - 1, 3).Text = .HannaCodes(i).Line
                    Grid.Cell(Grid.Rows - 1, 4).Text = .HannaCodes(i).Qty
                    Grid.Cell(Grid.Rows - 1, 5).Text = .HannaCodes(i).Um
                    Grid.Cell(Grid.Rows - 1, 6).Text = .HannaCodes(i).QtyToProduce
                    Grid.Cell(Grid.Rows - 1, 7).Text = .HannaCodes(i).LotNumber
                    
                    
                    If (.HannaCodes(i).bHide) Then Grid.RowHeight(Grid.Rows - 1) = 0
                
            Next
        
        
        End With

         Call SetHannaGridSpecific(Grid, True)
        .Refresh
        .AutoRedraw = True
    End With
ERR_END:
   
   On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox err.Description
    Resume Next

End Sub
