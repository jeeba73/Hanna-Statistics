Attribute VB_Name = "hProduction_05_FillFromfile"
Option Explicit


    
Public Function FillGridSTDPreparationFromFile(ByVal Grid As Grid, ByRef iSTDPreparation As RecipeForSTDPreparation, ByVal Index As Integer)

Dim i As Integer

    Select Case Index
        Case 1
            ' Hanna Codes
            Call GetCodeGrid(Grid, iSTDPreparation)
        Case 2
            ' Acquisition
            Call GetAcquisitionGrid(Grid, iSTDPreparation)
    End Select

    
End Function
    


Private Sub GetCodeGrid(ByVal Grid As Grid, ByRef iSTDPreparation As RecipeForSTDPreparation)
Dim i As Integer
Dim t As Integer
Dim HannaCount As Integer
Dim Variance As String
Dim VarDbl As Double
Dim PercStr As String

On Error GoTo ERR_GET:
        '------------------------------------------------
        '      RecipeForSTDPreparation  TABELLA Codici
        '------------------------------------------------
        
    With Grid
      
      
      .Rows = 1

        .AutoRedraw = False
    
        
        
        '.Cell(0, 1).Text = "Code"
        '.Cell(0, 2).Text = "Product Name"
        '.Cell(0, 3).Text = "Line"
        '.Cell(0, 4).Text = "Volume/Weight"
        '.Cell(0, 5).Text = "(um)"
        '.Cell(0, 6).Text = "Q.ty to produce"
        '.Cell(0, 7).Text = "Q.ty to produced"
        '.Cell(0, 8).Text = ""
        
        '.Cell(0, 9).Text = "%"
        '.Cell(0, 10).Text = "Recipe"
        '.Cell(0, 11).Text = "Mix"
        
        With iSTDPreparation
            HannaCount = .HannaCodesCount
            
            For i = 1 To HannaCount
                
                Grid.AddItem "", False
                Grid.Cell(Grid.Rows - 1, 1).Text = .HannaCodes(i).Code
                Grid.Cell(Grid.Rows - 1, 2).Text = .HannaCodes(i).ProductName
                Grid.Cell(Grid.Rows - 1, 3).Text = .HannaCodes(i).Line
                Grid.Cell(Grid.Rows - 1, 4).Text = .HannaCodes(i).Qty
                Grid.Cell(Grid.Rows - 1, 5).Text = .HannaCodes(i).Um
                Grid.Cell(Grid.Rows - 1, 6).Text = .HannaCodes(i).QtyToProduce
                Grid.Cell(Grid.Rows - 1, 7).Text = .HannaCodes(i).QtyProduced
                
                If .HannaCodes(i).QtyToProduce = "" Then .HannaCodes(i).QtyToProduce = "0"
                If .HannaCodes(i).QtyProduced = "" Then .HannaCodes(i).QtyProduced = "0"
                
                If CDbl(.HannaCodes(i).QtyProduced) > 0 And CDbl(.HannaCodes(i).QtyToProduce) > 0 Then
                
                    VarDbl = FormatNumber((.HannaCodes(i).QtyProduced / .HannaCodes(i).QtyToProduce), 4) * 100
                     
                    Select Case VarDbl
                        Case Is < 100
                            PercStr = "- "
                            VarDbl = FormatNumber(100 - VarDbl, 2)
                        Case Is = 100
                            PercStr = ""
                            VarDbl = VarDbl
                        Case Is > 100
                            PercStr = "+ "
                            VarDbl = VarDbl - 100
                    End Select
                                       
                    Variance = PercStr & VarDbl & " %"
                    Grid.Cell(Grid.Rows - 1, 9).Text = Variance

                    VarDbl = CDbl(.HannaCodes(i).QtyProduced) - CDbl(.HannaCodes(i).QtyToProduce)
                End If
                
                If CDbl(.HannaCodes(i).QtyToProduce) = 0 Then
                    VarDbl = CDbl(.HannaCodes(i).QtyProduced)
                End If
                Grid.Cell(Grid.Rows - 1, 10).Text = .HannaCodes(i).Recipe
                Grid.Cell(Grid.Rows - 1, 11).Text = .HannaCodes(i).Mix1 & IIf(Len(.HannaCodes(i).Mix2) > 0, ";" & .HannaCodes(i).Mix2, "")
                Grid.RowHeight(Grid.Rows - 1) = IIf(.HannaCodes(i).bHide, 0, Grid.RowHeight(Grid.Rows - 1))
                
                
                If CDbl(.HannaCodes(i).QtyProduced) > 0 Then
                    For t = 1 To Grid.Cols - 1

                       Grid.Cell(Grid.Rows - 1, t).FontBold = True
                       Grid.Cell(Grid.Rows - 1, t).ForeColor = &H404040 '&H644603
                
                        
                    Next
                    
                    Select Case VarDbl
                    
    
                        Case -.HannaCodes(i).QtyToProduce * 0.2 To -.HannaCodes(i).QtyToProduce * 0.02
                            Grid.Cell(Grid.Rows - 1, 8).BackColor = vbColorOrange
                        Case Is < -.HannaCodes(i).QtyToProduce * 0.2
                            Grid.Cell(Grid.Rows - 1, 8).BackColor = &HC0&
                        Case Else
                            Grid.Cell(Grid.Rows - 1, 8).BackColor = vbColorGreen
                    End Select
                
                End If
                
                
            Next
        
        
        End With
ERR_END:
         Call SetSTDPreparationHannaGridSpecific(Grid)
         .Column(11).AutoFit
         .Column(4).Alignment = cellRightCenter
         .Column(5).Alignment = cellLeftCenter
         
        .Refresh
        .AutoRedraw = True
    End With
    Exit Sub
ERR_GET:
   MsgBox err.Description
   Resume Next

End Sub

Public Function SetSTDPreparationHannaGridSpecific(ByVal Grd As Grid, Optional ByVal bSTDPreparation As Boolean)

Dim t As Integer
Dim i As Integer

With Grd
    For t = 1 To .Rows - 1
        
        
        
        .Cell(t, 1).FontSize = 11
        .Cell(t, 2).FontSize = 9
        .Cell(t, 1).FontBold = True
        .Cell(t, 1).ForeColor = &H404040
        

        .Cell(t, 6).BackColor = vbColorResults
        .Cell(t, 7).BackColor = vbColorResults
       ' .Cell(t, 14).BackColor = vbColorResults
    Next
    .Column(1).AutoFit
    
    If bSTDPreparation Then
        .Column(7).Width = 200
    Else
        .Column(8).AutoFit
    End If
    
    .ReadOnly = True
    .Refresh
    .AutoRedraw = True
End With
End Function

Private Sub GetAcquisitionGrid(ByVal Grid As Grid, ByRef uSTDPreparation As RecipeForSTDPreparation)
Dim i As Integer
Dim t As Integer

On Error GoTo ERR_GET:
        '------------------------------------------------
        '      Acquisition Grid
        '------------------------------------------------
    With Grid
      
      
      .Rows = 1

        .AutoRedraw = False
        .ReadOnly = False


    For i = 1 To UBound(uSTDPreparation.HannaCodes)

        With uSTDPreparation.HannaCodes(i)
          
          
          
            If .AcquisitionCount > 0 Then
                For t = 1 To .AcquisitionCount
                    Call STDPreparationAddNewRowInAcquisition(Grid, .Acquisitions(t))
                Next
            End If
        End With
    Next
    
        .FrozenCols = 3
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
Public Function STDPreparationAddNewRowInAcquisition(ByRef Grid2 As Grid, ByRef iAcquisition As ProdAcquisition)
Dim i As Integer
Dim t As Integer
With Grid2

 
    If iAcquisition.Code = "" Then Exit Function
    
        .AddItem "", False
        i = .Rows - 1

        .Cell(i, 1).Text = iAcquisition.Code
        .Cell(i, 2).Text = iAcquisition.QtyProduced
        .Cell(i, 3).Text = iAcquisition.LotNumber
        .Cell(i, 4).Text = iAcquisition.Operator
        .Cell(i, 5).Text = iAcquisition.DateProd
        .Cell(i, 6).Text = iAcquisition.WeekProd
        .Cell(i, 7).Text = iAcquisition.Machine
        .Cell(i, 8).Text = iAcquisition.Note
        
        .Cell(i, 9).Text = iAcquisition.AcquisitionTime
        .Cell(i, 10).Text = iAcquisition.ID
        .Cell(i, 11).Text = iAcquisition.Index
        
        
        .Cell(i, 12).Text = iAcquisition.Mix1Lot
        .Cell(i, 13).Text = iAcquisition.Mix2Lot
        
        .Cell(i, 14).Text = iAcquisition.ExpDate
        
        
                
        .Cell(i, 14).BackColor = vbColorResults
        .Cell(i, 3).BackColor = vbColorResults
        .Cell(i, 2).BackColor = vbColorResults
        
        .Cell(i, 2).Alignment = cellRightCenter
        

        
End With

End Function

