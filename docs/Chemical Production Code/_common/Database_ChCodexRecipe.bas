Attribute VB_Name = "Database_ChCodexRecipe"
Option Explicit
Private UserRmxRecipe() As RmxRecipe

Private AddRmxRecipe As RmxRecipe


Public Sub SetDatabaseComponentGrid(ByVal Grd As Grid)
Dim i As Integer
    With Grd
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 11
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "CH Code"
        .Cell(0, 2).Text = "Description"
        
        
        .Cell(0, 3).Text = "CAS"
        .Cell(0, 4).Text = "Q.ty/multiple"
        .Cell(0, 5).Text = "(um)"
        .Cell(0, 6).Text = "%"
        
      .Cell(0, 7).Text = "Tolerance %"
        
        .Cell(0, 8).Text = "Note"
        .Cell(0, 9).Text = "Mix"
        .Cell(0, 10).Text = "Critical RM"


        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 150
            .Cell(0, i).FontBold = True
            
        Next
        
        .Column(1).Width = 80
        .Column(2).Width = 300
        .Column(3).Width = 100
        .Column(5).Width = 80
        .Column(6).Width = 60
        .Column(7).Width = 90
        .Column(8).Width = 250
        
        .Column(9).Width = 0
        
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
  

End Sub





Public Function GetChemicalsPerRecipe(ByVal Grd As Grid, ByRef uRecipe As RecipeType) As Boolean
Dim rc As Boolean
Dim i As Integer


On Error GoTo ERR_CHEM:
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        .ReadOnly = False
        
        
        
        With dbTabRMxRecipe
            .Close
            .Open "SELECT *  FROM TabRMxRecipe order by id ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
            .filter = ""
            .filter = "RecipeCode='" & uRecipe.Code & "'"
            If .EOF Then
                GoTo ERR_END:
            Else
                .MoveFirst
                
                'ReDim uRecipe.RmxRecipe(.RecordCount)
                ReDim uRecipe.RmxRecipe(.RecordCount)
                uRecipe.RmxRecipeCount = .RecordCount
                
                For i = 1 To .RecordCount
                    uRecipe.RmxRecipe(i - 1).RecipeCode = uRecipe.Code
                    uRecipe.RmxRecipe(i - 1).ID = !ID
                    uRecipe.RmxRecipe(i - 1).CHCode = IIf(IsNull(Trim(!CHCode)), "", Trim(!CHCode))
                    uRecipe.RmxRecipe(i - 1).Description = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
                    uRecipe.RmxRecipe(i - 1).Cas = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
                    uRecipe.RmxRecipe(i - 1).Qty = CheckDot(IIf(IsNull(Trim(!Qty)), 0, Trim(!Qty)))
                    uRecipe.RmxRecipe(i - 1).Um = IIf(IsNull(Trim(!Um)), "g", Trim(!Um))
                    uRecipe.RmxRecipe(i - 1).Perc = CheckDot(IIf(IsNull(Trim(!Perc)), 0, Trim(!Perc)))
                    uRecipe.RmxRecipe(i - 1).Note = IIf(IsNull(Trim(!Note)), 0, Trim(!Note))
                    uRecipe.RmxRecipe(i - 1).TolerancePerc = CheckDot(IIf(IsNull(Trim(!TolerancePerc)), 1, Trim(!TolerancePerc)))
                    uRecipe.RmxRecipe(i - 1).bMix = !bMix
                    uRecipe.RmxRecipe(i - 1).CriticalRM = GetCriticalRM(uRecipe.RmxRecipe(i - 1).CHCode)
                    .MoveNext
                Next
            End If
        End With
       
       For i = LBound(uRecipe.RmxRecipe) To UBound(uRecipe.RmxRecipe) - 1
            .AddItem "", False
            .Cell(.Rows - 1, 1).Text = uRecipe.RmxRecipe(i).CHCode
            .Cell(.Rows - 1, 2).Text = uRecipe.RmxRecipe(i).Description
            .Cell(.Rows - 1, 3).Text = uRecipe.RmxRecipe(i).Cas
            .Cell(.Rows - 1, 4).Text = PadString(uRecipe.RmxRecipe(i).Qty)
            .Cell(.Rows - 1, 5).Text = uRecipe.RmxRecipe(i).Um
            .Cell(.Rows - 1, 6).Text = FormatNumber(uRecipe.RmxRecipe(i).Perc, 4)
            .Cell(.Rows - 1, 7).Text = FormatNumber(uRecipe.RmxRecipe(i).TolerancePerc, 2)
            .Cell(.Rows - 1, 8).Text = uRecipe.RmxRecipe(i).Note
            .Cell(.Rows - 1, 9).Text = uRecipe.RmxRecipe(i).bMix
            .Cell(.Rows - 1, 10).Text = uRecipe.RmxRecipe(i).CriticalRM
       Next

       
       ' adjust Grid color and all...
       
        Dim t As Integer
        For t = 1 To .Rows - 1
                
                If .Cell(t, 9).Text = True Then
                    .Cell(t, 1).FontBold = True
                    .Cell(t, 2).FontBold = True
                    .Cell(t, 3).FontBold = True
                    .Cell(t, 1).ForeColor = &H886010
                    .Cell(t, 2).ForeColor = &H886010
                    .Cell(t, 3).ForeColor = &H886010
                End If
                
                For i = 1 To .Cols - 1
                .Column(i).Alignment = cellLeftCenter
 
                .Cell(t, 7).BackColor = &HE0E0E0  ' vbColorResults
                .Cell(t, 8).BackColor = &HE0E0E0  ' vbColorResults
                .Cell(t, 4).BackColor = &HE0E0E0  ' vbColorResults
                .Cell(t, 5).BackColor = &HE0E0E0  ' vbColorResults
                .Cell(t, 6).BackColor = &HE0E0E0  ' vbColorResults
            Next
           
       Next
       .ReadOnly = False
       .Range(0, 4, 0, 5).Merge
       .Cell(0, 4).Alignment = cellCenterCenter
        .Column(1).AutoFit
       ' .Column(2).AutoFit
        '.Column(3).AutoFit
        .Column(4).Alignment = cellRightCenter
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellCenterCenter
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
ERR_END:
   On Error GoTo 0
   GetChemicalsPerRecipe = rc
   Exit Function
ERR_CHEM:
   rc = False
   MsgBox err.Description
   Resume Next
    

End Function

Public Function GetHannaCodePerRecipe(ByVal Grd As Grid, ByRef uRecipe As RecipeType) As Boolean
Dim rc As Boolean
Dim i As Integer


On Error GoTo ERR_CHEM:
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        .ReadOnly = False
        
        Dim t As Integer
        
        
        With dbTabCode
            .Close
            .Open "SELECT *  FROM TabCode order by id ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
            .filter = ""
            .filter = ""
            '.Filter = "Recipe='" & uRecipe.Code & "' or Mix1='" & uRecipe.Code & "' or Mix2='" & uRecipe.Code & "'"
            If .EOF Then
                GoTo ERR_END:
            Else
                .MoveFirst
      
       ' .Cell(0, 1).Text = "Code"
       ' .Cell(0, 2).Text = "Product Name"
       ' .Cell(0, 3).Text = "Line"
       ' .Cell(0, 4).Text = "Volume/Weight"
       ' .Cell(0, 5).Text = "(um)"
       ' .Cell(0, 6).Text = "Q.ty to produce"
       ' .Cell(0, 7).Text = "Recipe"
       ' .Cell(0, 8).Text = "Mix #1"
       ' .Cell(0, 9).Text = "Mix #2"
                
                
                t = 0
                For i = 1 To .RecordCount
                  If InStr(Trim(!Recipe), uRecipe.Code) Or InStr(Trim(!Mix1), uRecipe.Code) Or InStr(Trim(!Mix2), uRecipe.Code) Then
                    
                        ReDim Preserve uRecipe.HannaCodes(t)
                        uRecipe.HannaCodesCount = t
                    
                        uRecipe.HannaCodes(t).Code = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
                        uRecipe.HannaCodes(t).ID = !ID
                        uRecipe.HannaCodes(t).Line = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
                        uRecipe.HannaCodes(t).ProductName = IIf(IsNull(Trim(!ProductName)), "", Trim(!ProductName))
                        uRecipe.HannaCodes(t).Recipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
                        uRecipe.HannaCodes(t).Mix1 = IIf(IsNull(Trim(!Mix1)), "", Trim(!Mix1))
                        uRecipe.HannaCodes(t).Mix2 = IIf(IsNull(Trim(!Mix2)), "", Trim(!Mix2))
                        uRecipe.HannaCodes(t).Um = IIf(IsNull(Trim(!Um)), "g", Trim(!Um))
                        uRecipe.HannaCodes(t).Qty = IIf(IsNull(Trim(!Qty)), 0, Trim(!Qty))
                    
                        t = t + 1
                    End If
                    .MoveNext
                Next
            End If
        End With
        
       If t = 0 Then Exit Function
       
       For i = LBound(uRecipe.HannaCodes) To UBound(uRecipe.HannaCodes)
            .AddItem "", False
            .Cell(.Rows - 1, 1).Text = uRecipe.HannaCodes(i).Code
            .Cell(.Rows - 1, 2).Text = uRecipe.HannaCodes(i).ProductName
            '.Cell(.Rows - 1, 3).Text = uRecipe.HannaCodes(i).Line
            '.Cell(.Rows - 1, 4).Text = uRecipe.HannaCodes(i).Qty
            '.Cell(.Rows - 1, 5).Text = uRecipe.HannaCodes(i).Um
          ' la colonna 6 è per la RecipeForProduction
           ' .Cell(.Rows - 1, 7).Text = uRecipe.HannaCodes(i).Recipe
           ' .Cell(.Rows - 1, 8).Text = uRecipe.HannaCodes(i).Mix1
           ' .Cell(.Rows - 1, 9).Text = uRecipe.HannaCodes(i).Mix2
            
           ' .Cell(.Rows - 1, 4).BackColor = vbColorAzzurrino
            '.Cell(.Rows - 1, 5).BackColor = vbColorAzzurrino
       Next
       
       
       ' adjust Grid color and all...
       
      
        For t = 1 To .Rows - 1
            For i = 1 To .Cols - 1
                .Column(i).Alignment = cellLeftCenter
            Next
           
       Next
       .ReadOnly = False
       .Range(0, 4, 0, 5).Merge
       .Cell(0, 4).Alignment = cellCenterCenter
       
        '.Column(4).Alignment = cellRightCenter
       ' .Column(6).Alignment = cellCenterCenter
        '.Column(8).AutoFit
       ' .Column(9).AutoFit
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
ERR_END:
   On Error GoTo 0
   GetHannaCodePerRecipe = rc
   Exit Function
ERR_CHEM:
   rc = False
   MsgBox err.Description
   Resume Next
    


End Function



Public Function GetHannaCodePerGrid(ByVal Grd As Grid, ByRef HannaCodes() As HannaCode) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim HannaCode As String
On Error GoTo ERR_CHEM:
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grd

    ReDim HannaCodes(.Rows - 1)

    For i = 1 To .Rows - 1
    
        HannaCode = Trim(.Cell(i, 1).Text)
       
        If Grd.RowHeight(i) = 0 Then
            HannaCodes(i).bHide = True
        Else
            HannaCodes(i).bHide = False
        End If
       
        Call GetHannaCodexRecipe(HannaCode, HannaCodes(i), "")
        
        HannaCodes(i).QtyToProduce = Grd.Cell(i, 6).Text
        HannaCodes(i).LotNumber = Grd.Cell(i, 10).Text

  
    Next
    
End With
ERR_END:
   On Error GoTo 0
   GetHannaCodePerGrid = rc
   Exit Function
ERR_CHEM:
   rc = False
   MsgBox err.Description
   Resume Next
    


End Function






Public Function CopyUserChCodeInGrid(ByVal GridChemicals As Grid, ByVal UserChCode As String, ByRef uRecipe As RecipeType) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim ChemicalCount As Integer
    On Error GoTo ERR_CHEM:
    
    
    If UserChCode <> "" Then
        
        
    
        With GridChemicals
        
            ' check Code
            For t = 1 To .Rows - 1
                
                If Trim(.Cell(t, 1).Text) = Trim(UserChCode) Then
                    If F_MsgBox.DoShow("Chemical RM already in Recipe. Keep or Discard?", UserChCode, , "Keep", "Discard") Then
                        Exit For
                    Else
                        Exit Function
                    End If
                End If
               
           Next
           
        
        
        
        
            ChemicalCount = .Rows - 1
            ReDim Preserve uRecipe.RmxRecipe(ChemicalCount)
            
            Call GetChCodexRecipe(UserChCode, uRecipe.RmxRecipe(ChemicalCount))
            
             uRecipe.RmxRecipe(ChemicalCount).Um = uRecipe.UmMultiple
             
            .AddItem "", False
            .Cell(.Rows - 1, 1).Text = uRecipe.RmxRecipe(ChemicalCount).CHCode
            .Cell(.Rows - 1, 2).Text = uRecipe.RmxRecipe(ChemicalCount).Description
            .Cell(.Rows - 1, 3).Text = uRecipe.RmxRecipe(ChemicalCount).Cas
            .Cell(.Rows - 1, 4).Text = ""
            .Cell(.Rows - 1, 5).Text = SetUmComponent(uRecipe.RmxRecipe(ChemicalCount).Um)
            .Cell(.Rows - 1, 6).Text = ""
            .Cell(.Rows - 1, 7).Text = uRecipe.RmxRecipe(ChemicalCount).TolerancePerc
            .Cell(.Rows - 1, 8).Text = ""
            .Cell(.Rows - 1, 9).Text = uRecipe.RmxRecipe(ChemicalCount).bMix
            .Cell(.Rows - 1, 10).Text = uRecipe.RmxRecipe(ChemicalCount).CriticalRM

        
            For t = 1 To .Rows - 1
                If .Cell(t, 9).Text = True Then
                    .Cell(t, 1).FontBold = True
                    .Cell(t, 2).FontBold = True
                    .Cell(t, 3).FontBold = True
                    .Cell(t, 1).ForeColor = &H886010
                    .Cell(t, 2).ForeColor = &H886010
                    .Cell(t, 3).ForeColor = &H886010
                End If
                
                For i = 1 To .Cols - 1
                    .Column(i).Alignment = cellLeftCenter
                    .Cell(t, 7).BackColor = &HE0E0E0  ' vbColorResults
                    .Cell(t, 4).BackColor = &HE0E0E0  ' vbColorResults
                    .Cell(t, 6).BackColor = &HE0E0E0  ' vbColorResults
                    .Cell(t, 5).BackColor = &HE0E0E0  ' vbColorResults
                Next
               
           Next
            .ReadOnly = False
            .Range(0, 4, 0, 5).Merge
            .Cell(0, 4).Alignment = cellCenterCenter
       
            .Column(3).Alignment = cellCenterCenter
            .Column(4).Alignment = cellRightCenter
            .Column(6).Alignment = cellCenterCenter
            .ReadOnly = True
            .AutoRedraw = True
            .Refresh
            
        End With
        
    End If
    
    
    
ERR_END:
   On Error GoTo 0
   CopyUserChCodeInGrid = rc
   Exit Function
ERR_CHEM:
   rc = False
   MsgBox err.Description
   Resume Next

End Function



Public Function CopyUserHannaCodeInGrid(ByVal Grid1 As Grid, ByVal UserHannaCode As String, ByRef uRecipe As RecipeType) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim CodeCount As Integer
    On Error GoTo ERR_CHEM:
    
    
    If UserHannaCode <> "" Then
        
        
    
        With Grid1
        
            ' check Code
            For t = 1 To .Rows - 1
                
                If Trim(.Cell(t, 1).Text) = Trim(UserHannaCode) Then
                    If F_MsgBox.DoShow("Hanna Code Already exsists in Table. Keep or Discard?", UserHannaCode, , "Keep", "Discard") Then
                        Exit For
                    Else
                        Exit Function
                    End If
                End If
               
           Next
           
        
        
        
        
            CodeCount = .Rows - 1
            ReDim Preserve uRecipe.HannaCodes(CodeCount)
            
            Call GetHannaCodexRecipe(UserHannaCode, uRecipe.HannaCodes(CodeCount), uRecipe.Code)
            
             uRecipe.HannaCodes(CodeCount).Um = uRecipe.UmMultiple
             
            .AddItem "", False
            .Cell(.Rows - 1, 1).Text = uRecipe.HannaCodes(CodeCount).Code
            .Cell(.Rows - 1, 2).Text = uRecipe.HannaCodes(CodeCount).ProductName
            .Cell(.Rows - 1, 3).Text = uRecipe.HannaCodes(CodeCount).Line
            .Cell(.Rows - 1, 4).Text = uRecipe.HannaCodes(CodeCount).Qty
            .Cell(.Rows - 1, 5).Text = uRecipe.HannaCodes(CodeCount).Um
          ' la colonna 6 è per la RecipeForProduction
            .Cell(.Rows - 1, 7).Text = uRecipe.HannaCodes(CodeCount).Recipe
            .Cell(.Rows - 1, 8).Text = uRecipe.HannaCodes(CodeCount).Mix1
            .Cell(.Rows - 1, 9).Text = uRecipe.HannaCodes(CodeCount).Mix2
            

        
            For t = 1 To .Rows - 1
                For i = 1 To .Cols - 1
                    .Column(i).Alignment = cellLeftCenter
                Next
               
           Next
            .Column(3).Alignment = cellCenterCenter
            .Column(4).Alignment = cellRightCenter
            .Column(6).Alignment = cellCenterCenter
             .Column(8).AutoFit
             .Column(9).AutoFit
            .ReadOnly = True
            .AutoRedraw = True
            .Refresh
            
        End With
        
    End If
    
    
    
ERR_END:
   On Error GoTo 0
   CopyUserHannaCodeInGrid = rc
   Exit Function
ERR_CHEM:
   rc = False
   MsgBox err.Description
   Resume Next


End Function

Public Function GetChCodexRecipe(ByVal CHCode As String, ByRef uRMxRecipe As RmxRecipe)
    
    'Set uRmxRecipe = RmxRecipeClean
    
    With dbTabRawMaterial
        .filter = ""
        .filter = "Code='" & CHCode & "'"
        If .EOF Then
        Else
            .MoveFirst
            uRMxRecipe.ID = !ID
            uRMxRecipe.CHCode = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
            uRMxRecipe.Description = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
            uRMxRecipe.Cas = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
            uRMxRecipe.Um = IIf(IsNull(Trim(!Um)), "g", Trim(!Um))
            uRMxRecipe.bMix = !bMix
        End If
    End With
End Function

Public Function GetHannaCodexRecipe(ByVal Code As String, ByRef uHannaCode As HannaCode, ByVal NewRecipe As String)
    
    'Set uHannaCode = RmxRecipeClean
    
    With dbTabCode
        .filter = ""
        .filter = "Code='" & Code & "'"
        If .EOF Then
        Else
            .MoveFirst
            uHannaCode.ID = !ID
            uHannaCode.Code = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
            uHannaCode.ProductName = IIf(IsNull(Trim(!ProductName)), "", Trim(!ProductName))
            uHannaCode.Line = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
            uHannaCode.Qty = IIf(IsNull(Trim(!Qty)), 0, Trim(!Qty))
            uHannaCode.Um = IIf(IsNull(Trim(!Um)), "g", Trim(!Um))
            
            uHannaCode.Recipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
            If Len(Trim(uHannaCode.Recipe)) > 0 Then
                
                uHannaCode.Recipe = uHannaCode.Recipe & IIf(NewRecipe <> "", ";" & NewRecipe, "")
            Else
                uHannaCode.Recipe = NewRecipe
            End If
            
            !Recipe = uHannaCode.Recipe
            .Update
            uHannaCode.Mix1 = IIf(IsNull(Trim(!Mix1)), "", Trim(!Mix1))
            uHannaCode.Mix2 = IIf(IsNull(Trim(!Mix2)), "", Trim(!Mix2))

        End If
    End With
End Function


Public Function DeleteRecipePerCode(ByVal Code As String, ByVal NewRecipe As String)

Dim strRecipe As String
Dim strMix1 As String
Dim strMix2 As String
Dim Quanti As Integer
Dim Recipes() As String

With dbTabCode
        .filter = ""
        .filter = "Code='" & Code & "'"
        If .EOF Then
        Else
            .MoveFirst
            
            strRecipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
            strMix1 = IIf(IsNull(Trim(!Mix1)), "", Trim(!Mix1))
            strMix2 = IIf(IsNull(Trim(!Mix2)), "", Trim(!Mix2))
            

            If Len(strRecipe) > 0 Then
                
                Call SplitTextStringClassification("", strRecipe, Recipes(), Quanti)
            
                If Quanti > 0 Then
                    !Recipe = SetNeRecipesString(Recipes(), NewRecipe)
                End If
                
            
            End If
                    

             If Len(strMix1) > 0 Then
                
                Call SplitTextStringClassification("", strMix1, Recipes(), Quanti)
            
                If Quanti > 0 Then
                    !Mix1 = SetNeRecipesString(Recipes(), NewRecipe)
                End If
                
            
            End If
                               
            If Len(strMix1) > 0 Then
                
                Call SplitTextStringClassification("", strMix1, Recipes(), Quanti)
            
                If Quanti > 0 Then
                    !Mix2 = SetNeRecipesString(Recipes(), NewRecipe)
                End If
                
            
            End If
                                                   
            
            .Update
        End If
    End With
End Function

Private Function SetNeRecipesString(ByVal vetRecipes As Variant, ByVal RecipeCode As String) As String
Dim i As Integer
    For i = LBound(vetRecipes) To UBound(vetRecipes)
        If InStr(vetRecipes(i), RecipeCode) Then
        Else
            If vetRecipes(i) <> "" Then
            SetNeRecipesString = SetNeRecipesString & IIf(SetNeRecipesString = "", "", ";") & vetRecipes(i)
            End If
        End If
    Next
End Function

Public Function SetListOfString(ByVal resultString As String, ByVal NewItem As String) As String
Dim i As Integer
        If InStr(resultString, NewItem) Then
            SetListOfString = resultString
        Else
            SetListOfString = resultString & IIf(resultString = "", "", ";") & NewItem
        End If
End Function

Public Function SetNeHannaCodeQtyString(ByRef vetCode() As HannaCode) As String
Dim i As Integer
On Error GoTo ERR_SET:
    For i = LBound(vetCode) To UBound(vetCode)
        If vetCode(i).Code <> "" And vetCode(i).QtyToProduce <> "" Then
            SetNeHannaCodeQtyString = SetNeHannaCodeQtyString & IIf(SetNeHannaCodeQtyString = "", "", "  |  ") & vetCode(i).Code & " : " & vetCode(i).QtyToProduce
        End If
        
    Next
ERR_END:
    Exit Function
ERR_SET:
    Resume ERR_END
End Function

Public Function CheckPercentageByWeight(ByVal Grd As Grid, ByRef Perc As String, ByVal TotaleInGrammi As Double) As Boolean
Dim i As Integer
Dim singlePerc() As String
Dim CellValue As String
Dim MyUM As String


Dim rc As Boolean
Perc = 0
rc = False


    With Grd
        If .Rows > 1 Then
            
            MyUM = IIf(.Cell(1, 5).Text = "", "g", .Cell(1, 5).Text)

            If TotaleInGrammi = 0 Then Exit Function
            ReDim singlePerc(.Rows - 1)
            
            For i = 1 To .Rows - 1
            
                CellValue = .Cell(i, 4).Text
                If CellValue = "" Then Exit Function
    
            
                singlePerc(i) = FormatNumber(((CDbl(CellValue) * Um(MyUM)) / TotaleInGrammi) * 100, 3)
                .Cell(i, 6).Text = singlePerc(i)
                Perc = CDbl(Perc) + CDbl(singlePerc(i))
            Next
         
        End If
    End With
    
    rc = IIf(CDbl(Perc) = 100, True, False)
    CheckPercentageByWeight = rc


End Function


Public Function CheckChemicalsInRecipe(ByVal Grd As Grid) As Boolean
Dim rc As Boolean
Dim i As Integer
rc = False


    With Grd
        If .Rows > 1 Then
         
            For i = 1 To .Rows - 1
             If IsNumeric(.Cell(i, 4).Text) Then
                rc = True
             Else
                rc = False
                Exit For
             End If
            Next
        Else
            rc = True
        End If
    End With

    CheckChemicalsInRecipe = rc

End Function


Public Function SaveFullRecipe(ByVal Grd As Grid, ByVal RecipeCode As String) As Boolean

Dim rc As Boolean
Dim i As Integer
Dim CHCode As String

rc = False


    With dbTabRecipe
        .filter = ""
        .filter = "code='" & RecipeCode & "'"
        If .EOF Then
        Else
            !Mix = ""
        End If
    End With
    
    

    With Grd
        If .Rows > 1 Then
            DeleteRecipeComponentByCode RecipeCode
            
            For i = 1 To .Rows - 1
                CHCode = Trim(.Cell(i, 1).Text)
try:
                With dbTabRMxRecipe
                    .Close
                    .Open "SELECT *  FROM TabRMxRecipe order by id ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
                    .filter = ""
                    '.filter = "RecipeCode='" & RecipeCode & "' and CHCode='" & CHCode & "'"
                    'If .EOF Then
                        .AddNew
                   ' Else
                    'End If
                    !RecipeCode = RecipeCode
                    !CHCode = CHCode
                    !Description = Grd.Cell(i, 2).Text
                    !Cas = Grd.Cell(i, 3).Text
                    !Qty = Grd.Cell(i, 4).Text
                    !Um = Grd.Cell(i, 5).Text
                    !Perc = Grd.Cell(i, 6).Text
                    !TolerancePerc = Grd.Cell(i, 7).Text
                    !Note = Grd.Cell(i, 8).Text
                    !bMix = Grd.Cell(i, 9).Text
                    If !bMix Then
                        ' add Mix to Recipe...
                        Call AddMixtoRecipe(CHCode, RecipeCode)
                    End If
                    .Update
                    rc = True
                End With
             
             
            Next
        Else
            rc = True
        End If
    End With

    SaveFullRecipe = rc


End Function

Public Function DeleteRecipeComponentByCode(ByVal RecipeCode As String) As Boolean
Dim rc As Boolean
On Error GoTo ERR_DELETE
rc = True

    With dbTabRMxRecipe
         .Close
         .Open "SELECT *  FROM TabRMxRecipe order by id ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
     End With
            
    dbCode.Execute ("DELETE * FROM TabRMxRecipe WHERE RecipeCode='" & RecipeCode & "'")
    
ERR_END:
    On Error GoTo 0
    
    DeleteRecipeComponentByCode = rc
    Exit Function
ERR_DELETE:
    rc = False
    MsgBox err.Description
    Resume Next
    
End Function

Public Function AddMixtoRecipe(ByVal CHCode As String, ByVal RecipeCode As String) As Boolean
Dim rc As Boolean
Dim RecipeMix As String


    With dbTabRecipe
        .filter = ""
        .filter = "Code='" & RecipeCode & "'"
        If .EOF Then
        
        Else
            RecipeMix = IIf(IsNull(Trim(!Mix)), "", Trim(!Mix))

            If InStr(RecipeMix, CHCode) Then
                Exit Function
            ElseIf RecipeMix <> "" Then
                !Mix = !Mix & ";" & CHCode
            Else
                !Mix = CHCode
            End If
            
        End If
    
    End With
End Function

Public Function IfAllMixes(ByVal RecipeCode As String, Optional ByRef MixesCount As Integer) As Boolean
Dim rc As Boolean
Dim i As Integer
    MixesCount = 0
    With dbTabRMxRecipe
        .filter = ""
        .filter = "RecipeCode='" & RecipeCode & "'"
        If .EOF Then
            rc = False
        Else
        
            .MoveFirst
            rc = True
            For i = 1 To .RecordCount
                If !bMix = False Or IsNull(!bMix) Then
                    rc = False
                    Exit For
                End If
                .MoveNext
            Next
        End If
    End With
    IfAllMixes = rc
    MixesCount = i
End Function
Public Function GetChemicalCAS(ByVal strName As String) As String
If strName <> "" Then
    With dbTabRawMaterial
        .filter = ""
        .filter = "Code='" & strName & "'"
        If .EOF Then
            
            With dbTabRawMaterial
                .filter = ""
                .filter = "Code='" & strName & "'"
                If .EOF Then
                Else
                    GetChemicalCAS = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
                End If
            End With
            
        Else
            GetChemicalCAS = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
        End If
    End With
End If
End Function
Public Function GetChemicalDescription(ByVal strName As String) As String
If strName <> "" Then
    With dbTabRawMaterial
        .filter = ""
        .filter = "Code='" & strName & "'"
        If .EOF Then
            
            With dbTabRawMaterial
                .filter = ""
                .filter = "Code='" & strName & "'"
                If .EOF Then
                Else
                    GetChemicalDescription = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
                End If
            End With
            
        Else
            GetChemicalDescription = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
        End If
    End With
End If
End Function
