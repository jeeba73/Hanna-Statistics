Attribute VB_Name = "ChemicalProductionSetTypes"
Option Explicit



Public Function GetRecipeCodeByMix(ByRef uRecipe() As RecipeType, ByVal strMix As String, ByRef IndexRecipe As Integer, ByRef strRecipe As String)
Dim i As Integer
Dim t As Integer
    For i = 1 To UBound(uRecipe)
        If uRecipe(i).bHaveMixes Then
            For t = 0 To UBound(uRecipe(i).RmxRecipe)
                If uRecipe(i).RmxRecipe(t).bMix Then
                    If uRecipe(i).RmxRecipe(t).CHCode = strMix Then
                        IndexRecipe = i
                        strRecipe = uRecipe(i).Code
                        Exit Function
                    End If
                End If
            Next
        End If
    Next
End Function



Public Function SetMyRecipeByCode(ByVal RecipeCode As String, ByRef uRecipe As RecipeType)


On Error GoTo ERR_SET:
    If RecipeCode = "" Then Exit Function
    Dim i As Integer
    
    uRecipe.Code = RecipeCode
    uRecipe.Description = "Not Found in Database!"
     With dbTabRecipe
            
        .filter = ""
        .filter = "Code='" & RecipeCode & "'"
        If .EOF Then
            uRecipe.bUpdated = False
            Exit Function
        End If
        .MoveFirst
        uRecipe.Code = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
        uRecipe.Description = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
        uRecipe.Line = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
        uRecipe.Procedure = IIf(IsNull(Trim(!Procedure)), "", Trim(!Procedure))
        uRecipe.Rev = IIf(IsNull(Trim(!Rev)), "", Trim(!Rev))
        uRecipe.NoteRev = IIf(IsNull(Trim(!NoteRev)), "", Trim(!NoteRev))
        uRecipe.Exp = CheckDot(IIf(IsNull(Trim(!Exp)), 0, Trim(!Exp)))
        uRecipe.Density = CheckDot(IIf(IsNull(Trim(!Density)), 1, Trim(!Density)))
        uRecipe.MaxQty = CheckDot(IIf(IsNull(Trim(!MaxQty)), 0, Trim(!MaxQty)))
        uRecipe.UmMax = IIf(IsNull(Trim(!UmMax)), "g", Trim(!UmMax))
        uRecipe.MinQty = CheckDot(IIf(IsNull(Trim(!MinQty)), 0, Trim(!MinQty)))
        uRecipe.Multiple = CheckDot(IIf(IsNull(Trim(!Multiple)), 0, Trim(!Multiple)))
        uRecipe.UmMultiple = IIf(IsNull(Trim(!UmMultiple)), "g", Trim(!UmMultiple))
        uRecipe.MinQty2 = CheckDot(IIf(IsNull(Trim(!MinQty2)), 0, Trim(!MinQty2)))
        uRecipe.UmMinQty = IIf(IsNull(Trim(!UmMinQty)), "", Trim(!UmMinQty))
        uRecipe.Mix = IIf(IsNull(Trim(!Mix)), "", Trim(!Mix))
        uRecipe.bIsMix = IfRecipeIsMix(uRecipe)
        uRecipe.bUmMassa = SetbUmMassa(uRecipe.UmMultiple)
        uRecipe.bUpdated = True
        uRecipe.bNoPreparation = !bNoPreparation
        uRecipe.RevDate = IIf(IsNull(Trim(!RevDate)), "", Trim(!RevDate))
    End With

  

     uRecipe.bUpdated = SetRmxRecipeByRecipeCode(RecipeCode, uRecipe.RmxRecipe(), False)
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_SET:
    MsgBox err.Description
    Resume Next

End Function

Public Function IfRecipeIsMix(ByRef uRecipe As RecipeType)
With dbTabRawMaterial
    .filter = ""
    .filter = "Code='" & uRecipe.Code & "'"
    If .EOF Then
        IfRecipeIsMix = False
    Else
        IfRecipeIsMix = !bMix
        If !bMix Then
            uRecipe.Cas = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
            uRecipe.SpecifiedLocation = IIf(IsNull(Trim(!SpecifiedLocation)), "", Trim(!SpecifiedLocation))
            uRecipe.Location = IIf(IsNull(Trim(!Location)), "", Trim(!Location))
        End If
    End If
End With
End Function
Public Function IfRecipeIsMixString(ByVal uRecipe As String) As Boolean
IfRecipeIsMixString = False
With dbTabRawMaterial
    .filter = ""
    .filter = "Code='" & uRecipe & "'"
    If .EOF Then
        IfRecipeIsMixString = False
    Else
        IfRecipeIsMixString = !bMix
    End If
End With
End Function
Public Function IfRecipeNoPreparation(ByVal uRecipe As String) As Boolean
IfRecipeNoPreparation = False
With dbTabRecipe
    .filter = ""
    .filter = "Code='" & uRecipe & "'"
    If .EOF Then
        IfRecipeNoPreparation = False
    Else
        IfRecipeNoPreparation = !bNoPreparation
    End If
End With
End Function

Public Function IfRecipeHasMixes(ByVal uRecipe As String) As Boolean
IfRecipeHasMixes = False
With dbTabRecipe
    .filter = ""
    .filter = "Code='" & uRecipe & "'"
    If .EOF Then
        IfRecipeHasMixes = False
    Else
        IfRecipeHasMixes = IIf(IsNull(Trim(!Mix)) Or Trim(!Mix) = "", False, True)
    End If
End With
End Function





Public Function IfRecipeExsists(ByVal RecipeCode As String) As Boolean
Dim rc As Boolean

    With dbTabRecipe
        rc = True
        .filter = ""
        .filter = "Code='" & RecipeCode & "'"
        If .EOF Then rc = False
    End With
    
    IfRecipeExsists = rc
    
End Function


Public Function IfRecipeNotInGrid2(ByVal Recipe As String, ByVal Grid2 As Grid) As Boolean
Dim rc As Boolean
Dim i As Integer

    rc = True
    With Grid2
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                
                If Trim(UCase(.Cell(i, 1).Text)) = Trim(UCase(Recipe)) Then
                    rc = False
                    GoTo ERR_END
                End If
            
            
            Next
        End If
    End With
ERR_END:
    IfRecipeNotInGrid2 = rc
End Function

Public Function SetRmxRecipeByRecipeCode(ByVal RecipeCode As String, ByRef uRMxRecipe() As RmxRecipe, Optional ByVal bValue As Boolean = True, Optional ByVal UltimoMixese As Integer, Optional RmxRecipeCount As Integer) As Boolean

Dim Mixes() As RmxRecipe
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim RecipeRecordCount As Integer

On Error GoTo ERR_SET:

     If RecipeCode = "" Then Exit Function
     
     RecipeRecordCount = 0
     If UltimoMixese > 0 Then
        RecipeRecordCount = UltimoMixese + 1
     Else
        ReDim uRMxRecipe(RecipeRecordCount)
    End If
GetMixes:
     
    ' bValue = IfAllMixes(RecipeCode)
     With dbTabRMxRecipe
            
        .filter = ""
        .filter = "RecipeCode='" & RecipeCode & "'"
        If .EOF Then
            SetRmxRecipeByRecipeCode = False
            Exit Function
        End If
        .MoveFirst
        ReDim Preserve uRMxRecipe(RecipeRecordCount + .RecordCount - 1)
        
        t = 0
        For i = RecipeRecordCount To RecipeRecordCount + .RecordCount - 1
            uRMxRecipe(i).CHCode = IIf(IsNull(Trim(!CHCode)), "", Trim(!CHCode))
            uRMxRecipe(i).RecipeCode = RecipeCode
            uRMxRecipe(i).Description = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
            uRMxRecipe(i).Cas = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
            uRMxRecipe(i).Qty = CheckDot(IIf(IsNull(Trim(!Qty)), "", Trim(!Qty)))
            uRMxRecipe(i).Um = IIf(IsNull(Trim(!Um)), "", Trim(!Um))
            uRMxRecipe(i).Perc = CheckDot(IIf(IsNull(Trim(!Perc)), "", Trim(!Perc)))
            uRMxRecipe(i).TolerancePerc = CheckDot(IIf(IsNull(Trim(!TolerancePerc)), 1, Trim(!TolerancePerc)))
            uRMxRecipe(i).Note = IIf(IsNull(Trim(!Note)), "", Trim(!Note))
            uRMxRecipe(i).bMix = !bMix
            
            If uRMxRecipe(i).TolerancePerc = 0 Then uRMxRecipe(i).TolerancePerc = 1
            
            If !bMix Then
                t = t + 1
                ReDim Preserve Mixes(t)
                Mixes(t).CHCode = uRMxRecipe(i).CHCode
                
                Call SetMyRecipeMixByCode(uRMxRecipe(i).CHCode, uRMxRecipe(i))
                
            
            End If
            
            With dbTabRawMaterial
                .filter = ""
                .filter = "Code='" & uRMxRecipe(i).CHCode & "'"
                If .EOF Then
                Else
                    'uRMxRecipe(i).Specifications.ChemicalReactionLiquid = IIf(IsNull(Trim(!ChemicalReactionLiquid)), "", Trim(!ChemicalReactionLiquid))
                    'uRMxRecipe(i).Specifications.Classification = IIf(IsNull(Trim(!Classification)), "", Trim(!Classification))
                    'uRMxRecipe(i).Specifications.Pictograms = IIf(IsNull(Trim(!Pictograms)), "", Trim(!Pictograms))
                    'uRMxRecipe(i).Specifications.ManufacturerName = IIf(IsNull(Trim(!ManufacturerName)), "", Trim(!ManufacturerName))
                    'uRMxRecipe(i).Specifications.ManufacturerCode = IIf(IsNull(Trim(!ManufacturerCode)), "", Trim(!ManufacturerCode))
                    uRMxRecipe(i).Specifications.Location = IIf(IsNull(Trim(!Location)), "", Trim(!Location))
                    uRMxRecipe(i).Specifications.SpecifiedLocation = IIf(IsNull(Trim(!SpecifiedLocation)), "", Trim(!SpecifiedLocation))
                    uRMxRecipe(i).CriticalRM = IIf(IsNull(Trim(!CriticalRM)), "", Trim(!CriticalRM))
                End If
           End With
            
            .MoveNext
        Next

        RecipeRecordCount = RecipeRecordCount + .RecordCount - 1
    End With
    Dim X As Integer
    If t > 0 Then
        ' ho dei mix per cui prendo i Mixes
        For X = 1 To t
            
            RecipeCode = Mixes(X).CHCode
            Call SetComponentsInRecipesMix(RecipeCode, uRMxRecipe(), , UBound(uRMxRecipe()))
        Next
    
    End If
    
ERR_END:
    On Error GoTo 0
     RmxRecipeCount = RecipeRecordCount
    SetRmxRecipeByRecipeCode = True
    Exit Function
ERR_SET:
    MsgBox err.Description
    Resume Next

End Function
Public Function SetMyRecipeMixByCode(ByVal RecipeCode As String, ByRef uRMxMix As RmxRecipe)
On Error GoTo ERR_SET:
    If RecipeCode = "" Then Exit Function
    Dim i As Integer
     
     With dbTabRecipe
            
        .filter = ""
        .filter = "Code='" & RecipeCode & "'"
        If .EOF Then Exit Function
        .MoveFirst
        
        uRMxMix.Density = CheckDot(IIf(IsNull(Trim(!Density)), 1, Trim(!Density)))
        uRMxMix.MaxQty = CheckDot(IIf(IsNull(Trim(!MaxQty)), 0, Trim(!MaxQty)))
        uRMxMix.UmMax = IIf(IsNull(Trim(!UmMax)), "g", Trim(!UmMax))
        uRMxMix.MinQty = CheckDot(IIf(IsNull(Trim(!MinQty)), 0, Trim(!MinQty)))
        uRMxMix.Multiple = CheckDot(CheckDot(IIf(IsNull(Trim(!Multiple)), 0, Trim(!Multiple))))
        uRMxMix.UmMultiple = IIf(IsNull(Trim(!UmMultiple)), "g", Trim(!UmMultiple))
        uRMxMix.MinQty2 = CheckDot(IIf(IsNull(Trim(!MinQty2)), 0, Trim(!MinQty2)))
        uRMxMix.UmMinQty = IIf(IsNull(Trim(!UmMinQty)), "", Trim(!UmMinQty))
        uRMxMix.bUmMassa = SetbUmMassa(uRMxMix.UmMultiple)
      
    End With
ERR_END:
    On Error GoTo 0
    
    Exit Function
ERR_SET:
    MsgBox err.Description
    Resume Next
  
    
    
End Function





Public Function SetComponentsInRecipesMix(ByVal RecipeCode As String, ByRef uRMxRecipe() As RmxRecipe, Optional ByVal bValue As Boolean = True, Optional ByVal UltimoComponente As Integer) As Boolean

Dim Component() As RmxRecipe
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim RecipeRecordCount As Integer

     If RecipeCode = "" Then Exit Function
     
     RecipeRecordCount = 0
     If UltimoComponente > 0 Then
        RecipeRecordCount = UltimoComponente + 1
     Else
        ReDim uRMxRecipe(RecipeRecordCount)
    End If


     With dbTabRMxRecipe
            
        .filter = ""
        .filter = "RecipeCode='" & RecipeCode & "'"
        If .EOF Then
            SetComponentsInRecipesMix = False
            Exit Function
        End If
        .MoveFirst
        ReDim Preserve uRMxRecipe(RecipeRecordCount + .RecordCount - 1)
        
        t = 0
        For i = RecipeRecordCount To RecipeRecordCount + .RecordCount - 1
            uRMxRecipe(i).CHCode = IIf(IsNull(Trim(!CHCode)), "", Trim(!CHCode))
            uRMxRecipe(i).RecipeCode = RecipeCode
            uRMxRecipe(i).Description = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
            uRMxRecipe(i).Cas = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
            uRMxRecipe(i).Qty = CheckDot(IIf(IsNull(Trim(!Qty)), "", Trim(!Qty)))
            uRMxRecipe(i).Um = IIf(IsNull(Trim(!Um)), "", Trim(!Um))
            uRMxRecipe(i).Perc = CheckDot(IIf(IsNull(Trim(!Perc)), "", Trim(!Perc)))
            uRMxRecipe(i).TolerancePerc = CheckDot(IIf(IsNull(Trim(!TolerancePerc)), 1, Trim(!TolerancePerc)))
            uRMxRecipe(i).Note = IIf(IsNull(Trim(!Note)), "", Trim(!Note))
            uRMxRecipe(i).bMix = !bMix
            
            With dbTabRawMaterial
                .filter = ""
                .filter = "Code='" & uRMxRecipe(i).CHCode & "'"
                If .EOF Then
                Else
                    'uRMxRecipe(i).Specifications.ChemicalReactionLiquid = IIf(IsNull(Trim(!ChemicalReactionLiquid)), "", Trim(!ChemicalReactionLiquid))
                    'uRMxRecipe(i).Specifications.Classification = IIf(IsNull(Trim(!Classification)), "", Trim(!Classification))
                    'uRMxRecipe(i).Specifications.Pictograms = IIf(IsNull(Trim(!Pictograms)), "", Trim(!Pictograms))
                    'uRMxRecipe(i).Specifications.ManufacturerName = IIf(IsNull(Trim(!ManufacturerName)), "", Trim(!ManufacturerName))
                    'uRMxRecipe(i).Specifications.ManufacturerCode = IIf(IsNull(Trim(!ManufacturerCode)), "", Trim(!ManufacturerCode))
                    uRMxRecipe(i).Specifications.Location = IIf(IsNull(Trim(!Location)), "", Trim(!Location))
                    uRMxRecipe(i).Specifications.SpecifiedLocation = IIf(IsNull(Trim(!SpecifiedLocation)), "", Trim(!SpecifiedLocation))
                    uRMxRecipe(i).CriticalRM = IIf(IsNull(Trim(!CriticalRM)), "", Trim(!CriticalRM))
                    uRMxRecipe(i).Density = CheckDot(IIf(IsNull(Trim(!Density)), "1", Trim(!Density)))
                    
                End If
           End With
            
            .MoveNext
        Next

        RecipeRecordCount = RecipeRecordCount + .RecordCount - 1
    End With

   
    SetComponentsInRecipesMix = True

End Function























Public Function SetPackagingFromGrid(ByRef Packaging() As ProdWay, ByVal Grid As Grid) As Boolean
 Dim PackagingCount As Integer
    Dim i As Integer
    PackagingCount = Grid.Rows - 1

    
    ReDim Packaging(PackagingCount)
    
     
        '.Cell(0, 1).Text = "Recipe"
        '.Cell(0, 2).Text = "Line"
        '.Cell(0, 3).Text = "STDPreparation Way"
        '.Cell(0, 4).Text = "STDPreparation speed ( pcs/min )"
        '.Cell(0, 5).Text = "Estimated time machine ( h )"
        '.Cell(0, 6).Text = "Estimated Shift machine ( S )"
        
    For i = 1 To PackagingCount
        With Packaging(i)
            .Recipe = Grid.Cell(i, 1).Text
            .Line = Grid.Cell(i, 2).Text
            .STDPreparation = Grid.Cell(i, 3).Text
            .Speed = IIf(Grid.Cell(i, 4).Text = "", 0, Grid.Cell(i, 4).Text)
            
            .EstTimeH = IIf(Grid.Cell(i, 5).Text = "", 0, Grid.Cell(i, 5).Text)
            .EsttimeD = IIf(Grid.Cell(i, 6).Text = "", 0, Grid.Cell(i, 6).Text)
            
        End With
    Next
    
End Function

Public Function SetTotalsFromGrid(ByRef Totals() As Totals, ByVal Grid As Grid) As Boolean

    Dim TotalsCount As Integer
    Dim i As Integer
    On Error GoTo ERR_SET:
    TotalsCount = Grid.Rows - 1

    
    ReDim Totals(TotalsCount)
    
     
       ' .Cell(0, 1).Text = "Recipe"
       ' .Cell(0, 2).Text = "Description"
       ' .Cell(0, 3).Text = "Total Weight (Kg)"
       ' .Cell(0, 4).Text = "Total Volume (L)"
       ' .Cell(0, 5).Text = "Multiple"
       ' .Cell(0, 7).Text = "Min"
       ' .Cell(0, 8).Text = "Max"
       ' .Cell(0, 10).Text = "Min"
       ' .Cell(0, 11).Text = "Max"
       ' .Cell(0, 12).Text = "Min pcs"
       ' .Cell(0, 13).Text = "Multiple"
        
    For i = 1 To TotalsCount
        With Totals(i)
            .Recipe = Grid.Cell(i, 1).Text
            .Description = Grid.Cell(i, 2).Text
            .TotalWeighKg = Grid.Cell(i, 3).Text
            .TotalWeighL = Grid.Cell(i, 4).Text
            .TotalMultiple = Grid.Cell(i, 5).Text
            .CkMin = IIf(Grid.Cell(i, 8).BackColor = &H8000&, True, False)
            .CkMax = IIf(Grid.Cell(i, 11).BackColor = &H8000&, True, False)
            .Min = Grid.Cell(i, 7).Text
            .Max = Grid.Cell(i, 10).Text
            .Minpcs = Grid.Cell(i, 12).Text
            .Multiple = Grid.Cell(i, 13).Text
            .bMix = Grid.Cell(i, 15).Text
        End With
    Next
    
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_SET:
    SetTotalsFromGrid = False
    GoTo ERR_END

End Function


Public Function IfChemicalInRecipeType(ByVal CHCode As String, ByRef iRecipe As RecipeType, ByRef Index As Integer) As Boolean
Dim i As Integer
Dim rc As Boolean
rc = False
Index = 999
    With iRecipe
    
        For i = 0 To .RmxRecipeCount
        
        
            If CHCode = .RmxRecipe(i).CHCode Then
                Index = i
                rc = True
                Exit For
            End If
        Next
        
    End With
    
IfChemicalInRecipeType = rc
End Function
