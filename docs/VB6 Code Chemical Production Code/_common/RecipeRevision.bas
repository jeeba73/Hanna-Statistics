Attribute VB_Name = "RecipeRevision"
Option Explicit

Public Type RevisionHistory

     RevDate     As String
     Recipe      As String
     RevNumber   As String
     RevType     As String
     Description As String
     Operator    As String


End Type


Public Function SaveRecipeDataPerRevision(ByRef iRecipe As RecipeType)
Dim rc As Boolean
Dim Line As String
Dim Code As String
Dim Rev As String
Dim RevDate As String
Dim SettingName As String
Dim ExcelName As String


On Error GoTo ERR_SAVE:

    rc = True
    
    With iRecipe
        Code = .Code
        Line = .Line
        Rev = .Rev
        RevDate = .RevDate
    
        ' creo o verifico la dir di salvataggio
        Call SettSavePath(PathRecipe & Line)
        Call SettSavePath(PathRecipe & "Data")
        
        ' imposto la dir del file
        USER_PATH = PathRecipe & "Data\"
        
        SettingName = FormatNomeFile(Code & ".recipe")
        ExcelName = FormatNomeFile(Code & "_r" & Rev & "." & RevDate & ".xls")
        
        
        Call SaveRecipeInFile(iRecipe, SettingName)
        
        Call EsportaRecipeExcel(SettingName, ExcelName, iRecipe)
     

    End With
ERR_END:
    On Error GoTo 0
    SaveRecipeDataPerRevision = rc
    Exit Function
ERR_SAVE:
    rc = False
    MsgBox err.Description
    Resume Next

End Function


Private Function SaveRecipeInFile(ByRef iRecipe As RecipeType, ByVal SettingName As String)
Dim rc As Boolean
Dim RmxRecipeCount As Integer
Dim RevCount As Integer
Dim OldRev As String

On Error GoTo ERR_SAVE:

    rc = True
    
    
    CloseSettingDataFile
    
    RevCount = GetSettingData(SettingName, "Formulation Revision", "RevCount", 0)
    OldRev = GetSettingData(SettingName, "Revision" & RevCount, "Rev", "")
    
    If iRecipe.Rev <> OldRev Then RevCount = RevCount + 1
    
    With iRecipe
    
        SaveSettingData SettingName, "Formulation Revision", "RevCount", RevCount
        
        SaveSettingData SettingName, "Revision" & RevCount, "Rev", .Rev
        SaveSettingData SettingName, "Revision" & RevCount, "RevDate", .RevDate
        
        
        SaveSettingData SettingName, "Rev" & .Rev, "RevCount", RevCount
        
        
        SaveSettingData SettingName, "Rev" & .Rev, "RevDate", .RevDate
        SaveSettingData SettingName, "Rev" & .Rev, "Code", .Code
        SaveSettingData SettingName, "Rev" & .Rev, "bHaveMixes", .bHaveMixes
        SaveSettingData SettingName, "Rev" & .Rev, "bUmMassa", .bUmMassa
        SaveSettingData SettingName, "Rev" & .Rev, "bUpdated", .bUpdated
        SaveSettingData SettingName, "Rev" & .Rev, "bIsMix", .bIsMix
        SaveSettingData SettingName, "Rev" & .Rev, "Density", .Density
        SaveSettingData SettingName, "Rev" & .Rev, "Description", .Description
        SaveSettingData SettingName, "Rev" & .Rev, "Exp", .Exp
        SaveSettingData SettingName, "Rev" & .Rev, "ID", .ID
        SaveSettingData SettingName, "Rev" & .Rev, "Line", .Line
        
        SaveSettingData SettingName, "Rev" & .Rev, "MaxQty", .MaxQty
        SaveSettingData SettingName, "Rev" & .Rev, "MinQty", .MinQty
        SaveSettingData SettingName, "Rev" & .Rev, "MinQty2", .MinQty2
        SaveSettingData SettingName, "Rev" & .Rev, "Classification", .Classification
        SaveSettingData SettingName, "Rev" & .Rev, "Mix", .Mix
        
        SaveSettingData SettingName, "Rev" & .Rev, "bNoPreparation", .bNoPreparation

        SaveSettingData SettingName, "Rev" & .Rev, "Multiple", .Multiple
        SaveSettingData SettingName, "Rev" & .Rev, "MultipleMassa", .MultipleMassa
        SaveSettingData SettingName, "Rev" & .Rev, "MultipleToProduce", .MultipleToProduce
        SaveSettingData SettingName, "Rev" & .Rev, "NoteRev", .NoteRev
        SaveSettingData SettingName, "Rev" & .Rev, "Procedure", .Procedure
        SaveSettingData SettingName, "Rev" & .Rev, "ProcedureDate", .ProcedureDate
        SaveSettingData SettingName, "Rev" & .Rev, "ProductionWay.EsttimeD", .ProductionWay.EsttimeD
        SaveSettingData SettingName, "Rev" & .Rev, "ProductionWay.EstTimeH", .ProductionWay.EstTimeH
        SaveSettingData SettingName, "Rev" & .Rev, "ProductionWay.Line", .ProductionWay.Line
        SaveSettingData SettingName, "Rev" & .Rev, "ProductionWay.Production", .ProductionWay.Production
        SaveSettingData SettingName, "Rev" & .Rev, "ProductionWay.Speed", .ProductionWay.Speed
        
        SaveSettingData SettingName, "Rev" & .Rev, "Rev", .Rev
        
        SaveSettingData SettingName, "Rev" & .Rev, "TotalMultiple", .TotalMultiple
        SaveSettingData SettingName, "Rev" & .Rev, "TotalRecipe", .TotalRecipe
        SaveSettingData SettingName, "Rev" & .Rev, "TotalWeightKg", .TotalWeightKg
        SaveSettingData SettingName, "Rev" & .Rev, "TotalWeightL", .TotalWeightL
        SaveSettingData SettingName, "Rev" & .Rev, "UmMax", .UmMax
        SaveSettingData SettingName, "Rev" & .Rev, "UmMinQty", .UmMinQty
        SaveSettingData SettingName, "Rev" & .Rev, "bHide", .bHide
        SaveSettingData SettingName, "Rev" & .Rev, "UmMultiple", .UmMultiple
        SaveSettingData SettingName, "Rev" & .Rev, "UmTotalWeightKg", .UmTotalWeightKg
        SaveSettingData SettingName, "Rev" & .Rev, "UmTotalWeightL", .UmTotalWeightL

        '-----------------------------------------------------------
        ' RmxRecipe
        '-----------------------------------------------------------
        If .RmxRecipeCount >= 0 Then
            RmxRecipeCount = .RmxRecipeCount
            SaveSettingData SettingName, "Rev" & .Rev & " - RmxRecipe", "RmxRecipeCount", .RmxRecipeCount
            Call SetRevRmxRecipeInFile(.Rev, .RmxRecipe, RmxRecipeCount, SettingName)
        End If
    End With


CloseSettingDataFile


ERR_END:
    On Error GoTo 0
    SaveRecipeInFile = rc
    Exit Function
ERR_SAVE:
    rc = False
    MsgBox err.Description
    Resume Next

End Function
Public Function SetRevRmxRecipeInFile(ByVal Rev As String, ByRef RmxRecipe() As RmxRecipe, ByVal RmxRecipeCount As Integer, ByVal SettingName As String)
Dim i As Integer

    If RmxRecipeCount = 0 Then
        Exit Function
    End If
    
    For i = 0 To RmxRecipeCount
        
        With RmxRecipe(i)
        
            
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "bMix", .bMix
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "bUmMassa", .bUmMassa
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Cas", .Cas
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "CHCode", .CHCode
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Density", .Density
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Description", .Description
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "ID", .ID
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "MaxQty", .MaxQty
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "MinQty", .MinQty
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "MinQty2", .MinQty2
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Multiple", .Multiple
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "MultipleInCell", .MultipleInCell
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "MultipleMassa", .MultipleMassa
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "MultipleToProduce", .MultipleToProduce
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Note", .Note
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Perc", .Perc
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Qty", .Qty
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "CriticalRM", .CriticalRM
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Specifications.ChemicalReactionLiquid", .Specifications.ChemicalReactionLiquid
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Specifications.Classification", .Specifications.Classification
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Specifications.Location", .Specifications.Location
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Specifications.ManufacturerCode", .Specifications.ManufacturerCode
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Specifications.ManufacturerName", .Specifications.ManufacturerName
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Specifications.Pictograms", .Specifications.Pictograms
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Specifications.QtyToProduce", .Specifications.QtyToProduce
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Specifications.SpecifiedLocation", .Specifications.SpecifiedLocation
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Specifications.Tollerance", .Specifications.Tollerance
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Specifications.UmQty", .Specifications.UmQty
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "RecipeCode", .RecipeCode
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "TolerancePerc", .TolerancePerc
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "TheoreticalWeight", .TheoreticalWeight
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "TotalMultiple", .TotalMultiple
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "TotalWeightKg", .TotalWeightKg
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "TotalWeightL", .TotalWeightL
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "Um", .Um
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "UmMax", .UmMax
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "UmMinQty", .UmMinQty
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "UmMultiple", .UmMultiple
            SaveSettingData SettingName, "Rev" & Rev & " - RmxRecipe" & i, "UmTheoreticalWeight", .UmTheoreticalWeight

        End With

    Next
    
    CloseSettingDataFile
End Function

