Attribute VB_Name = "mod_Excel_Recipes"
Option Explicit



Private uRecipe As RecipeType



Public Function EsportaRecipeExcel(ByVal FileName As String, ByVal sString As String, ByRef iRecipe As RecipeType) As Boolean
    Dim rc As Boolean
    Dim Path As String
    Dim Line As String
    
    On Error GoTo ERR_IMP
    rc = True
   ' MsgBox USER_DESKTOP & "\" & "Backup VerPeriodica.xls"
   uRecipe = iRecipe
   SettingName = FileName
   Line = uRecipe.Line
   If SettingName = "" Then Exit Function
   
   If Line = "" Then
   Else

        If CreateExcel(False) Then
            NewExcelWorksheet (sString)
            Call SettSavePath(PathRecipe & Line)
            Path = PathRecipe & Line
            Debug.Print Path
            If CopyRecipeData(SettingName) Then
                Call SaveExcel(sString, Path)
                Call CloseExcel
                PopupMessage 2, "Excel file correctly generated...", , , iRecipe.Code
            Else
                rc = False
            End If
        Else
            rc = False
        End If
        
        
    End If
ERR_END:
    On Error GoTo 0
    EsportaRecipeExcel = rc
    Exit Function
ERR_IMP:
    rc = False
    MsgBox err.Description
    Resume ERR_END
End Function


Public Function CopyRecipeData(ByVal SettingName As String) As Boolean
Dim rc As Boolean
Dim i As Integer
    On Error GoTo ERR_COPY
    '---------------------------
    ' set excel page
    '---------------------------
   ' Call SetUnit
    Call FormatPage
    
    Call SetInformation(i)
    Call SetComponent(i)
    Call SetHannaCode(i)

    
    rc = True
ERR_END:
    On Error GoTo 0
    CopyRecipeData = rc
    Exit Function
ERR_COPY:
    rc = False
    MsgBox err.Description
    Resume Next
End Function

Private Function SetInformation(ByRef Riga As Integer)
Dim i As Integer
Dim Density As String
Riga = 2

With uRecipe

    Density = IIf(.Density = 0, "1", .Density)
    Call AddValue(Riga, 2, "Recipe", True)
    Call AddValue(Riga, 3, "Description", True)
    
    Call AddValue(Riga, 4, "Revision", True)
    Call AddValue(Riga, 5, "Revision Date", True)
    Call AddValue(Riga, 6, "Revision Note", True)
    Call AddValue(Riga, 7, "Line", True)
    Call AddValue(Riga, 8, "Exp", True)
    Call AddValue(Riga, 9, "Procedure", True)
    Call AddValue(Riga, 10, "Is Mix", True)
    
    Call AddValue(Riga + 2, 2, "Density", True)
    Call AddValue(Riga + 2, 3, "MaxQty", True)
    Call AddValue(Riga + 2, 4, "MinQty", True)
    Call AddValue(Riga + 2, 5, "MinQty2", True)
    Call AddValue(Riga + 2, 6, "Multiple", True)
    Call AddValue(Riga + 2, 7, "Mix", True)
    Call AddValue(Riga + 2, 8, "No Preparation ", True)
    Call AddValue(Riga + 2, 9, "Classification ", True)
    
    Call AddValue(Riga + 2, 10, "Procedure # ", True)
    Call AddValue(Riga + 2, 11, "Procedure Date ", True)
    
    
    Call AddValue(Riga + 1, 2, .Code)
    Call AddValue(Riga + 1, 3, .Description)
    
    Call AddValue(Riga + 1, 4, .Rev)
    Call AddValue(Riga + 1, 5, .RevDate)
    Call AddValue(Riga + 1, 6, .NoteRev)
    Call AddValue(Riga + 1, 7, .Line)
    Call AddValue(Riga + 1, 8, .Exp)
    Call AddValue(Riga + 1, 9, .Procedure)
    Call AddValue(Riga + 1, 10, IIf(.bIsMix, "X", ""))
    
    
    Call AddValue(Riga + 3, 2, Replace(Density, ",", "."))
    Call AddValue(Riga + 3, 3, .MaxQty & " " & .UmMax)
    Call AddValue(Riga + 3, 4, .MinQty & " " & .UmMax)
    Call AddValue(Riga + 3, 5, .MinQty2 & " " & .UmMinQty)
    Call AddValue(Riga + 3, 6, .Multiple & " " & .UmMultiple)
    Call AddValue(Riga + 3, 7, .Mix)
    Call AddValue(Riga + 3, 8, .Procedure)
    Call AddValue(Riga + 3, 9, "'" & .Procedure)

    
End With

    Riga = Riga + 6

End Function
Private Function SetComponent(ByRef Riga As Integer)
Dim i As Integer
Dim Density As String

    Call AddValue(Riga - 1, 2, "Recipe Component", True)
    Call AddValue(Riga, 2, "Chemical Code", True)
    Call AddValue(Riga, 3, "Description", True)
    Call AddValue(Riga, 4, "Cas", True)
    Call AddValue(Riga, 5, "Qty", True)
    Call AddValue(Riga, 6, "Density", True)
    Call AddValue(Riga, 7, "Perc", True)
    Call AddValue(Riga, 8, "TolerancePerc", True)
    Call AddValue(Riga, 9, "Critical RM", True)
    Call AddValue(Riga, 10, "Note", True)
    Call AddValue(Riga, 11, "Is Mix", True)
    
    
    For i = 0 To uRecipe.RmxRecipeCount - 1
        With uRecipe.RmxRecipe(i)
            
            Riga = Riga + 1
            Density = IIf(.Density = 0, "1", .Density)
            Call AddValue(Riga, 2, .CHCode)
            Call AddValue(Riga, 3, .Description)
            Call AddValue(Riga, 4, .Cas)
            Call AddValue(Riga, 5, .Qty & " " & .Um)
            Call AddValue(Riga, 6, Density)
            Call AddValue(Riga, 7, Replace(.Perc, ",", ".") & " %")
            Call AddValue(Riga, 8, Replace(.TolerancePerc, ",", ".") & " %")
            Call AddValue(Riga, 9, .CriticalRM)
            Call AddValue(Riga, 10, .Note)
            Call AddValue(Riga, 11, IIf(.bMix, "X", ""))
        End With
    Next

    Riga = Riga + 4

End Function

Private Function SetHannaCode(ByRef Riga As Integer)
Dim i As Integer
Dim MaxHannaCode As Integer
On Error Resume Next
    MaxHannaCode = -1
    MaxHannaCode = UBound(uRecipe.HannaCodes)
    
    If MaxHannaCode = -1 Then Exit Function
    Call AddValue(Riga - 1, 2, "Hanna Code", True)
    
    Call AddValue(Riga, 2, "Code", True)
    Call AddValue(Riga, 3, "Description", True)
    Call AddValue(Riga, 4, "Line", True)
    Call AddValue(Riga, 5, "Qty", True)

    
    For i = 0 To UBound(uRecipe.HannaCodes)
        With uRecipe.HannaCodes(i)
            
            Riga = Riga + 1
      
            Call AddValue(Riga, 2, .Code)
            Call AddValue(Riga, 3, .ProductName)
            Call AddValue(Riga, 4, .Line)
            Call AddValue(Riga, 5, .Qty & " " & .Um)
           
        End With
    Next

    Riga = Riga + 2

End Function

