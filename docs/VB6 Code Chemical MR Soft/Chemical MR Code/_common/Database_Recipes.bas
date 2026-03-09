Attribute VB_Name = "Database_Recipes"
Option Explicit
Public zRecipe As RecipeType

Public Function SetGridRecipe(ByRef Grd1 As Grid) As Boolean


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
        .RowHeight(0) = 0
        .Cols = 5
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Code"
        .Column(1).Width = 150
        .Cell(0, 2).Text = "Description"
        .Column(2).Width = 280
       
        .Cell(0, 3).Text = "ID"
        .Column(3).Width = 0
        .Cell(0, 4).Text = "Line"
        .Column(4).Width = 200
                

        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .DefaultFont.Size = 10 ' * m_ControlGridFontSize
        .DefaultFont.Bold = False
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
   
 
End Function


Public Function FillGridRecipe(ByRef Grd As Grid, Optional ByVal Code As String, Optional bMainForm As Boolean, Optional ByVal sString As String, Optional ByVal bOnlyCriticals As Boolean)
Dim i As Integer
Dim t As Integer
Dim r As Integer
Dim MaxCount As Integer
Dim bMancaFormulation As Boolean
    On Error GoTo ERR_GRID
    ' --------------------------------------
    '
    '  filtra TabReport e riempi Tabella
    '
    ' --------------------------------------
    If InStr(UCase(Code), UCase("Code")) Then Code = ""
    
    


    With Grd
        .Rows = 1
        .ReadOnly = True
        .AutoRedraw = False
        With dbTabRecipe
            .Close
            .Open "SELECT *  FROM TabRecipe order by Code ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
            .filter = ""
            If sString = "Recipe" Then
                If Code <> "" Then .filter = "Code like '*" & Trim(Code) & "*'"
            Else
                'If Code <> "" Then .Filter = "Code like '*" & Code & "*'"
            End If
            
            If .EOF Then Exit Function
            MaxCount = .RecordCount
            .MoveFirst
        End With
        
        For r = 1 To MaxCount
        
        
        
            If bOnlyCriticals Then
                
                If CheckCriticalRecipes(IIf(IsNull(Trim(dbTabRecipe!Code)), "", Trim(dbTabRecipe!Code))) = False Then GoTo cont:
                
            End If
       
       
            .AddItem "", False
            i = .Rows - 1
            .Cell(i, 0).Text = i
            .Cell(i, 1).Text = "  " & IIf(IsNull(Trim(dbTabRecipe!Code)), "", Trim(dbTabRecipe!Code))
             
             bMancaFormulation = GetRecipeFormulation(IIf(IsNull(Trim(dbTabRecipe!Code)), "", Trim(dbTabRecipe!Code)))
            
            .Cell(i, 1).Alignment = cellLeftCenter
            .Cell(i, 2).Text = "  " & IIf(IsNull(Trim(dbTabRecipe!Description)), "", Trim(dbTabRecipe!Description))
            .Cell(i, 2).Alignment = cellLeftCenter
            .Cell(i, 3).Text = dbTabRecipe!ID
            .Cell(i, 4).Text = IIf(IsNull(Trim(dbTabRecipe!Line)), "", Trim(dbTabRecipe!Line))
            For t = 1 To .Cols - 1
                If bMainForm Then
                    .Cell(i, t).ForeColor = vbColorTextDarkBlue
                Else
                    .Cell(i, t).ForeColor = vbColorDarkFont
                End If
                
                If bMancaFormulation Then
                    
                    .Cell(i, t).ForeColor = vbColorOrange
                 
                End If
                 
           
            Next
                
                If i > 1 Then
                
                If .Cell(i, 1).Text = .Cell(i - 1, 1).Text Then
                   For t = 1 To .Cols - 1
                    .Cell(i, t).BackColor = vbColorTextLightBlue
                   Next
                End If
    
            End If
cont:
            dbTabRecipe.MoveNext
        Next
ERR_END:
        If Not (bMainForm) Then .Column(1).AutoFit
        '.Column(2).AutoFit
        .AutoRedraw = True
        .Refresh
    End With

    Exit Function
ERR_GRID:
    MessageInfoTime = 2000
 
    PopupMessage 2, err.Description
    Resume Next
End Function


Private Function CheckCriticalRecipes(ByVal Recipe As String) As Boolean
Dim rc As Boolean
Dim strCritical As String
Dim strChemical As String

Dim i As Integer

    On Error GoTo ERR_CHECK:
    rc = True

    With dbTabRMxRecipe
        .filter = ""
        .filter = "RecipeCode='" & Replace(Recipe, "'", "''") & "'"
        
        If .EOF Then
            rc = False
        Else
            .MoveFirst
            For i = 1 To .RecordCount
                strChemical = IIf(IsNull(Trim(!CHCode)), "", Trim(!CHCode))
                
                If strChemical <> "" Then
                    
                    With dbTabRawMaterial
                        .filter = ""
                        .filter = "Code='" & strChemical & "'"
                        If .EOF Then
                            rc = False
                        Else
                            
                            strCritical = IIf(IsNull(Trim(!CriticalRM)), "", Trim(!CriticalRM))
                            rc = IIf(Len(strCritical) > 0, True, False)
                            If rc = True Then
                                GoTo ERR_END
                            End If
                        End If
                        
                    End With
                    
                Else
                    rc = False
                End If
            
                .MoveNext
            Next
            
        
        End If
    
    End With



ERR_END:

    CheckCriticalRecipes = rc

    Exit Function
ERR_CHECK:
    rc = False
    MessageInfoTime = 2000
 
    PopupMessage 2, err.Description
    GoTo ERR_END:
End Function


Public Function SetGridEditRecipe(ByRef Grd As Grid) As Boolean
   
    With Grd
        .Rows = 1
        .AutoRedraw = False
        .ReadOnly = False
        .DisplayFocusRect = True 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionFree
        .DefaultFont.Size = 10 '* m_ControlGridFontSize
        
        .DefaultFont.Bold = False
        .DefaultRowHeight = 35
        .ExtendLastCol = True
        .Cols = 3
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Column(1).Width = 250
        .Column(2).Width = 250
        .RowHeight(0) = 0
        .Rows = 21
        .Cols = 3
        .Cell(1, 1).Text = "  " & "Code"
        .Cell(2, 1).Text = "  " & "Description"
        .Cell(3, 1).Text = "  " & "Line"
        .Cell(4, 1).Text = "  " & "Recipe Revision #"
        .Cell(5, 1).Text = "  " & "Rev Date"
        .Cell(6, 1).Text = "  " & "Note"
        .Cell(7, 1).Text = "  " & "Exp ( Years )"
        .Cell(8, 1).Text = "  " & "Density"
        .Cell(9, 1).Text = "  " & "Max Q.ty"
        .Cell(10, 1).Text = "  " & "um Max Q.ty"
        .Cell(11, 1).Text = "  " & "Min Q.ty"
        .Cell(12, 1).Text = "  " & "Multiple"
        .Cell(13, 1).Text = "  " & "um Multiple"
        .Cell(14, 1).Text = "  " & "Min Q.ty Multiple"
        .Cell(15, 1).Text = "  " & "um Min Q.ty"
        .Cell(16, 1).Text = "  " & "Mix"
        .Cell(17, 1).Text = "  " & "No Preparation"
        .Cell(17, 2).CellType = cellCheckBox
        
        .Cell(18, 1).Text = "  " & "Classification"
        
        .Cell(19, 1).Text = "  " & "Procedure #"
        .Cell(20, 1).Text = "  " & "Procedure Date"
        
        
        
        .Cell(14, 2).Locked = True
        .Cell(16, 2).Locked = True
        
        .Cell(3, 2).CellType = cellButton ' line
        .Cell(10, 2).CellType = cellButton ' um max qty
        .Cell(13, 2).CellType = cellButton ' um multiple
        .Cell(15, 2).CellType = cellButton ' um min qty
        .ButtonLocked = True
        .ReadOnlyFocusRect = Solid
        
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            
        Next

        For i = 1 To .Rows - 1
            .Cell(i, 1).BackColor = &HF0F0F0 'vbColorUnabled
            .Cell(i, 1).ForeColor = vbColorDarkFont 'vbColorDarkFont 'vbColorForeFixed  ' vbColorTextDarkBlue
            .Cell(i, 1).FontBold = False
            .Cell(i, 1).Locked = True
            .Cell(i, 2).ForeColor = vbColorDarkFont
        Next
       
       ' .RowHeight(15) = 0
        .RowHeight(16) = 0
        .ReadOnly = False
        .AutoRedraw = True
        .Refresh
        
    End With
End Function


Public Sub CopyRecipeGrid2(ByVal Grd2 As Grid, ByVal lId As Long)
    If lId = 0 Then Exit Sub
    Dim i As Integer
    Dim RecipeCode As String
    Dim bFormulation As Boolean
    
    
    bFormulation = False

     With dbTabRecipe
        '.Close
        '.Open "SELECT *  FROM TabRecipe order by Code ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
        .filter = ""
        .filter = "ID='" & lId & "'"
        If .EOF Then Exit Sub
        .MoveFirst
        
        RecipeCode = Trim(!Code)
        
        
    End With
    
    With dbTabRMxRecipe
        .filter = ""
        .filter = "RecipeCode='" & RecipeCode & "'"
        If .EOF Then
        
        Else
            bFormulation = True
            '.MoveFirst
            'For i = 1 To .RecordCount
                    
                
                '.MoveNext
           ' Next
        
        End If
        
    End With
    
   
    With Grd2
       .AutoRedraw = False
       For i = 1 To .Rows - 1
            .Cell(i, 2).Text = IIf(IsNull(Trim(dbTabRecipe.fields(i))), "", Trim(dbTabRecipe.fields(i)))
       Next
        
        For i = 2 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter ' cellLeftCenter
        
        Next
        
    
        .Cell(12, 2).BackColor = vbColorAzzurrino
        .Cell(13, 2).BackColor = vbColorAzzurrino
        
        
        If .Cell(15, 2).Text = "" Then .Cell(15, 2).Text = "pcs"
        If bFormulation Then
            .Cell(1, 2).BackColor = &H8000&    'vbGreenColor
            .Cell(1, 2).ForeColor = vbWhite
        Else
            .Cell(1, 2).BackColor = &HFFFFFF
            .Cell(1, 2).ForeColor = vbBlack
        End If
        .RowHeight(15) = 0
        .Refresh
        .AutoRedraw = True
    End With
    
    

End Sub






Public Sub Grd2_Recipe_LeaveCell(ByVal Grd2 As Grid, ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean, ByVal lRow As Long)


Dim sValue As String
Dim sString As String
Dim MinQty As Double
Debug.Print "Leave ", Row, Col
With Grd2
    sValue = .Cell(Row, Col).Text
    sString = Trim(.Cell(Row, 1).Text)
    If Col = 2 Then
        If lRow = Row Then
        
            Select Case Row
                Case 1
                    ' CODE
                    If Len(sValue) = 0 Then
                        PopupMessage 2, "Warning : Code must be a valid value...."
                       
                    End If
                Case 8, 9, 11, 12
                    ' density
                    
                    If IsNumeric(sValue) Then
CheckQty:
                        Call SetMinQtyMultiple(Grd2)
                    
                    Else
                    
                        PopupMessage 2, "Warning : " & sString & "  must be a Number....", , , sString
                        .Cell(Row, Col).Text = ""
                    End If
                Case 10, 13, 15
                    GoTo CheckQty
                    
                Case 18
                    If IsDate(sValue) Then
                        .Cell(Row, Col).Text = FormatDataLAT(sValue)
                    Else
                        PopupMessage 2, "Please enter a valid Date...", , True, "Revision Date"
                        
                    End If
              
            
            End Select
        
        
            .Cell(Row, Col).Alignment = cellCenterCenter
        End If
    End If
    
End With

Exit Sub

err:
PopupMessage 2, sString
Grd2.Cell(Row, Col).Text = ""
Return
End Sub


Public Function SetMinQtyMultiple(ByVal Grid2 As Grid) As Double
Dim MinQtyMultiple As Double
Dim Col As Integer
Dim umMin As String
Dim UmMultiple As String
Dim umMinQtyMultiple As String
Dim Multiple As Double
Dim Density As Double
Dim MinQty As Double
Dim bUmMassa As Boolean

Col = 2

With Grid2

    
    If .Cell(8, Col).Text <> "" And .Cell(11, Col).Text <> "" And .Cell(12, Col).Text <> "" Then
        
        
        umMin = .Cell(10, Col).Text
        bUmMassa = SetbUmMassa(umMin)
        
        UmMultiple = .Cell(13, Col).Text
        umMinQtyMultiple = "pcs"
        Density = CDbl(.Cell(8, Col).Text)
        MinQty = CDbl(.Cell(11, Col).Text)
        Multiple = CDbl(.Cell(12, Col).Text)
    
    
        If SetbUmMassa(umMin) = SetbUmMassa(UmMultiple) Then
            Density = 1
        Else
            
        End If
        
        
        If umMin <> "" And UmMultiple <> "" And umMinQtyMultiple <> "" Then
        
            MinQtyMultiple = (MinQty * Um(umMin) * Density) / (Multiple * Um(UmMultiple))
            MinQtyMultiple = Int((MinQtyMultiple / Um(umMinQtyMultiple)))
            .Cell(14, Col).Text = MinQtyMultiple
            .Cell(14, Col).Alignment = cellCenterCenter
        
        End If
        
    End If
                              
End With
         SetMinQtyMultiple = MinQtyMultiple
End Function

Public Function SaveDatabaseRecipe(ByVal Grd2 As Grid, ByVal bClone As Boolean, ByVal OldCode As String) As Boolean
Dim rc As Boolean
Dim MyNewCode As String
Dim RangeMin As String
Dim RangeMax As String

On Error GoTo ERR_SAVE
rc = True
    MyNewCode = Trim(Grd2.Cell(1, 2).Text)
    
    If MyNewCode = "" Then
        PopupMessage 2, "Please Enter a valid Code!"
        Exit Function
    End If
    
    With dbTabRecipe
        .filter = ""
        .filter = "Code='" & MyNewCode & "'"
        If .EOF Then
            .AddNew
            If bClone Then
                '-------------------------------
                ' clona la ricetta+revision...
                '-------------------------------
                Call CloneRecipe(OldCode, MyNewCode)
            End If
        Else
            If F_MsgBox.DoShow("Code already exsist. Replace Info?") Then
            Else
                Exit Function
            End If
            
        End If
        
        Dim i As Integer
        Dim Value As String
        For i = 1 To Grd2.Rows - 1
        
            Value = Trim(Grd2.Cell(i, 2).Text)
            Select Case i
            
                Case 8, 9, 11, 12, 14
                    If IsNumeric(Value) = False Then GoTo cont
                Case 5
                    If IsDate(Value) Then
                    Else
                        Value = FormatDataLAT(Now())
                    End If
                Case 19 ' classification
            End Select
          
            .fields(i) = Value
cont:
        
        
        
        Next
        
        .Update
    End With


ERR_END:
    On Error GoTo 0
    If rc Then
        PopupMessage 2, "Code saved!", , , MyNewCode
    Else
        PopupMessage 2, "Warning : a problem occurred, please check all entries before Save"
    End If
    
    SaveDatabaseRecipe = rc
    Exit Function
    
ERR_SAVE:
    rc = False
    MsgBox err.Description
    Resume Next:

End Function

Private Function CloneRecipe(ByVal OldCode As String, ByVal MyNewCode As String)
Dim i As Integer
Dim ComponentRecipe() As RmxRecipe

dbCode.Execute ("DELETE * FROM TabRMxRecipe WHERE RecipeCode='" & MyNewCode & "'")

With dbTabRMxRecipe
    .Close
    .Open "SELECT *  FROM TabRMxRecipe order by id ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    .filter = ""
    .filter = "RecipeCode='" & OldCode & "'"
    If .EOF Then
    Else
        .MoveFirst
        
        ReDim ComponentRecipe(.RecordCount)
        
        For i = 0 To .RecordCount - 1
        
            ComponentRecipe(i).RecipeCode = MyNewCode
            ComponentRecipe(i).CHCode = IIf(IsNull(Trim(!CHCode)), "", Trim(!CHCode))
            ComponentRecipe(i).Description = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
            ComponentRecipe(i).Cas = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
            ComponentRecipe(i).Qty = IIf(IsNull(Trim(!Qty)), "", Trim(!Qty))
            ComponentRecipe(i).Um = IIf(IsNull(Trim(!Um)), "", Trim(!Um))
            ComponentRecipe(i).Perc = IIf(IsNull(Trim(!Perc)), "", Trim(!Perc))
            ComponentRecipe(i).Note = IIf(IsNull(Trim(!Note)), "", Trim(!Note))
            ComponentRecipe(i).bMix = !bMix
            ComponentRecipe(i).TolerancePerc = IIf(IsNull(Trim(!TolerancePerc)), "", Trim(!TolerancePerc))

            .MoveNext
        Next
  
        For i = 0 To .RecordCount - 1
            .AddNew
            !RecipeCode = ComponentRecipe(i).RecipeCode
            !CHCode = ComponentRecipe(i).CHCode
            !Description = ComponentRecipe(i).Description
            !Cas = ComponentRecipe(i).Cas
            !Qty = ComponentRecipe(i).Qty
            !Um = ComponentRecipe(i).Um
            !Perc = ComponentRecipe(i).Perc
            !Note = ComponentRecipe(i).Note
            !bMix = ComponentRecipe(i).bMix
            !TolerancePerc = ComponentRecipe(i).TolerancePerc
        
        Next
        
        .Update
    
    End If
    


dbCode.Execute ("DELETE * FROM TabRecipeRevisionHistory WHERE Recipe='" & MyNewCode & "'")

Dim RecipeRevisionHistory() As RevisionHistory

With dbTabRecipeRevisionHistory

    .Close
    .Open "SELECT *  FROM TabRecipeRevisionHistory order by id ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    .filter = ""
    .filter = "Recipe='" & OldCode & "'"
    If .EOF Then
    Else
        .MoveFirst
        
        ReDim RecipeRevisionHistory(.RecordCount)
        
        For i = 0 To .RecordCount - 1

            RecipeRevisionHistory(i).Recipe = MyNewCode
            RecipeRevisionHistory(i).RevDate = IIf(IsNull(Trim(!RevDate)), "", Trim(!RevDate))
            RecipeRevisionHistory(i).RevNumber = IIf(IsNull(Trim(!RevNumber)), "", Trim(!RevNumber))
            RecipeRevisionHistory(i).RevType = IIf(IsNull(Trim(!Type)), "", Trim(!Type))
            RecipeRevisionHistory(i).Description = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
            RecipeRevisionHistory(i).Operator = IIf(IsNull(Trim(!Operator)), "", Trim(!Operator))
            .MoveNext
        Next
        For i = 0 To .RecordCount - 1
            .AddNew
    
            !Recipe = RecipeRevisionHistory(i).Recipe
            If RecipeRevisionHistory(i).RevDate <> "" And IsDate(RecipeRevisionHistory(i).RevDate) Then
                !RevDate = RecipeRevisionHistory(i).RevDate
            End If
            !RevNumber = RecipeRevisionHistory(i).RevNumber
            !Type = RecipeRevisionHistory(i).RevType
            !Description = RecipeRevisionHistory(i).Description
            !Operator = RecipeRevisionHistory(i).Operator
        
        Next
        
        .Update
        
    End If

End With

End With


End Function

Public Sub AddComboRecipe(ByVal Combo1 As ComboBox)

    Combo1.Clear
    Combo1.AddItem "Recipe"
    Combo1.AddItem "Hanna Code"
    Combo1.ListIndex = 0
End Sub




Public Function GetStrRecipe(ByRef iRecipe() As RecipeType) As String
Dim i As Integer
    
    For i = 1 To UBound(iRecipe)
        If iRecipe(i).bHide = False And iRecipe(i).bUpdated Then
            If i > 1 Then
                If InStr(GetStrRecipe, iRecipe(i).Code) Then
                Else
                    GetStrRecipe = GetStrRecipe & IIf(GetStrRecipe = "", "", " ; ") & iRecipe(i).Code
                End If
            
            Else
                GetStrRecipe = iRecipe(i).Code
            End If
            
           
                
        End If
    Next
    GetStrRecipe = Trim(Left(GetStrRecipe, 255))
End Function

Public Function GetStrDescriptionRecipe(ByRef iRecipe() As RecipeType) As String
Dim i As Integer
    For i = 1 To UBound(iRecipe)
        If iRecipe(i).bHide = False And iRecipe(i).bUpdated Then
            If i > 1 Then
                If InStr(GetStrDescriptionRecipe, iRecipe(i).Description) Then
                Else
                    GetStrDescriptionRecipe = GetStrDescriptionRecipe & IIf(GetStrDescriptionRecipe = "", "", " ; ") & iRecipe(i).Description
                End If
            Else
                GetStrDescriptionRecipe = iRecipe(i).Description
            End If
        End If
    Next
    
    GetStrDescriptionRecipe = Trim(Left(GetStrDescriptionRecipe, 255))
End Function

Public Function GetStrLineRecipe(ByRef iRecipe() As RecipeType) As String
Dim i As Integer
    
    For i = 1 To UBound(iRecipe)
        If iRecipe(i).bHide = False And iRecipe(i).bUpdated Then
            If i > 1 Then
                If InStr(GetStrLineRecipe, iRecipe(i).Line) Then
                Else
                    GetStrLineRecipe = GetStrLineRecipe & IIf(GetStrLineRecipe = "", "", " ; ") & iRecipe(i).Line
                End If
            Else
                GetStrLineRecipe = iRecipe(i).Line
            End If
        End If
    Next
    GetStrLineRecipe = Trim(Left(GetStrLineRecipe, 255))
End Function


Public Function GetRecipeDescription(ByVal strName As String) As String
If strName <> "" Then
    With dbTabRecipe
        .filter = ""
        .filter = "Code='" & strName & "'"
        If .EOF Then
            
            With dbTabRawMaterial
                .filter = ""
                .filter = "Code='" & strName & "'"
                If .EOF Then
                Else
                    GetRecipeDescription = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
                End If
            End With
            
        Else
            GetRecipeDescription = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
        End If
    End With
End If
End Function

Public Function GetRecipeFormulation(ByVal Recipe As String) As Boolean
Dim rc As Boolean
    rc = False
    With dbTabRMxRecipe
        .filter = ""
        .filter = "RecipeCode='" & Replace(Trim(Recipe), "'", "''") & "'"
        If (.EOF) Then rc = True
    End With
    GetRecipeFormulation = rc
End Function





Public Function SerarchRecipePerLine(ByRef Grd1 As Grid, ByVal LineRecipe As String)
Dim bTutti As Boolean
Dim i As Integer

bTutti = False
If LineRecipe = "" Or InStr(UCase(LineRecipe), "ALL") Then
bTutti = True
End If

With Grd1
    .AutoRedraw = False
    If .Rows > 1 Then
            
        For i = 1 To .Rows - 1
            
            If bTutti Then
                .RowHeight(i) = 25
            Else
            
                If InStr(UCase(Trim(.Cell(i, 4).Text)), UCase(Trim(LineRecipe))) Then
                    .RowHeight(i) = 25
            
                Else
                    .RowHeight(i) = 0
                End If
            End If
        
        Next
    End If

    .Refresh
    .AutoRedraw = True
End With


End Function

Public Function GetRecipeIdByName(ByVal strName As String) As Long
If strName <> "" Then
    With dbTabRecipe
        .filter = ""
        .filter = "Code='" & strName & "'"
        If .EOF Then
            
            GetRecipeIdByName = 0
            
        Else
            GetRecipeIdByName = !ID
        End If
    End With
End If
End Function
