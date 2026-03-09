Attribute VB_Name = "Database_RawMaterials"

Option Explicit
Public Function SetGridChemicalRM(ByRef Grd1 As Grid) As Boolean


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
        .Cols = 4
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Code"
        .Column(1).Width = 120
        .Cell(0, 2).Text = "Description"
        .Column(2).Width = 250
        .Cell(0, 3).Text = "ID"
        .Column(3).Width = 0

        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
        Next
        .DefaultFont.Size = 9 ' * m_ControlGridFontSize
        .DefaultFont.Bold = False
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
   
 
End Function

Public Function SetGridEditChemicalRM(ByRef Grd As Grid) As Boolean
   
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
        .Rows = 16

        .Cell(1, 1).Text = "  " & "Code"
        .Cell(2, 1).Text = "  " & "Description"
        .Cell(3, 1).Text = "  " & "Cas"
        .Cell(4, 1).Text = "  " & "Chemical Reaction Liquid"
        .Cell(5, 1).Text = "  " & "Classification"
        .Cell(6, 1).Text = "  " & "Pictograms"
        .Cell(7, 1).Text = "  " & "Um"
        .Cell(8, 1).Text = "  " & "Manufacturer"
        .Cell(9, 1).Text = "  " & "Manufacturer Code"
        .Cell(10, 1).Text = "  " & "Location"
        .Cell(11, 1).Text = "  " & "Specified Location"
        .Cell(12, 1).Text = "  " & "Mix"
        .Cell(13, 1).Text = "  " & "Date Modified"
        .Cell(14, 1).Text = "  " & "Critical RM"
        .Cell(15, 1).Text = "  " & "Density"

        .Cell(12, 2).CellType = cellCheckBox
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


        
        .RowHeight(6) = 0
        .ReadOnly = False
        .AutoRedraw = True
        .Refresh
        
    End With
End Function
Public Sub CopyChemicalRMGrd1(ByRef Grd As Grid, Optional ByVal Code As String, Optional bMainForm As Boolean, Optional ByVal sString As String, Optional ByVal bMixes As Boolean, Optional ByVal bOnlyCriticals As Boolean)
Dim i As Integer
Dim t As Integer
Dim filterString As String
Dim strCritical As String

    filterString = UCase(Replace(Trim(Code), "'", "''"))
     sString = IIf(bMixes, " and bMix=true", "")
     

     
     With dbTabRawMaterial
        
        .filter = ""
        .filter = IIf(bMixes, "bMix=true", "")
        If filterString <> "" And Code <> "" Then
             .filter = "Code like '*" & filterString & "*'" & sString
             If .EOF Then
                .filter = "Code = '" & filterString & "'" & sString
             End If
        End If
        If .EOF Then Exit Sub
        .MoveFirst

    End With
    
    
    With Grd
       .AutoRedraw = False
       For i = 1 To dbTabRawMaterial.RecordCount
       
       If bOnlyCriticals Then
        strCritical = IIf(IsNull(Trim(dbTabRawMaterial!CriticalRM)), "", Trim(dbTabRawMaterial!CriticalRM))
        If Len(strCritical) = 0 Then GoTo cont:
            
       End If

        .AddItem "", False
           
            .Cell(.Rows - 1, 1).Text = IIf(IsNull(Trim(dbTabRawMaterial.fields(1))), "", Trim(dbTabRawMaterial.fields(1)))
            .Cell(.Rows - 1, 2).Text = IIf(IsNull(Trim(dbTabRawMaterial.fields(2))), "", Trim(dbTabRawMaterial.fields(2)))
        
         
         .Cell(.Rows - 1, 3).Text = dbTabRawMaterial!ID
         
         If dbTabRawMaterial!bMix Then
            .Cell(.Rows - 1, 1).FontBold = True
            .Cell(.Rows - 1, 1).ForeColor = vbColorTextDarkBlue
            .Cell(.Rows - 1, 2).FontBold = True
            .Cell(.Rows - 1, 2).ForeColor = vbColorTextDarkBlue
            
         End If
         
         
cont:
         dbTabRawMaterial.MoveNext
       Next

       .Refresh
       
       .ReadOnly = True
        .AutoRedraw = True
    End With

End Sub

Public Sub CopyChemicalRMGrd2(ByVal Grd2 As Grid, ByVal lId As Long)
Dim i As Integer
Dim t As Integer
    If lId = 0 Then Exit Sub
    
    On Error GoTo ERR_CH

     With dbTabRawMaterial
            
        .filter = ""
        .filter = "ID='" & lId & "'"
        If .EOF Then Exit Sub
        .MoveFirst

    End With
    
    
    With Grd2
       ' .DefaultFont.Size = 12 * m_ControlGridFontSize
       For i = 1 To .Rows - 1
            .Cell(i, 2).Text = IIf(IsNull(Trim(dbTabRawMaterial.fields(i))), "", Trim(dbTabRawMaterial.fields(i)))
       Next
        

        For i = 2 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter ' cellLeftCenter
            
        Next
        
       If .Cell(7, 2).Text = "ml" Then .Cell(7, 2).BackColor = vbColorAzzurrino
       If .Cell(15, 2).Text <> "1" Then .Cell(15, 2).BackColor = vbColorAzzurrino
       
   
    
    End With
    
ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_CH:
    MsgBox err.Description
    Resume Next

End Sub

Public Sub Grd2_ChemicalRM_LeaveCell(ByVal Grd2 As Grid, ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean, ByVal lRow As Long)


Dim sValue As String
Dim sString As String
Dim Um As String
Debug.Print "Leave ", Row, Col
With Grd2
    sValue = .Cell(Row, Col).Text
    
    Um = .Cell(6, Col).Text
    
    If Col = 2 Then
        If lRow = Row Then
        
            Select Case Row
                Case 1
                    ' CODE
                    If Len(sValue) = 0 Then
                        PopupMessage 2, "Warning : Code must be a valid value...."
                       
                    End If
                    
                Case 7
                    'um
                Case 15
                    '
                    If IsNumeric(sValue) Then
                        
                        If sValue <> 1 Then
                            
                            If Um <> "ml" Then
                                
                                If F_MsgBox.DoShow("Change measurement unit to 'ml' ?", "Um Raw Material", , "ml", "g") Then
                                    .Cell(7, 2).Text = "ml"

                                End If
                                
                            End If
                        
                        End If
                    Else
                        PopupMessage 2, "Warning : Denisty must be a valid value....", , , "Raw Material Density"
                    
                    End If
             
            End Select
        
        
        
        End If
    End If
    .Refresh
    .AutoRedraw = True
End With

Exit Sub

err:
PopupMessage 2, sString
Grd2.Cell(Row, Col).Text = ""
Return
End Sub

Public Function SaveDatabaseChemicalRM(ByVal Grd2 As Grid) As Boolean
Dim rc As Boolean
Dim MyNewCode As String
Dim bMix As Boolean

On Error GoTo ERR_SAVE
rc = True
    
    MyNewCode = Trim(Grd2.Cell(1, 2).Text)

    If MyNewCode = "" Then
        PopupMessage 2, "Please Enter a valid Code!"
        Exit Function
    End If
    
    With dbTabRawMaterial
    
        .filter = ""
        .filter = "Code='" & MyNewCode & "'"
        If .EOF Then
        
            .AddNew
        Else
            If F_MsgBox.DoShow("Code already exsist. Replace Info?") Then
            Else
                Exit Function
            End If
            
        End If
        Dim i As Integer
        For i = 1 To Grd2.Rows - 1
            If Grd2.Cell(15, 2).Text = "" Then
                
                If F_MsgBox.DoShow("Set Density = 1 ?", MyNewCode) Then
                    Grd2.Cell(15, 2).Text = "1"
                End If
            
            End If
            If i = 13 Then
                .fields(i) = Now()
            Else
                .fields(i) = Trim(Grd2.Cell(i, 2).Text)
            End If
            
          
        Next
        
        bMix = CBool(Trim(Grd2.Cell(12, 2).Text))
        .Update
    End With
    If bMix = False Then
        Call CheckTabRecipe(MyNewCode)
    End If

ERR_END:
    On Error GoTo 0
    If rc Then
        PopupMessage 2, "Code : " & MyNewCode & " saved!"
    Else
        PopupMessage 2, "Warning : a problem occurred, please check all entries before Save"
    End If
    
    SaveDatabaseChemicalRM = rc
    Exit Function
    
ERR_SAVE:
    rc = False
    MsgBox err.Description
    Resume Next:

End Function

Private Function CheckTabRecipe(ByVal MyNewCode As String)
    With dbTabRecipe
        .filter = ""
        .filter = "Code='" & Trim(MyNewCode) & "'"
        
        If .EOF Then
        
        Else
            
            If F_MsgBox.DoShow("Delete Code from Recipes ?", MyNewCode) Then
                
                .Delete
                .Update
                
                Call DeleteRecipeComponentByCode(MyNewCode)
            
            
            End If
            
        End If
    End With

End Function


Public Sub AddComboChemicalRM(ByVal Combo1 As ComboBox)
    Combo1.Clear
    Combo1.AddItem "Code"
    Combo1.ListIndex = 0
End Sub



Public Function AddComponentGrid(ByVal Grd As Grid, ByVal RecipeCode As String)

Dim i As Integer
  
        '.Cell(0, 1).Text = "CH Code"
        '.Cell(0, 2).Text = "Description"
        '.Cell(0, 3).Text = "CAS"
        '.Cell(0, 4).Text = "Q.ty/multiple"
        '.Cell(0, 5).Text = "(um)"
        '.Cell(0, 6).Text = "%"
        '.Cell(0, 9).Text = "Note"
        
        With Grd
            .Rows = 1
            .AutoRedraw = False
            With dbTabRMxRecipe
                .filter = ""
                .filter = "RecipeCode='" & RecipeCode & "'"
                If .EOF Then
                
                Else
                    .MoveFirst
                    For i = 1 To .RecordCount
                        Grd.AddItem "", False
                        Grd.Cell(Grd.Rows - 1, 1).Text = IIf(IsNull(Trim(!CHCode)), "", Trim(!CHCode))
                        Grd.Cell(Grd.Rows - 1, 2).Text = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
                        Grd.Cell(Grd.Rows - 1, 3).Text = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
                        Grd.Cell(Grd.Rows - 1, 4).Text = CheckDot(IIf(IsNull(Trim(!Qty)), "", Trim(!Qty)))
                        Grd.Cell(Grd.Rows - 1, 5).Text = IIf(IsNull(Trim(!Um)), "", Trim(!Um))
                        Grd.Cell(Grd.Rows - 1, 6).Text = CheckDot(IIf(IsNull(Trim(!Perc)), "", Trim(!Perc)))
                        Grd.Cell(Grd.Rows - 1, 9).Text = IIf(IsNull(Trim(!Note)), "", Trim(!Note))
                        Grd.Cell(Grd.Rows - 1, 10).Text = !bMix
                        Grd.Cell(Grd.Rows - 1, 11).Text = GetCriticalRM(IIf(IsNull(Trim(!CHCode)), "", Trim(!CHCode)))
                        Grd.Cell(Grd.Rows - 1, 2).FontSize = 9
                        Grd.Cell(Grd.Rows - 1, 4).Alignment = cellRightCenter
                        Grd.Cell(Grd.Rows - 1, 5).Alignment = cellLeftCenter
                       ' Grd.Cell(Grd.Rows - 1, 6).Alignment = cellCenterCenter
                        
                        .MoveNext
                    Next
                End If
                
            End With
            .ReadOnly = False
            .Range(0, 4, 0, 5).Merge
            .Cell(0, 4).Alignment = cellCenterCenter
            .Column(6).Alignment = cellLeftCenter
            .Column(2).AutoFit
            .Refresh
            .ReadOnly = True
            .AutoRedraw = True
        
        End With
        
        

End Function

Public Function GetNoteRM(ByVal RecipeCode As String, ByVal CHCode As String) As String
    
    If RecipeCode <> "" And CHCode <> "" Then
        With dbTabRMxRecipe
            .filter = ""
            .filter = "RecipeCode='" & RecipeCode & "' and CHCode='" & CHCode & "'"
            If .EOF Then
            Else
                GetNoteRM = IIf(IsNull(Trim(!Note)), "", Trim(!Note))
            End If
        End With
    End If
    
End Function
Public Function GetCriticalRM(ByVal Code As String) As String
    
    If Code <> "" Then
        With dbTabRawMaterial
            .filter = ""
            .filter = "Code='" & Code & "'"
            If .EOF Then
            Else
                GetCriticalRM = IIf(IsNull(Trim(!CriticalRM)), "", Trim(!CriticalRM))
                
            End If
        End With
    End If
    
End Function
