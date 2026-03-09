Attribute VB_Name = "QRCode"
Option Explicit


Public Function DeleteRecipeComponentByCode(ByVal Code As String) As Boolean
Dim rc As Boolean
On Error GoTo ERR_DELETE
rc = True

    With dbTabMR
         .Close
         .Open "SELECT *  FROM TabMR order by id ", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
     End With
            
    dbCode.Execute ("DELETE * FROM TabMR WHERE Code='" & Code & "'")
    
ERR_END:
    On Error GoTo 0
    
    DeleteRecipeComponentByCode = rc
    Exit Function
ERR_DELETE:
    rc = False
    MsgBox err.Description
    Resume Next
    
End Function



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
        .Rows = 15

        .Cell(1, 1).Text = "  " & "Code"
        .Cell(2, 1).Text = "  " & "Description"
        .Cell(3, 1).Text = "  " & "Supplier"
        .Cell(4, 1).Text = "  " & "MR NMP"
        .Cell(5, 1).Text = "  " & "Physical State"
        .Cell(6, 1).Text = "  " & "Density"
        .Cell(7, 1).Text = "  " & "Unit"
        .Cell(8, 1).Text = "  " & "Parameter"
        .Cell(9, 1).Text = "  " & "FW Parameter"
        .Cell(10, 1).Text = "  " & "Storage"
        .Cell(11, 1).Text = "  " & "MinQTY"
        .Cell(12, 1).Text = "  " & "MS Type"
        .Cell(13, 1).Text = "  " & "Stock QTY"
        .Cell(14, 1).Text = "  " & "Last Updated"
 
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            
        Next
        
        
        For i = 1 To .Rows - 1
            .Cell(i, 1).BackColor = &HF0F0F0 'vbColorUnabled
            .Cell(i, 1).ForeColor = vbColorDarkFont 'vbColorDarkFont 'vbColorForeFixed  ' vbColorBlueProgram
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
Public Sub CopyChemicalRMGrd1(ByRef Grd As Grid, Optional ByVal Code As String, Optional bMainForm As Boolean, Optional ByVal sString As String)
Dim i As Integer
Dim t As Integer
Dim filterString As String
Dim strCritical As String

    filterString = UCase(Replace(Trim(Code), "'", "''"))
    

     
     With dbTabMR
        
        .Filter = ""
        If filterString <> "" And Code <> "" Then
             .Filter = "Code like '*" & filterString & "*'"
             If .EOF Then
                .Filter = "Code = '" & filterString & "'"
             End If
        End If
        If .EOF Then Exit Sub
        .MoveFirst

    End With
    
    
    With Grd
       .AutoRedraw = False
       For i = 1 To dbTabMR.RecordCount
       
     

        .AddItem "", False
           
            .Cell(.Rows - 1, 1).Text = IIf(IsNull(Trim(dbTabMR.fields(1))), "", Trim(dbTabMR.fields(1)))
            .Cell(.Rows - 1, 2).Text = IIf(IsNull(Trim(dbTabMR.fields(2))), "", Trim(dbTabMR.fields(2)))
        
         
         .Cell(.Rows - 1, 3).Text = dbTabMR!ID
         
         
cont:
         dbTabMR.MoveNext
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

     With dbTabMR
            
        .Filter = ""
        .Filter = "ID='" & lId & "'"
        If .EOF Then Exit Sub
        .MoveFirst

    End With
    
    
    With Grd2
       ' .DefaultFont.Size = 12 * m_ControlGridFontSize
       For i = 1 To .Rows - 1
            .Cell(i, 2).Text = IIf(IsNull(Trim(dbTabMR.fields(i))), "", Trim(dbTabMR.fields(i)))
       Next
        

        For i = 2 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter ' cellLeftCenter
            
        Next
        
        If .Cell(5, 2).Text = "" Then
            ' .Cell(5, 2).BackColor = vbRed
        Else
             .Cell(5, 2).BackColor = vbColorAzzurrino
        End If
        
         If .Cell(12, 2).Text = "" Then
           '  .Cell(12, 2).BackColor = vbRed
        Else
             .Cell(12, 2).BackColor = vbColorAzzurrino
        End If
               
   
    
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
    
    With dbTabMR
    
        .Filter = ""
        .Filter = "Code='" & MyNewCode & "'"
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
           
            If i = 13 Then
            
                GoTo cont:
            End If
            If i = 14 Then
                .fields(i) = Now()
           Else
                .fields(i) = Trim(Grd2.Cell(i, 2).Text)
            End If
            
cont:
        Next

        .Update
    End With
   
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
    Debug.Print Trim(Grd2.Cell(i, 2).Text)
    MsgBox err.Description
    Resume Next:

End Function



Public Sub AddComboChemicalRM(ByVal Combo1 As ComboBox)
    Combo1.Clear
    Combo1.AddItem "Code"
    Combo1.ListIndex = 0
End Sub
