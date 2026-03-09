Attribute VB_Name = "Database_PhrasesH"
Option Explicit
Public Function SetGridFrasiH(ByRef Grd1 As Grid) As Boolean


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
        .Cell(0, 2).Text = "Statement"
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

Public Function SetGridEditFrasiH(ByRef Grd As Grid) As Boolean
   
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
        .Rows = 7

        .Cell(1, 1).Text = "  " & "Code"
        .Cell(2, 1).Text = "  " & "Statement"
        .Cell(3, 1).Text = "  " & "Hazard Category"
        .Cell(4, 1).Text = "  " & "Precaution"
        .Cell(5, 1).Text = "  " & "Safety Equipments"
        .Cell(6, 1).Text = "  " & "Pictogram"
        

        
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


        
        
        .ReadOnly = False
        .AutoRedraw = True
        .Refresh
        
    End With
End Function
Public Sub CopyFrasiHGrd1(ByRef Grd As Grid, Optional ByVal Code As String, Optional bMainForm As Boolean, Optional ByVal sString As String)
Dim i As Integer
Dim t As Integer


     With dbTabFrasiH
            
        .Filter = ""
        If sString <> "" And Code <> "" Then
            '.Filter = "STDPreparationWay ='" & Replace(Trim(Code), "'", "''") & "'"
             .Filter = "Code like '*" & Replace(Trim(Code), "'", "''") & "*'"
        End If
        If .EOF Then Exit Sub
        .MoveFirst

    End With
    
    
    With Grd
       .AutoRedraw = False
       For i = 1 To dbTabFrasiH.RecordCount
        .AddItem "", False
            .Cell(.Rows - 1, 1).Text = IIf(IsNull(Trim(dbTabFrasiH.fields(1))), "", Trim(dbTabFrasiH.fields(1)))
            .Cell(.Rows - 1, 2).Text = IIf(IsNull(Trim(dbTabFrasiH.fields(2))), "", Trim(dbTabFrasiH.fields(2)))
        
         
         .Cell(.Rows - 1, 3).Text = dbTabFrasiH!ID
         dbTabFrasiH.MoveNext
       Next
       .Refresh
       .ReadOnly = True
        .AutoRedraw = True
    End With

End Sub

Public Sub CopyFrasiHGrd2(ByVal Grd2 As Grid, ByVal lId As Long)
Dim i As Integer
Dim t As Integer
    If lId = 0 Then Exit Sub

     With dbTabFrasiH
            
        .Filter = ""
        .Filter = "ID='" & lId & "'"
        If .EOF Then Exit Sub
        .MoveFirst

    End With
    
    
    With Grd2
       ' .DefaultFont.Size = 12 * m_ControlGridFontSize
       For i = 1 To .Rows - 1
            .Cell(i, 2).Text = IIf(IsNull(Trim(dbTabFrasiH.fields(i))), "", Trim(dbTabFrasiH.fields(i)))
       Next
        

        For i = 2 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter ' cellLeftCenter
            
        Next
        
    
    
    End With

End Sub

Public Sub Grd2_FrasiH_LeaveCell(ByVal Grd2 As Grid, ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean, ByVal lRow As Long)


Dim sValue As String
Dim sString As String
Debug.Print "Leave ", Row, Col
With Grd2
    sValue = .Cell(Row, Col).Text
    If Col = 2 Then
        If lRow = Row Then
        
            Select Case Row
                Case 1
                    ' CODE
                    If Len(sValue) = 0 Then
                        PopupMessage 2, "Warning : Code must be a valid value...."
                       
                    End If
                    
                Case 2
             
            End Select
        
        
        
        End If
    End If
End With

Exit Sub

err:
PopupMessage 2, sString
Grd2.Cell(Row, Col).Text = ""
Return
End Sub

Public Function SaveDatabaseFrasiH(ByVal Grd2 As Grid) As Boolean
Dim rc As Boolean
Dim MyNewCode As String


On Error GoTo ERR_SAVE
rc = True
    
    MyNewCode = Trim(Grd2.Cell(1, 2).Text)

    If MyNewCode = "" Then
        PopupMessage 2, "Please Enter a valid Code!"
        Exit Function
    End If
    
    With dbTabFrasiH
    
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
        
            .fields(i) = Trim(Grd2.Cell(i, 2).Text)
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
    
    SaveDatabaseFrasiH = rc
    Exit Function
    
ERR_SAVE:
    rc = False
    MsgBox err.Description
    Resume Next

End Function

Public Sub AddComboFrasiH(ByVal Combo1 As ComboBox)
    Combo1.Clear
    Combo1.AddItem "Code"
    Combo1.ListIndex = 0
End Sub



