Attribute VB_Name = "Database_Pipette"
Option Explicit


Public Function SetGridPipette(ByVal Grd1 As Grid) As Boolean

       '------------------------------------------------
        '       SET TABELLA Pipette
        '------------------------------------------------
    With Grd1
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        .RowHeight(0) = 0
        
        .Cols = 8
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Pipette"
        .Column(1).Width = 150
        .Cell(0, 2).Text = "Volume adjustment"
        .Column(2).Width = 200
        
        .Cell(0, 3).Text = "Characteristic"
        .Column(3).Width = 100
        
    
        .Cell(0, 4).Text = "ID"
        .Column(4).Width = 0
        
        .Cell(0, 5).Text = "Volume Min"
        .Column(5).Width = 0
        .Cell(0, 6).Text = "Volume Max"
        .Column(6).Width = 0
        .Cell(0, 7).Text = "Unit"
        .Column(7).Width = 0
        
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


Public Function FillGridPipette(ByRef Grd As Grid, Optional ByVal Equipment As String, Optional bMainForm As Boolean, Optional ByVal sString As String)
Dim i As Integer
Dim t As Integer
Dim MaxCount As Integer

    On Error GoTo ERR_GRID
    ' --------------------------------------
    '
    '  filtra TabReport e riempi Tabella
    '
    ' --------------------------------------
    If InStr(UCase(Equipment), UCase("Equipment")) Then Equipment = ""
    
    


    With Grd
        .Rows = 1
        .ReadOnly = True
        .AutoRedraw = False
        With dbTabPipette
        
            .filter = ""
            If sString = "Equipment" Then
                If Equipment <> "" Then .filter = "Equipment like '*" & Equipment & "*'"
            'Else
               ' If Equipment <> "" Then .filter = "Equipment like '*" & Equipment & "*'"
            End If
            
            If .EOF Then Exit Function
            MaxCount = .RecordCount
            .MoveFirst
        End With
        
        For i = 1 To MaxCount
            .AddItem "", False
            .Cell(i, 0).Text = i
            .Cell(i, 1).Text = "  " & IIf(IsNull(Trim(dbTabPipette!Equipment)), "", Trim(dbTabPipette!Equipment))

            
            .Cell(i, 1).Alignment = cellLeftCenter
            .Cell(i, 2).Text = "  " & IIf(IsNull(Trim(dbTabPipette!VolumeAdjustment)), "", Trim(dbTabPipette!VolumeAdjustment))
            .Cell(i, 2).Alignment = cellLeftCenter
            .Cell(i, 3).Text = IIf(IsNull(Trim(dbTabPipette!Characteristic)), "", Trim(dbTabPipette!Characteristic))
            .Cell(i, 4).Text = dbTabPipette!ID
            
            
            .Cell(i, 5).Text = IIf(IsNull(Trim(dbTabPipette!VolMin)), "", Trim(dbTabPipette!VolMin))
            .Cell(i, 6).Text = IIf(IsNull(Trim(dbTabPipette!VolMax)), "", Trim(dbTabPipette!VolMax))
            .Cell(i, 7).Text = IIf(IsNull(Trim(dbTabPipette!Unit)), "", Trim(dbTabPipette!Unit))
            
            
            For t = 1 To .Cols - 1
                If bMainForm Then
                    .Cell(i, t).ForeColor = vbColorBlueProgram 'vbColorBlueProgram
                Else
                    .Cell(i, t).ForeColor = vbColorDarkFont
                End If
            Next
                
                If i > 1 Then
                
                If .Cell(i, 1).Text = .Cell(i - 1, 1).Text Then
                   For t = 1 To .Cols - 1
                    .Cell(i, t).BackColor = vbColorTextLightBlue
                   Next
                End If

            End If
            dbTabPipette.MoveNext
        Next
ERR_END:
        If Not (bMainForm) Then .Column(1).AutoFit
        .Column(2).AutoFit
        .AutoRedraw = True
        .Refresh
    End With

    Exit Function
ERR_GRID:
    MessageInfoTime = 2000
    PopupMessage 2, err.Description
    GoTo ERR_END:
End Function



Public Function SetGridEditPipetta(ByRef Grd As Grid) As Boolean
   
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
        
        .Rows = 14
        
        .Cell(1, 1).Text = "  " & "Pipette"
        .Cell(2, 1).Text = "  " & "Volume Adjustment"
        .Cell(3, 1).Text = "  " & "Characteristic"
        .Cell(4, 1).Text = "  " & "Volume MIN"
        .Cell(5, 1).Text = "  " & "Volume MAX"
        .Cell(6, 1).Text = "  " & "Unit"
        .Cell(7, 1).Text = "  " & "Decimal mL"
        .Cell(8, 1).Text = "  " & "Grad. Resolution"
        .Cell(9, 1).Text = "  " & "Acc "
        .Cell(10, 1).Text = "  " & "Acc % Min"
        .Cell(11, 1).Text = "  " & "Acc % Max"
        .Cell(12, 1).Text = "  " & "Immersion Deph"
        .Cell(13, 1).Text = "  " & "Wait Time"
      
        
     
     
     
        
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


Public Sub CopyPipettaGrd2(ByVal Grd2 As Grid, ByVal lId As Long)
    If lId = 0 Then Exit Sub
    Dim i As Integer



     With dbTabPipette
            
        .filter = ""
        .filter = "ID='" & lId & "'"
        If .EOF Then Exit Sub
        .MoveFirst
        
        
    End With
    
   
    With Grd2
       ' .DefaultFont.Size = 12 * m_ControlGridFontSize
       For i = 1 To .Rows - 1
            .Cell(i, 2).Text = (IIf(IsNull(Trim(dbTabPipette.fields(i))), "", Trim(dbTabPipette.fields(i))))
       Next
        

        For i = 2 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter ' cellLeftCenter
            
        Next
           .Cell(4, 2).BackColor = vbColorAzzurrino
        .Cell(5, 2).BackColor = vbColorAzzurrino
        
    
    
    End With

End Sub

Public Sub Grd2_Pipetta_LeaveCell(ByVal Grd2 As Grid, ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean, ByVal lRow As Long)


Dim sValue As String
Dim sString As String
Debug.Print "Leave ", Row, Col
With Grd2
    sValue = .Cell(Row, Col).Text
    If Col = 2 Then
        If lRow = Row Then
        
            Select Case Row
                Case 1
                    ' Pipette
                    If Len(sValue) = 0 Then
                        PopupMessage 2, "Warning : Pipette must be a valid value...."
                       
                    End If
            
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

Public Function SaveDatabasePipette(ByVal Grd2 As Grid, ByVal MyID As Long) As Boolean
Dim rc As Boolean
Dim MyNewEquipment As String
Dim RangeMin As String
Dim RangeMax As String

On Error GoTo ERR_SAVE
rc = True
    MyNewEquipment = Trim(Grd2.Cell(1, 2).Text)
    RangeMin = Trim(Grd2.Cell(12, 2).Text)
    RangeMax = Trim(Grd2.Cell(13, 2).Text)
    
    If MyNewEquipment = "" Then
        PopupMessage 2, "Please Enter a valid Pipette!"
        Exit Function
    End If
    
    With dbTabPipette
        .filter = ""
        If MyID > 0 Then
            .filter = "ID='" & MyID & "'"
        Else
        
            .filter = "Equipment='" & MyNewEquipment & "'"
        End If
        If .EOF Then
        
            .AddNew
         
        Else
            If F_MsgBox.DoShow("Pipette already exsist. Replace Info?") Then
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
        PopupMessage 2, "Pipette : " & MyNewEquipment & " saved!"
    Else
        PopupMessage 2, "Warning : a problem occurred, please check all entries before Save"
    End If
    
    SaveDatabasePipette = rc
    Exit Function
    
ERR_SAVE:
    rc = False
    MsgBox err.Description
    Resume Next:

End Function


Public Sub AddComboPipetta(ByVal Combo1 As ComboBox)

    Combo1.Clear
    Combo1.AddItem "Equipment"
    Combo1.ListIndex = 0
End Sub


