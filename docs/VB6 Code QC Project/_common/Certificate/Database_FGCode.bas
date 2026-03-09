Attribute VB_Name = "Database_FGCode"
Option Explicit


Public Function SetGridFGCode(ByRef Grd1 As Grid) As Boolean

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
        .Cell(0, 1).Text = "FG Code"
        .Column(1).Width = 150
        .Cell(0, 2).Text = "Description"
        .Column(2).Width = 200
        .Cell(0, 3).Text = "ID"
        .Column(3).Width = 0
        
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


Public Function FillGridFGCode(ByRef Grd As Grid, Optional ByVal Code As String, Optional bMainForm As Boolean, Optional ByVal sString As String)
Dim i As Integer
Dim t As Integer
Dim x As Integer
Dim MaxCount As Integer

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
        With dbTabFinishGood
        
            .filter = ""

            If Code <> "" Then .filter = "Code like '*" & Code & "*'"
        
            
            If .EOF Then Exit Function
            MaxCount = .RecordCount
            .MoveFirst
        End With
        i = 0
        For x = 1 To MaxCount
            .AddItem "", False
            i = i + 1
            .Cell(i, 0).Text = i
            
            
            
            
            .Cell(i, 1).Text = "  " & IIf(IsNull(Trim(dbTabFinishGood!Code)), "", Trim(dbTabFinishGood!Code))
            .Cell(i, 1).Alignment = cellLeftCenter
            .Cell(i, 2).Text = "  " & IIf(IsNull(Trim(dbTabFinishGood!Description)), "", Trim(dbTabFinishGood!Description))
            .Cell(i, 2).Alignment = cellLeftCenter
            .Cell(i, 3).Text = dbTabFinishGood!ID
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
cont:
            dbTabFinishGood.MoveNext
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
 
    PopupMessage 2, Err.Description
    Resume Next
End Function



Public Function SetGridEditFGCode(ByRef Grd As Grid) As Boolean
Dim i As Integer
Dim MaxRows As Long
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
        
        
        MaxRows = dbTabFinishGood.Fields.Count
        .Rows = MaxRows

       
        For i = 1 To .Rows - 1
        
         .Cell(i, 1).Text = dbTabFinishGood.Fields(i).Name
         
        Next
        
       
        
        
        For i = 1 To .Rows - 1
        
            .Cell(i, 1).BackColor = &HF0F0F0 'vbColorUnabled
            .Cell(i, 1).ForeColor = vbColorDarkFont 'vbColorDarkFont 'vbColorForeFixed  ' vbColorBlueProgram
            .Cell(i, 1).FontBold = False
            .Cell(i, 1).Locked = True
            .Cell(i, 2).ForeColor = vbColorDarkFont
             .Cell(i, 2).WrapText = True

        Next


        
        
        .ReadOnly = False
        .AutoRedraw = True
        .Refresh
        
    End With
End Function


Public Sub CopyFGCodeGrd2(ByVal Grd2 As Grid, ByVal lId As Long)
    If lId = 0 Then Exit Sub
    Dim i As Integer



     With dbTabFinishGood
            
        .filter = ""
        .filter = "ID='" & lId & "'"
        If .EOF Then Exit Sub
        .MoveFirst
        
        
    End With
    
   
    With Grd2
       ' .DefaultFont.Size = 12 * m_ControlGridFontSize
       For i = 1 To .Rows - 1
            .Cell(i, 2).Text = (IIf(IsNull(Trim(dbTabFinishGood.Fields(i))), "", Trim(dbTabFinishGood.Fields(i))))
       Next
        

        For i = 2 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter ' cellLeftCenter
            
        Next
           .Cell(11, 2).BackColor = vbColorAzzurrino
        .Cell(12, 2).BackColor = vbColorAzzurrino
        
    
    
    End With

End Sub

Public Sub Grd2_FGCode_LeaveCell(ByVal Grd2 As Grid, ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean, ByVal lRow As Long)


Dim sValue As String
Dim sString As String
Debug.Print "Leave ", Row, Col
With Grd2
    sValue = .Cell(Row, Col).Text
    If Col = 2 Then
        If lRow = Row Then
        
            Select Case Row
                Case 1
                    ' FGCode
                    If Len(sValue) = 0 Then
                        PopupMessage 2, "Warning : FGCode must be a valid value...."
                       
                    End If
            
            End Select
        
        
        
        End If
    End If
End With

Exit Sub

Err:
PopupMessage 2, sString
Grd2.Cell(Row, Col).Text = ""
Return
End Sub

Public Function SaveDatabaseFGCode(ByVal Grd2 As Grid) As Boolean
Dim rc As Boolean
Dim MyNewFGCode As String
Dim RangeMin As String
Dim RangeMax As String

On Error GoTo ERR_SAVE
rc = True
    MyNewFGCode = Trim(Grd2.Cell(1, 2).Text)
    RangeMin = Trim(Grd2.Cell(12, 2).Text)
    RangeMax = Trim(Grd2.Cell(13, 2).Text)
    
    If MyNewFGCode = "" Then
        PopupMessage 2, "Please Enter a valid FGCode!"
        Exit Function
    End If
    
    With dbTabFinishGood
        .filter = ""
        .filter = "Code='" & MyNewFGCode & "'"
        If .EOF Then
        
            .AddNew
         
        Else
            If F_MsgBox.DoShow("FGCode already exsist. Replace Info?") Then
            Else
                Exit Function
            End If
            
        End If
        
        Dim i As Integer
        For i = 1 To Grd2.Rows - 1
        
            .Fields(i) = Trim(Grd2.Cell(i, 2).Text)
        Next
        
        .Update
    End With


ERR_END:
    On Error GoTo 0
    If rc Then
        PopupMessage 2, "FGCode : " & MyNewFGCode & " saved!"
    Else
        PopupMessage 2, "Warning : a problem occurred, please check all entries before Save"
    End If
    
    SaveDatabaseFGCode = rc
    Exit Function
    
ERR_SAVE:
    rc = False
    MsgBox Err.Description
    Resume Next:

End Function
Public Function CancellaFGCodeByID(ByVal dbTab As ADODB.Recordset, ByVal MyID As Long) As Boolean
Dim rc As Boolean
Dim FGCode As String
On Error GoTo ERR_CAN



    If MyID = 0 Then Exit Function
    
    
    rc = True
    With dbTab
        .filter = ""
        .filter = "ID='" & MyID & "'"
        If .EOF Then
            rc = False
        Else
        
        FGCode = Trim(!Code)
      
        .Delete
        .Update
        End If
    End With
ERR_END:
    On Error GoTo 0
    If rc Then
        PopupMessage 2, "Record Deleted!"
    Else
        PopupMessage 2, "Warning : a problem occurred...."
    End If
    
    CancellaFGCodeByID = rc
    Exit Function
    
ERR_CAN:
    rc = False
    MsgBox Err.Description
    Resume ERR_END:
End Function

Public Sub AddComboFGCode(ByVal Combo1 As ComboBox)

    Combo1.Clear
    Combo1.AddItem "FGCode"
    Combo1.AddItem "MRFGCode"
    Combo1.ListIndex = 0
End Sub



