Attribute VB_Name = "Database_Classification"
Option Explicit
Public Function SetGridClassification(ByRef Grd1 As Grid) As Boolean


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
        .Column(1).Width = 200
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

Public Function SetGridEditClassification(ByRef Grd As Grid) As Boolean
   
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
        .Rows = 8

        .Cell(1, 1).Text = "  " & "Code"
        .Cell(2, 1).Text = "  " & "Description"
        .Cell(3, 1).Text = "  " & "Cas"
        .Cell(4, 1).Text = "  " & "Index"
        .Cell(5, 1).Text = "  " & "Cee"
        .Cell(6, 1).Text = "  " & "Recipe"
        .Cell(7, 1).Text = "  " & "Classification  "
        
        
        
        
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


        .RowHeight(3) = 0
        .RowHeight(4) = 0
        .RowHeight(5) = 0
        .RowHeight(6) = 0
        '.Cell(7, 1).FontBold = True
        '.Cell(7, 1).ForeColor = vbColorTextBlue
       '.Cell(7, 1).FontUnderline = True
        .ReadOnly = False
        .AutoRedraw = True
        .Refresh
        
    End With
End Function
Public Sub CopyClassificationGrd1(ByRef Grd As Grid, Optional ByVal Code As String, Optional bMainForm As Boolean, Optional ByVal sString As String)
Dim i As Integer
Dim t As Integer
Dim bMancaClassification As Boolean
Dim strClassification As String
Dim strCode As String
Dim classID As Long

    With dbTabCode
    
        .filter = ""
        If sString = "Recipe" Then
            If Code <> "" Then .filter = "Recipe like '*" & Code & "*'"
        Else
            If Code <> "" Then .filter = "Code like '*" & Code & "*'"
        End If
        
        If .EOF Then Exit Sub
       ' MaxCount = .RecordCount
        .MoveFirst
    End With
        

     'With dbTabCodeClassification
     '
     '   .filter = ""
     '   If sString <> "" And Code <> "" Then
     '       '.Filter = "ProductionWay ='" & Replace(Trim(Code), "'", "''") & "'"
     '        .filter = "Code like '*" & Replace(Trim(Code), "'", "''") & "*'"
     '   End If
     '   If .EOF Then Exit Sub
     '   .MoveFirst

     'End With
    
    
    With Grd
       .AutoRedraw = False
       For i = 1 To dbTabCode.RecordCount
        .AddItem "", False
            .Cell(.Rows - 1, 1).Text = IIf(IsNull(Trim(dbTabCode!Code)), "", Trim(dbTabCode!Code))
            .Cell(.Rows - 1, 2).Text = IIf(IsNull(Trim(dbTabCode!ProductName)), "", Trim(dbTabCode!ProductName))
        
            strCode = IIf(IsNull(Trim(dbTabCode.fields(1))), "", Trim(dbTabCode.fields(1)))
            strClassification = GetClassificationByCode(strCode, classID)
            
            bMancaClassification = IIf(Len(strClassification) > 0, False, True)
         
         .Cell(.Rows - 1, 3).Text = classID
         
         If bMancaClassification Then
            
            .Cell(.Rows - 1, 1).ForeColor = vbColorOrange
            .Cell(.Rows - 1, 2).ForeColor = vbColorOrange
            
         End If
         
         dbTabCode.MoveNext
       Next
       .Refresh
       .ReadOnly = True
        .AutoRedraw = True
    End With

End Sub

Private Function GetClassificationByCode(ByVal Code As String, ByRef ID As Long) As String
With dbTabCodeClassification
    
        .filter = ""
        If Code <> "" Then
            '.Filter = "ProductionWay ='" & Replace(Trim(Code), "'", "''") & "'"
             .filter = "Code ='" & Replace(Trim(Code), "'", "''") & "'"
        End If
        If .EOF Then
        Else
            ID = !ID
            GetClassificationByCode = IIf(IsNull(Trim(dbTabCodeClassification!Phrases)), "", Trim(dbTabCodeClassification!Phrases))
            
        End If
        

     End With
End Function

Public Sub CopyClassificationGrd2(ByVal Grd2 As Grid, ByVal lId As Long)
Dim i As Integer
Dim t As Integer
    If lId = 0 Then Exit Sub

     With dbTabCodeClassification
            
        .filter = ""
        .filter = "ID='" & lId & "'"
        If .EOF Then Exit Sub
        .MoveFirst

    End With
    
    
    With Grd2
       .AutoRedraw = False
       For i = 1 To .Rows - 1
            .Cell(i, 2).Text = IIf(IsNull(Trim(dbTabCodeClassification.fields(i))), "", Trim(dbTabCodeClassification.fields(i)))
       Next
        

        For i = 2 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter ' cellLeftCenter
            
        Next
        
        .RowHeight(3) = 0
        .RowHeight(4) = 0
        .RowHeight(5) = 0
        .RowHeight(6) = 0
        
        .Refresh
        .AutoRedraw = True
    
    
    End With

End Sub

Public Sub Grd2_Classification_LeaveCell(ByVal Grd2 As Grid, ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean, ByVal lRow As Long)


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

Public Function SaveDatabaseClassification(ByVal Grd2 As Grid) As Boolean
Dim rc As Boolean
Dim MyNewCode As String


On Error GoTo ERR_SAVE
rc = True
    
    MyNewCode = Trim(Grd2.Cell(1, 2).Text)

    If MyNewCode = "" Then
        PopupMessage 2, "Please Enter a valid Code!"
        Exit Function
    End If
    
    With dbTabCodeClassification
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
       ' For i = 1 To Grd2.Rows - 1
        
            .fields(1) = Trim(Grd2.Cell(1, 2).Text)
            .fields(2) = Trim(Grd2.Cell(2, 2).Text)
            .fields(7) = Trim(Grd2.Cell(7, 2).Text)
            .fields(8) = Now()
        
        .Update
    End With


ERR_END:
    On Error GoTo 0
    If rc Then
        PopupMessage 2, "Code : " & MyNewCode & " saved!"
    Else
        PopupMessage 2, "Warning : a problem occurred, please check all entries before Save"
    End If
    
    SaveDatabaseClassification = rc
    Exit Function
    
ERR_SAVE:
    rc = False
    
    MsgBox err.Description
    Resume Next

End Function

Public Sub AddComboClassification(ByVal Combo1 As ComboBox)
    Combo1.Clear
    Combo1.AddItem "Code"
    Combo1.AddItem "Recipe"
    Combo1.ListIndex = 0
End Sub


Public Function SetCodeClassification(ByVal Code As String, ByVal Description As String)
    
    With dbTabCodeClassification
        .filter = ""
        .filter = "Code='" & Code & "'"
        If .EOF Then
        
            .AddNew
        Else
           
        End If
        
        !Code = Code
        !Name = Description
        
        .Update
    End With

End Function


Public Function DeleteCodeClassification(ByVal Code As String)
 
 With dbTabCodeClassification
    .filter = ""
    .filter = "Code='" & Code & "'"
    If .EOF Then
    
       
    Else
       .Delete
       .Update
    End If

End With

End Function

Public Function SetAllClassificationByRecipe()
Dim i As Integer
 Dim Recipe As String
Dim ID As Long
Dim Classification As String
With dbTabRecipe
    .filter = ""
    If .EOF Then
    Else
        .MoveFirst
        For i = 1 To .RecordCount
            ID = !ID
            Classification = IIf(IsNull(Trim(!Classification)), "", Trim(!Classification))
            Recipe = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
            
            If Classification <> "" Then
                Call ClassificationInHannacode(True, ID, Recipe, , Classification)
            End If
            .MoveNext
        Next
    End If
SaveSetting App.Title, "Classification", "SetAllClassificationByRecipe", True
End With

End Function

Public Function SetCodeClassificationByRecipe(ByVal rc As Boolean, ByVal ID As Long, Optional ByVal HannaCode As String, Optional ByVal Classification As String)
Dim Recipe As String

    With dbTabRecipe
        .filter = ""
        .filter = "ID='" & ID & "'"
        If .EOF Then
        Else
            Recipe = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
            Classification = IIf(IsNull(Trim(!Classification)), "", Trim(!Classification))
            Call ClassificationInHannacode(True, ID, Recipe, , Classification)
        End If
    End With
    
End Function


Public Function ClassificationInHannacode(ByVal rc As Boolean, ByVal ID As Long, ByVal Recipe As String, Optional ByVal HannaCode As String, Optional ByVal Classification As String)


Dim i As Integer
Dim HannaCodeDescription As String

    
    
   If Classification <> "" Then GoTo hannac:
   

    
    
    
    If Recipe <> "" And Classification <> "" Then

hannac:

        '----------------------------------------------------
        ' ho Classification e Recipe cerco l'HannaCode....
        '----------------------------------------------------
        
        If HannaCode <> "" Then GoTo cont:
        
        With dbTabCode
            .filter = ""
            .MoveFirst
            For i = 1 To .RecordCount
                If Trim(!Recipe) = Trim(Recipe) Then ' Or Trim(!Mix1) = Recipe Or Trim(!Mix2) = Recipe Then
                
                    HannaCode = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
                    HannaCodeDescription = IIf(IsNull(Trim(!ProductName)), "", Trim(!ProductName))
                    If rc Then GoTo cont:
back:
                End If
    
                .MoveNext
            Next
            Exit Function
        End With
cont:
        '----------------------------------------------------
        ' ho HannaCode e inserisco in TabCodeClassification
        '----------------------------------------------------
        
        With dbTabCodeClassification
            .filter = ""
            .filter = "Code='" & HannaCode & "'"
            If .EOF Then .AddNew
            !Code = HannaCode
            !Name = HannaCodeDescription
            !Recipe = Recipe
            !Phrases = Classification
            !DateModified = Now()
            .Update
        End With
        
        If rc Then GoTo back:
        
    End If
End Function
