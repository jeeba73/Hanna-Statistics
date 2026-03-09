Attribute VB_Name = "Database_MR"
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
    MsgBox Err.Description
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
        .Rows = 21

        .Cell(1, 1).Text = "  " & "Code"
        .Cell(2, 1).Text = "  " & "Description"
        .Cell(3, 1).Text = "  " & "Supplier"
        .Cell(4, 1).Text = "  " & "MR NMP"
        .Cell(5, 1).Text = "  " & "Location"
        .Cell(6, 1).Text = "  " & "Physical State"
        .Cell(7, 1).Text = "  " & "Density"
        .Cell(8, 1).Text = "  " & "MR Unit"
        .Cell(9, 1).Text = "  " & "MR Parameter"
        .Cell(10, 1).Text = "  " & "FW MR Parameter"
        .Cell(11, 1).Text = "  " & "Storage"
        .Cell(12, 1).Text = "  " & "MinQTY"
        .Cell(14, 1).Text = "  " & "Stock QTY"
        .Cell(15, 1).Text = "  " & "Stock Unit"
        .Cell(16, 1).Text = "  " & "Last Updated"
        
        .Cell(17, 1).Text = "  " & "Reduction Expiration Days"
        
        .Cell(18, 1).Text = "  " & "Classification"
        .Cell(19, 1).Text = "  " & "MR Purity (%)"
        .Cell(20, 1).Text = "  " & "MR Value"
        
        
        .RowHeight(13) = 0
 
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
Public Sub CopyChemicalRMGrd1(ByRef Grd As Grid, Optional ByVal Code As String, Optional bMainForm As Boolean, Optional ByVal sString As String)
Dim i As Integer
Dim t As Integer
Dim filterString As String
Dim strCritical As String

    filterString = UCase(Replace(Trim(Code), "'", "''"))
    

     
     With dbTabMR
        
        .filter = ""
        If filterString <> "" And Code <> "" Then
             .filter = "Code like '*" & filterString & "*'"
             If .EOF Then
                .filter = "Code = '" & filterString & "'"
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

Public Sub CopyChemicalRMGrd2(ByVal Grd2 As Grid, ByVal lId As Long, ByRef MRLocation As String, ByRef MRDescription As String)
Dim i As Integer
Dim t As Integer
    If lId = 0 Then Exit Sub
    
    On Error GoTo ERR_CH

     With dbTabMR
            
        .filter = ""
        .filter = "ID='" & lId & "'"
        If .EOF Then Exit Sub
        .MoveFirst

    End With
    
    
    With Grd2
       ' .DefaultFont.Size = 12 * m_ControlGridFontSize
       For i = 1 To .Rows - 1
            .Cell(i, 2).Text = (IIf(IsNull(Trim(dbTabMR.fields(i))), "", Trim(dbTabMR.fields(i))))
       Next

        
       
       Dim strUnit As String
       
        strUnit = .Cell(8, 2).Text
        
        If Trim(.Cell(6, 2).Text) = "" Or IsNull(Trim(.Cell(6, 2).Text)) Then
        
            If InStr(UCase(strUnit), "L") Then
            
                dbTabMR!PhysicalState = "L"
                
            Else
                dbTabMR!PhysicalState = "S"
            
            End If
            .Cell(6, 2).Text = dbTabMR!PhysicalState
            dbTabMR.Update
            
        End If
        
        ' stock unit!!!!
        
        If Trim(.Cell(15, 2).Text) = "" Or IsNull(Trim(.Cell(15, 2).Text)) Then
        
            If InStr(UCase(strUnit), "L") Then
                dbTabMR!STOCK_UNIT = "L"
            Else
                dbTabMR!STOCK_UNIT = "g"
            End If
            .Cell(15, 2).Text = dbTabMR!STOCK_UNIT
            dbTabMR.Update
        End If
        
        If Trim(.Cell(17, 2).Text) = "" Or IsNull(Trim(.Cell(17, 2).Text)) Then
        

            dbTabMR!ReductionExpDays = "120"
            
           
            .Cell(17, 2).Text = dbTabMR!ReductionExpDays
            dbTabMR.Update
        End If
        
        If Trim(.Cell(5, 2).Text) = "" Or IsNull(Trim(.Cell(5, 2).Text)) Then

            dbTabMR!Location = GetLocation(Trim(.Cell(1, 2).Text))

            .Cell(5, 2).Text = dbTabMR!Location
            dbTabMR.Update
        End If
        
        
        
        For i = 2 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter ' cellLeftCenter
            
        Next
        
        If .Cell(6, 2).Text = "" Then
            ' .Cell(5, 2).BackColor = vbcolorred
        Else
             .Cell(6, 2).BackColor = vbColorAzzurrino
        End If
        
        If .Cell(13, 2).Text = "" Then
           '  .Cell(12, 2).BackColor = vbcolorred
        Else
             .Cell(13, 2).BackColor = vbColorAzzurrino
        End If
               
   
        .Cell(14, 2).BackColor = &HEDEBE7
        .Cell(15, 2).BackColor = &HEDEBE7
        .Cell(16, 2).BackColor = &HEDEBE7
        
     
        .Cell(14, 2).Locked = True
        .Cell(15, 2).Locked = True
        .Cell(16, 2).Locked = True
        
        
        MRLocation = Trim(.Cell(5, 2).Text)
        MRDescription = Trim(.Cell(2, 2).Text)
        
    End With
    
    
    
ERR_END:
    On Error GoTo 0
    
    
    Exit Sub
ERR_CH:
    MsgBox Err.Description
    Resume Next

End Sub

Public Sub Grd2_ChemicalRM_LeaveCell(ByVal Grd2 As Grid, ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean, ByVal lRow As Long)


Dim sValue As String
Dim sString As String
Dim Um As String
Debug.Print "Leave ", Row, Col
With Grd2
    sValue = .Cell(Row, Col).Text
    
    Um = .Cell(15, Col).Text
    
    If Col = 2 Then
        If lRow = Row Then
        
            Select Case Row
                Case 1
                    ' CODE
                    If Len(sValue) = 0 Then
                        PopupMessage 2, "Warning : Code must be a valid value...."
                       
                    End If
                Case 6
                    
                    If Len(sValue) = 0 Or (sValue <> "S" And sValue <> "L") Then
                        PopupMessage 2, "Warning : Physical State must be a S or L ...."
                    Else
                       If sValue = "S" Then .Cell(15, Col).Text = "g"
                       If sValue = "L" Then .Cell(15, Col).Text = "mL"
                    End If
                    
                Case 7
                    ' DENSITY
                    If IsNumeric(sValue) Then
                        
            
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

Err:
PopupMessage 2, sString
Grd2.Cell(Row, Col).Text = ""
Return
End Sub

Public Function SaveDatabaseChemicalRM(ByVal Grd2 As Grid, ByRef MRLocation As String, ByRef MRDescription As String) As Boolean
Dim rc As Boolean
Dim mrc As Boolean
Dim MyNewCode As String
Dim bMix As Boolean
Dim NewMRLocation As String
Dim NewMRDescription As String
Dim stockUnit As String

On Error GoTo ERR_SAVE
rc = True
    
    MyNewCode = Trim(Grd2.Cell(1, 2).Text)

    If MyNewCode = "" Then
        PopupMessage 2, "Please Enter a valid Code!"
        Exit Function
    End If
    
    With dbTabMR
    
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
        
        NewMRLocation = Trim(Grd2.Cell(5, 2).Text)
        NewMRDescription = Trim(Grd2.Cell(2, 2).Text)
        
        stockUnit = Trim(Grd2.Cell(15, 2).Text)
        
        For i = 1 To Grd2.Rows - 1
           
            If i = 13 Then
            
               GoTo cont:
            End If
            
            Debug.Print i & "   " & .fields(i).Name & "     " & Trim(Grd2.Cell(i, 2).Text)
            If i = 16 Then
                .fields(i) = Now()
           Else
                .fields(i) = Trim(Grd2.Cell(i, 2).Text)
                
            End If
            
                    
cont:
        Next

        .Update
    End With
    
    
    'If Trim(UCase(MRLocation)) = Trim(UCase(NewMRLocation)) Then
    'Else
    
    '------------------------------------------------
    ' correggo Stock unit , Location e Description
    '------------------------------------------------
    mrc = ModifyWarehouseData(MyNewCode, NewMRLocation, stockUnit, NewMRDescription)
    
    
    If mrc And Trim(UCase(MRLocation)) <> Trim(UCase(NewMRLocation)) Then
        PopupMessage 2, "All bottles moved to new location!" & vbCrLf & "New Location : " & NewMRLocation, , , MyNewCode
    End If
        
    MRLocation = NewMRLocation
 
    
    If mrc And Trim(UCase(MRDescription)) <> Trim(UCase(NewMRDescription)) Then
        PopupMessage 2, MyNewCode & " : All bottles Description Changed!" & vbCrLf & "New Description : " & NewMRDescription, , , MyNewCode
    End If
        
    MRDescription = NewMRDescription
    
    
   
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
    MsgBox Err.Description
    Resume Next:

End Function

Public Function ModifyWarehouseData(ByVal Code As String, ByVal Location As String, ByVal stockUnit As String, ByVal Description As String) As Boolean
Dim rc As Boolean
Dim i As Integer
    
  On Error GoTo ERR_SAVE
  
  rc = False
  
  If Location = "" Then GoTo ERR_END:
  
  With dbTabMRWarehouse
    .filter = ""
    .filter = "Code='" & Code & "' and bClosed=false"
    If .EOF Then
        GoTo ERR_END
    
    End If
    rc = True
    .MoveFirst
    For i = 1 To .RecordCount
        !Location = Location
        !Description = Description
        !stockUnit = stockUnit
        .MoveNext
    Next
   
  End With
  
  
    
     
ERR_END:
    On Error GoTo 0
 
    
    
    ModifyWarehouseData = rc
    Exit Function
    
ERR_SAVE:
    rc = False
    MsgBox Err.Description
    Resume Next:

End Function

Public Sub AddComboChemicalRM(ByVal Combo1 As ComboBox)
    Combo1.Clear
    Combo1.AddItem "Code"
    Combo1.ListIndex = 0
End Sub



Public Function AddPurityMR(ByVal MRCode As String, ByVal Purity As String) As Boolean

If IsNumeric(Purity) Then

Else
    Purity = 100
End If

    With dbTabMR
        .filter = ""
        .filter = "Code='" & MRCode & "'"
        If .EOF Then
        Else
            
            !MRPurity = Purity
            .Update
        End If
    End With
    
    

End Function


Public Function CreateMRExp(SupplierEXP As String, ReductionExpDays As String) As String
    '-------------------------------------------------
    ' MREXP
    ' Supplier EXP ( Date ) - ReductionExpDays
    '-------------------------------------------------
    If IsDate(SupplierEXP) And IsNumeric(ReductionExpDays) Then
       
       CreateMRExp = FormatDataLAT(DateAdd("d", -CInt(ReductionExpDays), SupplierEXP))
       
    End If

End Function
