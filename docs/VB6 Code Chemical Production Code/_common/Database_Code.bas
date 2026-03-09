Attribute VB_Name = "Database_Code"
Option Explicit
Public Function SetGridCode(ByRef Grd1 As Grid) As Boolean


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
        .Cell(0, 1).Text = "Code SFG"
        .Column(1).Width = 150
        .Cell(0, 2).Text = "Description"
        .Column(2).Width = 280
       
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


Public Function FillGridCode(ByRef Grd As Grid, Optional ByVal Code As String, Optional bMainForm As Boolean, Optional ByVal sString As String, Optional ByVal strLine As String)
Dim i As Integer
Dim t As Integer
Dim MaxCount As Integer
Dim filterString As String

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
        With dbTabCode
        
            .Close
            .Open "SELECT *  FROM TabCode ORDER BY Code", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
   
        
            .filter = ""

            
            If sString = "Recipe" Then
                If Code <> "" Then filterString = IIf(filterString <> "", filterString & " AND ", "") & "Recipe like '*" & Code & "*'"
            Else
                If Code <> "" Then filterString = IIf(filterString <> "", filterString & " AND ", "") & "Code like '*" & Code & "*'"
            End If
            
            
            If LCase(strLine) = "all lines" Or strLine = "" Then

            Else
            
                filterString = filterString & IIf(filterString = "", "", " and ") & "line like '*" & Replace(Trim(strLine), "'", "''") & "*'"
                
                
                
               ' Grid4.Column(2).Width = 0
            End If
            
            .filter = filterString
        
            
            If .EOF Then Exit Function
            MaxCount = .RecordCount
            .MoveFirst
        End With
        
        For i = 1 To MaxCount
            .AddItem "", False
            .Cell(i, 0).Text = i
            .Cell(i, 1).Text = "  " & IIf(IsNull(Trim(dbTabCode!Code)), "", Trim(dbTabCode!Code))

           
            .Cell(i, 1).Alignment = cellLeftCenter
            .Cell(i, 2).Text = "  " & IIf(IsNull(Trim(dbTabCode!ProductName)), "", Trim(dbTabCode!ProductName))
            .Cell(i, 2).Alignment = cellLeftCenter
            .Cell(i, 3).Text = dbTabCode!ID
            For t = 1 To .Cols - 1
                If bMainForm Then
                    .Cell(i, t).ForeColor = vbColorTextDarkBlue
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
            dbTabCode.MoveNext
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
    GoTo ERR_END:
End Function




Public Function SetGridEditCode(ByRef Grd As Grid) As Boolean
   
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
        .Rows = 18

        .Cell(1, 1).Text = "  " & "Hanna Code"
        .Cell(2, 1).Text = "  " & "Line"
        .Cell(3, 1).Text = "  " & "Std Q.ty to produce per Month"
        .Cell(4, 1).Text = "  " & "Rev loaded in printer"
        .Cell(5, 1).Text = "  " & "Product Name"
        .Cell(6, 1).Text = "  " & "Recipe #"
        .Cell(7, 1).Text = "  " & "Mix 1"
        .Cell(8, 1).Text = "  " & "Mix 2"
        .Cell(9, 1).Text = "  " & "Recipe Rev"
        .Cell(10, 1).Text = "  " & "Exp (years)"
        .Cell(11, 1).Text = "  " & "Um"
        .Cell(12, 1).Text = "  " & "Q.ty"
        .Cell(13, 1).Text = "  " & "Min Q.ty"
        .Cell(14, 1).Text = "  " & "Max Q.ty"
        .Cell(15, 1).Text = "  " & "Uncertantly from CoA"
        .Cell(16, 1).Text = "  " & "Procedure"
        .Cell(17, 1).Text = "  " & "Procedure Rev"
      
         .RowHeight(4) = 0
     
     
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


        
        .RowHeight(9) = 0
        .ReadOnly = False
        .AutoRedraw = True
        .Refresh
        
    End With
End Function


Public Sub CopyCodeGrd2(ByVal Grd2 As Grid, ByVal lId As Long)
    If lId = 0 Then Exit Sub
    Dim i As Integer
    
     With dbTabCode
     
        .filter = ""
        .filter = "ID='" & lId & "'"
        If .EOF Then Exit Sub
        .MoveFirst
        
    End With
    
    
   
        '.Cell(1, 1).Text = "  " & "Hanna Code"
        '.Cell(2, 1).Text = "  " & "Line"
        '.Cell(3, 1).Text = "  " & "Std"
        '.Cell(4, 1).Text = "  " & "Rev loaded in printer"
        '.Cell(5, 1).Text = "  " & "Product Name"
        '.Cell(6, 1).Text = "  " & "Recipe #"
        '.Cell(7, 1).Text = "  " & "Mix 1"
        '.Cell(8, 1).Text = "  " & "Mix 2"
        '.Cell(9, 1).Text = "  " & "Recipe Rev"
        '.Cell(10, 1).Text = "  " & "Exp (years)"
        '.Cell(11, 1).Text = "  " & "Um"
        '.Cell(12, 1).Text = "  " & "Q.ty"
        '.Cell(13, 1).Text = "  " & "Min Q.ty"
        '.Cell(14, 1).Text = "  " & "Max Q.ty"
        '.Cell(15, 1).Text = "  " & "Uncertantly from CoA"
        '.Cell(16, 1).Text = "  " & "Procedure"
        '.Cell(17, 1).Text = "  " & "Procedure Rev"
       
    
    
    
   
    With Grd2
       ' .DefaultFont.Size = 12 * m_ControlGridFontSize

        .Cell(1, 2).Text = IIf(IsNull(Trim(dbTabCode!Code)), "", Trim(dbTabCode!Code))
        .Cell(2, 2).Text = IIf(IsNull(Trim(dbTabCode!Line)), "", Trim(dbTabCode!Line))
        .Cell(3, 2).Text = IIf(IsNull(Trim(dbTabCode!STD)), "", Trim(dbTabCode!STD))
        .Cell(5, 2).Text = IIf(IsNull(Trim(dbTabCode!ProductName)), "", Trim(dbTabCode!ProductName))
        .Cell(6, 2).Text = IIf(IsNull(Trim(dbTabCode!Recipe)), "", Trim(dbTabCode!Recipe))
        .Cell(7, 2).Text = IIf(IsNull(Trim(dbTabCode!Mix1)), "", Trim(dbTabCode!Mix1))
        .Cell(8, 2).Text = IIf(IsNull(Trim(dbTabCode!Mix2)), "", Trim(dbTabCode!Mix2))
        .Cell(9, 2).Text = IIf(IsNull(Trim(dbTabCode!RecipeRev)), "", Trim(dbTabCode!RecipeRev))
        .Cell(10, 2).Text = IIf(IsNull(Trim(dbTabCode!Exp)), "", Trim(dbTabCode!Exp))
        .Cell(11, 2).Text = IIf(IsNull(Trim(dbTabCode!Um)), "", Trim(dbTabCode!Um))
        .Cell(12, 2).Text = IIf(IsNull(Trim(dbTabCode!Qty)), "", Trim(dbTabCode!Qty))
        .Cell(13, 2).Text = IIf(IsNull(Trim(dbTabCode!MinQty)), "", Trim(dbTabCode!MinQty))
        .Cell(14, 2).Text = IIf(IsNull(Trim(dbTabCode!MaxQty)), "", Trim(dbTabCode!MaxQty))
        .Cell(15, 2).Text = IIf(IsNull(Trim(dbTabCode!UncertantlyFromCoA)), "", Trim(dbTabCode!UncertantlyFromCoA))
        .Cell(16, 2).Text = IIf(IsNull(Trim(dbTabCode!Procedure)), "", Trim(dbTabCode!Procedure))
        .Cell(17, 2).Text = IIf(IsNull(Trim(dbTabCode!ProcedureRev)), "", Trim(dbTabCode!ProcedureRev))
     
  
        .RowHeight(4) = 0
        
        

        For i = 2 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter ' cellLeftCenter
            
        Next
           .Cell(11, 2).BackColor = vbColorAzzurrino
        .Cell(12, 2).BackColor = vbColorAzzurrino
        
    
    
    End With

End Sub

Public Sub Grd2_Code_LeaveCell(ByVal Grd2 As Grid, ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean, ByVal lRow As Long)


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

Public Function SaveDatabaseCode(ByVal Grd2 As Grid) As Boolean
Dim rc As Boolean
Dim MyNewCode As String
Dim RangeMin As String
Dim RangeMax As String

On Error GoTo ERR_SAVE
rc = True
    MyNewCode = Trim(Grd2.Cell(1, 2).Text)
    RangeMin = Trim(Grd2.Cell(13, 2).Text)
    RangeMax = Trim(Grd2.Cell(14, 2).Text)
    
    If MyNewCode = "" Then
        PopupMessage 2, "Please Enter a valid Code!"
        Exit Function
    End If
    
    With dbTabCode
        .filter = ""
        .filter = "Code='" & MyNewCode & "'"
        If .EOF Then
        
            .AddNew
            
            Call SetCodeClassification(MyNewCode, Trim(Grd2.Cell(2, 2).Text))
            
        Else
            If F_MsgBox.DoShow("Code already exsist. Replace Info?") Then
            Else
                Exit Function
            End If
            
        End If
        
        Dim i As Integer
        
        
       ' For i = 1 To Grd2.Rows - 1
          '  .fields(i) = Trim(Grd2.Cell(i, 2).Text)
       ' Next
       
        !Code = Grd2.Cell(1, 2).Text
        !Line = Grd2.Cell(2, 2).Text
        !STD = Grd2.Cell(3, 2).Text
        !ProductName = Grd2.Cell(5, 2).Text
        !Recipe = Grd2.Cell(6, 2).Text
        !Mix1 = Grd2.Cell(7, 2).Text
        !Mix2 = Grd2.Cell(8, 2).Text
        !RecipeRev = Grd2.Cell(9, 2).Text
        !Exp = Grd2.Cell(10, 2).Text
        !Um = Grd2.Cell(11, 2).Text
        !Qty = Grd2.Cell(12, 2).Text
        !MinQty = Grd2.Cell(13, 2).Text
        !MaxQty = Grd2.Cell(14, 2).Text
        !UncertantlyFromCoA = Grd2.Cell(15, 2).Text
        !Procedure = Grd2.Cell(16, 2).Text
        !ProcedureRev = Grd2.Cell(17, 2).Text
       
        
        
       ' .Cell(1, 2).Text = IIf(IsNull(Trim(dbTabCode!Code)), "", Trim(dbTabCode!Code))
       ' .Cell(2, 2).Text = IIf(IsNull(Trim(dbTabCode!Line)), "", Trim(dbTabCode!Line))
       ' .Cell(3, 2).Text = IIf(IsNull(Trim(dbTabCode!STD)), "", Trim(dbTabCode!STD))
       ' .Cell(5, 2).Text = IIf(IsNull(Trim(dbTabCode!ProductName)), "", Trim(dbTabCode!ProductName))
       ' .Cell(6, 2).Text = IIf(IsNull(Trim(dbTabCode!Recipe)), "", Trim(dbTabCode!Recipe))
       ' .Cell(7, 2).Text = IIf(IsNull(Trim(dbTabCode!Mix1)), "", Trim(dbTabCode!Mix1))
       ' .Cell(8, 2).Text = IIf(IsNull(Trim(dbTabCode!Mix2)), "", Trim(dbTabCode!Mix2))
       ' .Cell(9, 2).Text = IIf(IsNull(Trim(dbTabCode!RecipeRev)), "", Trim(dbTabCode!RecipeRev))
       ' .Cell(10, 2).Text = IIf(IsNull(Trim(dbTabCode!Exp)), "", Trim(dbTabCode!Exp))
       ' .Cell(11, 2).Text = IIf(IsNull(Trim(dbTabCode!Um)), "", Trim(dbTabCode!Um))
       ' .Cell(12, 2).Text = IIf(IsNull(Trim(dbTabCode!Qty)), "", Trim(dbTabCode!Qty))
       ' .Cell(13, 2).Text = IIf(IsNull(Trim(dbTabCode!MinQty)), "", Trim(dbTabCode!MinQty))
       ' .Cell(14, 2).Text = IIf(IsNull(Trim(dbTabCode!MaxQty)), "", Trim(dbTabCode!MaxQty))
      '  .Cell(15, 2).Text = IIf(IsNull(Trim(dbTabCode!UncertantlyFromCoA)), "", Trim(dbTabCode!UncertantlyFromCoA))
      '  .Cell(16, 2).Text = IIf(IsNull(Trim(dbTabCode!Procedure)), "", Trim(dbTabCode!Procedure))
       ' .Cell(17, 2).Text = IIf(IsNull(Trim(dbTabCode!ProcedureRev)), "", Trim(dbTabCode!ProcedureRev))
        
        
        
        .Update
    End With


ERR_END:
    On Error GoTo 0
    If rc Then
        PopupMessage 2, "Code : " & MyNewCode & " saved!"
    Else
        PopupMessage 2, "Warning : a problem occurred, please check all entries before Save"
    End If
    
    SaveDatabaseCode = rc
    Exit Function
    
ERR_SAVE:
    rc = False
    MsgBox err.Description
    Resume Next:

End Function
Public Function CancellaCodeByID(ByVal dbTab As ADODB.Recordset, ByVal MyID As Long) As Boolean
Dim rc As Boolean
Dim Code As String
On Error GoTo ERR_CAN



    If MyID = 0 Then Exit Function
    
    
    rc = True
    With dbTab
        .filter = ""
        .filter = "ID='" & MyID & "'"
        If .EOF Then
            rc = False
        Else
        
        Code = Trim(!Code)
        Call DeleteCodeClassification(Code)
        
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
    
    CancellaCodeByID = rc
    Exit Function
    
ERR_CAN:
    rc = False
    MsgBox err.Description
    Resume ERR_END:
End Function

Public Sub AddComboCode(ByVal Combo1 As ComboBox)

    Combo1.Clear
    Combo1.AddItem "Code"
    Combo1.AddItem "Recipe"
    Combo1.ListIndex = 0
End Sub




Public Function GetStrHannaCode(ByRef iHannaCode() As HannaCode, ByRef strRecipes As String) As String
Dim i As Integer
Dim sRecipe As String
    strRecipes = ""
    For i = 1 To UBound(iHannaCode)
        If iHannaCode(i).bHide = False Then
            If i > 1 Then
                sRecipe = ""
                If InStr(GetStrHannaCode, iHannaCode(i).Code) Then
                Else
                    sRecipe = GetRecipeByHannaCode(iHannaCode(i).Code)
                    If InStr(strRecipes, sRecipe) Then
                    Else
                        strRecipes = strRecipes & IIf(strRecipes = "", "", " ; ") & sRecipe
                    End If
                    GetStrHannaCode = GetStrHannaCode & IIf(GetStrHannaCode = "", "", " ; ") & iHannaCode(i).Code
                End If
            
            Else
                GetStrHannaCode = iHannaCode(i).Code
                strRecipes = GetRecipeByHannaCode(iHannaCode(i).Code)
            End If
            
           
                
        End If
    Next
    GetStrHannaCode = Trim(Left(GetStrHannaCode, 255))
    
End Function

Private Function GetRecipeByHannaCode(ByVal Code As String) As String
    
    With dbTabCode
        .filter = ""
        .filter = "Code='" & Replace(Code, "'", "''") & "'"
        If .EOF Then
            GetRecipeByHannaCode = ""
        Else
            GetRecipeByHannaCode = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
        End If
    End With
    
End Function
Public Function GetHannaCodeID(ByVal Code As String) As String
    
    With dbTabCode
        .filter = ""
        .filter = "Code='" & Replace(Code, "'", "''") & "'"
        If .EOF Then
            GetHannaCodeID = 0
        Else
            GetHannaCodeID = !ID
        End If
    End With
    
End Function
