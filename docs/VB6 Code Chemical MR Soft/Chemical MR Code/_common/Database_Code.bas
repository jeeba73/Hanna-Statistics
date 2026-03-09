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
        
        .Cols = 8
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Code SFG"
        .Column(1).Width = 150
        .Cell(0, 2).Text = "Description"
        .Column(2).Width = 200
        
        .Cell(0, 3).Text = "Line"
        .Column(3).Width = 100
        
        .Cell(0, 4).Text = "MR 1"
        .Column(4).Width = 100
        
               
        .Cell(0, 5).Text = "MR 2"
        .Column(5).Width = 100
                     
        .Cell(0, 6).Text = "Range Max"
        .Column(6).Width = 0
                     
        .Cell(0, 7).Text = "ID"
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


Public Function FillGridCode(ByRef Grd As Grid, Optional ByVal Code As String, Optional bMainForm As Boolean, Optional ByVal sString As String)
Dim i As Integer
Dim t As Integer
Dim X As Integer
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
        With dbTabCode
        
            .filter = ""
            If sString = "Chemical MR" Then
                If Code <> "" Then .filter = "STDMR like '*" & Code & "*'"
            Else
                If Code <> "" Then .filter = "Code like '*" & Code & "*'"
            End If
            
            If .EOF Then Exit Function
            MaxCount = .RecordCount
            .MoveFirst
        End With
        i = 0
        For X = 1 To MaxCount
             If IsNull(dbTabCode!QCMethod) Or dbTabCode!QCMethod = "" Then GoTo cont:
            .AddItem "", False
            i = i + 1
            .Cell(i, 0).Text = i
            
            
            If IsNumeric(Trim(dbTabCode!Decimal)) Then
            Else
                dbTabCode!Decimal = "0"
                dbTabCode.Update
                bAddNewDatabaseRelease = True
            End If
            
            .Cell(i, 1).Text = "  " & IIf(IsNull(Trim(dbTabCode!Code)), "", Trim(dbTabCode!Code))

            
            .Cell(i, 1).Alignment = cellLeftCenter
            .Cell(i, 2).Text = "  " & IIf(IsNull(Trim(dbTabCode!ProductName)), "", Trim(dbTabCode!ProductName))
            .Cell(i, 2).Alignment = cellLeftCenter
            .Cell(i, 3).Text = IIf(IsNull(Trim(dbTabCode!Line)), "", Trim(dbTabCode!Line))
            .Cell(i, 4).Text = IIf(IsNull(Trim(dbTabCode!STDMR)), "", Trim(dbTabCode!STDMR))
            .Cell(i, 5).Text = IIf(IsNull(Trim(dbTabCode!STDMR2)), "", Trim(dbTabCode!STDMR2))
            .Cell(i, 6).Text = IIf(IsNull(Trim(dbTabCode!RangeMax)), "", Trim(dbTabCode!RangeMax))
            .Cell(i, 7).Text = dbTabCode!ID
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
            dbTabCode.MoveNext
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



Public Function SetGridEditCode(ByRef Grd As Grid) As Boolean
Dim i As Integer
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
        
        .Rows = 76

        .Cell(1, 1).Text = "  " & "Hanna SFG Code"
        .Cell(2, 1).Text = "  " & "SFG Description"
        .Cell(3, 1).Text = "  " & "Line"
        .Cell(4, 1).Text = "  " & "Recipe"
        
        .Cell(5, 1).Text = "  " & "QC Method"
        .Cell(6, 1).Text = "  " & "Meter Family 1"
        .Cell(7, 1).Text = "  " & "Meter Family 2"
        
        .RowHeight(6) = 0
        .RowHeight(7) = 0
        
        .Cell(8, 1).Text = "  " & "Parameter Method"
        .Cell(9, 1).Text = "  " & "Hanna Formula"
        .Cell(10, 1).Text = "  " & "Measurement Unit"
        
        
        .Range(11, 1, 11, 2).Merge
        .Cell(11, 1).Text = "User manual parameter data"
        .Cell(12, 1).Text = "  " & "Range Min"
        .Cell(13, 1).Text = "  " & "Range Max"
        .Cell(14, 1).Text = "  " & "Decimal"
        
        .Range(15, 1, 15, 2).Merge
        .Cell(15, 1).Text = "Tolerance"
        .Cell(16, 1).Text = "  " & "Fixed"
        .Cell(17, 1).Text = "  " & "And / Or"
        .Cell(18, 1).Text = "  " & "Percentage (%)"
        .Cell(19, 1).Text = "  " & "QC Restriction (%)"
       ' .Cell(20, 1).Text = "  " & "STD MR"
       
       .RowHeight(15) = 0
       .RowHeight(16) = 0
       .RowHeight(17) = 0
       .RowHeight(18) = 0
       .RowHeight(19) = 0
        
        .Range(21, 1, 21, 2).Merge
        .Cell(21, 1).Text = "STD1"
        
        .Cell(22, 1).Text = "  " & "Value"
        .Cell(23, 1).Text = "  " & "Min"
        .Cell(24, 1).Text = "  " & "Max"
        
       
        .RowHeight(23) = 0
        .RowHeight(24) = 0
       
        
        .Range(25, 1, 25, 2).Merge
        .Cell(25, 1).Text = "STD2"
        
        .Cell(26, 1).Text = "  " & "Value"
        .Cell(27, 1).Text = "  " & "Min"
        .Cell(28, 1).Text = "  " & "Max"
        
   
        .RowHeight(27) = 0
        .RowHeight(28) = 0
       
        
        .Range(29, 1, 29, 2).Merge
        .Cell(29, 1).Text = "STD3"
        
        .Cell(30, 1).Text = "  " & "Value"
        .Cell(31, 1).Text = "  " & "Min"
        .Cell(32, 1).Text = "  " & "Max"
       
       
      
        .RowHeight(31) = 0
        .RowHeight(32) = 0
              
              
        .Range(33, 1, 33, 2).Merge
        .Cell(33, 1).Text = "STD4"
        
        .Cell(34, 1).Text = "  " & "Value"
        .Cell(35, 1).Text = "  " & "Min"
        .Cell(36, 1).Text = "  " & "Max"
        
        
      
        .RowHeight(35) = 0
        .RowHeight(36) = 0
       
        
         .Range(37, 1, 37, 2).Merge
        .Cell(37, 1).Text = "STD5"
        .Cell(38, 1).Text = "  " & "Value"
        .Cell(39, 1).Text = "  " & "Min"
        .Cell(40, 1).Text = "  " & "Max"
        
     
        .RowHeight(39) = 0
        .RowHeight(40) = 0
       
                
        
         .Range(41, 1, 41, 2).Merge
        .Cell(41, 1).Text = "STD6"
        .Cell(42, 1).Text = "  " & "Value"
        .Cell(43, 1).Text = "  " & "Min"
        .Cell(44, 1).Text = "  " & "Max"
        
        
        .RowHeight(43) = 0
        .RowHeight(44) = 0
       
                        
                
        .Range(45, 1, 45, 2).Merge
        .Cell(45, 1).Text = "pH 1"
        .Cell(46, 1).Text = "  " & "Value"
        .Cell(47, 1).Text = "  " & "Min"
        .Cell(48, 1).Text = "  " & "Max"
                
                
        .Range(49, 1, 49, 2).Merge
        .Cell(49, 1).Text = "pH 2"
        .Cell(50, 1).Text = "  " & "Value"
        .Cell(51, 1).Text = "  " & "Min"
        .Cell(52, 1).Text = "  " & "Max"
                
        .Range(53, 1, 53, 2).Merge
        .Cell(53, 1).Text = "pH 3"
        .Cell(54, 1).Text = "  " & "Value"
        .Cell(55, 1).Text = "  " & "Min"
        .Cell(56, 1).Text = "  " & "Max"
                             
                             
        .Range(57, 1, 57, 2).Merge
        .Cell(57, 1).Text = "Weight (mg)"
        .Cell(58, 1).Text = "  " & "Value"
        .Cell(59, 1).Text = "  " & "Min"
        .Cell(60, 1).Text = "  " & "Max"
                
        .Cell(61, 1).Text = "  " & "Certified"
        
        
        For i = 45 To 61
        
         .RowHeight(i) = 0
         
        Next
        
        
        .Cell(62, 1).Text = "  " & "Revision Date"
        
        
        .Cell(63, 1).Text = "  " & "MR1"
        .Cell(64, 1).Text = "  " & "MR2"
        .Cell(65, 1).Text = "  " & "MS1 Value"
        .Cell(66, 1).Text = "  " & "MS1 Volume (ml)"
        
        .Cell(67, 1).Text = "  " & "MS2 dil"
        .Cell(68, 1).Text = "  " & "MS2 Volume (ml)"
        
        
        .Cell(69, 1).Text = "  " & "MS EXP (days)"
        .Cell(70, 1).Text = "  " & "STD Matrix"
        .Cell(71, 1).Text = "  " & "STD Volume (ml)"
        .Cell(72, 1).Text = "  " & "STD EXP (days)"
        .Cell(73, 1).Text = "  " & "STD Note"
        .Cell(74, 1).Text = "  " & "FW Hanna Parameter"

        .Cell(75, 1).Text = "  " & "STD Storage"
        
        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            
        Next
        
        
        For i = 1 To .Rows - 1
        
            .Cell(i, 1).BackColor = &HF0F0F0 'vbColorUnabled
            .Cell(i, 1).ForeColor = vbColorDarkFont 'vbColorDarkFont 'vbColorForeFixed  ' vbColorBlueProgram
            .Cell(i, 1).FontBold = False
            .Cell(i, 1).Locked = True
            .Cell(i, 2).ForeColor = vbColorDarkFont
            
            
            If i = 11 Or i = 15 Or i = 21 Or i = 25 Or i = 29 Or i = 33 Or i = 37 Or i = 41 Or i = 45 Or i = 49 Or i = 53 Or i = 57 Then
                .Cell(i, 1).Alignment = cellCenterCenter
                .Cell(i, 1).BackColor = &HF0F0F0 ' vbColorTextBlue ' &HF0F0F0
                .Cell(i, 1).ForeColor = vbColorDarkFont ' vbColorDarkFont 'vbWhite  ' &HF0F0F0
            End If
            
        Next


        
        
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
    
   
    With Grd2
       ' .DefaultFont.Size = 12 * m_ControlGridFontSize
       For i = 1 To .Rows - 1
            .Cell(i, 2).Text = (IIf(IsNull(Trim(dbTabCode.fields(i))), "", Trim(dbTabCode.fields(i))))
       Next
        

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

Err:
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
    RangeMin = Trim(Grd2.Cell(12, 2).Text)
    RangeMax = Trim(Grd2.Cell(13, 2).Text)
    
    If MyNewCode = "" Then
        PopupMessage 2, "Please Enter a valid Code!"
        Exit Function
    End If
    
    With dbTabCode
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
    
    SaveDatabaseCode = rc
    Exit Function
    
ERR_SAVE:
    rc = False
    MsgBox Err.Description
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
    MsgBox Err.Description
    Resume ERR_END:
End Function

Public Sub AddComboCode(ByVal Combo1 As ComboBox)

    Combo1.Clear
    Combo1.AddItem "Code"
    Combo1.AddItem "MRCode"
    Combo1.ListIndex = 0
End Sub


