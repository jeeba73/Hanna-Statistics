Attribute VB_Name = "mod_Grid_fill"
Option Explicit


Public Function FillGridCode(ByRef Grd As Grid, Optional ByVal Code As String, Optional bMainForm As Boolean, Optional ByVal sString As String)
Dim i As Integer
Dim t As Integer
Dim MaxCount As Integer
Dim x As Integer
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
            If sString = "Recipe" Then
                If Code <> "" Then .filter = "Recipe like '*" & Code & "*'"
            Else
                If Code <> "" Then .filter = "Code like '*" & Code & "*'"
            End If
            
            If .EOF Then Exit Function
            MaxCount = .RecordCount
            .MoveFirst
        End With
        
        For x = 0 To MaxCount - 1
        
           ' If IsNull(dbTabCode!QCMethod) Or dbTabCode!QCMethod = "" Then GoTo cont:
                
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
            .Cell(i, 4).Text = IIf(IsNull(Trim(dbTabCode!Recipe)), "", Trim(dbTabCode!Recipe))
            .Cell(i, 5).Text = IIf(IsNull(Trim(dbTabCode!RangeMin)), "", Trim(dbTabCode!RangeMin))
            .Cell(i, 6).Text = IIf(IsNull(Trim(dbTabCode!RangeMax)), "", Trim(dbTabCode!RangeMax))
            .Cell(i, 7).Text = dbTabCode!ID
            For t = 1 To .Cols - 1
                If bMainForm Then
                    .Cell(i, t).ForeColor = vbColorDarkFont 'vbColorTextDarkBlue
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


Public Function SetCmbCode(ByRef cmb As ComboBox)
Dim i As Integer
Dim t As Integer
Dim MaxCount As Integer
Dim x As Integer
    On Error GoTo ERR_GRID
   

    With cmb
        .Clear
        
        With dbTabCode
        
            .filter = ""
            If .EOF Then Exit Function
            MaxCount = .RecordCount
            .MoveFirst
        End With
        
        For x = 0 To MaxCount - 1
        
           ' If IsNull(dbTabCode!QCMethod) Or dbTabCode!QCMethod = "" Then GoTo cont:
                
            .AddItem IIf(IsNull(Trim(dbTabCode!Code)), "", Trim(dbTabCode!Code))
            i = i + 1
           
cont:
            dbTabCode.MoveNext
        Next
       
ERR_END:
       
       .ListIndex = 0
    End With

    Exit Function
ERR_GRID:
    MessageInfoTime = 2000
 
    PopupMessage 2, Err.Description
    Resume Next
End Function

Public Function FillSTDToleranceGrid(ByVal sCode As String, ByRef Grd3 As Grid, Optional ByRef lbSpecification As Label, Optional ByRef SelectedCodeID As Long, Optional ByVal bColorCell As Boolean) As Boolean
Dim rc As Boolean
Dim sString As String
Dim strString As String
Dim i As Integer
On Error GoTo ERR_SEL
sString = sCode


Dim strDecimal As String
Dim strUnit‡DiMisura As String


On Error GoTo ERR_SEL:


If sString = "" Then
   
Else


End If

rc = True

    With dbTabCode
        .filter = ""
        If SelectedCodeID = 0 Then
            .filter = "Code='" & Replace(sCode, "'", "''") & "'"
            
        Else
            .filter = "ID='" & SelectedCodeID & "'"
        End If
        
        
        If .EOF Then
           ' PopupMessage 2, "Not a Valid Hanna Code"
           rc = False
           Grd3.Rows = 2
           GoTo ERR_END:
           
           
        Else
           
        End If
        SelectedCodeID = !ID
        Grd3.Rows = 3
        '------------------------
        ' trovato codice
        '------------------------
            If Not (lbSpecification Is Nothing) Then
            lbSpecification.Caption = "Hanna Code : " & Trim(!Code) & " (" & Trim(!ProductName) & ")"
            lbSpecification.FontSize = IIf(Len(lbSpecification) > 60, 14, 22)
           ' Debug.Print Len(lbSpecification)
            End If
            strUnit‡DiMisura = Trim(!MeasurementUnit)
            strDecimal = Trim(!Decimal)
            strDecimal = FormatDecimal(strDecimal)
            Grd3.RowHeight(2) = 50
            Grd3.SelectionMode = cellSelectionNone
            
            
            
              
        '.Cell(0, 1).Text = "Tolerance"
        '.Cell(1, 1).Text = "Fixed"
        '.Cell(1, 2).Text = "And / Or"
        '.Cell(1, 3).Text = "%"
        '.Cell(1, 4).Text = "Qc Restriction"
        '.Cell(0, 5).Text = "STD MR"
        '.Cell(0, 6).Text = "STD1"
        '.Cell(0, 9).Text = "STD2"
        '.Cell(0, 12).Text = "STD3"
        '.Cell(0, 15).Text = "STD4"
        '.Cell(0, 18).Text = "STD5"
        '.Cell(0, 21).Text = "STD6"
        '.Cell(0, 24).Text = "pH 1"
        '.Cell(0, 27).Text = "pH 2"
        '.Cell(0, 30).Text = "pH 3"
        '.Cell(0, 33).Text = "Weight"
        'Dim i As Integer
        'For i = 6 To 35 Step 3
        '    .Cell(1, i).Text = "Value"
        '    .Cell(1, i + 1).Text = "Min"
        '    .Cell(1, i + 2).Text = "Max"
        'Next
      

            
                        Grd3.Cell(0, 1).Text = "Tolerance ( " & strUnit‡DiMisura & " )"
                        Grd3.Cell(0, 6).Text = "STD1" & " ( " & strUnit‡DiMisura & " )"
                        Grd3.Cell(0, 9).Text = "STD2" & " ( " & strUnit‡DiMisura & " )"
                        Grd3.Cell(0, 12).Text = "STD3" & " ( " & strUnit‡DiMisura & " )"
                        Grd3.Cell(0, 15).Text = "STD4" & " ( " & strUnit‡DiMisura & " )"
                        Grd3.Cell(0, 18).Text = "STD5" & " ( " & strUnit‡DiMisura & " )"
                        Grd3.Cell(0, 21).Text = "STD6" & " ( " & strUnit‡DiMisura & " )"
                        Grd3.Cell(0, 33).Text = "Weight ( mg )"
              
            
            
          '  Debug.Print .Fields(.Fields.Count - 18).Name
            
            
            For i = 32 To .Fields.Count - 19
            
                Grd3.Column(i - 31).Width = 150
              
                strString = IIf(IsNull(Trim(.Fields(i))), "", Trim(.Fields(i)))
                Select Case i
                    Case 32
                        
                        If (strString = "0") Or (strString = "/") Then
                            strString = "No"
                        Else
                            strString = Format$(strString, strDecimal) & " "
                        End If
                    Case 34, 35
                        If (strString = "0") Or (strString = "/") Then
                            strString = "0"
                        Else
                            strString = strString & " %"
                        End If
                    Case 36
                    Case 37 To 54
                        
                        If InStr(strString, "/") Or Trim(strString) = "" Then Grd3.Column(i - 31).Width = 0
                        
                        strString = Format$(strString, strDecimal) & " "
                    Case 35 To 63
                        If InStr(strString, "/") Or Trim(strString) = "" Then Grd3.Column(i - 31).Width = 0
                    Case 64 To .Fields.Count - 19
                        If InStr(strString, "/") Or Trim(strString) = "" Then Grd3.Column(i - 31).Width = 0
                    
                End Select
               
                strString = Replace(strString, "/", "")
                
                
                Grd3.Cell(2, i - 31).Text = strString 'Format$(strString, strDecimal)
                
                
                If bColorCell Then Grd3.Cell(2, i - 31).BackColor = vbWhite
                If bColorCell Then
                    Select Case i
                        Case 32
                            Grd3.Cell(0, i - 31).BackColor = vbColorTextLightBlue
                        Case 37, 40, 43, 46, 49, 52 ', 38, 41, 44
                            Grd3.Cell(0, i - 31).BackColor = vbColorTextLightBlue
                            Grd3.Cell(2, i - 31).BackColor = vbColorTextLightBlue
                    End Select
                End If
                Grd3.Cell(2, i - 31).FontBold = True
            Next
        

    End With
ERR_END:
    On Error GoTo 0
    FillSTDToleranceGrid = rc
    Exit Function
ERR_SEL:
    rc = False
    MsgBox Err.Description
    Resume Next:
End Function
Public Function STDToleranceGridColumn(ByRef Grd3 As Grid, Optional ByVal bValue As Boolean) As Boolean
Dim rc As Boolean
Dim MyWidth As Integer
Dim i As Integer
On Error GoTo ERR_SEL

rc = True

    With Grd3
        For i = 1 To .Cols - 1
            Select Case i
                Case 6 To 23
                    If .Column(i).Width = 0 Then
                        MyWidth = 0
                    Else
                        MyWidth = 100
                    End If
                Case Else
                    MyWidth = IIf(bValue, 0, 100)
            End Select
            .Column(i).Width = MyWidth
        Next
    End With
ERR_END:
    On Error GoTo 0
    STDToleranceGridColumn = rc
    Exit Function
ERR_SEL:
    rc = False
    MsgBox Err.Description
    Resume Next:
End Function


Public Function GetMeanTable(ByRef Grd3 As Grid, ByVal SettingName As String)
    
Dim i As Integer
Dim t As Integer
Dim nRows As Long
Dim nCols As Long
        '  tabella Results
CloseSettingDataFile
    With Grd3
        .AutoRedraw = False
        nRows = GetSettingData(SettingName, "Evaluation QC", "Results Grid Rows", .Rows)
        nCols = GetSettingData(SettingName, "Evaluation QC", "Results Grid Cols", .Cols)
        .Rows = nRows
        .Cols = 8 ' nCols
        For i = 1 To .Rows - 1
            For t = 1 To .Cols - 1
                .Cell(i, t).Text = GetSettingData(SettingName, "Evaluation QC", "Results Grid Standard (" & i & ")  Column " & t, .Cell(i, t).Text)
                Debug.Print .Cell(i, t).Text
                .Cell(i, t).ForeColor = GetSettingData(SettingName, "Evaluation QC", "Results Grid Standard (" & i & ") Forecolor " & t, vbBlack)
            Next
            
          
            
        Next
        
        ' faccio un check se trovo righe bianche le cancello....
start:
        For i = 1 To .Rows - 1
            If .Cell(i, 1).Text = "" And .Cell(i, 2).Text = "" And .Cell(i, 3).Text = "" Then
                ' c'Ë qualcosa che non va.....
            .Cell(i, 1).SetFocus
          
            .ReadOnly = False
            .Selection.DeleteByRow
            .ReadOnly = True
            GoTo start:
            End If
        Next
        .AutoRedraw = True
        .Refresh
    End With
    
    
    CloseSettingDataFile
End Function
