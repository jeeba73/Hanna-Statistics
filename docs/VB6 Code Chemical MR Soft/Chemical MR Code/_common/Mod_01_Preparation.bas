Attribute VB_Name = "Mod_01_Preparation"
Option Explicit

Public Function SetPretarationType(ByRef Combo As ComboBox)
With Combo
    .Clear
    .AddItem "MS1"
    .AddItem "MS2"
    .AddItem "MRL"
    .ListIndex = 0
End With
End Function

Public Function SetType(ByRef Index As Integer) As String
Select Case Index
    Case 1
        SetType = "MS1"
    Case 2
        SetType = "MS2"
     Case 0
        SetType = "MRL"
End Select
End Function






Public Function GetPreparationFromDatabase(ByRef Grid As Grid, ByVal bClosed As Boolean, Optional ByVal userCode As String) As Boolean
Dim i As Integer
Dim t As Integer
Dim Count As Integer
Dim sString As String


sString = "bClosed='" & CStr(bClosed) & "'"

If UCase(userCode) = "SEARCH" Then userCode = ""
    
    If userCode = "" Then
    Else
        sString = sString & " and Code='" & Replace(Trim(userCode), "'", "''") & "'"
    End If

With Grid


    .Rows = 1
    .AutoRedraw = False
   
    
    With dbTabPreparation
        .filter = ""
        .filter = sString
        
        If .EOF Then
            Count = 0
        Else
            Count = .RecordCount
            .MoveFirst
        End If
        

        
        For i = 1 To Count
        

       ' .Cell(0, 1).Text = "HannaCode"
       ' .Cell(0, 2).Text = "Description"
       ' .Cell(0, 3).Text = "MRCode"
       ' .Cell(0, 4).Text = "Data Prep."
       ' .Cell(0, 5).Text = "Hour Prep."
        
       ' .Cell(0, 6).Text = "PrepWeek"
       ' .Cell(0, 7).Text = "Operator"
        
       ' .Cell(0, 8).Text = "QtyToProduce"
       ' .Cell(0, 9).Text = "Unit"
        
        '.Cell(0, 10).Text = "STD Matrix"
        '.Cell(0, 11).Text = "STD Exp"
        '.Cell(0, 12).Text = "STD Storage"
        
     
        '.Cell(0, 13).Text = "Note"
        
        
        '.Cell(0, 14).Text = "MS Type"
        '.Cell(0, 15).Text = "Closed"
        '.Cell(0, 16).Text = "ID"
        '.Cell(0, 17).Text = "Filename"
        
            
            If IsNull(Trim(!HannaCode)) Or Trim(!HannaCode) = "" Then
                .Delete
                .Update
            End If
        
            Grid.AddItem "", False
            

            Grid.Cell(Grid.Rows - 1, 1).Text = IIf(IsNull(Trim(!HannaCode)), "", Trim(!HannaCode))
            Grid.Cell(Grid.Rows - 1, 2).Text = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
           
           
           Grid.Cell(Grid.Rows - 1, 3).Text = IIf(IsNull(Trim(!MRCode)), "", Trim(!MRCode))
           Grid.Cell(Grid.Rows - 1, 4).Text = FormatDataLAT(IIf(IsNull(Trim(!DataPrep)), "", Trim(!DataPrep)))
           
           Grid.Cell(Grid.Rows - 1, 5).Text = IIf(IsNull(Trim(!HourPrep)), "", Trim(!HourPrep))
           Grid.Cell(Grid.Rows - 1, 6).Text = IIf(IsNull(Trim(!PrepWeek)), "", Trim(!PrepWeek))
           Grid.Cell(Grid.Rows - 1, 7).Text = IIf(IsNull(Trim(!Operator)), "", Trim(!Operator))
           Grid.Cell(Grid.Rows - 1, 8).Text = PadString((IIf(IsNull(Trim(!QtyToProduce)), "", Trim(!QtyToProduce))))
           
           Grid.Cell(Grid.Rows - 1, 9).Text = IIf(IsNull(Trim(!Unit)), "", Trim(!Unit))
           Grid.Cell(Grid.Rows - 1, 10).Text = (IIf(IsNull(Trim(!STDMatrix)), "", Trim(!STDMatrix)))
           Grid.Cell(Grid.Rows - 1, 11).Text = IIf(IsNull(Trim(!STDExp)), "", Trim(!STDExp))
           Grid.Cell(Grid.Rows - 1, 12).Text = IIf(IsNull(Trim(!STDStorage)), "", Trim(!STDStorage))
           
           Grid.Cell(Grid.Rows - 1, 13).Text = IIf(IsNull(Trim(!Note)), "", Trim(!Note))
           Grid.Cell(Grid.Rows - 1, 14).Text = SetType(IIf(IsNull(Trim(!MsType)), 0, Trim(!MsType)))
           Grid.Cell(Grid.Rows - 1, 15).Text = IIf(IsNull(Trim(!bClosed)), False, Trim(!bClosed))
           Grid.Cell(Grid.Rows - 1, 16).Text = !ID
           
           Grid.Cell(Grid.Rows - 1, 17).Text = IIf(IsNull(Trim(!FileName)), "", Trim(!FileName))


            Grid.Cell(Grid.Rows - 1, 1).FontBold = True
            Grid.Cell(Grid.Rows - 1, 1).ForeColor = &H473733
            Grid.Cell(Grid.Rows - 1, 7).FontBold = True
            Grid.Cell(Grid.Rows - 1, 7).ForeColor = &H473733
            
            Grid.Cell(Grid.Rows - 1, 8).FontBold = True
            Grid.Cell(Grid.Rows - 1, 8).ForeColor = &H473733
            Grid.Cell(Grid.Rows - 1, 9).FontBold = True
            Grid.Cell(Grid.Rows - 1, 9).ForeColor = &H473733
            
            
            
            Select Case !MsType
                
                Case "1"
                    Grid.Cell(Grid.Rows - 1, 14).BackColor = vbColorAzzurrino
                     Grid.Cell(Grid.Rows - 1, 3).BackColor = vbColorAzzurrino
                     
                Case "2"
                    Grid.Cell(Grid.Rows - 1, 14).BackColor = vbColorRosaTabella
                    Grid.Cell(Grid.Rows - 1, 3).BackColor = vbColorRosaTabella
                Case Else
                    Grid.Cell(Grid.Rows - 1, 14).BackColor = vbColorResults
                    Grid.Cell(Grid.Rows - 1, 3).BackColor = vbColorResults
            
            End Select
            
            ' Grid.Cell(Grid.Rows - 1, 3).FontBold = True
            ' Grid.Cell(Grid.Rows - 1, 14).FontBold = True
            
            Grid.Cell(Grid.Rows - 1, 3).Alignment = cellCenterCenter
            
            
            
            If !bManuale Then
                Grid.Cell(Grid.Rows - 1, 20).Text = "true"
                Grid.Cell(Grid.Rows - 1, 20).ForeColor = vbColorManualPreparation
                Grid.Cell(Grid.Rows - 1, 20).BackColor = vbColorManualPreparation
                
                'Grid.Cell(Grid.Rows - 1, 0).BackColor = vbColorManualPreparation
                Grid.Cell(Grid.Rows - 1, 1).BackColor = vbColorManualPreparation
                'Grid.Cell(Grid.Rows - 1, 2).BackColor = vbColorManualPreparation
                Grid.Cell(Grid.Rows - 1, 3).BackColor = vbColorManualPreparation
                
                Grid.Cell(Grid.Rows - 1, 1).ForeColor = vbWhite
                'Grid.Cell(Grid.Rows - 1, 2).ForeColor = vbWhite
                Grid.Cell(Grid.Rows - 1, 3).ForeColor = vbWhite
                
                
            
           Else
            
           End If
             
             
            
            .MoveNext
        Next
    
    End With
    For i = 1 To .Cols - 1
        .Column(i).AutoFit
        .Column(i).Width = .Column(i).Width * 1.1
    Next
    
    .Column(2).Width = 200
    .Column(15).Width = 0
     .Column(16).Width = 0
     .Column(17).Width = 0
     .Column(18).Width = 0
     .Column(19).Width = 0
     .Column(20).Width = 20
    .AllowUserReorderColumn = True
    .Column(4).Sort cellAscending
    

    .Refresh
    .AutoRedraw = True
    .ReadOnly = True


End With

End Function

Public Function GetHannaCodeFromDatabase(ByRef Grid As Grid, ByVal bClosed As Boolean, Optional ByVal userCode As String, Optional strLine As String) As Boolean
Dim i As Integer
Dim t As Integer
Dim Count As Integer
Dim sString As String


If UCase(userCode) = "SEARCH" Then userCode = ""
    
    If userCode = "" Then
    Else
        sString = "STDMR like '*" & Trim(userCode) & "*' or STDMR2 like '*" & Trim(userCode) & "*'"
    End If

With Grid

    .Rows = 1
    .AutoRedraw = False
   
    
    With dbTabCode
        .filter = ""
        .filter = sString
        
        If .EOF Then
            Count = 0
        Else
            Count = .RecordCount
            .MoveFirst
        End If
        

        
        For i = 1 To Count
        

            
        
            Grid.AddItem "", False
            
        '.Cell(0, 1).Text = "HannaCode"
        '.Cell(0, 2).Text = "Description"
        '.Cell(0, 3).Text = "MRCode"
        '.Cell(0, 4).Text = "Hanna Parameter"
        
        '.Cell(0, 5).Text = "FW Hanna Parameter"
        '.Cell(0, 6).Text = "STD Volume (ml)"
        '.Cell(0, 7).Text = "STD Matrix"
        '.Cell(0, 8).Text = "STD Exp"

        '.Cell(0, 9).Text = "STD Storage"
        '.Cell(0, 10).Text = "ID"
      
        

            Grid.Cell(Grid.Rows - 1, 1).Text = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
            Grid.Cell(Grid.Rows - 1, 2).Text = IIf(IsNull(Trim(!ProductName)), "", Trim(!ProductName))
           
           
           Grid.Cell(Grid.Rows - 1, 3).Text = IIf(IsNull(Trim(!STDMR)), "", Trim(!STDMR)) & IIf(IsNull(Trim(!STDMR2)) Or Trim(!STDMR2) = "", "", " - " & Trim(!STDMR2))
           Grid.Cell(Grid.Rows - 1, 4).Text = IIf(IsNull(Trim(!ParameterMethod)), "", Trim(!ParameterMethod))
           Grid.Cell(Grid.Rows - 1, 5).Text = (IIf(IsNull(Trim(!ParameterFormula)), "", Trim(!ParameterFormula)))
           
           Grid.Cell(Grid.Rows - 1, 6).Text = IIf(IsNull(Trim(!STDVolume)), "", Trim(!STDVolume))
           
           Grid.Cell(Grid.Rows - 1, 7).Text = (IIf(IsNull(Trim(!STDMatrix)), "", Trim(!STDMatrix)))
           Grid.Cell(Grid.Rows - 1, 8).Text = IIf(IsNull(Trim(!STDExp)), "", Trim(!STDExp))
           
           Grid.Cell(Grid.Rows - 1, 9).Text = IIf(IsNull(Trim(!STDNote)), "", Trim(!STDNote))
           Grid.Cell(Grid.Rows - 1, 10).Text = IIf(IsNull(Trim(!STDStorage)), "", Trim(!STDStorage))
           Grid.Cell(Grid.Rows - 1, 11).Text = !ID
          
         
            Grid.Cell(Grid.Rows - 1, 1).FontBold = True
            Grid.Cell(Grid.Rows - 1, 1).ForeColor = &H473733
            Grid.Cell(Grid.Rows - 1, 3).FontBold = True
            Grid.Cell(Grid.Rows - 1, 3).ForeColor = &H473733

            Grid.Cell(Grid.Rows - 1, 6).FontBold = True
            Grid.Cell(Grid.Rows - 1, 6).ForeColor = &H473733
         
            
            .MoveNext
        Next
    
    End With
    For i = 1 To .Cols - 1
        .Column(i).AutoFit
        .Column(i).Width = .Column(i).Width * 1.1
    Next
    
    .Column(2).Width = 200
    .Column(11).Width = 0
    
    .AllowUserReorderColumn = True
    .Column(1).Sort cellAscending
    

    .Refresh
    .AutoRedraw = True
    .ReadOnly = True


End With

End Function



Public Function DeleteSelectedPreparation(ByVal PreparationID As Long, ByVal FileName As String) As Boolean
    
  Dim rc As Boolean
  
  On Error GoTo ERR_DEL
  
    rc = True
    
    dbChemicalMR.Execute "DELETE * FROM TabAcquisition where PreparationID=" & PreparationID
    dbChemicalMR.Execute "DELETE * FROM TabPreparationNotes where PreparationID=" & PreparationID
    dbChemicalMR.Execute "DELETE * FROM TabMotherSolution where PreparationID=" & PreparationID


    
    If FileName <> "" Then
    
    
    If FileExists(USER_TEMP_PATH & FileName) Then
        Kill USER_TEMP_PATH & FileName
    ElseIf FileExists(USER_DATA_PATH & FileName) Then
        Kill USER_DATA_PATH & FileName
    Else
        PopupMessage 2, "No Preparation file found...", , True, FileName
    End If
    
    
    
    
    
    End If
    
    
    With dbTabPreparation
        .filter = ""
        .filter = "ID='" & PreparationID & "'"
        If .EOF Then
        Else
            .Delete
            .Update
        End If
    End With
    
ERR_END:
    
   On Error GoTo 0
   DeleteSelectedPreparation = rc
    Exit Function
ERR_DEL:
    rc = False
    MsgBox Err.Description
    Resume Next
   
    
End Function

