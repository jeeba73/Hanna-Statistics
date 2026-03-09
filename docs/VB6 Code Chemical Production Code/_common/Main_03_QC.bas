Attribute VB_Name = "Main_03_QC"
Option Explicit

Public Function PassToQCInfile(ByRef RecipeQC As QCType) As Boolean
Dim i As Integer
Dim PassDate As String

Dim rc As Boolean

    On Error GoTo ERR_SET:
    
    rc = True
    
    
     With RecipeQC
        
        
        USER_PATH = USER_PREPARATION_PATH
        
        If FileExists(USER_PATH & .SettingName) = False Then
            rc = False
            GoTo ERR_END
        End If
    
        
        CloseSettingDataFile
        PassDate = GetSettingData(.SettingName, "Preparation", "PassToQC Date", "")
        CloseSettingDataFile
    
        If PassDate = "" Then
            SaveSettingData .SettingName, "Preparation", "PassToQC", True
            SaveSettingData .SettingName, "Preparation", "PassToQC Operator", .Operator
            SaveSettingData .SettingName, "Preparation", "PassToQC Date", .Date
        End If
    End With
    
    Call SetQcStatus(RecipeQC)
    Call SetTabPreparationToQC(RecipeQC)

ERR_END:
    On Error GoTo 0
    CloseSettingDataFile
    PassToQCInfile = rc
    Exit Function
ERR_SET:
    rc = False
    Resume Next
End Function

Public Function SetTabPreparationToQC(ByRef RecipeQC As QCType) As Boolean
Dim i As Integer
Dim rc As Boolean
Dim bMix As Boolean
Dim sSep As String

    On Error GoTo ERR_SET:
    
    rc = True
    
    bMix = IfRecipeIsMixString(Trim(RecipeQC.RecipeCode))
    
    With dbTabPreparation
        .filter = ""
        .filter = "ID='" & RecipeQC.ID & "'"
        If .EOF Then
            rc = False
        Else
            !QCStatus = RecipeQC.Status
            !Operator = RecipeQC.Operator
            !QCDate = RecipeQC.Date
            !QCNote = RecipeQC.Note
            !PassToQC = True
            
            If RecipeQC.Correction <> "" Then
                If IsNull(!Correction) = False Then sSep = ";"
                !Correction = !Correction & sSep & RecipeQC.Correction
                !CorrectionDate = !CorrectionDate & sSep & RecipeQC.CorrectionDate
            End If
            
            !QCOperator = RecipeQC.QCOperator
            
            .Update
        End If

    
    End With
   
ERR_END:
    On Error GoTo 0

    SetTabPreparationToQC = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox err.Description
    Resume Next
End Function



Public Function SetQcStatus(ByRef RecipeQC As QCType) As Boolean

Dim i As Integer
Dim rc As Boolean

    On Error GoTo ERR_SET:
    
    rc = True
    USER_PATH = USER_PREPARATION_PATH
    
    
    With RecipeQC
        If FileExists(USER_PATH & .SettingName) = False Then
            rc = False
            GoTo ERR_END
        End If
        
        
        CloseSettingDataFile
        
        .Index = GetSettingData(.SettingName, "QC", "Count", 0)
        
        i = .Index + 1
        
        CloseSettingDataFile
        
        SaveSettingData .SettingName, "QC", "Count", i
        
        SaveSettingData .SettingName, "QC", "Status" & i, .Status
        SaveSettingData .SettingName, "QC", "Operator" & i, .Operator
        SaveSettingData .SettingName, "QC", "Date" & i, .Date
        SaveSettingData .SettingName, "QC", "Note" & i, .Note
        SaveSettingData .SettingName, "QC", "Registration" & i, .Registration
        
        SaveSettingData .SettingName, "QC", "QCOperator" & i, .QCOperator
        SaveSettingData .SettingName, "QC", "Correction" & i, .Correction
        SaveSettingData .SettingName, "QC", "CorrectionDate" & i, .CorrectionDate
        
        
        .Index = i
        DoEvents
        
    End With
ERR_END:
    On Error GoTo 0
    CloseSettingDataFile
    DoEvents
    
    SetQcStatus = rc
    Exit Function
ERR_SET:
    rc = False
    Resume Next

End Function






Public Function GetQcInGrid(ByVal Grid As Grid, Optional ByVal strLine As String, Optional ByVal strRecipe As String, Optional ByVal bClosed As Boolean, Optional ByRef Frame2 As Frame, Optional ByVal bTuttiRecord As Boolean)

Dim i As Integer
Dim MaxCount As Integer
Dim DiffQty As Double
Dim sString As String
Dim QtyProduced As Double
Dim QtyToProduce As Double
Dim bAllMixes As Boolean
Dim strFileName As String
Dim strStatus As String
Dim MixesCount As Integer
Dim bPassed As Boolean
Dim bMix As Boolean
On Error GoTo ERR_GET:

    If strRecipe = "" Then
    Else
        
        If LCase(strRecipe) = "search" Then
        
            sString = ""
        Else
            sString = " and Recipe like '*" & Replace(Trim(strRecipe), "'", "''") & "*'"
        End If
        
    End If
    
    If strLine = "" Then
    Else
        
        If LCase(strLine) = "all lines" Or strLine = "" Then
        
           
        Else
            sString = sString & " and Line like '*" & Replace(Trim(strLine), "'", "''") & "*'"
        End If
        
    End If
    
    Dim dDate As Date
    dDate = DateAdd("m", -6, Date)
    If bClosed Then sString = sString & " and DataRecipe>=#" & dDate & "# "
    
    
    With Grid
        .Rows = 1
        .AutoRedraw = False
        .Column(1).Width = 0
       
        With dbTabPreparation
            .filter = ""
            .filter = "PassToQC=true  and QcClosed=" & (bClosed) & sString
            If .EOF Then
                GoTo exitMe
            End If
            .MoveLast
            
            MaxCount = .RecordCount
            
            If bTuttiRecord = False And bClosed Then
                MaxCount = IIf(MaxCount > 80, 80, MaxCount)
            End If
        
            For i = 1 To MaxCount
            
                bPassed = False
                bMix = False
                strStatus = (IIf(IsNull(Trim(!QCStatus)), "", Trim(!QCStatus)))
                strRecipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
                
                If IfRecipeNoPreparation(strRecipe) Then GoTo cont:
                
                Grid.AddItem "", False
                
              
                bAllMixes = IfAllMixes(strRecipe, MixesCount)
                strFileName = IIf(IsNull(Trim(!FileName)), "", Trim(!FileName))

                Grid.Cell(Grid.Rows - 1, 1).Text = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
                Grid.Cell(Grid.Rows - 1, 2).Text = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
                Grid.Cell(Grid.Rows - 1, 3).Text = IIf(IsNull(Trim(!HannaCode)), "", Trim(!HannaCode))
                Grid.Cell(Grid.Rows - 1, 3).FontBold = True
                Grid.Cell(Grid.Rows - 1, 15).Text = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
        
        
                Grid.Cell(Grid.Rows - 1, 4).Text = IIf(IsNull(Trim(!numPrepWeek)), "", Trim(!numPrepWeek))
                Grid.Cell(Grid.Rows - 1, 5).Text = IIf(IsNull(Trim(!PrepWeek)), "", Trim(!PrepWeek))
                Grid.Cell(Grid.Rows - 1, 6).Text = IIf(IsNull(Trim(!PrepDate)), "", Trim(!PrepDate))

                Grid.Cell(Grid.Rows - 1, 7).Text = strStatus
                Grid.Cell(Grid.Rows - 1, 8).Text = (IIf(IsNull(Trim(!QCDate)), "", Trim(!QCDate)))
                Grid.Cell(Grid.Rows - 1, 9).Text = IIf(IsNull(Trim(!FileName)), "", Trim(!FileName))
                Grid.Cell(Grid.Rows - 1, 12).Text = !ID
                Grid.Cell(Grid.Rows - 1, 13).Text = IIf(IsNull(Trim(!Operator)), "", Trim(!Operator))
                Grid.Cell(Grid.Rows - 1, 14).Text = IIf(IsNull(Trim(!QCNote)), "", Trim(!QCNote))
                Grid.Cell(Grid.Rows - 1, 16).Text = IIf(IsNull(Trim(!Lot)), "", Trim(!Lot))
                Grid.Cell(Grid.Rows - 1, 16).Alignment = cellRightCenter
                Grid.Cell(Grid.Rows - 1, 16).FontBold = True
                Grid.Cell(Grid.Rows - 1, 16).ForeColor = vbColorTextBlue

                Grid.Cell(Grid.Rows - 1, 7).Alignment = cellCenterCenter
                Grid.Cell(Grid.Rows - 1, 8).Alignment = cellCenterCenter
                Grid.Cell(Grid.Rows - 1, 13).Alignment = cellCenterCenter
                
                Grid.Cell(Grid.Rows - 1, 7).BackColor = vbColorResults
                Grid.Cell(Grid.Rows - 1, 8).BackColor = vbColorResults
                Grid.Cell(Grid.Rows - 1, 13).BackColor = vbColorResults
                
                Select Case Trim(!QCStatus)
                    Case "Passed"
                        Grid.Cell(Grid.Rows - 1, 7).BackColor = vbColorGreen
                        Grid.Cell(Grid.Rows - 1, 7).ForeColor = vbWhite
                        
                        bPassed = True
                       ' bMix = IfRecipeIsMixString(Trim(!Recipe))
                        'If bMix Then Grid.RowHeight(Grid.Rows - 1) = 0
                        
                    Case "Waiting"
                        Grid.Cell(Grid.Rows - 1, 7).BackColor = vbColorAzzurrino
                    Case "Failed"
                        Grid.Cell(Grid.Rows - 1, 7).BackColor = vbColorRed
                        Grid.Cell(Grid.Rows - 1, 7).ForeColor = vbWhite
                End Select
                
cont:
                
                CloseSettingDataFile
                

                
                .MovePrevious
            Next
        End With
        

        .Cell(0, 4).Text = "# Prep. Week"
        .Cell(0, 5).Text = "Prep. Week"
        .Cell(0, 6).Text = "Preparation Date"

        
        .Column(0).Width = 0
        .Column(14).AutoFit
        

        .Column(1).AutoFit
        .Column(2).AutoFit
        .Column(4).AutoFit
        .Column(5).AutoFit
        .Column(6).AutoFit
        .Column(15).AutoFit
        .Column(16).AutoFit
        
exitMe:


        If bClosed Then
            .Column(8).Sort cellDescending
        Else
            .Column(6).Sort cellDescending
        End If
        .SelectionMode = cellSelectionByRow
        .Refresh
        .AutoRedraw = True
    End With

    Frame2.Visible = IIf(Grid.Rows > 1, False, True)
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_GET:
    Resume Next

End Function




Public Function GetQCPerRecipeInGrid7(ByVal Grid7 As Grid, Optional ByVal strFileName As String)

Dim i As Integer
Dim MaxCount As Integer
Dim RecipeQCInGrid() As QCType



On Error GoTo ERR_GET:


    Call GetQcStatus(RecipeQCInGrid, strFileName, MaxCount)
    
   
    With Grid7
        .Rows = 1
        .AutoRedraw = False

            For i = 1 To MaxCount
            

        
                .AddItem "", False
                
                
                '.Cell(0, 1).Text = "QC status"
                '.Cell(0, 2).Text = "QC Date"
                '.Cell(0, 3).Text = "Operator"
                '.Cell(0, 4).Text = "Note"
               
                .Cell(.Rows - 1, 1).Text = RecipeQCInGrid(i).Status
                .Cell(.Rows - 1, 2).Text = RecipeQCInGrid(i).Date
                .Cell(.Rows - 1, 3).Text = RecipeQCInGrid(i).Operator
                .Cell(.Rows - 1, 4).Text = RecipeQCInGrid(i).Note
                .Cell(.Rows - 1, 5).Text = RecipeQCInGrid(i).Registration
                
                .Cell(.Rows - 1, 6).Text = RecipeQCInGrid(i).QCOperator
                .Cell(.Rows - 1, 7).Text = RecipeQCInGrid(i).Correction
                .Cell(.Rows - 1, 8).Text = RecipeQCInGrid(i).CorrectionDate
                

  
                Select Case Trim(RecipeQCInGrid(i).Status)
                    Case "Passed"
                        .Cell(.Rows - 1, 1).BackColor = vbColorGreen
                        .Cell(.Rows - 1, 1).ForeColor = vbWhite
                    Case "Waiting"
                        .Cell(.Rows - 1, 1).BackColor = vbColorAzzurrino
                    Case "Failed"
                        .Cell(.Rows - 1, 1).BackColor = vbColorRed
                        .Cell(.Rows - 1, 1).ForeColor = vbWhite
                End Select
                
            Next
      
        
exitMe:
        .Column(4).Alignment = cellLeftCenter
        .Column(4).AutoFit
        .SelectionMode = cellSelectionByRow
        .Refresh
        .AutoRedraw = True
    End With

ERR_END:
    On Error GoTo 0
    Exit Function
ERR_GET:
    Resume Next

End Function





Public Function GetQcStatus(ByRef RecipeQC() As QCType, ByVal SettingName As String, ByRef MaxCount As Integer) As Boolean

Dim i As Integer

Dim rc As Boolean

    On Error GoTo ERR_SET:
    
    rc = True
    USER_PATH = USER_PREPARATION_PATH
    If FileExists(USER_PATH & SettingName) Then
    ElseIf FileExists(USER_PREPARATION_PATH & "Data\" & SettingName) Then
        USER_PATH = USER_PREPARATION_PATH & "Data\"
    Else
        rc = False
        GoTo ERR_END
    End If
        
    CloseSettingDataFile
    
    MaxCount = GetSettingData(SettingName, "QC", "Count", 0)
    
    
    ReDim RecipeQC(MaxCount)
    
    For i = 1 To MaxCount
    
        With RecipeQC(i)
    
            .Status = GetSettingData(SettingName, "QC", "Status" & i, .Status)
            .Operator = GetSettingData(SettingName, "QC", "Operator" & i, .Operator)
            .Date = FormatDateTime(GetSettingData(SettingName, "QC", "Date" & i, .Date), vbShortDate)
            .Note = GetSettingData(SettingName, "QC", "Note" & i, .Note)
            .Registration = GetSettingData(SettingName, "QC", "Registration" & i, .Registration)
            
            .QCOperator = GetSettingData(SettingName, "QC", "QCOperator" & i, .QCOperator)
            .Correction = GetSettingData(SettingName, "QC", "Correction" & i, .Correction)
            .CorrectionDate = GetSettingData(SettingName, "QC", "CorrectionDate" & i, .CorrectionDate)
            
            .Index = i
            
        End With
    Next
ERR_END:
    On Error GoTo 0
    CloseSettingDataFile
    USER_PATH = USER_PREPARATION_PATH
    GetQcStatus = rc
    Exit Function
ERR_SET:
    rc = False
    Resume Next

End Function



Public Function RecipeCloseQC(ByVal PreparationFileName As String, ByVal PreparationID As String) As Boolean

Dim i As Integer

Dim rc As Boolean
Dim RecipeQC As QCType

    On Error GoTo ERR_SET:
    
    rc = True
    
    SettingName = PreparationFileName
    If FileExists(USER_PREPARATION_PATH & SettingName) Then
        
    ElseIf FileExists(USER_PREPARATION_PATH & "data\" & SettingName) Then
        GoTo cont
    Else
        rc = False
        GoTo ERR_END
    End If
    
    CloseSettingDataFile
    
    SaveSettingData SettingName, "iPreparation", "bOpen", False, USER_PREPARATION_PATH
    SaveSettingData SettingName, "QC Closed", "Manually", True, USER_PREPARATION_PATH
    SaveSettingData SettingName, "QC Closed", "Date", Now(), USER_PREPARATION_PATH
    SaveSettingData SettingName, "QC Closed", "Operator", MyOperatore.Name, USER_PREPARATION_PATH
    
    CloseSettingDataFile
    
    FileCopy USER_PREPARATION_PATH & SettingName, USER_PREPARATION_PATH & "data\" & SettingName
    Kill USER_PREPARATION_PATH & SettingName
    

cont:
    With RecipeQC
        .ID = PreparationID
        
    End With
    
    
    Call SetTabPreparationQcClosed(RecipeQC, True)
    
ERR_END:
    On Error GoTo 0
    CloseSettingDataFile
    RecipeCloseQC = rc
    Exit Function
ERR_SET:
    rc = False
    Resume Next

End Function

Public Function MoveRecipeInProductionDatabase(ByVal PreparationFileName As String, ByVal PreparationID As String, ByVal bIsMix As Boolean) As Boolean

Dim i As Integer
Dim rfpSettingName As String
Dim rc As Boolean
Dim RecipeQC As QCType
Dim RFPFILE_PATH As String
Dim userRfp As RfpDetails
    On Error GoTo ERR_SET:
    
    rc = True
    
    SettingName = PreparationFileName
    If FileExists(USER_PREPARATION_PATH & SettingName) Then
        
    ElseIf FileExists(USER_PREPARATION_PATH & "data\" & SettingName) Then
        GoTo cont
    Else
        rc = False
        GoTo ERR_END
    End If
    
    
    CloseSettingDataFile
    
    SaveSettingData SettingName, "iPreparation", "bOpen", False, USER_PREPARATION_PATH
    
    CloseSettingDataFile
    
    FileCopy USER_PREPARATION_PATH & SettingName, USER_PREPARATION_PATH & "data\" & SettingName
    Kill USER_PREPARATION_PATH & SettingName
    
cont:

    With RecipeQC
        .ID = PreparationID
        
    End With
    

               
    
    Call SetTabPreparationQcClosed(RecipeQC, False)
    
    If Not (bIsMix) Then
    
        If SetDatabaseTabProduction(RecipeQC) Then
        
        Else
            rc = False
        End If
    Else
        '------------------------------------------------------------------------------------------------------------
        ' è un mix probabilmente di una recipe senza preparation per cui salvo i dati di preparation delle mix!!!!
        '------------------------------------------------------------------------------------------------------------
        
        With dbTabPreparation
            .filter = ""
            .filter = "ID='" & PreparationID & "'"
            
            If .EOF Then
            
            Else
                userRfp.PrepDate = IIf(IsNull(Trim(!PrepDate)), "", Trim(!PrepDate))
                userRfp.PrepWeek = IIf(IsNull(Trim(!PrepWeek)), "", Trim(!PrepWeek))
                userRfp.numPrepWeek = IIf(IsNull(Trim(!numPrepWeek)), "", Trim(!numPrepWeek))
                userRfp.ExpDate = IIf(IsNull(Trim(!ExpDate)), "", Trim(!ExpDate))
                rfpSettingName = IIf(IsNull(Trim(!RfpFileName)), "", Trim(!RfpFileName))
            End If
        
        End With
        
        
        If rfpSettingName <> "" Then
        
            If FileExists(USER_TEMP_PATH & rfpSettingName) Then
                RFPFILE_PATH = USER_TEMP_PATH
            ElseIf FileExists(USER_DATA_PATH & rfpSettingName) Then
                RFPFILE_PATH = USER_DATA_PATH
            ElseIf FileExists(USER_PRODUCTION_PATH & rfpSettingName) Then
                RFPFILE_PATH = USER_PRODUCTION_PATH
            End If
            
            CloseSettingDataFile
            
            With userRfp
                SaveSettingData rfpSettingName, "iRecipeForProduction", "PreparationDate", .PrepDate, RFPFILE_PATH
                SaveSettingData rfpSettingName, "iRecipeForProduction", "PreparationLot", .Lot, RFPFILE_PATH
                SaveSettingData rfpSettingName, "iRecipeForProduction", "PrepWeek", .PrepWeek, RFPFILE_PATH
                SaveSettingData rfpSettingName, "iRecipeForProduction", "numPrepWeek", .numPrepWeek, RFPFILE_PATH
                SaveSettingData rfpSettingName, "iRecipeForProduction", "ExpDate", .ExpDate, RFPFILE_PATH
            End With
            
            CloseSettingDataFile
        End If
    
    
    
    
    
    End If
    
ERR_END:
    On Error GoTo 0
    CloseSettingDataFile
    MoveRecipeInProductionDatabase = rc
    Exit Function
ERR_SET:
    rc = False
    Resume Next

End Function






Public Function SetTabPreparationQcClosed(ByRef RecipeQC As QCType, ByVal bManually As Boolean) As Boolean
Dim i As Integer
Dim rc As Boolean

    On Error GoTo ERR_SET:
    
    rc = True
    

    
    With dbTabPreparation
        .filter = ""
        .filter = "ID='" & RecipeQC.ID & "'"
        If .EOF Then
            rc = False
        Else
            !PassToQC = True
            !bClosed = True
            !QcClosed = True
            !Operator = MyOperatore.Name
            !QCDate = FormatDataLAT(Now())
            If bManually Then
                !QCStatus = "Closed"
                !QCNote = "Manually Closed"
            End If
            RecipeQC.SettingName = !RfpFileName
            .Update
        End If
    End With
   
ERR_END:
    On Error GoTo 0

    SetTabPreparationQcClosed = rc
    Exit Function
ERR_SET:
    rc = False
    Resume Next
End Function



'--------------------------------------------------------------------
' check scanned QRCode
'---------------------------------------------------------------------


Public Function QRCodeQCToTabPreparation(ByRef QRCode As QRCodeType, Optional ByVal bClosed As Boolean) As Boolean

Dim rc As Boolean
Dim i As Integer
Dim sString As String

On Error GoTo ERR_GET:

        rc = True
        
        With QRCode
            If bCODLine Then
                ' dalla release 1.3.22
                ' le etichette verso QC vengono stampate con IL SFG Lot
                ' ora devo riconoscerlo...
                sString = "SFGLot like '*" & .Lot & "*' and ExpDate='" & .Exp & "'"
            Else

                sString = "Lot='" & .Lot & "' and ExpDate='" & .Exp & "'"
            
            End If
    
        End With
        
        
        
        With dbTabPreparation
            .filter = ""
            '.filter = "PassToQC=true  and QcClosed=" & (bClosed) & " And " & sString
             .filter = "PassToQC=true  and " & sString
            If .EOF Then
                rc = False
            Else
                QRCode.FileName = IIf(IsNull(Trim(!FileName)), "", Trim(!FileName))
            End If
         
         End With
         
        
ERR_END:
    On Error GoTo 0

    QRCodeQCToTabPreparation = rc
    Exit Function
ERR_GET:
    rc = False
    MsgBox err.Description
    Resume Next
End Function

Public Function SearchQCInTab(ByRef UserQrCode As QRCodeType, ByRef Grid As Grid) As Boolean
    Dim rc As Boolean
    Dim i As Integer
    On Error GoTo ERR_QR:
    
    rc = True
    
    With Grid
        If .Rows > 1 Then
            
            For i = 1 To .Rows - 1
                If .Cell(i, 2).Text = UserQrCode.Recipe And .Cell(i, 16).Text = UserQrCode.Lot Then
                    .Cell(i, 2).EnsureVisible
                    .Cell(i, 2).SetFocus
                    Exit For
                End If
            Next
            rc = False
        End If
    End With
ERR_END:
    On Error GoTo 0
    SearchQCInTab = rc
    Exit Function
ERR_QR:
    rc = False
    Resume ERR_END:


End Function

