Attribute VB_Name = "gPreparation_02_Functions"
Option Explicit

Public Function SetInitSTDTheoreticalWeight(ByRef uPreparation As RecipeForProduction) As Boolean
Dim STDType As Integer
With uPreparation.HannaCode

    If .MS1val <> "" Then
        STDType = 1
    ElseIf .MS2Dil <> "" Then
        STDType = 2
    Else
        STDType = 0
    End If

    .STDType = STDType

    uPreparation.MsType = .STDType
    uPreparation.Recipe.Type = .STDType

End With

    Call SetSTDTheoreticalWeight(STDType, uPreparation)




End Function

Public Function SetSTDTheoreticalWeight(ByVal STDType As Integer, ByRef uPreparation As RecipeForProduction)
Select Case STDType
    
    Case 0
        ' MR (LIQUIDO)
        Call PreparationMR(uPreparation)
        
    Case 1
        ' MS1 (SOLIDO)
        Call PreparationMS(uPreparation, 1)
    Case 2
        ' MS2(LIQUIDO)
        Call PreparationMS(uPreparation, 2)



End Select


End Function

Private Function PreparationMR(ByRef uPreparation As RecipeForProduction)
Dim ConcHannaParameter As Double
Dim Unit As String
Dim i As Integer

On Error GoTo ERR_PREP

With uPreparation
    ConcHannaParameter = .HannaCode.ConcHannaParameter
    Unit = .HannaCode.UnitMR
    If ConcHannaParameter > 0 Then
    Else
        .HannaCode.ConcHannaParameter = FindConcParameter(uPreparation)
   
    End If
    
        If .HannaCode.STDVolume = "" Then Exit Function
        If ConcHannaParameter = 0 Then Exit Function
        
        For i = 1 To .Recipe.STDcount
            .Recipe.STD(i).TheoreticalWeight = FormatNumber(.Recipe.STD(i).Value * CDbl(.HannaCode.STDVolume) / (ConcHannaParameter * UmMS(.HannaCode.MeasurementUnit)), 3)     '.HannaCode.Decimal + 1)
        
        Next
        
    
End With

ERR_END:
    
   On Error GoTo 0

    Exit Function
ERR_PREP:
    
    MsgBox Err.Description
    Resume Next
   
    
    
End Function

Private Function FindConcParameter(ByRef uPreparation As RecipeForProduction) As Double
Dim ConcParameter As Double
FindConcParameter = 0
With uPreparation.HannaCode
If .MR.FWParameter <> 0 Then
    With uPreparation.HannaCode
        ConcParameter = (.MR.MRPurity / 100) * .MR.MRValue * .FWHannaParameter / .MR.FWParameter
    End With
End If
End With
FindConcParameter = ConcParameter

End Function


Public Function MSol(ByRef uPreparation As RecipeForProduction, ByVal Index As Integer)
'------------------------------------
' mother solution
'------------------------------------

On Error GoTo ERR_PREP

With uPreparation


    ' Dil e conc li calcolo????
    ' vedi foglio excel
    
    If Index = 2 Then
         .HannaCode.ConcHannaParameter = FindConcParameter(uPreparation)
        .MS.DilConc = .HannaCode.ConcHannaParameter / .HannaCode.MS2Dil
        
         .MS.DilConc = .HannaCode.MS2Dil
        .MS.Volume = .HannaCode.MS2vol
        .MS.Exp = IIf(.HannaCode.MSEXP = "", 1, .HannaCode.MSEXP)
        
        'If (.MS.Volume) = 0 Then .MS.Volume = 100
        .MS.Unit = .HannaCode.MeasurementUnit
        
        .MS.Qty = .MS.Volume / .MS.DilConc
        .MS.Value = (.HannaCode.MR.MRValue * (.HannaCode.MR.MRPurity / 100) * UmMS(.MS.Unit) * .HannaCode.FWHannaParameter) / (.MS.DilConc * .HannaCode.MR.FWParameter)
        .HannaCode.ConcHannaParameter = .MS.Value
        
    
    ElseIf Index = 1 Then
        .MS.DilConc = .HannaCode.MS1val
        .MS.Value = .HannaCode.MS1val
        .MS.Volume = .HannaCode.MS1vol
        .MS.Exp = IIf(.HannaCode.MSEXP = "", 1, .HannaCode.MSEXP)
        'If (.MS.Volume) = 0 Then .MS.Volume = 100
        .MS.Unit = .HannaCode.MeasurementUnit

        .MS.Qty = .MS.Volume * .MS.DilConc / 1000000
       ' .MS.Value = (.HannaCode.FWHannaParameter * (.HannaCode.MR.MRPurity / 100) * .MS.DilConc / .HannaCode.MR.FWParameter)
        .HannaCode.ConcHannaParameter = .MS.Value * UmMS(.MS.Unit) * (.HannaCode.MR.MRPurity / 100) * .HannaCode.FWHannaParameter / .HannaCode.MR.FWParameter

    
    End If
    
End With

ERR_END:
    
   On Error GoTo 0

    Exit Function
ERR_PREP:
    
    MsgBox Err.Description
    Resume Next
   
End Function


Private Function PreparationMS(ByRef uPreparation As RecipeForProduction, ByVal Index As Integer)
Dim ConcHannaParameter As Double
Dim Unit As String
Dim i As Integer

On Error GoTo ERR_PREP

With uPreparation

    Call MSol(uPreparation, Index)

    ConcHannaParameter = .HannaCode.ConcHannaParameter
    Unit = .HannaCode.UnitMR
        If .HannaCode.STDVolume = "" Then Exit Function
        
        For i = 1 To .Recipe.STDcount
            .Recipe.STD(i).TheoreticalWeight = FormatNumber(.Recipe.STD(i).Value * CDbl(.HannaCode.STDVolume) / ConcHannaParameter, 3) '.HannaCode.Decimal + 1)
        
        Next
        
    
End With

ERR_END:
    
   On Error GoTo 0

    Exit Function
ERR_PREP:
    
    MsgBox Err.Description
    Resume Next
   
    
    
End Function

Public Function GetMRCodeFromHannaCode(ByVal HannaCode As String) As String
GetMRCodeFromHannaCode = ""
If HannaCode = "" Then Exit Function
    With dbTabCode
        .filter = ""
        .filter = "Code='" & HannaCode & "'"

        If .EOF Then
        Else
            GetMRCodeFromHannaCode = IIf(IsNull(Trim(!STDMR)), "", Trim(!STDMR))
        End If
    
    End With
End Function


Public Function ColorTolerance(ByVal Variance As Double, ByVal Toll As Double, ByRef bAcquisitiTutti As Boolean, ByRef bCorrection As Boolean) As OLE_COLOR
Dim rc As Boolean
Dim MyColor As OLE_COLOR
    
    rc = False
   
    Select Case Abs(Variance) - Toll
    Case Is <= 0
        ' in tolleranza
        MyColor = &H8000&
        rc = True
    Case TolerancePerc To Toll
        MyColor = vbColorOrange
         rc = True
    Case Toll To Toll * 30
        MyColor = &HC0&
        bCorrection = True
    Case Is > Toll * 30
         rc = False
    Case Else
        rc = False
    
    End Select
    
    bAcquisitiTutti = rc
 
    ColorTolerance = MyColor
    
End Function



Public Function OpenProductClassification(ByVal Code As String, ByVal Index As Integer)
Dim MyID As Long

    If Code = "" Then Exit Function
    
    If Index = 0 Then
        MyID = 1
        
    Else
    
        MyID = GetIDMR(Code)
    
    
    
    End If
    
    If MyID > 0 Then Call F_PICTOGRAM.DoShow(MyID, Index, Code)

End Function




Public Function AddNewRowInAcquisition(ByRef Grid2 As Grid, ByRef iAcquisition As PrepAcquisition)
Dim i As Integer
Dim t As Integer
With Grid2

 
        .AddItem "", False
        i = .Rows - 1
      
        .Cell(i, 1).Text = iAcquisition.Bottle
        .Cell(i, 2).Text = iAcquisition.MRLot
        .Cell(i, 3).Text = iAcquisition.STDNumber
        .Cell(i, 4).Text = iAcquisition.STDValue
        .Cell(i, 5).Text = PadString(iAcquisition.STDQty)
        .Cell(i, 6).Text = iAcquisition.STDUnit
        .Cell(i, 7).Text = PadString(iAcquisition.ActualWeight) ' & " mL"
        .Cell(i, 8).Text = iAcquisition.Note
        .Cell(i, 9).Text = iAcquisition.Operator
        .Cell(i, 10).Text = iAcquisition.AcquisitionTime
        .Cell(i, 11).Text = iAcquisition.ID
        .Cell(i, 12).Text = iAcquisition.LeftInBottle ' & " mL"
        .Cell(i, 13).Text = iAcquisition.CodicePipetta
        .Cell(i, 14).Text = iAcquisition.PipettaType
        
        .Cell(i, 15).Text = iAcquisition.ScaleID
        .Cell(i, 16).Text = iAcquisition.GlasswareID
        
        .Cell(i, 17).Text = iAcquisition.MotherSolutionDate
        
        .Cell(i, 18).Text = iAcquisition.MNP
        .Cell(i, 19).Text = iAcquisition.ExpMR
        
        
        
        
        .Cell(i, 7).BackColor = vbColorResults
        .Cell(i, 7).Alignment = cellRightCenter
        

        .Column(2).Alignment = cellCenterCenter
        .Column(3).Alignment = cellCenterCenter
        .Column(4).Alignment = cellCenterCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellCenterCenter
        .Column(9).Alignment = cellCenterCenter
        .Column(13).Alignment = cellCenterCenter
        .Column(14).Alignment = cellCenterCenter
        .Column(15).Alignment = cellCenterCenter
        
        .Column(16).Alignment = cellCenterCenter
        .Column(17).Alignment = cellCenterCenter
        
        .Cell(i, 12).Alignment = cellRightCenter
        .Cell(i, 12).BackColor = vbColorResults
        
End With

End Function

Public Function AggiornaTabPreparation(ByRef PreparationID As Long, ByRef uPreparation As RecipeForProduction)
With dbTabPreparation
    .filter = ""
    If PreparationID <> 0 Then
        .filter = "ID='" & PreparationID & "'"
        If .EOF Then
            .AddNew
        End If
    Else
        .AddNew
    
    End If

        !HannaCode = uPreparation.HannaCode.Code
        !Description = uPreparation.HannaCode.Description
        !MRCode = uPreparation.MRCode
        !DataPrep = uPreparation.DataPrep
        !HourPrep = uPreparation.HourPrep
        !PrepWeek = uPreparation.PrepWeek
        !Operator = uPreparation.Operator
        !QtyToProduce = uPreparation.QtyToProduce
        !QtyProduced = uPreparation.QtyProduced
        !Unit = "mL"
        !STDMatrix = uPreparation.HannaCode.STDMatrix
        !STDExp = uPreparation.HannaCode.STDExp
        
        !STDStorage = uPreparation.HannaCode.STDStorage
        !bClosed = uPreparation.bClosed
        !Note = uPreparation.Note
        !FileName = uPreparation.FileName
        !MsType = uPreparation.MsType
        !bManuale = uPreparation.bManual
        
    
        .Update
        
        uPreparation.ID = !ID
        PreparationID = !ID

End With

End Function

Public Function SearchIDBottle(ByVal Code As String, ByVal strNumber As String, ByVal strLot As String, ByRef ID As Long) As Boolean
Dim rc As Boolean

    ID = 0
    rc = True
    With dbTabMRWarehouse
        .filter = ""
        .filter = "Code='" & Code & "' and Bottle='" & strNumber & "' and Lot='" & strLot & "'"
        If .EOF Then
        
            rc = False
        Else
        
            ID = !ID
            
        End If
    End With
    
    SearchIDBottle = rc


End Function

Public Function GetAllSTDTotalWeight(ByRef uRecipe As RecipeType, ByRef TotalWeight As Double)
Dim i As Integer
    For i = 1 To uRecipe.STDcount
        With uRecipe.STD(i)
        
           ' .ActualWeight = FormatNumber(.TheoreticalWeight - .RealWeight, 3)
            
            TotalWeight = TotalWeight + .TheoreticalWeight
            
        End With
    
    Next
    
    TotalWeight = FormatNumber(TotalWeight, 3)

End Function

Public Function GetPreparatonID(ByVal HannaCode As String, ByVal DataPrep As String, ByVal HourPrep As String, ByVal PrepWeek As String, ByRef bIsClosed As Boolean) As Long

Dim ID As Long
ID = 0

    With dbTabPreparation
        .filter = ""
        .filter = "HannaCode='" & HannaCode & "' and  DataPrep='" & DataPrep & "' and HourPrep='" & HourPrep & "' and PrepWeek='" & PrepWeek & "'"
        If .EOF Then
        Else
            bIsClosed = !bClosed
            ID = !ID
        End If
        

    End With
    
    GetPreparatonID = ID
    
End Function







Public Function CaricaPipette(ByRef Grd As Grid, Optional ByVal Qty As Double)
Dim i As Integer
Dim t As Integer
Dim y As Integer
Dim MaxCount As Integer
Dim PipMin As Double
Dim PipMax As Double




    On Error GoTo ERR_GRID
    ' --------------------------------------
    '
    '  filtra TabPipette e riempi Tabella
    '
    ' --------------------------------------
  

    With Grd
        .Rows = 1
        .ReadOnly = True
        .AutoRedraw = False
        With dbTabPipette
        
            .filter = ""
            .filter = ""
            
            If .EOF Then Exit Function
            MaxCount = .RecordCount
            .MoveFirst
        End With
        
        i = 0
        
        For y = 1 To MaxCount
        
        
            
            
        
            PipMin = IIf(IsNull(Trim(dbTabPipette!VolMin)), 0, Trim(dbTabPipette!VolMin))
            PipMax = IIf(IsNull(Trim(dbTabPipette!VolMax)), 0, Trim(dbTabPipette!VolMax))
            
            If Qty >= PipMin And Qty <= PipMax Then
        
            Else
                GoTo cont:
            End If
        
        
        
            .AddItem "", False
            i = i + 1
            .Cell(i, 0).Text = i
            .Cell(i, 1).Text = "  " & IIf(IsNull(Trim(dbTabPipette!Equipment)), "", Trim(dbTabPipette!Equipment))
            .Cell(i, 2).Text = "  " & IIf(IsNull(Trim(dbTabPipette!VolumeAdjustment)), "", Trim(dbTabPipette!VolumeAdjustment))
            .Cell(i, 3).Text = IIf(IsNull(Trim(dbTabPipette!Characteristic)), "", Trim(dbTabPipette!Characteristic))
            .Cell(i, 4).Text = dbTabPipette!ID
            .Cell(i, 5).Text = IIf(IsNull(Trim(dbTabPipette!VolMin)), "", Trim(dbTabPipette!VolMin))
            .Cell(i, 6).Text = IIf(IsNull(Trim(dbTabPipette!VolMax)), "", Trim(dbTabPipette!VolMax))
            .Cell(i, 7).Text = IIf(IsNull(Trim(dbTabPipette!Unit)), "", Trim(dbTabPipette!Unit))

                If i > 1 Then
                
                If .Cell(i, 1).Text = .Cell(i - 1, 1).Text Then
                   For t = 1 To .Cols - 1
                    .Cell(i, t).BackColor = vbColorTextLightBlue
                   Next
                End If

            End If
cont:
            dbTabPipette.MoveNext
        Next
ERR_END:

        .RowHeight(0) = 35
        
        .Column(1).Alignment = cellLeftCenter
        .Column(1).Alignment = cellLeftCenter
                   
        .Column(1).AutoFit
        .Column(2).AutoFit
        .AutoRedraw = True
        .Refresh
    End With

    Exit Function
ERR_GRID:
    MessageInfoTime = 2000
   PopupMessage 2, Err.Description
   GoTo ERR_END:
End Function

