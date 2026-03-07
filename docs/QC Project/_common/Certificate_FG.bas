Attribute VB_Name = "Certificate_FG"
Option Explicit



Public Type RegSet
    Code As String
    Lot As String
    Exp As String


End Type

Public Type CertResult
    MaxCount            As Integer
    STDValue()          As Double
    STDLotResult()      As Double
    STDYcalc()          As Double
    Uncertainty()       As Double
    UncertPerc()        As Double
    
    MinValue            As Double
    MaxValue            As Double
    MedValue            As Double
    
    a                   As Double
    b                   As Double
    
    df                  As Double
    sy                  As Double
    ssx                 As Double
    MethodStDeviation   As Double
    MethodVariation     As Double
    
    RSS                 As Double
    TSS                 As Double
    r                   As Double
    n                   As Integer
    tval                As Double ' fattore di copertura
    
    ConfidenceInterval  As Double

End Type

Public Type GraphType

    ' Incertezza tipo taratura
    
    Replicat            As Integer
    x()                 As Double
    Ypred()             As Double
    sYpre()             As Double
    Lplim()             As Double
    Uplim()             As Double
    LplimGrph()         As Double
    UplimGrph()         As Double
        
End Type

Public Type STDCert
    STDValue            As Double
    AverageResult       As Double
    Passed              As Boolean
    
End Type

Public Type ResType
    TargetValue         As Double
    LotValue            As Double
    Passed              As Boolean
    
End Type

Public Type CrtSTD
    Value               As String
    Average             As String
    gdl                 As ResType
    Slope               As ResType
    Intersect           As ResType
    ReagentBlank        As ResType
    Variation           As ResType
    Confidence          As ResType
    StdDeviation        As ResType

End Type

Public Type CertType
    ProductName         As String
    ProductCode         As String
    Method              As String
    RangePPM            As String
    LotNumber           As String
    BestUseBefore       As String
    DateAnalisys        As String
    ReferenceMeter      As String
    ReferenceMeterDescription      As String
    ReferenceSTD        As String
    Wavelenght          As String
    CellMM              As String
    RefSTDNote1         As String
    RefSTDNote2         As String
    RangeFormula        As String
    CalibrationFunction As CrtSTD
    LotCalculation      As CertResult
    STD()               As STDCert
    GraphCert           As GraphType
    
    
    UserDecimal         As String

End Type

Public myCertificate As CertType
Public MyCertificateClean As CertType


Public Function SetGridFGCodeCertificate(ByRef Grd As Grid) As Boolean
Dim i As Integer
Dim MaxRows As Long
    With Grd
        .Rows = 1
        .AutoRedraw = False
        .ReadOnly = False
        .DisplayFocusRect = True 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionFree
        .DefaultFont.Size = 9 '* m_ControlGridFontSize
        
        .DefaultFont.Bold = False
        .DefaultRowHeight = 35
        .ExtendLastCol = True
        .Cols = 3
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Column(1).Width = 250
        .Column(2).Width = 250
        
        .RowHeight(0) = 0
        
        .Rows = 16
       
        
        
         .Cell(1, 1).Text = "Product Name"
         .Cell(2, 1).Text = "Product Code"
         .Cell(3, 1).Text = "Method"
         .Cell(4, 1).Text = "Range ppm (as O2)"
         .Cell(5, 1).Text = "Lot number"
         .Cell(6, 1).Text = "Best use before"
         .Cell(7, 1).Text = "Date of analysis"
         .Cell(8, 1).Text = "Reference meter"
         .Cell(9, 1).Text = "Reference standard"
         .Cell(10, 1).Text = "Wavelenght nm"
         .Cell(11, 1).Text = "Cell mm"
         .Cell(12, 1).Text = "Reference standard Note 1"
         .Cell(13, 1).Text = "Reference standard Note 2"
         .Cell(14, 1).Text = "Reference meter Description"
         .Cell(15, 1).Text = "Range Formula"
      
        
       
        
        
        For i = 1 To .Rows - 1
        
            .Cell(i, 1).BackColor = &HF0F0F0 'vbColorUnabled
            .Cell(i, 1).ForeColor = vbColorDarkFont 'vbColorDarkFont 'vbColorForeFixed  ' vbColorBlueProgram
            .Cell(i, 1).FontBold = False
            .Cell(i, 1).Locked = True
            .Cell(i, 2).Locked = False
            .Cell(i, 2).ForeColor = vbColorDarkFont
             .Cell(i, 2).WrapText = True

        Next

        .Cell(1, 2).Locked = False
        .Cell(2, 2).Locked = False

        
        
        .ReadOnly = False
        .AutoRedraw = True
        .Refresh
        
    End With
End Function


Public Sub CopyFGCodeCertificateGrd2(ByVal Grd2 As Grid, ByVal lId As Long, ByRef myCertificate As CertType)
    If lId = 0 Then Exit Sub
    Dim i As Integer


     With dbTabFinishGood
            
        .filter = ""
        .filter = "ID='" & lId & "'"
        If .EOF Then Exit Sub
        .MoveFirst
        
        
    End With
    
    
    With Grd2

         
         .Cell(1, 2).Text = IIf(IsNull(Trim(dbTabFinishGood!Description)), "", Trim(dbTabFinishGood!Description))
         .Cell(2, 2).Text = IIf(IsNull(Trim(dbTabFinishGood!Code)), "", Trim(dbTabFinishGood!Code))
         .Cell(3, 2).Text = IIf(IsNull(Trim(dbTabFinishGood!Method)), "", Trim(dbTabFinishGood!Method))
         .Cell(4, 2).Text = IIf(IsNull(Trim(dbTabFinishGood!RangePPM)), "", Trim(dbTabFinishGood!RangePPM))
         .Cell(5, 2).Text = ""
         .Cell(6, 2).Text = ""
         .Cell(7, 2).Text = myCertificate.DateAnalisys
         .Cell(8, 2).Text = IIf(IsNull(Trim(dbTabFinishGood!RefMeter)), "", Trim(dbTabFinishGood!RefMeter))
         .Cell(9, 2).Text = IIf(IsNull(Trim(dbTabFinishGood!RefSTD)), "", Trim(dbTabFinishGood!RefSTD))
         .Cell(10, 2).Text = IIf(IsNull(Trim(dbTabFinishGood!Wavelength)), "", Trim(dbTabFinishGood!Wavelength))
         .Cell(11, 2).Text = IIf(IsNull(Trim(dbTabFinishGood!Cell)), "", Trim(dbTabFinishGood!Cell))

         .Cell(12, 2).Text = IIf(IsNull(Trim(dbTabFinishGood!RefSTDNote)), "", Trim(dbTabFinishGood!RefSTDNote))
         .Cell(13, 2).Text = IIf(IsNull(Trim(dbTabFinishGood!RefSTDNote2)), "", Trim(dbTabFinishGood!RefSTDNote2))
         .Cell(14, 2).Text = IIf(IsNull(Trim(dbTabFinishGood!RefMeterDescription)), "", Trim(dbTabFinishGood!RefMeterDescription))
         .Cell(15, 2).Text = IIf(IsNull(Trim(dbTabFinishGood!RangeFormula)), "", Trim(dbTabFinishGood!RangeFormula))

        
        
        For i = 2 To .Rows - 1
            .AutoFitRowHeight (i)
            .Cell(i, 2).Locked = True
        Next
        
        .Cell(5, 2).Locked = False
        .Cell(6, 2).Locked = False
        .Cell(7, 2).Locked = False
        
        .Cell(5, 1).FontBold = True
        .Cell(6, 1).FontBold = True
        .Cell(7, 1).FontBold = True
        
      
        
        For i = 2 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            
            
        Next
           .Cell(1, 2).BackColor = vbColorAzzurrino
       .Cell(2, 2).BackColor = vbColorAzzurrino
        
        .Column(1).AutoFit
    
    End With

End Sub
Public Function SetGridFGCodeCalFunction(ByRef Grd As Grid) As Boolean
Dim i As Integer
Dim MaxRows As Long
    With Grd
        .Rows = 1
        .AutoRedraw = False
        .ReadOnly = False
        .DisplayFocusRect = True 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionFree
        .DefaultFont.Size = 9 '* m_ControlGridFontSize
        
        .DefaultFont.Bold = False
        .DefaultRowHeight = 35
        .ExtendLastCol = False
       
        
        .Cols = 4
        .Rows = 8
        
        .DefaultRowHeight = 40
       
        .Cell(0, 1).Text = "Target Value"
        .Cell(0, 2).Text = "Lot Value"
        .Cell(0, 3).Text = "Passed"
        
        .Cell(1, 0).Text = "n"
        .Cell(2, 0).Text = "Standard Deviation"
        .Cell(3, 0).Text = "Method variation coefficient [%]"
        .Cell(4, 0).Text = "Confidence interval (95%)"

        .Cell(5, 0).Text = "Slope"
        .Cell(6, 0).Text = "Ordinate intersect ppm"
        .Cell(7, 0).Text = "Blank Value [Absorbance]"
        
        
        
      
             .Column(0).Alignment = cellLeftCenter
             .Column(1).Alignment = cellCenterCenter
             .Column(2).Alignment = cellCenterCenter
             .Column(3).Alignment = cellCenterCenter
      
        
           For i = 0 To .Cols - 1
        
            .Column(i).Width = 150
            
            
        Next
        
        
        
        
        For i = 0 To .Rows - 1
        
            .Cell(i, 0).WrapText = True
           
            .Cell(i, 1).ForeColor = vbColorDarkFont 'vbColorDarkFont 'vbColorForeFixed  ' vbColorBlueProgram
            .Cell(i, 1).FontBold = False
            
            .AutoFitRowHeight (i)
             
        Next

        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
End Function


Public Function SetGridFGCodeLotResult(ByRef Grd As Grid) As Boolean
Dim i As Integer
Dim MaxRows As Long
    With Grd
        .Rows = 1
        .AutoRedraw = False
        .ReadOnly = False
        .DisplayFocusRect = True 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionFree
        .DefaultFont.Size = 9 '* m_ControlGridFontSize
        
        .DefaultFont.Bold = False
        .DefaultRowHeight = 35
        .ExtendLastCol = False
       
        
        .Cols = 4

        
        .Column(0).Width = 30
        .Cell(0, 0).Text = "STD"
        .Cell(0, 1).Text = "Standard value ppm"
        .Cell(0, 2).Text = "Average Result ppm"
        .Cell(0, 3).Text = "k"
        
      
      
             .Column(0).Alignment = cellCenterCenter
             .Column(1).Alignment = cellCenterCenter
          
        
        For i = 0 To .Cols - 1
        
            .Column(i).Width = 150
            
            
        Next
        
         
        
        For i = 0 To .Rows - 1
        
            .Cell(i, 0).WrapText = True
           
            .Cell(i, 1).ForeColor = vbColorDarkFont 'vbColorDarkFont 'vbColorForeFixed  ' vbColorBlueProgram
            .Cell(i, 1).FontBold = False
            
            .AutoFitRowHeight (i)
             
        Next

        .Column(3).Width = 0
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
End Function




Public Function SetGridFGCodeLotCalculationLinearRegression(ByRef Grd As Grid) As Boolean
Dim i As Integer
Dim MaxRows As Long
    With Grd
        .Rows = 2
        .AutoRedraw = False
        .ReadOnly = False
        .DisplayFocusRect = True 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionFree
        .DefaultFont.Size = 9 '* m_ControlGridFontSize
        
        .DefaultFont.Bold = False
        .DefaultRowHeight = 35
        .ExtendLastCol = True
       
        
        .Cols = 11

        
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Ordinate intersect a (ABS)"
        .Cell(0, 2).Text = "Sensitivity (slope) b (mg/L)"
        .Cell(0, 3).Text = "df"
        .Cell(0, 4).Text = "t crit."
        .Cell(0, 5).Text = "MedValue"
        .Cell(0, 6).Text = "Residual STD dev s(y) (ABS)"
        .Cell(0, 7).Text = "SSx"
        .Cell(0, 8).Text = "Method standard deviation (mg/L)"
        .Cell(0, 9).Text = "Method variation coefficient %"
        .Cell(0, 10).Text = "Max Uncertainty sample "
        
      
      
        
        For i = 0 To .Cols - 1
        
            .Column(i).Width = 150
            .Column(i).Alignment = cellCenterCenter
            .Cell(0, i).WrapText = True
        Next
        
         
     
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
End Function

Public Function SetGridReagentSet(ByRef Grd As Grid) As Boolean
Dim i As Integer
Dim MaxRows As Long
    With Grd
        .Rows = 4
        .AutoRedraw = False
        .ReadOnly = False
        .DisplayFocusRect = True 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionFree
        .DefaultFont.Size = 9 '* m_ControlGridFontSize
        
        .DefaultFont.Bold = False
        .DefaultRowHeight = 35
        .ExtendLastCol = False
       
        
        .Cols = 6
        .Rows = 4
        
        .RowHeight(0) = 0
        
        '.Column(0).Width = 0
        '.Cell(0, 1).Text = ""
        '.Cell(0, 2).Text = "Ypred"
        '.Cell(0, 3).Text = "s(Ypre.)"
        '.Cell(0, 4).Text = "Lplim(Y)"
        '.Cell(0, 5).Text = "Uplim(Y)"
        
      
        
        For i = 0 To .Cols - 1
        
            .Column(i).Width = 120
            .Column(i).Alignment = cellCenterCenter
            .Cell(0, i).WrapText = True
        Next
        
         
     
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
End Function



Public Function SetGridCalibrationUncertainty(ByRef Grd As Grid) As Boolean
Dim i As Integer
Dim MaxRows As Long
    With Grd
        .Rows = 2
        .AutoRedraw = False
        .ReadOnly = False
        .DisplayFocusRect = True 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionFree
        .DefaultFont.Size = 9 '* m_ControlGridFontSize
        
        .DefaultFont.Bold = False
        .DefaultRowHeight = 35
        .ExtendLastCol = False
       
        
        .Cols = 6

        
        .Column(0).Width = 10
        .Cell(0, 1).Text = "X"
        .Cell(0, 2).Text = "Ypred"
        .Cell(0, 3).Text = "s(Ypre.)"
        .Cell(0, 4).Text = "Lplim(Y)"
        .Cell(0, 5).Text = "Uplim(Y)"
        
      
        
        For i = 0 To .Cols - 1
        
            .Column(i).Width = 120
            .Column(i).Alignment = cellCenterCenter
            .Cell(0, i).WrapText = True
        Next
        
         
     
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
End Function



Public Function SetGridFGCodeLotCalculation(ByRef Grd As Grid) As Boolean
Dim i As Integer
Dim MaxRows As Long
    With Grd
        .Rows = 1
        .AutoRedraw = False
        .ReadOnly = False
        .DisplayFocusRect = True 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionFree
        .DefaultFont.Size = 9 '* m_ControlGridFontSize
        
        .DefaultFont.Bold = False
        .DefaultRowHeight = 35
        .ExtendLastCol = False
       
        
        .Cols = 8

        
        .Column(0).Width = 30
        .Cell(0, 1).Text = "Level"
        .Cell(0, 2).Text = "Standard value"
        .Cell(0, 3).Text = "Lot Result ABS"
        .Cell(0, 3).WrapText = True
        .Cell(0, 4).Text = "Ycalc"
        
        .Cell(0, 5).Text = "u"
        .Cell(0, 6).Text = "u (%)"
        .Cell(0, 7).Text = "U"
        
        
      
      
             .Column(0).Alignment = cellCenterCenter
             .Column(1).Alignment = cellCenterCenter
          
        
        For i = 0 To .Cols - 1
        
            .Column(i).Width = 150
            
            
        Next
        
         
        
        For i = 0 To .Rows - 1
        
            .Cell(i, 0).WrapText = True
           
           ' .Cell(i, 1).ForeColor = vbColorDarkFont 'vbColorDarkFont 'vbColorForeFixed  ' vbColorBlueProgram
            '.Cell(i, 1).FontBold = False
            
           ' .AutoFitRowHeight (i)
             
        Next

      '  .Column(3).Width = 0
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
End Function

Public Sub CopyFGCodeCertificateGrid2(ByVal Grd2 As Grid, ByVal lId As Long, ByRef CalibrationFunction As CrtSTD)
    If lId = 0 Then Exit Sub
    Dim i As Integer



     With dbTabFinishGood
            
        .filter = ""
        .filter = "ID='" & lId & "'"
        If .EOF Then Exit Sub
        .MoveFirst
        
        
    End With
    
    
    With Grd2

        .AutoRedraw = False
        
        
        ' .Cell(1, 0).Text = "n"
       ' .Cell(2, 0).Text = "Standard Deviation [mg/L NH4+]"
       ' .Cell(3, 0).Text = "Method variation coefficient [%]"
       ' .Cell(4, 0).Text = "Confidence interval (95%) [mg/L NH4+]"

       ' .Cell(5, 0).Text = "Slope"
       ' .Cell(6, 0).Text = "Ordinate intersect ppm"
       ' .Cell(7, 0).Text = "Blank Value [Absorbance]"
       

       
       ' .Cell(2, 0).Text = "Standard Deviation [mg/L NH4+]"
       ' .Cell(3, 0).Text = "Method variation coefficient [%]"
        '.Cell(4, 0).Text = "Confidence interval (95%) [mg/L NH4+]"
        
        
        .Cell(1, 1).Text = IIf(IsNull(Trim(dbTabFinishGood!gdl)), 0, Trim(dbTabFinishGood!gdl))
        .Cell(2, 1).Text = IIf(IsNull(Trim(dbTabFinishGood!StdDeviation)), 0, Trim(dbTabFinishGood!StdDeviation))
        .Cell(3, 1).Text = IIf(IsNull(Trim(dbTabFinishGood!MethVar)), 0, Trim(dbTabFinishGood!MethVar))
        .Cell(4, 1).Text = IIf(IsNull(Trim(dbTabFinishGood!ConfInt)), 0, Trim(dbTabFinishGood!ConfInt))
        
        .Cell(5, 1).Text = IIf(IsNull(Trim(dbTabFinishGood!Slope)), 0, Trim(dbTabFinishGood!Slope))
        .Cell(6, 1).Text = IIf(IsNull(Trim(dbTabFinishGood!OrdinateIntersect)), 0, Trim(dbTabFinishGood!OrdinateIntersect))
        .Cell(7, 1).Text = IIf(IsNull(Trim(dbTabFinishGood!ReagentBlank)), 0, Trim(dbTabFinishGood!ReagentBlank))

        
        
       If .Cell(1, 1).Text <> "" Then CalibrationFunction.gdl.TargetValue = CDbl(.Cell(1, 1).Text)
       If .Cell(2, 1).Text <> "" Then CalibrationFunction.StdDeviation.TargetValue = CDbl(.Cell(2, 1).Text)
       If .Cell(3, 1).Text <> "" Then CalibrationFunction.Variation.TargetValue = CDbl(.Cell(3, 1).Text)
       If .Cell(4, 1).Text <> "" Then CalibrationFunction.Confidence.TargetValue = CDbl(.Cell(4, 1).Text)
       If .Cell(5, 1).Text <> "" Then CalibrationFunction.Slope.TargetValue = CDbl(.Cell(5, 1).Text)
       If .Cell(6, 1).Text <> "" Then CalibrationFunction.Intersect.TargetValue = CDbl(.Cell(6, 1).Text)
       If .Cell(7, 1).Text <> "" Then CalibrationFunction.ReagentBlank.TargetValue = CDbl(.Cell(7, 1).Text)

        
        
       
         .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    
    End With

End Sub
