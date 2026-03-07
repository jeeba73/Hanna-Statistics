Attribute VB_Name = "Evaluation"
Option Explicit

Function CalcolaSTDEV(arr() As Double) As Double
    Dim somma As Double
    Dim sommaQuadrati As Double
    Dim conta As Integer
    Dim valore As Double
    Dim i As Integer

    somma = 0
    sommaQuadrati = 0
    conta = UBound(arr) - LBound(arr) + 1

    For i = LBound(arr) To UBound(arr)
        valore = arr(i)
        somma = somma + valore
        sommaQuadrati = sommaQuadrati + valore * valore
    Next i

    If conta > 1 Then
        CalcolaSTDEV = Sqr((sommaQuadrati - somma * somma / conta) / (conta - 1))
    Else
        CalcolaSTDEV = 0
    End If
End Function


Public Function SetGridEvaluationResults(ByRef Grd As Grid) As Boolean
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
        .Column(1).Width = 180
        .Column(2).Width = 20
        
        .RowHeight(0) = 0
        
        .Rows = 11
       
        
        
         .Cell(1, 1).Text = "STD  "
         .Cell(2, 1).Text = "# Readings  "
         .Cell(3, 1).Text = "# Tests  "
         .Cell(4, 1).Text = "Total Average  "
         .Cell(5, 1).Text = "# Readings (Sel)  "
         .Cell(6, 1).Text = "# Tests (Sel)  "
         .Cell(7, 1).Text = "Mean Value (Sel)  "
         .Cell(8, 1).Text = "Std Deviation  "
         .Cell(9, 1).Text = "Std Deviation %  "
         .Cell(10, 1).Text = "Repeatability  "
       
      
        
       
        
        
        For i = 1 To .Rows - 1
        
            .Cell(i, 1).BackColor = &HF0F0F0 'vbColorUnabled
            .Cell(i, 1).ForeColor = vbColorDarkFont 'vbColorDarkFont 'vbColorForeFixed  ' vbColorBlueProgram
            .Cell(i, 1).FontBold = False
            .Cell(i, 1).Locked = True
            .Cell(i, 2).ForeColor = vbColorDarkFont
            .Cell(i, 1).Alignment = cellRightCenter
         
        Next


        
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With

End Function
