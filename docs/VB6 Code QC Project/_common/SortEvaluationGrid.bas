Attribute VB_Name = "SortEvaluationGrid"
Option Explicit




Public Function SortGrid(ByRef Grd3 As Grid)

Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim STDValue As String
Dim STDNum As Integer
Dim Cell() As String
Dim CellForecolor() As String
Dim temp() As String
Dim tempForecolor() As String
Dim j As Integer


With Grd3
     If .Rows <= 1 Then Exit Function

    .ReadOnly = True
    .AutoRedraw = False
   
    ReDim Cell(.Rows, .Cols)
    ReDim temp(.Cols)
    ReDim tempForecolor(.Cols)
    ReDim CellForecolor(.Rows, .Cols)
    For i = 1 To .Rows - 1
        For t = 1 To .Cols - 1
            Cell(i, t) = .Cell(i, t).Text
            CellForecolor(i, t) = .Cell(i, t).ForeColor
        Next
    Next
    
   
    For i = 1 To .Rows - 1
        For j = i + 1 To .Rows - 1
            If CDbl(Cell(i, 3)) > CDbl(Cell(j, 3)) Then
                For t = 1 To .Cols - 1
                    temp(t) = Cell(i, t)
                    tempForecolor(t) = CellForecolor(i, t)
                    Cell(i, t) = Cell(j, t)
                    CellForecolor(i, t) = CellForecolor(j, t)
                    Cell(j, t) = temp(t)
                    CellForecolor(j, t) = tempForecolor(t)
                Next
            End If
        Next j
    Next i
    

    For i = 1 To .Rows - 1
        For t = 1 To .Cols - 1
             .Cell(i, t).Text = Cell(i, t)
             .Cell(i, t).ForeColor = CellForecolor(i, t)
        Next
    Next
    
fine:
   ' .Column(6).Sort cellAscending
    .AutoRedraw = True
    .Refresh
End With



End Function



