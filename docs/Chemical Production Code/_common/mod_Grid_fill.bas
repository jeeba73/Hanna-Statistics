Attribute VB_Name = "mod_Grid_fill"
Option Explicit



Public Function GetMeanTable(ByRef Grd3 As Grid, ByVal SettingName As String)
    
Dim i As Integer
Dim t As Integer
Dim nRows As Long
Dim nCols As Long
        '  tabella Results

    With Grd3
        .AutoRedraw = False
        nRows = GetSettingData(SettingName, "Evaluation QC", "Results Grid Rows", .Rows)
        nCols = GetSettingData(SettingName, "Evaluation QC", "Results Grid Cols", .Cols)
        .Rows = nRows
        .Cols = nCols
        For i = 1 To .Rows - 1
            For t = 1 To .Cols - 1
                .Cell(i, t).Text = GetSettingData(SettingName, "Evaluation QC", "Results Grid Standard (" & i & ")  Column " & t, .Cell(i, t).Text)
                .Cell(i, t).ForeColor = GetSettingData(SettingName, "Evaluation QC", "Results Grid Standard (" & i & ") Forecolor " & t, vbBlack)
            Next
        Next
        .AutoRedraw = True
        .Refresh
    End With
    
    
    CloseSettingDataFile
End Function
