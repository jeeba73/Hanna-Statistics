Attribute VB_Name = "mod_SearchInGrid"
Option Explicit


Public Function SearchInGrid(ByRef Grid As Grid, ByVal str As String, ByVal bShowAll As Boolean, Optional ByVal Column As Integer = 1)

Dim i As Integer
Dim UserColumn As Integer

UserColumn = IIf(Column > 0, Column, 1)

str = Trim(str)

If str = "" Then bShowAll = True

With Grid
    .AutoRedraw = False
    If .Rows > 1 Then
        
        For i = 1 To .Rows - 1
            If bShowAll Then
                 .RowHeight(i) = 25
            Else
                If InStr(UCase(.Cell(i, UserColumn).Text), UCase(str)) Then
                    
                    .RowHeight(i) = 25
                
                Else
                    .RowHeight(i) = 0
                
                End If
            End If
            
        
        Next

    End If
    .Refresh
    .AutoRedraw = True
End With
End Function
