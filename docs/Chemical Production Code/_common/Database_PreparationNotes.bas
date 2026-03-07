Attribute VB_Name = "Database_PreparationNotes"
Option Explicit
Public Function GetProductionNotes(ByVal Grd As Grid, ByVal FileName As String) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer


        
    If FileName = "" Then Exit Function
        
    With Grd
        .Rows = 1
        .AutoRedraw = False
    
        With dbTabProductionNotes
            .Close
            .Open "SELECT *  FROM TabProductionNotes order by NoteDate", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
            .filter = ""
            .filter = "FileName='" & FileName & "'"
            If .EOF Then
            Else
                .MoveLast
                Grd.DefaultRowHeight = 40
                For i = 1 To .RecordCount
                    Grd.AddItem "", False
                    t = Grd.Rows - 1
                    Grd.Cell(t, 1).Text = FormatDataLAT(IIf(IsNull(Trim(!NoteDate)), "", Trim(!NoteDate)))
                    Grd.Cell(t, 2).Text = IIf(IsNull(Trim(!Type)), "", Trim(!Type))
                    Grd.Cell(t, 3).Text = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
                    Grd.Cell(t, 3).FontSize = 8
                    Grd.Cell(t, 3).WrapText = True
                    Grd.Cell(t, 4).Text = IIf(IsNull(Trim(!Operator)), "", Trim(!Operator))

                    Grd.Cell(t, 5).Text = !ID
                  
                    .MovePrevious
                Next
                
            End If
    
    
    
        End With
        .SelectionMode = cellSelectionByRow
        .Column(1).Alignment = cellCenterCenter
        .Column(1).Width = 100
        .Column(2).Width = 200
        .Column(3).Width = 700
        .Column(4).Width = 100
        .Column(5).Width = 0
        .Refresh
        .AutoRedraw = True
        
    End With



End Function

Public Function GetPreparationNotes(ByVal Grd As Grid, ByVal FileName As String) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer


        
    If FileName = "" Then Exit Function
        
    With Grd
        .Rows = 1
        .AutoRedraw = False
    
        With dbTabPreparationNotes
            .Close
            .Open "SELECT *  FROM TabPreparationNotes order by NoteDate", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
            .filter = ""
            .filter = "FileName='" & FileName & "'"
            If .EOF Then
            Else
                .MoveLast
                Grd.DefaultRowHeight = 40
                For i = 1 To .RecordCount
                    Grd.AddItem "", False
                    t = Grd.Rows - 1
                    Grd.Cell(t, 1).Text = FormatDataLAT(IIf(IsNull(Trim(!NoteDate)), "", Trim(!NoteDate)))
                    Grd.Cell(t, 2).Text = IIf(IsNull(Trim(!Type)), "", Trim(!Type))
                    Grd.Cell(t, 3).Text = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
                    Grd.Cell(t, 3).FontSize = 8
                    Grd.Cell(t, 3).WrapText = True
                    Grd.Cell(t, 4).Text = IIf(IsNull(Trim(!Operator)), "", Trim(!Operator))

                    Grd.Cell(t, 5).Text = !ID
                  
                    .MovePrevious
                Next
                
            End If
    
    
    
        End With
        .SelectionMode = cellSelectionByRow
        .Column(1).Alignment = cellCenterCenter
        .Column(1).Width = 100
        .Column(2).Width = 200
        .Column(3).Width = 700
        .Column(4).Width = 100
        .Column(5).Width = 0
        .Refresh
        .AutoRedraw = True
        
    End With



End Function


Public Sub SetGridNotes(ByVal Grd As Grid)
Dim i As Integer
    With Grd
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 6
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Date"
        .Cell(0, 2).Text = "Type"
        .Cell(0, 3).Text = "Description"
        .Cell(0, 4).Text = "Operator"
        .Cell(0, 5).Text = "ID"
     


        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 150
            .Cell(0, i).FontBold = True
            
        Next
        
        .Column(1).Width = 80
        .Column(2).Width = 100
        .Column(4).Width = 300
        
        .Column(4).Width = 80
        .Column(5).Width = 0
        
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
  

End Sub

