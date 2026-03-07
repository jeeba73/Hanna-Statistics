Attribute VB_Name = "Database_RecipeRevision"
Option Explicit

Public Function GetRecipeRevision(ByVal Grd As Grid, ByVal Recipe As String) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer


    Recipe = Trim(Recipe)
        
    If Recipe = "" Then Exit Function
        
    With Grd
        .Rows = 1
        .AutoRedraw = False
    
        With dbTabRecipeRevisionHistory
            .Close
            .Open "SELECT *  FROM TabRecipeRevisionHistory order by RevDate", dbChemicalQC, adOpenKeyset, adLockOptimistic, adCmdText
            .filter = ""
            .filter = "Recipe='" & Replace(Recipe, "'", "''") & "'"
            If .EOF Then
            Else
                .MoveLast
                Grd.DefaultRowHeight = 50
                For i = 1 To .RecordCount
                    Grd.AddItem "", False
                    t = Grd.Rows - 1
                    Grd.Cell(t, 1).Text = FormatDataLAT(IIf(IsNull(Trim(!RevDate)), "", Trim(!RevDate)))
                    Grd.Cell(t, 2).Text = IIf(IsNull(Trim(!RevNumber)), "", Trim(!RevNumber))
                    Grd.Cell(t, 3).Text = IIf(IsNull(Trim(!type)), "", Trim(!type))
                    Grd.Cell(t, 4).Text = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
                    Grd.Cell(t, 4).FontSize = 8
                    Grd.Cell(t, 4).WrapText = True
                    Grd.Cell(t, 5).Text = IIf(IsNull(Trim(!Operator)), "", Trim(!Operator))
                    Grd.Cell(t, 6).Text = !ID
                  
                    .MovePrevious
                Next
                
            End If
    
    
    
        End With
        .SelectionMode = cellSelectionByRow
        .Column(2).Alignment = cellCenterCenter
        
        .Column(2).Width = 200
        .Column(3).Width = 100
        .Column(4).Width = 500
        .Refresh
        .AutoRedraw = True
        
    End With



End Function


Public Sub SetGridRecipeRevision(ByVal Grd As Grid)
Dim i As Integer
    With Grd
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 11
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Rev. Date"
        .Cell(0, 2).Text = "Rev. Number"
        
        
        .Cell(0, 3).Text = "Type"
        .Cell(0, 4).Text = "Description"
        .Cell(0, 5).Text = "Operator"
        .Cell(0, 6).Text = "ID"
     


        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 150
            .Cell(0, i).FontBold = True
            
        Next
        
        .Column(1).Width = 80
        .Column(2).Width = 80
        .Column(3).Width = 100
        
        .Column(4).Width = 300
        .Column(5).Width = 80
        
        .Column(6).Width = 0
        
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
  

End Sub
