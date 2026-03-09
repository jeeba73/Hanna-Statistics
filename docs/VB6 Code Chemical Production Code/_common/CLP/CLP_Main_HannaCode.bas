Attribute VB_Name = "CLP_Main_HannaCode"
Option Explicit


Public Sub AddHannaCodeinGrid(ByVal Grd As Grid)
Dim i As Integer
Dim sString As String
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        With dbTabCode
            .filter = ""
            
           ' If txCode <> "" Then sString = "Code like '*" & Trim(txCode) & "*'"
           ' If InStr(LCase(cmbLine), "all lines") Then
           '
           ' Else
            
            If InStr(LCase(UserLine), "all lines") Or UserLine = "" Then
             
            Else
                sString = " line='" & UserLine & "'"
           End If
           .filter = sString
           
           If .EOF Then
           
           Else
           
                .MoveFirst
           
                For i = 1 To .RecordCount
                    Grd.AddItem "", False
                    
        '.Cell(0, 2).Text = "Product Name"
        '.Cell(0, 3).Text = "Line"
        '.Cell(0, 4).Text = "Volume/Weight"
        '.Cell(0, 5).Text = "um"
        '.Cell(0, 6).Text = "Q.ty to produce"
        '.Cell(0, 7).Text = "Recipe"
        '.Cell(0, 8).Text = "Mix #1"
        'p.Cell(0, 9).Text = "Mix #2"
                    Grd.Cell(Grd.Rows - 1, 0).Text = i
                    Grd.Cell(Grd.Rows - 1, 1).Text = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
                    Grd.Cell(Grd.Rows - 1, 2).Text = IIf(IsNull(Trim(!ProductName)), "", Trim(!ProductName))
                    Grd.Cell(Grd.Rows - 1, 3).Text = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
                    Grd.Cell(Grd.Rows - 1, 4).Text = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
                    Grd.Cell(Grd.Rows - 1, 5).Text = IIf(IsNull(Trim(!Mix1)), "", Trim(!Mix1))
                    Grd.Cell(Grd.Rows - 1, 6).Text = IIf(IsNull(Trim(!Mix2)), "", Trim(!Mix2))
                    Grd.Cell(Grd.Rows - 1, 7).Text = !ID
                    .MoveNext
                Next
           End If
        
        
        End With
        
        
        Dim t As Integer
        For t = 1 To .Rows - 1
            
            For i = 1 To .Cols - 1
                .Column(i).Alignment = IIf(i > 2, cellCenterCenter, cellLeftCenter)
            Next
       Next
     
        '.Column(1).AutoFit
        '.Column(4).AutoFit
        '.Column(5).AutoFit
        '.Column(6).AutoFit
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
    

End Sub

