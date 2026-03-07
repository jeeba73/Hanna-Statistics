Attribute VB_Name = "CLP_Main_Recipe"
Option Explicit



Public Sub AddRecipeinGrid(ByVal Grd As Grid)
Dim i As Integer
Dim sString As String
Dim bMancaFormulation As Boolean
Dim t As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
      
        
        With dbTabRecipe
            .filter = ""
           
         
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
                    Grd.Cell(Grd.Rows - 1, 0).Text = i
                    Grd.Cell(Grd.Rows - 1, 1).Text = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
                    
                     ' bMancaFormulation = GetRecipeFormulation(IIf(IsNull(Trim(!Code)), "", Trim(!Code)))
            
            
                    Grd.Cell(Grd.Rows - 1, 2).Text = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
                    Grd.Cell(Grd.Rows - 1, 3).Text = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
                    Grd.Cell(Grd.Rows - 1, 4).Text = IIf(IsNull(Trim(!Mix)), "", Trim(!Mix))
                    Grd.Cell(Grd.Rows - 1, 5).Text = !Id
                    
                     If IsNull(Trim(!Classification)) Or Trim(!Classification) = "" Then
                                
            
                                Grd.Cell(Grd.Rows - 1, 6).BackColor = vbColorOrange
                         

                       End If
                       
                         
                 
                    
                    .MoveNext
                Next
           End If
        
        
        End With
        
        .Column(2).AutoFit
        .Column(5).Width = 0
        .Column(6).Width = 10
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
    

End Sub


