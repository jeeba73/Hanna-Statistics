Attribute VB_Name = "mod_Excel_ExportProduction"
Option Explicit

Private uProduction As RecipeForProduction



Public Function EsportaProductionExcel(ByVal FileName As String, ByVal sString As String, ByRef iProduction As RecipeForProduction) As Boolean
    Dim rc As Boolean
    On Error GoTo ERR_IMP
    rc = True
   ' MsgBox USER_DESKTOP & "\" & "Backup VerPeriodica.xls"
   uProduction = iProduction
   SettingName = FileName
   
   If SettingName = "" Then Exit Function
   
    
        If CreateExcel(False) Then
            NewExcelWorksheet (sString)
            If CopyChemicalProductionData(SettingName) Then
                Call SaveExcel(sString)
                Call CloseExcel
                
                With dbTabProduction
                    .filter = ""
                    .filter = "FileName='" & SettingName & "'"
                    If .EOF Then
                    Else
                        !ExcelDone = True
                        .Update
                    End If
                End With
                
                PopupMessage 2, "Excel file correctly generated..."
            Else
                rc = False
            End If
        Else
            rc = False
        End If
ERR_END:
    On Error GoTo 0
    EsportaProductionExcel = rc
    Exit Function
ERR_IMP:
    rc = False
    MsgBox err.Description
    Resume ERR_END
End Function


Public Function CopyChemicalProductionData(ByVal SettingName As String) As Boolean
Dim rc As Boolean
Dim i As Integer
    On Error GoTo ERR_COPY
    '---------------------------
    ' set excel page
    '---------------------------
   ' Call SetUnit
    Call FormatPage
    
    Call SetInformation(i)
    Call SetHannaCode(i)
    Call SetAcquisitionGrid(i)
    Call SetNotesGrid(i)
    
    rc = True
ERR_END:
    On Error GoTo 0
    CopyChemicalProductionData = rc
    Exit Function
ERR_COPY:
    rc = False
    MsgBox err.Description
    Resume Next
End Function
Private Sub SetRecipeDetails(ByRef Riga As Integer, ByVal Code As String)

With dbTabRecipe
    
    .filter = ""
    .filter = "Code='" & Code & "'"
    If .EOF Then
    Else
    
        Call AddValue(Riga, 4, "Line", True)
        Call AddValue(Riga, 5, IIf(IsNull(Trim(!Line)), "", Trim(!Line)))
        Call AddValue(Riga, 6, "Procedure", True)
        Call AddValue(Riga, 7, IIf(IsNull(Trim(!Procedure)), "", Trim(!Procedure)))
        Call AddValue(Riga, 8, "Rev", True)
        Call AddValue(Riga, 9, IIf(IsNull(Trim(!Rev)), "", Trim(!Rev)))
        Call AddValue(Riga, 10, "NoteRev", True)
        Call AddValue(Riga, 11, IIf(IsNull(Trim(!NoteRev)), "", Trim(!NoteRev)))
        Call AddValue(Riga, 10, "Exp", True)
        Call AddValue(Riga, 11, IIf(IsNull(Trim(!Exp)), "", Trim(!Exp)))
        Call AddValue(Riga + 1, 4, "Density", True)
        Call AddValue(Riga + 1, 5, IIf(IsNull(Trim(!Density)), "", Trim(!Density)))
        Call AddValue(Riga + 1, 6, "MaxQty", True)
        Call AddValue(Riga + 1, 7, IIf(IsNull(Trim(!MaxQty)), "", Trim(!MaxQty) & !UmMax))
        Call AddValue(Riga + 1, 8, "MinQty", True)
        Call AddValue(Riga + 1, 9, IIf(IsNull(Trim(!MinQty)), "", Trim(!MinQty) & !UmMax))
        Call AddValue(Riga + 1, 10, "Multiple", True)
        Call AddValue(Riga + 1, 11, IIf(IsNull(Trim(!MinQty)), "", Trim(!MinQty)))
        Call AddValue(Riga + 1, 12, "Mix", True)
        Call AddValue(Riga + 1, 13, IIf(IsNull(Trim(!Mix)), "", Trim(!Mix)))
        
        
    End If

End With
End Sub
Private Sub SetInformation(ByRef Riga As Integer)
Dim i As Integer
Dim sString As String
Dim rc As Boolean
Riga = 2

With uProduction

    Call AddValue(Riga + 2, 2, "Production", True, True)
    'Call AddValue(Riga, 2, "Recipe", True)
    'Call AddValue(Riga, 3, .Recipes(1).Code)
    
    
    'Call SetRecipeDetails(Riga, .Recipes(1).Code)
    
    
    'Call AddValue(Riga + 1, 2, "Description", True)
    'Call AddValue(Riga + 1, 3, .Recipes(1).Description)
    
   ' Call AddValue(Riga + 3, 2, "Recipe for Production/Preparation Details", True)
    
    
  
    Call AddValue(Riga + 5, 2, "Recipe by", True)
    Call AddValue(Riga + 5, 3, .RecipeBy)
    Call AddValue(Riga + 5, 4, "Preparation Date", True)
    Call AddValue(Riga + 5, 5, FormatDataLAT(.PreparationDate))
    Call AddValue(Riga + 5, 6, "# Preparation Week", True)
    Call AddValue(Riga + 5, 7, .numPrepWeek)
    
    Call AddValue(Riga + 6, 2, "Planned Preparation Week", True)
    Call AddValue(Riga + 6, 3, "'" & .PlannedPrepWeek)
    Call AddValue(Riga + 6, 4, "Preparation Week", True)
    Call AddValue(Riga + 6, 5, "'" & .PrepWeek)
    Call AddValue(Riga + 6, 6, "Planning Reference", True)
    Call AddValue(Riga + 6, 7, .PlanningReference)
    
    Call AddValue(Riga + 7, 2, "Note", True)
    Call AddValue(Riga + 7, 3, .Note)
    Call AddValue(Riga + 7, 4, "Operator Preparation", True)
    Call AddValue(Riga + 7, 5, .OperatorPrep)
  
End With

    Riga = Riga + 10
 
End Sub
Private Sub SetHannaCode(ByRef Riga As Integer)
Dim i As Integer
Dim t As Integer
Dim HannaCount As Integer
Dim Variance As String
Dim VarDbl As Double
Dim PercStr As String


    Riga = Riga + 2
    
    Call AddValue(Riga - 1, 2, "Hanna Code Table", True, True)



On Error GoTo ERR_GET:
        '------------------------------------------------
        '      PRODUCTION  TABELLA HANNA CODE
        '------------------------------------------------
        '.Cell(0, 1).Text = "Code"
        '.Cell(0, 2).Text = "Product Name"
        '.Cell(0, 3).Text = "Line"
        '.Cell(0, 4).Text = "Volume/Weight"
        '.Cell(0, 5).Text = "(um)"
        '.Cell(0, 6).Text = "Q.ty to produce"
        '.Cell(0, 7).Text = "Q.ty to produced"
        '.Cell(0, 8).Text = ""
        
        '.Cell(0, 9).Text = "%"
        '.Cell(0, 10).Text = "Recipe"
        '.Cell(0, 11).Text = "Mix"
        
        Call AddValue(Riga, 2, "Code", True)
        Call AddValue(Riga, 3, "Product Name", True)
        Call AddValue(Riga, 4, "Line", True)
        Call AddValue(Riga, 5, "Volume/Weight", True)
        Call AddValue(Riga, 6, "(um)", True)
        Call AddValue(Riga, 7, "Q.ty to produce", True)
        Call AddValue(Riga, 8, "Q.ty  produced", True)
        Call AddValue(Riga, 9, "%", True)
        Call AddValue(Riga, 10, "Recipe", True)
        Call AddValue(Riga, 11, "Mix", True)
       ' Call AddValue(Riga, 12, "Exp Date", True)

        
        With uProduction
            HannaCount = .HannaCodesCount
            
            For i = 1 To HannaCount
            
            
                If CDbl(.HannaCodes(i).QtyToProduce) = 0 And CDbl(.HannaCodes(i).QtyProduced) = 0 Then GoTo cont
                If .HannaCodes(i).bHide Then GoTo cont
                
                Riga = Riga + 1
                
                
                Call AddValue(Riga, 2, .HannaCodes(i).Code)
                Call AddValue(Riga, 3, .HannaCodes(i).ProductName)
                Call AddValue(Riga, 4, .HannaCodes(i).Line)
                Call AddValue(Riga, 5, Replace(.HannaCodes(i).Qty, ",", "."))
                Call AddValue(Riga, 6, .HannaCodes(i).Um)
                Call AddValue(Riga, 7, Replace(.HannaCodes(i).QtyToProduce, ",", "."))
                Call AddValue(Riga, 8, Replace(.HannaCodes(i).QtyProduced, ",", "."))
                

                If .HannaCodes(i).QtyToProduce = "" Then .HannaCodes(i).QtyToProduce = "0"
                If .HannaCodes(i).QtyProduced = "" Then .HannaCodes(i).QtyProduced = "0"
                
                If CDbl(.HannaCodes(i).QtyProduced) > 0 And CDbl(.HannaCodes(i).QtyToProduce) > 0 Then
                
                    VarDbl = FormatNumber((.HannaCodes(i).QtyProduced / .HannaCodes(i).QtyToProduce), 4) * 100
                     
                    Select Case VarDbl
                        Case Is < 100
                            PercStr = "'- "
                            VarDbl = FormatNumber(100 - VarDbl, 2)
                        Case Is = 100
                            PercStr = ""
                            VarDbl = VarDbl
                        Case Is > 100
                            PercStr = "'+ "
                            VarDbl = VarDbl
                    End Select
                                       
                    Variance = PercStr & VarDbl & " %"

                    Call AddValue(Riga, 9, Replace(Variance, ",", "."))
                    
                    VarDbl = CDbl(.HannaCodes(i).QtyProduced) - CDbl(.HannaCodes(i).QtyToProduce)
                Else
                    Call AddValue(Riga, 9, "/")
                    
                End If
                
                If CDbl(.HannaCodes(i).QtyToProduce) = 0 Then
                    VarDbl = CDbl(.HannaCodes(i).QtyProduced)
                End If
                Call AddValue(Riga, 10, .HannaCodes(i).Recipe)
                Call AddValue(Riga, 11, .HannaCodes(i).Mix1 & IIf(Len(.HannaCodes(i).Mix2) > 0, ";" & .HannaCodes(i).Mix2, ""))
                
               ' If .HannaCodes(i).ExpDate = "" Then .HannaCodes(i).ExpDate = SetExpDate(.PreparationDate, GetRecipeExp(.HannaCodes(i).Recipe))
                
               ' Call AddValue(Riga, 12, "'" & .HannaCodes(i).ExpDate)
                
                
                
              
              
              
                'If CDbl(.HannaCodes(i).QtyProduced) > 0 Then
                '    For t = 1 To Grid.Cols - 1

                '       Grid.Cell(Grid.Rows - 1, t).FontBold = True
                '       Grid.Cell(Grid.Rows - 1, t).ForeColor = &H404040 '&H644603
                
                        
                '    Next
                    
                '    Select Case VarDbl
                    
    
                '        Case -.HannaCodes(i).QtyToProduce * 0.2 To -.HannaCodes(i).QtyToProduce * 0.02
                '            Grid.Cell(Grid.Rows - 1, 8).BackColor = vbColorOrange
                '        Case Is < -.HannaCodes(i).QtyToProduce * 0.2
                '            Grid.Cell(Grid.Rows - 1, 8).BackColor = &HC0&
                '        Case Else
                '            Grid.Cell(Grid.Rows - 1, 8).BackColor = vbColorGreen
                '    End Select
                
                ' End If
                
cont:
            Next
        
        
        End With
ERR_END:

    Riga = Riga + 2
    
    Exit Sub
ERR_GET:
   MsgBox err.Description
   Resume Next

End Sub


Private Sub SetAcquisitionGrid(ByRef Riga As Integer)
Dim i As Integer
Dim t As Integer
Dim RecipeCount As Integer
Dim Variance As Double
Dim VariancePerc As Double
Dim TotalRealWeight As Double
Dim bUmMassa As Boolean
Dim Density As Double
Dim bRecalculate As Boolean
Dim PesoIntolleranza As Double
Dim MyColor As OLE_COLOR

Dim iAcquisition As PrepAcquisition
 
On Error GoTo ERR_GET:


    Riga = Riga + 2
    
    Call AddValue(Riga - 1, 2, "Acquisition Table", True, True)
    
    '------------------------------------------------
    '      Acquisition Grid
    '------------------------------------------------
    
    Call AddValue(Riga, 2, "Code", True)
    Call AddValue(Riga, 3, "QtyProduced", True)
    Call AddValue(Riga, 4, "LotNumber", True)
    Call AddValue(Riga, 5, "Operator", True)
    Call AddValue(Riga, 6, "DateProd", True)
    Call AddValue(Riga, 7, "WeekProd", True)
    Call AddValue(Riga, 8, "Machine", True)
    Call AddValue(Riga, 9, "Note", True)
    Call AddValue(Riga, 10, "AcquisitionTime", True)
    Call AddValue(Riga, 11, uProduction.HannaCodes(i).Mix1 & " Lot", True)
    Call AddValue(Riga, 12, uProduction.HannaCodes(i).Mix2 & " Lot", True)
    Call AddValue(Riga, 13, "Exp Date", True)
    
   ' uProduction.HannaCodes(i).Mix1
   ' uProduction.HannaCodes(i).Mix2
   
     For i = 1 To UBound(uProduction.HannaCodes)

        With uProduction.HannaCodes(i)
          
          
          
            If .AcquisitionCount > 0 Then
                For t = 1 To .AcquisitionCount
                    Call ProductionAddNewRowInAcquisitionExcel(Riga, .Acquisitions(t))
                Next
            End If
        End With
    Next
    



ERR_END:
   Riga = Riga + 2
   On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox err.Description
    Resume Next
End Sub

Private Function ProductionAddNewRowInAcquisitionExcel(ByRef Riga As Integer, ByRef iAcquisition As ProdAcquisition)
                
                
               
                
                    Riga = Riga + 1
                    
                   
                    
                    
                    Call AddValue(Riga, 2, iAcquisition.Code)
                    Call AddValue(Riga, 3, CStr(Replace(iAcquisition.QtyProduced, ",", ".")))
                    Call AddValue(Riga, 4, "'" & CStr(iAcquisition.LotNumber))
                    Call AddValue(Riga, 5, iAcquisition.Operator)
                    Call AddValue(Riga, 6, FormatDataLAT(iAcquisition.DateProd))
                    Call AddValue(Riga, 7, "'" & iAcquisition.WeekProd)
                    Call AddValue(Riga, 8, iAcquisition.Machine)
                    Call AddValue(Riga, 9, iAcquisition.Note)
                    Call AddValue(Riga, 10, CStr(iAcquisition.AcquisitionTime))
                    Call AddValue(Riga, 11, "'" & CStr(iAcquisition.Mix1Lot))
                    Call AddValue(Riga, 12, "'" & CStr(iAcquisition.Mix2Lot))
                    Call AddValue(Riga, 13, "'" & iAcquisition.ExpDate)
                    
                    
                    
                    
End Function
Private Sub SetNotesGrid(ByRef Riga As Integer)
Dim i As Integer
Dim HannaCount As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    
On Error GoTo ERR_GET:
    
    Riga = Riga + 2
    
    Call AddValue(Riga - 1, 2, "Production Notes", True, True)
    

        

            
                    Call AddValue(Riga, 2, "Date", True)
                    Call AddValue(Riga, 3, "Type", True)
                    Call AddValue(Riga, 4, "Description", True)
                    Call AddValue(Riga, 5, "Operator", True)

        With dbTabProductionNotes
            .filter = ""
            .filter = "FileName='" & SettingName & "'"
            If .EOF Then
            Else
                .MoveFirst
                For i = 1 To .RecordCount
                
                    
                   
                        Riga = Riga + 1
    
                        Call AddValue(Riga, 2, IIf(IsNull(Trim(!NoteDate)), "", Trim(!NoteDate)))
                        Call AddValue(Riga, 3, IIf(IsNull(Trim(!Type)), "", Trim(!Type)))
                        Call AddValue(Riga, 4, IIf(IsNull(Trim(!Description)), "", Trim(!Description)))
                        Call AddValue(Riga, 5, IIf(IsNull(Trim(!Operator)), "", Trim(!Operator)))
                      
                     
                     .MoveNext
                    
                Next
            End If
        
        
        End With

ERR_END:
   Riga = Riga + 2
   On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox err.Description
    Resume Next

End Sub
