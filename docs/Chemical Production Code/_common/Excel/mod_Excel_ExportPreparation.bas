Attribute VB_Name = "mod_Excel_ExportPreparation"
Option Explicit


Private uPreparation As RecipeForProduction



Public Sub ExportPreparationAfterQC(ByVal SettingName As String, ByVal RecipeCode As String, ByRef Lot As String, ByVal rfpSettingName As String)
    Dim i As Integer
    Dim RFPFILE_PATH As String
    Dim ExcelFilename As String
    ' export LOT Excel
    If SettingName = "" Then
    Else
    
        If FileExists(USER_TEMP_PATH & rfpSettingName) Then
            RFPFILE_PATH = USER_TEMP_PATH
        ElseIf FileExists(USER_DATA_PATH & rfpSettingName) Then
            RFPFILE_PATH = USER_DATA_PATH
        ElseIf FileExists(USER_PRODUCTION_PATH & rfpSettingName) Then
            RFPFILE_PATH = USER_PRODUCTION_PATH
        End If
            
    
        If FileExists(USER_PREPARATION_PATH & SettingName) Then
            USER_PATH = USER_PREPARATION_PATH
        ElseIf FileExists(USER_PREPARATION_PATH & "Data\" & SettingName) Then
            USER_PATH = USER_PREPARATION_PATH & "Data\"
        Else
         
            PopupMessage 2, "No file Preparation found...", , True, SettingName
            Exit Sub
        End If
        
        Call PreparationGetSetting(uPreparation, SettingName, RecipeCode)
        
            If uPreparation.Recipes(1).PreparationLotMix = "" Or uPreparation.Recipes(1).PreparationLotMix = "0000" Then
             
check:
             
             Call CheckPreparationLot(uPreparation.Recipes(1).PreparationLotMix, uPreparation.Recipes(1).Line, True, uPreparation)
             
             Lot = uPreparation.Recipes(1).PreparationLotMix
             
             If F_MsgBox.DoShow("Use this Lot noumber?", "Lot Number = " & uPreparation.Recipes(1).PreparationLotMix) Then
             
             Else
                Lot = uPreparation.Recipes(1).PreparationLotMix
InputLot:
                    If F_InputBox.DoShow("Please enter Lot Number.", "Preparation", , , , Lot) Then
                        
                        If Len(Lot) <> 4 Then
                              
                                PopupMessage 2, "Lot must be 4 digits : es. 0001", , True, "LOT NUMBERT ERROR"
                                GoTo check
                            
                            End If
                        
                            If CheckPreparationLot(Lot, uPreparation.Recipes(1).Line, False, uPreparation) Then

                                        uPreparation.Recipes(1).PreparationLotMix = Lot
                                    
                                  
                            Else
                                PopupMessage 2, "Lot Number already exsists"
                                If F_MsgBox.DoShow("Lot Number already exsists" & vbCrLf & "Use this Lot noumber?", "Lot Number = " & uPreparation.Recipes(1).PreparationLotMix) Then
             
                                Else
                                    GoTo InputLot:
                                
                                End If
                            End If
    
                        
                        
                    End If
                    
             End If
             
             
             If uPreparation.Recipes(1).bIsMix Then
             
             Else
                CloseSettingDataFile
                
                For i = 1 To UBound(uPreparation.HannaCodes)
                   uPreparation.HannaCodes(i).LotNumber = uPreparation.Recipes(1).PreparationLotMix
                   SaveSettingData SettingName, "HannaCode" & i, "LotNumber", uPreparation.HannaCodes(i).LotNumber
                Next
                
                CloseSettingDataFile
                
                For i = 1 To UBound(uPreparation.HannaCodes)
                 
                   SaveSettingData rfpSettingName, "HannaCode" & i, "LotNumber", uPreparation.HannaCodes(i).LotNumber, RFPFILE_PATH
                   
                   
                Next
                
                End If
           
        End If
        
        CloseSettingDataFile
        
        ExcelFilename = "PREP_" & FormatNomeFile(Trim(uPreparation.Recipes(1).Code) & "." & Trim(uPreparation.Recipes(1).Line) & "." & uPreparation.numPrepWeek & "." & uPreparation.PrepWeek & "." & Lot)
        
        
        
        PopupMessage 2, "Exporting data to Excel : please wait...." & vbCrLf & ExcelFilename
        If Len(ExcelFilename) > 40 Then ExcelFilename = Left$(ExcelFilename, 40)
        Call EsportaPreparationExcel(SettingName, ExcelFilename, uPreparation)
    End If


End Sub



Public Function EsportaPreparationExcel(ByVal FileName As String, ByVal sString As String, ByRef iPreparation As RecipeForProduction) As Boolean
    Dim rc As Boolean
    On Error GoTo ERR_IMP
    rc = True
   ' MsgBox USER_DESKTOP & "\" & "Backup VerPeriodica.xls"
   uPreparation = iPreparation
   SettingName = FileName
   sString = FormatNomeFile(sString)
   If SettingName = "" Then Exit Function
   
    
        If CreateExcel(False) Then
            NewExcelWorksheet (sString)
            If CopyChemicalPreparationData(SettingName) Then
                Call SaveExcel(sString)
                Call CloseExcel
                
                With dbTabPreparation
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
    EsportaPreparationExcel = rc
    Exit Function
ERR_IMP:
    rc = False
    MsgBox err.Description
    Resume ERR_END
End Function


Public Function CopyChemicalPreparationData(ByVal SettingName As String) As Boolean
Dim rc As Boolean
Dim i As Integer
    On Error GoTo ERR_COPY
    '---------------------------
    ' set excel page
    '---------------------------
   ' Call SetUnit
    Call FormatPage
    
    Call SetInformation(i)
    Call SetComponent(i)
    Call SetAcquisitionGrid(i)
    
    If uPreparation.Recipes(1).bIsMix Then
    
    Else
      Call SetHannaCodeGrid(i)
    End If
   
    Call SetNotesGrid(i)
    
    rc = True
ERR_END:
    On Error GoTo 0
    CopyChemicalPreparationData = rc
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
        Call AddValue(Riga, 11, IIf(IsNull(Trim(!Exp)), "", "'" & Trim(!Exp)))
       

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
Riga = 4

With uPreparation

    Call AddValue(Riga - 2, 2, "Preparation", True, True)
    Call AddValue(Riga, 2, "Recipe", True)
    Call AddValue(Riga, 3, .Recipes(1).Code)
    
    
    Call SetRecipeDetails(Riga, .Recipes(1).Code)
    
    
    Call AddValue(Riga + 1, 2, "Description", True)
    Call AddValue(Riga + 1, 3, .Recipes(1).Description)
    
   ' Call AddValue(Riga + 3, 2, "Recipe for production/Preparation Details", True)
    
    Call AddValue(Riga + 5, 2, "Recipe by", True)
    Call AddValue(Riga + 5, 3, .RecipeBy)
    Call AddValue(Riga + 5, 4, "Preparation Date", True)
    Call AddValue(Riga + 5, 5, .PreparationDate)
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
    Call AddValue(Riga + 7, 4, "Operator", True)
    Call AddValue(Riga + 7, 5, .OperatorPrep)
    Call AddValue(Riga + 7, 6, "Exp Date", True)
    Call AddValue(Riga + 7, 7, "'" & .ExpDate)
    
     
    If .Recipes(1).bIsMix Then
    
        Call AddValue(Riga + 8, 2, "Preparation Lot (Mix)", True)
        Call AddValue(Riga + 8, 3, .Recipes(1).PreparationLotMix)
    
    End If
        
    
End With

    Riga = Riga + 12
 
End Sub
Private Sub SetComponent(ByRef Riga As Integer)
Dim i As Integer
Dim t As Integer
Dim x As Integer
Dim RecipeCount As Integer
Dim Variance As Double
Dim VariancePerc As Double
Dim TotalRealWeight As Double
Dim TheorWeight As Double
Dim bUmMassa As Boolean
Dim Density As Double
Dim bRecalculate As Boolean
Dim PesoIntolleranza As Double
Dim MyColor As OLE_COLOR
Dim InputWeight As String
Dim bCorrection As Boolean
Dim Perc As String
Dim TheorW As String
    Riga = Riga + 2
    
    Call AddValue(Riga - 1, 2, "Component Table", True, True)



On Error GoTo ERR_GET:
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
        
     
        
        CloseSettingDataFile

        With uPreparation.Recipes(1)
            bUmMassa = .bUmMassa
            Density = .Density
            
            Call AddValue(Riga, 2, "Code", True)
            Call AddValue(Riga, 3, "Description", True)
            Call AddValue(Riga, 4, "Cas", True)
            Call AddValue(Riga, 5, "%", True)
            Call AddValue(Riga, 6, "Theoretical weight", True)
            Call AddValue(Riga, 7, "Real weight", True)
            
            Call AddValue(Riga, 8, "Variance", True)
            Call AddValue(Riga, 9, "Variance Perc", True)
            
            Call AddValue(Riga, 10, "Real Perc", True)
            Call AddValue(Riga, 11, "Note", True)
            Call AddValue(Riga, 12, "Mix", True)
            
            For i = 0 To .RmxRecipeCount
         
                     
                'If uPreparation.Recipes(1).code <> .RmxRecipe(i).RecipeCode Then GoTo cont:
                'If .RmxRecipe(i).bDeleted Then GoTo cont:

                Riga = Riga + 1
                
                If .RmxRecipe(i).bAddedInPreparation Or .RmxRecipe(i).bCorrection Then
                    Call AddValue(Riga, 2, .RmxRecipe(i).CHCode, , , , True)
                    Call AddValue(Riga, 3, .RmxRecipe(i).Description, , , , True)
                    Call AddValue(Riga, 4, .RmxRecipe(i).Cas, , , , True)
                
                Else
                    Call AddValue(Riga, 2, .RmxRecipe(i).CHCode)
                    Call AddValue(Riga, 3, .RmxRecipe(i).Description)
                    Call AddValue(Riga, 4, .RmxRecipe(i).Cas)
                End If
                
                
                Perc = Replace(GetSettingData(SettingName, "Recipes" & 1 & " - RmxRecipe" & i, "Perc", ""), ",", ".")
                TheorW = Replace(PadString(GetSettingData(SettingName, "Recipes" & 1 & " - RmxRecipe" & i, "TheoreticalWeight", "")), ",", ".")
              
                
                If .RmxRecipe(i).bAddedInPreparation Then
                    If .RmxRecipe(i).TheoreticalWeight <> .RmxRecipe(i).RealWeight Then
                        Call AddValue(Riga, 6, TheorW)
                    Else
                        Call AddValue(Riga, 6, "")
                    End If
                    Call AddValue(Riga, 6, "")
                    
                Else
                    Call AddValue(Riga, 5, (Perc) & " %")
                    Call AddValue(Riga, 6, (TheorW))
                End If
                
                
                Call AddValue(Riga, 7, Replace(PadString(.RmxRecipe(i).RealWeight), ",", "."))
              
                If .RmxRecipe(i).bAddedInPreparation Then
                    
                    If .RmxRecipe(i).TheoreticalWeight <> .RmxRecipe(i).RealWeight Then
                        GoTo PrintVariance
                    Else
                        Call AddValue(Riga, 8, Replace(PadString(.RmxRecipe(i).RealWeight), ",", "."))
                        Call AddValue(Riga, 9, "-")
                    End If
                    
                Else
PrintVariance:
                    Call AddValue(Riga, 8, Replace(PadString(.RmxRecipe(i).Variance), ",", "."))
                    Call AddValue(Riga, 9, Replace(FormatNumber(.RmxRecipe(i).VariancePerc, 2), ",", ".") & "%")
                   
                End If
                
                
                Call AddValue(Riga, 10, Replace(.RmxRecipe(i).RealPerc, ",", ".") & " %")
                Call AddValue(Riga, 11, .RmxRecipe(i).Note)
                Call AddValue(Riga, 12, IIf(.RmxRecipe(i).bMix, "X", ""))
                
            Next
            
            Riga = Riga + 2
        
            Call AddValue(Riga, 2, "TotalWeight (Kg)", True)
            Call AddValue(Riga, 3, Replace(PadString(.TotalWeightKg), ",", "."))
            
            Call AddValue(Riga, 4, "Real Weight (Kg)", True)
            Call AddValue(Riga, 5, Replace(PadString(.ActualWeight), ",", "."))
            
            Variance = .ActualWeight - .TotalWeightKg
            VariancePerc = (Variance / .TotalWeightKg) * 100
            
            Call AddValue(Riga, 6, "Variance (Kg)", True)
            Call AddValue(Riga, 7, Replace(PadString(Variance), ",", "."))
            Call AddValue(Riga, 8, "Variance Perc", True)
            Call AddValue(Riga, 9, Replace(FormatNumber(VariancePerc, 2), ",", ".") & "%")
           

            
            If Not (bUmMassa) Then
            
                ' tot Litri
                
                Riga = Riga + 1
                
                Call AddValue(Riga, 2, "TotalWeight (L)", True)
                Call AddValue(Riga, 3, Replace(PadString(.TotalWeightKg / Density), ",", "."))
                
                Call AddValue(Riga, 4, "Real Weight (L)", True)
                Call AddValue(Riga, 5, Replace(PadString(.ActualWeight / Density), ",", "."))
                
                Variance = Variance / Density
               ' VariancePerc = (Variance / .TotalWeightKg) * 100
                
                Call AddValue(Riga, 6, "Variance (L)", True)
                Call AddValue(Riga, 7, Replace(PadString(Variance), ",", "."))
                Call AddValue(Riga, 8, "Variance Perc", True)
                Call AddValue(Riga, 9, Replace(FormatNumber(VariancePerc, 2), ",", ".") & "%")
            

            
            End If
    End With


    Riga = Riga + 3
 
ERR_END:
   CloseSettingDataFile
   
   On Error GoTo 0
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
    With uPreparation.Recipes(1)
            bUmMassa = .bUmMassa
            Density = .Density
            If .AcquisitionCount > 0 Then
 
            Call AddValue(Riga, 2, "Code", True)
            Call AddValue(Riga, 3, "Description", True)
            Call AddValue(Riga, 4, "Cas", True)
            Call AddValue(Riga, 5, "Real Weight (g)", True)
            Call AddValue(Riga, 6, "Manufacturer", True)
            Call AddValue(Riga, 7, "Manufacturer Code", True)
            Call AddValue(Riga, 8, "Manufacturer Lot", True)
            Call AddValue(Riga, 9, "Delivery Date", True)
            Call AddValue(Riga, 10, "Qty Delivered", True)
            Call AddValue(Riga, 11, "Week Delivery", True)
            Call AddValue(Riga, 12, "Package", True)
            Call AddValue(Riga, 13, "Note", True)
            Call AddValue(Riga, 14, "Operator", True)
            Call AddValue(Riga, 15, "Acquisition Time", True)
            Call AddValue(Riga, 16, "Recalculation", True)
            Call AddValue(Riga, 17, "Added Chemical (in recipe)", True)
                
                For i = 1 To .AcquisitionCount
                
                    Riga = Riga + 1
                    
                    iAcquisition = uPreparation.Recipes(1).Acquisitions(i)
                    
                    
                    Call AddValue(Riga, 2, iAcquisition.PrepBarcode.Code)
                    Call AddValue(Riga, 3, iAcquisition.PrepBarcode.ChemicalName)
                    Call AddValue(Riga, 4, iAcquisition.PrepBarcode.Cas)
                    Call AddValue(Riga, 5, Replace(PadString(iAcquisition.ActualWeight), ",", "."))
                    Call AddValue(Riga, 6, iAcquisition.PrepBarcode.Manufacturer)
                    Call AddValue(Riga, 7, iAcquisition.PrepBarcode.ManufacturerCode)
                    Call AddValue(Riga, 8, "'" & CStr(iAcquisition.PrepBarcode.ManufacturerLot))
                    Call AddValue(Riga, 9, iAcquisition.PrepBarcode.DeliveryDate)
                    Call AddValue(Riga, 10, iAcquisition.PrepBarcode.QtyDelivered)
                    Call AddValue(Riga, 11, iAcquisition.PrepBarcode.WeekDelPackageNumber)
                    Call AddValue(Riga, 12, "'" & CStr(iAcquisition.PrepBarcode.Package))
                    Call AddValue(Riga, 13, iAcquisition.Note)
                    Call AddValue(Riga, 14, iAcquisition.Operator)
                    Call AddValue(Riga, 15, CStr(iAcquisition.AcquisitionTime))
                    

                       
                       If iAcquisition.bRecalculation Then
                           Call AddValue(Riga, 16, "X")
                        Else
                            Call AddValue(Riga, 16, "")
                       End If
                       If iAcquisition.bRecipeComponent = False Then
                          Call AddValue(Riga, 17, "X")
                        Else
                            Call AddValue(Riga, 17, "")
                       End If
                       
                    
                    
                Next
            End If
            
            
    End With

ERR_END:
   Riga = Riga + 3
   On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox err.Description
    Resume Next
End Sub

Private Sub SetHannaCodeGrid(ByRef Riga As Integer)
Dim i As Integer
Dim HannaCount As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    
On Error GoTo ERR_GET:
    
    Riga = Riga + 2
    
    Call AddValue(Riga - 1, 2, "Hanna Code Table", True, True)
    

        
        With uPreparation
            HannaCount = .HannaCodesCount
            
                    Call AddValue(Riga, 2, "Code", True)
                    Call AddValue(Riga, 3, "Product Name", True)
                    Call AddValue(Riga, 4, "Line", True)
                    Call AddValue(Riga, 5, "Volume/Weight", True)
                    Call AddValue(Riga, 6, "(um)", True)
                    Call AddValue(Riga, 7, "Q.ty to produce", True)
                    Call AddValue(Riga, 8, "Lot Number", True)
                    
            
            For i = 1 To HannaCount
            
                If (.HannaCodes(i).bHide) Then
                Else
               
                    Riga = Riga + 1

                    Call AddValue(Riga, 2, "'" & .HannaCodes(i).Code)
                    Call AddValue(Riga, 3, "'" & .HannaCodes(i).ProductName)
                    Call AddValue(Riga, 4, "'" & .HannaCodes(i).Line)
                    Call AddValue(Riga, 5, "'" & Replace(.HannaCodes(i).Qty, ",", "."))
                    Call AddValue(Riga, 6, "'" & .HannaCodes(i).Um)
                    Call AddValue(Riga, 7, "'" & .HannaCodes(i).QtyToProduce)
                    Call AddValue(Riga, 8, "'" & CStr(.HannaCodes(i).LotNumber))
                 End If
                
            Next
        
        
        End With

ERR_END:
   Riga = Riga + 2
   On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox err.Description
    Resume Next

End Sub
Private Sub SetNotesGrid(ByRef Riga As Integer)
Dim i As Integer
Dim HannaCount As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    
On Error GoTo ERR_GET:
    
    Riga = Riga + 2
    
    Call AddValue(Riga - 1, 2, "Preparation Notes", True, True)
    

        

            
                    Call AddValue(Riga, 2, "Date", True)
                    Call AddValue(Riga, 3, "Type", True)
                    Call AddValue(Riga, 4, "Description", True)
                    Call AddValue(Riga, 5, "Operator", True)

        With dbTabPreparationNotes
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
