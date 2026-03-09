Attribute VB_Name = "mod_Excel_ExportPreparation"
Option Explicit


Private uPreparation As RecipeForProduction
Private MsType As Integer
Private strMS As String
Private bManualPrepration As Boolean

Public Function EsportaPreparationExcel(ByVal FileName As String, ByVal sString As String, ByRef iPreparation As RecipeForProduction, Optional ByVal bManual As Boolean) As Boolean
    Dim rc As Boolean
    On Error GoTo ERR_IMP
    rc = True
   ' MsgBox USER_DESKTOP & "\" & "Backup VerPeriodica.xls"
   uPreparation = iPreparation
   SettingName = FileName
   sString = FormatNomeFile(sString)
   
   bManualPrepration = bManual
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
    MsgBox Err.Description
    Resume Next
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
    
    If bManualPrepration Then
    Else
        Call SetHannaCode(i)
        Call SetMR(i)
    
    End If
    
    
    Call SetComponent(i)
    Call SetAcquisitionGrid(i)
    Call SetMotherSolutionGrid(i)
    
    Call SetNotesGrid(i)
    

    rc = True
ERR_END:
    On Error GoTo 0
    CopyChemicalPreparationData = rc
    Exit Function
ERR_COPY:
    rc = False
    MsgBox Err.Description
    Resume Next
End Function

Private Sub SetRecipeDetails(ByRef Riga As Integer, ByVal Code As String)

End Sub

Private Sub SetInformation(ByRef Riga As Integer)
Dim i As Integer
Dim sString As String
Dim rc As Boolean
Riga = 4

With uPreparation

    Call AddValue(Riga, 2, "Preparation", True, True)
    Call AddValue(Riga + 2, 2, "Hanna Code", True)
    Call AddValue(Riga + 2, 3, .HannaCode.Code)
    
    Call SetRecipeDetails(Riga, .HannaCode.Code)
    
    
    Call AddValue(Riga + 2, 4, "Description", True)
    Call AddValue(Riga + 2, 5, .HannaCode.Description)
    
    If bManualPrepration Then
    Else
    
        Call AddValue(Riga + 3, 2, "MR Code", True)
        Call AddValue(Riga + 3, 3, .HannaCode.MR.Code)
        
        Call AddValue(Riga + 3, 4, "MR Description", True)
        Call AddValue(Riga + 3, 5, .HannaCode.MR.Description)
    
    End If
    
    Call AddValue(Riga + 5, 2, "Preparation Date", True)
    Call AddValue(Riga + 5, 3, "'" & FormatDataLAT((.DataPrep)))
    Call AddValue(Riga + 5, 4, "Preparation Hour", True)
    Call AddValue(Riga + 5, 5, "'" & .HourPrep)
    Call AddValue(Riga + 5, 6, "Preparation Week", True)
    Call AddValue(Riga + 5, 7, "'" & .PrepWeek)
    
    If bManualPrepration Then
    Else

        Call AddValue(Riga + 6, 2, "FW Hanna Parameter", True)
        Call AddValue(Riga + 6, 3, "'" & .HannaCode.MR.FWParameter)
        
        Call AddValue(Riga + 6, 4, "Measurement Unit", True)
        Call AddValue(Riga + 6, 5, "'" & .HannaCode.STDUnit)
        
       
    
    End If
    
    Call AddValue(Riga + 6, 6, "STD Matrix", True)
    Call AddValue(Riga + 6, 7, "'" & .HannaCode.STDMatrix)
    
    Call AddValue(Riga + 6, 8, "STD Exp (days)", True)
    Call AddValue(Riga + 6, 9, "'" & .HannaCode.STDExp)
    
    Call AddValue(Riga + 6, 10, "STD Exp (Date)", True)
    Call AddValue(Riga + 6, 11, "'" & .HannaCode.STDExpDate)
    
    Call AddValue(Riga + 6, 12, "Storage STD", True)
    Call AddValue(Riga + 6, 13, "'" & .HannaCode.STDStorage)
    

    
    Call AddValue(Riga + 7, 2, "Note", True)
    Call AddValue(Riga + 7, 3, .Note)
    Call AddValue(Riga + 7, 4, "Operator", True)
    Call AddValue(Riga + 7, 5, .Operator)
    
    
    Call AddValue(Riga + 10, 2, "STD Volume/Weight", True)
    Call AddValue(Riga + 10, 3, Replace(.HannaCode.STDVolume, ",", "."))
                     
    MsType = .MsType
    
    Select Case .MsType
        
        Case 1, 2
            strMS = "MS"
        Case Else
            strMS = "MR"
    End Select
    
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
    
    Call AddValue(Riga - 1, 2, "STD Table", True, True)



On Error GoTo ERR_GET:
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
        
     
        
        CloseSettingDataFile

        With uPreparation.Recipe

            If bManualPrepration Then
                strMS = "MR"
            
                Call AddValue(Riga, 2, "MR Code", True)
            Else
                Call AddValue(Riga, 2, "STD Number", True)
            End If
            
            Call AddValue(Riga, 3, "STD Value", True)
            Call AddValue(Riga, 4, strMS & " Qty", True)
            Call AddValue(Riga, 5, strMS & " Acquired", True)
            Call AddValue(Riga, 6, "Variance", True)
            Call AddValue(Riga, 7, "Variance Perc", True)
            Call AddValue(Riga, 8, "Note", True)

            For i = 1 To .STDcount
         
                     
                Riga = Riga + 1
                
                    If .STD(i).Note = "" Then .STD(i).Note = "-"
               
                    If bManualPrepration Then
                        Call AddValue(Riga, 2, "'" & (.STD(i).MRCode))
                    Else
                        Call AddValue(Riga, 2, "'" & str(.STD(i).NUMBER))
                    End If
                
                    Call AddValue(Riga, 3, "'" & str(.STD(i).Value))
 
                    Call AddValue(Riga, 4, "'" & str(.STD(i).TheoreticalWeight))
                    Call AddValue(Riga, 5, "'" & str(.STD(i).RealWeight))
 
                    Call AddValue(Riga, 6, "'" & str(-.STD(i).Variance))
                    Call AddValue(Riga, 7, "'" & str(-.STD(i).VariancePerc) & " %")
                    Call AddValue(Riga, 8, "'" & .STD(i).Note)
                
            Next
            
            Riga = Riga + 2
        
            
            
            
    End With


    Riga = Riga + 3
 
ERR_END:
   CloseSettingDataFile
   
   On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox Err.Description
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
    With uPreparation.Recipe
            bUmMassa = .bUmMassa
          
            If .AcquisitionCount > 0 Then
            


   ' With userAcquisition
    
   '     .AcquisitionTime = Now()
   '     .ActualWeight = txAcquisition(13)
   '     .LeftInBottle = uBottle(0).StockQTY & " " & uBottle(0).StockUnit
   '     .Bottle = uBottle(0).EntryBottle ' txStock(1)
   '     .Code = uRecipe.Code
   '     .DatePrep = uPreparation.DataPrep
   '     .FileName = SettingName
   '     .HannaCode = uPreparation.HannaCode.Code
   '     .HourPrep = uPreparation.HourPrep
       ' .MotherSolutionDate=
   '     .MRLot = uBottle(0).Lot ' txAcquisition(0)
   '     .MSType = uPreparation.MSType
   '     .Note = txAcquisition(17)
   '     .Operator = MyOperatore.Name
   '     .PreparationID = PreparationID
   '     .STDNumber = txAcquisition(15)
    
   '     .STDUnit = txAcquisition(19)
   '     .STDQty = txAcquisition(10)
   '     .STDValue = txAcquisition(16)
   '     .WeekPrep = uPreparation.PrepWeek
   '     .CodicePipetta = txAcquisition(9)

   ' End With



            Call AddValue(Riga, 2, "Index", True)
            Call AddValue(Riga, 3, "AcquisitionTime", True)
            Call AddValue(Riga, 4, "Bottle", True)
            Call AddValue(Riga, 5, "Lot", True)
            If bManualPrepration Then
                Call AddValue(Riga, 6, "MR Code", True)
            Else
                Call AddValue(Riga, 6, "STDNumber", True)
            End If
            Call AddValue(Riga, 7, "STDValue", True)
            Call AddValue(Riga, 8, "STDUnit", True)
            Call AddValue(Riga, 9, strMS & " Qty", True)
            
            Call AddValue(Riga, 10, strMS & " Acquired", True)
            Call AddValue(Riga, 11, "Pipette", True)
            Call AddValue(Riga, 12, "LeftInBottle", True)
            Call AddValue(Riga, 13, "Operator", True)
            
            Call AddValue(Riga, 14, "Note", True)
            
            Call AddValue(Riga, 15, "Pipetta Code", True)
            Call AddValue(Riga, 16, "Pipetta Type", True)
            
            Call AddValue(Riga, 17, "Scale Code", True)
            Call AddValue(Riga, 18, "GlassWare Code", True)
            If bManualPrepration Then
            Else
            Call AddValue(Riga, 19, "MS Date", True)
            End If
            
            Call AddValue(Riga, 20, "MNP", True)
            Call AddValue(Riga, 21, "Exp.MR", True)
            
            
        
          
                
                For i = 1 To .AcquisitionCount
                
                    Riga = Riga + 1
                    
                    iAcquisition = uPreparation.Recipe.Acquisitions(i)
                    
                    
                    Call AddValue(Riga, 2, "'" & iAcquisition.Index)
                    Call AddValue(Riga, 3, "'" & iAcquisition.AcquisitionTime)
                    Call AddValue(Riga, 4, "'" & iAcquisition.Bottle)
                    Call AddValue(Riga, 5, "'" & iAcquisition.MRLot)
                    
                    If bManualPrepration Then
                        Call AddValue(Riga, 6, "'" & iAcquisition.Code)
                    Else
                        Call AddValue(Riga, 6, "'" & iAcquisition.STDNumber)
                    End If
                    
                    
                    Call AddValue(Riga, 7, "'" & iAcquisition.STDValue)
                    Call AddValue(Riga, 8, "'" & iAcquisition.STDUnit)
                    Call AddValue(Riga, 9, "'" & iAcquisition.STDQty)
                    
                    Call AddValue(Riga, 10, "'" & iAcquisition.ActualWeight)
                    Call AddValue(Riga, 11, "'" & iAcquisition.CodicePipetta)
                    Call AddValue(Riga, 12, "'" & iAcquisition.LeftInBottle)
                    Call AddValue(Riga, 13, "'" & iAcquisition.Operator)
                    Call AddValue(Riga, 14, "'" & iAcquisition.Note)
                    
                    Call AddValue(Riga, 15, "'" & iAcquisition.CodicePipetta)
                    Call AddValue(Riga, 16, "'" & iAcquisition.PipettaType)
                    
                    Call AddValue(Riga, 17, "'" & iAcquisition.ScaleID)
                    Call AddValue(Riga, 18, "'" & iAcquisition.GlasswareID)
                    If bManualPrepration Then
                    Else
                    Call AddValue(Riga, 19, "'" & iAcquisition.MotherSolutionDate)
                    End If
                    
                    Call AddValue(Riga, 20, "'" & iAcquisition.MNP)
                    Call AddValue(Riga, 21, "'" & iAcquisition.ExpMR)
    
                    
                Next
            End If
            
            
    End With

ERR_END:
   Riga = Riga + 3
   On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox Err.Description
    Resume Next
End Sub

Private Sub SetMotherSolutionGrid(ByRef Riga As Integer)
Dim i As Integer
Dim HannaCount As Integer
        '------------------------------------------------
        '      Mother Solution
        '------------------------------------------------
    
On Error GoTo ERR_GET:
    
    Riga = Riga + 2
    
    
    
   

        
        With uPreparation.MotherSol
           If .MsType = 0 Then Exit Sub
           
           
            Call AddValue(Riga - 1, 2, "Mother Solution", True, True)
    
           
            Call AddValue(Riga, 2, "HannaCode", True)
            Call AddValue(Riga, 3, "MR Qty Required", True)
            Call AddValue(Riga, 4, "DataPrep", True)
            Call AddValue(Riga, 5, "HourPrep", True)
            Call AddValue(Riga, 6, "PrepWeek", True)
            Call AddValue(Riga, 7, "MS Preparation Date", True)
            Call AddValue(Riga, 8, "MS QtyProduced", True)
            Call AddValue(Riga, 9, "MS DataExp", True)
            
            Call AddValue(Riga, 10, "MR Bottle Number", True)
            Call AddValue(Riga, 11, "MR Bottle Lot", True)
            Call AddValue(Riga, 12, "MR Bottle Qty", True)
            Call AddValue(Riga, 13, "Note", True)
            
            
            Riga = Riga + 1
            
            
            
            Call AddValue(Riga, 2, "'" & .HannaCode)
            
            Call AddValue(Riga, 3, "'" & FormatNumber(uPreparation.MS.Qty, 3) & " mL")
            
            Call AddValue(Riga, 4, "'" & FormatDataLAT(CStr(.DataPrep)))
            Call AddValue(Riga, 5, "'" & .HourPrep)
            Call AddValue(Riga, 6, "'" & .WeekPrep)
            Call AddValue(Riga, 7, "'" & FormatDataLAT(CStr(.DataMS)))
            Call AddValue(Riga, 8, "'" & .QtyProduced & " mL")
            Call AddValue(Riga, 9, "'" & FormatDataLAT(CStr(.DataExp)))
            
            Call AddValue(Riga, 10, "'" & .Bottle.EntryBottle)
            Call AddValue(Riga, 11, "'" & .Bottle.Lot)
            Call AddValue(Riga, 12, "'" & .Bottle.StockQTY & " mL")
            Call AddValue(Riga, 13, "'" & .Note)
          
            
        End With

ERR_END:
   Riga = Riga + 2
   On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox Err.Description
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
    


        With dbTabPreparationNotes
            .filter = ""
            .filter = "FileName='" & SettingName & "'"
            If .EOF Then
            Else
                .MoveFirst
                
                
                    Call AddValue(Riga - 1, 2, "Preparation Notes", True, True)

                    Call AddValue(Riga, 2, "Date", True)
                    Call AddValue(Riga, 3, "Type", True)
                    Call AddValue(Riga, 4, "Description", True)
                    Call AddValue(Riga, 5, "Operator", True)
                    
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
    MsgBox Err.Description
    Resume Next

End Sub



Private Sub SetHannaCode(ByRef Riga As Integer)
Dim i As Integer
Dim t As Integer
Dim x As Integer

    Riga = Riga + 2
    
    Call AddValue(Riga - 1, 2, "Hanna Code", True, True)



On Error GoTo ERR_GET:
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
        
     
        
        CloseSettingDataFile

        With uPreparation.HannaCode

            
            Call AddValue(Riga, 2, "Hanna Code", True)
            Call AddValue(Riga, 3, "Description", True)
            Call AddValue(Riga, 4, "Recipe/Meter", True)

            Call AddValue(Riga, 5, "FW Hanna Parameter", True)
            Call AddValue(Riga, 6, "Measurement Unit", True)
            Call AddValue(Riga, 7, "Decimal", True)
            Call AddValue(Riga, 8, "MS1 Value ", True)
            Call AddValue(Riga, 9, "MS1 Volume (ml)", True)
            Call AddValue(Riga, 10, "MS2dil ", True)
            Call AddValue(Riga, 11, "MS2 Vol. (ml)", True)
            Call AddValue(Riga, 12, "STD Matrix", True)
            Call AddValue(Riga, 13, "STD Vol.(mL)", True)
            Call AddValue(Riga, 14, "STD Exp.", True)
            Call AddValue(Riga, 15, "Storage STD ", True)

                Riga = Riga + 1
                
            
            Call AddValue(Riga, 2, "'" & .Code)
            Call AddValue(Riga, 3, "'" & .Description)
            Call AddValue(Riga, 4, "'" & .Recipe)
            Call AddValue(Riga, 5, "'" & .FWHannaParameter)
            Call AddValue(Riga, 6, "'" & .MeasurementUnit)
            Call AddValue(Riga, 7, "'" & .Decimal)
            Call AddValue(Riga, 8, "'" & .MS1val)
            Call AddValue(Riga, 9, "'" & .MS1vol)
            Call AddValue(Riga, 10, "'" & .MS2Dil)
            Call AddValue(Riga, 11, "'" & .MS2vol)
            Call AddValue(Riga, 12, "'" & .STDMatrix)
            Call AddValue(Riga, 13, "'" & .STDVolume)
            Call AddValue(Riga, 14, "'" & .STDExp)
            Call AddValue(Riga, 15, "'" & .STDStorage)
        

        
            
            Riga = Riga + 2
        
            
            
            
    End With


    Riga = Riga + 3
 
ERR_END:
   CloseSettingDataFile
   
   On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox Err.Description
    Resume Next

End Sub


Private Sub SetMR(ByRef Riga As Integer)
Dim i As Integer
Dim t As Integer
Dim x As Integer

    Riga = Riga + 2
    
    Call AddValue(Riga - 1, 2, "Chemical MR", True, True)



On Error GoTo ERR_GET:
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
        
     
        
        CloseSettingDataFile

        With uPreparation.HannaCode.MR

            
            Call AddValue(Riga, 2, "Code", True)
            Call AddValue(Riga, 3, "Description", True)
            Call AddValue(Riga, 4, "MR physical state", True)

            Call AddValue(Riga, 5, "MR Density 20°C(g/mL)", True)
      
            Call AddValue(Riga, 6, "MR MNP", True)
            Call AddValue(Riga, 7, "MR PURITY %", True)
            Call AddValue(Riga, 8, "MR Value", True)
            Call AddValue(Riga, 9, "MR Unit", True)
            Call AddValue(Riga, 10, "MR Parameter", True)
            Call AddValue(Riga, 11, "FW MR Parameter", True)
        
                Riga = Riga + 1
                
            
            Call AddValue(Riga, 2, "'" & .Code)
            Call AddValue(Riga, 3, "'" & .Description)
            Call AddValue(Riga, 4, "'" & .PhysicalState)
            Call AddValue(Riga, 5, "'" & .Density)
            Call AddValue(Riga, 6, "'" & .MNP)
            Call AddValue(Riga, 7, "'" & .MRPurity & " %")
            Call AddValue(Riga, 8, "'" & .MRValue)
            Call AddValue(Riga, 9, "'" & .Unit)
            Call AddValue(Riga, 10, "'" & .Parameter)
            Call AddValue(Riga, 11, "'" & .FWParameter)
       
         
            
            Riga = Riga + 2
        
            
            
            
    End With


    Riga = Riga + 3
 
ERR_END:
   CloseSettingDataFile
   
   On Error GoTo 0
    Exit Sub
ERR_GET:
    MsgBox Err.Description
    Resume Next

End Sub


