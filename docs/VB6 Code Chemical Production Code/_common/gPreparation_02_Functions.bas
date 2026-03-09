Attribute VB_Name = "gPreparation_02_Functions"
Option Explicit


Public Function ColorTolerance(ByVal Variance As Double, ByVal Toll As Double, ByRef bRecalculate As Boolean, ByRef bCorrection As Boolean) As OLE_COLOR
Dim rc As Boolean
Dim MyColor As OLE_COLOR
    
    rc = False
   
    Select Case Abs(Variance) - Toll
    Case Is < 0
        ' in tolleranza
        MyColor = &H8000&
        rc = True
    Case 0 To Toll / 2
        MyColor = vbColorOrange
        
    Case Is > Toll / 2
        MyColor = &HC0&
        bCorrection = True
    
    End Select
    
    bRecalculate = Not (rc)
 
    ColorTolerance = MyColor
    
End Function



Public Function OpenProductCalssification(ByVal Code As String, ByVal Index As Integer)
Dim MyID As Long

    If Code = "" Then Exit Function
    
    Select Case Index
    
        Case 0
            ' è un hannacode...
            MyID = 1
        Case 1
            ' è un raw material
            MyID = GetIDRowMaterial(Code)
        Case 2
            ' è una ricetta Recipe
            MyID = GetRecipeIdByName(Code)
            
            ' aggiorno HannaCode...
            Call SetCodeClassificationByRecipe(True, MyID)
    End Select
    
    If MyID > 0 Then Call F_PICTOGRAM.DoShow(MyID, Index, Code)

End Function




Public Function AddNewComponentToRecipe(ByRef iRmxRecipe() As RmxRecipe, ByRef MaxCount As Integer, ByVal ChemicalCode As String, ByVal RecipeCode As String, ByVal ActualWeight As Double)


    With dbTabRawMaterial
        .filter = ""
        .filter = "Code='" & ChemicalCode & "'"
        If .EOF Then
        Else
        
            MaxCount = MaxCount + 1
            ReDim Preserve iRmxRecipe(MaxCount)
            iRmxRecipe(MaxCount).CHCode = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
            iRmxRecipe(MaxCount).Description = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
            iRmxRecipe(MaxCount).Cas = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
            iRmxRecipe(MaxCount).bMix = !bMix 'IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
            iRmxRecipe(MaxCount).RecipeCode = RecipeCode
            iRmxRecipe(MaxCount).TheoreticalWeight = ActualWeight
            iRmxRecipe(MaxCount).RealWeight = ActualWeight
            iRmxRecipe(MaxCount).UmTheoreticalWeight = "g"
            iRmxRecipe(MaxCount).TolerancePerc = 1
            iRmxRecipe(MaxCount).bAddedInPreparation = True
            iRmxRecipe(MaxCount).CriticalRM = IIf(IsNull(Trim(!CriticalRM)), "", Trim(!CriticalRM))
        End If
        
    End With
     

End Function

Public Function AddNewRowInAcquisition(ByRef Grid2 As Grid, ByRef iAcquisition As PrepAcquisition)
Dim i As Integer
Dim t As Integer
With Grid2

 
        .AddItem "", False
        i = .Rows - 1
        '.Cell(0, 1).Text = "Code"
        '.Cell(0, 2).Text = "Description"
        '.Cell(0, 3).Text = "CAS"
        '.Cell(0, 4).Text = "Real Weight (g)"
        '.Cell(0, 5).Text = "Manufacturer"
        '.Cell(0, 6).Text = "Manufacturer Code"
        '.Cell(0, 7).Text = "Manufacturer Lot"
        '.Cell(0, 8).Text = "Delivery Date"
        '.Cell(0, 9).Text = "Qty Delivered"
        ''.Cell(0, 10).Text = "Week Delivery"
        '.Cell(0, 11).Text = "Package"
        '.Cell(0, 12).Text = "Note"
        '.Cell(0, 13).Text = "Operator"
        '.Cell(0, 14).Text = "Acquisition Time"
        .Cell(i, 1).Text = iAcquisition.PrepBarcode.Code
        .Cell(i, 2).Text = iAcquisition.PrepBarcode.ChemicalName
        .Cell(i, 3).Text = iAcquisition.PrepBarcode.Cas
        .Cell(i, 4).Text = PadString(iAcquisition.ActualWeight)
        .Cell(i, 5).Text = iAcquisition.PrepBarcode.Manufacturer
        .Cell(i, 6).Text = iAcquisition.PrepBarcode.ManufacturerCode
        .Cell(i, 7).Text = iAcquisition.PrepBarcode.ManufacturerLot
        .Cell(i, 8).Text = iAcquisition.PrepBarcode.DeliveryDate
        .Cell(i, 9).Text = iAcquisition.PrepBarcode.QtyDelivered
        .Cell(i, 10).Text = iAcquisition.PrepBarcode.WeekDelPackageNumber
        .Cell(i, 11).Text = iAcquisition.PrepBarcode.Package
        .Cell(i, 12).Text = iAcquisition.Note
        .Cell(i, 13).Text = iAcquisition.Operator
        .Cell(i, 14).Text = iAcquisition.AcquisitionTime
        .Cell(i, 15).Text = iAcquisition.ID
        .Cell(i, 16).Text = iAcquisition.ExpDate
        
        
        
        .Cell(i, 4).BackColor = vbColorResults
        .Cell(i, 4).Alignment = cellRightCenter
        
        If iAcquisition.bRecalculation Then
            '.Cell(i, 0).BackColor = &HC0&
            For t = 1 To 3
                .Cell(i, t).ForeColor = &HC0&
                .Cell(i, t).FontBold = True
            Next
        End If
        If iAcquisition.bRecipeComponent = False Then
           ' .Cell(i, 0).BackColor = &HFFFF&
            For t = 1 To 3
            
                .Cell(i, t).BackColor = &HC0C0&    ' &HFFFF&
                .Cell(i, t).ForeColor = vbWhite
                
                .Cell(i, t).FontBold = True
               ' .Cell(i, t).BackColor = &HC0C0C0
            Next
        End If
        
     
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Alignment = cellCenterCenter
        .Column(8).Alignment = cellCenterCenter
        .Column(9).Alignment = cellCenterCenter
        .Column(10).Alignment = cellCenterCenter
        .Column(11).Alignment = cellCenterCenter
        
End With

End Function



Public Function AggiornaTabPreparation(ByVal PreparationID As Long, ByRef uPreparation As RecipeForProduction)

With dbTabPreparation
    .filter = ""
    .filter = "ID='" & PreparationID & "'"
    If .EOF Then
    
    Else
        !QtyToProduce = uPreparation.Recipes(1).TotalWeightKg
        !QtyProduced = uPreparation.Recipes(1).ActualWeight
        
        !PrepWeek = uPreparation.PrepWeek
        !PrepDate = FormatDataLAT(uPreparation.PreparationDate)
        !ExpDate = uPreparation.ExpDate
        !numPrepWeek = uPreparation.numPrepWeek
        !bPesatoTuttiComponenti = uPreparation.bPesatoTuttiComponenti
        !bCorrection = uPreparation.bCorrection
        If uPreparation.QCCount > 0 Then
            
            !PassToQC = True
            !QCStatus = uPreparation.QCStatus(uPreparation.QCCount).Status
            !QCOperator = uPreparation.QCStatus(uPreparation.QCCount).Operator
            !QCDate = uPreparation.QCStatus(uPreparation.QCCount).Date
            !QCNote = uPreparation.QCStatus(uPreparation.QCCount).Note
            
        End If
        Call SetHannaCodeAndLot(uPreparation)
        .Update
    End If


End With

End Function


Private Function SetHannaCodeAndLot(ByRef uPreparation As RecipeForProduction)
Dim strLot As String
Dim strCode As String
Dim strSep As String
Dim i As Integer


    With dbTabPreparation
        If .EOF Then
        Else
            strCode = IIf(IsNull(Trim(!HannaCode)), "", Trim(!HannaCode))
            strLot = IIf(IsNull(Trim(!Lot)), "", Trim(!Lot))
        End If
    End With
    
    
    With uPreparation

       strLot = .Recipes(1).PreparationLotMix
      If strCode = "" And Not (.Recipes(1).bIsMix) Then strCode = .HannaCodes(1).Code
        
    End With
    
    
    With dbTabPreparation
        If .EOF Then
        Else
            !HannaCode = IIf(IsNull(!HannaCode), Trim(strCode), !HannaCode)
            !Lot = Trim(strLot)
            .Update
        End If
        
    End With


End Function


Public Function GetRawMaterialManufacturer(ByVal Code As String, ByRef Manufacturer As String, ByRef ManufacturerCode As String)
Dim rc As Boolean

    With dbTabRawMaterial
        .filter = ""
        .filter = "Code='" & Code & "'"
        If .EOF Then
        Else
            Manufacturer = IIf(IsNull(Trim(!ManufacturerName)), "", Trim(!ManufacturerName))
            ManufacturerCode = IIf(IsNull(Trim(!ManufacturerCode)), "", Trim(!ManufacturerCode))
        End If
    End With
End Function
