Attribute VB_Name = "ChemicalMRSetTypes"
Option Explicit

Public Function SetHannaCodeByCode(ByVal Code As String, ByRef uCode As HannaCode)


On Error GoTo ERR_SET:
    If Code = "" Then Exit Function
    Dim i As Integer
    
    uCode.Code = Code
    uCode.Description = "Not Found in Database!"
     With dbTabCode
            
        .filter = ""
        .filter = "Code='" & Code & "'"
        If .EOF Then
            
            Exit Function
        End If
        .MoveFirst
        uCode.Code = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
        
        uCode.Decimal = IIf(IsNull(Trim(!Decimal)), 0, Trim(!Decimal))
        uCode.Description = IIf(IsNull(Trim(!ProductName)), "", Trim(!ProductName))
     
        uCode.FWHannaParameter = CheckDot(IIf(IsNull(Trim(!FWParameterFormula)) Or Trim(!FWParameterFormula) = "", 0, Trim(!FWParameterFormula)))
        uCode.ID = !ID
        uCode.MeasurementUnit = IIf(IsNull(Trim(!MeasurementUnit)), "", Trim(!MeasurementUnit))
        
        uCode.MR.Code = IIf(IsNull(Trim(!STDMR)), "", Trim(!STDMR))
        
        uCode.MS1val = (IIf(IsNull(Trim(!MS1val)), "", Trim(!MS1val)))
        uCode.MS1vol = (IIf(IsNull(Trim(!MS1vol)), "", Trim(!MS1vol)))
        uCode.MS2Dil = (IIf(IsNull(Trim(!MS2Dil)), "", Trim(!MS2Dil)))
        uCode.MS2vol = (IIf(IsNull(Trim(!MS2vol)), "", Trim(!MS2vol)))
        uCode.MSEXP = IIf(IsNull(Trim(!MSEXP)), "", Trim(!MSEXP))
        uCode.STDType = IIf(uCode.MS1val <> "", 1, IIf(uCode.MS2Dil <> "", 2, 0))
        
        uCode.Hannaformula = IIf(IsNull(Trim(!ParameterFormula)), "", Trim(!ParameterFormula))
        uCode.ParameterMethod = IIf(IsNull(Trim(!ParameterMethod)), "", Trim(!ParameterMethod))
        
        uCode.RangeMax = (IIf(IsNull(Trim(!RangeMax)), "", Trim(!RangeMax)))
        uCode.RangeMin = (IIf(IsNull(Trim(!RangeMin)), "", Trim(!RangeMin)))
        uCode.Recipe = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
        
       
      

        uCode.STDExp = IIf(IsNull(Trim(!STDExp)), "", Trim(!STDExp))
        uCode.STDMatrix = (IIf(IsNull(Trim(!STDMatrix)), "", Trim(!STDMatrix)))
        uCode.STDMR2 = IIf(IsNull(Trim(!STDMR2)), "", Trim(!STDMR2))
        uCode.STDNote = IIf(IsNull(Trim(!STDNote)), "", Trim(!STDNote))
        uCode.STDStorage = IIf(IsNull(Trim(!STDStorage)), "", Trim(!STDStorage))
        
        If (IsNull(Trim(!STDVolume)) Or Trim(!STDVolume) = "") Then
            !STDVolume = 500
            .Update
        End If
        uCode.STDVolume = (IIf(IsNull(Trim(!STDVolume)) Or Trim(!STDVolume) = "", "500", Trim(!STDVolume)))
       
        uCode.UnitMR = "mL"
        
       
        
        Call SetMRFromDatabase(uCode.MR.Code, uCode.MR)
        uCode.STDUnit = uCode.MR.Unit
         Call SetSTDFromDatabase(uCode.Code, uCode.STD(), uCode.STDcount)

        If uCode.MR.FWParameter <> 0 Then
            uCode.ConcHannaParameter = FormatNumber((uCode.MR.MRPurity / 100) * uCode.MR.MRValue * uCode.FWHannaParameter / uCode.MR.FWParameter, uCode.Decimal + 2)
        End If
    
    End With


ERR_END:
    On Error GoTo 0
    Exit Function
ERR_SET:
    MsgBox Err.Description
    Resume Next

End Function


Public Function SetMRFromDatabase(ByVal Code As String, ByRef MR As MRType) As Boolean
Dim i As Integer
Dim rc As Boolean

    On Error GoTo ERR_SET:
    
    
    
    With dbTabMR
        
        .filter = ""
        .filter = "Code='" & Replace(Code, "'", "''") & "'"
        If .EOF Then
            rc = False
            GoTo ERR_END
        Else
            rc = True
        End If
        
       MR.Classification = IIf(IsNull(Trim(!Classification)), "", Trim(!Classification))
       MR.Code = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
       MR.Density = IIf(IsNull(Trim(!Density)), 1, Trim(!Density))
       MR.Description = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
       MR.FWParameter = IIf(IsNull(Trim(!FWParameter)) Or Trim(!FWParameter) = "", 0, Trim(!FWParameter))
       MR.Location = IIf(IsNull(Trim(!Location)), "", Trim(!Location))
       MR.MinQty = IIf(IsNull(Trim(!MinQty)), 0, Trim(!MinQty))
       MR.MNP = IIf(IsNull(Trim(!MNP)), "", Trim(!MNP))
       MR.Modified = IIf(IsNull(Trim(!Modified)), "", Trim(!Modified))
       MR.MRPurity = IIf(IsNull(Trim(!MRPurity)), 100, Trim(!MRPurity))
       MR.MRValue = IIf(IsNull(Trim(!MRValue)) Or Trim(!MRValue) = "", 1000, Trim(!MRValue))
       MR.Parameter = IIf(IsNull(Trim(!Parameter)), "", Trim(!Parameter))
       MR.PhysicalState = IIf(IsNull(Trim(!Code)), "", Trim(!PhysicalState))
      
       MR.ReductionExpDays = IIf(IsNull(Trim(!ReductionExpDays)), 120, Trim(!ReductionExpDays))
       If Trim(MR.ReductionExpDays) = "" Then MR.ReductionExpDays = 120
       
       MR.Rev = IIf(IsNull(Trim(!Rev)), "", Trim(!Rev))
       MR.STOCK_QTY = IIf(IsNull(Trim(!STOCK_QTY)), "", Trim(!STOCK_QTY))
       MR.STOCK_UNIT = IIf(IsNull(Trim(!STOCK_UNIT)), "", Trim(!STOCK_UNIT))
       MR.StorageT = IIf(IsNull(Trim(!StorageT)), "", Trim(!StorageT))
       MR.Supplier = IIf(IsNull(Trim(!Supplier)), "", Trim(!Supplier))
       MR.Unit = IIf(IsNull(Trim(!Unit)), "", Trim(!Unit))
      
       
    End With
    
    
ERR_END:
    On Error GoTo 0

    SetMRFromDatabase = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox Err.Description
    Resume Next
 
End Function

Public Function SetSTDFromDatabase(ByVal HannaCode As String, ByRef STD() As STD, ByRef Count As Integer) As Boolean
Dim i As Integer
Dim rc As Boolean

    On Error GoTo ERR_SET:
    
    With dbTabCode
        
        .filter = ""
        .filter = "Code='" & Replace(HannaCode, "'", "''") & "'"
        If .EOF Then
            rc = False
            GoTo ERR_END
        Else
            rc = True
        End If
        
        ReDim STD(0)
        
        i = 0
    
       If Not (IsNull(!STD1Value)) And IsNumeric(!STD1Value) Then
            If Trim(!STD1Value) >= 0 Then
                ReDim Preserve STD(1)
                STD(1).NUMBER = 1
                STD(1).Value = CDbl(Trim(!STD1Value))
                i = 1
            End If
       End If
             
       If Not (IsNull(!STD2Value)) And IsNumeric(!STD2Value) Then
            If Trim(!STD2Value) >= 0 Then
                ReDim Preserve STD(2)
                STD(2).NUMBER = 2
                STD(2).Value = CDbl(Trim(!STD2Value))
                i = 2
            End If
       End If
       
       If Not (IsNull(!STD3Value)) And IsNumeric(!STD3Value) Then
            If Trim(!STD3Value) > 0 Then
                ReDim Preserve STD(3)
                STD(3).NUMBER = 3
                STD(3).Value = CDbl(Trim(!STD3Value))
                i = 3
            End If
       End If
       
       If Not (IsNull(!STD4Value)) And IsNumeric(!STD4Value) Then
            If Trim(!STD4Value) > 0 Then
                ReDim Preserve STD(4)
                STD(4).NUMBER = 4
                STD(4).Value = CDbl(Trim(!STD4Value))
                i = 4
            End If
       End If
       
       If Not (IsNull(!STD5Value)) And IsNumeric(!STD5Value) Then
            If Trim(!STD5Value) > 0 Then
                ReDim Preserve STD(5)
                STD(5).NUMBER = 5
                STD(5).Value = CDbl(Trim(!STD5Value))
                i = 5
            End If
       End If
       
       If Not (IsNull(!STD6Value)) And IsNumeric(!STD6Value) Then
            If Trim(!STD6Value) > 0 Then
                ReDim Preserve STD(6)
                STD(6).NUMBER = 6
                STD(6).Value = CDbl(Trim(!STD6Value))
                i = 6
            End If
       End If
    End With

ERR_END:
    On Error GoTo 0
    Count = i
    SetSTDFromDatabase = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox Err.Description
    Resume Next

End Function

