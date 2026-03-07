Attribute VB_Name = "MaterialRequisitionFunctions"
Option Explicit
Private SettingName As String


Public Function MaterialRequisitionDeleteRecord(ByRef Grd As Grid)

    Grd.ReadOnly = False
    Grd.Selection.DeleteByRow
    Grd.ReadOnly = True
    
End Function


Public Function AddTotalWeightMixesAllRecipes(ByRef iRecipe() As RecipeType, ByVal strMixCode As String, ByRef TotalWeightKg As Double, ByRef TotalWeightL As Double, ByRef MultipleToProduce As Double) As Boolean
Dim i As Integer
Dim t As Integer

    For i = 1 To UBound(iRecipe)
        If iRecipe(i).bHide = False And iRecipe(i).bUpdated Then
              If iRecipe(i).Code = strMixCode Then
                    TotalWeightKg = TotalWeightKg + iRecipe(i).TotalWeightKg
                    TotalWeightL = TotalWeightKg / IIf(iRecipe(i).Density = 0, 1, iRecipe(i).Density)
                    Exit For
                End If
            For t = 0 To UBound(iRecipe(i).RmxRecipe)
            
              
                
                If iRecipe(i).RmxRecipe(t).CHCode = strMixCode Then
                    If iRecipe(i).RmxRecipe(t).TheoreticalWeight > 0 Then
                        Debug.Print "Recipe Code " & iRecipe(i).Code & " component " & iRecipe(i).RmxRecipe(t).CHCode & " component Recipe " & iRecipe(i).RmxRecipe(t).RecipeCode
                        If iRecipe(i).RmxRecipe(t).UmTheoreticalWeight = "" Then iRecipe(i).RmxRecipe(t).UmTheoreticalWeight = "g"
                        TotalWeightKg = TotalWeightKg + iRecipe(i).RmxRecipe(t).TheoreticalWeight * Um(iRecipe(i).RmxRecipe(t).UmTheoreticalWeight) / 1000
                        TotalWeightL = TotalWeightKg / IIf(iRecipe(i).RmxRecipe(t).Density = 0, 1, iRecipe(i).RmxRecipe(t).Density)
                    End If
                End If
                
            Next
        End If
    
    Next
    
    'Call AddTotalWeightRecipeMixAllRecipes(iRecipe(), strMixCode, TotalWeightKg, TotalWeightL, MultipleToProduce)
    
    
End Function
Public Function AddTotalWeightRecipeMixAllRecipes(ByRef iRecipe() As RecipeType, ByVal strMixCode As String, ByRef TotalWeightKg As Double, ByRef TotalWeightL As Double, ByRef MultipleToProduce As Double) As Boolean
Dim i As Integer
Dim t As Integer

    For i = 1 To UBound(iRecipe)
        If iRecipe(i).bHide = False And iRecipe(i).bUpdated And iRecipe(i).bIsMix Then
            
                If iRecipe(i).Code = strMixCode Then
                
                
                    If iRecipe(i).TotalWeightKg > 0 Then
                        TotalWeightKg = TotalWeightKg + iRecipe(i).TotalWeightKg * Um("kg") / 1000
                        TotalWeightL = TotalWeightL + (iRecipe(i).TotalWeightKg * Um("kg") / 1000) / IIf(iRecipe(i).Density = 0, 1, iRecipe(i).Density)
                    End If
                End If
                
           
        End If
    
    Next
End Function

Public Function SetMaterialRequisitionAllMixes(ByRef RecipeMaterialReq() As RecipeType, ByVal Grid4 As Grid) As Boolean
Dim rc As Boolean
Dim t As Integer
Dim i As Integer
Dim Count As Integer
Dim ComponentCode As String
Dim IndexComponent As Integer
Dim TotalWeightKg As Double
Dim TotalWeightL As Double
Dim MultipleTP As Double

On Error GoTo ERR_SET
    
    Count = 0
    
    rc = True
    
    
    ReDim RecipeMaterialReq(Count)
        
    With Grid4
        If .Rows < 2 Then Exit Function
        
            For i = 1 To .Rows - 1
            
            
                If .RowHeight(i) > 0 Then
                
                    If CBool(.Cell(i, 15).Text) And .Cell(i, 1).Text <> "" Then
                    
                        Count = Count + 1
                        
                        ReDim Preserve RecipeMaterialReq(Count)
                        
                        Call SetMyRecipeByCode(.Cell(i, 1).Text, RecipeMaterialReq(Count))
                        

                        
                      
                        
                        RecipeMaterialReq(Count).TotalWeightKg = CDbl(.Cell(i, 3).Text)
                        RecipeMaterialReq(Count).UmTotalWeightKg = "kg"
                        
                        
                    
                    End If
                    
                End If
                                            
            
            Next
            
            
        End With

ERR_END:
    On Error GoTo 0
    SetMaterialRequisitionAllMixes = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox err.Description
    Resume Next

End Function


Public Function SetMaterialRequisition(ByRef Recipe() As RecipeType, ByRef RecipeMaterialReq As RecipeType, ByVal i As Integer, ByVal bValue As Boolean) As Boolean
Dim rc As Boolean
Dim t As Integer
Dim Count As Integer
Dim ComponentCode As String
Dim IndexComponent As Integer
Dim TotalWeightKg As Double
Dim TotalWeightL As Double
Dim MultipleTP As Double


On Error GoTo ERR_SET
    
    Count = 0
    
    rc = True
    
        Recipe(i).bHaveMixes = False 'IfAllMixes(Recipe(i).Code)
    
        ReDim Preserve RecipeMaterialReq.RmxRecipe(Count)
        
        For t = LBound(Recipe(i).RmxRecipe) To UBound(Recipe(i).RmxRecipe)
        
        
        
           With Recipe(i).RmxRecipe(t)
                  
                  If Recipe(i).bHaveMixes = False Then GoTo addComponent
                  
                  If .bMix And bValue = False Then
                  
                      If .RecipeCode = Recipe(i).Code Then
                          ' è una ricetta non la metto...
                      Else
                          GoTo addComponent
                      End If
                      
                  ElseIf .bMix And bValue And .RecipeCode = Recipe(i).Code Then
                      GoTo addComponent:
                      
                  ElseIf bValue = False Then
addComponent:
                      
                      If Count > 0 Then
                      
                          'ComponentCode = .CHCode
                          IndexComponent = CheckRmxRecipeInRecipe(RecipeMaterialReq.RmxRecipe(), .CHCode, Recipe(i).Code)
                          If IndexComponent = -1 Then
Aggiungi:
                              
                             ReDim Preserve RecipeMaterialReq.RmxRecipe(Count)
                             
               
                             
                             RecipeMaterialReq.RmxRecipe(Count) = Recipe(i).RmxRecipe(t)
                             
                            
                             If bValue Then
                                TotalWeightKg = 0
                                Call AddTotalWeightMixesAllRecipes(Recipe, Recipe(i).RmxRecipe(t).CHCode, TotalWeightKg, TotalWeightL, MultipleTP)

                                RecipeMaterialReq.RmxRecipe(Count).TheoreticalWeight = TotalWeightKg
                                RecipeMaterialReq.RmxRecipe(Count).UmTheoreticalWeight = "kg"

                             End If

                             Count = Count + 1
                          Else

                              RecipeMaterialReq.RmxRecipe(IndexComponent).TheoreticalWeight = RecipeMaterialReq.RmxRecipe(IndexComponent).TheoreticalWeight + Recipe(i).RmxRecipe(t).TheoreticalWeight
                              
                          End If
                      Else
                          GoTo Aggiungi
                      End If
        
                  End If
            End With
                  
            
        Next


ERR_END:
    On Error GoTo 0
    SetMaterialRequisition = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox err.Description
    Resume Next

End Function

Public Function SetMaterialRequisitionPreparation(ByRef Recipe() As RecipeType, ByRef RecipeMaterialReq As RecipeType, ByRef Grid As Grid, ByVal bPreparation As Boolean, Optional ByVal FileName As String) As Boolean
Dim rc As Boolean
Dim t As Integer
Dim Count As Integer
Dim ComponentCode As String
Dim IndexComponent As Integer
Dim TotalWeightKg As Double
Dim TotalWeightL As Double
Dim MultipleTP As Double
Dim i As Integer
Dim X As Integer


On Error GoTo ERR_SET

    rc = True
    
        i = 1
    
        ReDim Preserve RecipeMaterialReq.RmxRecipe(Recipe(1).RmxRecipeCount)
        
        For t = 0 To UBound(Recipe(i).RmxRecipe)
        
        
        
           With Grid
                  

                  ReDim Preserve RecipeMaterialReq.RmxRecipe(t)
                  
    
                  
                    RecipeMaterialReq.RmxRecipe(t) = Recipe(i).RmxRecipe(t)
                     If IsNumeric(Grid.Cell(t + 1, 6).Text) Then
                        TotalWeightKg = FormatNumber(Grid.Cell(t + 1, 6).Text / 1000, 2)  ' trasformo in kg
                     End If
                     If bPreparation Then
                        RecipeMaterialReq.RmxRecipe(t).TheoreticalWeight = IIf(IsNumeric(Grid.Cell(t + 1, 6).Text), Grid.Cell(t + 1, 6).Text, 0)
                        RecipeMaterialReq.RmxRecipe(t).UmTheoreticalWeight = "g"
                        RecipeMaterialReq.RmxRecipe(t).RealWeight = Recipe(1).RmxRecipe(t).RealWeight
                        
                        Call GetManufacturerLot(RecipeMaterialReq.RmxRecipe(t), FileName)
                     Else
                        RecipeMaterialReq.RmxRecipe(t).TheoreticalWeight = TotalWeightKg
                        RecipeMaterialReq.RmxRecipe(t).UmTheoreticalWeight = "kg"
                        RecipeMaterialReq.RmxRecipe(t).RealWeight = FormatNumber(Recipe(1).RmxRecipe(t).RealWeight / 1000, 2) ' trasformo in kg
                     End If
                    
                     Call GetRMLocation(RecipeMaterialReq.RmxRecipe(t))
                     
       
            End With
                  
            
        Next


ERR_END:
    On Error GoTo 0
    SetMaterialRequisitionPreparation = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox err.Description
    Resume Next

End Function

Public Function GetManufacturerLot(ByRef uRMxRecipe As RmxRecipe, ByVal FileName As String)
Dim i As Integer
Dim Lot As String
With dbTabAcquisition
    .filter = ""
    .filter = "FileName='" & FileName & "' and Code='" & uRMxRecipe.CHCode & "'"
    If .EOF Then
       uRMxRecipe.ManufacturerLot = ""
    Else
        If .RecordCount > 1 Then
            .MoveFirst
            For i = 1 To .RecordCount
                Lot = IIf(IsNull(Trim(!ManufacturerLot)), "", Trim(!ManufacturerLot))
                If Lot <> "" Then
                    uRMxRecipe.ManufacturerLot = uRMxRecipe.ManufacturerLot & IIf(InStr(uRMxRecipe.ManufacturerLot, Lot), "", IIf(Len(uRMxRecipe.ManufacturerLot) > 0, ";", "") & Lot)
                End If
                .MoveNext
            Next
        Else
            uRMxRecipe.ManufacturerLot = IIf(IsNull(Trim(!ManufacturerLot)), "", Trim(!ManufacturerLot))
        End If
    End If
End With
End Function

Public Function ChangeMaterialPreparationReqQty(ByVal Grid6 As Grid, Index As Long, ByVal Row As Long)
Dim Um As String
Dim Qty As String
Dim OriginQty As String
Dim sString As String
Dim Value() As String

With Grid6
    
    If Index > 0 And Index <= .Rows - 1 Then
        
        sString = .Cell(Index, Row).Text
    
        If sString <> "" Then
            
            'Value = Split(Trim(sString), " ")
            'Debug.Print UBound(Value)
            'If UBound(Value) = 1 Then
            
                OriginQty = sString
                Qty = sString
               
                sString = .Cell(Index, 1).Text
                
                If F_InputBox.DoShow("Confirm or Change Qty", sString, , , , Qty) Then

                    If Qty <> "" Then
                    
                        .Cell(Index, Row).Text = (Qty)
                    
                    End If
                End If
            'End If
                
        End If
    End If


End With





End Function


Public Function ChangeMaterialReqQty(ByVal Grid6 As Grid, Index As Long)
Dim Um As String
Dim Qty As String
Dim OriginQty As String
Dim sString As String
Dim Value() As String

With Grid6
    
    If Index > 0 And Index <= .Rows - 1 Then
        
        sString = .Cell(Index, 4).Text
    
        If sString <> "" Then
            
            'Value = Split(Trim(sString), " ")
            'Debug.Print UBound(Value)
            'If UBound(Value) = 1 Then
            
                OriginQty = sString
                Qty = sString
               
                sString = .Cell(Index, 1).Text
                
                If F_InputBox.DoShow("Confirm or Change Qty", sString, , , , Qty) Then

                    If Qty <> "" Then
                    
                        .Cell(Index, 4).Text = (Qty)
                    
                    End If
                End If
            'End If
                
        End If
    End If


End With





End Function



Public Sub AddMixesToMaterialReqGrid(ByVal Grd As Grid, ByRef ReqRecipe() As RecipeType)
Dim i As Integer
Dim t As Integer
Dim X As Integer

    With Grd
    
        .Rows = 1
        .AutoRedraw = False
        
        For t = 1 To UBound(ReqRecipe)
            
                .AddItem "", False
                .Cell(.Rows - 1, 1).Text = ReqRecipe(t).Code
                .Cell(.Rows - 1, 2).Text = ReqRecipe(t).Description
                .Cell(.Rows - 1, 3).Text = ReqRecipe(t).Cas
                .Cell(.Rows - 1, 4).Text = PadString(ReqRecipe(t).TotalWeightKg) & " " & ReqRecipe(t).UmTotalWeightKg
                .Cell(.Rows - 1, 5).Text = ReqRecipe(t).Location
                .Cell(.Rows - 1, 6).Text = ReqRecipe(t).SpecifiedLocation
            
                '.Cell(0, 1).Text = "CH Code"
                '.Cell(0, 2).Text = "Description"
                '.Cell(0, 3).Text = "CAS"
                '.Cell(0, 4).Text = "Q.ty Required"
                '.Cell(0, 5).Text = "Location"
                '.Cell(0, 6).Text = "Specified Location"
        
                
        
                For i = 1 To .Cols - 1
                    .Column(i).Alignment = cellLeftCenter
                    .Column(i).Width = 150
                    .Cell(0, i).FontBold = True
                Next
                
                '.Column(0).Width = 0
                .Column(2).Width = 250
                .Column(3).Width = 100
                .Column(5).Width = 100
                .Column(4).Alignment = cellRightCenter
                .Column(5).Alignment = cellCenterCenter
                .Column(6).Alignment = cellCenterCenter
           
        Next
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
 

End Sub
Public Sub AddRecipeToMaterialReqGrid(ByVal Grd As Grid, ByRef ReqRecipe() As RecipeType, Optional ByVal bPreparation As Boolean)
Dim i As Integer
Dim t As Integer
Dim X As Integer
Dim strWeight As String
Dim MyUM As String
Dim strRealWeight As String


    With Grd
    
        .Rows = 1
        .AutoRedraw = False
        
        For t = 1 To UBound(ReqRecipe)
            For X = LBound(ReqRecipe(t).RmxRecipe) To UBound(ReqRecipe(t).RmxRecipe)
            
                If InStr(ReqRecipe(t).RmxRecipe(X).CHCode, "DW") Then GoTo cont:
           
                .AddItem "", False
                .Cell(.Rows - 1, 1).Text = ReqRecipe(t).RmxRecipe(X).CHCode
                .Cell(.Rows - 1, 2).Text = ReqRecipe(t).RmxRecipe(X).Description
                .Cell(.Rows - 1, 3).Text = ReqRecipe(t).RmxRecipe(X).Cas
                If ReqRecipe(t).RmxRecipe(X).Density = 0 Then ReqRecipe(t).RmxRecipe(X).Density = 1
                If ReqRecipe(t).RmxRecipe(X).Density <> 1 Then
                
                    '-----------------------------
                    ' Density <> 1 è un liquido!
                    '-----------------------------
                    
                    With dbTabRawMaterial
                        .filter = ""
                        .filter = "Code='" & ReqRecipe(t).RmxRecipe(X).CHCode & "'"
                        If .EOF Then
                            GoTo ex:
                        Else
                            MyUM = IIf(IsNull(!Um), "ml", !Um)
                        
                          '  ReqRecipe(t).RmxRecipe(X).Density = IIf(IsNull(Trim(!Density)), 1, Trim(!Density))
                        End If
                        
                    End With
                    '--------------------------------
                    ' trasformare tutto in ml!!!
                    '--------------------------------
                    strWeight = PadString((ReqRecipe(t).RmxRecipe(X).TheoreticalWeight / ReqRecipe(t).RmxRecipe(X).Density)) & "  " & MyUM
                    If bPreparation Then
                        
                        MyUM = "ml"
                        strWeight = PadString((ReqRecipe(t).RmxRecipe(X).TheoreticalWeight / ReqRecipe(t).RmxRecipe(X).Density)) & "  " & MyUM
                         strRealWeight = PadString((ReqRecipe(t).RmxRecipe(X).RealWeight / ReqRecipe(t).RmxRecipe(X).Density)) & "  " & MyUM
                        
                    End If
                Else
ex:
                    strWeight = PadString(ReqRecipe(t).RmxRecipe(X).TheoreticalWeight) & "  " & ReqRecipe(t).RmxRecipe(X).UmTheoreticalWeight
                    If bPreparation Then
                        '--------------------------------
                        ' trasformare tutto in g!!!
                        '--------------------------------
                         ReqRecipe(t).RmxRecipe(X).UmTheoreticalWeight = "g"
                         strWeight = PadString(ReqRecipe(t).RmxRecipe(X).TheoreticalWeight) & "  " & ReqRecipe(t).RmxRecipe(X).UmTheoreticalWeight
                        strRealWeight = PadString(ReqRecipe(t).RmxRecipe(X).RealWeight) & "  " & ReqRecipe(t).RmxRecipe(X).UmTheoreticalWeight
                    End If
                End If
                
                .Cell(.Rows - 1, 4).Text = strWeight
                
                If bPreparation Then
                
                    .Cell(.Rows - 1, 5).Text = strRealWeight
                    .Cell(.Rows - 1, 6).Text = ReqRecipe(t).RmxRecipe(X).ManufacturerLot
                    .Cell(.Rows - 1, 7).Text = ReqRecipe(t).RmxRecipe(X).Specifications.Location
                    .Cell(.Rows - 1, 8).Text = ReqRecipe(t).RmxRecipe(X).Specifications.SpecifiedLocation
                    
                                
                Else
                
                    .Cell(.Rows - 1, 5).Text = ReqRecipe(t).RmxRecipe(X).Specifications.Location
                    .Cell(.Rows - 1, 6).Text = ReqRecipe(t).RmxRecipe(X).Specifications.SpecifiedLocation
                
                End If
                    
                '.Cell(0, 1).Text = "CH Code"
                '.Cell(0, 2).Text = "Description"
                '.Cell(0, 3).Text = "CAS"
                '.Cell(0, 4).Text = "Q.ty Required"
                '.Cell(0, 5).Text = "Location"
                '.Cell(0, 6).Text = "Specified Location"
        
                
        
                For i = 1 To .Cols - 1
                    .Column(i).Alignment = cellLeftCenter
                    .Column(i).Width = 150
                    .Cell(0, i).FontBold = True
                Next
                
                '.Column(0).Width = 0
                .Column(2).Width = 250
                .Column(3).Width = 100
                .Column(5).Width = 100
                .Column(4).Alignment = cellRightCenter
                .Column(5).Alignment = cellCenterCenter
                .Column(6).Alignment = cellCenterCenter
cont:
            Next
        Next
        .Column(3).AutoFit
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
 

End Sub



Public Function MaterialRequisitionSaveSettingsFile(ByVal Grd As Grid, ByRef txDocument() As String, ByVal FileName As String, ByVal IndexRecipe As Integer) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim MRCount As Integer

On Error GoTo ERR_SAVE

    SettingName = FileName


    rc = True
    
    If USER_PATH = "" Then USER_PATH = USER_TEMP_PATH
    
    If SettingName = "" Then
        Exit Function
        'SettingName = FormatNomeFile(txDocument(0) & "." & txDocument(1) & "." & txDocument(2) & "." & txDocument(3)) & "." & USER_ESTENSIONE
    End If

    DoEvents
    
    CloseSettingDataFile


    For i = LBound(txDocument) To UBound(txDocument)
        SaveSettingData SettingName, "Material Requisition" & IndexRecipe, "txDocument(" & i & ")", txDocument(i)
    Next
    
    With Grd
        SaveSettingData SettingName, "Material Requisition" & IndexRecipe, "Rows", .Rows - 1
        For i = 1 To .Rows - 1
            For t = 1 To .Cols - 1
                 SaveSettingData SettingName, "Material Requisition" & IndexRecipe, "Grd(" & i & "," & t & ")", .Cell(i, t).Text
            Next
        Next
    End With


ERR_END:
    On Error GoTo 0
    
     CloseSettingDataFile

     MaterialRequisitionSaveSettingsFile = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function



Public Function MaterialRequisitionSaveSettingsTempFile(ByVal Grd As Grid, ByRef txDocument() As String, ByRef FileName As String, ByVal strHannaCode As String, ByVal strRecipe As String) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim MRCount As Integer

On Error GoTo ERR_SAVE

    SettingName = FileName


    rc = True
    
    If USER_PATH = "" Then USER_PATH = USER_TEMP_PATH
    
    If SettingName = "" Then
        SettingName = FormatNomeFile(txDocument(0) & "." & txDocument(1) & "." & txDocument(2) & "." & txDocument(3)) & "." & USER_ESTENSIONE
    End If
    
    
    If FileExists(USER_PATH & SettingName) Then
        
        If F_MsgBox.DoShow("Warning : Material Requisition already created.", , True, "Overwrite", "Exit") Then
            
            Kill USER_PATH & SettingName
            DoEvents
        Else
            'PopupMessage 2, "Check all Document fields and save ...."
            rc = False
            GoTo ERR_END:
        End If
    End If
    
    DoEvents
    
    CloseSettingDataFile


    For i = LBound(txDocument) To UBound(txDocument)
        SaveSettingData SettingName, "Material Requisition", "txDocument(" & i & ")", txDocument(i)
    Next
    
    SaveSettingData SettingName, "Material Requisition", "strHannaCode", Trim(strHannaCode)
    SaveSettingData SettingName, "Material Requisition", "strRecipe", Trim(strRecipe)
    
  
    With Grd
        SaveSettingData SettingName, "Material Requisition", "Rows", .Rows - 1
        For i = 1 To .Rows - 1
            For t = 1 To .Cols - 1
                 SaveSettingData SettingName, "Material Requisition", "Grd(" & i & "," & t & ")", .Cell(i, t).Text
            Next
        Next
    End With


ERR_END:
    On Error GoTo 0
    
     CloseSettingDataFile
     FileName = SettingName
     MaterialRequisitionSaveSettingsTempFile = rc
     Exit Function
ERR_SAVE:
     rc = False
     MsgBox err.Description
     Resume Next
End Function



Public Function CheckStrMaterialRequisition(ByVal OldMaterialRequisition As String, ByVal NewMaterialRequisition As String) As String
    Dim strResults As String
    
    strResults = OldMaterialRequisition
    
    If InStr(strResults, NewMaterialRequisition) Then
    Else
        strResults = strResults & " ; " & NewMaterialRequisition
    End If
    CheckStrMaterialRequisition = strResults
    
End Function





