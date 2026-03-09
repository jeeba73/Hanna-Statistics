Attribute VB_Name = "AddInStock_01_Grid"
Option Explicit

Private Grid() As Grid

Public Function SetAllRecipeForProductionGrid(ByVal Grid As Variant) As Boolean



' 0 code


    
   ' Call SetCodeGrid(Grid(0))
    Call SetStockTable(Grid(0))

End Function


Public Sub SetCodeGrid(ByVal Grd As Grid)
Dim i As Integer
        '------------------------------------------------
        '      RecipeForProduction  TABELLA Codici
        '------------------------------------------------
    With Grd
      
      
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click

        .SelectionMode = cellSelectionNone
        
        .DefaultRowHeight = 25
        
        
        .Cols = 9
        
        .RowHeight(0) = 35
        
        .ReadOnly = False
        .DefaultFont.Size = 11 ' * m_ControlGridFontSize
        '.DefaultFont.Bold = False
        
        .DefaultFont.Name = "Calibri"
        
        
        .Cell(0, 1).Text = "Code"
        .Cell(0, 2).Text = "Product Name"
        .Cell(0, 3).Text = "Line"
        .Cell(0, 4).Text = "Volume/Weight"
        .Cell(0, 5).Text = "(um)"
        .Cell(0, 6).Text = "Q.ty to produce"
        .Cell(0, 7).Text = "MRCode"
        .Cell(0, 8).Text = "Mix"

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
            .Column(i).Width = 150
            .Cell(0, i).FontBold = True
            If i > 7 Then .Column(i).Width = 100
        Next
        .Column(1).Width = 100
        .Column(3).Width = 100
        .Column(2).Width = 250
        .Column(4).Width = 120
        .Column(5).Width = 80
        .Column(7).Width = 200
        .Column(8).Width = 400
        .Column(4).Alignment = cellRightCenter
        .Column(8).Alignment = cellCenterCenter
        .BoldFixedCell = True
        
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
   

End Sub



Public Function SetChemicalMRStockQTY(ByVal ChemicalMR As String, ByVal StockQTY As Double, ByVal stockUnit As String, ByRef MR() As MRType) As Boolean
Dim i As Integer
Dim NewCode As Integer
For i = 0 To UBound(MR)

    If MR(i).Code = Trim(ChemicalMR) Then
        MR(i).STOCK_QTY = MR(i).STOCK_QTY + StockQTY
        MR(i).STOCK_UNIT = stockUnit

        GoTo cont:
    End If
Next
' non l'ho trovato....
NewCode = UBound(MR) + 1
ReDim Preserve MR(NewCode)

 MR(NewCode).Code = Trim(ChemicalMR)
 MR(NewCode).STOCK_QTY = MR(NewCode).STOCK_QTY + StockQTY
 
 MR(NewCode).STOCK_UNIT = stockUnit


cont:

End Function

Public Function UpdateStockQTY_MRDatabase(ByRef MR() As MRType) As Boolean
Dim i As Integer
Dim t As Integer
Dim ChemicalMR As String

    With dbTabMR
        .filter = ""
        .filter = ""
        If .EOF Then Exit Function
        
        .MoveFirst
        For t = 1 To .RecordCount
            ChemicalMR = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
            For i = 0 To UBound(MR)
                If MR(i).Code = Trim(ChemicalMR) And ChemicalMR <> "" Then
                    !STOCK_QTY = MR(i).STOCK_QTY
                    !STOCK_UNIT = IIf(MR(i).STOCK_UNIT <> "", MR(i).STOCK_UNIT, !STOCK_UNIT)
                    
                    'If IsNull(!STOCK_UNIT) Or !STOCK_UNIT = "" Then
                        If Not (IsNull(!PhysicalState)) Or !PhysicalState <> "" Then
                            
                            If !PhysicalState = "L" Then
                                !STOCK_UNIT = "mL"
                            Else
                                !STOCK_UNIT = "g"
                                
                            End If

                        End If
                    'End If
                    
                    
                    Exit For
                End If
                
            Next
            .MoveNext
        Next
    End With


End Function




Public Function SetMR(ByVal MRCode As String, ByRef uMR As MRType) As Boolean
Dim rc As Boolean
    With dbTabMR
        .filter = ""
        .filter = "Code='" & MRCode & "'"
        
        If .EOF Then
            uMR = MRTypeClean
            rc = False
        Else
            rc = True
        
            uMR.Code = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
            uMR.Description = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
            uMR.Supplier = IIf(IsNull(Trim(!Supplier)), "", Trim(!Supplier))
            uMR.MNP = IIf(IsNull(Trim(!MNP)), "", Trim(!MNP))
            uMR.PhysicalState = IIf(IsNull(Trim(!PhysicalState)), "", Trim(!PhysicalState))
            uMR.Density = IIf(IsNull(Trim(!Density)), "", Trim(!Density))
            uMR.Unit = IIf(IsNull(Trim(!Unit)), "", Trim(!Unit))
            uMR.Parameter = IIf(IsNull(Trim(!Parameter)), "", Trim(!Parameter))
            uMR.FWParameter = IIf(IsNull(Trim(!FWParameter)) Or Trim(!FWParameter) = "", 0, Trim(!FWParameter))
            uMR.StorageT = IIf(IsNull(Trim(!StorageT)), "", Trim(!StorageT))
            uMR.MinQty = IIf(IsNull(Trim(!MinQty)) Or Trim(!MinQty) = "", 0, Trim(!MinQty))
            uMR.STOCK_QTY = IIf(IsNull(Trim(!STOCK_QTY)) Or Trim(!STOCK_QTY) = "", 0, Trim(!STOCK_QTY))
            uMR.STOCK_UNIT = IIf(IsNull(Trim(!STOCK_UNIT)), "", Trim(!STOCK_UNIT))
            uMR.ReductionExpDays = IIf(IsNull(Trim(!ReductionExpDays)), 120, Trim(!ReductionExpDays))
            If Trim(uMR.ReductionExpDays) = "" Then uMR.ReductionExpDays = 120
            uMR.Classification = IIf(IsNull(Trim(!Classification)), "", Trim(!Classification))
            uMR.bMassa = IIf(InStr(UCase(uMR.PhysicalState), "L"), False, True)
            uMR.Location = IIf(IsNull(Trim(!Location)), "", Trim(!Location))
            
            If uMR.Location = "" Then
                 uMR.Location = GetLocation(uMR.Code)
            End If
        End If
    
    End With
    SetMR = rc
    
End Function

Public Function GetLocation(ByVal Code As String) As String

If Code = "" Then Exit Function

With dbTabMRWarehouse
    .filter = ""
    .filter = "Code='" & Code & "'"
    If .EOF Then
    Else
        .MoveLast
        GetLocation = IIf(IsNull(Trim(!Location)), "", Trim(!Location))
    End If
End With
End Function


Public Function GetStockQTY(ByVal Code As String, ByRef stockUnit As String) As Double
Dim i As Integer
Dim SotckQTY As Double


On Error GoTo ERR_ADD

    If Code = "" Then
        GetStockQTY = 0
        Exit Function
    End If
    
    
    With dbTabMRWarehouse
        .filter = ""
        .filter = "Code='" & Code & "' and bClosed=False"
        If .EOF Then
        Else
            .MoveFirst
            For i = 1 To .RecordCount
            
                If Not (IsNull(!MREXP)) Then
                    If CDate(Trim(!MREXP)) < FormatDateTime(Now, vbShortDate) Then
                        GoTo cont:
                    End If
                End If
                
                If Not (IsNull(!SupplierEXP)) Then
                    If CDate(Trim(!SupplierEXP)) < FormatDateTime(Now, vbShortDate) Then
                        GoTo cont:
                    End If
                End If
                                
                If Not (IsNull(!Finished)) Then
                    GoTo cont:
                End If
                                    
                SotckQTY = SotckQTY + CheckDot(IIf(IsNull(Trim(!StockQTY)), 0, Trim(!StockQTY)))
                stockUnit = IIf(IsNull(Trim(!stockUnit)), "mL", Trim(!stockUnit))
cont:
                .MoveNext
            Next
            If stockUnit = "mL" Then
                If (SotckQTY / 1000) > 1 Then
                    
                    SotckQTY = FormatNumber(SotckQTY / 1000, 3)
                    stockUnit = "L"
                End If
            End If
            
        End If
    End With
ERR_END:
    On Error GoTo 0
    SotckQTY = FormatNumber(SotckQTY, 3)
    GetStockQTY = SotckQTY
    Exit Function
ERR_ADD:
    MsgBox Err.Description
    Resume Next
End Function

