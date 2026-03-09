Attribute VB_Name = "Database_Warehouse_Entry"
Option Explicit


Public Function GetDatabaseWareHouseEntry(ByVal ID As Long, ByVal i As Integer, ByVal bSearch As Boolean, ByRef iWarehouse() As WareHouseEntry) As Boolean

Dim PurityPerc As Double
Dim MNP As String
With dbTabMRWarehouse
     ReDim Preserve iWarehouse(i)
     
     If bSearch Then
                
            .filter = ""
            .filter = "ID='" & ID & "'"
            
            If .EOF Then
            
                GetDatabaseWareHouseEntry = False
                Exit Function
            Else
                GetDatabaseWareHouseEntry = True
               
            End If
     End If
            
            
        PurityPerc = 100
        If IsNumeric(IIf(IsNull(Trim(!Purity)), "", Trim(!Purity))) Then
            PurityPerc = CheckDot(Trim(!Purity))
            If CheckDot(Trim(!Purity)) < 1 Then
                PurityPerc = CheckDot(Trim(!Purity)) * 100
            End If
        End If
         iWarehouse(i).ArrivedTime = FormatDataLAT(IIf(IsNull(Trim(!ArrivedTime)), "", Trim(!ArrivedTime)))
         iWarehouse(i).EntryBottle = IIf(IsNull(Trim(!Bottle)), "", Trim(!Bottle))
         iWarehouse(i).Density = CheckDot(IIf(IsNull(Trim(!Density)), 1, Trim(!Density)))
         iWarehouse(i).Description = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
         iWarehouse(i).Finished = FormatDataLAT(IIf(IsNull(Trim(!Finished)), "", Trim(!Finished)))
         iWarehouse(i).FWParameter = IIf(IsNull(Trim(!FWParameter)) Or Trim(!FWParameter) = "", 0, Trim(!FWParameter))
         iWarehouse(i).ID = !ID
         iWarehouse(i).Location = IIf(IsNull(Trim(!Location)), "", Trim(!Location))
         iWarehouse(i).Lot = IIf(IsNull(Trim(!Lot)), "", Trim(!Lot))
         iWarehouse(i).MRCode = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
         iWarehouse(i).MREXP = IIf(IsNull(Trim(!MREXP)), "", Trim(!MREXP))
         
  
         Call GetSupplier(iWarehouse(i).MRCode, MNP)
         iWarehouse(i).MNP = MNP
         
         
         iWarehouse(i).MRValueConcentration = (IIf(IsNull(Trim(!MRValue)), "", Trim(!MRValue)))
         iWarehouse(i).U = (IIf(IsNull(Trim(!U)), "", Trim(!U)))
         iWarehouse(i).Note = IIf(IsNull(Trim(!Note)), "", Trim(!Note))
         iWarehouse(i).NumberBottle = 1 ' IIf(IsNull(Trim(!Bottle)), "", Trim(!Bottle))
         iWarehouse(i).Open = FormatDataLAT(IIf(IsNull(Trim(!Open)), "", Trim(!Open)))
         iWarehouse(i).Operator = IIf(IsNull(Trim(!Operator)), "", Trim(!Operator))
         iWarehouse(i).Parameter = IIf(IsNull(Trim(!Parameter)), "", Trim(!Parameter))
         iWarehouse(i).Purity = PurityPerc
         iWarehouse(i).Status = IIf(IsNull(Trim(!Status)), 0, Trim(!Status))
         iWarehouse(i).StockQTY = CheckDot(IIf(IsNull(Trim(!StockQTY)), "", Trim(!StockQTY)))
         iWarehouse(i).stockUnit = IIf(IsNull(Trim(!stockUnit)), "", Trim(!stockUnit))
         iWarehouse(i).SupplierEXP = IIf(IsNull(Trim(!SupplierEXP)), "", Trim(!SupplierEXP))
         iWarehouse(i).Unit = IIf(IsNull(Trim(!Unit)), "", Trim(!Unit))
         
End With
            
   End Function
