Attribute VB_Name = "frAddInStock_01_Functions"
Option Explicit

Public Function GetLastUMBottleLetter(ByRef iMRWarehouse As WareHouseEntry) As Boolean
Dim LastLetter As String

LastLetter = ""

With iMRWarehouse
    If .MRCode <> "" Then
        
        With dbTabMRWarehouse
            .Close
            .Open "SELECT *  FROM TabMRWarehouse order by Bottle", dbChemicalMR, adOpenKeyset, adLockOptimistic, adCmdText
   
            .filter = ""
            .filter = "Lot='" & iMRWarehouse.Lot & "'"
            If .EOF Then
            Else
                .MoveLast
                LastLetter = IIf(IsNull(Trim(!Bottle)) Or Trim(!Bottle) = "", "", Trim(!Bottle))
            End If
    
        End With
    End If
    
    .LastLetter = LastLetter

End With



End Function



Public Function GetLetter(ByRef LastLetter As String) As String

Dim NUMBER As String
Dim Letter As String

If LastLetter = "" Then

    GetLetter = "0A"

Else
    NUMBER = Left(LastLetter, 1)
    Letter = Right(LastLetter, 1)
    
    If Letter < "Z" Then
    
        Letter = Chr(Asc(Letter) + 1)
    
    ElseIf Letter = "Z" Then
    
        NUMBER = NUMBER + 1
        Letter = "A"
    
    
    End If
    

    GetLetter = NUMBER & Letter

End If




End Function


Public Function SaveWarehouseEntryInDatabase(ByRef iBottle As WareHouseEntry, ByVal bModify As Boolean, Optional ByVal bPreparation As Boolean, Optional ByVal ID As Long) As Boolean

Dim rc As Boolean

Dim NumBottle As Integer
Dim LastLetter As String
Dim NewLetter   As String
Dim i As Integer
Dim strBottle As String
rc = True
On Error GoTo ERR_SAVE:

With iBottle


    If bModify Then
        NumBottle = 1
    Else
        NumBottle = .NumberBottle
    End If
    
    
    If NumBottle > 0 And .MRCode <> "" Then


        For i = 0 To NumBottle - 1
        
            If bModify Then
                strBottle = iBottle.EntryBottle
            Else
                strBottle = iBottle.Bottle(i)
            End If
        
            With dbTabMRWarehouse
            
                    .filter = ""
                    
                    If ID > 0 Then
                        .filter = "ID='" & ID & "'"
                    Else
                        .filter = "Code='" & iBottle.MRCode & "' and bottle='" & strBottle & "' and Lot='" & iBottle.Lot & "'"
                    End If
                
                If .EOF Then
                    .AddNew
                Else
                    If bModify And Not (bPreparation) Then
                        If F_MsgBox.DoShow("Modify Warehouse Stock Entry?", iBottle.MRCode & " | " & strBottle & " | " & iBottle.Lot) Then
                        Else
                            
                            Exit Function
                        
                        End If
                    
                    End If
                End If
                
                
                
                
                !StockQTY = FormatNumber(iBottle.StockQTY, 3)
                If IsDate(iBottle.Open) Then !Open = iBottle.Open
                If IsDate(iBottle.Finished) Then
                     !Finished = iBottle.Finished
                     If IsDate(!Open) Then
                     Else
                         !Open = iBottle.Finished
                     End If
                     !Status = 2
                End If
                     
                If ID > 0 Then GoTo cont:  ' resto non serve aggiornarlo....
                
                !Bottle = strBottle
              
                
                !Code = iBottle.MRCode
                !Description = iBottle.Description
                
                !Lot = iBottle.Lot
                !Density = iBottle.Density
                !Purity = IIf(iBottle.Purity > 1, iBottle.Purity, iBottle.Purity * 100)
                !MRValue = iBottle.MRValueConcentration
                !U = iBottle.U
                !Unit = iBottle.Unit
                !Parameter = iBottle.Parameter
                !FWParameter = iBottle.FWParameter
                !Location = iBottle.Location
               
                !stockUnit = iBottle.stockUnit
                !ArrivedTime = iBottle.ArrivedTime
                !Status = (iBottle.Status)
                !U = iBottle.U
               
                If IsDate(iBottle.MREXP) Then
               
                   ' iBottle.MREXP = CreateMRExp((iBottle.SupplierEXP), uMR.ReductionExpDays)
                    !MREXP = iBottle.MREXP
                    
                End If
                
                
                
                !SupplierEXP = iBottle.SupplierEXP
                
                !Note = iBottle.Note
                !Operator = iBottle.Operator
cont:
               ' !bBarcode = iBottle.bBarcode
                !bClosed = IIf(Trim(!Finished) <> "", True, False)
                
                .Update
                
            End With
     
        Next
        
    End If

End With

'----------------------------------------------------------------
' ho deciso che quando salvo un lotto aggiorno la purity dell'MR
'----------------------------------------------------------------




Call AddPurityMR(iBottle.MRCode, iBottle.Purity)


ERR_END:
    On Error GoTo 0
   ' dbTabMRWarehouse.Close
   ' dbTabMRWarehouse.Open "SELECT *  FROM TabMRWarehouse order by Code", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    SaveWarehouseEntryInDatabase = rc
    Exit Function
ERR_SAVE:
    rc = False
    MsgBox Err.Description
    Resume Next
    
End Function




Public Function DeleteWharehouseEntry(ByVal ID As Long, ByRef Grid1 As Grid) As Boolean

Dim rc As Boolean


Dim i As Integer
rc = True

If F_MsgBox.DoShow("Delete Entry?", "Warehouse") = False Then
    DeleteWharehouseEntry = False
    Exit Function
End If

On Error GoTo ERR_DEL:

    With dbTabMRWarehouse
        .filter = ""
        .filter = "ID='" & ID & "'"
        If .EOF Then
            rc = False
        Else
            .Delete
            .Update
        End If

    End With

    Grid1.ReadOnly = False
    Grid1.Selection.DeleteByRow
    Grid1.ReadOnly = True
    
    
ERR_END:
    On Error GoTo 0
    DeleteWharehouseEntry = rc
    Exit Function
ERR_DEL:
    rc = False
    MsgBox Err.Description
    Resume Next
End Function

