Attribute VB_Name = "Acquisition"
Option Explicit

Public Function DeleteRowInTabAcquisition(ByVal AcquisitionID As Long)
With dbTabAcquisition
    .filter = ""
    .filter = "ID='" & AcquisitionID & "'"
    If .EOF Then
    Else
        .Delete
        .Update
    End If
End With
End Function


Public Function DeleteRowInTabProductionAcquisition(ByVal ProductionID As Long)
With dbTabProdHistory
    .filter = ""
    .filter = "ID='" & ProductionID & "'"
    If .EOF Then
    Else
        .Delete
        .Update
    End If
End With
End Function



Public Function CheckQtyProducedAcquisition(ByVal ProductionID As Long, ByVal HannaCode As String) As Double
Dim i As Integer
With dbTabProdHistory
    .filter = ""
    .filter = "ProductionID='" & ProductionID & "' and Code='" & HannaCode & "'"
    If .EOF Then
        CheckQtyProducedAcquisition = 0
    Else
        .MoveFirst
        For i = 1 To .RecordCount
        
            CheckQtyProducedAcquisition = CheckQtyProducedAcquisition + CDbl(IIf(IsNull(Trim(!QtyProduced)), 0, Trim(!QtyProduced)))
            .MoveNext
        Next
        
    End If
End With
End Function



Public Function GetAcquisitionCount(ByVal ProductionID As Long)
With dbTabProdHistory
    .filter = ""
    .filter = "ID='" & ProductionID & "'"
    If .EOF Then
        GetAcquisitionCount = 0
    Else
         GetAcquisitionCount = .RecordCount
    End If
End With
End Function




