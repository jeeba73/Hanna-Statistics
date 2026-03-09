Attribute VB_Name = "Acquisition"
Option Explicit
Public Const TolerancePerc As Double = 0.01   ' 1%

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

 'dbChemicalMR.Execute "DELETE * FROM TabAcquisition where ID=" & AcquisitionID
   
    
End Function
