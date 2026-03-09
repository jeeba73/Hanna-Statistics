Attribute VB_Name = "mod_Status"
Option Explicit


Public Function GetStatus(ByVal Index) As String

    Select Case Index
        Case 0
            GetStatus = "In Stock"
        Case 1
            GetStatus = "Opened"
        Case 2
            GetStatus = "Finished"
    
    End Select

End Function


Public Function IndexStatus(ByVal str As String) As Integer



    Select Case UCase(str)
        Case UCase("In Stock")
            IndexStatus = 0
        Case UCase("Opened")
            IndexStatus = 1
        Case UCase("Finished")
            IndexStatus = 2
        Case Else
            IndexStatus = 0
    End Select




End Function



Public Function SetCmbStatus(ByRef cmb As ComboBox)

With cmb
    .Clear
    .AddItem GetStatus(0)
    .AddItem GetStatus(1)
    .AddItem GetStatus(2)
    .ListIndex = 0
End With



End Function

