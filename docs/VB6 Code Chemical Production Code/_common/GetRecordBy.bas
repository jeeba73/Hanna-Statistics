Attribute VB_Name = "GetRecordBy"
Option Explicit


Public Function GetChemicalRMbyID(ByVal ID As Long) As String
    With dbTabRawMaterial
        .filter = ""
        .filter = "ID='" & ID & "'"
        If .EOF Then
            GetChemicalRMbyID = "Not Found"
        Else
            GetChemicalRMbyID = Trim(!Code)
        End If
    End With
End Function


Public Function GetIDRowMaterial(ByVal Code As String) As Long
If Code = "" Then Exit Function
With dbTabRawMaterial
    .filter = ""
    .filter = "Code='" & Code & "'"
    If .EOF Then
    Else
        GetIDRowMaterial = !ID
    End If

End With
End Function

