Attribute VB_Name = "GetRMby"
Option Explicit


Public Function GetMRbyID(ByVal ID As Long) As String
    With dbTabMR
        .Filter = ""
        .Filter = "ID='" & ID & "'"
        If .EOF Then
            GetMRbyID = "Not Found"
        Else
            GetMRbyID = Trim(!Code)
        End If
    End With
End Function


Public Function GetIDMR(ByVal Code As String) As Long
If Code = "" Then Exit Function
With dbTabMR
    .Filter = ""
    .Filter = "Code='" & Code & "'"
    If .EOF Then
    Else
        GetIDMR = !ID
    End If

End With
End Function


