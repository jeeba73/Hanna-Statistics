Attribute VB_Name = "GetRecordBy"
Option Explicit





Public Function GetIDMR(ByVal Code As String) As Long
If Code = "" Then Exit Function
With dbTabMR
    .filter = ""
    .filter = "Code='" & Code & "'"
    If .EOF Then
    Else
        GetIDMR = !ID
    End If

End With
End Function

