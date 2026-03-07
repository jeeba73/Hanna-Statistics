Attribute VB_Name = "SortArray"
Option Explicit


Public Sub OrdinaArray(ByRef V() As String)
Dim i, j  As Integer
Dim temp As String
For i = 1 To UBound(V)
 For j = i + 1 To UBound(V)
  If V(i) > V(j) Then
  temp = V(i)
    V(i) = V(j)
    V(j) = temp
  End If
 Next j
Next i

End Sub
