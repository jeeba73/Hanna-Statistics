Attribute VB_Name = "mod_grafico"
Option Explicit


Public cYMax
Public coef_graph
Public coef_graph_2
Public Xmin
Public Xmax
Public Ydev
Public cDevSt(600, 10) As Double
Public cStep
Public Accettazione As Boolean
Public Chiudi As Boolean
Public Function Controllo(pezzi)
    
    Select Case pezzi
        Case 0 To 100
    End Select
End Function

Function Controllo_Media(media As Double, Scarto As Double) As Boolean
Dim cMAX_MEDIA
On Error GoTo err:

Controllo_Media = True
'cMAX_MEDIA = cNominale - cCOST * scarto
If media >= cMAX_MEDIA Then
    ElseIf media < cMAX_MEDIA Then
        Controllo_Media = False
End If

exit_sub:
Exit Function
err:
Resume exit_sub
End Function
