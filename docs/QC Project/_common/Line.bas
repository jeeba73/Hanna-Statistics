Attribute VB_Name = "Mod__Line"
Option Explicit


Public LineList() As String
Public bCheckLineFromDatabase As Boolean




Public UserLine As String
Public UserLineIndex As Integer



Public Function GetUserLine()


UserLine = GetSetting(App.Title, "Stazione", "UserLine", "")
UserLineIndex = GetSetting(App.Title, "Stazione", "UserLineIndex", 0)


End Function
Public Function SetUserLine(ByVal Line As String, ByVal Index As Integer)

If Line <> "" Then
    SaveSetting App.Title, "Stazione", "UserLine", Line
    SaveSetting App.Title, "Stazione", "UserLineIndex", Index
    
    UserLine = Line
    UserLineIndex = Index
    
End If

End Function




Public Function SetLine(ByVal cmb As ComboBox, Optional bAll As Boolean) As Boolean
Dim i As Integer
On Error GoTo ERR_SET:
    If bCheckLineFromDatabase = False Then GetAllLine
    With cmb
        .Clear
        If bAll Or UserLine = "All Lines" Then .AddItem "All Lines"
         For i = LBound(LineList) To UBound(LineList)
            .AddItem LineList(i)
         Next
        
        ' .ListIndex = 0
        ' If .ListCount = 0 Then Cmb.Visible = False
        ' If UserLineIndex <= .ListCount And bAll Then .ListIndex = UserLineIndex
        
        
        cmb = UserLine
        
        
    End With
    bCheckLineFromDatabase = True
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_SET:
   ' MsgBox err.Description
    GoTo ERR_END
End Function

Public Function GetAllLine() As Boolean
Dim i As Integer
Dim t As Integer
Dim strLine As String

    With dbTabCode
        .filter = ""
        If .EOF Then
        
        Else
            .MoveFirst
            t = 0
            For i = 1 To .RecordCount
            
                strLine = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
                If strLine = "" Then GoTo cont
                If t > 0 Then
                    If GetIndexArStrOneDim(LineList(), strLine) = -1 Then
Aggiungi:
                        
                       ReDim Preserve LineList(t)
                       LineList(t) = strLine
                       t = t + 1
            
                       
                    End If
                Else
                    GoTo Aggiungi
                End If
cont:
                .MoveNext
            Next
        
           
        End If
    
    End With
End Function




Public Function GetIndexArStrOneDim(AR() As String, ByVal Val As String) As Long
    Dim i As Long, ei As Long
    
    GetIndexArStrOneDim = -1
 
    On Error Resume Next
        ei = UBound(AR)
        If err.NUMBER <> 0 Then Exit Function
     On Error GoTo 0
 
    For i = 0 To ei
        If UCase(AR(i)) = UCase(Val) Then GetIndexArStrOneDim = i: Exit For
    Next
 
End Function

