Attribute VB_Name = "modLine"
Option Explicit

Public UserLine As String
Public UserLineIndex As Integer
Public bCODLine As Boolean


Public Function GetUserLine()


UserLine = GetSetting(App.Title, "Stazione", "UserLine", "")
UserLineIndex = GetSetting(App.Title, "Stazione", "UserLineIndex", 0)

bCODLine = IIf(InStr(UserLine, "59"), True, False)
If Not bCODLine Then
    bCODLine = IIf(InStr(UserLine, "COD"), True, False)
End If
End Function
Public Function SetUserLine(ByVal Line As String, ByVal Index As Integer)

If Line <> "" Then
    SaveSetting App.Title, "Stazione", "UserLine", Line
    SaveSetting App.Title, "Stazione", "UserLineIndex", Index
    
    UserLine = Line
    UserLineIndex = Index
    
End If

bCODLine = IIf(InStr(UserLine, "59"), True, False)
If Not bCODLine Then
    bCODLine = IIf(InStr(UserLine, "COD"), True, False)
End If

End Function

