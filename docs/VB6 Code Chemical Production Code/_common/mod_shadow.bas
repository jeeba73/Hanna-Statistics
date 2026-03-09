Attribute VB_Name = "mod_shadow"
Private Declare Function GetClassLong Lib "user32.dll" Alias "GetClassLongA" _
  (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetClassLong Lib "user32.dll" Alias "SetClassLongA" _
  (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GCL_STYLE = -26
Private Const CS_DROPSHADOW = &H20000

Public Sub DropShadow(hWnd As Long)
    SetClassLong hWnd, GCL_STYLE, GetClassLong(hWnd, GCL_STYLE) Or CS_DROPSHADOW
End Sub

