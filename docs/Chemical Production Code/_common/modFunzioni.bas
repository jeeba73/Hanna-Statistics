Attribute VB_Name = "modFunzioni"
Option Explicit
Public Const VK_TAB = &H9
         
Public Declare Sub keybd_event Lib "user32" _
  (ByVal bVk As Byte, _
   ByVal bScan As Byte, _
   ByVal dwFlags As Long, _
   ByVal dwExtraInfo As Long)
   

Private Type Societa
    Enabled As Boolean
    Department As String
    Description As String
    Workstation As String
    LineLeader As String
    email As String
    ServerUserID As String
    ServerPassword As String
    ServerFTP As String
    ServerWorkPath As String
End Type




Public MyWorkStation As Societa
'Public MyFormula As Formula
'Public MyFormulaClean As Formula


Public Function FindWorkStation() As Boolean
    Dim rc As Boolean
    On Error GoTo ERR_FIND
    rc = True
    With dbTabLaboratorio
        .filter = ""
        rc = Not (.EOF)
    End With
ERR_END:
    On Error GoTo 0
    FindWorkStation = rc
    Exit Function
ERR_FIND:
    rc = False
    Resume ERR_END
End Function


Public Function TxtToNumber(KeyAscii As Integer) As Integer
    '------------------------------------------------------------
    '   restituisce solo virgole e controlla se è un numero
    '------------------------------------------------------------
    Select Case KeyAscii
        Case 8
            GoTo NUMBER
        Case 13
            GoTo NUMBER
        Case 46
          '  KeyAscii = 0
           ' SendKeys (",")
    End Select
    If Not IsNumeric(Chr$(KeyAscii)) And Not (KeyAscii = 44 Or KeyAscii = 46 Or KeyAscii = 45) Then
        KeyAscii = 0
    End If
 
NUMBER:
    
    TxtToNumber = KeyAscii
End Function
