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
    LaboratoryManager As String
    Phone As String
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


Public Function StandardCal(ByVal sValue As String, ByVal Fixed As Double, ByVal AndOr As String, ByVal Perc As Double, ByVal Restr As Double, ByVal sDecimal As String, ByRef Min As String, ByRef Max As String) As Boolean
Dim rc As Boolean
Dim Value As Double
Dim MenoValue As Double
Dim PiùValue As Double
Dim Index As Integer
Dim MenoValueOr As Double
Dim PiùValueOr As Double

Dim sRisMeno As String
Dim sRisPiù As String

    On Error GoTo ERR_CAL
    rc = True
    MenoValue = 0
    PiùValue = 0
    
    If sValue = "/" Or Not (IsNumeric(sValue)) Then
        
        sRisMeno = "/"
        sRisPiù = "/"
        Min = sValue
        Max = sValue
        GoTo ERR_END
    
    End If

    Value = CDbl(sValue)
    
    
    Select Case UCase(AndOr)
        Case "&"
            Index = 0
        Case UCase("or")
            Index = 1
        Case Else
            Index = 2
    End Select
    
    
    Select Case Index
        Case 0 ' AND
            MenoValue = Value - (Fixed) - (Value * Perc * Restr)
            PiùValue = Value + (Fixed) + (Value * Perc * Restr)

            If MenoValue < 0 Then MenoValue = 0
        Case 1 ' OR
        
            MenoValue = Value - (Fixed * Restr)
            PiùValue = Value + (Fixed * Restr)
            MenoValueOr = Value - (Value * Perc * Restr)
            PiùValueOr = Value + (Value * Perc * Restr)
            
            If MenoValue > MenoValueOr Then
                MenoValue = MenoValueOr
                PiùValue = PiùValueOr
            End If
            
            If MenoValue < 0 Then MenoValue = 0
            
        Case Else ' /
            MenoValue = Value
            PiùValue = Value
    End Select
    
    sRisMeno = Format$(MenoValue, sDecimal)
    sRisPiù = Format$(PiùValue, sDecimal)
    
ERR_END:
    On Error GoTo 0
    Min = sRisMeno
    Max = sRisPiù
    StandardCal = rc
    Exit Function
ERR_CAL:
    rc = False
    MsgBox err.Description
    Resume Next
End Function

