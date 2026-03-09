Attribute VB_Name = "STDCalculation"
Option Explicit



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


