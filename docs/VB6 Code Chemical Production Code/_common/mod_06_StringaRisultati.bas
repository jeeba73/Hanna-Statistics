Attribute VB_Name = "mod_06_StringaRisultati"
Option Explicit

Public Function SetFormattazioneRisultatiLAT(ByVal strRisultato As String) As String
Dim sString As String
Dim sStringDecimali As String
Dim i As Integer
Dim sLenght As Integer
Dim GruppiDiTre As Integer
Dim MyArray() As String
Dim MyUM As String



strRisultato = Replace(strRisultato, ".", "")

MyUM = GetUM(strRisultato)

strRisultato = Trim(Replace(strRisultato, MyUM, ""))

MyArray = Split(strRisultato, ",")

If UBound(MyArray) = 1 Then
' ho un risultato con la virgola....
' considero la prima parte....

  sLenght = Len(MyArray(1))
   
   sStringDecimali = MyArray(1)
    
   GruppiDiTre = Int(sLenght / 3)
   
    If GruppiDiTre > 0 Then
        For i = 1 To GruppiDiTre
    
            sStringDecimali = Trim(mID(sStringDecimali, 1, (3 * i) + (i - 1)) & " " & Right(sStringDecimali, Len(sStringDecimali) - (3 * i) - (i - 1)))
        Next
    Else
        sStringDecimali = MyArray(1)
    End If
    
    
End If

   sLenght = Len(MyArray(0))
   
   sString = MyArray(0)
    
   GruppiDiTre = Int(sLenght / 3)
   
    If GruppiDiTre > 0 Then
        For i = 1 To GruppiDiTre
    
            sString = Trim(mID(MyArray(0), 1, Len(sString) - (3 * i) - (i - 1)) & " " & Right(sString, (3 * i) + (i - 1)))
        Next
    Else
        sString = MyArray(0)
    End If
    
    
    SetFormattazioneRisultatiLAT = Trim(sString & IIf(sStringDecimali <> "", "," & sStringDecimali, "") & " " & MyUM)


End Function


Private Function GetUM(ByVal strRisultato As String) As String

If InStr(LCase(strRisultato), "ng") Then
    
    GetUM = "ng"

ElseIf InStr(LCase(strRisultato), "mg") Then
    
    GetUM = "mg"
    
ElseIf InStr(LCase(strRisultato), "kg") Then
    
    GetUM = "kg"


ElseIf InStr(LCase(strRisultato), "g") Then
    
    GetUM = "g"

End If

End Function
