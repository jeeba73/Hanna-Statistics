Attribute VB_Name = "mod_Aggiornamento"
Option Explicit

Public StgVersione As String
Public StgVersioneftp As String

Public Function GetVerSoft(ByVal vFile As String) As String
  '-------------------------------------------
  ' va a prendere la versione attuale
  '-------------------------------------------
    Dim NF As Integer
    Dim sString As String
    
    
    On Error GoTo ERR_GET
    
    NF = FreeFile()
    Open USER_UPDATE_PATH & vFile For Input As NF
      Line Input #1, sString   ' Assegna la riga a una variabile.
    Close NF
ERR_END:
    On Error GoTo 0
    GetVerSoft = sString
    Exit Function
ERR_GET:
    
    Resume ERR_END
    
  End Function
  
  
Public Function EsistonoAggiornamenti(sVerLoc As String, sVerWeb As String) As Boolean
 On Error Resume Next
    Dim vLoc, vWeb, i As Integer
    Dim a, b As Integer
    If sVerLoc = "" Then
        EsistonoAggiornamenti = False
    Else
        vLoc = Split(sVerLoc, ".")
        vWeb = Split(sVerWeb, ".")
        For i = 0 To 2
            a = CInt(vLoc(i))
            a = IIf(i = 2 And a < 10, a * 10, a)
           
            b = CInt(vWeb(i))
             b = IIf(i = 2 And b < 10, b * 10, b)
          If a < b Then
            EsistonoAggiornamenti = True
            Exit For
            ElseIf a > b Then
               ' MsgBox "Attenzione. Si possiede una versione più aggiornata di quella che si sta cercando di scaricare", vbCritical, "Smart Update"
                Exit Function
          End If
        Next
    End If
    On Error GoTo 0
  End Function

