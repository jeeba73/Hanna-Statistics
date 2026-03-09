Attribute VB_Name = "mod_logfile"
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function OpenProcess Lib "kernel32" _
  (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
   ByVal dwProcessId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
   (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" _
   (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
  
Private RetVal As Long


Public Function CreateLogFile(ByVal strAccesso As String) As Boolean
Dim rc As Boolean
Dim sString As String

  '---------------------------------------------
  ' Aggiorno il file Log in MY_LOG_PATH
  '---------------------------------------------

    On Error GoTo DownloadError
    DoEvents
    rc = True
     
   ' If FileExists(MY_LOG_PATH & GetNomeMese) = False Then CreateDataLog
   
    If strAccesso = "" Then strAccesso = Trnslate("Accesso senza Password")

    sString = Date & " - " & Time & " : " & strAccesso
    
    Open MY_LOG_PATH & "\" & GetNomeMese & ".txt" For Append As #1
        Print #1, sString
    Close #1
ERR_END:
    On Error GoTo 0
    CreateLogFile = rc
    Exit Function
DownloadError:
    
    rc = False
    Resume ERR_END
End Function




Public Function GetNomeMese() As String

Dim NomeMese(12) As String
Dim i As Integer
NomeMese(1) = Trnslate("Gennaio")
NomeMese(2) = Trnslate("Febbraio")
NomeMese(3) = Trnslate("Marzo")
NomeMese(4) = Trnslate("Aprile")
NomeMese(5) = Trnslate("Maggio")
NomeMese(6) = Trnslate("Giugno")
NomeMese(7) = Trnslate("Luglio")
NomeMese(8) = Trnslate("Agosto")
NomeMese(9) = Trnslate("Settembre")
NomeMese(10) = Trnslate("Ottobre")
NomeMese(11) = Trnslate("Novembre")
NomeMese(12) = Trnslate("Dicembre")

i = Month(Date)

GetNomeMese = Year(Date) & Trnslate(" - Accessi ") & NomeMese(i)

End Function



