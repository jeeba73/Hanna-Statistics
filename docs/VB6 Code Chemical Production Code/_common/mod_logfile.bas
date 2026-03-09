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
   
    If strAccesso = "" Then strAccesso = "Accesso senza Password"

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
NomeMese(1) = "Gennaio"
NomeMese(2) = "Febbraio"
NomeMese(3) = "Marzo"
NomeMese(4) = "Aprile"
NomeMese(5) = "Maggio"
NomeMese(6) = "Giugno"
NomeMese(7) = "Luglio"
NomeMese(8) = "Agosto"
NomeMese(9) = "Settembre"
NomeMese(10) = "Ottobre"
NomeMese(11) = "Novembre"
NomeMese(12) = "Dicembre"

i = Month(Date)

GetNomeMese = year(Date) & " - Accessi " & NomeMese(i)

End Function



Public Function CreateClassificationLogFile(ByVal strAccesso As String) As Boolean
Dim rc As Boolean
Dim sString As String

  '---------------------------------------------
  ' Aggiorno il file Log in MY_LOG_PATH
  '---------------------------------------------

    On Error GoTo DownloadError
    DoEvents
    rc = True
     
   ' If FileExists(MY_LOG_PATH & GetNomeMese) = False Then CreateDataLog
   
    If strAccesso = "" Then strAccesso = "Accesso senza Password"

    sString = Date & " - " & Time & " : " & strAccesso
    
    Open USER_DOCUMENTI & "\" & "Classification View - " & GetNomeMese & ".txt" For Append As #1
        Print #1, sString
    Close #1
ERR_END:
    On Error GoTo 0
    CreateClassificationLogFile = rc
    Exit Function
DownloadError:
    
    rc = False
    Resume ERR_END
End Function
