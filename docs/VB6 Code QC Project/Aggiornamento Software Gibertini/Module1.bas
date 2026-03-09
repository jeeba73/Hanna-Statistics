Attribute VB_Name = "mod_apriechiudifiles"
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function OpenProcess Lib "kernel32" _
  (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
   ByVal dwProcessId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
   (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" _
   (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
  
Private RetVal As Long

Public Function ApriEseguibile(ByVal vFile As String)
    'apre la calcolatrice
    RetVal = Shell(vFile, 1)
End Function

Public Function ChiudiEseguibile(ByVal vFile As String)
    'Chiude calcolatrice
    Dim hP As Long
    Dim lExC As Long

    hP = OpenProcess(PROCESS_ALL_ACCESS, 0&, RetVal)
    If hP Then
        GetExitCodeProcess hP, lExC
        If lExC Then TerminateProcess hP, lExC
    End If
End Function




Public Function CreateVerFile() As Boolean
Dim rc As Boolean
Dim sString As String

  '-------------------------------------------
  ' Aggiorno il file versione in \update
  '-------------------------------------------

    On Error GoTo DownloadError
    DoEvents
    rc = True
     
    Kill App.Path & "\update\" & "nver.txt"

    sString = App.Major & "." & App.Minor & "." & App.Revision

    Open App.Path & "\update\" & "nver.txt" For Append As #1
        Print #1, sString
    Close #1
ERR_END:
    On Error GoTo 0
    CreateVerFile = rc
    Exit Function
DownloadError:
    MsgBox Err.Description
    rc = False
    Resume ERR_END
End Function


