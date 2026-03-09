Attribute VB_Name = "mod_Inet"
Option Explicit
Private ITC As Inet

Public Function GetFileFromUrl(ByVal Inet1 As Inet, ByRef url As String, ByRef vFile As String) As Boolean
Dim fileBytes() As Byte
Dim fileNum As Integer
Dim rc As Boolean
Dim a
  '-------------------------------------------
  ' scarica quella da ftp
  '-------------------------------------------
    Do Until ITCReady(Inet1, False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    
    If FileExists(USER_UPDATE_PATH & vFile) Then Kill USER_UPDATE_PATH & vFile
    
    
   'MsgBox USER_UPDATE_PATH
    On Error GoTo DownloadError
    DoEvents
    rc = True
    fileBytes() = Inet1.OpenURL(url & vFile, icByteArray)

  
    fileNum = FreeFile
    Open USER_UPDATE_PATH & vFile For Binary Access Write As #fileNum
    Put #fileNum, , fileBytes()
    Close #fileNum
ERR_END:
    On Error GoTo 0
   ' MsgBox url & vFile
    GetFileFromUrl = rc
    Exit Function
DownloadError:
    MsgBox err.Description
    rc = False
    Resume ERR_END
End Function



Public Function UploadMyInfo(ByVal Inet1 As Inet)
Dim UpdateFileString As String
Dim UploadedFileString As String

'On Error Resume Next
    With Inet1
        .url = "ftp://ftp.bilsoft.it"
        .UserName = "9620198@aruba.it"
        .Password = "qjE4Kb7NGhUF"
    End With
    
    
    SetWorkStation
     
    UploadedFileString = FormatNomeFile(MyWorkStation.Department & MyWorkStation.Description)
    
    If UploadedFileString = "" Then
        UploadedFileString = WorkstationID & ".txt"
    Else
        UploadedFileString = UploadedFileString & ".txt"
    End If
     
    Call CreateUpdteFile(UploadedFileString)
    
    UpdateFileString = PC_DOCUMENTI & UploadedFileString
    
    Call PutFileinUrl(Inet1, UploadedFileString, UpdateFileString)
    
End Function


Private Function PutFileinUrl(ByVal Inet1 As Inet, ByRef url As String, ByRef vFile As String) As Boolean
Dim fileBytes() As Byte
Dim LocalPath As String
Dim RemotePath As String
Dim fileNum As Integer
Dim rc As Boolean
Dim a
  '-------------------------------------------
  ' upload file stazione
  '-------------------------------------------


    On Error GoTo DownloadError
    DoEvents
    rc = True

  LocalPath = vFile
  RemotePath = url
  
  Debug.Print vFile
  Debug.Print url

  With Inet1
    .Execute , "CD " & Chr(34) & "bilsoft.it/Download/" & PROGRAM_NAME & "/User/" & Chr(34)
    
    Do Until ITCReady(Inet1, False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    
  .Execute , "PUT " & Chr(34) & LocalPath & Chr(34) & " " & Chr(34) & RemotePath & Chr(34)

  End With
ERR_END:
    On Error GoTo 0
   ' MsgBox url & vFile
    PutFileinUrl = rc
    Exit Function
DownloadError:
   ' MsgBox err.Description
    rc = False
    Resume ERR_END
End Function
Private Function ITCReady(ByVal Inet1 As Inet, ShowMessage As Boolean)
'Check the state of itc, if it is not executing return true
If Inet1.StillExecuting Then
    ITCReady = False
    If ShowMessage Then
        MsgBox "Please wait.  FTP is still executing", vbInformation + vbOKOnly, "Busy"
    End If
Else
    ITCReady = True
End If
End Function
