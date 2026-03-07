Attribute VB_Name = "FTP"
Option Explicit
Public sFTPServer As String
Public sFTPUser As String
Public sFTPPass As String
Public bFTPBinary As Boolean
Public bFTPConn As Boolean
Public bFTPInvioAutomatico As Boolean
Public bFTPCancellafile As Boolean
Public sFTPpathOrigin As String
Public sFTPPathServer As String
Public bSospendiFTP As Boolean
Public bFTP As Boolean
Public TestTime As Double
Public bStampanteOK As Boolean
Public bFTPOK As Boolean
Public bInvioVerifiche As Boolean
Public bInvioRichieste As Boolean


Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef _
lpSFlags As Long, ByVal dwReserved As Long) As Long
Const INTERNET_CONNECTION_MODEM = 1
Const INTERNET_CONNECTION_LAN = 2
Const INTERNET_CONNECTION_PROXY = 4
Const INTERNET_CONNECTION_MODEM_BUSY = 8
Dim flags As Long



Public Sub mWait()
   Screen.MousePointer = vbHourglass
End Sub

Public Sub mOk()
   Screen.MousePointer = vbDefault
End Sub



Public Function CheckInternetConn() As Boolean
Dim rc As Boolean
    rc = True
    If InternetGetConnectedState(flags, 0) = 0 Then
        rc = False
        'MsgBox "Attenzione. Il Computer Non è connesso ad internet." & vbCrLf & "Controllare la connessione e riavviare l'update", vbCritical
    ElseIf flags = INTERNET_CONNECTION_MODEM Then
        'MsgBox "Sei connesso con il Modem" ' connessione attiva via modem
    ElseIf flags = INTERNET_CONNECTION_LAN Then
        'MsgBox "Sei connesso con LAN" ' connessione attiva via LAN
    ElseIf flags = INTERNET_CONNECTION_PROXY Then
        'MsgBox "Sei connesso via Proxy" ' connessione attiva via proxy
    ElseIf flags = 18 Then
        'MsgBox "Sei connesso via Wifi"
    End If
    CheckInternetConn = rc
End Function





Public Function connettiFTP(ByVal mFTP As cFTP) As Boolean
Dim rc As Boolean

   mWait
   rc = True
   On Error GoTo ERR_FTP
   
   If CheckInternetConn Then
 
        If mFTP.OpenConnection(sFTPServer, sFTPUser, sFTPPass) Then
          ' mFTP.SetFTPDirectory "/"
           'Me.Caption = "Connesso a " & sFTPServer & " come " & sFTPUser & " - " & Now
         Else
             rc = False
        End If
   Else
        rc = False
    End If
ERR_END:
   mOk
   connettiFTP = rc
   Exit Function
ERR_FTP:
   MsgBox err.Description
   rc = False
   Resume Next
End Function

Public Function fnFTP_MASTER(ByVal mFTP As cFTP, ByVal sFTPPathStazione As String, ByVal sStazione As String, Optional ByVal bUpload As Boolean = False) As Boolean
  
    Dim rc As Boolean
    Dim rcFTP As Boolean
    Dim i As Integer
    Dim t As Integer
    Dim RealFile As Integer
    
   mFTP.SetModePassive
   mFTP.SetTransferBinary

    
    If bFTPConn Then
        mFTP.SetModeActive
    Else
        mFTP.SetModePassive
    End If
    If bFTPBinary Then
        mFTP.SetTransferBinary
    Else
        mFTP.SetTransferASCII
    End If
    
   
   mWait
    rc = connettiFTP(mFTP)
    rcFTP = rc
   
  
    
    If rc Then
    
        If bUpload Then
            RealFile = EnumDir(sFTPpathOrigin)
          
            If RealFile = 0 Then
                fnFTP_MASTER = False
                Exit Function
            End If
            
             
            
            If UploadFTP(sFTPpathOrigin, sFTPPathStazione & "/MASTER", i, mFTP) Then
            Else
               rcFTP = False
                PopupMessage 0, ("Attenzione : I files non sono stati caricati sul Server"), , True, ("Trasferimento file")
                
            End If
             mFTP.CloseConnection
        Else
            If DownloadFTP(sFTPpathOrigin, sFTPPathStazione, i, mFTP) Then
                If i > 1 Then PopupMessage 1, "Scaricati " & i & " files", sStazione
            Else
               rcFTP = False
                PopupMessage 1, ("Attenzione : I files non sono stati scaricati sul Server"), , True, ("Trasferimento file")
            End If
             mFTP.CloseConnection
        End If
        
    'End If
    Else
       ' 'MessageInfoTime = 3000
       PopupMessage 2, ("Connessione Server non riuscita. Verificare in Opzioni le impostazioni del server..."), , True, ("Trasferimento file")
       bFTPOK = False
    End If
    mOk
    fnFTP_MASTER = rcFTP
    

End Function
Public Function fnFTP(ByVal mFTP As cFTP, Optional ByVal bUpload As Boolean = True) As Boolean
  
    Dim rc As Boolean
    Dim rcFTP As Boolean
    Dim i As Integer
    Dim t As Integer
    Dim RealFile As Integer
    
   mFTP.SetModePassive
   mFTP.SetTransferBinary

    
    If bFTPConn Then
        mFTP.SetModeActive
    Else
        mFTP.SetModePassive
    End If
    If bFTPBinary Then
        mFTP.SetTransferBinary
    Else
        mFTP.SetTransferASCII
    End If
    
   
   mWait
    rc = connettiFTP(mFTP)
    rcFTP = rc
    
    'MessageInfoTime = 3000
    If rc Then
    
        If bUpload Then
            RealFile = EnumDir(sFTPpathOrigin)
            If RealFile = 0 Then
                fnFTP = False
                Exit Function
            End If
            If UploadFTP(sFTPpathOrigin, sFTPPathServer, i, mFTP) Then

            Else
               rcFTP = False
                 PopupMessage 2, ("Attenzione : I files non sono stati scaricati sul Server"), , True, ("Trasferimento file")
            End If
             mFTP.CloseConnection
        Else
            If DownloadFTP(sFTPpathOrigin & "\MASTER\", sFTPPathServer & "/MASTER", i, mFTP) Then

                If i > 1 Then PopupMessage 1, ("Scaricati ") & i & (" files da Server")
            Else
                PopupMessage 2, ("Attenzione : I files non sono stati scaricati sul Server"), , True, ("Trasferimento file")
            End If
             mFTP.CloseConnection
        End If
        
    'End If
    Else
       PopupMessage 2, ("Connessione Server non riuscita. Verificare in Opzioni le impostazioni del server..."), , True, ("Trasferimento file")
    End If
    mOk
    fnFTP = rcFTP
    

End Function
Public Function fnFTPbyName(ByVal mFTP As cFTP, Optional ByVal bUpload As Boolean = True, Optional ByVal sPathServer As String, Optional sNomeFile As String, Optional NomeStazione As String) As Boolean
  
    Dim rc As Boolean
    Dim rcFTP As Boolean
    Dim i As Integer
    Dim t As Integer
    Dim RealFile As Integer
    
   mFTP.SetModePassive
   mFTP.SetTransferBinary

    
    If bFTPConn Then
        mFTP.SetModeActive
    Else
        mFTP.SetModePassive
    End If
    If bFTPBinary Then
        mFTP.SetTransferBinary
    Else
        mFTP.SetTransferASCII
    End If
    
   
   mWait
    rc = connettiFTP(mFTP)
    rcFTP = rc
    
    If rc Then
     
    Else
        PopupMessage 2, ("Connessione Server non riuscita. Verificare in Opzioni le impostazioni del server..."), , True, ("Trasferimento file")
    End If
    mOk
    fnFTPbyName = rcFTP
    

End Function

Private Function DownloadFTP(ByVal MyPath As String, ByVal myServer As String, Optional ByRef i As Integer = 0, Optional ByVal mFTP As cFTP) As Boolean
   Dim rc As Boolean
   Dim strRemote As String
   Dim strFile As String
   Dim pathOrigin As String
   Dim pathFile As String
    Dim Item As cDirItem

   On Error GoTo ERR_UPLOAD
   
   rc = True
    mWait
    pathOrigin = Trim(MyPath)

    i = 0

   mFTP.SetFTPDirectory myServer
   mFTP.GetDirectoryListing "*.*"

   For Each Item In mFTP.Directory
      
    If Item.Normal Then
       strFile = pathOrigin & Item.FileName
        pathFile = Trim(Item.FileName)


         If Item.FileName <> "" Then
             If Not mFTP.FTPDownloadFile(strFile, pathFile) Then
                 PopupMessage 1, mFTP.GetLastErrorMessage, , True
                 rc = False
               Else
                ' ok salvato, lo cancello
                
                    If mFTP.DeleteFTPFile(pathFile) Then
    
                    Else
                       PopupMessage 1, mFTP.GetLastErrorMessage, , True
                    End If
                   
                   
             End If
         End If
    i = i + 1
    End If
    
    Next
    mOk
ERR_END:
   On Error GoTo 0
   DownloadFTP = rc
   Exit Function
ERR_UPLOAD:
   rc = False
   MsgBox err.Description
   Resume Next
End Function



Public Function UploadFTP(ByVal MyPath As String, ByVal myServer As String, Optional ByRef i As Integer = 0, Optional ByVal mFTP As cFTP) As Boolean
   Dim rc As Boolean
   Dim strRemote As String
   Dim strFile As String
   Dim pathOrigin As String
   Dim pathFile As String
   Dim PathDaFile As String
   Dim ID_TABLAVORI_MASTER As Long
   On Error GoTo ERR_UPLOAD
   
   rc = True
    mWait
    pathOrigin = Trim(MyPath)

    strFile = Dir(pathOrigin)
    

    i = 0

    

    Do While strFile > ""
       
        strFile = Dir()
        i = i + 1
    Loop

    
    'pb.Visible = False
    mOk
   ' pb.Visible = False

    
ERR_END:
   On Error GoTo 0
   UploadFTP = rc
   Exit Function
ERR_UPLOAD:
   rc = False
   MsgBox err.Description
   Resume Next
End Function
