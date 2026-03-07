Attribute VB_Name = "mod_internet"
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef _
lpSFlags As Long, ByVal dwReserved As Long) As Long
Const INTERNET_CONNECTION_MODEM = 1
Const INTERNET_CONNECTION_LAN = 2
Const INTERNET_CONNECTION_PROXY = 4
Const INTERNET_CONNECTION_MODEM_BUSY = 8
Dim flags As Long

Public Function CheckInternetConn() As Boolean
Dim rc As Boolean
    rc = True
    If InternetGetConnectedState(flags, 0) = 0 Then
        rc = False
        MsgBox "Attenzione. Il Computer Non è connesso ad internet." & vbCrLf & "Controllare la connessione e riavviare l'update", vbCritical
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

