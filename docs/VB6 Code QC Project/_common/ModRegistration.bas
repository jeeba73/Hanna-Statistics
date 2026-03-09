Attribute VB_Name = "Mod_Registration"
Option Explicit
Public DataLicenza As Date
Public bPrimoAvvio As Boolean
Public Const ExpDays = 21 'giorni di demoo
Public DemoDate As Date
Public bDemo As Boolean
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function GetSpecificheLicenza()
    
    
    DataLicenza = GetSetting(App.Title, "Autorizzazione", "Data", date)
    bPrimoAvvio = GetSetting(App.Title, "Autorizzazione", "Primo avvio", True)
    
    If bPrimoAvvio Then SaveSetting App.Title, "Autorizzazione", "Primo avvio", False
    
End Function

Public Function GetLicency(ByVal Frm As Form, ByVal VisualForm As Boolean) As Boolean
    On Error Resume Next
    Dim rc As Boolean
    rc = True
    
    GetLicency = False
    Call GetSpecificheLicenza
    If VisualForm = False Then
        Dim bValue As Boolean
        bValue = GetSetting(App.Title, "Autorizzazione", "demo", True)
        bDemo = bValue
        If bValue Then
            If CheckTimeDemo = False Then
                bValue = GetSetting(App.Title, "Autorizzazione", "done", True)
                If bValue Then
                    '--------------------------------------
                    '       blocco tutto
                    '--------------------------------------
                    VisualForm = True
                    GetLicency = False
                Else
                    
                    
                
                    Exit Function
                End If
            Else
                VisualForm = True
            End If
        Else
            
            Exit Function
        End If
    End If
        UploadDownloadMessageCounter = 0
        If F_RegForm.DoShow(VisualForm) Then
            '--------------------------------------
            '  procedo con la versione registrata
            '--------------------------------------
            Frm.Caption = App.Title
            GetLicency = True
            SaveSetting App.Title, "Autorizzazione", "demo", False
            PopupMessage 2, "Registration Successfull", , , App.Title
        Else
            '--------------------------------------
            '       demo demo demo demo
            '--------------------------------------
            SaveSetting App.Title, "Autorizzazione", "demo", True
            Frm.Caption = App.Title & " - Versione Demo dal " & DateDemo
            GetLicency = True
             PopupMessage 2, "30 days Demo Version activated..", , True, App.Title
        End If
       
    On Error GoTo 0
End Function
Private Function DateDemo() As Date
    DateDemo = GetSetting(App.Title, "Autorizzazione", "DemoDate", date)
End Function
Public Function WorkstationID() As String
  Dim sBuffer As String * 255

  If GetComputerNameA(sBuffer, 255&) > 0 Then
    WorkstationID = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
  Else
    WorkstationID = "?"
  End If

End Function

Public Function NetworkUserName() As String
  Dim lpBuff   As String * 25
  Dim RetVal   As Long

  RetVal = GetUserName(lpBuff, 25)
  ' trim off any trailing spaces found in the name
  NetworkUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

End Function


Public Sub SaveRegFile(ByVal MyKey As String, MyUserID As String, MyPassword As String, MyRegKey As String)
Dim FileName
FileName = App.Path & "\" & MyKey & ".sys"

Open FileName For Output As #1    ' Open file.
    Write #1, "User ID : " & MyUserID
    Write #1, "Password : " & MyPassword
    Write #1, "Registration Key : " & MyRegKey
    Write #1, "Activation Key : " & MyKey
Close #1    ' Close file.

End Sub

Public Function RegKey(ByVal UserID As String, ByVal PassMe As String)

    RegKey = "Rf7j" & Left((UserID), 4) & Right(StrReverse(PassMe), 3) & Mid(StrReverse(WorkstationID()), 3, 4)

End Function



Public Function CheckActKey(ByVal UnserID As String, ByVal Password As String, ByVal MyRegKey As String, ByVal MyActKey As String, ByRef HelpStr As String)
    

    If ActKey(MyRegKey, UnserID, Password, MyActKey) Then
        '--------------------
        '  attivazione OK
        '--------------------
        CheckActKey = True
        
       ' Call SaveRegFile(MyActKey, UnserID, Password, MyRegKey)

    Else
        '--------------------
        '  ERRORE!!!!!!!!
        '--------------------
        CheckActKey = False
        HelpStr = "La chiave di atttivazione non è Valida. Accertarsi di avere inserito correttamente i dati"

    End If

End Function


Private Function ActKey(ByVal RegUserKey As String, ByVal UserID As String, ByVal Password As String, ByVal MyActKey As String) As Boolean
    Dim OkString As String
    Dim YourString As String
    ActKey = True
    OkString = Left(StrReverse(UserID), 2) & "808" & Right(Password, 3) & "RG" & Right(RegUserKey, 3)
    YourString = Trim(MyActKey)
    
    If OkString <> YourString Then ActKey = False
    
End Function

Public Function CheckTimeDemo() As Boolean
    
  
    DemoDate = GetSetting(App.Title, "Autorizzazione", "DemoDate", date)
    
    If (DemoDate + ExpDays) <= date Then
    CheckTimeDemo = True
    Else
    CheckTimeDemo = False
    End If
    
End Function
