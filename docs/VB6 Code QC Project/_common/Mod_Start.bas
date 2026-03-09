Attribute VB_Name = "Mod_Start"
Option Explicit

Public MyAttualeOperatore As String
Public myUm As String

Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Declare Sub DisableProcessWindowsGhosting Lib "user32" ()
Private Function FindProcess(Process) As Long
    Dim objWMIService, colProcesses, objProcess
   
   Set objWMIService = GetObject("winmgmts:")
   Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process")
   For Each objProcess In colProcesses
       If LCase(Process) = LCase(objProcess.Caption) Then
            FindProcess = objProcess.ProcessID
            Exit For
        End If
   Next
End Function
Public Sub Main()

    Dim usa_p As Integer
    Dim rc As Boolean
    
   On Error Resume Next
   
    If App.PrevInstance = True Then
         'PopupMessage 2, "Application is already running."
          AppActivate FindProcess(App.EXEName & ".exe")
         End
    End If

    bFullScreen = GetSetting(App.Title, "Opzioni", "Full Screen Mode", False)
   'If Screen.Width - 19200 > 1000 Then
    PictureMaxScreen = App.PATH & "\Images\BackGroung2048x1159.jpg"
  ' End If
   Dim iccex As tagInitCommonControlsEx
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   
   
   
    Call GetSeparator
   
    Call GetProgramSettings
    
    '------------------------------------
    '  carica le cartelle del programma
    '------------------------------------
        
    Load_Folder
    
     '------------------------------
    ' apri l'archivio
    '------------------------------
    
    BackupDatabase
    
    rc = SetArchivio
    
     '--------------------------------
    '           licenza
    '--------------------------------

    GetLicency F_MAIN, False
    DoEvents

    '--------------------------------
    '  controlla il nuovo database
    '--------------------------------
    Call CheckCodeDB
    
    rc = True
    
    If rc Then
          With dbTabUserAccount
            
              bExistAccount = IIf(.EOF, False, True)
              
              If bExistAccount Then
                .filter = ""
                .filter = "IndexPrivilege=3"
                bExistAdministrator = IIf(.EOF, False, True)
                .filter = ""
                .filter = "IndexPrivilege=2"
                bExistTCO = IIf(.EOF, False, True)
                .filter = ""
                .filter = "IndexPrivilege=1"
                bExistManager = IIf(.EOF, False, True)
                               
              
              End If
              
          End With
          


         '------------------------------
         ' Login Avvio
         '------------------------------

         bLoginAvvio = GetSetting(App.Title, "Settings", "utilizza_pass", False)

         If bExistAccount And bLoginAvvio Then
              If frmLogin.DoShow(0) Then
              
              Else
                  End
              End If
         End If
         
         
       Call CreateLogFile(MyOperatore.Name)
       Call CreateVerFile
    
         '------------------------------
         ' procedi e avvia il programma
         '------------------------------
            
    F_MAIN.Timer1.Enabled = True
    F_MAIN.TimerIntro.Enabled = True
    
    
            Call SetStart
              
        '-------------------------------

    End If
   On Error GoTo 0
   
End Sub


Private Function SetStart(Optional Index As Integer = 0)
    
    F_MAIN.DoShow
        If MyOperatore.Name <> "" Then F_MAIN.Label5 = MyOperatore.Name
    
    
    
End Function


Public Function SetArchivio() As Boolean
    Dim rc As Boolean
    
    rc = True
    On Error GoTo ERR_SET
    
    dbPath = GetSetting(App.Title, "ARCHIVIO", "PATH", APP_DATA_FOLDER)
    MydbName = dbCodeName 'GetSetting(App.Title, "ARCHIVIO", "NOME", dbName)
    
   
    
    If m_CreateArchivio(dbPath, MydbName) Then
        
        Else
            rc = False
            PopupMessage 2, ("Warning : Database error...."), , True, App.Title
            End
    End If
ERR_END:
    On Error GoTo 0
    SetArchivio = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox err.Description
    Resume ERR_END
End Function

Public Function SetWorkStation() As Boolean
    Dim rc As Boolean
    
    On Error GoTo ERR_SET
    
    rc = FindWorkStation
    
    With MyWorkStation
        If rc Then
            .Enabled = True
                
            .Department = IIf(IsNull(Trim(dbTabLaboratorio!Department)), "", Trim(dbTabLaboratorio!Department))
            .Workstation = IIf(IsNull(Trim(dbTabLaboratorio!Workstation)), "", Trim(dbTabLaboratorio!Workstation))
            .LaboratoryManager = IIf(IsNull(Trim(dbTabLaboratorio!LaboratoryManager)), "", Trim(dbTabLaboratorio!LaboratoryManager))
            .Phone = IIf(IsNull(Trim(dbTabLaboratorio!Phone)), "", Trim(dbTabLaboratorio!Phone))
            .Description = IIf(IsNull(Trim(dbTabLaboratorio!Description)), "", Trim(dbTabLaboratorio!Description))
            .ServerUserID = IIf(IsNull(Trim(dbTabLaboratorio!ServerUserID)), "", Trim(dbTabLaboratorio!ServerUserID))
            .ServerPassword = IIf(IsNull(Trim(dbTabLaboratorio!ServerPassword)), "", Trim(dbTabLaboratorio!ServerPassword))
            .ServerFTP = IIf(IsNull(Trim(dbTabLaboratorio!ServerFTP)), "", Trim(dbTabLaboratorio!ServerFTP))
            .ServerWorkPath = IIf(IsNull(Trim(dbTabLaboratorio!ServerWorkPath)), "", Trim(dbTabLaboratorio!ServerWorkPath))
    
        Else
            .Enabled = False
        End If
    End With

ERR_END:
    On Error GoTo 0
    SetWorkStation = rc
    Exit Function
ERR_SET:
    rc = False
    Resume Next
End Function
