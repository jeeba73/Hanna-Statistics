VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Controllo Aggiornamenti Online"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8820
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":1782
   ScaleHeight     =   5310
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00963D01&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   6480
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton Command2 
         Caption         =   "Richiedi in seguito"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   600
         TabIndex        =   12
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Aggiorna software"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   3600
         TabIndex        =   11
         Top             =   2040
         Width           =   4095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C87034&
         Height          =   2895
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   8295
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leggi prima le specifiche dell'aggiornamento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   600
         MouseIcon       =   "Form1.frx":50DFD
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   1320
         Width           =   4560
      End
      Begin VB.Label lbVersione 
         BackStyle       =   0  'Transparent
         Caption         =   "Versione"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   560
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Si consiglia di procedere e aggiornare il programma"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "E' disponibile una nuova versione di Verifica Periodica :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   240
         Width           =   5535
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1600
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   7080
      Top             =   1320
   End
   Begin VB.TextBox txVer 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox txVer 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   8520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   240
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00963D01&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2520
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton Command2 
         Caption         =   "Installa e Riavvia"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   2400
         TabIndex        =   14
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label lbflash 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scarico dal server la versione aggiornata"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   360
         Width           =   4140
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Per rendere effettive le modifiche riavviare il programma cliccando il pulsante in basso."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   1920
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   720
         Width           =   4815
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C87034&
         Height          =   2895
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   8295
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00963D01&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   240
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton Command1 
         Caption         =   "Esci"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   5280
         TabIndex        =   21
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ricerca aggiornamenti"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   1200
         TabIndex        =   20
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C87034&
         Height          =   2895
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   8295
      End
      Begin VB.Label lbVer 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1080
         TabIndex        =   22
         Top             =   1080
         Width           =   5535
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Non sono disponibili nuove versioni del programma"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   1080
         TabIndex        =   19
         Top             =   600
         Width           =   6105
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "Form1.frx":51107
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Controlla Aggiornamenti"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C87034&
      Height          =   495
      Left            =   1200
      TabIndex        =   25
      Top             =   120
      Width           =   4155
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Aggiornamento automatico del software."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   24
      Top             =   600
      Width           =   5415
   End
   Begin SoftwareWebUpdate.ucAniGIF ucAniGIF1 
      Height          =   3060
      Left            =   2880
      Top             =   2040
      Width           =   3180
      _ExtentX        =   7938
      _ExtentY        =   7938
      GIF             =   "Form1.frx":57959
      DelayLoad       =   0
   End
   Begin VB.Label lbRicerca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ricerca di nuove versioni di "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   3315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ricerca degli aggiornamenti in corso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   7560
      Width           =   6015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versione disponibile al download :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1920
      TabIndex        =   1
      Top             =   7320
      Width           =   2880
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versione attualmente installata :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2280
      TabIndex        =   0
      Top             =   8520
      Visible         =   0   'False
      Width           =   2745
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Private StgVersione As String   ' stringa per scaricare il file txt
Private StgVersioneftp As String  ' stringa per scaricare il file exe

Private f_doShow As Boolean
Dim fso As New FileSystemObject

Private Function SmartUpdate()

    Timer1.Enabled = True
    Frame3.Visible = False
    Frame1.Visible = False
    ProgressBar1.Value = 10
    
    Call GetSpecificheAggiornamento
      
    ProgressBar1.Value = 30
    
    
   CloseSettingDataFile
    If GetFileFromUrl("http://www.bilsoft.it/Download/" & PROGRAM_NAME & "/Update/", PROGRAM_NAME & ".txt") Then
      '-----------------------------------------------------------------------
      ' ok esiste il file con la versione
      '-----------------------------------------------------------------------
      StgVersioneftp = GetVerSoft(PROGRAM_NAME & ".txt")
      txVer(1) = StgVersioneftp
      
      ProgressBar1.Value = 50
      CloseSettingDataFile
      
      If EsistonoAggiornamenti(StgVersione, StgVersioneftp) Then
          '-----------------------------------
          ' esistono gli aggiornamenti
          '-----------------------------------
          lbVersione = PROGRAM_NAME & " r" & txVer(1)
          lbVersione.ForeColor = &H8000&
          ProgressBar1.Value = 80
          Frame3.Visible = False
          Frame1.Visible = True

      Else
          '-----------------------------------
          ' NO aggiornamenti
          '-----------------------------------

      
          ProgressBar1.Value = 80
          Frame3.Visible = True
          lbVer = "L'attuale versione r" & txVer(0) & " è la più aggiornata."
          lbVer.ForeColor = &H80FF&
        ' ucAniGIF1.Action = gfaStop
      End If
    End If
    ProgressBar1.Value = 100
    Timer1.Enabled = False
    Label2.ForeColor = &H800000
    ucAniGIF1.Action = gfaStop
    ucAniGIF1.Visible = False
    CloseSettingDataFile
End Function


Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
            End
        Case 1
            Call SmartUpdate
    End Select
End Sub

Private Sub Command2_Click(Index As Integer)

    
    
    Select Case Index
    

        Case 0
            Frame1.Visible = False
            Frame2.Visible = True
            Label8.Visible = False
            Timer1.Enabled = True
            
            
            
            Call ChiudeProgramma
            
            Label7 = ""
            Command2(3).Enabled = False
            'lbflash
            lbflash = "Scarico dal server la versione aggiornata r" & txVer(1)
            lbflash.ForeColor = vbWhite
            ProgressBar1.Value = 10
            ' creiamo il backup della versione attuale
            Call CreaBackupExe(PROGRAM_NAME)
            ProgressBar1.Value = 50


            If GetFileFromUrl("http://www.bilsoft.it/Download/" & PROGRAM_NAME & "/Update/", PROGRAM_EXE_NAME) Then

            
                lbflash.Visible = False
                ProgressBar1.Value = 100
                Label7 = "Il file è stato scaricato correttamente."
                Label7.ForeColor = &H80FF&
                Command2(3).Enabled = True
                Label8.Visible = True
                
                
                SaveSetting "Update " & PROGRAM_NAME, "UPDATE", "ESEGUITO", True
                SaveSetting "Update " & PROGRAM_NAME, "UPDATE", "IN DATA", Date
                SaveSetting "Update " & PROGRAM_NAME, "UPDATE", "PRIMO AVVIO", True
                
            Else
                lbflash.Visible = False
                SaveSetting "Update " & PROGRAM_NAME, "UPDATE", "ESEGUITO", False
                SaveSetting "Update " & PROGRAM_NAME, "UPDATE", "IN DATA", ""
                SaveSetting "Update " & PROGRAM_NAME, "UPDATE", "PRIMO AVVIO", False
                Label7 = "Impossibile scaricare il file..."
                Label7.ForeColor = &H80&
                
                
            End If
            Timer1.Enabled = False
           ' lbflash.ForeColor = vbBlack
            ucAniGIF1.Visible = False
            
            
        Case 1
            '-----------------------
            ' esci, alla prossima
            '-----------------------
            ucAniGIF1.Visible = False
            Unload Me
            End
        
        Case 3
            '-----------------------
            ' riavvia e aggiorna
            '-----------------------
           
            DoEvents
            
            SostituisceExe PROGRAM_NAME
            DoEvents
            
            ApreProgramma
            Unload Me
            End
    End Select
End Sub

Private Sub Form_Initialize()
Load_DOC_FOLDER
DoEvents
PROGRAM_NAME = GetExeName

If CheckInternetConn = False Then End
End Sub

Private Sub Form_Load()

    Me.Visible = True
    With Inet
        .url = "ftp://ftp.bilsoft.it"
        .UserName = "9620198@aruba.it"
        .Password = "qjE4Kb7NGhUF"
    End With
    
    StgVersione = PROGRAM_VERSIONE ' GetVerSoft("nver.txt")
    txVer(0) = StgVersione
  
    If f_doShow = False Then
        Call SmartUpdate
        f_doShow = True
    End If
    
    Me.ZOrder
  
End Sub

Private Sub Form_Activate()
    'WindowsXPC1.InitSubClassing
    ProgressBar1.Max = 100
    
    Frame2.Move Frame1.Left, Frame1.Top, Frame1.Width, Frame1.Height
    Frame3.Move Frame1.Left, Frame1.Top, Frame1.Width, Frame1.Height
    
    Me.Caption = "Controllo Aggiornamenti Online " ' r" & App.Major & "." & App.Minor & "." & App.Revision
    
    lbRicerca = "Ricerca di nuove versioni di " & PROGRAM_NAME
    Label5 = "E' disponibile una nuova versione :" ' & PROGRAM_NAME & " :"
    ucAniGIF1.Action = gfaPlay
    ucAniGIF1.ZOrder
End Sub

Public Function GetExeName() As String
  '-------------------------------------------
  ' va a prendere la versione attuale
  '-------------------------------------------
  Dim mName As String

   
   If PC_DOCUMENTI <> "" Then
    mName = GetSettingData("nome.txt", "Programma", "Nome", "", PC_DOCUMENTI)
    If mName = "" Then
    Else
    PROGRAM_EXE_NAME = GetSettingData("nome.txt", "Programma", "ExeName", "", PC_DOCUMENTI)
    PROGRAM_VERSIONE = GetSettingData("nome.txt", "Versione", "Rel.", "", PC_DOCUMENTI)
    USER_UPDATE_PATH = GetSettingData("nome.txt", "Aggiornamento", "Path", "", PC_DOCUMENTI)
   End If
   End If
   CloseSettingDataFile
ERR_END:
    On Error GoTo 0
    GetExeName = mName
    Exit Function
ERR_GET:
    Resume ERR_END
End Function


Public Function GetFileFromUrl(ByRef url As String, ByRef vFile As String) As Boolean
Dim fileBytes() As Byte
Dim fileNum As Integer
Dim rc As Boolean
Dim a
  '-------------------------------------------
  ' scarica quella da ftp
  '-------------------------------------------

    If FileExists(USER_UPDATE_PATH & vFile) Then Kill USER_UPDATE_PATH & vFile
    
    
   'MsgBox USER_UPDATE_PATH
    On Error GoTo DownloadError
    DoEvents
    rc = True
    fileBytes() = Inet.OpenURL(url & vFile, icByteArray)
    ProgressBar1.Value = 70
    
    fileNum = FreeFile
    Open USER_UPDATE_PATH & vFile For Binary Access Write As #fileNum
    Put #fileNum, , fileBytes()
    Close #fileNum
    ProgressBar1.Value = 80
ERR_END:
    On Error GoTo 0
   ' MsgBox url & vFile
    GetFileFromUrl = rc
    Exit Function
DownloadError:
    MsgBox Err.Description
    rc = False
    Resume ERR_END
End Function

Public Function GetValidazioneFromUrl(ByRef url As String, ByRef vFile As String) As Boolean
Dim fileBytes() As Byte
Dim fileNum As Integer
Dim rc As Boolean
Dim a


    If FileExists(App.Path & "\" & vFile) Then Kill App.Path & "\" & vFile
    
    
   'MsgBox USER_UPDATE_PATH
    On Error GoTo DownloadError
    DoEvents
    rc = True
    fileBytes() = Inet.OpenURL(url & vFile, icByteArray)
    ProgressBar1.Value = 70
    
    fileNum = FreeFile
    Open App.Path & "\" & vFile For Binary Access Write As #fileNum
    Put #fileNum, , fileBytes()
    Close #fileNum
    ProgressBar1.Value = 80
ERR_END:
    On Error GoTo 0
    GetValidazioneFromUrl = rc
    Exit Function
DownloadError:
    MsgBox Err.Description
    rc = False
    Resume ERR_END
End Function

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


 Private Function EsistonoAggiornamenti(sVerLoc As String, sVerWeb As String) As Boolean
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

Private Function CreaBackupExe(ByVal vString As String)

    If FileExists(App.Path & "\" & PROGRAM_EXE_NAME) Then
        FileCopy App.Path & "\" & PROGRAM_EXE_NAME, App.Path & "\OLD_" & PROGRAM_EXE_NAME
        
        Kill App.Path & "\" & PROGRAM_EXE_NAME
    End If
    
End Function


Private Function GetSpecificheAggiornamento() As String
   If GetFileFromUrl("http://www.bilsoft.it/Download/" & PROGRAM_NAME & "/Update/", PROGRAM_NAME & "_info.doc") Then
   End If
   If GetFileFromUrl("http://www.bilsoft.it/Download/" & PROGRAM_NAME & "/Update/", PROGRAM_NAME & "_info.rtf") Then
   End If
   If GetValidazioneFromUrl("http://www.bilsoft.it/Download/" & PROGRAM_NAME & "/Update/", "Validazione " & Trim(PROGRAM_NAME) & ".pdf") Then
   End If
End Function

Public Function ApriTesto(strDownloadFile)

If Dir(strDownloadFile) <> "" Then
Dim lngResult As Long
'lngResult = ShellExecute(Me.hwnd, "Open", strDownloadFile, "", "", vbNormalFocus)

ShellExecute Me.hWnd, vbNullString, strDownloadFile, vbNullString, App.Path, 1
lngResult = lngResult
End If
End Function

Private Sub Form_Resize()
Frame2.Top = Frame3.Top
Frame2.Left = Frame3.Left
Frame1.Top = Frame3.Top
Frame1.Left = Frame3.Left
'ucAniGIF1.Move Me.Width / 2 - ucAniGIF1.Width / 2, Me.Height / 2 - ucAniGIF1.Height / 2 '+ ProgressBar1.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
CloseSettingDataFile
End
End Sub

Private Sub Label11_Click()
ShellExecute 0&, "open", USER_UPDATE_PATH & PROGRAM_NAME & "_info.rtf", "", "", vbNormalFocus
End Sub


Private Sub Timer1_Timer()
    lbflash.Visible = Not (lbflash.Visible)
    Label2.ForeColor = IIf(Label2.ForeColor = &H800000, &H8000&, &H800000)
End Sub


Private Sub ChiudeProgramma()
ChiudiEseguibile (PROGRAM_NAME)
End Sub

Private Sub ApreProgramma()
ApriEseguibile (App.Path & "\" & PROGRAM_EXE_NAME)
DoEvents
Unload Me
End Sub


Private Function SostituisceExe(ByVal vString As String)
If FileExists(USER_UPDATE_PATH & PROGRAM_EXE_NAME) Then

    FileCopy USER_UPDATE_PATH & PROGRAM_EXE_NAME, App.Path & "\" & PROGRAM_EXE_NAME
    DoEvents
    Kill USER_UPDATE_PATH & PROGRAM_EXE_NAME
End If
End Function
