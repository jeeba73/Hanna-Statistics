VERSION 5.00
Begin VB.Form F_PERCORSO_ARCHIVIO 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Percorso archivio dbPeriodica.mdb"
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   16395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMD_BUTTON 
      Caption         =   "Reset Default Path"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   720
      TabIndex        =   8
      Top             =   3840
      Width           =   4335
   End
   Begin VB.CommandButton CMD_BUTTON 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   11280
      TabIndex        =   4
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton CMD_BUTTON 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   13800
      TabIndex        =   3
      Top             =   3840
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00964901&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   720
      ScaleHeight     =   975
      ScaleWidth      =   15015
      TabIndex        =   0
      Top             =   840
      Width           =   15015
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00964901&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   240
         Picture         =   "F_PERCORSO_ARCHIVIO.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can change database path. Click on  ... and select path."
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
         Left            =   1080
         TabIndex        =   9
         Top             =   600
         Width           =   6045
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Database Path"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   3825
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00964901&
      BorderStyle     =   0  'None
      Caption         =   "Percorso file"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   720
      TabIndex        =   5
      Top             =   2160
      Width           =   15015
      Begin VB.TextBox txt_STAMPA 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   405
         Left            =   120
         TabIndex        =   7
         Text            =   "...."
         Top             =   410
         Width           =   13335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   13680
         TabIndex        =   6
         Top             =   380
         Width           =   1095
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   16455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C87034&
      Height          =   4905
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   16395
   End
End
Attribute VB_Name = "F_PERCORSO_ARCHIVIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_path As String
Private m_name As String
Private m_old_name As String
Private m_rc As Boolean
Private MyBackupPath As String
Private MyBackupName As String

Public Function DoShow(Optional ByRef t_path As String, Optional ByRef t_nome As String, Optional Index As Integer = 0) As Boolean



    Dim m_FlgLoading As Boolean
    Dim MyPath As String
    
    On Error GoTo ERR_SHOW
   
    
    MyPath = GetSetting(App.Title, "ARCHIVIO", "PATH", "")
    If MyPath = "" Then
        txt_STAMPA = APP_DATA_FOLDER
    Else
        txt_STAMPA = MyPath
    End If
    
    

    
    
    m_rc = False

    m_FlgLoading = False
    
    Me.Show vbModal
    
    If m_rc = True Then
            'in caso di modifiche
        t_path = m_path
        t_nome = m_name
    End If
    
    DoShow = m_rc
    
ERR_END:
    On Error GoTo 0
    Exit Function
    

ERR_SHOW:

    m_rc = False
    Resume ERR_END
 
End Function


Private Sub CMD_BUTTON_Click(Index As Integer)
'Dim rc As Boolean
On Error Resume Next
    Select Case Index
        Case 0
            m_rc = SetPercorsoArchivio
            If m_rc Then
                UploadDownloadMessageCounter = 0
                PopupMessage 2, ("Database path : OK"), , , ("Database Path")
                Unload Me
            Else
                PopupMessage 2, ("Warning : wrong selected path"), , , ("Database Path")
            End If
           ' m_rc = True
        Case 1
            m_rc = False
            Unload Me
        Case 2
             txt_STAMPA = APP_DATA_FOLDER
            m_rc = SetPercorsoArchivio
            If m_rc Then
                UploadDownloadMessageCounter = 0
                PopupMessage 2, ("Database path : OK"), , , ("Database Path")
                Unload Me
            Else
                PopupMessage 2, ("Warning : no dbChemical.mdb in selected folder"), , True
            End If
             
    End Select
        
End Sub

Private Sub Command1_Click()

    Dim szFilename As String
  '  szFilename = DialogFile(Me.hwnd, 1, "Open", m_name, "Microsoft Word" & Chr(0) & "*.doc" & Chr(0) & "Tutti i files" & Chr(0) & "*.*", "", "doc")
    
    szFilename = BrowseFolder(Me.hWnd, ("Please select dbChemicalQC Path:"))
            
    If szFilename = "" Then
        CMD_BUTTON(1).Enabled = False
        Exit Sub
    Else
        txt_STAMPA = szFilename
    End If

   
End Sub



Private Function SetPercorsoArchivio() As Boolean
    Dim rc As Boolean
    Dim MyPath As String
    rc = True
    On Error GoTo ERR_SET
    MyPath = txt_STAMPA
    If MyPath = "" Then
        dbPath = APP_DATA_FOLDER
        txt_STAMPA = APP_DATA_FOLDER
    Else
    
        If InStr(MyPath, ".mdb") > 0 Then
            MyPath = Left$(MyPath, Len("dbWeighCheck.mdb"))
        Else
            If Right$(MyPath, 1) = "\" Then
            Else
                MyPath = MyPath & "\"
            End If
        End If
        dbPath = MyPath
        
    End If
    
    
    
    SaveSetting App.Title, "ARCHIVIO", "PATH", dbPath
    MydbName = dbCodeName 'GetSetting(App.Title, "ARCHIVIO", "NOME", dbName)
    
    
    If m_CreateArchivio(dbPath, MydbName) Then
        Else
            'rc = False
           ' SaveSetting App.Title, "ARCHIVIO", "PATH", APP_DATA_FOLDER
    End If
ERR_END:
    On Error GoTo 0
    SetPercorsoArchivio = rc
    Exit Function
ERR_SET:
    rc = False
    
    MsgBox err.Description
    Resume ERR_END
End Function



