VERSION 5.00
Begin VB.Form F_DUPLICA 
   BackColor       =   &H00473733&
   BorderStyle     =   0  'None
   Caption         =   "Gestione Database"
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   16635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00964901&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   15480
      MouseIcon       =   "F_DUPLICAARCHIVIO.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   4455
      TabIndex        =   13
      Top             =   5880
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   1
         Left            =   1920
         MouseIcon       =   "F_DUPLICAARCHIVIO.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "F_DUPLICAARCHIVIO.frx":0614
         Top             =   195
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   0
      Left            =   10680
      MouseIcon       =   "F_DUPLICAARCHIVIO.frx":39F6
      ScaleHeight     =   855
      ScaleWidth      =   4455
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   0
         Left            =   2040
         MouseIcon       =   "F_DUPLICAARCHIVIO.frx":3D00
         MousePointer    =   99  'Custom
         Picture         =   "F_DUPLICAARCHIVIO.frx":400A
         Top             =   195
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.CommandButton Command1 
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
      Left            =   12840
      TabIndex        =   11
      Top             =   4920
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
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
      Left            =   9480
      TabIndex        =   10
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   960
      TabIndex        =   0
      Top             =   2280
      Width           =   14775
      Begin VB.TextBox Filepath 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   435
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   12975
      End
      Begin VB.TextBox Destinationpath 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   435
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   12975
      End
      Begin VB.CommandButton Browsedestination 
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
         Left            =   13440
         TabIndex        =   2
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Browsefile 
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
         Left            =   13440
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Filelabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   360
         TabIndex        =   6
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Destinationlabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   960
      ScaleHeight     =   1095
      ScaleWidth      =   14775
      TabIndex        =   7
      Top             =   840
      Width           =   14775
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   360
         Picture         =   "F_DUPLICAARCHIVIO.frx":73EC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can duplicate or backup Access database. Select destination and press OK"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   600
         Width           =   7635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backup Database"
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
         Left            =   1200
         TabIndex        =   9
         Top             =   120
         Width           =   3120
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
      Width           =   16695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C87034&
      Height          =   5850
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   16635
   End
End
Attribute VB_Name = "F_DUPLICA"
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

Public Function DoShow() As Boolean

    
    Dim m_FlgLoading As Boolean
    
    On Error GoTo ERR_SHOW
    
    
    m_rc = False
 
    Filepath = GetSetting(App.Title, "ARCHIVIO", "PATH", APP_DATA_FOLDER) & dbCodeName
    Destinationpath = USER_DOCUMENTI & dbCodeName
       
    
    m_FlgLoading = False
    
    Me.Show vbModal
    
    If m_rc = True Then
     
    End If
    
    DoShow = m_rc
    
ERR_END:
    On Error GoTo 0
    Exit Function
    

ERR_SHOW:

    m_rc = False
    Resume ERR_END
 
End Function

Private Sub Browsedestination_Click()

    Dim szFilename As String
    szFilename = GetFolder(Me.hWnd, USER_DOCUMENTI)
    If szFilename = "" Then Exit Sub
    Destinationpath = szFilename & "\" & dbCodeName
    
    Picture2(0).Enabled = True
    
End Sub

Private Sub Browsefile_Click()

    Dim szFilename As String
    szFilename = DialogFile(Me.hWnd, 1, "Open", m_name, "Database Access" & Chr(0) & "*.mdb" & Chr(0) & "Tutti i files" & Chr(0) & "*.*", "", "mdb")
    If szFilename = "" Then Exit Sub
    Filepath = szFilename
    
    Picture2(0).Enabled = True
   
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo ERR_COPY
 Select Case Index
    Case 0
        m_rc = True
        If Destinationpath = "" Then Destinationpath = USER_DOCUMENTI & dbCodeName

        If dbCode.State Then dbCode.Close
        FileCopy Filepath, Destinationpath
        
        Call SetArchivio
        Picture2(0).Enabled = False
    Case 1
        m_rc = False
        Unload Me
End Select
ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_COPY:
    PopupMessage 2, ("Impossibile procedere con l'operazione:") & vbCrLf & err.Description, , True
    Resume ERR_END:
End Sub

Private Sub Picture2_Click(Index As Integer)
Command1_Click Index
End Sub
