VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form About 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5445
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   6810
      TabIndex        =   9
      Top             =   0
      Width           =   6810
      Begin VB.PictureBox FlagPic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   240
         ScaleHeight     =   975
         ScaleWidth      =   1305
         TabIndex        =   10
         Top             =   120
         Width           =   1305
         Begin SHDocVwCtl.WebBrowser FlagBrowser 
            Height          =   1770
            Left            =   -120
            TabIndex        =   11
            Top             =   -120
            Width           =   2625
            ExtentX         =   4630
            ExtentY         =   3122
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin VB.Image Image1 
         Height          =   585
         Left            =   1800
         Picture         =   "About.frx":0000
         Top             =   360
         Width           =   4905
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6810
      TabIndex        =   5
      Top             =   4830
      Width           =   6810
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Close"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdAbout 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Email Contact"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSysInfo 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&System Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   3840
      ScaleHeight     =   2295
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   1800
      Width           =   2655
      Begin VB.Label lTicDesc 
         BackStyle       =   0  'Transparent
         Caption         =   $"About.frx":18C0
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lTicTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Warehouse Inventory 1.0."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lTicTitle2 
         BackStyle       =   0  'Transparent
         Caption         =   "CMMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lTicDesc2 
         BackStyle       =   0  'Transparent
         Caption         =   $"About.frx":1953
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Timer timTic 
      Interval        =   35
      Left            =   6120
      Top             =   4080
   End
   Begin VB.Image Image3 
      Height          =   60
      Left            =   0
      Picture         =   "About.frx":1A10
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   6840
   End
   Begin VB.Image Image5 
      Height          =   15
      Left            =   3840
      Picture         =   "About.frx":2157
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Image Image6 
      Height          =   60
      Left            =   0
      Picture         =   "About.frx":289E
      Stretch         =   -1  'True
      Top             =   1170
      Width           =   6840
   End
   Begin VB.Image Image4 
      Height          =   15
      Left            =   3840
      Picture         =   "About.frx":2FE5
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Still to Come:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lscrolltext 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4065
      TabIndex        =   14
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   360
      Picture         =   "About.frx":372C
      Top             =   2483
      Width           =   240
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   2820
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "My Message"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   720
      TabIndex        =   13
      Top             =   2370
      Width           =   2775
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'PLAY MUSIC
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public strFileToPlay As String
Public bPlaying As Boolean

Public PicPath As String    'BROWSER FILE
Dim strName As String       'SCROLL TEXT
Dim CN As Integer           'SCROLL TEXT

'OPEN MUSIC
Public Sub OpenMovie()
    
If strFileToPlay <> "" Then

    mciSendString "open " & strFileToPlay & " type MPEGVideo", 0, 0, 0

End If

End Sub

'PLAY MUSIC
Public Sub PlayMovie()
    
If strFileToPlay <> "" Then
    
    mciSendString "play " & strFileToPlay, 0, 0, 0
    bPlaying = True
                
End If

End Sub

'STOP MUSIC
Public Sub CloseAll()
    
    mciSendString "close all", 0, 0, 0

End Sub


Private Sub cmdAbout_Click()
                                                                            'DATS MA EMAIL :)
MsgBox "Comments & Suggestions are Welcomed" & vbCrLf & vbCrLf & "---   kaleemullah@windowslive.com   ---", , Me.Caption

End Sub

Private Sub cmdClose_Click()

'STOP MUSIC
Call CloseAll
Unload Me

End Sub

Private Sub Form_Load()

'CENTER THIS FORM
Me.Left = (Screen.Width - Me.Width) \ 2
Me.top = (Screen.Height - Me.Height) \ 2

'SHOW GIF FILE IN BROWSER (A TRICKY WAY TO WORK WITH GIF FILES, IS'NT IT :)
PicPath = App.Path & "\Extra\flag.html" 'PAKISTAN FLAG
FlagBrowser.Navigate PicPath

'INITIAL POSITION TICKER
lTicTitle.top = 2700
lTicDesc.top = 3000

'PLAY MUSIC
Dim PathMusic
PathMusic = App.Path & "\Extra\Music.mp3"
strFileToPlay = PathMusic
strFileToPlay = """" & strFileToPlay & """"
Call OpenMovie
Call PlayMovie

CN = 0 'INITIAL VALUE SCROLL TEXT

Label3.Caption = Label3.Caption & App.Major & "." & App.Minor & "." & App.Revision 'APP DETAILS

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

timTic.Enabled = True   'ENABLE TICKER MOVE

End Sub


Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

timTic.Enabled = False  'DISABLE TICKER MOVE

End Sub

Private Sub timTic_Timer()

'USE THE LABEL TOP TO ACT AS MARQUEE DIRECTION UP :)

lTicTitle.top = lTicTitle.top - 20
lTicDesc.top = lTicDesc.top - 20
On Error Resume Next
lTicTitle2.top = lTicTitle2.top - 20
On Error Resume Next
lTicDesc2.top = lTicDesc2.top - 20


If lTicDesc.top = 1400 Then

    lTicTitle2.Visible = True
    lTicDesc2.Visible = True
    lTicTitle2.top = 2700
    lTicDesc2.top = 3000
    
End If

If lTicDesc2.top = -1400 Then

    lTicTitle.top = 2700
    lTicDesc.top = 3000
    lTicTitle2.Visible = False
    lTicDesc2.Visible = False

End If


'SCROLL TEXT
strName = "No copyrights for this program..."

CN = CN + 1 'ADD SINGLE LETTER

If CN = 100 Then
        
    CN = -1 'START OVER

Else

    lscrolltext.Caption = Left(strName, CN) 'SHOW strName

End If

End Sub
