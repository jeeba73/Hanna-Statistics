VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Message"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   4860
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   360
      TabIndex        =   16
      Top             =   2520
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   143
      TabIndex        =   12
      Top             =   4080
      Width           =   4575
      Begin VB.OptionButton WTaskbar 
         Caption         =   "Below Taskbar"
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
         Left            =   2524
         TabIndex        =   14
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton Taskbar 
         Caption         =   "Above Taskbar"
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
         Left            =   641
         TabIndex        =   13
         Top             =   840
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "If the Message stops above the taskbar height, try Below Taskbar option and see what goes..."
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
         Height          =   705
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4395
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5040
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "MainForm.frx":058A
         Stretch         =   -1  'True
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Two types of Message Effects are available. Make sure you select either of those before clicking on Show Message..."
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
         Height          =   705
         Left            =   1080
         TabIndex        =   9
         Top             =   150
         Width           =   3795
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4860
      TabIndex        =   7
      Top             =   5700
      Width           =   4860
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
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
         Left            =   3480
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdShowMessage 
         Caption         =   "&Show Message"
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
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "&About"
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
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   5000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   5000
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.TextBox txtDesc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1515
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "MainForm.frx":13CC
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Text            =   "VerPeriodica 4 Upload file"
      Top             =   1440
      Width           =   3255
   End
   Begin VB.OptionButton FadeEffect 
      Caption         =   "Fade Effect"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.OptionButton PopEffect 
      Caption         =   "Popup Effect"
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
      Left            =   3120
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   60
      Left            =   0
      Picture         =   "MainForm.frx":13FB
      Stretch         =   -1  'True
      Top             =   960
      Width           =   5040
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Message Title:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I AM WORKING ON PROGRAMS LIKE "WAREHOUSE INVENTORY SYSTEM" & COMPUTERIZED MAINTENANCE MANAGEMENT SYSTEM".
'FOR IT I REQUIRED A MESSAGE TO POPUP DURING THE WORK TO ALERT ACTIVE USER OR USERS ABOUT THE EVENT
'(e.g. ANY USER LOGGED IN, RECORD SAVED ETC.). I TRIED THE SYSTRAY
'ICON BALLON MESSAGE BUT THE CODING WAS'NT SO EASY TO WORK WITH. THEN I TRIED TO MAKE A POPUP MESSAGE JUST LIKE
'MSN (R) MESSENGER HAS, THIS WORKED GREAT ON ALL THE SYSTEMS AND USERS LIKED IT TOO, THOUGH IT ALSO HAS SOME BUGS
'BUT THAT WILL ISA BE EASILY REMOVED IN NEXT VERSION. I GUESS WHO IS READING THIS MAY ALSO LIKE TO USE IT
'AS A MESSAGE........ :)
'DONT FORGET TO VOTE:



Private Sub cmdAbout_Click()

Unload About
About.Show vbModal, Me

End Sub

Private Sub cmdClose_Click()

End

End Sub

Private Sub cmdShowMessage_Click()



Call PopupMessage(IIf(FadeEffect.Value, 0, 1), txtDesc.Text)


End Sub



Private Sub VecchiaSub()
UploadDownloadMessageCounter = UploadDownloadMessageCounter + 1

ReDim Preserve myUploadMessageForm(1 To UploadDownloadMessageCounter) As Form

Set myUploadMessageForm(UploadDownloadMessageCounter) = New MessageDownloadFade

ReDim Preserve myDownloadMessageForm(1 To UploadDownloadMessageCounter) As Form

Set myDownloadMessageForm(UploadDownloadMessageCounter) = New MessageUploadFade


If FadeEffect.Value = True Then

   ' MessageDownloadFade.Hide   'IN CASE ONLY WHEN IT IS BEING USED

    'THIS IS THE ONLY CODE TO CALL THE MESSAGE
  '  Set frm = New MessageUploadFade

    'Unload MessageUploadFade
DoEvents
        If txtTitle.Text = "" Then txtTitle.Text = "Message Title Here"     'IF EMPTY ENTRY
        If txtDesc.Text = "" Then txtDesc.Text = "Description goes here"    'IF EMPTY ENTRY

   ' MessageUploadFade.lTitle = txtTitle.Text
    myDownloadMessageForm(UploadDownloadMessageCounter).lDescription = UploadDownloadMessageCounter

   ' MessageUploadFade.Show 'SHOW MESSAGE
    myDownloadMessageForm(UploadDownloadMessageCounter).Show 'SHOW MESSAGE
End If

If PopEffect.Value = True Then

   ' MessageUploadFade.Hide   'IN CASE ONLY WHEN IT IS BEING USED
DoEvents
    'THIS IS THE ONLY CODE TO CALL THE MESSAGE
   ' Unload MessageDownloadFade

        If txtTitle.Text = "" Then txtTitle.Text = "Message Title Here"     'IF EMPTY ENTRY
        If txtDesc.Text = "" Then txtDesc.Text = "Description goes here"    'IF EMPTY ENTRY

  '  MessageDownloadFade.lTitle = txtTitle.Text
    myUploadMessageForm(UploadDownloadMessageCounter).lDescription = txtDesc.Text & " - " & UploadDownloadMessageCounter

    myUploadMessageForm(UploadDownloadMessageCounter).Show 'SHOW MESSAGE

End If
End Sub

Private Sub Command1_Click()
PopupMessage 2, "timer 100", , True, "Info"
End Sub

Private Sub Form_Load()
UploadDownloadMessageCounter = 0
'CENTER THIS FORM
Me.Left = (Screen.Width - Me.Width) \ 2
Me.top = (Screen.Height - Me.Height) \ 2

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'I DON'T LIKE THIS METHODE OF CLOSING THE FORMS........ :(

Cancel = True
Exit Sub

End Sub
