VERSION 5.00
Begin VB.Form MessageUploadFade 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer MessageFade 
      Interval        =   30
      Left            =   0
      Top             =   1320
   End
   Begin VB.Timer FadeDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   1320
   End
   Begin VB.OptionButton FadeIn 
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton FadeOut 
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbStazione 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STAZIONE 1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   760
      Width           =   1200
   End
   Begin VB.Label lDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Files caricati sul Server..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00964901&
      Height          =   855
      Left            =   1680
      TabIndex        =   4
      Top             =   1080
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Label lTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "BilCal Upload"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.Image Image3 
      Height          =   1185
      Left            =   120
      Picture         =   "MessageFade.frx":0000
      Top             =   360
      Width           =   1275
   End
   Begin VB.Label lClose 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "  X  "
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
      Left            =   360
      MouseIcon       =   "MessageFade.frx":7E7F
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   0
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00963D01&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "MessageUploadFade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'MESSAGE ON TOP
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

'GET SCREEN RIGHT CORNER
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long

Private Const SPI_GETWORKAREA = 48

'FORM FADE
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Public bSonoGi‡Aperto As Boolean

Dim Fade As Integer

'FORM FADE
Public Function MakeTransparent(ByVal hWnd As Long, Perc As Integer) As Long

Dim msg As Long
On Error Resume Next

If Perc < 0 Or Perc > 255 Then
  
    MakeTransparent = 1

Else
  
    msg = GetWindowLong(hWnd, GWL_EXSTYLE)
    msg = msg Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, msg
    SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
    MakeTransparent = 0

End If

If err Then
  
      MakeTransparent = 2

End If

End Function

'MESSAGE ON TOP
Public Sub MessageOnTop(hWindow As Long, bTopMost As Boolean)
    
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
    
Dim wFlags
Dim placement
    
wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    
placement = HWND_TOPMOST
    
SetWindowPos hWindow, placement, 0, 0, 0, 0, wFlags

End Sub

'POSITION THE FORM IN THE RIGHT CORNER OF SCREEN


Private Sub FadeDelay_Timer()

FadeDelay.Interval = FadeDelay.Interval + 1000

If FadeDelay.Interval > 2999 Then

    FadeIn.value = False
    FadeOut.value = True
    FadeDelay.Enabled = False
    MessageUploadFade.Enabled = True

End If

End Sub

Public Function DoShow(Optional ByVal bRed As Boolean) As Boolean
    On Error GoTo ERR_SHOW
    
    If bRed Then lDescription.ForeColor = vbRed

Call MessageOnTop(Me.hWnd, True) 'MESSAGE ON TOP

PlaceMessageInLowerRight Me, 0, 0 'MESSAGE PLACEMENT

Fade = 0
MakeTransparent Me.hWnd, 0
If lbStazione = "" Then lDescription.Top = lbStazione.Top
    Me.Show vbModal
ERR_END:
    On Error GoTo 0

    Exit Function
ERR_SHOW:

    Resume ERR_END
End Function


'POSITION THE FORM IN THE RIGHT CORNER OF SCREEN
Private Sub PlaceMessageInLowerRight(ByVal Frm As Form, ByVal right_margin As Single, ByVal bottom_margin As Single)
Dim frmTop As Long
Dim wa_info As RECT

'If MainForm.Taskbar.Value = True Then

    If SystemParametersInfo(SPI_GETWORKAREA, 0, wa_info, 0) <> 0 Then

        'GOT POSITION, PLACE THE FORM NOW
        Frm.Left = ScaleX(wa_info.Right, vbPixels, vbTwips) - Width - right_margin
        
        frmTop = (ScaleY(wa_info.Bottom, vbPixels, vbTwips) - Height * UploadDownloadMessageCounter - bottom_margin)
        If frmTop < 0 Then UploadDownloadMessageCounter = 1
        
        Frm.Top = ScaleY(wa_info.Bottom, vbPixels, vbTwips) - Height * UploadDownloadMessageCounter - bottom_margin
        bSonoGi‡Aperto = True
    End If

'End If

'If MainForm.WTaskbar.Value = True Then
        
    'DID NOT GOT THE WORK AREA BOUNDS, USE THE ENTIRE SCREEN
   ' frm.Left = Screen.Width - Width - right_margin
    'frm.top = Screen.Height - Height - bottom_margin
    
'End If

End Sub

Private Sub lClose_Click()

FadeDelay.Enabled = False   'DONT ALLOW TO PROCEED TO DELAY TIME
FadeIn.value = False        'SKIP MESSAGE FADE IN
FadeOut.value = True        'ALLOW MESSAGE FADE OUT
MessageUploadFade.Enabled = True  'KEEP MESSAGE FADING

End Sub

Private Sub lClose_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

lClose.ForeColor = vbWhite

End Sub

Private Sub lClose_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

lClose.ForeColor = vbBlack

End Sub

'FORM FADE IN EFFECT CONTROL TIMER
Private Sub messagefade_Timer()

'THE FADING EFFECT SPEED CAN BE CONTROLLED BY CHANGING THE MESSGEFADE INTERVAL AND / OR
'BY CHANGING THE BELOW ADDITION AND SUBRATION VALUE (i.e. 20)

If FadeIn.value = True Then

    If Fade <= 255 Then
 
        MakeTransparent Me.hWnd, Fade
        Fade = Fade + FadeTime                'FADE IN MESSAGE UNTIL IT IS FULLY APPEARED

    Else

       ' MessageUploadFade.Enabled = False     'STOP WHEN FULLY APPEARED
        FadeDelay.Enabled = True        'START DELAY TIMER
        MakeTransparent Me.hWnd, 255    '255 = FULLY APPEARED

    End If

End If

If FadeOut.value = True Then

    If Fade >= 0 Then
 
        MakeTransparent Me.hWnd, Fade
        Fade = Fade - FadeTime                'FADE OUT MESSAGE UNTIL IT IS FULLY DISAPPEARED

    Else

        MessageUploadFade.Enabled = False     'STOP WHEN FULLY DISAPPEARED
        MakeTransparent Me.hWnd, 0      '0 = FULLY DISAPPEARED
        On Error Resume Next
        MessageFade.Enabled = False
        Me.Hide                         'HIDE WHEN DONE
        Unload Me
    End If

End If

End Sub


