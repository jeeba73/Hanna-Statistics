VERSION 5.00
Begin VB.Form MessageInfo 
   Appearance      =   0  'Flat
   BackColor       =   &H00644603&
   BorderStyle     =   0  'None
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11355
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   2985
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer MessageFade 
      Interval        =   10
      Left            =   0
      Top             =   1320
   End
   Begin VB.Timer FadeDelay 
      Enabled         =   0   'False
      Interval        =   100
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C87034&
      Height          =   2985
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   11355
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   1
      Left            =   480
      Picture         =   "MessageInfo.frx":0000
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9960
      MouseIcon       =   "MessageInfo.frx":33E2
      MousePointer    =   99  'Custom
      Picture         =   "MessageInfo.frx":36EC
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Files caricati sul Server..."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   9375
      WordWrap        =   -1  'True
   End
   Begin VB.Label lTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ChemicalProduction"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   6000
   End
   Begin VB.Label lClose 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "  X  "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      MouseIcon       =   "MessageInfo.frx":6ACE
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   0
      Width           =   10935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00644603&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00322009&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   11415
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "MessageInfo.frx":6DD8
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "MessageInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'MESSAGE ON TOP
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
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
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Public bSonoGi‡Aperto As Boolean

Private m_rc As Boolean
Dim Fade As Integer



'FORM FADE
Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long

Dim Msg As Long
On Error Resume Next

If MessageInfoTime = 0 Then MessageInfoTime = 1300

If Perc < 0 Or Perc > 255 Then
  
    MakeTransparent = 1

Else
  
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    Msg = Msg Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
    MakeTransparent = 0

End If

If Err Then
  
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

If FadeDelay.Interval > MessageInfoTime Then

    FadeIn.Value = False
    FadeOut.Value = True
    FadeDelay.Enabled = False
    MessageInfo.Enabled = True

End If

End Sub

Public Function DoShow(Optional ByVal bRed As Boolean, Optional ByVal MyImage As Image, Optional ByVal bButton As Boolean) As Boolean
    On Error GoTo ERR_SHOW
    mOk
    m_rc = False
    If MyImage Is Nothing Then
    Else
    Set Image3(0) = MyImage
    End If
    If bRed Then
        lDescription.ForeColor = vbColorRed
        Image3(1).Visible = True
        Image3(0).Visible = False
    Else

    End If
    
    Image1.Visible = bButton

    Fade = 0
    MakeTransparent Me.hwnd, 0

    Me.Show vbModal
    
    If m_rc Then
    
    End If
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
End Sub

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Initialize()
lTitle = PROGRAM_NAME
End Sub

Private Sub Image1_Click()
m_rc = True
Unload Me
End Sub

Private Sub Image3_Click(Index As Integer)
Unload Me
End Sub

Private Sub lClose_Click()

FadeDelay.Enabled = False   'DONT ALLOW TO PROCEED TO DELAY TIME
FadeIn.Value = False        'SKIP MESSAGE FADE IN
FadeOut.Value = True        'ALLOW MESSAGE FADE OUT
MessageInfo.Enabled = True  'KEEP MESSAGE FADING

End Sub

Private Sub lClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

lClose.ForeColor = vbWhite

End Sub

Private Sub lClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

lClose.ForeColor = vbBlack

End Sub

Private Sub lDescription_Click()
Unload Me
End Sub

Private Sub lTitle_Change()


lTitle.FontSize = IIf(Len(lTitle) > 25, 20, 36)
    If Len(lTitle) > 25 Then
        lTitle.Top = lTitle.Top + 240
    End If

End Sub

Private Sub lTitle_Click()
Unload Me
End Sub

'FORM FADE IN EFFECT CONTROL TIMER
Private Sub messagefade_Timer()

'THE FADING EFFECT SPEED CAN BE CONTROLLE D BY CHANGING THE MESSGEFADE INTERVAL AND / OR
'BY CHANGING THE BELOW ADDITION AND SUBRATION VALUE (i.e. 20)

If FadeIn.Value = True Then

    If Fade <= 255 Then
 
        MakeTransparent Me.hwnd, Fade
        Fade = Fade + FadeTime               'FADE IN MESSAGE UNTIL IT IS FULLY APPEARED

    Else

        'MessageInfo.Enabled = False     'STOP WHEN FULLY APPEARED
        FadeDelay.Enabled = True        'START DELAY TIMER
        MakeTransparent Me.hwnd, 255    '255 = FULLY APPEARED

    End If

End If

If FadeOut.Value = True Then

    If Fade >= 0 Then
 
        MakeTransparent Me.hwnd, Fade
        Fade = Fade - FadeTime            'FADE OUT MESSAGE UNTIL IT IS FULLY DISAPPEARED

    Else

        MessageInfo.Enabled = False     'STOP WHEN FULLY DISAPPEARED
        MakeTransparent Me.hwnd, 0      '0 = FULLY DISAPPEARED
        MessageFade.Enabled = False
        On Error Resume Next
        Me.Hide                         'HIDE WHEN DONE
        Unload Me
    End If

End If

End Sub




