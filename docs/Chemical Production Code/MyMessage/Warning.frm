VERSION 5.00
Begin VB.Form Warning 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13770
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   5160
      Picture         =   "Warning.frx":0000
      ScaleHeight     =   3840
      ScaleWidth      =   3840
      TabIndex        =   4
      Top             =   1560
      Width           =   3840
   End
   Begin VB.Label lTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preparation : Warning"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00644603&
      Height          =   1080
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   9015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00644603&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00524029&
      Height          =   255
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   13815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00644603&
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   7080
      MouseIcon       =   "Warning.frx":1744
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   7080
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00745613&
      Caption         =   "SI"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   1080
      MouseIcon       =   "Warning.frx":1A4E
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   7080
      Width           =   5655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      Visible         =   0   'False
      X1              =   6885
      X2              =   6885
      Y1              =   360
      Y2              =   3960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00877773&
      Height          =   7875
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   13770
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Warning.frx":1D58
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   600
      TabIndex        =   1
      Top             =   5760
      Width           =   12615
   End
   Begin VB.Image Image 
      Height          =   480
      Index           =   0
      Left            =   12840
      MouseIcon       =   "Warning.frx":1DFC
      MousePointer    =   99  'Custom
      Picture         =   "Warning.frx":2106
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image 
      Height          =   480
      Index           =   1
      Left            =   11760
      MouseIcon       =   "Warning.frx":54E8
      MousePointer    =   99  'Custom
      Picture         =   "Warning.frx":57F2
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   1
      Left            =   360
      Picture         =   "Warning.frx":8BD4
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Warning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_rc As Boolean



Public Function DoShow(Optional ByVal Text As String = "", Optional ByVal Title As String = "", Optional ByVal bGood As Boolean = True, Optional ByVal SiTitle As String = "", Optional ByVal NoTitle As String = "", Optional MyImage As Image) As Boolean

    On Error GoTo ERR_SHOW
    mOk
    m_rc = False


  
    If bGood = False Then
        Image3(0).Visible = False
        Image3(1).Visible = True
    
    
    End If
    If Len(Title) = 0 Then Title = PROGRAM_NAME
    lTitle = Title
    lTitle.ForeColor = IIf(bGood, lTitle.ForeColor, vbRed)
    Label2 = Text
    
    Label1(0) = IIf(Len(SiTitle) > 0, SiTitle, "Yes")
    Label1(1) = IIf(Len(NoTitle) > 0, NoTitle, "No")
   ' If Len(Text) < 50 Then Label2.Alignment = vbCenter
   Label1(0).ZOrder
    Me.Show vbModal
    
    If m_rc = True Then

    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    Resume ERR_END
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Label1_Click 0
End If
End Sub

Private Sub Form_Resize()
Shape1(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Image_Click(Index As Integer)
Select Case Index
    Case 0
        m_rc = False
        
    Case 1
        m_rc = True
End Select
Unload Me
End Sub

Private Sub Label1_Click(Index As Integer)
 Select Case Index

        Case 0
            Image_Click 1
        Case 1
            Image_Click 0
    End Select
    
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)



Label1(Index).BackColor = vbColorButtonSI
Select Case Index
    Case 0
        Label1(1).BackColor = vbColorButtonMouseOver
 
    Case 1
        Label1(0).BackColor = vbColorButtonMouseOver
    Case 2
    
End Select
End Sub



Private Sub Label2_Change()
If Len(Label2) > 80 Then
Label2.FontSize = 16
End If
End Sub

Private Sub lTitle_Change()
lTitle.Top = lTitle.Top + 100
    lTitle.FontSize = IIf(Len(lTitle) > 15, 30, 48)

End Sub

