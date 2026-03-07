VERSION 5.00
Begin VB.Form F_MsgBox 
   BackColor       =   &H00473733&
   BorderStyle     =   0  'None
   ClientHeight    =   4710
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
   ScaleHeight     =   4710
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
   Begin VB.Image Image3 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "MsgBox.frx":0000
      Top             =   600
      Width           =   480
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
      MouseIcon       =   "MsgBox.frx":33E2
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3720
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
      MouseIcon       =   "MsgBox.frx":36EC
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3720
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
   Begin VB.Label lTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CalWeight"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   48
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1080
      Left            =   1080
      TabIndex        =   2
      Top             =   360
      Width           =   4260
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00877773&
      Height          =   2175
      Index           =   0
      Left            =   600
      Top             =   720
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"MsgBox.frx":39F6
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   12255
   End
   Begin VB.Image Image 
      Height          =   480
      Index           =   0
      Left            =   12840
      MouseIcon       =   "MsgBox.frx":3A9A
      MousePointer    =   99  'Custom
      Picture         =   "MsgBox.frx":3DA4
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image 
      Height          =   480
      Index           =   1
      Left            =   11760
      MouseIcon       =   "MsgBox.frx":7186
      MousePointer    =   99  'Custom
      Picture         =   "MsgBox.frx":7490
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   1
      Left            =   360
      Picture         =   "MsgBox.frx":A872
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "F_MsgBox"
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

    If MyImage Is Nothing Then
    Else
    Set Image3(1) = MyImage
    End If
    
  
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

Private Sub Picture1_Click()

End Sub
