VERSION 5.00
Begin VB.Form F_InputBox 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13770
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   2040
      Width           =   11655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00964901&
      Caption         =   "NO"
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
      Height          =   615
      Index           =   1
      Left            =   7080
      MouseIcon       =   "InputBox.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3120
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "YES"
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
      Height          =   615
      Index           =   0
      Left            =   1080
      MouseIcon       =   "InputBox.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3120
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
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   13815
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "InputBox.frx":0614
      Top             =   720
      Width           =   480
   End
   Begin VB.Label lTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chemical QC"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   675
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   3465
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   2175
      Index           =   0
      Left            =   480
      Top             =   720
      Width           =   10575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Requested field..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   390
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   3090
   End
   Begin VB.Image Image 
      Height          =   480
      Index           =   0
      Left            =   12840
      MouseIcon       =   "InputBox.frx":39F6
      MousePointer    =   99  'Custom
      Picture         =   "InputBox.frx":3D00
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image 
      Height          =   480
      Index           =   1
      Left            =   11760
      MouseIcon       =   "InputBox.frx":70E2
      MousePointer    =   99  'Custom
      Picture         =   "InputBox.frx":73EC
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "F_InputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_rc As Boolean

Private MyreportString As String
Private IsNumericAnswer As Boolean

Public Function DoShow(Optional ByVal Text As String = "", Optional ByVal Title As String = "", Optional ByVal bGood As Boolean = True, Optional ByVal SiTitle As String = "", Optional ByVal NoTitle As String = "", Optional ByRef ReportString As String = "", Optional ByRef MyImage As Image, Optional ByRef IsNum As Boolean) As Boolean

    On Error GoTo ERR_SHOW
    
    If MyImage Is Nothing Then
    Else
    Set Image3(1) = MyImage
    End If
    IsNumericAnswer = IsNum
    mOk
    m_rc = False
    Me.TOp = F_MAIN.TOp + 600
    Me.Left = Screen.Width / 2 - Me.Width / 2
    If Len(Title) = 0 Then Title = ProjectName
    MyreportString = ReportString
    Text1 = MyreportString
    lTitle = Title
    lTitle.ForeColor = IIf(bGood, lTitle.ForeColor, vbRed)
    Label2 = Text
    
    Label1(0) = IIf(Len(SiTitle) > 0, SiTitle, "Save")
    Label1(1) = IIf(Len(NoTitle) > 0, NoTitle, "Exit")
   ' If Len(Text) < 50 Then Label2.Alignment = vbCenter
    
    Me.Show vbModal
    
    If m_rc = True Then
        
        
        ReportString = MyreportString
    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    Resume ERR_END
End Function



Private Sub Form_Resize()
Shape1(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
Private Sub Image_Click(Index As Integer)
Select Case Index
    Case 0
        m_rc = False
        
    Case 1
        m_rc = IIf(Len(Text1) > 0, True, False)
     
End Select
Unload Me
End Sub

Private Sub Label1_Click(Index As Integer)
MyreportString = Text1
 Select Case Index

        Case 0
            m_rc = True
            Image_Click 1
        Case 1
            m_rc = False
            Image_Click 0
    End Select
    
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


Label1(Index).BackColor = vbColorTextDarkBlue
Select Case Index
    Case 0
        Label1(1).BackColor = &H8000000D
       ' Label1(2).ForeColor = vbColorTextBlue
    Case 1
        Label1(0).BackColor = &H8000000D
       ' Label1(2).ForeColor = vbColorTextBlue
    
    Case 2
       ' Label1(1).ForeColor = vbColorTextBlue
        'Label1(0).ForeColor = vbColorTextBlue
End Select

End Sub


Private Sub lTitle_Change()
'lTitle.TOp = lTitle.TOp + 100
   ' lTitle.FontSize = IIf(Len(lTitle) > 15, 24, 30)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)


   
If KeyAscii = 13 Then
    Label1_Click 0
End If

If IsNumericAnswer Then KeyAscii = TxtToNumber(KeyAscii)
End Sub
