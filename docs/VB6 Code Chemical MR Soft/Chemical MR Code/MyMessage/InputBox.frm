VERSION 5.00
Begin VB.Form F_InputBox 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4170
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
   ScaleHeight     =   4170
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2040
      Width           =   11655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00886010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00886010&
      Height          =   375
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   13815
   End
   Begin VB.Label lbNumeric 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter numeric field"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00745613&
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
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
      MouseIcon       =   "InputBox.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3240
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00644603&
      Caption         =   "SI"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
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
      MouseIcon       =   "InputBox.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3240
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
   Begin VB.Image Image3 
      Height          =   480
      Index           =   1
      Left            =   300
      Picture         =   "InputBox.frx":0614
      Top             =   720
      Width           =   480
   End
   Begin VB.Label lTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chemical MR"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00644603&
      Height          =   825
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   4005
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   2175
      Index           =   0
      Left            =   480
      Top             =   720
      Width           =   10575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inserire il campo Richiesto"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   2970
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
Private bNumber As Boolean
Private MyreportString As String


Public Function DoShow(Optional ByVal Text As String = "", Optional ByVal Title As String = "", Optional ByVal bGood As Boolean = True, Optional ByVal SiTitle As String = "", Optional ByVal NoTitle As String = "", Optional ByRef ReportString As String = "", Optional ByRef MyImage As Image, Optional ByRef bNumero As Boolean, Optional ByVal FormTop As Double) As Boolean

    On Error GoTo ERR_SHOW
    
    If MyImage Is Nothing Then
    Else
    Set Image3(1) = MyImage
    End If
    mOk
    m_rc = False
    Me.Top = FormTop + 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    
    If Len(Title) = 0 Then Title = PROGRAM_NAME
    MyreportString = ReportString
    Text1 = MyreportString
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
    lTitle = Title
    lTitle.ForeColor = IIf(bGood, lTitle.ForeColor, vbColorRed)
    Label2 = Text
    
    Label1(0) = IIf(Len(SiTitle) > 0, SiTitle, "Save")
    Label1(1) = IIf(Len(NoTitle) > 0, NoTitle, "Exit")
    
    bNumber = bNumero
    lbNumeric.Visible = bNumber
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
        m_rc = True ' IIf(Len(Text1) > 0, True, False)
     
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

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)


Label1(Index).BackColor = vbColorButtonSI
Select Case Index
    Case 0
        Label1(1).BackColor = vbColorButtonMouseOver
       ' Label1(2).ForeColor = vbColorTextBlue
    Case 1
        Label1(0).BackColor = vbColorButtonMouseOver
       ' Label1(2).ForeColor = vbColorTextBlue
    
    Case 2
       ' Label1(1).ForeColor = vbColorTextBlue
        'Label1(0).ForeColor = vbColorTextBlue
End Select

End Sub


Private Sub lTitle_Change()
lTitle.Top = lTitle.Top + 100
    lTitle.FontSize = IIf(Len(lTitle) > 15, 30, 40)

End Sub

Private Sub Text1_DblClick()
  Text1 = FormatDataLAT(Date)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If bNumber Then KeyAscii = TxtToNumber(KeyAscii)
If KeyAscii = 13 Then
    Label1_Click 0
End If
End Sub
