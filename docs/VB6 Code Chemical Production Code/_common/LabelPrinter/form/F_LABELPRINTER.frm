VERSION 5.00
Begin VB.Form F_LABELPRINTER 
   BackColor       =   &H00966C3E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Label Printer "
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13335
   Icon            =   "F_LABELPRINTER.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   13335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   615
      Left            =   10440
      TabIndex        =   8
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Frame FrameInside 
      BackColor       =   &H00966C3E&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   4095
      Index           =   6
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   12855
      Begin VB.TextBox Text13 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1560
         Width           =   11775
      End
      Begin VB.CheckBox Check18 
         BackColor       =   &H00966C3E&
         Caption         =   "Use Laber Printer ( Brother Printer )"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   600
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "+"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   2
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Select Label Printer "
         Height          =   615
         Left            =   840
         TabIndex        =   1
         Top             =   2760
         Width           =   11775
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label Path"
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
         Left            =   840
         TabIndex        =   6
         Top             =   1155
         Width           =   1080
      End
      Begin VB.Label lbLabelPrinter 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No printer selected"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   810
         TabIndex        =   5
         Top             =   2040
         Width           =   11790
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label Printer Settings"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   13215
   End
End
Attribute VB_Name = "F_LABELPRINTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_rc As Boolean
Private m_da As String
Private m_a As String

Private MyID As Long
Public Function DoShow() As Boolean
    
    On Error GoTo ERR_SHOW
    
    m_rc = False






   Text13 = MyPathLabel_Brother
        
        lbLabelPrinter = IIf(PRINTERNAME = "", "No printer selected", PRINTERNAME)
        
        Check18.Value = IIf(bStampaOk, 1, 0)
    









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



Private Sub Command1_Click()
MyPathLabel_Brother = Text13
SaveSetting App.Title, "PATH", "TEMPLATE LABEL", MyPathLabel_Brother
 
Unload Me
End Sub

Private Sub Form_Activate()
    '
    DropShadow Me.hwnd

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            m_rc = False
            Unload Me
    End Select
End Sub







Private Sub Check18_Click()
Dim rc As Boolean

rc = IIf(Check18.Value = 1, True, False)
'If rc Then

    SaveSetting App.Title, "LABEL PRINTER", "bUtilizzo", rc
  
If rc Then
  
 
    If Me.Visible Then SearchInfoLabelPrinter
    
  
    'Image1.Visible = bStampaOk
    
    SaveSetting App.Title, "LABEL PRINTER", "bUtilizzo", bStampaOk
    
    'Call SetFormPrinter
    
    If MyPathLabel_Brother = "" And Me.Visible Then
        PopupMessage 2, "Please, Set Label Path..."
        Command3_Click 2
        
    End If
    
End If
    
'End If
End Sub


Private Sub Command3_Click(Index As Integer)
Dim NumRecord As Long
    Select Case Index

        Case 2
            Call SetTemplateLabel(Me)
            
             Text13 = MyPathLabel_Brother
    End Select
End Sub


Private Sub Command11_Click()

    If SelezionoStampante Then
        bStampanteSelezionata = True
        
        SearchInfoLabelPrinter
    End If
            
    lbLabelPrinter = IIf(PRINTERNAME = "", "No printer selected", PRINTERNAME)

End Sub

