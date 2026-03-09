VERSION 5.00
Begin VB.Form F_PRINTER_SETTING 
   BackColor       =   &H00644603&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00964901&
      Height          =   615
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   13335
   End
   Begin VB.Label lTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Printer Setup"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   27.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   360
      Width           =   2955
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "F_PRINTER_SETTING.frx":0000
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   330
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   2280
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00765F2D&
      Caption         =   "Salva"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   2880
      MouseIcon       =   "F_PRINTER_SETTING.frx":33E2
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3480
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00967F4D&
      Caption         =   "Annulla"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   8880
      MouseIcon       =   "F_PRINTER_SETTING.frx":36EC
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3480
      Width           =   5655
   End
End
Attribute VB_Name = "F_PRINTER_SETTING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private m_rc As Boolean
Private mIndex As Integer

Public Function DoShow() As Boolean
    On Error GoTo ERR_SHOW
    m_rc = False
    Dim i As Integer
    
     
    Call RiempiCombo
    Call ScriviPredefinita
  
    
    Me.Show vbModal
    If m_rc = True Then
            'in caso di modifiche
           
    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    Resume ERR_END
End Function


Private Sub Command1_Click(Index As Integer)
    
    Select Case Index
    
        Case 0
            SetPrinter
            m_rc = True
        Case 1
            m_rc = False
    End Select
    Unload Me
    

End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)





Label1(Index).BackColor = &H765F2D ' &H967F4D
Select Case Index
    Case 0
        Label1(1).BackColor = &H967F4D ' &H765F2D

    Case 1
        Label1(0).BackColor = &H967F4D '&H765F2D

    Case 2

End Select

End Sub
Private Sub RiempiCombo()
  Dim prn As Printer
  For Each prn In Printers
    Combo1.AddItem prn.DeviceName
  Next
  Combo1.ListIndex = 0
  Combo1 = Printer.DeviceName
End Sub

Private Sub ScriviPredefinita()
  lb(0) = "La stampante predefinita è " + vbCrLf + Printer.DeviceName
End Sub

Private Sub SetPrinter()
  Dim prn As Printer
  For Each prn In Printers
    If prn.DeviceName = Combo1.Text Then
      Set Printer = prn
      Exit For
    End If
  Next
  lb(0) = "La stampante scelta è " + vbCrLf + Printer.DeviceName
End Sub

Private Sub Label1_Click(Index As Integer)
Command1_Click Index
End Sub

