VERSION 5.00
Begin VB.Form F_RegForm 
   BackColor       =   &H00644603&
   BorderStyle     =   0  'None
   Caption         =   "Registrazione software"
   ClientHeight    =   12000
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   19200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12000
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXT_REG 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00964901&
      Height          =   390
      Index           =   1
      Left            =   6000
      TabIndex        =   6
      Top             =   5400
      Width           =   7335
   End
   Begin VB.TextBox TXT_REG 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00964901&
      Height          =   390
      Index           =   0
      Left            =   6000
      TabIndex        =   5
      Top             =   4440
      Width           =   7335
   End
   Begin VB.TextBox TXT_REG 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
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
      Height          =   405
      Index           =   2
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   7440
      Width           =   7335
   End
   Begin VB.CommandButton CMD_REGISTRA 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Get Registration key"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   6000
      TabIndex        =   3
      Top             =   6120
      Width           =   7335
   End
   Begin VB.TextBox TXT_REG 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   6000
      TabIndex        =   2
      Top             =   8640
      Width           =   7335
   End
   Begin VB.CommandButton CMD_REGISTRA 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Apply"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   8520
      TabIndex        =   1
      Top             =   9840
      Width           =   4815
   End
   Begin VB.CommandButton CMD_REGISTRA 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Demo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   6000
      TabIndex        =   0
      Top             =   9840
      Width           =   2415
   End
   Begin VB.Label lbProdotto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   4800
      TabIndex        =   14
      Top             =   600
      Width           =   9615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   14400
      X2              =   14400
      Y1              =   0
      Y2              =   11880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   4800
      X2              =   4800
      Y1              =   0
      Y2              =   11880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   9600
      X2              =   9600
      Y1              =   120
      Y2              =   12000
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9360
      MouseIcon       =   "F_Reg_Form.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "F_Reg_Form.frx":030A
      Top             =   10920
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00AAB9A5&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   6000
      TabIndex        =   13
      Top             =   4920
      Width           =   1230
   End
   Begin VB.Label Label2 
      BackColor       =   &H00AAB9A5&
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   6000
      TabIndex        =   12
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Key"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   11
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Activation Key"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   3
      Left            =   6000
      TabIndex        =   10
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Software Registration"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   6000
      TabIndex        =   9
      Top             =   2520
      Width           =   2820
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"F_Reg_Form.frx":36EC
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
      Height          =   855
      Left            =   6000
      TabIndex        =   8
      Top             =   3000
      Width           =   7335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrizione"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00644603&
      Height          =   405
      Index           =   4
      Left            =   9000
      TabIndex        =   7
      Top             =   1560
      Width           =   1185
   End
End
Attribute VB_Name = "F_RegForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_rc As Boolean

Private m_UserID As String
Private m_PassMe As String
Private m_RegKey As String
Private m_ActKey As String
Private m_DemoDate As String
Private m_Demo As Boolean
Private HelpString(2) As String
Private Expired As Boolean
Private VisibleMe As Boolean
Private Registration As Boolean

Public Function DoShow(ByVal FormVisible As Boolean) As Boolean

    Dim m_FlgLoading As Boolean
    
    On Error GoTo ERR_SHOW
    
    m_rc = False
    Registration = False
    m_FlgLoading = False
    Expired = False
    VisibleMe = FormVisible
    
    F_MAIN.TimeriNTRO.Enabled = False
    
    Me.Show vbModal
    If Registration = False Then
            
            If m_rc = True Then
                    'in caso di modifiche
                    
                    SaveSetting App.Title, "Autorizzazione", "done", False
                    SaveSetting App.Title, "Autorizzazione", "Demo", False
                    SaveSetting App.Title, "Autorizzazione", "DemoDate", Date
                    
            Else
                '----------------------------
                '   DEMO DEMO DEMO
                '-----------------------------
                
                If Expired Then
                    SaveSetting App.Title, "Autorizzazione", "done", True
                    SaveSetting App.Title, "Autorizzazione", "Demo", True
                    End
                Else
                    If m_Demo Then
                    Else
                        SaveSetting App.Title, "Autorizzazione", "done", False
                        SaveSetting App.Title, "Autorizzazione", "Demo", True
                        SaveSetting App.Title, "Autorizzazione", "DemoDate", Date
                    End If
                    
                End If
            End If
        
    End If
    
    DoShow = m_rc
 
ERR_END:
    On Error GoTo 0
    Exit Function
    

ERR_SHOW:

    m_rc = False
    Resume ERR_END
 
End Function

Private Sub Form_Load()
'SaveLanguageFile

End Sub
Private Sub Form_Activate()

F_MAIN.Timer1.Enabled = False

mOk

       lbProdotto = App.Title & " v" & App.Major & "." & App.Minor
       Label2(4) = "Hanna Instrument : Qc Control "
        
        
        
        If GetSoftSetting Then
            '--------------------
            ' tutto bene!!!
            '--------------------
            If VisibleMe Then
                Call RegTrue
            Else
                m_rc = True
                Unload Me
            End If
        Else
            '--------------------
            ' già in Demo
            '--------------------
            If VisibleMe Then
            Else
                m_rc = False
                Unload Me
            End If
            
        End If
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape
            If m_rc = False Then
                If MsgBox("Attenzione il programma non è registrato.Procedo con la versione Demo?", vbInformation + vbOKCancel) = vbOK Then
                    Unload Me
                Else
                
                End If
            Else
            
            End If
            
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set F_RegForm = Nothing
End Sub

Private Sub Image1_Click()
  
    If CMD_REGISTRA(0).Enabled = False And CMD_REGISTRA(2).Enabled = False Then
        Unload Me
        End
    Else
        Unload Me
    
    End If
    


End Sub

Private Sub TXT_REG_Change(Index As Integer)
    
    If Index = 3 Then
            
            TXT_REG(Index).BackColor = IIf(Len(TXT_REG(Index)) > 0, vbWhite, &H808080)
            
            If CheckActKey(TXT_REG(0), TXT_REG(1), TXT_REG(2), TXT_REG(3), HelpString(1)) Then
                TXT_REG(Index).ForeColor = &H8000&
                CMD_REGISTRA(2).Enabled = True
                CMD_REGISTRA(0).Enabled = False
            Else
                TXT_REG(Index).ForeColor = vbBlack
                CMD_REGISTRA(2).Enabled = False
                CMD_REGISTRA(0).Enabled = True
            End If
    End If
    
    CMD_REGISTRA(1).Enabled = IIf(Len(TXT_REG(0)) > 0 And Len(TXT_REG(1)) > 0, True, False)

End Sub

Private Sub TXT_REG_Click(Index As Integer)
   ' TXT_REG(Index).SelStart = 0
   ' TXT_REG(Index).SelLength = Len(TXT_REG(Index))
    
End Sub

Private Sub TXT_REG_GotFocus(Index As Integer)

   ' TXT_REG(Index).SelStart = 0
   ' TXT_REG(Index).SelLength = Len(TXT_REG(Index))
    If Index = 3 Then TXT_REG(3).BackColor = vbWhite

End Sub

Private Sub TXT_REG_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub



Private Sub CMD_REGISTRA_Click(Index As Integer)
    Select Case Index
        Case 0
            '-------------------------
            '      demo
            '-------------------------
            
            Call DemoVersion
        Case 1
            Call CreaReg
        Case 2
            '-------------------------
            '      registra
            '-------------------------

            Call Registra
      
    End Select
    
End Sub

Private Function CreaReg()

TXT_REG(2).BackColor = vbWhite
TXT_REG(2) = RegKey(TXT_REG(0), TXT_REG(1))

End Function

Private Sub Registra()
    '-----------------------
    '     Registro
    '-----------------------
    If Registration = False Then
        If CheckActKey(TXT_REG(0), TXT_REG(1), TXT_REG(2), TXT_REG(3), HelpString(1)) Then
            Call TrueRegistration
            m_rc = True
            Unload Me
            
        Else
            TXT_REG(3).SetFocus
            m_rc = False
        End If
    Else
        Unload Me
    End If
End Sub



Private Sub DemoVersion()
'--------------------------
    'm_Demo = True
    m_rc = False
    If MsgBox("Attenzione il programma non è registrato.Procedo con la versione Demo?", vbInformation + vbOKCancel) = vbOK Then

        Unload Me
    Else
        'TXT_REG(0).SetFocus
    End If
End Sub
Private Function GetDemoExp() As Boolean

        GetDemoExp = True
        
        If GetSetting(App.Title, "Autorizzazione", "done", False) Then
            '--------------------------------
            '       soft Expirated
            '--------------------------------
            CMD_REGISTRA(0).Enabled = False
            GetDemoExp = False
            HelpString(0) = "La versione Demo è scaduta. Registrare il programma. Inserire una User ID e una password quindi premere avanti. "
        Else
            '--------------------------------
            ' conta quanto mancaaaaaa
            '--------------------------------
            If CheckTimeDemo() = False Then
                '--------------------------------
                '       mancano giorniii
                '--------------------------------
                HelpString(0) = "Il programma è attualmente in versione Demo dal : " & m_DemoDate & vbCrLf & _
                "E' possibile registrarlo correttamente inserendo UserID, Password per ottenere la chiave di registrazione."

                GetDemoExp = True
            Else
                '--------------------------------
                '       soft Expirated
                '--------------------------------
            
                CMD_REGISTRA(0).Enabled = False
                HelpString(0) = "La versione Demo è scaduta. Registrare il programma. Inserire una User ID e una password quindi premere avanti. "
                GetDemoExp = False
            End If
        End If
End Function

Private Function GetSoftSetting() As Boolean
        On Error Resume Next
        Registration = False
        
        m_UserID = GetSetting(App.Title, "Autorizzazione", "ID", "")
        m_PassMe = GetSetting(App.Title, "Autorizzazione", "Pass", "")
        m_RegKey = GetSetting(App.Title, "Autorizzazione", "regKey", "")
        m_ActKey = GetSetting(App.Title, "Autorizzazione", "ActKey", "")
        m_DemoDate = GetSetting(App.Title, "Autorizzazione", "DemoDate", "")
        m_Demo = GetSetting(App.Title, "Autorizzazione", "Demo", False)
    
    If m_Demo Then
        GetSoftSetting = False
        CMD_REGISTRA(0).Enabled = True

        If GetDemoExp Then
         
        Else
            '-----------------
            ' scadutoo
            '------------------
            HelpString(0) = "La versione Demo è scaduta. Procedere alla registrazione"
            Expired = True
            VisibleMe = True
        End If
    Else
        '---------------------
        ' prima atttivazione
        '---------------------
        VisibleMe = True
        CMD_REGISTRA(0).Enabled = True
           
        If m_ActKey <> "" Then
                '---------------------
                ' programma registrato
                '---------------------
                GetSoftSetting = True
                HelpString(0) = "Il Programma è attualmente registrato. Premere Esci o ESC per uscire."
                CMD_REGISTRA(0).Enabled = False
                Registration = True
                m_rc = True
                
                Dim TxCount As Integer
                For TxCount = 0 To 3
                    TXT_REG(TxCount).Locked = True
                    TXT_REG(TxCount).BackColor = &H8000000F
                Next
        Else
        
                HelpString(0) = "Procedura di attivazione del software. Inserire una UserID e una Password quindi premere Avanti. Demo per visionare il programma per " & ExpDays & " giorni."
                GetSoftSetting = False
        End If
            
        
    End If
End Function

Private Sub RegTrue()

    TXT_REG(0) = m_UserID
    TXT_REG(1) = m_PassMe
    TXT_REG(2) = m_RegKey
    TXT_REG(3) = m_ActKey
    

End Sub

Private Sub EndingRegProc()
    If Len(TXT_REG(1)) = 0 Then TXT_REG(1) = "00000"
    TXT_REG(2) = RegKey(TXT_REG(0), TXT_REG(1))
    TXT_REG(3).SetFocus
End Sub




Private Sub TrueRegistration()
            
             SaveSetting App.Title, "Autorizzazione", "done", True
             SaveSetting App.Title, "Autorizzazione", "ID", Trim(TXT_REG(0))
             SaveSetting App.Title, "Autorizzazione", "Pass", Trim(TXT_REG(1))
             SaveSetting App.Title, "Autorizzazione", "regKey", Trim(TXT_REG(2))
             SaveSetting App.Title, "Autorizzazione", "ActKey", Trim(TXT_REG(3))
             SaveSetting App.Title, "Autorizzazione", "Demo", False
            
             SaveSetting App.Title, "Autorizzazione", "Data", Date
             SaveSetting App.Title, "Autorizzazione", "Primo avvio", True
             

End Sub
