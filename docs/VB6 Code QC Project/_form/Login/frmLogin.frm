VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   8955
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   15390
   FillColor       =   &H00008000&
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   15390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      Caption         =   "   Account List"
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
      Height          =   660
      Left            =   4080
      TabIndex        =   7
      Top             =   7320
      Width           =   7215
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
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
      Height          =   585
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4800
      Width           =   7215
   End
   Begin VB.ComboBox cmbAccount 
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
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.TextBox txUserName 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
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
      Height          =   465
      Left            =   4080
      TabIndex        =   0
      Top             =   3240
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Login"
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
      Height          =   615
      Index           =   0
      Left            =   4080
      MouseIcon       =   "frmLogin.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6120
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00964901&
      Caption         =   "Exit"
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
      Height          =   615
      Index           =   1
      Left            =   7800
      MouseIcon       =   "frmLogin.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   6120
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   255
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   15375
   End
   Begin VB.Label lClose 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "  X  "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      MouseIcon       =   "frmLogin.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   0
      Width           =   10935
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Operator"
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
      Height          =   825
      Left            =   90
      TabIndex        =   6
      Top             =   840
      Width           =   15255
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   4200
      Width           =   1845
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   2640
      Width           =   1965
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Pass As String
Private Operatore As String
Dim CurrentUser As String
Private bSeAdmin As Boolean
Private bSonoAdmin As Boolean
Private IndexPrivilege As Integer
Private strPrivilege As String
Private ReqPrivilege As Integer
Private m_rc As Boolean

Public Function DoShow(Optional ByVal IndPrivilege As Integer = 0) As Boolean
    
     On Error GoTo ERR_SHOW

    m_rc = False
    ReqPrivilege = IndPrivilege

    Operatore = MyOperatore.Name
    strPrivilege = GetPosition(ReqPrivilege)
    
    lbInfo.Caption = "Login " & strPrivilege
    

    Me.Show vbModal
    
    If m_rc = True Then
            MyOperatore.Name = Operatore
            MyOperatore.IndexPrivilege = IndexPrivilege
    End If
    
    DoShow = m_rc
    
ERR_END:
    On Error GoTo 0
    Exit Function
    

ERR_SHOW:

    m_rc = False
    Resume ERR_END
 

End Function


Private Sub CheckUserName()
    bSonoAdmin = False
     With dbTabUserAccount
         .filter = ""
         .filter = "USERID='" & Trim(txUserName) & "'"
         If .EOF Then
         Else
             Pass = Trim(!Password)
             IndexPrivilege = !IndexPrivilege
            
            If IndexPrivilege >= ReqPrivilege Then
                ' tutto bene....
            Else
                PopupMessage 2, "Warning :  " & strPrivilege & " only can proceed." & vbCrLf & "Please select another account....", , True, strPrivilege
                txUserName = ""

             End If
         End If
     End With
End Sub

Private Sub Check1_Click()
Dim rc As Boolean
rc = IIf(Check1.Value = 1, True, False)
SaveSetting App.Title, "Settings", "Lista Operatori", Check1.Value
cmbAccount.Visible = rc
If rc Then cmbAccount.SetFocus
End Sub



Private Sub cmbAccount_Click()
txUserName = cmbAccount
    If txUserName <> "" Then
        txtPassword.SetFocus
    End If


End Sub

Private Sub cmdCancel_Click()
    m_rc = False
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    
    Call CheckUserName
    If txtPassword.Text = "" Then
        txtPassword.SetFocus
        Exit Sub
    End If
    If txtPassword.Text = Pass Then
        DoEvents
        PassCorrect
    Else
       
        PopupMessage 2, "Wrong Password", , True
        'txtPassword = ""
        txtPassword = ""
        txtPassword.SetFocus
    End If
    

End Sub

Public Sub PassCorrect()
    m_rc = True
    Operatore = Trim(txUserName)
    
    Unload frmLogin
End Sub

Private Sub cmbAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txUserName <> "" Then
        txtPassword.SetFocus
    End If


End If
End Sub

Private Sub Form_Activate()
Dim a As Integer

    With dbTabUserAccount
        .filter = ""
        If .EOF Then
            MessageInfoTime = 2000
            PopupMessage 2, "Please go to Settings... and Enter User Account"
            cmbAccount.Clear
            Check1.Visible = False
            m_rc = True
            Unload Me
        Else
            Call AddUserCombo
            Check1.Value = 1 'GetSetting(App.Title, "Settings", "Lista Operatori", 0)
            
            txUserName = ""
            .MoveFirst
            
            If ReqPrivilege > 0 Then
                For a = 1 To .RecordCount
                    If !IndexPrivilege = ReqPrivilege Then
                        txUserName = Trim(!UserID)
                        txtPassword.SetFocus
                        cmbAccount.Visible = False
                        Check1.Value = 0
                    Else
                    End If
                    .MoveNext
                Next
            Else
setoper:
                If Operatore <> "" Then
                    txUserName = Operatore
                    txtPassword.SetFocus
                Else
                    txUserName.SetFocus
                End If
                If Check1.Value = 1 Then cmbAccount.SetFocus
            End If
        End If
    End With
End Sub
Private Sub AddUserCombo()
Dim i As Integer
cmbAccount.Clear
With dbTabUserAccount
    .MoveFirst
    For i = 1 To .RecordCount
        cmbAccount.AddItem IIf(IsNull(Trim(!UserID)), "", Trim(!UserID))
        .MoveNext
    Next
End With
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
    m_rc = False
    Unload Me
End Select
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
    Case 0
        cmdLogin_Click
    Case 1
        cmdCancel_Click
End Select
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txUserName <> "" Then
            Label1_Click 0
        End If
    End If
End Sub

Private Sub txUserName_KeyPress(KeyAscii As Integer)
'Debug.Print KeyAscii
If KeyAscii = 13 Then
    If Len(txUserName) = 0 Then
        PopupMessage 2, "Enter valid Username...", , True
    Else
        Call CheckUserName
        txtPassword.SetFocus
    End If

End If
End Sub
