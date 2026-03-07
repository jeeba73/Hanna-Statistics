VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   6705
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
   ScaleHeight     =   7095
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frCommandInside 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   11
      Top             =   5040
      Width           =   4575
      Begin VB.Label lbCommandInside 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   4575
      End
   End
   Begin VB.Frame frCommandInside 
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   9
      Top             =   5760
      Width           =   4575
      Begin VB.Label lbCommandInside 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   4575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   16695
      Begin VB.Image Image3 
         Height          =   240
         Index           =   1
         Left            =   120
         Picture         =   "frmLogin.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lbInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   480
         TabIndex        =   8
         Top             =   135
         Width           =   6135
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "   Account List"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   1080
      TabIndex        =   4
      Top             =   3600
      Width           =   4575
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2880
      Width           =   4575
   End
   Begin VB.ComboBox cmbAccount 
      BackColor       =   &H00FFFFFF&
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
      Height          =   435
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox txUserName 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
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
      Height          =   465
      Left            =   1080
      TabIndex        =   0
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Shape shInside 
      BackColor       =   &H00808080&
      BorderColor     =   &H00808080&
      Height          =   7095
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   6705
   End
   Begin VB.Label lClose 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "  X  "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      MouseIcon       =   "frmLogin.frx":33E2
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   0
      Width           =   10935
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      TabIndex        =   3
      Top             =   2520
      Width           =   1080
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      TabIndex        =   2
      Top             =   1320
      Width           =   1170
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
                      '  txtPassword.SetFocus
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

Private Sub frCommandInside_Click(Index As Integer)
Label1_Click Index
End Sub

Private Sub lbCommandInside_Click(Index As Integer)
Label1_Click Index
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
