VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form frmPreferenze 
   BackColor       =   &H00303030&
   Caption         =   "Database"
   ClientHeight    =   12000
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   19200
   Icon            =   "FrmPassword.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   12000
   ScaleMode       =   0  'User
   ScaleWidth      =   19290.42
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   840
      ScaleHeight     =   465
      ScaleWidth      =   375
      TabIndex        =   27
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00404040&
      Caption         =   "Tecnical Office"
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
      Height          =   465
      Left            =   1200
      MaskColor       =   &H8000000F&
      TabIndex        =   26
      Top             =   7680
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   19215
      TabIndex        =   22
      Top             =   1080
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QC Operator Default User"
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
         Index           =   3
         Left            =   720
         TabIndex        =   24
         Top             =   120
         Width           =   4350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter UserName and Password : Set Administrator / Laboratory Manager / Statup Login"
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
         Height          =   225
         Index           =   3
         Left            =   720
         TabIndex        =   23
         Top             =   600
         Width           =   7200
      End
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00404040&
      Caption         =   "Laboratory Manager"
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
      Height          =   465
      Left            =   1200
      MaskColor       =   &H8000000F&
      TabIndex        =   7
      Top             =   8640
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.CheckBox pass_ok 
      BackColor       =   &H00404040&
      Caption         =   "Startup Login"
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
      Height          =   465
      Left            =   1200
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   5760
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.PictureBox chPass 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   840
      ScaleHeight     =   465
      ScaleWidth      =   375
      TabIndex        =   21
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   840
      ScaleHeight     =   465
      ScaleWidth      =   375
      TabIndex        =   20
      Top             =   8640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      Caption         =   "Administrator"
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
      Height          =   465
      Left            =   1200
      MaskColor       =   &H8000000F&
      TabIndex        =   5
      Top             =   6720
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   840
      ScaleHeight     =   465
      ScaleWidth      =   375
      TabIndex        =   19
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PicMainMenu 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   4
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   19215
      TabIndex        =   8
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MouseIcon       =   "FrmPassword.frx":33E2
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   13
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   735
            MouseIcon       =   "FrmPassword.frx":36EC
            MousePointer    =   99  'Custom
            Picture         =   "FrmPassword.frx":39F6
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Save Account"
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
            Height          =   225
            Index           =   0
            Left            =   465
            MouseIcon       =   "FrmPassword.frx":6DD8
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   720
            Width           =   1080
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   1920
         MouseIcon       =   "FrmPassword.frx":70E2
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   11
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clear form"
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
            Height          =   225
            Index           =   1
            Left            =   510
            MouseIcon       =   "FrmPassword.frx":73EC
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   720
            Width           =   870
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   735
            MouseIcon       =   "FrmPassword.frx":76F6
            MousePointer    =   99  'Custom
            Picture         =   "FrmPassword.frx":7A00
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   3840
         MouseIcon       =   "FrmPassword.frx":ADE2
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   9
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Delete Account"
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
            Height          =   225
            Index           =   2
            Left            =   0
            MouseIcon       =   "FrmPassword.frx":B0EC
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   735
            MouseIcon       =   "FrmPassword.frx":B3F6
            MousePointer    =   99  'Custom
            Picture         =   "FrmPassword.frx":B700
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Manager"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   15420
         TabIndex        =   15
         Top             =   360
         Width           =   3270
      End
   End
   Begin VB.TextBox Pass 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00661D01&
      Height          =   615
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   4680
      Width           =   7455
   End
   Begin VB.TextBox Pass 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00661D01&
      Height          =   600
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   3240
      Width           =   7455
   End
   Begin FlexCell.Grid GrdAccount 
      Height          =   7200
      Left            =   9120
      TabIndex        =   18
      Top             =   3240
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   12700
      AllowUserReorderColumn=   -1  'True
      AllowUserSort   =   -1  'True
      Appearance      =   0
      BackColor1      =   14737632
      BackColor2      =   14737632
      BackColorActiveCellSel=   12632256
      BackColorBkg    =   14737632
      BackColorFixed  =   12632256
      BackColorFixedSel=   12632256
      BackColorScrollBar=   -2147483635
      BackColorSel    =   8421504
      BorderColor     =   9849089
      CellBorderColor =   16512
      CellBorderColorFixed=   16777215
      Cols            =   10
      DefaultFontName =   "Calibri"
      DefaultFontSize =   12
      DefaultFontBold =   -1  'True
      DisplayDateTimeMask=   -1  'True
      FixedRowColStyle=   0
      ForeColorFixed  =   4210752
      GridColor       =   16777215
      ReadOnly        =   -1  'True
      Rows            =   10
      SelectionMode   =   3
      MultiSelect     =   0   'False
      DateFormat      =   2
      EnterKeyMoveTo  =   1
      BackColorComment=   -2147483635
      AllowUserPaste  =   2
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Account Manager"
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
      Height          =   225
      Index           =   4
      Left            =   8160
      MouseIcon       =   "FrmPassword.frx":EAE2
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   11520
      Width           =   1890
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   1
      Left            =   7800
      MouseIcon       =   "FrmPassword.frx":EDEC
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   10560
      Width           =   2655
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   9162.949
      X2              =   9162.949
      Y1              =   0
      Y2              =   11880
   End
   Begin VB.Label lbMenuHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esci"
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   0
      Left            =   8955
      TabIndex        =   17
      Top             =   10080
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   18325.9
      Y1              =   10560
      Y2              =   10560
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "FrmPassword.frx":F0F6
      Height          =   480
      Index           =   1
      Left            =   8880
      MouseIcon       =   "FrmPassword.frx":124D8
      MousePointer    =   99  'Custom
      Picture         =   "FrmPassword.frx":127E2
      Top             =   10920
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "User List"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Index           =   2
      Left            =   9120
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Index           =   1
      Left            =   840
      TabIndex        =   4
      Top             =   4200
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   2760
      Width           =   1740
   End
End
Attribute VB_Name = "frmPreferenze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Operatore As String
Private IndexPrivilege As Long
Private m_rc As Boolean
Private bSeAdmin As Boolean
Private bSeManager As Boolean


Public Function DoShow(Optional ByVal IndPrivilege As Boolean, Optional ByVal User As String, Optional ByVal Passw As String) As Boolean
    
     On Error GoTo ERR_SHOW
    
    m_rc = True
    
    IndexPrivilege = IndPrivilege
    
    Operatore = User
    Pass(0) = User
    Pass(1) = Passw
    
    Call SetGrdAccount


    With dbTabUserAccount
    
    
        .filter = ""
        .filter = ""
        
        If .EOF Then
            Call SettPrimoAvvio
            GrdAccount.Rows = 1
            GoTo start:
        End If
    
    
        .filter = ""
        .filter = "IndexPrivilege=1"
        If .EOF Then
    
            SetFormPrivilege (1)
        Else
            bExistManager = True
        End If
        .filter = ""
        .filter = "IndexPrivilege=2"
        If .EOF Then
            SetFormPrivilege (2)
        Else
            bExistTCO = True
        End If
        .filter = ""
        .filter = "IndexPrivilege=3"
        If .EOF Then
            SetFormPrivilege (3)
        Else
            bExistAdministrator = True
        End If
    End With
    
    
    pass_ok.Value = (GetSetting(App.Title, "Settings", "utilizza_pass", 0))
    
    'Frame1.Enabled = Not (IIf(bLoginAvvio = 0, True, False))
    
    
    Call RiempiGrid
    
    
    If Operatore <> "" Then Call CheckUtente
start:
    
    Me.Show vbModal
    
    If m_rc = True Then
           ' Operatore = MyOperatore ' GetSetting(App.Title, "Settings", "Operatore", "Nominativo operatore")
            MyOperatore.Name = Operatore
            MyOperatore.IndexPrivilege = IndexPrivilege
    End If
    
    DoShow = m_rc
    
ERR_END:
    On Error GoTo 0
    Exit Function
    

ERR_SHOW:

    m_rc = False
    MsgBox err.Description
    Resume ERR_END
 

End Function

Private Sub CheckUtente(Optional ByVal bCambioPrivilegi As Boolean = True)
Dim UserPrivilege As Integer
    If Operatore <> "" Then

        With dbTabUserAccount
            .filter = ""
            .filter = "USERID='" & Operatore & "'"
            If .EOF Then
                SetFormPrivilege 0
                If Not (bExistAdministrator) Then SetFormPrivilege 3
            Else
                Pass(1) = Trim(!Password)
                UserPrivilege = !IndexPrivilege
                If bCambioPrivilegi Then IndexPrivilege = !IndexPrivilege
                SetFormPrivilege (UserPrivilege)
                
                If bCambioPrivilegi Then
                    If IndexPrivilege = 1 Then
                        Label1(3) = "QC Laboratory Manager Account"
                        Check2.Value = 1
                    End If
                    If IndexPrivilege = 2 Then
                        Check3.Value = 1
                        Label1(3) = "QC Tecnical Office (TCO) Account"
                    End If
                    
                    If IndexPrivilege = 3 Then
                        Check1.Value = 1
                        Label1(3) = "QC Administrator Account"
                    End If
                End If
                
            End If
        
        End With
    
    End If

End Sub

Private Sub SetFormPrivilege(ByVal IndPrivilege As Integer)
Dim rc As Boolean
Dim rc2 As Boolean
Dim rc3 As Boolean
Dim rc4 As Boolean
rc = False

Select Case IndPrivilege
    Case 0
        rc = False
        rc2 = False
        rc3 = False
        rc4 = False
    Case 1
        rc = False
        rc2 = True
        rc3 = True
        rc4 = False
        Check2.Value = 1
    Case 2
        rc = False
        rc2 = False
        rc3 = False
        rc4 = True
        Check3.Value = 1
    Case 3
        rc = True
        rc2 = True
        rc3 = True
        rc4 = True
        Check1.Value = 1
End Select
Check3.Visible = rc4
Picture3.Visible = rc4
Check1.Visible = rc
Picture1.Visible = rc
Check2.Visible = rc2
Picture2.Visible = rc2
pass_ok.Visible = rc3
chPass.Visible = rc3

End Sub
Private Sub Check3_Click()
Check2.Value = IIf(Check3.Value = 1, 0, Check2.Value)
Check1.Value = IIf(Check3.Value = 1, 0, Check1.Value)
Check3.ForeColor = IIf(Check3.Value = 1, vbColorOrange, vbWhite)
'IndexPrivilege = IIf(Check3.Value = 1, 2, IndexPrivilege)
End Sub

Private Sub Check1_Click()
Check2.Value = IIf(Check1.Value = 1, 0, Check2.Value)
Check3.Value = IIf(Check1.Value = 1, 0, Check3.Value)
Check1.ForeColor = IIf(Check1.Value = 1, vbColorOrange, vbWhite)
'IndexPrivilege = IIf(Check1.Value = 1, 3, IndexPrivilege)
End Sub
Private Sub Check2_Click()
Check1.Value = IIf(Check2.Value = 1, 0, Check1.Value)
Check3.Value = IIf(Check2.Value = 1, 0, Check3.Value)
Check2.ForeColor = IIf(Check2.Value = 1, vbColorOrange, vbWhite)
'IndexPrivilege = IIf(Check2.Value = 1, 2, IndexPrivilege)
End Sub
Public Sub Com_pass_Click(Index As Integer)
Select Case Index
    Case 0
        With dbTabUserAccount
            .filter = ""
            bExistAccount = IIf(.EOF, False, True)
        End With
        Unload Me
    Case 1
        Call SalvaPass
        
       
    Case 2
        If Pass(0) = "" Then Exit Sub
        If F_MsgBox.DoShow("Delete Account: " & Pass(0) & " ?") Then
           ' Com_pass(2).Enabled = False
            Call cancella
        End If
    Case 3
        Pass(0) = ""
        Pass(1) = ""
        Pass(0).SetFocus
    Case 4
        Pass(0) = ""
        Pass(1) = ""
        Check1.Value = 0
        Check2.Value = 0
        Check3.Value = 0
        IndexPrivilege = 0
        Pass(0).SetFocus
        SetFormPrivilege 0
End Select

Call RiempiGrid

End Sub
Private Sub SalvaPass()
If Pass(0) = "" Then Exit Sub
If Pass(1) = "" Then Exit Sub
    On Error Resume Next
        With dbTabUserAccount
          .filter = ""
          .filter = "USERID='" & Pass(0) & "'"
          If .EOF Then
            .AddNew
            PopupMessage 2, "Add new Account : " & Pass(0)
          End If
          !UserID = Trim(Pass(0))
          !Password = Trim(Pass(1))
          !IndexPrivilege = CheckIndex ' = IIf(Check1.value = 1, 2, IndexPrivilege)
          .Update
        End With
        
    ' Com_pass(2).Enabled = True
    ' Com_pass(1).Enabled = False
End Sub
Private Function CheckIndex() As Long
IndexPrivilege = 0
IndexPrivilege = IIf(Check2.Value = 1, 1, 0)
IndexPrivilege = IIf(Check3.Value = 1, 2, IndexPrivilege)
IndexPrivilege = IIf(Check1.Value = 1, 3, IndexPrivilege)
CheckIndex = IndexPrivilege
End Function

Private Sub cancella()

    With dbTabUserAccount
        .filter = ""
        .filter = "USERID='" & Pass(0) & "'"
        If .EOF Then
            
            Else
            .Delete
        End If
    End With
    
    Pass(0) = ""
    Pass(1) = ""
    IndexPrivilege = 0
   ' Com_pass(2).Enabled = False
   ' Com_pass(1).Enabled = False
                    
End Sub




Private Sub DefaultMenuLabel_Click(Index As Integer)
    Select Case Index
        Case 1
            Com_pass_Click 0
    
    End Select
End Sub

Private Sub Form_Activate()
'vbColorDarkUnabled
Me.Caption = App.EXEName & " :: Imposta Password "
    
End Sub

Private Sub Form_Load()

  'WindowsXPC1.InitSubClassing
  DropShadow Me.hWnd
  
    On Local Error Resume Next
  

End Sub

 
Private Function RiempiGrid()
Dim a As Integer
    GrdAccount.Rows = 1
    GrdAccount.AutoRedraw = False
    With dbTabUserAccount
        .filter = ""
        If .EOF Then
            Pass(0) = GetSetting(App.Title, "Autorizzazione", "ID", "")
            Pass(1) = GetSetting(App.Title, "Autorizzazione", "Pass", "")
            SetFormPrivilege (2)
        Else
        
            .MoveFirst
            For a = 1 To .RecordCount
                With GrdAccount
                    .AddItem "", False
                    .Cell(.Rows - 1, 0).Text = a
                    .Cell(.Rows - 1, 1).Text = Trim(dbTabUserAccount!UserID)
                    .Cell(.Rows - 1, 2).Text = GetPosition(dbTabUserAccount!IndexPrivilege)
                End With
                
                .MoveNext
            Next
                GrdAccount.AutoRedraw = True
                GrdAccount.Refresh
        End If
    End With
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
    Unload Me
End Select
End Sub

Private Sub GrdAccount_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

If FirstRow > 0 Then

    If IndexPrivilege = 3 Then
    
        Operatore = GrdAccount.Cell(FirstRow, 1).Text
        Pass(0) = Operatore
        Pass(1) = ""
        CheckUtente (False)
        
    Else
        PopupMessage 2, "Administrator User can View/Modify data...", , , "Account"
    End If


End If



End Sub

Private Sub Image3_Click(Index As Integer)
    Select Case Index
        Case 0
            Com_pass_Click 1
        Case 1
            Com_pass_Click 4
        Case 2
            Com_pass_Click 2
    End Select
End Sub

Private Sub PicMenu_Click(Index As Integer)
Image3_Click Index
End Sub

Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To PicMenu.Count - 1
    If i = Index Then
        PicMenu(i).BackColor = &H505050
    Else
        PicMenu(i).BackColor = vbColorDarkFont
    End If
Next
End Sub


Private Sub Pass_Change(Index As Integer)
   ' bLoginAvvio.Value = IIf(Len(Pass(Index)) < 1, 0, 1)
   ' Com_pass(1).Enabled = IIf(Len(Pass(Index)) > 0, True, False)
    
End Sub

Private Sub Pass_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
            Case 0
                Pass(1).SetFocus
            Case 1
               If F_MsgBox.DoShow("Save Password?") Then
                    'salvo
                    With dbTabUserAccount
                        .filter = ""
                        .filter = "USERID='" & Pass(0) & "'"
                        If .EOF Then
                            .AddNew
                            !UserID = Trim(Pass(0))
                        End If
                            !Password = Trim(Pass(1))
                            
                    End With
               Else
                    Pass(1) = ""
               End If
    
    End Select

End If
End Sub

Private Sub pass_ok_Click()
SaveSetting App.Title, "Settings", "utilizza_pass", pass_ok
'Com_pass(1).Enabled = True
End Sub


Private Sub SetGrdAccount()
Dim i As Integer
    With GrdAccount
  
  

        .AutoRedraw = False
        .Cols = 3
        .DefaultRowHeight = 35 '* m_ControlGridRowHeight
        
        .Cell(0, 0).Text = "#"
        .Cell(0, 1).Text = "Username"
        .Cell(0, 2).Text = "Position"
        
        .Column(0).Width = 35 '* m_ControlGridColWidth
        .Column(1).Width = 350 '* m_ControlGridColWidth
        .Column(2).Width = 200 '* m_ControlGridColWidth
        For i = 0 To .Cols - 1
        .Cell(0, i).FontSize = 14 ' * m_ControlGridFontSize
        .Cell(0, i).ForeColor = vbWhite
        Next
        .DefaultFont.Size = 12 * m_ControlGridFontSize
        .DefaultFont.Bold = True
        .AutoRedraw = True
        .Refresh
    

        
    
    End With


End Sub



Private Sub SettPrimoAvvio()


SetFormPrivilege 3


End Sub
