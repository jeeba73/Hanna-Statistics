VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form FormCodes 
   BackColor       =   &H00886010&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15390
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   15390
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frInside 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "&H00F0F0F0&"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   15255
      Begin VB.ComboBox cmbLine 
         Height          =   375
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   5520
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Frame frCommandInside 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   9000
         TabIndex        =   10
         Top             =   5520
         Visible         =   0   'False
         Width           =   3015
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Select"
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
            TabIndex        =   11
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00886010&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   5400
         TabIndex        =   6
         Top             =   2280
         Visible         =   0   'False
         Width           =   5055
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empty List..."
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
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   7
            Top             =   555
            Width           =   1155
         End
      End
      Begin VB.Frame frCommandInside 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Caption         =   "Image14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   12120
         TabIndex        =   4
         Top             =   5520
         Width           =   3015
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
            TabIndex        =   5
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "l"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   15255
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   120
            X2              =   15120
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Database"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00644603&
            Height          =   255
            Left            =   14160
            TabIndex        =   3
            Top             =   240
            Width           =   885
         End
         Begin VB.Label lbInside 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hanna Codes Database "
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00644603&
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   120
            Width           =   2835
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00E0E0E0&
            X1              =   120
            X2              =   15120
            Y1              =   480
            Y2              =   480
         End
      End
      Begin FlexCell.Grid Grid1 
         Height          =   4575
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   720
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   8070
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColor1      =   15790320
         BackColor2      =   15790320
         BackColorBkg    =   15790320
         BackColorFixed  =   15790320
         BackColorFixedSel=   15790320
         BackColorScrollBar=   15592423
         BorderColor     =   15790320
         CellBorderColor =   15790320
         CellBorderColorFixed=   15790320
         Cols            =   5
         DefaultFontName =   "Segoe UI"
         DefaultFontSize =   8.25
         DisplayRowIndex =   -1  'True
         DrawMode        =   1
         DefaultRowHeight=   20
         FixedRowColStyle=   0
         ForeColorFixed  =   6571523
         GridColor       =   15790320
         Rows            =   1
         ScrollBarStyle  =   0
         SelectionMode   =   3
         MultiSelect     =   0   'False
         DateFormat      =   0
      End
      Begin VB.TextBox txSearch 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   5520
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line"
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   5520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   15120
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Code"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   5560
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FormCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyID As Long
Private m_rc As Boolean
Private uHannaCode As String

Public Function DoShow(Optional ByRef UserHannaCode As String, Optional FormTop As Double, Optional ByRef ID As Long) As Boolean

    On Error GoTo ERR_SHOW
    
    m_rc = False
    mOk
    If FormTop = 0 Then FormTop = Screen.Height / 2 - Me.Height / 2
    Me.Top = FormTop + 300
    Me.Left = Screen.Width / 2 - Me.Width / 2

    Call SetHannaCodeGrid(Grid1)
    Call SetLine(cmbLine, True)

    Me.Show vbModal

    If m_rc = True Then
        UserHannaCode = uHannaCode
        ID = MyID
    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    Resume ERR_END
End Function

Private Sub cmbLine_Click()
InsertCode Grid1
End Sub


Private Sub InsertCode(ByRef Grid1 As Grid)
 Call GetHannaCodeFromDatabase(Grid1, False, , cmbLine)
End Sub

Private Sub frCommandInside_Click(Index As Integer)
Select Case Index
    Case 0
        m_rc = IIf(MyID > 0, True, False)
    Case 1
        m_rc = False
End Select

Unload Me


End Sub

Private Sub Grid1_DblClick()
If MyID > 0 Then frCommandInside_Click 0
End Sub

Private Sub Grid1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

MyID = 0

frCommandInside(0).Visible = False
If FirstRow > 0 Then
    MyID = Grid1.Cell(FirstRow, 11).Text
    uHannaCode = Trim(Grid1.Cell(FirstRow, 1).Text)
   
    frCommandInside(0).Visible = True
End If



End Sub

Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub



Private Function SearcMRInTable(ByVal rc As Boolean)
 Call MRSearchInGrid(Grid1, txSearch, rc)
End Function

Private Sub txSearch_Change()
Dim rc As Boolean
rc = IIf(Len(txSearch) > 0, False, True)
SearcMRInTable rc
End Sub

