VERSION 5.00
Begin VB.Form FormAcquisition 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acquisition"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   16260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   16260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txAcquisition 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   13
      Left            =   3240
      TabIndex        =   34
      Text            =   "-21 %"
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox txAcquisition 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   12
      Left            =   3240
      TabIndex        =   32
      Text            =   "-200,221"
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox txAcquisition 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Index           =   11
      Left            =   6600
      TabIndex        =   30
      Text            =   "1229,998"
      Top             =   3360
      Width           =   4095
   End
   Begin VB.TextBox txAcquisition 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   10
      Left            =   3240
      TabIndex        =   28
      Text            =   "1300,400"
      Top             =   3150
      Width           =   2415
   End
   Begin VB.TextBox txAcquisition 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   13200
      TabIndex        =   26
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txAcquisition 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   11160
      TabIndex        =   24
      Top             =   2040
      Width           =   840
   End
   Begin VB.TextBox txAcquisition 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   6720
      TabIndex        =   22
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txAcquisition 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   1800
      TabIndex        =   20
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txAcquisition 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   13200
      TabIndex        =   18
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txAcquisition 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   6720
      TabIndex        =   16
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txAcquisition 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   1800
      TabIndex        =   14
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txAcquisition 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   13200
      TabIndex        =   12
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txAcquisition 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   6720
      TabIndex        =   10
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "l"
      Height          =   615
      Index           =   5
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   15255
      Begin VB.Label lbInside 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recipe Acquisition"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00105010&
         Height          =   270
         Index           =   5
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   2070
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00B0B0B0&
         X1              =   0
         X2              =   15240
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.TextBox txAcquisition 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1800
      TabIndex        =   6
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Frame frCommandInside 
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      Caption         =   "Image14"
      Height          =   495
      Index           =   1
      Left            =   12720
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
   Begin VB.Frame frCommandInside 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Label lbCommandInside 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
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
         TabIndex        =   3
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.Frame frCommandInside 
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      Caption         =   "Image14"
      Height          =   495
      Index           =   2
      Left            =   6480
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Label lbCommandInside 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Product Classification"
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
         Index           =   2
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.Label lbAcquisition 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Variance %"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   2040
      TabIndex        =   35
      Top             =   4440
      Width           =   1005
   End
   Begin VB.Label lbAcquisition 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Variance"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   2250
      TabIndex        =   33
      Top             =   3840
      Width           =   825
   End
   Begin VB.Label lbAcquisition 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actual Weight"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   6600
      TabIndex        =   31
      Top             =   3000
      Width           =   1305
   End
   Begin VB.Label lbAcquisition 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Theoretical Weight"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   1365
      TabIndex        =   29
      Top             =   3120
      Width           =   1710
   End
   Begin VB.Label lbAcquisition 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Package"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   12240
      TabIndex        =   27
      Top             =   2040
      Width           =   795
   End
   Begin VB.Label lbAcquisition 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Week Delivery"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   9675
      TabIndex        =   25
      Top             =   2040
      Width           =   1305
   End
   Begin VB.Label lbAcquisition 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty delivered"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   5385
      TabIndex        =   23
      Top             =   2040
      Width           =   1230
   End
   Begin VB.Label lbAcquisition 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   465
      TabIndex        =   21
      Top             =   2040
      Width           =   1230
   End
   Begin VB.Label lbAcquisition 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer Lot "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   11505
      TabIndex        =   19
      Top             =   1560
      Width           =   1590
   End
   Begin VB.Label lbAcquisition 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer Code"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4860
      TabIndex        =   17
      Top             =   1560
      Width           =   1755
   End
   Begin VB.Label lbAcquisition 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   360
      TabIndex        =   15
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lbAcquisition 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cas"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   12360
      TabIndex        =   13
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lbAcquisition 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   4920
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00B0B0B0&
      X1              =   480
      X2              =   15720
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lbAcquisition 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM Code"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "FormAcquisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private UserChCode As String
Private m_rc As Boolean
Private MyID As Long
Private Acquisition As PrepAcquisition
Private AcquisitionClean As PrepAcquisition
Private RecipeCode As String



       '.Cell(0, 1).Text = "Code"
       ' .Cell(0, 2).Text = "Description"
       ' .Cell(0, 3).Text = "CAS"
       ' .Cell(0, 4).Text = "Theorethical weight (g)"
       ' .Cell(0, 5).Text = "Real Weight (g)"
       ' .Cell(0, 6).Text = "Variance (g)"
       ' .Cell(0, 7).Text = "Variance %"
       '
       ' .Cell(0, 8).Text = "Manufacturer"
       ' .Cell(0, 9).Text = "Manufacturer Code"
       ' .Cell(0, 10).Text = "Manufacturer Lot"
       ' .Cell(0, 11).Text = "Delivery Date"
       ' .Cell(0, 12).Text = "Qty Delivered"
       ' .Cell(0, 13).Text = "Week Delivery"
       ' .Cell(0, 14).Text = "Package"
       '
       ' .Cell(0, 15).Text = "Note"
       ' .Cell(0, 16).Text = "bMix"
        

Public Function DoShow(Optional ByRef CHCode As String, Optional FormTop As Double, Optional ByVal rCode As String) As Boolean
Dim rc As Boolean
    On Error GoTo ERR_SHOW
    
    m_rc = False
    mOk
    Acquisition = AcquisitionClean
    FormTop = Screen.Height / 2 - Me.Height / 2
    Me.Top = FormTop + 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    
    
    RecipeCode = rCode
    lbInside(5) = RecipeCode & " : Acquisitions"
  
 
    Me.Show vbModal
    
    

    
    If m_rc = True Then
        CHCode = UserChCode
    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    Resume ERR_END
End Function



Private Sub frCommandInside_Click(Index As Integer)
Select Case Index
    Case 0
        m_rc = IIf(Len(UserChCode) > 0, True, False)
    Case 1
        m_rc = False
    Case 2
        Call F_PICTOGRAM.DoShow(MyID, 1)
        Exit Sub
End Select

Unload Me


End Sub


Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub



Private Sub txFormulation_Click(Index As Integer)

End Sub


Private Sub FillUserRMCode(ByVal userCode As String)

End Sub


