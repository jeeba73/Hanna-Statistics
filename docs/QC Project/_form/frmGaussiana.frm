VERSION 5.00
Begin VB.Form F_GRAPH 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   12000
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   19200
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGaussiana.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmGaussiana.frx":6852
   ScaleHeight     =   12000
   ScaleMode       =   0  'User
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   2160
      Top             =   10800
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00964901&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8175
      Left            =   1440
      MouseIcon       =   "frmGaussiana.frx":2472B
      MousePointer    =   99  'Custom
      TabIndex        =   44
      Top             =   2640
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open pdf folder"
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
         Index           =   6
         Left            =   0
         MouseIcon       =   "frmGaussiana.frx":24A35
         MousePointer    =   99  'Custom
         TabIndex        =   54
         Top             =   6840
         Width           =   19185
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   4
         Left            =   8760
         TabIndex        =   53
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   5
         Left            =   9360
         MouseIcon       =   "frmGaussiana.frx":24D3F
         MousePointer    =   99  'Custom
         Picture         =   "frmGaussiana.frx":25049
         Top             =   6240
         Width           =   480
      End
      Begin VB.Label lbPrint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12345566"
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
         Index           =   2
         Left            =   9600
         TabIndex        =   52
         Top             =   4920
         Width           =   3360
      End
      Begin VB.Label lbPrint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12345566"
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
         Left            =   9600
         TabIndex        =   51
         Top             =   4320
         Width           =   3360
      End
      Begin VB.Label lbPrint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12345566"
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
         Left            =   9600
         TabIndex        =   50
         Top             =   3720
         Width           =   3360
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Number"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   2
         Left            =   6285
         TabIndex        =   49
         Top             =   4920
         Width           =   2970
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code SFG"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   1
         Left            =   6600
         TabIndex        =   48
         Top             =   4320
         Width           =   2655
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Number"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   0
         Left            =   6600
         TabIndex        =   47
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Graph QC : Average"
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
         Left            =   0
         TabIndex        =   46
         Top             =   2520
         Width           =   19140
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Printing pdf file..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   885
         Left            =   0
         TabIndex        =   45
         Top             =   1560
         Width           =   19140
      End
   End
   Begin ChemicalQC.Graph Graph1 
      Height          =   8175
      Left            =   0
      TabIndex        =   40
      Top             =   2520
      Width           =   19215
      _ExtentX        =   33893
      _ExtentY        =   14420
      State           =   "frmGaussiana.frx":2842B
   End
   Begin VB.PictureBox PicMainMenu 
      BackColor       =   &H00303030&
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
      Height          =   1095
      Index           =   4
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   19215
      TabIndex        =   19
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   3
         Left            =   11160
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   70
         Top             =   660
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   2
         Left            =   10080
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   69
         Top             =   660
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   1
         Left            =   9000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   68
         Top             =   660
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   200
         Index           =   0
         Left            =   7920
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   63
         Top             =   660
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00303030&
         Caption         =   "Show Value"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   5640
         TabIndex        =   43
         Top             =   180
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00303030&
         Caption         =   "Show Bar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   345
         Left            =   5640
         TabIndex        =   42
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00303030&
         Caption         =   "Show Meter Color"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   7920
         TabIndex        =   41
         Top             =   180
         Width           =   2775
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MouseIcon       =   "frmGaussiana.frx":284E7
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   22
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   735
            MouseIcon       =   "frmGaussiana.frx":287F1
            MousePointer    =   99  'Custom
            Picture         =   "frmGaussiana.frx":28AFB
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Average"
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
            Left            =   60
            MouseIcon       =   "frmGaussiana.frx":2BEDD
            MousePointer    =   99  'Custom
            TabIndex        =   23
            Top             =   720
            Width           =   1875
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   1920
         MouseIcon       =   "frmGaussiana.frx":2C1E7
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   20
         Top             =   0
         Width           =   1935
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gaussian"
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
            Left            =   105
            MouseIcon       =   "frmGaussiana.frx":2C4F1
            MousePointer    =   99  'Custom
            TabIndex        =   21
            Top             =   720
            Width           =   1800
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   735
            MouseIcon       =   "frmGaussiana.frx":2C7FB
            MousePointer    =   99  'Custom
            Picture         =   "frmGaussiana.frx":2CB05
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Meter4"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   11520
         TabIndex        =   67
         Top             =   660
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Meter3"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   10440
         TabIndex        =   66
         Top             =   660
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Meter2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   9360
         TabIndex        =   65
         Top             =   660
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Meter1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   8280
         TabIndex        =   64
         Top             =   660
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Graph QC : Average"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   450
         Left            =   14685
         TabIndex        =   24
         Top             =   360
         Width           =   4125
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   8280
      ScaleHeight     =   1335
      ScaleWidth      =   1215
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frame_DATI 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   19215
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Height          =   285
         Index           =   2
         Left            =   2205
         TabIndex        =   55
         Top             =   960
         Width           =   2295
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   4800
         ScaleHeight     =   975
         ScaleWidth      =   5655
         TabIndex        =   34
         Top             =   240
         Width           =   5655
         Begin VB.CheckBox Check1 
            Caption         =   "All Readings"
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
            Left            =   3600
            TabIndex        =   37
            Top             =   300
            Width           =   2175
         End
         Begin VB.ComboBox cmbSTD 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   220
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00964901&
            Caption         =   "Standard "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   120
            TabIndex        =   36
            Top             =   225
            Width           =   2100
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Height          =   285
         Index           =   0
         Left            =   2200
         TabIndex        =   31
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Height          =   285
         Index           =   1
         Left            =   2200
         TabIndex        =   30
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00964901&
         Caption         =   "Recipe  "
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
         Index           =   7
         Left            =   240
         TabIndex        =   56
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lbl_ris 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   4
         Left            =   15660
         TabIndex        =   10
         Top             =   960
         Width           =   840
      End
      Begin VB.Label lbl_ris 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   3
         Left            =   15660
         TabIndex        =   9
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lbl_ris 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   1
         Left            =   15660
         TabIndex        =   8
         Top             =   600
         Width           =   840
      End
      Begin VB.Label lbl_ris 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   6
         Left            =   17600
         TabIndex        =   39
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00964901&
         Caption         =   "# Selected Test  "
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
         Index           =   10
         Left            =   13680
         TabIndex        =   38
         Top             =   240
         Width           =   1950
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00964901&
         Caption         =   "Lot Number  "
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
         Index           =   9
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00964901&
         Caption         =   "Code SFG  "
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
         Index           =   5
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00964901&
         Caption         =   "# Out of Range  "
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
         Index           =   4
         Left            =   10500
         TabIndex        =   17
         Top             =   600
         Width           =   2430
      End
      Begin VB.Label lbl_ris 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   2
         Left            =   12960
         TabIndex        =   16
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004080&
         Caption         =   "Max  "
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
         Index           =   8
         Left            =   16530
         TabIndex        =   13
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lbl_ris 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   7
         Left            =   17600
         TabIndex        =   12
         Top             =   600
         Width           =   840
      End
      Begin VB.Label lbl_ris 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   5
         Left            =   12960
         TabIndex        =   11
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl_ris 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   0
         Left            =   12960
         TabIndex        =   7
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004080&
         Caption         =   "Min  "
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
         Index           =   6
         Left            =   16530
         TabIndex        =   6
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00964901&
         Caption         =   "STD Value  "
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
         Index           =   3
         Left            =   13680
         TabIndex        =   5
         Top             =   600
         Width           =   1950
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00964901&
         Caption         =   "% Out of Range  "
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
         Index           =   2
         Left            =   10500
         TabIndex        =   4
         Top             =   960
         Width           =   2430
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00964901&
         Caption         =   "Mean Value  "
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
         Index           =   1
         Left            =   13680
         TabIndex        =   3
         Top             =   960
         Width           =   1950
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00964901&
         Caption         =   "# Selected Readings  "
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
         Index           =   0
         Left            =   10500
         TabIndex        =   2
         Top             =   240
         Width           =   2430
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
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
      Height          =   3735
      Left            =   0
      MouseIcon       =   "frmGaussiana.frx":2FEE7
      MousePointer    =   99  'Custom
      ScaleHeight     =   3735
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   2520
      Width           =   4335
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STD"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Width           =   450
      End
      Begin VB.Label lbl_VALUE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom Out"
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
      Index           =   5
      Left            =   4950
      MouseIcon       =   "frmGaussiana.frx":301F1
      MousePointer    =   99  'Custom
      TabIndex        =   62
      Top             =   11640
      Width           =   840
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom  In"
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
      Left            =   3720
      MouseIcon       =   "frmGaussiana.frx":304FB
      MousePointer    =   99  'Custom
      TabIndex        =   61
      Top             =   11640
      Width           =   720
   End
   Begin VB.Image PicZOOM 
      Height          =   480
      Index           =   0
      Left            =   3840
      MouseIcon       =   "frmGaussiana.frx":30805
      MousePointer    =   99  'Custom
      Picture         =   "frmGaussiana.frx":30B0F
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image PicZOOM 
      Height          =   480
      Index           =   1
      Left            =   5160
      MouseIcon       =   "frmGaussiana.frx":33EF1
      MousePointer    =   99  'Custom
      Picture         =   "frmGaussiana.frx":341FB
      Top             =   11040
      Width           =   480
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move Forward"
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
      Left            =   17760
      MouseIcon       =   "frmGaussiana.frx":375DD
      MousePointer    =   99  'Custom
      TabIndex        =   60
      Top             =   11640
      Width           =   1200
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move Previous"
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
      Left            =   15675
      MouseIcon       =   "frmGaussiana.frx":378E7
      MousePointer    =   99  'Custom
      TabIndex        =   59
      Top             =   11640
      Width           =   1230
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print Chart"
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
      Left            =   270
      MouseIcon       =   "frmGaussiana.frx":37BF1
      MousePointer    =   99  'Custom
      TabIndex        =   58
      Top             =   11640
      Width           =   900
   End
   Begin VB.Label L 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Graph"
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
      Left            =   9000
      MouseIcon       =   "frmGaussiana.frx":37EFB
      MousePointer    =   99  'Custom
      TabIndex        =   57
      Top             =   11640
      Width           =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   0
      X2              =   19200
      Y1              =   10695
      Y2              =   10695
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   2
      Left            =   15240
      MouseIcon       =   "frmGaussiana.frx":38205
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   10800
      Width           =   2175
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   0
      Left            =   17640
      MouseIcon       =   "frmGaussiana.frx":3850F
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   10800
      Width           =   1575
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1095
      Index           =   1
      Left            =   7560
      MouseIcon       =   "frmGaussiana.frx":38819
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   10800
      Width           =   4335
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   3
      Left            =   0
      MouseIcon       =   "frmGaussiana.frx":38B23
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   10680
      Width           =   2295
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "frmGaussiana.frx":38E2D
      Height          =   480
      Index           =   0
      Left            =   18240
      MouseIcon       =   "frmGaussiana.frx":3C20F
      MousePointer    =   99  'Custom
      Picture         =   "frmGaussiana.frx":3C519
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   480
      MouseIcon       =   "frmGaussiana.frx":3F8FB
      MousePointer    =   99  'Custom
      Picture         =   "frmGaussiana.frx":3FC05
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "frmGaussiana.frx":42FE7
      Height          =   480
      Index           =   1
      Left            =   9360
      MouseIcon       =   "frmGaussiana.frx":463C9
      MousePointer    =   99  'Custom
      Picture         =   "frmGaussiana.frx":466D3
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   15960
      MouseIcon       =   "frmGaussiana.frx":49AB5
      MousePointer    =   99  'Custom
      Picture         =   "frmGaussiana.frx":49DBF
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   4
      Left            =   5880
      MouseIcon       =   "frmGaussiana.frx":4D1A1
      MousePointer    =   99  'Custom
      Picture         =   "frmGaussiana.frx":4D4AB
      Top             =   11040
      Visible         =   0   'False
      Width           =   480
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
      Index           =   2
      Visible         =   0   'False
      X1              =   14400
      X2              =   14400
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
      Left            =   9435
      TabIndex        =   25
      Top             =   10200
      Visible         =   0   'False
      Width           =   390
   End
End
Attribute VB_Name = "F_GRAPH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const Filename As String = "C:\Woof.dat"
Private Const FILENAME_PIC As String = "C:\Woof.bmp"

Private Const STYLE_MUSIC As Long = 0
Private Const STYLE_TASKMAN As Long = 1
Private Const STYLE_GRAPH As Long = 2
Private Const STYLE_CUSTOM As Long = 3
Private Const STYLE_BAR As Long = 4
Private Const STYLE_LINE As Long = 5
Private Const STYLE_PROGRESS As Long = 6
Private i As Integer
Private mlngStyle   As Long



Private Const GRID_SPACING = 120
Private Const HANDLE_WIDTH = 80
Private Const HANDLE_HALF_WIDTH = HANDLE_WIDTH / 2
Public Opt_GRAPHico As String
Public Sel_X
Public Sel_Y
Public NumReadings As Integer
Public sFilename As String
Public INDEX_G As Integer

Private TipoGrafico As Integer
Private bActivate As Boolean
Private m_rc As Boolean
Private cMedia  As Double
Private cDeviazione As Double
Private cTotali As Integer
Private cUm As String
Private cNominale As Double
Private cQMR As Double

Private MaxDev As Double
Private x_height As Double
Private y_height As Double
Private xoffset As Double
Private percent_data_x As Double
Private yoffset As Double
Private percent_data_y As Double
Private xData As Double
Private yData As Double
Private Zero As Integer
Private pltColor As OLE_COLOR
Private x_data_screen As Double
Private y_data_screen As Double
Private StartTime As String
Private StopTime As String
Private cStep
Private IndexFormProcedura As Integer
Private MyLot As String
Private MyCode As String
Private bAllValue As Boolean
Private STDCount As Integer
Private STD() As String
Private MeterNumber As Integer
Private virgola As Integer
Private UserDecimal As String
Private MeasurementUnit As String
Private MaxY(10) As Double
Private MaxYGauss(10) As Double
Private MaxValue(10) As Double
Private X0Value(10) As Double
Private MyHeight As Long
Private MyWidth As Long
Private bPrintMe As Boolean

Private Sub SaveSizes()
Dim i As Integer
Dim Ctl As Control
' Save the controls' positions and sizes.
On Error GoTo ERR_SAVE
ReDim m_ControlPositions(1 To Controls.Count)
i = 1
For Each Ctl In Controls
    With m_ControlPositions(i)
        If TypeOf Ctl Is Line Then
            .Left = Ctl.X1
            .Top = Ctl.Y1
            .Width = Ctl.X2 - Ctl.X1
            .Height = Ctl.Y2 - Ctl.Y1
        ElseIf TypeOf Ctl Is Menu Then
        ElseIf TypeOf Ctl Is Inet Then
        ElseIf TypeOf Ctl Is Timer Then
        

        Else
            .Left = Ctl.Left
            'MsgBox (TypeName(ctl))
            .Top = Ctl.Top
            .Width = Ctl.Width
            .Height = Ctl.Height
            On Error Resume Next
            .FontSize = Ctl.Font.Size
            
            'MsgBox (TypeName(ctl))
            On Error GoTo 0
        End If
    End With
    i = i + 1
Next Ctl
' Save the form's size.
ERR_END:
On Error GoTo 0

m_FormWid = ScaleWidth
m_FormHgt = ScaleHeight
Exit Sub
ERR_SAVE:
Resume Next
End Sub



Private Sub ResizeControls()
Dim i As Integer
Dim Ctl As Control
Dim x_scale As Single
Dim y_scale As Single
' Don't bother if we are minimized.
On Error GoTo ERR_SAVE
If WindowState = vbMinimized Then Exit Sub
' Get the form's current scale factors.
x_scale = ScaleWidth / m_FormWid
y_scale = ScaleHeight / m_FormHgt
' Position the controls.
i = 1

m_ControlGridFontSize = y_scale
m_ControlGridColWidth = x_scale
m_ControlGridRowHeight = y_scale



For Each Ctl In Controls
    With m_ControlPositions(i)
        If TypeOf Ctl Is Line Then
            Ctl.X1 = x_scale * .Left
            Ctl.Y1 = y_scale * .Top
            Ctl.X2 = Ctl.X1 + x_scale * .Width
            Ctl.Y2 = Ctl.Y1 + y_scale * .Height
        ElseIf TypeOf Ctl Is Timer Then
        ElseIf TypeOf Ctl Is Inet Then
        ElseIf TypeOf Ctl Is Grid Then
           Ctl.Left = x_scale * .Left
            Ctl.Top = y_scale * .Top
            Ctl.Width = x_scale * .Width
            Ctl.Height = y_scale * .Height

        Else
            Ctl.Left = x_scale * .Left
           ' MsgBox (TypeName(Ctl))
            Ctl.Top = y_scale * .Top
            Ctl.Width = x_scale * .Width
            If Not (TypeOf Ctl Is ComboBox) Then
                ' Cannot change height of ComboBoxes.
                Ctl.Height = y_scale * .Height
            End If
            On Error Resume Next
            Ctl.Font.Size = y_scale * .FontSize
            On Error GoTo 0
        End If
    End With
    i = i + 1
Next Ctl
Exit Sub
ERR_SAVE:
Resume Next
End Sub

Private Sub Check1_Click()
bAllValue = IIf(Check1.Value = 1, True, False)
Check1.ForeColor = IIf(bAllValue, vbColorOrange, &H964901)
GetMaxY (INDEX_STD)
Call Dati_Lotto(INDEX_STD)
End Sub

Private Sub Check2_Click()
Dim rc As Boolean
Graph1.ShowBars = IIf(Check2.Value = 1, True, False)
rc = IIf(Check2.Value = 1, True, False)
SaveSetting App.Title, "Graph QC", "Check2", Check2.Value
Check2.ForeColor = IIf(Check2.Value = 1, vbColorOrange, vbWhite)
Check3.Visible = IIf(Check2.Value = 1, True, False)
If rc = False Then
    
Else
    rc = IIf(Check3.Value = 1, True, False)
    
End If
ViewMeterLegend rc
End Sub


Private Sub ViweMeterAfterGaussian()
Dim rc As Boolean
rc = IIf(Check2.Value = 1, True, False)
If rc = False Then
Else
    rc = IIf(Check3.Value = 1, True, False)
End If
ViewMeterLegend rc
End Sub





Private Sub Check3_Click()
Dim i As Integer
Dim rc As Boolean
Graph1.ShowBarsColor = IIf(Check3.Value = 1, True, False)
SaveSetting App.Title, "Graph QC", "Check3", Check3.Value
Check3.ForeColor = IIf(Check3.Value = 1, vbColorOrange, vbWhite)

rc = IIf(Check3.Value = 1, True, False)
ViewMeterLegend rc
End Sub

Private Sub ViewMeterLegend(ByVal bValue As Boolean)


    For i = 0 To 3
        If i <= MeterNumber - 1 Then
            Picture5(i).Visible = bValue
            Label9(i).Visible = bValue
        Else
            Picture5(i).Visible = False
            Label9(i).Visible = False
        End If
    
    Next

End Sub

Private Sub Check4_Click()
Graph1.ShowValue = IIf(Check4.Value = 1, True, False)
SaveSetting App.Title, "Graph QC", "Check4", Check4.Value
Check4.ForeColor = IIf(Check4.Value = 1, vbColorOrange, vbWhite)
End Sub

Private Sub DefaultMenuLabel_Click(Index As Integer)
Dim MyHeight As Long
Dim MyWidth As Long

Dim MyIndex As Integer
Select Case Index
    Case 0
        ' vai avanti
        If IndexFormProcedura = PicMenu.Count - 1 Then
            MyIndex = 0
        Else
            MyIndex = IndexFormProcedura + 1
        End If
        PicMenu_Click MyIndex
    Case 1
        'If bFormSaved Then
             MyGraphicCheck = MyGraphicCheckClean
            Unload Me
       ' Else
            
       ' End If
    Case 2
        ' torna indietro
        If IndexFormProcedura = 0 Then
            MyIndex = PicMenu.Count - 1
        Else
            MyIndex = IndexFormProcedura - 1
        End If
        PicMenu_Click MyIndex
    Case 3
    
        Call PrintGraph
   
    Case 4
 
        OpenWithDefault USER_DOCUMENTI & "\" & PathReport
    Case 5
      
    Case 6
       ' Label6_Click
End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37
        DefaultMenuLabel_Click 2
    Case 39
        DefaultMenuLabel_Click 0
End Select
End Sub

Private Sub Form_Load()

Dim i As Integer
If Screen.Width - Me.Width > 1000 And bFullScreen Then
    Me.WindowState = 2
   ' For i = 0 To PicMain.Count - 1
       ' PicMain(i).Picture = LoadPicture(PictureMaxScreen)
        
    'Next '
    Me.Picture = LoadPicture(PictureMaxScreen)
End If
End Sub

Private Sub Frame1_Click()
Frame1.Visible = False
End Sub

Private Sub Image3_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub Label3_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub PicMenu_Click(Index As Integer)

If IndexFormProcedura = Index Then
Else
    Call SelectProcedura(Index)
End If
End Sub
Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To PicMenu.Count - 1
    If i = IndexFormProcedura Then
        ' lascialo stare...
    ElseIf i = Index Then
        PicMenu(i).BackColor = &H505050
    Else
        PicMenu(i).BackColor = vbColorDarkFont
    End If
Next
End Sub

Private Function SelectProcedura(ByVal Index As Integer) As Boolean
Dim rc As Boolean
Dim i As Integer
If Index > 3 Then Exit Function
For i = 0 To PicMenu.Count - 1
    If i = Index Then
        PicMenu(i).BackColor = vbColorForeFixed
    Else
        PicMenu(i).BackColor = vbColorDarkFont
    End If
Next
blTable = "Grapf QC : " & Label3(Index)
Select Case Index
    Case 0
        cmbSTD_Click
        'Picture4(0).BackColor = &H8000&
        PdfFileName = FormatNomeFile_ATP(Text1(0) & "_" & Text1(1) & "_ST" & cmbSTD & "_QCMean")
        
        rc = True
    Case 1
       'Picture4(0).BackColor = &H6000&
       PdfFileName = FormatNomeFile_ATP(Text1(0) & "_" & Text1(1) & "_ST" & cmbSTD & "_QCGauss")
        Gaussiana INDEX_STD
        rc = False
    Case 2
       ' Picture4(0).BackColor = &H5000&
        rc = False
End Select
For i = 0 To 4
   ' Text1(i).Locked = Not (rc)
Next
Picture4.Visible = rc
Check2.Visible = rc
Check3.Visible = IIf(Check2.Value = 1, rc, False)
Check4.Visible = rc

If rc = False Then
    ViewMeterLegend False
Else

    ViweMeterAfterGaussian
End If


'Label2(4) = Label3(Index)
IndexFormProcedura = Index
'PicMain(Index).Visible = True
'PicMain(Index).ZOrder
'blTable = "Grapgh QC : " & Label2(IndexFormProcedura)
'Cleanform (False)
    lbPrint(0) = Trim(Text1(0))
    lbPrint(1) = Trim(Text1(1))
    lbPrint(2) = Trim(cmbSTD)
    Label6 = blTable
    
End Function

Private Sub Cleanform(ByVal bValue As Boolean, Optional ByVal Index As Integer = 0)

End Sub



Public Function DoShow(ByRef Index As Integer, Optional ByRef sLot As String, Optional ByRef sCode As String, Optional ByVal lngID As Long, Optional ByVal MyImage As Image, Optional ByVal Filename As String, Optional ByVal STDNumber As String) As Boolean
    Dim rc As Boolean
    Dim m_FlgLoading As Boolean
    Dim mIndex As Integer
    
    On Error GoTo ERR_SHOW
    
    m_rc = False
    
    'If NumPesate = 0 Then GoTo ERR_END
    bActivate = True
    Picture1.Cls
    INDEX_G = 0
    INDEX_STD = 1
    absIndex = 1
    
    rc = GetSTDValueForGraph
    If Not (rc) Then
        Exit Function
    End If

    If sLot <> "" And sCode <> "" Then
        Call GetCodeInformation(sLot, sCode, lngID)
    Else
        PopupMessage 2, "Please select a valid Code/Lot..."
        Unload Me
    End If
    
    If STDNumber <> "" Then
        cmbSTD = "    " & STDNumber
    Else
        Dati_Lotto 1
    End If
    SelectProcedura 0

    mOk

    

Me.Show vbModal

    If m_rc = True Then
         
        sLot = MyLot
        sCode = MyCode
    End If
    
    DoShow = m_rc
    
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_SHOW:
    MsgBox err.Description
    m_rc = False
    Resume ERR_END
End Function



Private Sub Form_Activate()
    DropShadow Me.hWnd
    frame_DATI.BackColor = vbColorTextLightBlue '&H8000000D '
    Picture4.BackColor = frame_DATI.BackColor
    Check1.BackColor = Picture4.BackColor
    
    
    Picture5(3).BackColor = &H964901
    Picture5(0).BackColor = &HA65911
    Picture5(1).BackColor = &HEBC99B
    Picture5(2).BackColor = &H8000000D


  Frame1.Move Graph1.Left, Graph1.Top, Graph1.Width, Graph1.Height
  Frame1.ZOrder
End Sub
Private Sub Form_Initialize()


Check2.Value = GetSetting(App.Title, "Graph QC", "Check2", 1)
Check3.Value = GetSetting(App.Title, "Graph QC", "Check3", 0)
Check3_Click
Check4.Value = GetSetting(App.Title, "Graph QC", "Check4", 0)

SaveSizes

End Sub






Public Sub Dati_Lotto(Optional ByVal Index As Integer)


On Error GoTo ERR_DATI:

Dim cx As Double
Dim a As Integer
Dim t As Integer
Dim s As Integer
Dim nOutOfRange As Integer
'Index = 1
cMedia = 0
cDeviazione = 0
t = 0
s = 0
nOutOfRange = 0
Dim Tolerance As Double
Dim ToleranceNeg As Double
Dim PointValue As Double
Dim RealValue As Double


    MyChemicalQC.DataControllo(Index).STDRef = CDbl(STD(absIndex, 1))

   Tolerance = CDbl(STD(absIndex, 3)) ' MyChemicalQC.DataControllo(Index).STDRef + MyChemicalQC.DataControllo(Index).s3
   ToleranceNeg = CDbl(STD(absIndex, 2)) 'MyChemicalQC.DataControllo(Index).STDRef - MyChemicalQC.DataControllo(Index).s3
   
   If bAllValue Then
        Label1(0) = "# All Readings  "
        Label1(10) = "# All Tests  "
   Else
        Label1(0) = "# Selected Readings  "
        Label1(10) = "# Selected Tests  "
   End If
   
   
   'SetStyle 3, -Tolerance * 2, Tolerance * 2
   
    SetStyle 3, -MaxY(Index), MaxY(Index)
   
   
   
    Graph1.Redraw = False
   
       
    Graph1.s1 = MyChemicalQC.DataControllo(Index).s
    Graph1.s2 = MyChemicalQC.DataControllo(Index).s2
    
               
   
    With MyChemicalQC


    '-----------------------------------------------------------------------------
    '
    '            calcolo media
    '
    '------------------------------------------------------------------------------
        NumReadings = .STDtest(Index).MaxReadings
    
         For a = 1 To NumReadings
         
            If (.STDtest(Index).Readings(a).Value) = "" Then
                 Debug.Print "dato nullo"
            Else
                If (.STDtest(Index).Readings(a).Value) < 0 Then
                    Debug.Print "dato negativo"
                Else
                    If (.STDtest(Index).Readings(a).bSelectedValue) Or bAllValue Then
                        cx = .STDtest(Index).Readings(a).Value
                        t = t + 1
                        cMedia = cMedia + cx
                       ' MyGraphicCheck.STDtest(Index) = .STDtest(Index)
                        If cx > Tolerance Or cx < ToleranceNeg Then
                            nOutOfRange = nOutOfRange + 1
                        End If
                        PointValue = CDbl(.STDtest(Index).Readings(a).Value) - CDbl(.DataControllo(Index).STDRef)
                        RealValue = CDbl(.STDtest(Index).Readings(a).Value)
                        If MaxValue(Index) < RealValue Then
                            MaxValue(Index) = RealValue
                        End If
                        
                        AddPoint PointValue, CInt(.STDtest(Index).Readings(a).Meter), RealValue
                      
                    Else
                        Debug.Print "dato non registrato"
                    End If
                End If
            End If
         Next
         
         Graph1.Redraw = True
         If t = 0 Then Exit Sub
         
         
         
       ' StopTime = .STDTest(a - 1).Time
         cMedia = cMedia / t
        .NumReadings = t
        
        .DataControllo(Index).media = cMedia
        
    '-----------------------------------------------------------------------------
    '
    '            calcolo deviazione standard
    '
    '------------------------------------------------------------------------------
        s = 0
        MaxDev = 0
         For a = 1 To NumReadings
            If (.STDtest(Index).Readings(a).Value) = "" Then
            Else
                 If (.STDtest(Index).Readings(a).Value) <= 0 Then
                Else
                    If (.STDtest(Index).Readings(a).bSelectedValue) Or bAllValue Then
                        s = s + 1
                        cx = .STDtest(Index).Readings(a).Value
                        cDeviazione = cDeviazione + (cMedia - cx) ^ 2
                        cDevSt(s, Index) = Sqr(cDeviazione / s)
                        If cDevSt(s, Index) > MaxDev Then MaxDev = cDevSt(s, Index)
                    End If
                 End If
            End If
        Next
        If nOutOfRange = 0 Then
         .DataControllo(Index).OutOfRangeData = 0
         .DataControllo(Index).OutOfRangeDataPerc = 0
        Else
         .DataControllo(Index).OutOfRangeData = nOutOfRange
         .DataControllo(Index).OutOfRangeDataPerc = FormatNumber(100 * nOutOfRange / t, 1)
        End If
        

        
        cDeviazione = Sqr(cDeviazione / t)
        
        
        If cDeviazione = 0 Then cDeviazione = 0.001  ' per poter disegnare devo avere dev>0
        .DataControllo(Index).devst = cDeviazione
        
        
        cTotali = .NumReadings
           
        cUm = MeasurementUnit
        cNominale = .DataControllo(Index).STDRef
        cQMR = .DataControllo(Index).s2 * 2

        
    '-----------------------------------------------------------------------------
    '
    '            stampa a schermo i dati
    '
    '------------------------------------------------------------------------------
        lbl_ris(0) = t
        lbl_ris(1) = STD(absIndex, 1) ' & cUm
        lbl_ris(2) = .DataControllo(Index).OutOfRangeData
        If bAllValue Then
              lbl_ris(3) = MyChemicalQC.STDtest(Index).NumTest
           Else
             lbl_ris(3) = MyChemicalQC.STDtest(Index).SelTest
        End If
        lbl_ris(4) = Format$(cMedia, UserDecimal)
        lbl_ris(5) = .DataControllo(Index).OutOfRangeDataPerc & "%"
        lbl_ris(6) = STD(absIndex, 2)
        lbl_ris(7) = STD(absIndex, 3)
        
        ' se prendo i dati senza EVALUATION allora le medie non ci saranno
        If MyChemicalQC.STDtest(Index).TotalMean = 0 Then
            MyChemicalQC.STDtest(Index).TotalMean = lbl_ris(4)
            Graph1.MeanValue = MyChemicalQC.STDtest(INDEX_STD).TotalMean - STD(absIndex, 1)
        End If
        If MyChemicalQC.STDtest(Index).SelecMean = 0 Then
            MyChemicalQC.STDtest(Index).SelecMean = lbl_ris(4)
            Graph1.MeanValue = MyChemicalQC.STDtest(INDEX_STD).SelecMean - STD(absIndex, 1)
        End If
        'MyChemicalQC.STDtest(Index).NumTest = lbl_ris(3)

           
    '-----------------------------------------------------------------------------
    '
    '            assegnazione variabili grafico > MyGraphicCheck
    '
    '------------------------------------------------------------------------------
           
        MyGraphicCheck = MyChemicalQC
        
    End With

    Ydev = IIf(MaxDev = 0, 1, MaxDev)
    
    cYMax = (1 / (cDeviazione * Sqr(2 * 3.14)))
    Graph1.YMaxDevST = cYMax
    'cYMax = IIf(cYMax = 0, cTotali, cYMax)
    MaxYGauss(INDEX_STD) = cYMax * 1.2

ERR_END:
    On Error GoTo 0
    
    Exit Sub
ERR_DATI:
    MsgBox err.Description
    Resume Next

End Sub




Private Sub Form_Resize()
Dim Index As Integer

On Local Error Resume Next



ResizeControls


End Sub



Private Sub Form_Unload(Cancel As Integer)
bActivate = False
Set F_GRAPH = Nothing
End Sub

Private Sub lbl_VALUE_Click()
lbl_VALUE.Visible = False
End Sub

Private Sub mnuesci_Click()
Unload Me
End Sub

Private Sub mnustampa_Click()
Dim str_grafico As String
str_grafico = Trim(blTable)
Graph1.PrintPicture = True
End Sub



Private Sub Picture1_DblClick()
'If Me.WindowState = 2 Then
'    Me.WindowState = 0
'    Else
'    Me.WindowState = 2
'End If
'
End Sub

Private Sub Picture1_LostFocus()
lbl_VALUE.Visible = False
End Sub



Public Sub Gaussiana(ByVal Index As Integer)  ', XData As Double, YData As Double, pltColor As Long)
Dim j As Integer
Dim PointValue As Double
Dim RealValue As Double
Dim MaxXValue As Double
Dim Partenza As Double


    With MyChemicalQC
    
    
        .DataControllo(Index).media = FormatNumber(.DataControllo(Index).media, virgola)
        SetStyle 2, 0, MaxYGauss(Index) '* 1.5
        
        Graph1.Redraw = False
        Graph1.s1 = CDbl(STD(absIndex, 2)) ' yData
        Graph1.s2 = CDbl(STD(absIndex, 3)) 'yData
        Graph1.devst = FormatNumber(.DataControllo(Index).devst, virgola + 2)
       ' Graph1.MeanValue =
        
        If MaxValue(Index) * 2 > .DataControllo(Index).media * 2 Then
            X0Value(Index) = MaxValue(Index)
        Else
            X0Value(Index) = .DataControllo(Index).media
        End If
        
        Graph1.X0Value = X0Value(Index)
        If .DataControllo(Index).devst = 0 Then Exit Sub
        
        'Partenza = .DataControllo(Index).media - .DataControllo(Index).s
        
        For j = -2000 To 2000
           'xData = STD(absIndex, 3) * 2 * ((j / 1000))
           'xData = STD(absIndex, 3) * ((j / 1000) + 1)
           
           
            If STD(absIndex, 1) < 1 Or (STD(absIndex, 3)) = 0 Then
                'xData = X0Value(Index) * ((j / 2000) + 1)
                xData = ((.DataControllo(Index).media) * ((j / 1000))) + .DataControllo(Index).media
            Else
            
                'Partenza = (((STD(absIndex, 3)) * 10 / .DataControllo(Index).media * ((j / 2000))) / 2) * .DataControllo(Index).devst * 2
                
                Partenza = .DataControllo(Index).s2 * 2 * (j / 1000)
                
                xData = Partenza + .DataControllo(Index).media
                
                
            End If
            
      
            yData = (1 / (.DataControllo(Index).devst * Sqr(2 * 3.14)) * Exp(-(xData - .DataControllo(Index).media) ^ 2 / (2 * .DataControllo(Index).devst ^ 2)))
            
            
            If xData < 0 Then yData = 0
            
            If xData = 0 Then
                Debug.Print j
            End If
            If FormatNumber(cYMax, 7) = FormatNumber(yData, 7) Then
                Debug.Print j
            End If
            PointValue = yData ' CDbl(.STDTest(Index).Readings(j).Value) - CDbl(.DataControllo(Index).STDRef)
            RealValue = xData ' yData 'CDbl(.STDTest(Index).Readings(j).Value)
            AddPoint PointValue, 1, RealValue

        Next j
        Graph1.Redraw = True
    End With
    Exit Sub
err:


End Sub




Private Function GetSTDValueForGraph() As Boolean
Dim t As Integer
Dim k As Integer
Dim rc As Boolean
Dim Tolerance As Double
        STDCount = GetSettingData(SettingName, "Graph QC", "STDCount", 0)
        virgola = GetSettingData(SettingName, "Code Information", "Decimal", 0)
        UserDecimal = FormatDecimal(GetSettingData(SettingName, "Code Information", "Decimal", 0))
        MeasurementUnit = GetSettingData(SettingName, "Code Information", "MeasurementUnit", "")
        
    If STDCount = 0 Then
        
        rc = CheckReading
        If Not (rc) Then
            GetSTDValueForGraph = False
        End If
        Exit Function
    
    End If
        
    With cmbSTD
      '  .Clear
        ReDim STD(STDCount, 3) As String
        For t = 1 To STDCount
            STD(t, 0) = GetSettingData(SettingName, "Graph QC", "STDNumber" & t, "")
            STD(t, 1) = GetSettingData(SettingName, "Graph QC", "STDValue" & t, "")
            STD(t, 2) = GetSettingData(SettingName, "Graph QC", "STDMin" & t, "")
            STD(t, 3) = GetSettingData(SettingName, "Graph QC", "STDMax" & t, "")
            
            If InStr(STD(t, 1), "/") Or STD(t, 1) = "" Then
                STD(t, 1) = 0
            End If
            If InStr(STD(t, 2), "/") Or STD(t, 2) = "" Then
                STD(t, 2) = 0
            End If
            If InStr(STD(t, 3), "/") Or STD(t, 3) = "" Then
                STD(t, 3) = 0
            End If
            CloseSettingDataFile
            
            Tolerance = CDbl(STD(t, 3)) - CDbl(STD(t, 1))
            k = STD(t, 0)
            
            With MyChemicalQC.DataControllo(k)
                .s = CDbl(STD(t, 2)) - CDbl(STD(t, 1))
                .s2 = CDbl(STD(t, 3)) - CDbl(STD(t, 1))
                .STDRef = CDbl(STD(t, 1))
                .STDMin = STD(t, 2)
                .STDMax = STD(t, 3)
                .STDNumber = STD(t, 0)
            End With
            
            If STD(t, 0) <> "" Then
                .AddItem "    " & STD(t, 0)
                
            End If
        Next
        If t > 0 Then
            .ListIndex = 0
        End If
          
    End With
    
    GetSTDValueForGraph = True

End Function

Private Function GetReadingsFormFile()
Dim rc As Boolean
Dim t       As Integer
Dim i       As Integer
Dim j       As Integer
Dim r       As Integer
Dim MaxTest As Integer
Dim NumReading As Integer
Dim RedingDiff As Double
Dim SelTestCount As Integer
Dim MyCheckString As String
Dim k As Integer
On Error GoTo ERR_READ

    MeterNumber = GetSettingData(SettingName, "Information QC", "MeterNumber", 0)
    
    
  rc = IIf(Check3.Value = 1, True, False)
ViewMeterLegend rc


    If MeterNumber = 0 Then
        PopupMessage 2, "Warning : Please Fill Information QC and Reading QC ...", , True
        CloseSettingDataFile
        Unload Me
        Exit Function
    End If

    For k = 1 To STDCount
        NumReading = 0
        t = CInt(STD(k, 0))
        MaxTest = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Tests", 0)
        
        If MaxTest = 0 Then
        
                ' non ci sono test oppure non ho ancora fatto evaluation
                Call CheckReading
                Exit Function
        End If
        
        MyChemicalQC.STDtest(t).MaxReadings = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Readings", 0)
        MyChemicalQC.STDtest(t).SelReadings = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Selected Readings", 0)
        MyChemicalQC.STDtest(t).NumTest = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Tests", 0)
        MyChemicalQC.STDtest(t).SelTest = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Selected Tests", 0)
        MyChemicalQC.STDtest(t).TotalMean = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Total Average", 0)
        MyChemicalQC.STDtest(t).SelecMean = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Selected Average", 0)
        
        SelTestCount = 0
        For i = 1 To MaxTest
            
            For j = 1 To MeterNumber
                With MyChemicalQC.STDtest(t)
                    NumReading = NumReading + 1
                    .Readings(NumReading).Value = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Test " & i & " Meter " & j & " Value", "")
                    .Readings(NumReading).Meter = j
                    .Readings(NumReading).bSelectedValue = GetSettingData(SettingName, "Graph QC", "Standard " & t & " Test " & i & " Meter " & j & " Selected", "TRUE")
                    If .Readings(NumReading).bSelectedValue = True Then
                        SelTestCount = SelTestCount + 1
                    End If
                End With

            Next

        Next
            If MyChemicalQC.STDtest(t).SelTest = 0 Then
                 MyChemicalQC.STDtest(t).SelTest = Int(SelTestCount / MeterNumber)
            End If
        'With MyChemicalQC.STDtest(t)
        'For r = 1 To NumReading
        '    If .Readings(r).Value = "" Then
        '    Else
        '        If .Readings(r).bSelectedValue Then
        '            RedingDiff = Abs(CDbl(.Readings(r).Value) - CDbl(STD(absIndex, 1)))
        '            If RedingDiff > MaxY(t) Then
        '               MaxY(t) = Abs(RedingDiff) * 1.5
        '            End If
        '        End If
        '    End If
        'Next
        'If MaxY(t) < MyChemicalQC.DataControllo(t).s2 Then
        '    MaxY(t) = MyChemicalQC.DataControllo(t).s2 * 1.5
        'End If
        'End With
    Next
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_READ:
    MsgBox err.Description
    Resume Next


End Function



Private Function GetMaxY(ByVal t As Integer)

Dim r               As Integer
Dim NumReadings     As Integer
Dim RedingDiff      As Double

With MyChemicalQC.STDtest(t)
    NumReadings = .MaxReadings
    MaxY(t) = 0
    For r = 1 To NumReadings
        If .Readings(r).Value = "" Then
        Else
            If .Readings(r).bSelectedValue Or bAllValue Then
                RedingDiff = Abs(CDbl(.Readings(r).Value) - CDbl(STD(absIndex, 1)))
                
                If RedingDiff > MaxY(t) Then
                   MaxY(t) = Abs(RedingDiff) * 1.5
                End If
            End If
        End If
    Next
    If MaxY(t) < MyChemicalQC.DataControllo(t).s2 * 1.8 Then
        MaxY(t) = MyChemicalQC.DataControllo(t).s2 * 1.8
    End If
    End With
    

    

End Function
Private Sub cmbSTD_Click()
Dim Index As Integer
    Index = cmbSTD
    INDEX_STD = Index
    absIndex = cmbSTD.ListIndex + 1
    GetMaxY (INDEX_STD)
    Graph1.Redraw = False
    Call Dati_Lotto(INDEX_STD)
  
    'Label2 = "STD" & INDEX_STD
    Graph1.Redraw = True
End Sub




Private Sub GetCodeInformation(ByVal sLot As String, ByVal sCode As String, Optional ByVal MyID As Long)
Dim MeasurementUnit As String
    ' attenzione , se ho un file allora lo importo , altrimenti prendo i dati del Code
    
    MyID = GetSettingData(SettingName, "Code Information", "ID", 0)
    
    With dbTabCode
        .filter = ""
        If MyID > 0 Then
            .filter = "ID='" & MyID & "'"
        Else
        
            .filter = "Code='" & sCode & "'"
        End If
    

        If .EOF Then
            MessageInfoTime = 2000
            PopupMessage 2, "Cannot find Hanna Code  :  " & sCode & vbCrLf & "Please Enter Code Information..."
            'Unload Me
        Else
        
            MeasurementUnit = IIf(IsNull(Trim(!MeasurementUnit)), "", " " & Trim(!MeasurementUnit))
            Text1(0) = sLot
            Text1(1) = sCode
            Text1(2) = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
          
            
            
            GetFormSettingName


    
        End If
    
    End With




End Sub
Private Function GetFormSettingName()
Dim i As Integer
Dim NumReadings As String
   If FileExists(USER_TEMP_PATH & SettingName) Then
   ElseIf FileExists(USER_DATA_PATH & SettingName) Then
       ' PopupMessage 2, "Lot : " & Text1(0) & vbCrLf & "Code : " & Text1(1) & vbCrLf & "Is Closed..."
        USER_PATH = USER_DATA_PATH
   Else
        Exit Function
   End If
   
    CloseSettingDataFile
    
    'For i = 0 To Text1.Count - 1
    '   Text1(i) = GetSettingData(SettingName, "Information QC", "Text1" & i, Text1(i))
    'Next


    '---------------------------------------------------------
    ' numero di letture
    
   
    With MyChemicalQC
        .Lot = GetSettingData(SettingName, "Information QC", "Text10", "")
        .HannaCode.code = GetSettingData(SettingName, "Information QC", "Text11", "")
        .HannaCode.Description = GetSettingData(SettingName, "Information QC", "Text12", "")
        .Exp = GetSettingData(SettingName, "Information QC", "Text13", "")
        .HannaCode.Line = GetSettingData(SettingName, "Information QC", "Text14", "")
        .HannaCode.Recipe = GetSettingData(SettingName, "Information QC", "Text15", "")
        .date = GetSettingData(SettingName, "Information QC", "Modification Date", "")
        .HannaCode.Decimal = FormatDecimal(GetSettingData(SettingName, "Code Information", "Decimal", 0))
        .HannaCode.MeasurementUnit = GetSettingData(SettingName, "Code Information", "MeasurementUnit", "")
        .ProdFirst = GetSettingData(SettingName, "Information QC", "Text123", "")
        .ProdLast = GetSettingData(SettingName, "Information QC", "Text124", "")
    End With

    '---------------------------------------------------------

    GetReadingsFormFile

    CloseSettingDataFile
End Function



Private Sub AddPoint(ByVal Value As Double, ByVal i As Integer, ByVal RealValue As Double)
Dim lngIndex    As Long
Dim lngValue    As Double
Dim blnRedraw   As Boolean

    
    With Graph1
        blnRedraw = .Redraw
        '.Redraw = False
        Select Case mlngStyle
            Case STYLE_MUSIC
                '.Points.Clear
                'For lngIndex = 1 To IIf(Graph1.FixedPoints = 0, Graph1.Points.Count, Graph1.FixedPoints)
                '    lngValue = (Rnd * 80) + 50
                '    .Points.Add lngValue, i
                '    i = i + 1
                'Next lngIndex
            Case Else
                lngValue = Value
                .Points.Add lngValue, i, RealValue
                .BarColor = &H964901   'vbRed
        End Select
        '.Redraw = blnRedraw
    End With

End Sub

Private Sub SetStyle(ByVal plngIndex As Long, ByVal plngMin As Double, ByVal plngMax As Double)
Dim lngX    As Long
    Picture1.Visible = False
    With Graph1
        mlngStyle = plngIndex
        .Redraw = False
        .Points.Clear
        .MaxValue = plngMax
        .MinValue = plngMin
        .MeanValue = 0
        .STDValue = STD(absIndex, 1)
        If bAllValue Then
             .MeanValue = MyChemicalQC.STDtest(INDEX_STD).TotalMean - STD(absIndex, 1)
        Else
             .MeanValue = MyChemicalQC.STDtest(INDEX_STD).SelecMean - STD(absIndex, 1)
        End If
       ' .X0Value = X0Value(INDEX_STD)
        .IntVirgola = virgola
        .XGridInc = 1
        .YGridInc = 0.02
        .ShowBars = IIf(Check2.Value = 1, True, False)
        .ShowBarsColor = IIf(Check3.Value = 1, True, False)
        .ShowValue = IIf(Check4.Value = 1, True, False)
        


        
        If bPrintMe Then
        
            .BackColor = vbWhite
            .AxisColor = &H964901
            .GridColor = RGB(110, 135, 190)
            .LineColor = RGB(255, 255, 255) 'vbBlack '
            .PointColor = RGB(255, 0, 0)
            .BarColor = &HECD2BF  ' &HC0FFFF
            .MeanColor = vbBlack
            .NumberColor = &HC0C0C0
        Else
            .NumberColor = vbBlack ' vbWhite '
            .BackColor = RGB(121, 145, 200)  '&H101010    '
            .AxisColor = &H964901
            .GridColor = RGB(110, 135, 190)
            .LineColor = RGB(255, 255, 255)
            .PointColor = RGB(255, 0, 0)
            .MeanColor = &HC0C0C0
            .BarColor = &HECD2BF  ' &HC0FFFF
        
        End If
        
                
        Select Case plngIndex
            Case STYLE_MUSIC

            Case STYLE_TASKMAN
                
            Case STYLE_GRAPH ' GAUSSIANA
                .FadeIn = False
                .ShowGaussian = True
                .FixedPoints = 0
                .BarWidth = 0.8
                .XGridInc = 90
                .AxisColor = 0
              
                .MeanValue = FormatNumber(cMedia, virgola)
                .ShowLines = True
                .ShowPoints = False
                .ShowBars = False
                ' ---------------------------------------
                
            Case STYLE_CUSTOM  ' MEDIE
                .BarWidth = 0.8
                .FixedPoints = 0
                .ShowGaussian = False

                .ShowAxis = True
                .ShowGrid = True
                .ShowPoints = True
                .ShowLines = True
                .FadeIn = True


                
                ' --------------------------------------

        End Select
        .Redraw = True
    End With
End Sub



Private Function CheckReading() As Boolean
Dim rc As Boolean
Dim t As Integer
Dim Tolerance As Double

    CheckReading = True
    
    STDCount = GetSettingData(SettingName, "Graph QC", "STDCount", 0)
    virgola = GetSettingData(SettingName, "Code Information", "Decimal", 0)
    
    rc = GetSTDTable
    If Not (rc) Then
        CheckReading = False
        Exit Function
    End If
    
    bAllValue = True
    Check1.Visible = False
    
    If STDCount = 0 Then
        
        Exit Function
    
    End If
    
    GetReadingsFormFile
    
    With cmbSTD
      '  .Clear
        For t = 1 To STDCount
         
            If InStr(STD(t, 1), "/") Or STD(t, 1) = "" Then
                STD(t, 1) = 0
            End If
            If InStr(STD(t, 2), "/") Or STD(t, 2) = "" Then
                STD(t, 2) = 0
            End If
            If InStr(STD(t, 3), "/") Or STD(t, 3) = "" Then
                STD(t, 3) = 0
            End If
            CloseSettingDataFile
            
            Tolerance = CDbl(STD(t, 3)) - CDbl(STD(t, 1))
            With MyChemicalQC.DataControllo(t)
                .s = CDbl(STD(t, 2)) - CDbl(STD(t, 1))
                .s2 = CDbl(STD(t, 3)) - CDbl(STD(t, 1))
                .STDRef = CDbl(STD(t, 1))
            End With
            If STD(t, 0) <> "" Then
                .AddItem "    " & STD(t, 0)
                
            End If
        Next
        If t > 0 Then
            .ListIndex = 0
        End If
          
    End With
    
    

End Function



Private Function GetSTDTable() As Boolean
    Dim rc As Boolean
    Dim MyRows As String
    Dim MyCols As String
    Dim i As Integer
    Dim t As Integer
   
    Dim STDNumber As String
    Dim STDValue As String
    Dim STDMin As String
    Dim STDMax As String
    
    On Error GoTo ERR_GET
    
    rc = False
    
    MyRows = GetSettingData(SettingName, "Reading QC", "Grd2 Rows", 1)
    MyCols = GetSettingData(SettingName, "Reading QC", "Grd2 Cols", 1)
    
    If MyRows < 2 Then
        ' non ho fatto neanche un reading....
        PopupMessage 2, "Warning : this Lot has NO Readings..." & vbCrLf & "Please enter at least 1 reading", , True
        CloseSettingDataFile
        GetSTDTable = False
        Exit Function
    End If

    t = 0
    ReDim STD(10, 4) As String
    If MyRows > 1 Then
        If MyCols > 1 Then

            For i = 1 To MyRows - 1
                
                    STDNumber = GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "1", "")
                    STDValue = GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "2", "")
                    STDMin = GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "27", "")
                    STDMax = GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "28", "")
                    
                    
                    
                    If STDNumber <> "" And STDMin <> "" And STDMin <> "" Then
                    
                        If STDNumber = "/" Then STDNumber = "1"
                        If STDMin = "/" Then STDMin = "0"
                        If STDMax = "/" Then STDMax = "0"
                        
                        If t > 0 Then
                            If GetIndexArStr(STD(), STDNumber) = -1 Then
Aggiungi:
                                
                                t = t + 1
                    
                                STD(t, 0) = STDNumber
                                STD(t, 1) = STDValue
                                STD(t, 2) = STDMin
                                STD(t, 3) = STDMax
                                rc = True
                            End If
                        Else
                            GoTo Aggiungi
                        End If
                    End If
            Next
        End If
    End If
ERR_END:
    On Error GoTo 0
    CloseSettingDataFile
    If rc Then
        STDCount = t
        If t > 0 Then
            Call SaveValueForGraph
            Call GetSelectedTest
        Else
            ' non ho fatto neanche un reading....
            PopupMessage 2, "Warning : this Lot has NO Standard..." & vbCrLf & "Please enter at least 1 valid STD", , True
            CloseSettingDataFile
            Unload Me
            Exit Function
        End If
    End If
    Exit Function
ERR_GET:
    rc = False
    MsgBox err.Description
    Resume Next
End Function
Private Function SaveValueForGraph()
 Dim t As Integer
        SaveSettingData SettingName, "Graph QC", "STDCount", STDCount
        For t = 1 To STDCount
            SaveSettingData SettingName, "Graph QC", "STDNumber" & t, STD(t, 0)
            SaveSettingData SettingName, "Graph QC", "STDValue" & t, STD(t, 1)
            SaveSettingData SettingName, "Graph QC", "STDMin" & t, STD(t, 2)
            SaveSettingData SettingName, "Graph QC", "STDMax" & t, STD(t, 3)
        Next
        
End Function


Private Function GetSelectedTest()
Dim i As Integer
Dim t As Integer
Dim rc As Boolean
Dim lRows As Long
Dim lCols As Long
Dim ReadingStandard(99) As String
Dim numSelectedStandard As Integer
Dim STDtest(99, 6) As String

    CloseSettingDataFile
    
    On Error GoTo ERR_GET:
    
    MeterNumber = GetSettingData(SettingName, "Information QC", "MeterNumber", 0)
    If MeterNumber = 0 Then
        PopupMessage 2, "Warning : Please enter at least 1 Meter in Information QC", , True
        CloseSettingDataFile
        Unload Me
        Exit Function
    End If
    'For numSelectedStandard = 1 To STDCount
    

           lRows = GetSettingData(SettingName, "Reading QC", "Grd2 Rows", 0)
           lCols = GetSettingData(SettingName, "Reading QC", "Grd2 Cols", 0)
            For i = 1 To lRows - 1
                
                ReadingStandard(i) = GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "1", "")
                
                For t = 11 To 10 + MeterNumber
                    STDtest(i, t - 11) = GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & t, "")
                    
                Next
            Next

    'Next
        
    CloseSettingDataFile
        
    ' salvo per grafico.....
    'For numSelectedStandard = 1 To STDCount
        For t = 1 To lRows - 1
            For i = 1 To MeterNumber
                SaveSettingData SettingName, "Graph QC", "Standard " & ReadingStandard(t) & " Test " & t & " Meter " & i & " Value", STDtest(t, i - 1)
               ' SaveSettingData SettingName, "Graph QC", "Standard " & ReadingStandard(t) & " Test " & t & " Meter " & i & " Selected", "True"
            Next
        Next
   ' For t = 1 To STDCount
       ' MyChemicalQC.STDtest(t).MaxReadings = (lRows - 1) * MeterNumber
       ' MyChemicalQC.STDtest(t).SelReadings = (lRows - 1) * MeterNumber
       ' MyChemicalQC.STDtest(t).NumTest = (lRows - 1)
       ' MyChemicalQC.STDtest(t).SelTest = (lRows - 1)
                
       ' SaveSettingData SettingName, "Graph QC", "Standard " & t & " Total Readings", MyChemicalQC.STDtest(t).MaxReadings
        'SaveSettingData SettingName, "Graph QC", "Standard " & t & " Selected Readings", MyChemicalQC.STDtest(t).SelReadings
        'SaveSettingData SettingName, "Graph QC", "Standard " & t & " Total Tests", MyChemicalQC.STDtest(t).NumTest
        'SaveSettingData SettingName, "Graph QC", "Standard " & t & " Selected Tests", MyChemicalQC.STDtest(t).SelTest
       ' SaveSettingData SettingName, "Graph QC", "Standard " & t & " Total Average", MyChemicalQC.STDtest(t).TotalMean
       'SaveSettingData SettingName, "Graph QC", "Standard " & t & " Selected Average", MyChemicalQC.STDtest(t).SelecMean
   ' Next
ERR_END:
    
    CloseSettingDataFile
    Exit Function
ERR_GET:
    MsgBox err.Description
    Resume Next
End Function

Private Sub PicZOOM_Click(Index As Integer)

    Select Case IndexFormProcedura
    
        Case 0
        
            With Graph1
                Select Case Index
                    Case 0
                        ' zoom +
                        MaxY(INDEX_STD) = MaxY(INDEX_STD) * 0.7
                    Case 1
                        ' zoom -
                        MaxY(INDEX_STD) = MaxY(INDEX_STD) * 1.3
                            
                End Select
            End With
            
            SetStyle 3, -MaxY(INDEX_STD), MaxY(INDEX_STD)
            Dati_Lotto (INDEX_STD)
            
        Case 1
        
            With Graph1
                Select Case Index
                    Case 0
                        ' zoom +
                        MaxYGauss(INDEX_STD) = MaxYGauss(INDEX_STD) * 0.7
                    Case 1
                        ' zoom -
                        MaxYGauss(INDEX_STD) = MaxYGauss(INDEX_STD) * 1.3
                End Select
            End With
            
            Gaussiana (INDEX_STD)
        
        
    End Select

End Sub

Private Sub Timer1_Timer()
Label5.Caption = "File correctly saved..."
Timer1.Enabled = False
End Sub

Private Function PrintGraph()
Dim rc As Boolean

On Error GoTo ERR_PRINT:
    rc = True
    
    With MyGraphicCheck
        .DataControllo(INDEX_STD).Operator = MyOperatore.Name
    End With
        Label5.Caption = "Printing pdf file..."
        Frame1.Visible = True
        Timer1.Enabled = True
        MyHeight = Graph1.Height
        MyWidth = Graph1.Width
        
        With Graph1
            .BackColor = vbWhite
            .AxisColor = &H964901
            .GridColor = RGB(110, 135, 190)
            .LineColor = vbBlack 'RGB(255, 255, 255)
            .PointColor = RGB(255, 0, 0)
            .BarColor = &HECD2BF  ' &HC0FFFF
            .MeanColor = vbBlack
            .NumberColor = &HC0C0C0
            .Height = 4755
            .Width = 16310
        End With
        DoEvents
        Call SetPDFPrinter("", USER_DOCUMENTI & "\" & PathReport & "\" & PdfFileName)
        
        
        ' ok vai in STAMPA
        
        
        mnustampa_Click
        DoEvents
        
        
        ' uscito da STAMPA
        
        
        With Graph1
        
            .NumberColor = &HC0C0C0
            .BackColor = RGB(121, 145, 200)
            .AxisColor = &H964901
            .GridColor = RGB(110, 135, 190)
            .LineColor = RGB(255, 255, 255)
            .PointColor = RGB(255, 0, 0)
            .MeanColor = &HC0C0C0
            .BarColor = &HECD2BF  ' &HC0FFFF
            .Height = MyHeight
            .Width = MyWidth
        End With
ERR_END:
    
    CloseSettingDataFile
    Exit Function
ERR_PRINT:
    MsgBox err.Description
    Resume Next

End Function
