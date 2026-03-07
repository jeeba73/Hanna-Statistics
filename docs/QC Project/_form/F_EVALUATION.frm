VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form F_EVALUATION 
   BackColor       =   &H00808080&
   Caption         =   "Evaluation"
   ClientHeight    =   12000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
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
   Icon            =   "F_EVALUATION.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12510
   ScaleMode       =   0  'User
   ScaleWidth      =   19200
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   0
      Left            =   0
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Frame Frame2 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2535
         Left            =   3480
         TabIndex        =   23
         Top             =   2520
         Visible         =   0   'False
         Width           =   12255
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00004000&
            BorderStyle     =   0  'None
            FillColor       =   &H00004000&
            ForeColor       =   &H8000000D&
            Height          =   615
            Left            =   8040
            ScaleHeight     =   615
            ScaleWidth      =   3975
            TabIndex        =   91
            Top             =   1680
            Visible         =   0   'False
            Width           =   3975
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackColor       =   &H00004000&
               BackStyle       =   0  'Transparent
               Caption         =   "Update STD from Database"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   0
               TabIndex        =   92
               Top             =   120
               Width           =   3975
            End
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H000080DF&
            BorderColor     =   &H000060BF&
            Height          =   2535
            Left            =   0
            Top             =   0
            Width           =   12255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "6 - Laboratory Manager Only can Close Lots"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   1560
            TabIndex        =   58
            Top             =   2040
            Width           =   4260
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "4 - Save STD Mean Value"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   1560
            TabIndex        =   57
            Top             =   1320
            Width           =   2430
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "5 - Check Results"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   1560
            TabIndex        =   56
            Top             =   1680
            Width           =   1680
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3 - Check ph Range : Select ph to view pH number and Range"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   26
            Top             =   960
            Width           =   5925
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2 - Select Value from Readings Table : Select/Deselect CheckBoxes ( at least 80% of all value )"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1560
            TabIndex        =   25
            Top             =   600
            Width           =   9195
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1 - Select Standard from SFG Standard Table to show all STD Readings"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   24
            Top             =   240
            Width           =   6780
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   520
            Picture         =   "F_EVALUATION.frx":33E2
            Top             =   960
            Width           =   480
         End
      End
      Begin FlexCell.Grid Grid1 
         Height          =   5655
         Left            =   14760
         TabIndex        =   95
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   9975
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColorActiveCellSel=   15790320
         BackColorBkg    =   15790320
         BackColorFixed  =   15790320
         BackColorFixedSel=   15790320
         BackColorScrollBar=   15592423
         BorderColor     =   15790320
         CellBorderColor =   15790320
         CellBorderColorFixed=   15790320
         Cols            =   5
         DefaultFontName =   "Century Gothic"
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DefaultRowHeight=   20
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   16777215
         Rows            =   1
         ScrollBarStyle  =   0
         MultiSelect     =   0   'False
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         FillColor       =   &H00004000&
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   6600
         ScaleHeight     =   615
         ScaleWidth      =   4095
         TabIndex        =   89
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00004000&
            BackStyle       =   0  'Transparent
            Caption         =   "Update STD from Database"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   90
            Top             =   120
            Width           =   4095
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   13
         Left            =   16800
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   3720
         Width           =   1935
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         FillColor       =   &H00004000&
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   14760
         ScaleHeight     =   855
         ScaleWidth      =   3975
         TabIndex        =   61
         Top             =   1200
         Width           =   3975
         Begin VB.Label lbSTDNumber 
            Alignment       =   2  'Center
            BackColor       =   &H00004000&
            BackStyle       =   0  'Transparent
            Caption         =   "STD Number"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   62
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   6
         Left            =   16800
         TabIndex        =   37
         Top             =   5760
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   5
         Left            =   14760
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   5760
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   4
         Left            =   14760
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   4800
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   3
         Left            =   14760
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3720
         Width           =   1935
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         FillColor       =   &H00004000&
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   14760
         ScaleHeight     =   855
         ScaleWidth      =   4095
         TabIndex        =   28
         Top             =   6720
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Image ImageTAV 
            Height          =   480
            Index           =   5
            Left            =   1755
            MouseIcon       =   "F_EVALUATION.frx":67C4
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":6ACE
            Top             =   170
            Width           =   480
         End
      End
      Begin FlexCell.Grid Grd2 
         Height          =   6615
         Left            =   4560
         TabIndex        =   96
         Top             =   960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   11668
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColorActiveCellSel=   15790320
         BackColorBkg    =   15790320
         BackColorFixed  =   15790320
         BackColorFixedSel=   15790320
         BackColorScrollBar=   15592423
         BorderColor     =   15790320
         CellBorderColor =   15790320
         CellBorderColorFixed=   15790320
         Cols            =   5
         DefaultFontName =   "Century Gothic"
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DefaultRowHeight=   20
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   15790320
         Rows            =   1
         ScrollBarStyle  =   0
         MultiSelect     =   0   'False
      End
      Begin FlexCell.Grid Grd1 
         Height          =   6615
         Left            =   240
         TabIndex        =   97
         Top             =   960
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   11668
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColorActiveCellSel=   15790320
         BackColorBkg    =   15790320
         BackColorFixed  =   15790320
         BackColorFixedSel=   15790320
         BackColorScrollBar=   15592423
         BorderColor     =   15790320
         CellBorderColor =   15790320
         CellBorderColorFixed=   15790320
         Cols            =   5
         DefaultFontName =   "Century Gothic"
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DefaultRowHeight=   20
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   15790320
         Rows            =   1
         ScrollBarStyle  =   0
         MultiSelect     =   0   'False
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   2
         Left            =   14760
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2280
         Width           =   3975
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   5
         Left            =   14760
         MousePointer    =   99  'Custom
         Picture         =   "F_EVALUATION.frx":9EB0
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QC Results Table"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   15480
         MouseIcon       =   "F_EVALUATION.frx":D292
         MousePointer    =   99  'Custom
         TabIndex        =   99
         Top             =   480
         Width           =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Mean Value ( Selected )"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   4
         Left            =   14760
         TabIndex        =   34
         Top             =   4440
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "# Tests"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   13
         Left            =   16800
         TabIndex        =   64
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label lb80Pec 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select at least 70% of alla value"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   6405
         TabIndex        =   55
         Top             =   480
         Width           =   8010
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "# Tests ( S )"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   6
         Left            =   16800
         TabIndex        =   38
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "# Reading ( S )"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   5
         Left            =   14760
         TabIndex        =   36
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "# Readings"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   3
         Left            =   14760
         TabIndex        =   32
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Total Average"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   2
         Left            =   14760
         TabIndex        =   30
         Top             =   2280
         Width           =   3975
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   1
         Left            =   4560
         Picture         =   "F_EVALUATION.frx":D59C
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Readings"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   5160
         MouseIcon       =   "F_EVALUATION.frx":1097E
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SFG Standard"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   840
         MouseIcon       =   "F_EVALUATION.frx":10C88
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   480
         Width           =   1380
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   3
         Left            =   240
         Picture         =   "F_EVALUATION.frx":10F92
         Top             =   360
         Width           =   480
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Left            =   9360
         Picture         =   "F_EVALUATION.frx":14374
         Top             =   8880
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operator"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   360
         TabIndex        =   7
         Top             =   9120
         Width           =   975
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   7440
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   0
      Left            =   0
      MouseIcon       =   "F_EVALUATION.frx":17756
      MousePointer    =   99  'Custom
      ScaleHeight     =   1815
      ScaleWidth      =   2775
      TabIndex        =   16
      Top             =   1080
      Width           =   2775
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recipe / Reagent"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   555
         MouseIcon       =   "F_EVALUATION.frx":17A60
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   0
         Left            =   1200
         Picture         =   "F_EVALUATION.frx":17D6A
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.PictureBox PicMenuBar 
      BackColor       =   &H00303030&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   19215
      TabIndex        =   9
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   4
         Left            =   3840
         MouseIcon       =   "F_EVALUATION.frx":1B14C
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   93
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   4
            Left            =   735
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":1B456
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Certificate"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   4
            Left            =   60
            MouseIcon       =   "F_EVALUATION.frx":1DE48
            MousePointer    =   99  'Custom
            TabIndex        =   94
            Top             =   720
            Width           =   1875
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   3
         Left            =   9600
         MouseIcon       =   "F_EVALUATION.frx":1E152
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Graph QC"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   3
            Left            =   0
            MouseIcon       =   "F_EVALUATION.frx":1E45C
            MousePointer    =   99  'Custom
            TabIndex        =   68
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   735
            MouseIcon       =   "F_EVALUATION.frx":1E766
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":1EA70
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   7680
         MouseIcon       =   "F_EVALUATION.frx":21E52
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   65
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   735
            MouseIcon       =   "F_EVALUATION.frx":2215C
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":22466
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reading QC"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   2
            Left            =   465
            MouseIcon       =   "F_EVALUATION.frx":25848
            MousePointer    =   99  'Custom
            TabIndex        =   66
            Top             =   720
            Width           =   960
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MouseIcon       =   "F_EVALUATION.frx":25B52
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   12
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "F_EVALUATION.frx":25E5C
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":26166
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "All Readings"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   60
            MouseIcon       =   "F_EVALUATION.frx":29548
            MousePointer    =   99  'Custom
            TabIndex        =   13
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
         MouseIcon       =   "F_EVALUATION.frx":29852
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   10
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Mean value"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   1
            Left            =   60
            MouseIcon       =   "F_EVALUATION.frx":29B5C
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   720
            Width           =   1875
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":29E66
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Evaluation QC"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   12150
         TabIndex        =   14
         Top             =   360
         Width           =   6540
      End
   End
   Begin ChemicalQC.ctlCalendar ctlCalendar1 
      Height          =   6960
      Left            =   7080
      TabIndex        =   73
      Top             =   3240
      Visible         =   0   'False
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   12277
      ShowLastMonthButton=   -1  'True
      ShowNextMonthButton=   -1  'True
      ShowLastMonthDays=   -1  'True
      ShowNextMonthDays=   -1  'True
      ShowTodayLabel  =   -1  'True
      ColorBackgroundHeader=   9849089
      ColorForegroundHeader=   16777215
      ColorSelectedBack=   9849089
      ColorSelectedFore=   16777215
      ColorToday      =   255
      ColorDayColumn  =   16777215
      ColorAlarms     =   0
      ColorBackground =   12632256
      ColorForeground =   4210752
      ColorButtons    =   -2147483633
      ColorLastNextMonthDayColor=   8421504
      ColorLine       =   0
      ColorWeekNumber =   8421504
      WeekStartsWith  =   1
      ShowSelected    =   -1  'True
      ShowToolTipText =   -1  'True
      ShowWeekNumbers =   0   'False
      ShowWeekNumberLeft=   -1  'True
      AllowRightClick =   0   'False
      UseAlarms       =   0   'False
      ShowShortDays   =   0   'False
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Whitney-Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDay {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontToday {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontColumn {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Whitney-Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00606060&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   2760
      TabIndex        =   15
      Top             =   1080
      Width           =   16455
      Begin VB.Frame Frame3 
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1455
         Left            =   6000
         TabIndex        =   42
         Top             =   240
         Visible         =   0   'False
         Width           =   8895
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   10
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "3.23"
            Top             =   840
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   7
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   45
            Text            =   "0.4"
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   8
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   44
            Text            =   "2.4"
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   9
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "3.09"
            Top             =   840
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "MAX"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   10
            Left            =   6000
            TabIndex        =   52
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "MIN"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   9
            Left            =   4080
            TabIndex        =   51
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "MAX"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   8
            Left            =   1920
            TabIndex        =   50
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "MIN"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   7
            Left            =   0
            TabIndex        =   49
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lb 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "STD Range"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   7
            Left            =   0
            TabIndex        =   48
            Top             =   80
            Width           =   3735
         End
         Begin VB.Label lb 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "pH Range"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   9
            Left            =   4080
            TabIndex        =   47
            Top             =   80
            Visible         =   0   'False
            Width           =   3735
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "hjkhkj"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "hhhh"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   1695
         Left            =   5880
         TabIndex        =   77
         Top             =   0
         Visible         =   0   'False
         Width           =   8895
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   12
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   1080
            Width           =   3495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   11
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Label lb 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "Closing Information"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   1
            Left            =   360
            TabIndex        =   82
            Top             =   320
            Width           =   7095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "by"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   12
            Left            =   3960
            TabIndex        =   81
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "Validation Date"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   11
            Left            =   360
            TabIndex        =   79
            Top             =   720
            Width           =   3495
         End
      End
      Begin VB.Label lb 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "Lot"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   0
         Left            =   1200
         TabIndex        =   41
         Top             =   320
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "Code SFG"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   1
         Left            =   3240
         TabIndex        =   20
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "Lot Number"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   0
         Left            =   1200
         TabIndex        =   18
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   1
      Left            =   0
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   21
      Top             =   2880
      Width           =   19215
      Begin VB.Frame frCommandInside 
         BackColor       =   &H00208040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   11520
         TabIndex        =   87
         Top             =   5640
         Width           =   3500
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Passed"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   3
            Left            =   0
            TabIndex        =   88
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.Frame frCommandInside 
         BackColor       =   &H00A88030&
         BorderStyle     =   0  'None
         Caption         =   "Image14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   7920
         TabIndex        =   85
         Top             =   5640
         Width           =   3500
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Waiting"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   86
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.Frame frCommandInside 
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         Caption         =   "Image14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   4320
         TabIndex        =   83
         Top             =   5640
         Width           =   3500
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Failed"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   84
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.PictureBox PicInformation 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   12600
         MouseIcon       =   "F_EVALUATION.frx":2D248
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   5295
         TabIndex        =   75
         Top             =   6600
         Visible         =   0   'False
         Width           =   5295
         Begin VB.Label Lab 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Goto Lot Information QC"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   1635
            MouseIcon       =   "F_EVALUATION.frx":2D552
            MousePointer    =   99  'Custom
            TabIndex        =   76
            Top             =   660
            Width           =   2010
         End
         Begin VB.Image Im 
            Height          =   480
            Left            =   2400
            MouseIcon       =   "F_EVALUATION.frx":2D85C
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":2DB66
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         FillColor       =   &H00004000&
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   6960
         ScaleHeight     =   855
         ScaleWidth      =   5295
         TabIndex        =   40
         Top             =   6600
         Visible         =   0   'False
         Width           =   5295
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "CLOSE LOT"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   54
            Top             =   240
            Width           =   5295
         End
         Begin VB.Image ImageTAV 
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "F_EVALUATION.frx":30F48
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":31252
            Top             =   180
            Width           =   480
         End
      End
      Begin FlexCell.Grid Grd3 
         Height          =   4095
         Left            =   2160
         TabIndex        =   98
         Top             =   1440
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   7223
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColorActiveCellSel=   15790320
         BackColorBkg    =   15790320
         BackColorFixed  =   15790320
         BackColorFixedSel=   15790320
         BackColorScrollBar=   15592423
         BorderColor     =   15790320
         CellBorderColor =   15790320
         CellBorderColorFixed=   15790320
         Cols            =   5
         DefaultFontName =   "Century Gothic"
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DefaultRowHeight=   20
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   15790320
         Rows            =   1
         ScrollBarStyle  =   0
         MultiSelect     =   0   'False
      End
      Begin VB.Image DelImage 
         Height          =   480
         Left            =   17160
         Picture         =   "F_EVALUATION.frx":34634
         Top             =   1440
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lbClose 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"F_EVALUATION.frx":37A16
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2580
         TabIndex        =   74
         Top             =   480
         Visible         =   0   'False
         Width           =   13965
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   9360
         MouseIcon       =   "F_EVALUATION.frx":37AA8
         MousePointer    =   99  'Custom
         Picture         =   "F_EVALUATION.frx":37DB2
         Top             =   6720
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Only Laboratory Manager can Close Lots"
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   7560
         TabIndex        =   53
         Top             =   7200
         Width           =   4140
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Specifications Table"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   2160
         MouseIcon       =   "F_EVALUATION.frx":3B194
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   2
         Left            =   2760
         Picture         =   "F_EVALUATION.frx":3B49E
         Top             =   1200
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   4
      Left            =   0
      MouseIcon       =   "F_EVALUATION.frx":3E880
      MousePointer    =   99  'Custom
      TabIndex        =   59
      Top             =   10680
      Width           =   1935
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Index           =   1
      Left            =   8400
      MouseIcon       =   "F_EVALUATION.frx":3EB8A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   10800
      Width           =   2655
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   3
      Left            =   3840
      MouseIcon       =   "F_EVALUATION.frx":3EE94
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   10680
      Width           =   1815
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   0
      Left            =   17520
      MouseIcon       =   "F_EVALUATION.frx":3F19E
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   10680
      Width           =   1575
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   2
      Left            =   14880
      MouseIcon       =   "F_EVALUATION.frx":3F4A8
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   10680
      Width           =   2175
   End
   Begin VB.Label La 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move Forward"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   17640
      MouseIcon       =   "F_EVALUATION.frx":3F7B2
      MousePointer    =   99  'Custom
      TabIndex        =   72
      Top             =   11595
      Width           =   1200
   End
   Begin VB.Label La 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move Previous"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   15630
      MouseIcon       =   "F_EVALUATION.frx":3FABC
      MousePointer    =   99  'Custom
      TabIndex        =   71
      Top             =   11600
      Width           =   1230
   End
   Begin VB.Label La 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Procedure"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   9000
      MouseIcon       =   "F_EVALUATION.frx":3FDC6
      MousePointer    =   99  'Custom
      TabIndex        =   70
      Top             =   11600
      Width           =   1200
   End
   Begin VB.Label La 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reading QC Info"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   4125
      MouseIcon       =   "F_EVALUATION.frx":400D0
      MousePointer    =   99  'Custom
      TabIndex        =   69
      Top             =   11595
      Width           =   1320
   End
   Begin VB.Label lbOperator 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   300
      TabIndex        =   60
      Top             =   11595
      Width           =   645
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   4
      Left            =   360
      MouseIcon       =   "F_EVALUATION.frx":403DA
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":406E4
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   15960
      MouseIcon       =   "F_EVALUATION.frx":43AC6
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":43DD0
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_EVALUATION.frx":471B2
      Height          =   480
      Index           =   1
      Left            =   9360
      MouseIcon       =   "F_EVALUATION.frx":4A594
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":4A89E
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   4560
      MouseIcon       =   "F_EVALUATION.frx":4DC80
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":4DF8A
      Top             =   11040
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   431.698
      X2              =   16836.23
      Y1              =   11133.9
      Y2              =   11133.9
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
      TabIndex        =   1
      Top             =   10200
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   12950.94
      X2              =   12950.94
      Y1              =   0
      Y2              =   12384.9
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   4316.981
      X2              =   4316.981
      Y1              =   0
      Y2              =   12384.9
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   8633.963
      X2              =   8633.963
      Y1              =   125.1
      Y2              =   12510
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_EVALUATION.frx":5136C
      Height          =   480
      Index           =   0
      Left            =   18240
      MouseIcon       =   "F_EVALUATION.frx":5474E
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":54A58
      Top             =   11040
      Width           =   480
   End
End
Attribute VB_Name = "F_EVALUATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IndexFormProcedura As Integer
Private IndexMainProcedura As Integer
Private IndexTextSelected As Integer
Private IndexText As Integer
Private MyLot As String

Private MyCodeID As String
Private MyCode As String
Private m_rc As Boolean
Private bFormSaved As Boolean
Private MeasurementUnit As String
Private UserDecimal As String
Private intDecimal As Integer
Private STD() As String
Private MeterNumber As Integer
Private pHNumber As Integer
Private pHMin(3) As String
Private pHMax(3) As String
Private MinNumber80Pec(10) As Integer
Private SelectedSTDNumber As Integer
Private strPassed As String
Private numSelectedStandard As String
Private AndOr As String

Private lCol As Long
Private lRow As Long
Private lRowMean As Long
Private bHoChiusoilLotto As Boolean
Private bAnotherFormCalled As Boolean
Private STDCount As Integer
Private strQC As String
Private bReadingClosed As Boolean
Private bSaveQC As Boolean
Private QCIndex As Integer
Private RangeMin As String
Private RangeMax As String
Private bUpdatedSTDFromDatabase As Boolean
Private UNIT_PP As String


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
Public Function DoShow(ByRef Index As Integer, Optional ByRef sLot As String, Optional ByRef sCode As String, Optional ByVal lngID As Long, Optional MyImage As Image, Optional FileName As String) As Boolean

    On Error GoTo ERR_SHOW
    
    'Set DefaultMenu(4) = MyImage
    IndexMainProcedura = Index
    m_rc = False
    bFormSaved = False
    
    
    QCIndex = 99
    
    MyCode = sCode
    
    
    
 
    
    
    SettingName = FileName
   
  
    MinNumberSelecterPerc = GetSetting(App.Title, "Options", "MinNumberSelecterPerc", 0.7)
    
    
    Call GrdRisultati(Grd3, "ppm")
    
    CheckUser 1
    FormPulisciTutto

            
    MyCodeID = lngID

       
    If sLot <> "" And sCode <> "" Then
        Call GetCodeInformation(sLot, sCode, lngID)
        
       
    Else
        PopupMessage 2, "Please select a valid Code/Lot..."
        Unload Me
    End If
    


    
    SelectProcedura 0
    mOk
    
    
    If MyOperatore.IndexPrivilege >= 1 Then
            
        Text1(11).Locked = False
        Text1(12).Locked = False
        Text1(12) = MyOperatore.Name
    
    End If
    
    
    Call SetGridEvaluationResults(Grid1)
    
    Me.Show vbModal
    
    If m_rc = True Then
        Index = IndexMainProcedura
        sLot = MyLot
        sCode = MyCode
    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    MsgBox Err.Description
    Resume ERR_END
End Function






Private Sub DefaultMenuLabel_Click(Index As Integer)
Dim MyIndex As Integer
Select Case Index
    Case 0
        ' vai avanti
        If IndexFormProcedura = 1 Then
            MyIndex = 0
        Else
            MyIndex = IndexFormProcedura + 1
        End If
        PicMenu_Click MyIndex
    Case 1
        'If bFormSaved Then
        If Not (bHoChiusoilLotto) Then
        
            If bUpdatedSTDFromDatabase Then
            
                If F_MsgBox.DoShow("STD Updated from Database." & vbCrLf & "Save new values?") Then SetSTDInFile
            
            End If
            
            SaveSelectedTest
            
            SaveResultsTable
            
            SaveQC
        End If
            'USER_PATH = USER_TEMP_PATH
            
            Unload Me
       ' Else
            
       ' End If
    Case 2
        ' torna indietro
        If IndexFormProcedura = 0 Then
            MyIndex = 1
        Else
            MyIndex = IndexFormProcedura - 1
        End If
        PicMenu_Click MyIndex
    Case 3
        Frame2.Visible = Not (Frame2.Visible)
    Case 4
         frmLogin.DoShow 1
        CheckUser 1
    Case 5
       ' Label7_Click
    Case 6
       ' Label6_Click
End Select

End Sub


Private Function CheckUser(ByVal Index As Integer)
Dim rc As Boolean
     
    rc = True
    Text1(11).Locked = True
    Text1(12).Locked = True
            
    If MyOperatore.IndexPrivilege < 1 Then rc = False
    If MyOperatore.Name = "" Then rc = False
    lbOperator = MyOperatore.Name
    Picture2.Visible = rc
    
    
    If Not (Grd3.Rows > 1) Then
        Picture2.Visible = False
    End If
    
    
    If rc Then
        If Grd3.Rows > 1 Then
            Text1(12) = MyOperatore.Name
            If Text1(11) = "" Then
                Text1(11) = FormatDataLAT(Now)
            Else
            
            End If
            Text1(11).Locked = False
            Text1(12).Locked = False
            
            
        End If
    Else
         Text1(12) = ""
    End If
    
    
End Function




Private Sub DelImage_Click()
If lRowMean > 0 Then
    If F_MsgBox.DoShow("Delete Mean Row : " & Grd3.Cell(lRowMean, 1).Text) Then
        
        Grd3.ReadOnly = False
        Grd3.Selection.DeleteByRow
        Grd3.ReadOnly = True
    
    End If

End If
End Sub

Private Sub DisableImage_Click()
PopupMessage 2, "Warning : Administrator Only can Operate...", , True
End Sub

Private Sub Form_Initialize()

Call SetPicForm
Call SetGrid

SaveSizes
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37
        DefaultMenuLabel_Click 2
    Case 39
        DefaultMenuLabel_Click 0
End Select
End Sub
Private Sub FormPulisciTutto()
Dim Ctl As Control
    For Each Ctl In Controls
        If TypeOf Ctl Is TextBox Then
            Ctl = ""
        ElseIf TypeOf Ctl Is Label Then
            If InStr(Ctl.Caption, "SET") Then
            Else
            Ctl.BackColor = vbColorLabelUnabled
            End If

        ElseIf TypeOf Ctl Is Grid Then
            Ctl.Rows = 1
        End If
    Next Ctl
    
   ' Picture3.BackColor = vbColorLabelUnabled
End Sub

Private Sub Form_Load()
IndexFormProcedura = 99
Dim i As Integer
'If Screen.Width - Me.Width > 1000 And bFullScreen Then
   ' Me.WindowState = 2
    For i = 0 To PicMain.Count - 1
        PicMain(i).Picture = LoadPicture(PictureMaxScreen)
        
   Next '
    'Me.Picture = LoadPicture(PictureMaxScreen)
'End If
End Sub

Private Sub Form_Resize()

ResizeControls
End Sub

Private Sub Frame2_Click()
Frame2.Visible = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Picture1.BackColor = &H4000&
Picture2.BackColor = &H4000&

Picture5.BackColor = &H4000&
Picture6.BackColor = &H4000&
End Sub

Private Sub frCommandInside_Click(Index As Integer)

    bSaveQC = True

    Call SetQC(Index)
    
    

End Sub


Private Function SetQC(ByVal Index As Integer)

    
    Select Case Index
        Case 0
            strQC = "Waiting"
            
        Case 1
            
            strQC = "Failed"
              
           
        Case 3
            strQC = "Passed"

        Case Else
            bSaveQC = False
            Exit Function
            
    End Select
    
    Picture4(0).BackColor = frCommandInside(Index).BackColor
    Frame3.BackColor = Picture4(0).BackColor
    Frame1.BackColor = Picture4(0).BackColor
     Frame4.BackColor = Picture4(0).BackColor
    Label2(6) = "QC : " & strQC
    
End Function

Private Sub Grd3_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

lRowMean = 0

    DelImage.Visible = False
    
    
    If FirstRow > 0 Then
        
        DelImage.Visible = True
        lRowMean = FirstRow
    
    End If
End Sub

Private Sub Label6_Click()
Picture5_Click
End Sub

Private Sub Label7_Click()
Picture5_Click
End Sub

Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub

Private Sub grd1_Click()
Frame2.Visible = False
End Sub

Private Sub Grd1_LostFocus()
'Grd1.Cell(0, 0).SetFocus
End Sub

Private Sub Grd1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim t As Integer
    SelectedSTDNumber = 0
    strPassed = ""
    Picture1.Visible = False
    PicMenu(3).Visible = False
If FirstRow > 0 Then
    Call SaveSelectedTest
    SelectedSTDNumber = CInt((Grd1.Cell(FirstRow, 1).Text))
    t = CInt((Grd1.Cell(FirstRow, 3).Text))
    numSelectedStandard = t 'STD(t, 1)
    Text1(7) = STD(t, 2)
    Text1(8) = STD(t, 3)
    Text1(9) = ""
    Text1(10) = ""
    DoEvents
    
    Call FillReadingsTablePerStandard(MeterNumber, SelectedSTDNumber)
    
    PicMenu(3).Visible = True
    Picture1.Visible = True
End If

End Sub

Private Sub Grd2_Click()
With Grd2
   Select Case lCol
        Case 1, 3, 5, 7, 9, 11, 13, 15, 17
            SetMeanValue (SelectedSTDNumber)
    End Select
End With
End Sub

Private Sub CheckpHTrue(ByVal t As Long)

Dim i As Integer
Dim rc1 As Boolean
Dim rc2 As Boolean
Dim rc3 As Boolean
Dim bValue As OLE_COLOR

    ' se è pH-----------------------------
    
    rc1 = True
    rc2 = True
    rc3 = True
    
    With Grd2
        .AutoRedraw = False
        rc1 = .Cell(t, MeterNumber * 2 + 1).Text
        rc2 = .Cell(t, MeterNumber * 2 + 3).Text
        rc3 = .Cell(t, MeterNumber * 2 + 5).Text
    
        If rc1 And rc2 And rc3 Then
    
            Else
            
            For i = 1 To MeterNumber
                .Cell(t, i * 2).BackColor = vbColorLabelUnabled
                .Cell(t, i * 2 - 1).BackColor = vbColorLabelUnabled
                .Cell(t, 0).BackColor = vbColorLabelUnabled
            Next
            DoEvents
            
        End If
        .AutoRedraw = True
        .Refresh
    End With

End Sub

Private Sub Grd2_LeaveCell(ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
    'Select Case Col
    '    Case 1, 3, 5, 7, 9, 11
    '        SetMeanValue (SelectedSTDNumber)
    'End Select
End Sub

Private Sub Grd2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
lCol = 0
lRow = 0
If FirstCol > 0 Then lCol = FirstCol
If FirstRow > 0 Then lRow = FirstRow
With Grd2
    Select Case FirstCol
        Case MeterNumber * 2 + 1, MeterNumber * 2 + 2
           ' lb(9) = "pH " & .Cell(FirstRow, MeterNumber * 2 + 5).Text & " Range"
           ' Text1(9) = .Cell(FirstRow, MeterNumber * 2 + 6).Text
           ' Text1(10) = .Cell(FirstRow, MeterNumber * 2 + 7).Text
        Case Else
    End Select

End With



End Sub

Private Sub Grd2_SetCellText(ByVal Row As Long, ByVal Col As Long, ByVal Text As String, Cancel As Boolean)
   ' Select Case Col
   '     Case 1, 3, 5, 7, 9, 11
   '         SetMeanValue (SelectedSTDNumber)
   ' End Select
End Sub

Private Sub Image2_Click()
   CheckPrivilege 1
   CheckUser 1
End Sub

Private Sub ImageTAV_Click(Index As Integer)
Select Case Index
    Case 5
        Picture1_Click
End Select
End Sub

Private Sub Label3_Click()

    Picture2_Click

End Sub

Private Sub PicMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Len(Text1(IndexText)) = 0 Then Text1(IndexText).BackColor = vbColorUnabled
IndexText = 0
Picture1.BackColor = &H4000&
Picture2.BackColor = &H4000&

Picture5.BackColor = &H4000&
Picture6.BackColor = &H4000&

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Integer
For i = 0 To PicMenu.Count - 1
    If i = IndexFormProcedura Then
    Else
        PicMenu(i).BackColor = vbColorDarkFont
    End If
Next
Picture1.BackColor = &H4000&
Picture2.BackColor = &H4000&
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set F_EVALUATION = Nothing
End Sub

Private Sub Image3_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub Label1_Click(Index As Integer)
Dim rc As Boolean

    Select Case Index
        Case 14
            rc = True
        Case 15
            rc = False
        Case Else
            Exit Sub
    End Select
    
    Label1(14).BackColor = IIf(rc, Picture4(0).BackColor, &H808080)
    Label1(15).BackColor = IIf(Not (rc), Picture4(0).BackColor, &H808080)
    
End Sub

Private Sub Label2_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub PicMenu_Click(Index As Integer)

DelImage.Visible = False

    If IndexFormProcedura = Index Then
    Else
    
        Call SelectProcedura(Index)
    
        
    End If
End Sub


Private Function SelectProcedura(ByVal Index As Integer) As Boolean
Dim rc As Boolean
Dim i As Integer

    'If Index > 3 Then Exit Function
    For i = 0 To PicMenu.Count - 1
        If i = Index Then
            PicMenu(i).BackColor = vbColorForeFixed
        Else
            PicMenu(i).BackColor = vbColorDarkFont
            
        End If
    Next
    Set Image4(0) = Image3(Index)
    Select Case Index
        Case 0
            Picture4(0).BackColor = &H4080&
           ' rc = False
           Frame3.Visible = True
           Frame4.Visible = False
        Case 1
           ' rc = True
            
            Frame4.Visible = True
            Picture4(0).BackColor = &H4060&
            Frame3.Visible = False
            lb(1).BackColor = Picture4(0).BackColor
            If CheckRequestedFiles Then
            
            End If
        Case 2
           ' rc = False
           ' Picture4(0).BackColor = &H60DF&
           ' reading QC
           'PopupMessage 2, "Open Reading form..."
           OpenReadingQC
           Exit Function
           
        Case 3
            ' GRAPH QC
            If Grd2.Rows > 1 Then
                bAnotherFormCalled = True
                SaveSelectedTest
                
                SaveResultsTable
                F_GRAPH.WindowState = Me.WindowState
                F_GRAPH.Top = Me.Top
                F_GRAPH.Left = Me.Left
                F_GRAPH.DoShow IndexFormProcedura, Text1(0), Text1(1), , , SettingName, CStr(SelectedSTDNumber)
                PicMenu_Click 0
            
            Else
                PopupMessage 2, "Please Select a valid Standard with at least 1 Test to open Graph QC...", , True
            End If
            bAnotherFormCalled = False
            Exit Function
        Case 4
        
            bAnotherFormCalled = True
            SaveSelectedTest
            SaveResultsTable
            
            Dim MyFGCode As String
            Dim MyFGID As Long
            
            If FormCodes.DoShow(MyFGCode, , MyFGID) Then
                F_CERTIFICATE.WindowState = Me.WindowState
                F_CERTIFICATE.Top = Me.Top
                F_CERTIFICATE.Left = Me.Left
                F_CERTIFICATE.DoShow IndexFormProcedura, Text1(0), MyFGCode, MyFGID, , SettingName
            End If
            Exit Function
    End Select
    Label2(6) = Label2(Index)
    IndexFormProcedura = Index
    PicMain(Index).Visible = True
    PicMain(Index).ZOrder
   ' blTable = Label2(IndexFormProcedura)
    Cleanform (False)
    
    SetQC (QCIndex)
End Function

Private Sub Cleanform(ByVal bValue As Boolean, Optional ByVal Index As Integer = 0)

End Sub

Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
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


Private Sub SetPicForm()
Dim i As Integer


For i = 0 To PicMain.Count - 1
    PicMain(i).Left = 0
    PicMain(i).Top = PicMenuBar(0).Height + Frame1.Height
    PicMain(i).Width = Me.Width
    PicMain(i).Height = Line1.Y1 - PicMain(i).Top
Next


For i = 0 To Text1.Count - 1
    Text1(i).BackColor = IIf(Len(Text1(i)) > 0, vbWhite, vbColorUnabled)
Next
End Sub

Private Sub Picture1_Click()
' salva i risultati calcolati
Call SaveResults
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Picture1.BackColor = &H8000&
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call SaveProcedure
End Sub

Private Sub Picture2_Click()
' chiudi il lotto
If Grd3.Rows < 2 Then
    MessageInfoTime = 2000
    PopupMessage 2, "This Lot cannot be closed : Please save at least 1 Mean Value..."
Else
    Call CloseLot
End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Picture2.BackColor = &H8000&
End Sub

Private Sub Picture5_Click()
If F_MsgBox.DoShow("This operation will Import stored STD Value Min & Max form Database.." & "Continue?", "STD : " & Text1(1)) Then
    UpdateSTDFromDatabase
End If
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Picture5.BackColor = &H8000&
End Sub

Private Sub Picture6_Click()
Picture5_Click
End Sub

Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Picture6.BackColor = &H8000&
End Sub

Private Sub Text1_Change(Index As Integer)
Dim rc As Boolean
rc = IIf(Len(Text1(Index)) > 0, True, False)
Text1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbWhite, vbColorUnabled)
Label1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextBlue, vbColorLabelUnabled)

Select Case Index
    Case 0
        lb(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextDarkBlue, vbColorLabelUnabled)
    Case 2, 4
        ' SELECTED MEAN!
        
        
        
    Case 7
       lb(7).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextDarkBlue, vbColorLabelUnabled)

    Case 9
       lb(9).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextDarkBlue, vbColorLabelUnabled)

End Select

End Sub
Private Sub Text1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Index = IndexText Then
   '
Else
    If Len(Text1(Index)) = 0 Then Text1(Index).BackColor = vbColorgotFocus
    If Len(Text1(IndexText)) = 0 Then Text1(IndexText).BackColor = vbColorUnabled

End If
IndexText = Index
End Sub
Private Sub Text1_Click(Index As Integer)
Text1(Index).BackColor = vbWhite
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index < Text1.Count - 1 Then Text1(Index + 1).SetFocus
End If
End Sub
Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).BackColor = vbWhite
ctlCalendar1.Visible = False
Select Case Index
    Case 11
        ctlCalendar1.ZOrder
        ctlCalendar1.Visible = True
End Select
End Sub

Private Sub ctlCalendar1_DateClicked(inputDate As Date)
'Select Case IndexTextSelected
  'Case 11
    Text1(11) = FormatDataLAT(CStr(inputDate))
'Case Else
'End Select
ctlCalendar1.Visible = False
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbWhite, vbColorUnabled)
End Sub


Private Function SaveProcedure()
Dim rc As Boolean

    rc = SaveForm
    'PicMenu(3).Visible = rc
    If rc Then
        MyLot = Text1(0)
        MyCode = Text1(1)
        PopupMessage 2, blTable & " Saved..."
    Else
    
    End If
    
    m_rc = rc
    bFormSaved = rc

End Function
Private Function SaveForm() As Boolean

Dim rc As Boolean
On Error GoTo ERR_SAVE:

    rc = True
    
ERR_END:
    On Error GoTo 0
    SaveForm = rc
    Exit Function
ERR_SAVE:
    rc = False
    GoTo ERR_END:
End Function

Private Function SetGrid()

       '------------------------------------------------
        '       SET TABELLA STD
        '------------------------------------------------
    With Grd1
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
      '  .DefaultFont.Size = 12 '* m_ControlGridFontSize
        .DefaultRowHeight = 30
        .Cols = 4
        .Cell(0, 0).Text = "n."
        .ReadOnly = False
        .Column(0).Width = 0
        .Range(0, 1, 0, 2).Merge
        .Cell(0, 1).Text = "Standard"
        .Column(1).Width = 129
        '.Cell(0, 2).Text = "STD Value"
        .Column(2).Width = 150
        .Cell(0, 3).Text = "t"
        .Column(3).Width = 0
        .Cell(0, 1).BackColor = vbColorTextLightBlue
        '.Cell(0, 1).BackColor = vbColorTextLightBlue
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With

End Function

Private Sub SetGrd2(ByVal Index As Integer)
 With Grd2
      .Rows = 1

        .AutoRedraw = False
        .ReadOnly = False
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
      '  .DefaultFont.Size = 12
          .DefaultRowHeight = 30
        .Cols = 20
        .FixedCols = 0
        
        .Cell(0, 0).Text = "# Test"
        .Column(0).Width = 65
        .Cell(0, 0).BackColor = vbColorTextLightBlue
        .Range(0, 1, 0, 2).Merge
        .Cell(0, 1).Text = "Meter 1"
        .Column(1).CellType = cellCheckBox
        .Column(1).Width = 30
        .Column(2).Width = 100
        .Cell(0, 1).BackColor = vbColorTextLightBlue

        If Index > 1 Then
            .Range(0, 3, 0, 4).Merge
            .Cell(0, 3).Text = "Meter 2"
            .Column(3).CellType = cellCheckBox
            .Column(3).Width = 30
            .Column(4).Width = 100
            .Cell(0, 3).BackColor = vbColorTextLightBlue

             If Index > 2 Then
                .Range(0, 5, 0, 6).Merge
                .Cell(0, 5).Text = "Meter 3"
                .Column(5).CellType = cellCheckBox
                .Column(5).Width = 30
                .Column(6).Width = 100
                .Cell(0, 5).BackColor = vbColorTextLightBlue
                 If Index > 3 Then
                    .Range(0, 7, 0, 8).Merge
                    .Cell(0, 7).Text = "Meter 4"
                    .Column(7).CellType = cellCheckBox
                    .Column(7).Width = 30
                    .Column(8).Width = 100
                    .Cell(0, 7).BackColor = vbColorTextLightBlue
                End If
            End If
        End If
    
    
    ' pH1
    
        If pHMin(1) <> "" Then
            .Range(0, Index * 2 + 1, 0, Index * 2 + 2).Merge
            .Cell(0, Index * 2 + 1).Text = "pH 1"
            .Column(Index * 2 + 1).CellType = cellCheckBox
            .Column(Index * 2 + 1).Width = 30
            .Column(Index * 2 + 2).Width = 80
        Else
            .Column(Index * 2 + 1).Width = 0
            .Column(Index * 2 + 2).Width = 0
        End If
        
    ' pH2
            
        If pHMin(2) <> "" Then
            .Range(0, Index * 2 + 3, 0, Index * 2 + 4).Merge
            .Cell(0, Index * 2 + 3).Text = "pH 2"
            .Column(Index * 2 + 3).CellType = cellCheckBox
            .Column(Index * 2 + 3).Width = 30
            .Column(Index * 2 + 4).Width = 80
        Else
            .Column(Index * 2 + 3).Width = 0
            .Column(Index * 2 + 4).Width = 0
        End If
        
    ' pH3
              
        If pHMin(3) <> "" Then
            .Range(0, Index * 2 + 5, 0, Index * 2 + 6).Merge
            .Cell(0, Index * 2 + 5).Text = "pH 3"
            .Column(Index * 2 + 5).CellType = cellCheckBox
            .Column(Index * 2 + 5).Width = 30
            .Column(Index * 2 + 6).Width = 80
        Else
            .Column(Index * 2 + 5).Width = 0
            .Column(Index * 2 + 6).Width = 0
        End If

              
        
        
        
        .Cell(0, Index * 2 + 3 + (pHNumber - 1) * 2).Text = "Index"
        .Column(Index * 2 + 3 + (pHNumber - 1) * 2).Width = 0
        .Cell(0, Index * 2 + 4 + (pHNumber - 1) * 2).Text = "Type"
        .Column(Index * 2 + 4 + (pHNumber - 1) * 2).Width = 0
        
        .Cols = Index * 2 + 5 + (pHNumber - 1) * 2
        
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .AutoRedraw = True
        .Refresh
    End With
    
End Sub


Private Function GetFormSettingName()
Dim i As Integer
Dim t As Integer

    ' doppio check... si sa mai...
    
   If FileExists(USER_TEMP_PATH & SettingName) Then
        USER_PATH = USER_TEMP_PATH
   ElseIf FileExists(USER_DATA_PATH & SettingName) Then
        'PopupMessage 2, "Lot : " & Text1(0) & vbCrLf & "Code : " & Text1(1) & vbCrLf & "Is Closed..."
        USER_PATH = USER_DATA_PATH
   Else
        PopupMessage 2, "Lot : " & Text1(0) & vbCrLf & "Code : " & Text1(1) & vbCrLf & "Warning : Information QC not found...", , True

   End If
   
   
    CloseSettingDataFile
    
 
    SelectedSTDNumber = 0
    
    ' get variables
    
    
    bReadingClosed = GetSettingData(SettingName, "Reading", "Closed", False)
    
    GetSavedQC
    
    UserDecimal = FormatDecimal(GetSettingData(SettingName, "Code Information", "Decimal", 0))
    intDecimal = GetSettingData(SettingName, "Code Information", "Decimal", 0)
    MeasurementUnit = IIf(IsNull(Trim(dbTabCode!MeasurementUnit)), "", Trim(dbTabCode!MeasurementUnit))
    MeterNumber = GetSettingData(SettingName, "Information QC", "MeterNumber", 1)
    pHNumber = GetSettingData(SettingName, "Information QC", "pHNumber", 1)
    
    Text1(11) = GetSettingData(SettingName, "Close QC", "Validation Date", "")
    Text1(12) = GetSettingData(SettingName, "Close QC", "Operator", "")
    
    
    
    
    
    Call GetphValue
    
    ' Riempi tabella STANDARD
    
     Call GetSTDTable
     
    ' imposta tabella Meter
    
     SetGrd2 (MeterNumber)
    
    GetMeanTable Grd3, SettingName
    Call SortGrid(Grd3)
   'Grd3.Cell(0, 2).Text = "Target Value " & AndOr & " U [" & UNIT_PP & "]"
    CloseSettingDataFile
    
    Call SetQCFromfile
    
End Function

Private Sub SetQCFromfile()


If Grd3.Rows > 1 Then
    frCommandInside(0).Visible = bReadingClosed
    frCommandInside(1).Visible = bReadingClosed
    frCommandInside(3).Visible = bReadingClosed
    
    If bReadingClosed Then
    Else
    
    End If
Else
    frCommandInside(0).Visible = False
    frCommandInside(1).Visible = False
    frCommandInside(3).Visible = False
    
    
End If

End Sub
Private Sub GetCodeInformation(ByVal sLot As String, ByVal sCode As String, Optional ByVal MyID As Long)
Dim MeasurementUnit As String
    ' attenzione , se ho un file allora lo importo , altrimenti prendo i dati del Code
    
   GetFormSettingName
   
   
     
    With dbTabCode
        .filter = ""
        If MyCodeID > 0 Then
            .filter = "ID='" & MyCodeID & "'"
            Debug.Print !Code
        Else
        
            .filter = "Code='" & sCode & "'"
        End If
    
     
            'Select Case IIf(IsNull(Trim(!AndOr)), "&", Trim(!AndOr))
                   ' Case "&"
                       ' AndOr = Chr$(177)
                    'Case UCase("or")
                    
        '--------------------------------------------------
        ' carattere AND fisso nella tabella evaluation
        '--------------------------------------------------
        
        AndOr = Chr$(247)
        
        
                   ' Case Else
                    
               ' End Select
                
               ' Debug.Print AndOr
                
       
        If .EOF Then
            MessageInfoTime = 2000
            PopupMessage 2, "Cannot find Hanna Code  :  " & sCode & vbCrLf & "Please Enter Code Information..."
            'Unload Me
        Else
        
            MeasurementUnit = IIf(IsNull(Trim(!MeasurementUnit)), "", " " & Trim(!MeasurementUnit))
            
            Text1(0) = sLot
            Text1(1) = sCode
            
            If InStr(MeasurementUnit, "mg") Then
                UNIT_PP = "ppm"
            Else
                UNIT_PP = "ppb"
            End If
            Grd3.Cell(0, 1).Text = "Standard Value [" & UNIT_PP & "]"
            Grd3.Cell(0, 2).Text = "Target Value [" & UNIT_PP & "]"
            Grd3.Cell(0, 3).Text = "Mean Value [" & UNIT_PP & "]"
            Grd3.Cell(0, 4).Text = "Tot Average [" & UNIT_PP & "]"
    
        End If
    
    End With




End Sub


Private Sub GetSTDTable()
    Dim rc As Boolean
    Dim t As Integer
    
    Dim STDNumber As String
    Dim STDValue As String
    Dim STDMin As String
    Dim STDMax As String
    
    
    On Error GoTo ERR_GET
    
    
    rc = True
    
    STDCount = CInt(GetSettingData(SettingName, "Graph QC", "STDCount", 0))
    
    If STDCount = 0 Then
        rc = False
        GoTo ERR_END
    End If
    
    ReDim STD(STDCount, 4) As String
        
    For t = 1 To STDCount

        STD(t, 0) = GetSettingData(SettingName, "Graph QC", "STDNumber" & t, "")
        STD(t, 1) = GetSettingData(SettingName, "Graph QC", "STDValue" & t, "")
        STD(t, 2) = GetSettingData(SettingName, "Graph QC", "STDMin" & t, "")
        STD(t, 3) = GetSettingData(SettingName, "Graph QC", "STDMax" & t, "")
     
    Next
    
    
ERR_END:
    On Error GoTo 0
    
    CloseSettingDataFile
    
    If rc Then
        Call FillSTDGrid
    End If
    Exit Sub
ERR_GET:
    rc = False
    MsgBox Err.Description
    Resume Next
End Sub

Private Function SetSTDInFile()

Dim rc As Boolean
Dim t As Integer
Dim STDNumber As String
Dim STDValue As String
Dim STDMin As String
Dim STDMax As String
On Error GoTo ERR_SET
If STDCount = 0 Then
    STDCount = CInt(GetSettingData(SettingName, "Graph QC", "STDCount", 0))
End If
 
    If STDCount = 0 Then
        rc = False
        GoTo ERR_END
    End If
    
    For t = 1 To UBound(STD)
    
        SaveSettingData SettingName, "Graph QC", "STDNumber" & t, STD(t, 0)
        SaveSettingData SettingName, "Graph QC", "STDValue" & t, STD(t, 1)
        SaveSettingData SettingName, "Graph QC", "STDMin" & t, STD(t, 2)
        SaveSettingData SettingName, "Graph QC", "STDMax" & t, STD(t, 3)

    
    Next
   
ERR_END:
    On Error GoTo 0
    
    CloseSettingDataFile
    
    If rc Then
        PopupMessage 2, "STD updated..."
    End If
    Exit Function
ERR_SET:
    rc = False
    MsgBox Err.Description
    Resume Next


End Function

Private Function UpdateSTDFromDatabase()
   Dim rc As Boolean
    Dim t As Integer
    
    Dim STDNumber As String
    Dim STDValue As String
    Dim STDMin As String
    Dim STDMax As String
    
    
    On Error GoTo ERR_GET
    
    
    rc = True
    
    STDCount = CInt(GetSettingData(SettingName, "Graph QC", "STDCount", 0))
    
    If STDCount = 0 Then
        rc = False
        GoTo ERR_END
    End If
    
    ReDim STD(STDCount, 4) As String
        
    For t = 1 To STDCount

        STD(t, 0) = GetSettingData(SettingName, "Graph QC", "STDNumber" & t, "")
        
        ' importo gli standard dal database
        With dbTabCode

                Debug.Print !Code

        
            If .EOF Then
            
                STD(t, 1) = GetSettingData(SettingName, "Graph QC", "STDValue" & t, "")
                STD(t, 2) = GetSettingData(SettingName, "Graph QC", "STDMin" & t, "")
                STD(t, 3) = GetSettingData(SettingName, "Graph QC", "STDMax" & t, "")
        
            Else
                
                Debug.Print !Code
                
                If IsNumeric(STD(t, 0)) Then
                    Select Case CInt(STD(t, 0))
                        Case 1
                            If IsNumeric(dbTabCode!STD1Value) = False Then GoTo cont
                            STD(t, 1) = IIf(IsNull(Trim(dbTabCode!STD1Value)), "", Trim(dbTabCode!STD1Value))      '& "Value"
                            STD(t, 2) = IIf(IsNull(Trim(dbTabCode!STD1Min)), "", Trim(dbTabCode!STD1Min))      '& "Min"
                            STD(t, 3) = IIf(IsNull(Trim(dbTabCode!STD1Max)), "", Trim(dbTabCode!STD1Max))       '& "Max"
                        
                        Case 2
                            If IsNumeric(dbTabCode!STD2Value) = False Then GoTo cont
                            STD(t, 1) = IIf(IsNull(Trim(dbTabCode!STD2Value)), "", Trim(dbTabCode!STD2Value))      '& "Value"
                            STD(t, 2) = IIf(IsNull(Trim(dbTabCode!STD2Min)), "", Trim(dbTabCode!STD2Min))      '& "Min"
                            STD(t, 3) = IIf(IsNull(Trim(dbTabCode!STD2Max)), "", Trim(dbTabCode!STD2Max))       '& "Max"
                            
                        Case 3
                            If IsNumeric(dbTabCode!STD3Value) = False Then GoTo cont
                            STD(t, 1) = IIf(IsNull(Trim(dbTabCode!STD3Value)), "", Trim(dbTabCode!STD3Value))      '& "Value"
                            STD(t, 2) = IIf(IsNull(Trim(dbTabCode!STD3Min)), "", Trim(dbTabCode!STD3Min))      '& "Min"
                            STD(t, 3) = IIf(IsNull(Trim(dbTabCode!STD3Max)), "", Trim(dbTabCode!STD3Max))       '& "Max"
                        
                        Case 4
                            If IsNumeric(dbTabCode!STD4Value) = False Then GoTo cont
                            STD(t, 1) = IIf(IsNull(Trim(dbTabCode!STD4Value)), "", Trim(dbTabCode!STD4Value))      '& "Value"
                            STD(t, 2) = IIf(IsNull(Trim(dbTabCode!STD4Min)), "", Trim(dbTabCode!STD4Min))      '& "Min"
                            STD(t, 3) = IIf(IsNull(Trim(dbTabCode!STD4Max)), "", Trim(dbTabCode!STD4Max))       '& "Max"
                        Case 5
                            If IsNumeric(dbTabCode!STD5Value) = False Then GoTo cont
                            STD(t, 1) = IIf(IsNull(Trim(dbTabCode!STD5Value)), "", Trim(dbTabCode!STD5Value))      '& "Value"
                            STD(t, 2) = IIf(IsNull(Trim(dbTabCode!STD5Min)), "", Trim(dbTabCode!STD5Min))      '& "Min"
                            STD(t, 3) = IIf(IsNull(Trim(dbTabCode!STD5Max)), "", Trim(dbTabCode!STD5Max))       '& "Max"
                        
                        Case 6
                            If IsNumeric(dbTabCode!STD6Value) = False Then GoTo cont
                            STD(t, 1) = IIf(IsNull(Trim(dbTabCode!STD6Value)), "", Trim(dbTabCode!STD6Value))      '& "Value"
                            STD(t, 2) = IIf(IsNull(Trim(dbTabCode!STD6Min)), "", Trim(dbTabCode!STD6Min))      '& "Min"
                            STD(t, 3) = IIf(IsNull(Trim(dbTabCode!STD6Max)), "", Trim(dbTabCode!STD6Max))
                        Case Else
                        
                            STD(t, 1) = GetSettingData(SettingName, "Graph QC", "STDValue" & t, "")
                            STD(t, 2) = GetSettingData(SettingName, "Graph QC", "STDMin" & t, "")
                            STD(t, 3) = GetSettingData(SettingName, "Graph QC", "STDMax" & t, "")
                    
                    End Select
                    
                    
                Else
cont:
                    STD(t, 1) = GetSettingData(SettingName, "Graph QC", "STDValue" & t, "")
                    STD(t, 2) = GetSettingData(SettingName, "Graph QC", "STDMin" & t, "")
                    STD(t, 3) = GetSettingData(SettingName, "Graph QC", "STDMax" & t, "")
                
                End If
                
            End If
        End With
        
     
    Next
    
    
ERR_END:
    On Error GoTo 0
    
    CloseSettingDataFile
    
    If rc Then
        bUpdatedSTDFromDatabase = True
        PopupMessage 2, "STD updated..."
        Call FillSTDGrid
    End If
    Exit Function
ERR_GET:
    rc = False
    MsgBox Err.Description
    Resume Next

End Function

Private Function FillSTDGrid()

Dim i As Integer
Dim t As Integer
With Grd1
    .Rows = 1
    .ReadOnly = True
    .AutoRedraw = False
    t = 1
    For t = 1 To STDCount
        .AddItem "", False
        .Cell(.Rows - 1, 1).Text = STD(t, 0)
        .Cell(.Rows - 1, 2).Text = STD(t, 1)
        .Cell(.Rows - 1, 3).Text = t
        .Cell(.Rows - 1, 1).ForeColor = vbColorDarkUnabled
        .Cell(.Rows - 1, 2).ForeColor = vbColorDarkUnabled ' vbColorForeFixed
       
    Next
    .Column(1).Sort cellAscending
    .AutoRedraw = True
    .Refresh
End With

End Function




Private Function FillReadingsTablePerStandard(ByVal NumMeter As Integer, ByVal STDNumber As Integer) As Boolean
Dim rc As Boolean
Dim AllTests As String
Dim MyCols As String
Dim MeterValue As String
Dim sString As String
Dim ReadingSelected As Boolean
Dim pHValue As String
Dim MinValue As String
Dim MaxValue As String
Dim MyColor As OLE_COLOR
Dim i As Integer
Dim t As Integer
Dim RowCount As Integer
Dim ReadingCount As Integer
Dim STDNumberGrid As String
Dim NumberTest As String
Dim TestType As String
On Error GoTo ERR_GET
    rc = True
    
    Grd2.Rows = 1

    AllTests = CInt(GetSettingData(SettingName, "Graph QC", "Standard " & STDNumber & " Total Tests", 1))

    MinValue = Text1(7)
    MaxValue = Text1(8)
            
    If AllTests >= 1 Then
      
        With Grd2
         '   .DefaultFont.Size = 12 ' * m_ControlGridFontSize
            .AutoRedraw = False
            
            RowCount = 0
            ReadingCount = 0
            
            For i = 1 To AllTests
            
                    .AddItem "", False
                    '.Cell(.Rows - 1, t * 2 - 1).Text =
                    
                       NumberTest = GetSettingData(SettingName, "Graph QC", "Standard " & STDNumber & " Test " & i & " Real Test", "")
                       SaveSettingData SettingName, "Reading QC", "Grd2 Row" & NumberTest & " STD Test", i
                        .Cell(.Rows - 1, 0).Text = NumberTest
                    For t = 1 To NumMeter
                         
                         MeterValue = GetSettingData(SettingName, "Graph QC", "Standard " & STDNumber & " Test " & i & " Meter " & t & " Value", "")
                         ReadingSelected = GetSettingData(SettingName, "Graph QC", "Standard " & STDNumber & " Test " & i & " Meter " & t & " Selected", "TRUE")
                         
                         Call CheckValue(MeterValue, MinValue, MaxValue, MyColor)
                         
                         If MeterValue <> "" Then
                            ReadingCount = ReadingCount + 1
                            .Cell(.Rows - 1, t * 2 - 1).Text = ReadingSelected
                        Else
                            .Cell(.Rows - 1, t * 2 - 1).Text = False
                            .Cell(.Rows - 1, t * 2 - 1).Locked = True
                        End If
                        
                        .Cell(.Rows - 1, t * 2).Text = MeterValue       ' GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & 10 + t, "")
                        .Cell(.Rows - 1, t * 2).ForeColor = MyColor     ' GetSettingData(SettingName, "Reading QC", "Grd2 Fore" & i & 10 + t, vbBlack)
                        .Cell(.Rows - 1, t * 2).Locked = True
                        If IsNumeric(NumberTest) Then
                            SaveSettingData SettingName, "Reading QC", "Grd2 Fore Row" & NumberTest & " Col" & 10 + t, MyColor
                            
                        End If
                    Next
                    
                    
                    For t = 0 To 2
                      
                      pHValue = GetSettingData(SettingName, "Graph QC", "Standard " & STDNumber & " Test " & i & " pH" & t + 1 & " " & " Value", "")
                      
                      Call CheckpHValue(pHValue, pHMin(t + 1), pHMax(t + 1), MyColor)
                      
                      .Cell(.Rows - 1, NumMeter * 2 + 1 + t * 2).Text = True
                      .Cell(.Rows - 1, NumMeter * 2 + 2 + t * 2).Text = pHValue
                      .Cell(.Rows - 1, NumMeter * 2 + 2 + t * 2).ForeColor = MyColor
                      .Cell(.Rows - 1, NumMeter * 2 + 2 + t * 2).Locked = True
                    Next
            Next
            
            MinNumber80Pec(STDNumber) = Int(ReadingCount * MinNumberSelecterPerc) '  Int((.Rows - 1) * NumMeter * MinNumberSelecterPerc)  ' STDNumber=0 80% dei totali
            
            .AutoRedraw = True
            .ReadOnly = False
            .Refresh
        End With
     
    End If
    
    

    
    
ERR_END:
    On Error GoTo 0
    
    lb80Pec = "Select at least " & MinNumberSelecterPerc * 100 & "% of all value : " & MinNumber80Pec(STDNumber)
    
    If rc Then
        If STDNumber > 0 Then Call SetMeanValue(STDNumber)
    End If
    FillReadingsTablePerStandard = rc
    Exit Function
ERR_GET:
    rc = False
  
    MsgBox Err.Description
    Resume Next
    

    
End Function
Public Function CheckpHValue(ByVal Value As String, ByVal MinValue As String, ByVal MaxValue As String, ByRef MyColor As OLE_COLOR)

    If IsNumeric(Value) Then
        If IsNumeric(MinValue) And IsNumeric(MaxValue) Then
            If CDbl(Value) > 0 Then
                If CDbl(Value) >= CDbl(MinValue) And CDbl(Value) <= CDbl(MaxValue) Then
                    MyColor = vbBlack
                    
                Else
                    MyColor = vbRed
                   
                End If
            End If
        Else
            MyColor = vbBlack
        End If
    End If
End Function

Public Function CheckValue(ByVal MeterValue As String, ByVal MinValue As String, ByVal MaxValue As String, ByRef MyColor As OLE_COLOR)

    Debug.Print CDbl(MaxValue)
    If MaxValue = 0 Then
    
        GoTo OK:
    
    End If
    If IsNumeric(MeterValue) Then
        If IsNumeric(MinValue) And IsNumeric(MaxValue) Then
            If CDbl(MeterValue) >= 0 Then
                If CDbl(MeterValue) >= CDbl(MinValue) And CDbl(MeterValue) <= CDbl(MaxValue) Then
                    MyColor = vbBlack
                    
                Else
                    MyColor = vbRed
                   
                End If
            End If
        Else
OK:
            MyColor = vbBlack
        End If
    End If
End Function

Private Function SetMeanValue(ByVal STDNumber As Integer) As Boolean
Dim sSTDNumerString As String
Dim ReadingSelected As Boolean
Dim SelectedAverage As Double
Dim SelectedValue As Double
Dim ReadValue As Double
Dim SelectedCount As Integer
Dim TotalAverage As Double
Dim TotalReadingsCount As Integer
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim pHTrue As Boolean
Dim pHTrue1 As Boolean
Dim pHTrue2 As Boolean
Dim pHTrue3 As Boolean
Dim SelRows() As Integer

Dim Selected() As Double
Dim StdDeviation As Double
Dim StdDeviatioPerc As Double
Dim Repeatability As Double

    On Error GoTo ERR_SET
    
    
    If Me.Visible = False Then Exit Function

    rc = True
    
    ReDim SelRows(MeterNumber) As Integer
    
    
    If STDNumber = 0 Then ' totali
        sSTDNumerString = "STD All"
        lbSTDNumber.BackColor = vbColorDarkUnabled
      '  Picture3.BackColor = lbSTDNumber.BackColor
        Picture1.Visible = False
    Else
        sSTDNumerString = "STD Number : " & STDNumber
        lbSTDNumber.BackColor = &H4080&
       ' Picture3.BackColor = lbSTDNumber.BackColor
       Picture1.Visible = True
    End If


   ' SelectedAverage
    SelectedCount = 0
    SelectedAverage = 0
    SelectedValue = 0
    With Grd2
        .AutoRedraw = False
        If .Rows = 1 Then
            rc = False
            GoTo ERR_END
        End If
        For i = 1 To .Rows - 1
            
            For t = 1 To MeterNumber
                If Trim(.Cell(i, t * 2).Text) = "" Or IsNull(.Cell(i, t * 2).Text) Then
                    GoTo salta
                Else
                    If IsNumeric((.Cell(i, t * 2).Text)) Then
                    ReadValue = CDbl(.Cell(i, t * 2).Text)
                    Else
                        ReadValue = 0
                    End If
                End If
                
                TotalAverage = TotalAverage + ReadValue
                TotalReadingsCount = TotalReadingsCount + 1

                pHTrue1 = (.Cell(i, MeterNumber * 2 + 1).Text)
                pHTrue2 = (.Cell(i, MeterNumber * 2 + 3).Text)
                pHTrue3 = (.Cell(i, MeterNumber * 2 + 5).Text)
                pHTrue = IIf(pHTrue1 And pHTrue2 And pHTrue3, True, False)

                
                ReadingSelected = .Cell(i, t * 2 - 1).Text = True
               
                
                If ReadingSelected Then ' se ho la spunta lo conto
                    If .Cell(i, t * 2).Text <> "" Then
                    
                        If pHTrue Then ' controllo il PH! se ho la spunta conto tutto....
                            SelRows(t) = SelRows(t) + 1
                            SelectedValue = ReadValue ' CDbl(.Cell(i, t * 2).Text)
                            SelectedCount = SelectedCount + 1
                            SelectedAverage = SelectedAverage + SelectedValue
                            
                             ReDim Preserve Selected(SelectedCount - 1)
                             Selected(SelectedCount - 1) = SelectedValue
                            
                        End If
                    End If
                    .Cell(i, t * 2 - 1).BackColor = vbWhite '&HF0F0F0 '
                    .Cell(i, t * 2).BackColor = vbWhite '&HF0F0F0 '
                    .Cell(i, 0).BackColor = vbWhite '&HF0F0F0 '
                Else
                    .Cell(i, t * 2 - 1).BackColor = vbColorMediumFixed
                    .Cell(i, t * 2).BackColor = vbColorMediumFixed
                    .Cell(i, 0).BackColor = vbColorMediumFixed '&HF0F0F0 '
                End If
salta:
                .Cell(0, t * 2).BackColor = vbColorLightFixed
            Next
            CheckpHTrue i
        Next
        .AutoRedraw = True
        .Refresh
    End With
    If SelectedCount = 0 Then
        SelectedAverage = 0
   
    Else
        SelectedAverage = SelectedAverage / SelectedCount
    End If
    

    
    If TotalReadingsCount = 0 Then GoTo ERR_END:
    
    
    TotalAverage = TotalAverage / TotalReadingsCount '((Grd2.Rows - 1) * MeterNumber)
    
'Dim StdDeviation As Double
'Dim StdDeviatioPerc As Double
'Dim Repeatability As Double
    
     StdDeviation = Format$(CalcolaSTDEV(Selected()), UserDecimal)
     If SelectedAverage <> 0 Then
     StdDeviatioPerc = Format$((StdDeviation / SelectedAverage) * 100, UserDecimal)
     Else
     StdDeviatioPerc = 0
     End If
     Repeatability = Format$(StdDeviation * 2.26 * Sqr(2), UserDecimal)
    
    If MinNumber80Pec(STDNumber) <= SelectedCount Then ' OK
    
    Else
        rc = False
    End If
    Dim Max As Integer
    For i = 0 To UBound(SelRows)
        If SelRows(i) > Max Then
            Max = SelRows(i)
        End If
    Next
ERR_END:
    On Error GoTo 0
    lbSTDNumber = sSTDNumerString
    Text1(2) = FormatNumber(TotalAverage, intDecimal)
    Text1(2) = Format$(Text1(2), UserDecimal)
    
    Text1(3) = TotalReadingsCount '* MeterNumber
    Text1(13) = (Grd2.Rows - 1)
    Text1(4) = FormatNumber(SelectedAverage, intDecimal)
    Text1(4) = Format$(Text1(4), UserDecimal)
    Text1(5) = SelectedCount
    Text1(5).ForeColor = IIf(rc, vbBlack, vbRed)
    Text1(6) = Max
    
    
    '-----------------------------------------
    ' riempio la  griglia Grid1
    '-----------------------------------------
    
    
         '.Cell(1, 1).Text = "STD"
         '.Cell(2, 1).Text = "# Readings"
         '.Cell(3, 1).Text = "# Tests"
         '.Cell(4, 1).Text = "Total Average"
         '.Cell(5, 1).Text = "# Readings (Selected)"
         '.Cell(6, 1).Text = "# Tests (Selected)"
         '.Cell(7, 1).Text = "Mean Value (Selected)"
         '.Cell(8, 1).Text = "Std Deviation"
         '.Cell(9, 1).Text = "Std Deviation %"
         '.Cell(10, 1).Text = "Repeatability"
       
    With Grid1
        .AutoRedraw = False
        
         .Cell(1, 2).Text = STDNumber
         .Cell(2, 2).Text = TotalReadingsCount
         .Cell(3, 2).Text = (Grd2.Rows - 1)
         .Cell(4, 2).Text = Format$(TotalAverage, UserDecimal)
         .Cell(5, 2).Text = SelectedCount
         .Cell(6, 2).Text = Max
         .Cell(7, 2).Text = Format$(SelectedAverage, UserDecimal)
         .Cell(8, 2).Text = StdDeviation
         .Cell(9, 2).Text = StdDeviatioPerc & " %"
         .Cell(10, 2).Text = Repeatability
         
         For i = 1 To .Rows - 1
            .Cell(i, 2).FontBold = False
            .Cell(i, 2).ForeColor = vbColorBlueProgram
            .Cell(i, 2).Alignment = cellCenterCenter
            .Cell(i, 1).Alignment = cellRightCenter
         Next
         
         .Refresh
         .AutoRedraw = True
    End With
    
    Call CheckValuePassed(2)
    Call CheckValuePassed(4)
    
    FillGrd3
    
    
    CloseSettingDataFile
    
    SaveSettingData SettingName, "Evaluation QC", "StdDeviation" & STDNumber, StdDeviation
    SaveSettingData SettingName, "Evaluation QC", "StdDeviatioPerc" & STDNumber, StdDeviatioPerc
    SaveSettingData SettingName, "Evaluation QC", "Repeatability" & STDNumber, Repeatability
    
    
    CloseSettingDataFile

        
    Picture1.Visible = rc
    SetMeanValue = rc
    Exit Function
ERR_SET:
    rc = False
  
    MsgBox Err.Description
    Resume Next
End Function

Private Sub CheckValuePassed(ByVal Index As Integer)

'Debug.Print CDbl(Text1(Index)), CDbl(Text1(7)), CDbl(Text1(8))
    If Text1(7) <> "" And Text1(8) <> "" And Text1(7) <> "/" And Text1(8) <> "/" Then
        If CDbl(Text1(Index)) >= CDbl(Text1(7)) And CDbl(Text1(Index)) <= CDbl(Text1(8)) Then
            Text1(Index).ForeColor = vbBlack
            strPassed = "YES"
        Else
            Text1(Index).ForeColor = vbRed
            strPassed = "NO"
        End If
    End If
    
End Sub
Private Sub SaveQC()

    bSaveQC = IIf(bReadingClosed, bSaveQC, False)

    If bSaveQC Then
    
        If F_MsgBox.DoShow("Save QC?", "QC : " & strQC) Then
    
            CloseSettingDataFile
            SaveSettingData SettingName, "Evaluation QC", "ResultQC", strQC
            SaveSettingData SettingName, "Evaluation QC", "ResultQC Date", Now()
            SaveSettingData SettingName, "Evaluation QC", "ResultQC Operator", MyOperatore.Name
            CloseSettingDataFile
            
        End If
        
        
    End If

End Sub

Private Sub GetSavedQC()

    
    strQC = GetSettingData(SettingName, "Evaluation QC", "ResultQC", "")
    
    Select Case strQC
       
        Case "Waiting"
            QCIndex = 0
        Case "Failed"
            QCIndex = 1
        Case "Passed"
            QCIndex = 3
    End Select
       
    
End Sub

Private Function SaveSelectedTest()
Dim i As Integer
Dim t As Integer
Dim rc As Boolean
Dim ReadingAvg As Double
Dim MaxCount As Integer
Dim Reading As Double
    CloseSettingDataFile
    
    If SelectedSTDNumber = 0 Then Exit Function
    
    
    With Grd2
        For t = 1 To .Rows - 1
            ReadingAvg = 0
            MaxCount = 0
            For i = 2 To MeterNumber * 2 Step 2
               ' If .Cell(t, i).BackColor = vbColorLabelUnabled Then
                If UCase(.Cell(t, i - 1).Text) = "FALSE" Then
                    SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Test " & t & " Meter " & i / 2 & " Selected", "False"
                Else
                    MaxCount = MaxCount + 1
                    Reading = CDbl(.Cell(t, i).Text)
                    ReadingAvg = ReadingAvg + Reading
                    SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Test " & t & " Meter " & i / 2 & " Selected", .Cell(t, i - 1).Text
                End If
            Next
            If MaxCount = 0 Then
                SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Test " & t & " ReadingAvg", "NULL"
            Else
                ReadingAvg = ReadingAvg / MaxCount
                SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Test " & t & " ReadingAvg", ReadingAvg
            End If
            
        Next
    End With
    
    rc = CheckAverage
    
    If rc Then
    
salva:
        If Text1(2) <> "" Then SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Total Average", Text1(2)
        If Text1(3) <> "" Then SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Total Readings", Text1(3)
        If Text1(13) <> "" Then SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Total Tests", Text1(13)
        If Text1(4) <> "" Then SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Selected Average", Text1(4)
        If Text1(5) <> "" Then SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Selected Readings", Text1(5)
        If Text1(6) <> "" Then SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Selected Tests", Text1(6)
        'Debug.Print Text1(6)
    Else
    
        SetMeanValue (SelectedSTDNumber)
        GoTo salva:
    
    End If
    CloseSettingDataFile
    
End Function

Private Function CheckAverage() As Boolean
Dim rc As Boolean
Dim i As Integer

    rc = True
    If Text1(2) = "" Then rc = False
    If Text1(3) = "" Then rc = False
    If Text1(13) = "" Then rc = False
    If Text1(4) = "" Then rc = False
    If Text1(5) = "" Then rc = False
    If Text1(6) = "" Then rc = False
    CheckAverage = rc
    
    
End Function

Private Function SaveResults() As Boolean

Dim rc As Boolean
    
Dim i As Integer
Dim t As Integer

On Error GoTo ERR_SAVE
rc = True
    
    
    If SelectedSTDNumber = 0 Then Exit Function
    
    Call SaveSelectedTest
        ' salva tabella Test SelectedSTDNumber
    Call FillGrd3

    CloseSettingDataFile
    
    PutLotInDatabase

    CloseSettingDataFile


ERR_END:
    On Error GoTo 0
    SaveResults = rc
    If rc Then
        SetQCFromfile
        PopupMessage 2, "Mean Value for Standard n " & STD(numSelectedStandard, 0) & " (" & STD(numSelectedStandard, 1) & ") " & vbCrLf & "Saved...", , , "Mean Results"
    End If
    Exit Function
ERR_SAVE:
    rc = False
    MsgBox Err.Description
    Resume Next
    
End Function

Private Function PutLotInDatabase() As Boolean

Dim rc As Boolean

On Error GoTo ERR_SAVE:
rc = True
With dbTabReport
    .filter = ""
    .filter = "Lot='" & Trim(Text1(0)) & "' and Code='" & Trim(Text1(1)) & "' and NomeFile='" & SettingName & "'"
    
    If .EOF Then
        ' come è possibile????
        Exit Function
    Else
    End If
        !Evaluation = IIf(Grd3.Rows > 1, True, False)
        .Update
End With

ERR_END:
    On Error GoTo 0
    PutLotInDatabase = rc
    Exit Function
ERR_SAVE:
    MsgBox Err.Description
    rc = False
    Resume ERR_END
End Function

Private Function SaveResultsTable()
    
Dim i As Integer
Dim t As Integer

    ' salva tabella Results
    
    'MsgBox USER_PATH

    With Grd3
        If .Rows > 1 Then
            SaveSettingData SettingName, "Evaluation QC", "Results Grid Rows", .Rows
            SaveSettingData SettingName, "Evaluation QC", "Results Grid Cols", .Cols
            For i = 0 To .Rows - 1
                For t = 1 To .Cols - 1
                    SaveSettingData SettingName, "Evaluation QC", "Results Grid Standard (" & i & ")  Column " & t, .Cell(i, t).Text
                    'If .Cell(i, t).ForeColor <> 0 Then
                        SaveSettingData SettingName, "Evaluation QC", "Results Grid Standard (" & i & ") Forecolor " & t, .Cell(i, t).ForeColor
                    'End If
                Next
            Next
        End If
    End With
    
    
    CloseSettingDataFile
End Function

Private Function FillGrd3()
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim STDValue As String
Dim STDNum As Integer
Dim rc1 As Boolean
Dim rc2 As Boolean

t = numSelectedStandard


STDNum = CInt(STD(t, 0))
STDValue = STD(t, 1)

  'AndOr = Chr$(177)
  AndOr = Chr$(247)

If STDNum = 0 Then Exit Function
With Grd3
    .ReadOnly = True
    .AutoRedraw = False
    If .Rows <= 1 Then GoTo Aggiungi:
    .Cell(0, 2).Text = "Target Value " & AndOr & " U [" & UNIT_PP & "]"
    For i = 1 To .Rows - 1
    
        If STDNum = Trim(.Cell(i, 5).Text) And STDValue = Trim(.Cell(i, 6).Text) Then
            'c'è già
            .Cell(i, 1).Text = STDValue
            .Cell(i, 2).Text = STD(t, 2) & " " & AndOr & " " & STD(t, 3)
            .Cell(i, 3).Text = Text1(4)
            .Cell(i, 4).Text = Text1(2)
            rc1 = checkMeanValue(Text1(4), t)
            rc2 = checkMeanValue(Text1(2), t)
            .Cell(i, 3).ForeColor = IIf(rc1, vbBlack, vbRed) '  Text1(4).ForeColor
            .Cell(i, 4).ForeColor = IIf(rc2, vbBlack, vbRed)
            strPassed = IIf(rc1 And rc2, "YES", "NO")
            .Cell(i, 7).Text = strPassed
            GoTo fine:
        End If
    Next
Aggiungi:
            ' lo aggiungo
            .AddItem "", False
            .Cell(.Rows - 1, 1).Text = STDValue
            .Cell(.Rows - 1, 2).Text = STD(t, 2) & " " & AndOr & " " & STD(t, 3)
            .Cell(.Rows - 1, 3).Text = Text1(4)
            .Cell(.Rows - 1, 4).Text = Text1(2)
            .Cell(i, 3).ForeColor = Text1(4).ForeColor
            .Cell(i, 4).ForeColor = Text1(2).ForeColor
            .Cell(.Rows - 1, 5).Text = STDNum
            .Cell(.Rows - 1, 6).Text = STDValue
            .Cell(.Rows - 1, 7).Text = strPassed
fine:
    .Column(6).Sort cellAscending
    .AutoRedraw = True
    .Refresh
End With

Call SortGrid(Grd3)


End Function

Private Function checkMeanValue(ByVal mean As Double, ByVal t As Integer) As Boolean
Dim rc As Boolean
rc = True
If CDbl(mean) >= CDbl(STD(t, 2)) And CDbl(mean) <= CDbl(STD(t, 3)) Then
Else
    rc = False
End If
checkMeanValue = rc

End Function

Private Function GetphValue()
Dim i As Integer

    'For i = 1 To pHNumber
    '    pHMin(i) = GetSettingData(SettingName, "Information QC", "pHMin" & i, 0)
    '    pHMax(i) = GetSettingData(SettingName, "Information QC", "pHMax" & i, 0)
    'Next

Dim Num As Integer
Dim sString As String

    Num = 0
    
    With dbTabCode
        For i = 1 To 3
            Num = Num + 1
            sString = GetSettingData(SettingName, "Information QC", "pHValue" & Num, "0")
            
            
            
            If (sString <> "" Or sString <> "/") And IsNumeric(sString) And sString <> "0" Then
                'ph(Num, 0) = sString
    
                pHMin(Num) = GetSettingData(SettingName, "Information QC", "pHMin" & Num, "0")
                pHMax(Num) = GetSettingData(SettingName, "Information QC", "pHMax" & Num, "0")
            Else
                pHMin(Num) = ""
                pHMax(Num) = ""
            End If
                   
        Next
    End With

End Function

Private Function CloseLot()
Dim StartDate As String
Dim rc As Boolean

' chiudiamo il Lotto!!!!

rc = True

    ' salvo gli ultimi cambiamenti.....


    If bHoChiusoilLotto Then
        PopupMessage 2, "Lot already Closed..."
        Exit Function
    End If
    
    CloseSettingDataFile
    
    StartDate = GetSettingData(SettingName, "File Information", "Creation Date", "")
    
    If StartDate = "" Then
        StartDate = FormatDataLAT(date)
        SaveSettingData SettingName, "File Information", "Creation Date", StartDate
    End If
    
    
    StartDate = FormatDateTime(StartDate, vbShortDate)
    CloseSettingDataFile
    
    If Text1(11) = "" Then
        Text1(11) = FormatDataLAT(Now)
    End If

    SaveSelectedTest
    
    SaveResultsTable
    
    
    SaveSettingData SettingName, "Close QC", "Date", FormatDateTime(Now, vbShortDate)
    SaveSettingData SettingName, "Close QC", "Validation Date", Text1(11)
    SaveSettingData SettingName, "Close QC", "Operator", Text1(12)
    
    CloseSettingDataFile
    


On Error GoTo ERR_SAVE:
rc = True
With dbTabReport
    .filter = ""
    .filter = "Lot='" & Trim(Text1(0)) & "' and Code='" & Trim(Text1(1)) & "' and NomeFile='" & SettingName & "'"
    
    If .EOF Then
        ' come è possibile????
        
        'StartDate
        .filter = ""
        .filter = "Lot='" & Trim(Text1(0)) & "' and Code='" & Trim(Text1(1)) & "' and StartDate='" & StartDate & "'"
        If .EOF Then
            PopupMessage 2, "No data available for Lot = " & Trim(Text1(0)) & " Code = " & Trim(Text1(1))
            Exit Function
        End If
        
        
      '  mrc=CheckFileName(
      '  If mrc = False Then Exit Function
    Else
    End If
        !Nomefile = SettingName
        !Finished = IIf(Grd3.Rows > 1, True, False)
        !Evaluation = True
        .Update
End With

ERR_END:
    On Error GoTo 0
    
    If rc Then
    
        If USER_PATH = USER_DATA_PATH Then
            ' non c'è bisogno di spostarlo/cancellarlo----
        Else
            FileCopy USER_PATH & SettingName, USER_DATA_PATH & SettingName
            Kill USER_PATH & SettingName
        End If
        PopupMessage 2, "Lot closed..." & vbCrLf & "Lot : " & Trim(Text1(0)) & "  Code : " & Trim(Text1(1))
      
        USER_PATH = USER_DATA_PATH
    End If
    bHoChiusoilLotto = rc
    CloseLot = rc
    Exit Function
ERR_SAVE:
    MsgBox Err.Description
    rc = False
    Resume ERR_END

End Function


Private Sub OpenReadingQC()
On Error GoTo ERR_OPEN
F_READING.WindowState = Me.WindowState
F_READING.Left = Me.Left
F_READING.Top = Me.Top


If F_READING.Visible = True Then
    Unload Me

Else
    If F_READING.DoShow(IndexFormProcedura, Text1(0), Text1(1), , , SettingName) Then
        
      GetFormSettingName
    
    End If

End If

ERR_END:


Exit Sub

ERR_OPEN:

'MsgBox err.Description & " " & err.NUMBER

Unload Me


Exit Sub


End Sub

Private Function CheckRequestedFiles() As Boolean
Dim rc As Boolean
Dim sString As String
Dim FiledString As String
On Error GoTo ERR_CHECK



    ' campi obbligatori

    rc = True

    FiledString = "Lot Expiration"
    sString = GetSettingData(SettingName, "Information QC", "Text13", 1)
    GoSub CheckValue
    
    
    FiledString = "Preparation Week"
    sString = GetSettingData(SettingName, "Information QC", "Text121", 1)
    GoSub CheckValue
    
     
 
    FiledString = "Prep. Operator"
    sString = GetSettingData(SettingName, "Information QC", "Text122", 1)
    GoSub CheckValue
    
    
    FiledString = "First Day Prod."
    sString = GetSettingData(SettingName, "Information QC", "Text123", 1)
    GoSub CheckValue
    

    FiledString = "Last Day Prod."
    sString = GetSettingData(SettingName, "Information QC", "Text124", 1)
    GoSub CheckValue
    
    
    FiledString = "Machine"
    sString = GetSettingData(SettingName, "Information QC", "Text125", 1)
    GoSub CheckValue
    
             
    




ERR_END:
    On Error GoTo 0
    PicInformation.Left = Picture2.Left
    If Not (rc) Then
        'MessageInfoTime = 2000
        'PopupMessage 2, "This Lot cannot be closed : please fill all Requested fields" & vbCrLf & FiledString & " is missing." & FiledString & "Goto Lot Information QC"
        lbClose.Visible = True
        lbClose.ForeColor = &H40C0&     ' vbred
        PicInformation.Visible = True
        lbClose = "[Close Lot]  Please, fill all the required fields : Lot Expiration , Preparation Week, Prep. Operator, First day Prod. Last Day Prod., Machine"
    Else
        lbClose.Visible = True
        lbClose.ForeColor = vbColorGreen
        lbClose = "[Close Lot] This Lot has all the Information required..."
        If Grd3.Rows > 1 Then
        Else
            lbClose.ForeColor = &H40C0& ' vbred
            lbClose = "[Close Lot]  This Lot needs at least 1 Mean Value"
            
        End If
        
        PicInformation.Visible = False
    End If
    CheckRequestedFiles = rc
    Exit Function
ERR_CHECK:
    MsgBox Err.Description
    rc = False
    Resume ERR_END

CheckValue:
    If sString = "" Then
    rc = False
    GoTo ERR_END:
    Else
        Return
    End If

End Function




Private Sub PicInformation_Click()
MyLot = Text1(0)
MyCode = Text1(1)
bAnotherFormCalled = True
F_INFORMATION.WindowState = Me.WindowState
F_INFORMATION.Left = Me.Left
F_INFORMATION.Top = Me.Top

If F_INFORMATION.DoShow(IndexFormProcedura, Text1(0), Text1(1), , , SettingName) Then
    Form_Initialize

    lbOperator = MyOperatore.Name
  
End If
CheckRequestedFiles
bAnotherFormCalled = False

End Sub


