VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form frmProduction 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Chemical Production"
   ClientHeight    =   12000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19200
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Production.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   12000
   ScaleWidth      =   19200
   Begin VB.PictureBox PBContainerViewport 
      BackColor       =   &H00FFFFFF&
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
      Height          =   9975
      Index           =   0
      Left            =   0
      ScaleHeight     =   9975
      ScaleWidth      =   19245
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   960
      Width           =   19245
      Begin VB.PictureBox PBContainer 
         BackColor       =   &H00E0E0E0&
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
         Height          =   45000
         Left            =   0
         ScaleHeight     =   45000
         ScaleMode       =   0  'User
         ScaleWidth      =   19155
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   -34080
         Width           =   19155
         Begin VB.Frame frInside 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Caption         =   "&H00E0E0E0&"
            Height          =   9015
            Index           =   3
            Left            =   960
            TabIndex        =   113
            Top             =   35000
            Visible         =   0   'False
            Width           =   17055
            Begin VB.ComboBox cmbRevType 
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   6360
               Style           =   2  'Dropdown List
               TabIndex        =   138
               Top             =   5760
               Visible         =   0   'False
               Width           =   2415
            End
            Begin VB.TextBox txRevision 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
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
               Left            =   6360
               TabIndex        =   139
               Top             =   5760
               Width           =   2415
            End
            Begin VB.Frame frExcel2 
               BackColor       =   &H00206020&
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
               Left            =   960
               TabIndex        =   125
               Top             =   6960
               Width           =   3015
               Begin VB.Label Label4 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Export Excel"
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
                  Left            =   0
                  TabIndex        =   126
                  Top             =   120
                  Width           =   3015
               End
               Begin VB.Image Image1 
                  Height          =   480
                  Left            =   120
                  MousePointer    =   99  'Custom
                  OLEDropMode     =   1  'Manual
                  Picture         =   "Production.frx":29F2
                  Top             =   0
                  Width           =   480
               End
            End
            Begin VB.TextBox txRevision 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
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
               Height          =   585
               Index           =   3
               Left            =   2160
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   124
               Top             =   6240
               Width           =   13815
            End
            Begin VB.TextBox txRevision 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
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
               Left            =   11160
               TabIndex        =   123
               Top             =   5760
               Width           =   2415
            End
            Begin VB.TextBox txRevision 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
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
               Left            =   2160
               TabIndex        =   122
               Top             =   5760
               Width           =   2415
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00886010&
               BorderStyle     =   0  'None
               Caption         =   "&H00E0E0E0&"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   4
               Left            =   0
               TabIndex        =   119
               Top             =   0
               Width           =   17055
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Production Notes"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
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
                  Left            =   7605
                  TabIndex        =   121
                  Top             =   75
                  Width           =   1965
               End
               Begin VB.Label Label23 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Left            =   16845
                  TabIndex        =   120
                  Top             =   120
                  Width           =   45
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
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
               Height          =   495
               Index           =   12
               Left            =   12960
               TabIndex        =   117
               Top             =   6960
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Clear form"
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
                  Index           =   12
                  Left            =   0
                  TabIndex        =   118
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00886010&
               BorderStyle     =   0  'None
               Height          =   1335
               Left            =   5880
               TabIndex        =   114
               Top             =   2400
               Width           =   5055
               Begin VB.Label lbChem 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Enter all fileds and Save"
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
                  Left            =   1380
                  TabIndex        =   116
                  Top             =   720
                  Width           =   2340
               End
               Begin VB.Label lbChem 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Empty List..."
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   5
                  Left            =   0
                  TabIndex        =   115
                  Top             =   360
                  Width           =   4995
               End
            End
            Begin FlexCell.Grid Grid4 
               Height          =   4695
               Left            =   0
               TabIndex        =   127
               TabStop         =   0   'False
               Top             =   600
               Width           =   17055
               _ExtentX        =   30083
               _ExtentY        =   8281
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
               BoldFixedCell   =   0   'False
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
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Type"
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
               Index           =   19
               Left            =   5520
               TabIndex        =   140
               Top             =   5760
               Width           =   735
            End
            Begin VB.Label lbFunction 
               BackStyle       =   0  'Transparent
               Height          =   855
               Index           =   5
               Left            =   8400
               TabIndex        =   135
               Top             =   7440
               Width           =   1815
            End
            Begin VB.Label lbFunction 
               BackStyle       =   0  'Transparent
               Height          =   855
               Index           =   4
               Left            =   6840
               TabIndex        =   134
               Top             =   7440
               Width           =   1575
            End
            Begin VB.Label lbRevision 
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
               Index           =   2
               Left            =   240
               TabIndex        =   133
               Top             =   6240
               Width           =   1695
            End
            Begin VB.Label lbRevision 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Operator"
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
               Left            =   10080
               TabIndex        =   132
               Top             =   5760
               Width           =   855
            End
            Begin VB.Label lbRevision 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
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
               Left            =   240
               TabIndex        =   131
               Top             =   5760
               Width           =   1695
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Delete Note"
               ForeColor       =   &H00808080&
               Height          =   255
               Left            =   8640
               TabIndex        =   130
               Top             =   7875
               Width           =   1170
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Save Note"
               ForeColor       =   &H00808080&
               Height          =   255
               Left            =   7080
               TabIndex        =   129
               Top             =   7875
               Width           =   1005
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   5
               Left            =   7440
               Picture         =   "Production.frx":5DD4
               ToolTipText     =   "4000"
               Top             =   7485
               Width           =   240
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   4
               Left            =   9120
               Picture         =   "Production.frx":67D6
               ToolTipText     =   "4000"
               Top             =   7485
               Width           =   240
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Enter/Edit/ Delete  Revision Specifics : Enter all fields and Save or Export to Excel"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   240
               Index           =   2
               Left            =   5265
               TabIndex        =   128
               Top             =   8640
               Width           =   6435
            End
         End
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
            Height          =   10335
            Index           =   0
            Left            =   960
            TabIndex        =   63
            Top             =   240
            Width           =   17175
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
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
               Height          =   495
               Index           =   13
               Left            =   11040
               TabIndex        =   136
               Top             =   8160
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Production Notes"
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
                  Index           =   13
                  Left            =   0
                  TabIndex        =   137
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frExcel 
               BackColor       =   &H00206020&
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
               Left            =   6240
               TabIndex        =   109
               Top             =   8760
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbExcel 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Export Excel"
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
                  Left            =   0
                  TabIndex        =   110
                  Top             =   120
                  Width           =   3015
               End
               Begin VB.Image Image 
                  Height          =   480
                  Left            =   120
                  MousePointer    =   99  'Custom
                  Picture         =   "Production.frx":71D8
                  Top             =   0
                  Width           =   480
               End
            End
            Begin VB.Frame Frame3 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   1
               Left            =   0
               TabIndex        =   87
               Top             =   3360
               Width           =   17175
               Begin VB.Image ImViewRecipes 
                  Height          =   240
                  Index           =   0
                  Left            =   360
                  Picture         =   "Production.frx":A5BA
                  Top             =   120
                  Width           =   240
               End
               Begin VB.Image ImViewRecipes 
                  Height          =   240
                  Index           =   1
                  Left            =   960
                  Picture         =   "Production.frx":AFBC
                  Top             =   120
                  Width           =   240
               End
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hanna Code Production"
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
                  Height          =   345
                  Index           =   1
                  Left            =   2640
                  TabIndex        =   89
                  Top             =   120
                  Width           =   11325
               End
               Begin VB.Label lbExpand 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "+ "
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   20.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00105010&
                  Height          =   480
                  Left            =   16440
                  TabIndex        =   88
                  Top             =   0
                  Width           =   840
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
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
               Index           =   3
               Left            =   14160
               TabIndex        =   85
               Top             =   8160
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Production Hystory table"
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
                  Index           =   3
                  Left            =   0
                  TabIndex        =   86
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame Frame2 
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
               Left            =   6120
               TabIndex        =   83
               Top             =   5160
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
                  Index           =   0
                  Left            =   1920
                  TabIndex        =   84
                  Top             =   555
                  Width           =   1155
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
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
               Left            =   0
               TabIndex        =   81
               Top             =   8760
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Add Hanna Code"
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
                  TabIndex        =   82
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
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
               Index           =   1
               Left            =   3120
               TabIndex        =   79
               Top             =   8760
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Delete Component from Recipe"
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
                  TabIndex        =   80
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00D0D0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   300
               Index           =   5
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   78
               Top             =   2280
               Width           =   13575
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
               Index           =   5
               Left            =   960
               TabIndex        =   76
               Top             =   840
               Width           =   15255
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe for Production | Preparation"
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
                  TabIndex        =   77
                  Top             =   120
                  Width           =   3915
               End
               Begin VB.Line Line8 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00D0D0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   300
               Index           =   4
               Left            =   13440
               Locked          =   -1  'True
               TabIndex        =   75
               Top             =   1920
               Width           =   2535
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00D0D0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   300
               Index           =   3
               Left            =   3960
               Locked          =   -1  'True
               TabIndex        =   74
               Top             =   1920
               Width           =   1695
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00D0D0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   300
               Index           =   2
               Left            =   14280
               Locked          =   -1  'True
               TabIndex        =   73
               Top             =   1560
               Width           =   1695
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00D0D0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   300
               Index           =   1
               Left            =   8400
               Locked          =   -1  'True
               TabIndex        =   72
               Top             =   1560
               Width           =   1695
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00D0D0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   300
               Index           =   0
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   71
               Top             =   1560
               Width           =   3255
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00D0D0C0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   300
               Index           =   6
               Left            =   8400
               Locked          =   -1  'True
               TabIndex        =   70
               Top             =   1920
               Width           =   1695
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
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
               Index           =   6
               Left            =   6240
               TabIndex        =   68
               Top             =   8160
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Product Calssification"
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
                  Index           =   6
                  Left            =   0
                  TabIndex        =   69
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
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
               Index           =   5
               Left            =   3120
               TabIndex        =   66
               Top             =   8160
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Add Production"
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
                  Index           =   5
                  Left            =   0
                  TabIndex        =   67
                  Top             =   120
                  Width           =   3015
               End
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
               Index           =   4
               Left            =   0
               TabIndex        =   64
               Top             =   8160
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Save Production"
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
                  Index           =   4
                  Left            =   0
                  TabIndex        =   65
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin FlexCell.Grid Grid1 
               Height          =   3855
               Left            =   0
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   3960
               Width           =   17175
               _ExtentX        =   30295
               _ExtentY        =   6800
               AllowUserSort   =   -1  'True
               Appearance      =   0
               BackColor1      =   14737632
               BackColor2      =   14737632
               BackColorBkg    =   14737632
               BackColorFixed  =   14737632
               BackColorFixedSel=   14737632
               BackColorScrollBar=   15592423
               BorderColor     =   14737632
               CellBorderColor =   14737632
               CellBorderColorFixed=   14737632
               Cols            =   5
               DefaultFontName =   "Segoe UI"
               DefaultFontSize =   9.75
               BoldFixedCell   =   0   'False
               DisplayRowIndex =   -1  'True
               DrawMode        =   1
               DefaultRowHeight=   20
               FixedRowColStyle=   0
               ForeColorFixed  =   6571523
               GridColor       =   14737632
               Rows            =   1
               ScrollBarStyle  =   0
               SelectionMode   =   3
               MultiSelect     =   0   'False
               DateFormat      =   0
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Select Production Hystory table to view a List of all Productions"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   240
               Index           =   17
               Left            =   12090
               TabIndex        =   108
               Top             =   9000
               Width           =   5025
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Select Hanna Code to Add Production or view Product Classification"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   240
               Index           =   16
               Left            =   11520
               TabIndex        =   107
               Top             =   8760
               Width           =   5565
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Recipe without Preparation : Set # Prep.Week before start Production.   Preparation Date + Preparation Week are Auto filled"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   240
               Index           =   7
               Left            =   960
               TabIndex        =   104
               Top             =   480
               Width           =   9960
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00D0D0D0&
               X1              =   120
               X2              =   17280
               Y1              =   7920
               Y2              =   7920
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00B0B0B0&
               X1              =   960
               X2              =   16200
               Y1              =   2880
               Y2              =   2880
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Note"
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
               Left            =   1200
               TabIndex        =   97
               Top             =   2280
               Width           =   450
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Planning Reference"
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
               Left            =   11280
               TabIndex        =   96
               Top             =   1920
               Width           =   1725
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Planned Preparation Week"
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
               Left            =   1200
               TabIndex        =   95
               Top             =   1920
               Width           =   2385
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "# Prep Week"
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
               Left            =   12720
               TabIndex        =   94
               Top             =   1560
               Width           =   1335
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "Preparation Date"
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
               Left            =   6360
               TabIndex        =   93
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "Recipe by"
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
               Left            =   1200
               TabIndex        =   92
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Preparation Week"
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
               Left            =   6360
               TabIndex        =   91
               Top             =   1920
               Width           =   1620
            End
         End
         Begin VB.TextBox txQRCode 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   840
            TabIndex        =   62
            Text            =   "Text1"
            Top             =   17760
            Visible         =   0   'False
            Width           =   14535
         End
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
            Height          =   6480
            Index           =   1
            Left            =   840
            TabIndex        =   19
            Top             =   11040
            Width           =   17175
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   11
               Left            =   3120
               TabIndex        =   60
               Top             =   5880
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Delete Production"
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
                  Index           =   11
                  Left            =   0
                  TabIndex        =   61
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   10
               Left            =   0
               TabIndex        =   58
               Top             =   5880
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Add Production"
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
                  Index           =   10
                  Left            =   0
                  TabIndex        =   59
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
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
               Index           =   2
               Left            =   14160
               TabIndex        =   24
               Top             =   5880
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hanna Codes Table"
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
                  TabIndex        =   25
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00D0D0D0&
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
               Height          =   540
               Index           =   8
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Width           =   17175
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Production History Table"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   12.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00105010&
                  Height          =   285
                  Index           =   8
                  Left            =   7095
                  TabIndex        =   21
                  Top             =   105
                  Width           =   2955
               End
            End
            Begin FlexCell.Grid Grid2 
               Height          =   4935
               Left            =   0
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   600
               Width           =   17175
               _ExtentX        =   30295
               _ExtentY        =   8705
               AllowUserSort   =   -1  'True
               Appearance      =   0
               BackColor1      =   14737632
               BackColor2      =   14737632
               BackColorActiveCellSel=   14737632
               BackColorBkg    =   14737632
               BackColorFixed  =   14737632
               BackColorFixedSel=   14737632
               BackColorScrollBar=   15592423
               BorderColor     =   14737632
               CellBorderColor =   14737632
               CellBorderColorFixed=   14737632
               Cols            =   5
               DefaultFontName =   "Segoe UI"
               DefaultFontSize =   9.75
               BoldFixedCell   =   0   'False
               DisplayRowIndex =   -1  'True
               DrawMode        =   1
               DefaultRowHeight=   20
               FixedRowColStyle=   0
               ForeColorFixed  =   8937488
               GridColor       =   14737632
               Rows            =   1
               ScrollBarStyle  =   0
               SelectionMode   =   3
               MultiSelect     =   0   'False
               DateFormat      =   0
            End
            Begin VB.Line Line12 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   17280
               Y1              =   5640
               Y2              =   5640
            End
         End
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
            Height          =   8655
            Index           =   2
            Left            =   840
            TabIndex        =   16
            Top             =   18480
            Visible         =   0   'False
            Width           =   17295
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
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
               Index           =   14
               Left            =   13680
               TabIndex        =   111
               Top             =   2040
               Width           =   2415
            End
            Begin VB.Frame frLotMix 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   615
               Left            =   600
               TabIndex        =   99
               Top             =   2040
               Visible         =   0   'False
               Width           =   11535
               Begin VB.TextBox txAcquisition 
                  Alignment       =   2  'Center
                  BackColor       =   &H000000FF&
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
                  Index           =   13
                  Left            =   6600
                  TabIndex        =   102
                  Top             =   120
                  Width           =   2415
               End
               Begin VB.TextBox txAcquisition 
                  Alignment       =   2  'Center
                  BackColor       =   &H000000FF&
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
                  Index           =   12
                  Left            =   1680
                  TabIndex        =   100
                  Top             =   120
                  Width           =   2415
               End
               Begin VB.Label lbAcquisition 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mix 2 Lot"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Index           =   13
                  Left            =   5760
                  TabIndex        =   103
                  Top             =   120
                  Width           =   690
               End
               Begin VB.Label lbAcquisition 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mix 1 Lot"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Index           =   12
                  Left            =   840
                  TabIndex        =   101
                  Top             =   120
                  Width           =   690
               End
            End
            Begin VB.ComboBox ComboMachine 
               BackColor       =   &H00F0F0F0&
               Height          =   375
               Left            =   13680
               Style           =   2  'Dropdown List
               TabIndex        =   98
               Top             =   960
               Visible         =   0   'False
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
               Index           =   11
               Left            =   6720
               TabIndex        =   56
               Top             =   4800
               Width           =   7455
            End
            Begin VB.PictureBox PicTolerance 
               BorderStyle     =   0  'None
               Height          =   135
               Left            =   6720
               ScaleHeight     =   135
               ScaleWidth      =   4095
               TabIndex        =   55
               Top             =   4200
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   7
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   53
               Text            =   "-200,221"
               Top             =   3840
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0F0F0&
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
               Height          =   400
               Index           =   9
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   51
               Text            =   "-21 %"
               Top             =   4800
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0F0F0&
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
               Height          =   400
               Index           =   8
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   49
               Text            =   "-200,221"
               Top             =   4320
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
               Height          =   600
               Index           =   10
               Left            =   6720
               TabIndex        =   47
               Text            =   "1229,998"
               Top             =   3600
               Width           =   4095
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   6
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   45
               Text            =   "1300,400"
               Top             =   3360
               Width           =   2415
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   9
               Left            =   11160
               TabIndex        =   37
               Top             =   5760
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
                  Index           =   9
                  Left            =   0
                  TabIndex        =   38
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00307030&
               BorderStyle     =   0  'None
               Height          =   495
               Index           =   8
               Left            =   8040
               TabIndex        =   35
               Top             =   5760
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
                  Index           =   8
                  Left            =   0
                  TabIndex        =   36
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   7
               Left            =   14280
               TabIndex        =   33
               Top             =   5760
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Exit Production"
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
                  Index           =   7
                  Left            =   0
                  TabIndex        =   34
                  Top             =   120
                  Width           =   3015
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
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   1320
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
               Left            =   7200
               TabIndex        =   31
               Top             =   1320
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
               Index           =   2
               Left            =   13680
               TabIndex        =   30
               Top             =   1320
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
               Left            =   2280
               TabIndex        =   29
               Top             =   1680
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
               Left            =   7200
               TabIndex        =   28
               Top             =   1680
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
               Index           =   5
               Left            =   13680
               TabIndex        =   27
               Top             =   1680
               Width           =   2415
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00886010&
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
               Height          =   540
               Index           =   2
               Left            =   0
               TabIndex        =   17
               Top             =   0
               Width           =   17295
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Production Details"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
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
                  Left            =   7590
                  TabIndex        =   18
                  Top             =   120
                  Width           =   2085
               End
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Exp Date"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   14
               Left            =   12720
               TabIndex        =   112
               Top             =   2040
               Width           =   765
            End
            Begin VB.Label lbAcquisition 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mixes : Fill Mix Lot 1 and 2"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   240
               Index           =   15
               Left            =   240
               TabIndex        =   106
               Top             =   6840
               Width           =   2025
            End
            Begin VB.Label lbAcquisition 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Required Fileds : Lot + Machine "
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   240
               Index           =   30
               Left            =   240
               TabIndex        =   105
               Top             =   6600
               Width           =   2535
            End
            Begin VB.Label lbAcquisition 
               BackStyle       =   0  'Transparent
               Caption         =   "Note"
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
               Index           =   11
               Left            =   6720
               TabIndex        =   57
               Top             =   4560
               Width           =   1095
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qty Produced"
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
               Left            =   1995
               TabIndex        =   54
               Top             =   3840
               Width           =   1230
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
               Index           =   9
               Left            =   1995
               TabIndex        =   52
               Top             =   4800
               Width           =   1230
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
               Index           =   8
               Left            =   1995
               TabIndex        =   50
               Top             =   4320
               Width           =   1230
            End
            Begin VB.Label lbAcquisition 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Production"
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
               Left            =   6720
               TabIndex        =   48
               Top             =   3240
               Width           =   975
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qty to Produce"
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
               Left            =   1755
               TabIndex        =   46
               Top             =   3360
               Width           =   1470
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hanna code"
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
               Left            =   1080
               TabIndex        =   44
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00B0B0B0&
               X1              =   840
               X2              =   16080
               Y1              =   2760
               Y2              =   2760
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date Production"
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
               Left            =   5520
               TabIndex        =   43
               Top             =   1320
               Width           =   1575
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Machine"
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
               Left            =   12480
               TabIndex        =   42
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Week Production"
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
               Left            =   480
               TabIndex        =   41
               Top             =   1680
               Width           =   1695
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Operator"
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
               Left            =   6270
               TabIndex        =   40
               Top             =   1680
               Width           =   825
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Lot "
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
               Left            =   13245
               TabIndex        =   39
               Top             =   1680
               Width           =   330
            End
         End
      End
   End
   Begin VB.PictureBox PicHover 
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
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   675
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   675
      Begin VB.Label imOver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "é"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   175
         TabIndex        =   15
         Top             =   80
         Width           =   330
      End
      Begin VB.Label lblHoverClick 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   570
         Left            =   60
         TabIndex        =   14
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.PictureBox PBFooter 
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
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   4
      Top             =   11040
      Width           =   19215
      Begin VB.Timer TimerBeginForm 
         Interval        =   1
         Left            =   8400
         Top             =   120
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Move Forward"
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
         Height          =   225
         Index           =   12
         Left            =   17745
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   660
         Width           =   1200
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Move Previous"
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
         Height          =   225
         Index           =   11
         Left            =   15345
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   660
         Width           =   1230
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Production"
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
         Height          =   225
         Index           =   7
         Left            =   8940
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   660
         Width           =   1260
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   0
         Left            =   9360
         MousePointer    =   99  'Custom
         Picture         =   "Production.frx":B9BE
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   3
         Left            =   15600
         MousePointer    =   99  'Custom
         Picture         =   "Production.frx":EDA0
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   4
         Left            =   18240
         MousePointer    =   99  'Custom
         Picture         =   "Production.frx":12182
         Top             =   120
         Width           =   480
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   8760
         TabIndex        =   7
         Top             =   -120
         Width           =   1695
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   3
         Left            =   14760
         TabIndex        =   6
         Top             =   -120
         Width           =   2175
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   4
         Left            =   17280
         TabIndex        =   5
         Top             =   -120
         Width           =   1935
      End
   End
   Begin VB.PictureBox PBTitle 
      BackColor       =   &H00644603&
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
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   0
      Top             =   0
      Width           =   19215
      Begin ChemicalMR.ucScrollAdd ucScrollAdd1 
         Left            =   9600
         Top             =   120
         _ExtentX        =   1138
         _ExtentY        =   423
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00A48643&
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
         Height          =   1095
         Index           =   0
         Left            =   0
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   2175
         TabIndex        =   1
         Top             =   0
         Width           =   2175
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   840
            MousePointer    =   99  'Custom
            Picture         =   "Production.frx":15564
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Production"
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
            Height          =   225
            Index           =   0
            Left            =   90
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   640
            Width           =   2070
         End
      End
      Begin VB.Label lbLine 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Production"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   26
         Top             =   195
         Visible         =   0   'False
         Width           =   19215
      End
      Begin VB.Label lbWait 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "Wait : Loading Data..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5760
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   7575
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Production"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   16770
         TabIndex        =   3
         Top             =   195
         Width           =   2160
      End
   End
End
Attribute VB_Name = "frmProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_rc As Boolean



Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type


Private m_ControlPositions() As ControlPositionType
Private m_FormWid As Single
Private m_FormHgt As Single

Private IndexProcedura As Integer
Private IndexDashCommInside As Integer
Private IndexVisibleFrame As Integer

Private SelectedCode As String

Private uRecipe() As RecipeType
Private uHannaCodes() As HannaCode
Private SelectedMixCode As String
Private SelectedRecipeCode As String


Private lRowHanna As Long
Private lColHanna As Long

Private lRowRecipe As Long
Private lColRecipe As Long

Private lRowMixes As Long
Private lColMixes As Long

Private lRowMaterialReq As Long
Private lColMaterialReq As Long

Private lRowCombo As Long

Private IndexRecipe As Integer
Private indexMix As Integer
Private IndexComponent   As Integer


Private STDPreparationWay() As ProdWay
Private uSTDPreparation As RecipeForSTDPreparation
Private uSTDPreparationClean As RecipeForSTDPreparation
Private uMaterialRequisition As MaterialRequisition

Private SettingName As String
Private bImportata As Boolean
Private bIfDataPath As Boolean
Private bfrInsideMoveTop As Boolean

Private bCancelUpdate As Boolean

Private RecipeCode As String
Private HannaCode As String

Private Frame3Top As Long
Private Grid1Height As Long
Private STDPreparationID As Long
Private userAcquisition As ProdAcquisition
Private userAcquisitionClean As ProdAcquisition

Private AcquisitionID As Long
Private AcquisitionHannaCode As String
Private AcquisitionQty As Double
Private lAcquisitionRow As Long
Private lAcquisitionIndex As Integer

Private ComponentID As Long
Private ComponentHannaCode As String
Private HannaCodeQty As Double
Private lhannaCodeRow As Long
Private IndexCode As Integer
Private bNoPreparationRecipe As Boolean
Private bSTDPreparationClosed  As Boolean

Private NotesID As Long

Private Sub SetColumnWidth()

Dim ctl As Control
Dim i As Integer
For Each ctl In Controls
    If TypeOf ctl Is Grid Then
            For i = 1 To ctl.Cols - 1
                ctl.Column(i).Width = (m_ControlGridColWidth / m_ControlGridColWidthOld) * ctl.Column(i).Width
          Next
    End If
Next



m_ControlGridFontSizeOld = m_ControlGridFontSize
m_ControlGridColWidthOld = m_ControlGridColWidth
m_ControlGridRowHeightOld = m_ControlGridRowHeight


 Frame3Top = Frame3(1).Top
 Grid1Height = Grid1.Height

    
End Sub


Public Function DoShow(Optional ByVal FileName As String, Optional ByVal STDPreparation_ID As Long) As Boolean

    On Error GoTo ERR_SHOW
    
    m_rc = False
    mOk
    STDPreparationID = STDPreparation_ID
    bIfDataPath = IIf(USER_PATH = USER_STD_PREPARATION_PATH, True, False)

    SettingName = FileName
        
    If SettingName = "" Then
        '----------------------------------------------
        ' è una STDPreparation senza nessuna preparation!
        '----------------------------------------------
        Call SetSTDPreparationNoPreparation
    End If
    
    bImportata = IIf(FileName <> "", True, False)



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






Private Sub ComboMachine_Click()
If ComboMachine.ListIndex > 0 Then
    txAcquisition(2) = ComboMachine
Else
    txAcquisition(2) = ""
End If
ComboMachine.Visible = False
End Sub

Private Sub Form_Activate()
Me.WindowState = MainWindowState
End Sub







Private Sub Grid1_DblClick()
If bIfDataPath Then AddSTDPreparation
End Sub

Private Sub Grid1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)


    lbLine.Visible = False
    
    IndexCode = 0
    lhannaCodeRow = 0
    frCommandInside(5).Visible = False
    frCommandInside(6).Visible = False
    
    If FirstRow > 0 Then
        'If Grid1.Cell(FirstRow, 7).Text <> "" Then
            lhannaCodeRow = FirstRow
            HannaCode = Trim(Grid1.Cell(FirstRow, 1).Text)
           
           ' HannaCodeQty = Trim(Grid1.Cell(FirstRow, 7).Text)
       ' End If
        If bSTDPreparationClosed = False Then
            frCommandInside(6).Visible = IIf(HannaCode <> "", bIfDataPath, False)
            frCommandInside(5).Visible = IIf(HannaCode <> "", bIfDataPath, False)
            
            
            lbLine.Caption = HannaCode
            lbLine.Visible = True
            
        End If
        
        
    End If
    
End Sub


Private Sub Grid2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
AcquisitionID = 0
AcquisitionHannaCode = ""
AcquisitionQty = 0
lAcquisitionRow = 0
If FirstRow > 0 Then
    lAcquisitionRow = FirstRow
    AcquisitionID = Grid2.Cell(FirstRow, 10).Text
    AcquisitionHannaCode = Grid2.Cell(FirstRow, 1).Text
    AcquisitionQty = CDbl(Grid2.Cell(FirstRow, 2).Text)
    lAcquisitionIndex = Grid2.Cell(FirstRow, 11).Text
    frCommandInside(11).Visible = Not (bSTDPreparationClosed)
    
End If

End Sub

Private Sub Image_Click()
frExcel_Click
End Sub

Private Sub ImViewRecipes_Click(Index As Integer)
    Select Case Index
        Case 0
            ' ripristina tutte le Rows..
            Call ViewHannaCodeInSTDPreparation(uHannaCodes, Grid1, True)
        Case 1
            Call ViewHannaCodeInSTDPreparation(uHannaCodes, Grid1, False)
    
    End Select
End Sub

Private Sub lbExcel_Click()
frExcel_Click
End Sub

Private Sub lbExpand_Click()
Grid1.ZOrder
Frame3(1).ZOrder
    If Trim(lbExpand) = "+" Then
        lbExpand = "-"
        Frame3(1).Top = 0
        Grid1.Top = Frame3(1).Top + Frame3(1).Height + 60
        Grid1.Height = Line5.y1 - Grid1.Top - 120
        
    Else
        lbExpand = "+"
        
        Frame3(1).Top = Frame3Top
        Grid1.Top = Frame3(1).Top + Frame3(1).Height + 60
        Grid1.Height = Grid1Height

    
    End If
End Sub

Private Sub TimerBeginForm_Timer()
    
    
Call StartUpForm

TimerBeginForm.Enabled = False

End Sub

Private Sub StartUpForm()
    
    Call InitForm
    
    Dim i As Integer
    
    
    
    
    
    If bfrInsideMoveTop = False Then
        For i = 3 To frInside.UBound
            frInside(i).Top = frInside(i).Top - (frInside(2).Height) * m_ControlGridRowHeight
        Next
        bfrInsideMoveTop = True
    End If
    

   
    '--------------------------------------
    '
    '   Recipe importata
    '
    '--------------------------------------
  
    If bImportata Then
        GetFileInfo
    Else
        If txFormulation(1) = "" Then
            txFormulation(1) = FormatDataLAT(Now())
        End If
        
        If txFormulation(3) = "" Then
            txFormulation(3) = PreparationWeek(Now())
            
        End If
        If txFormulation(6) = "" Then
            txFormulation(6) = PreparationWeek(Now())
        End If
        
        
    End If
    
    
    '--------------------------------------


End Sub
Private Sub InitForm()



  
    uSTDPreparation = uSTDPreparationClean
    
    ReDim uRecipe(0)
    
    SelectedCode = ""
    SelectedMixCode = ""
    SelectedRecipeCode = ""
    lRowHanna = 0
    lColHanna = 0
    lRowRecipe = 0
    lColRecipe = 0
    lRowMixes = 0
    lColMixes = 0
    lRowCombo = 0
    IndexRecipe = 0
    indexMix = 0
    
    lRowMaterialReq = 0
    lColMaterialReq = 0
    
    
    Dim Grid(10) As Grid
    
    Set Grid(0) = Grid1
    Set Grid(1) = Grid2
    'Set Grid(2) = Grid3
    'Set Grid(3) = Grid4
    'Set Grid(4) = Grid5
   ' Set Grid(6) = Grid7
    
    Call SetGridNotes(Grid4)
    
    Call SetAllSTDPreparationGrid(Grid())
    Call SetColumnWidth
    
    Grid1.FrozenCols = 2
    Grid2.FrozenCols = 2
   ' Grid3.FrozenCols = 2
   ' Grid4.FrozenCols = 2
   ' Grid5.FrozenCols = 2
  
   ' Grid7.FrozenCols = 2
   '
    
    

End Sub
Private Sub Form_Load()


    PBContainer.Top = 0
    PBContainer.Left = 0
    PBContainer.Width = Me.Width
  
    
    

    ucScrollAdd1.AddScroll PBContainerViewport(0)
    ucScrollAdd1.TrackMouseWheel Vertical
    ucScrollAdd1.ResizeTargetOnFormResize 0, 0
    ucScrollAdd1.UCScrollV.ShowButtons = False
    ucScrollAdd1.UCScrollH.ShowButtons = False
    
    
    Dim i As Integer
    If Screen.Width - Me.Width > 1000 And bFullScreen Then
        Me.WindowState = 2
    
    End If


    For i = PBContainerViewport.LBound To PBContainerViewport.UBound
        PBContainerViewport(i).Move 0, PBTitle.Height, Me.ScaleWidth, Me.ScaleHeight - PBTitle.Height
    Next
  
    RSBottom PicHover, Me, -1350
    RSRight PicHover, Me, -450
 

    PBContainerViewport(0).ZOrder
    PBFooter.ZOrder
    
    
    
    
    
End Sub

Private Sub Form_Resize()




    
    lbWait.Left = Me.Width / 2 - lbWait.Width / 2
    PBTitle.Width = Me.Width
    PBFooter.Top = Me.ScaleHeight - PBFooter.Height
    PBFooter.Width = Me.Width
 
    
    'Resize the container (needed to show the full bottom box on maximized state)
    'First resize our container
    ucScrollAdd1.ContainerW = Me.ScaleWidth
    'But also need to resize PBContainer wich hide the width of the bottom box

    
    
      ResizeControls

    SetColumnWidth
    
   ' MainWindowState = Me.WindowState
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmSTDPreparation = Nothing
End Sub


Private Sub txFormulation_Change(Index As Integer)
Dim rc As Boolean
    rc = IIf(Len(txFormulation(Index)) > 0, True, False)

    
    txFormulation(Index).BackColor = IIf(rc, &HD0D0C0, vbRed)
    

    
End Sub


Private Sub txFormulation_Click(Index As Integer)


If Index = 2 Then

With uSTDPreparation

      
        
            
        If .PrepWeek = "" Then .PrepWeek = PreparationWeek(Now())
        
            Dim strPrep As String
            PopupMessage 2, "Please enter Number Preparation Week", , , "STDPreparation"
            If F_InputBox.DoShow("Number Preparation Week", "STDPreparation", , , , strPrep) Then
                .numPrepWeek = strPrep
            End If
     
         
        txFormulation(2) = .numPrepWeek
        

        
    End With



End If


End Sub



Private Sub ucScrollAdd1_ScrollH(Value As Long)
    Form_Resize
End Sub
Private Sub PicHover_Click()
ucScrollAdd1.UCScrollV.ScrollToValue 0
End Sub
Private Sub lblHoverClick_Click()
    ucScrollAdd1.UCScrollV.ScrollToValue 0
    
End Sub
Private Sub imOver_Click()
ucScrollAdd1.UCScrollV.ScrollToValue 0
End Sub

'========================================
'Vertical scroll event
'========================================
Private Sub ucScrollAdd1_ScrollV(Value As Long)
    
    'Just log the value for no reason
   
        
    PicHover.ZOrder
    'Show a button to scroll to top
    If Not (ucScrollAdd1.UCScrollV Is Nothing) Then
        If (ucScrollAdd1.UCScrollV.Value > 100) Then
            PicHover.Visible = True
        Else
            PicHover.Visible = False
        End If
    Else
        PicHover.Visible = False
    End If
    
    
    If ucScrollAdd1.UCScrollV.Value <= frInside(0).Top Then
        IndexVisibleFrame = 0
    ElseIf ucScrollAdd1.UCScrollV.Value > frInside(0).Top And ucScrollAdd1.UCScrollV.Value <= frInside(1).Top Then
        IndexVisibleFrame = 1
    
    ElseIf ucScrollAdd1.UCScrollV.Value > frInside(1).Top And ucScrollAdd1.UCScrollV.Value <= frInside(2).Top Then
        IndexVisibleFrame = 2
    ElseIf ucScrollAdd1.UCScrollV.Value > frInside(2).Top And ucScrollAdd1.UCScrollV.Value <= frInside(3).Top Then
        IndexVisibleFrame = 3
   ' ElseIf ucScrollAdd1.UCScrollV.Value > frInside(3).Top And ucScrollAdd1.UCScrollV.Value <= frInside(4).Top Then
       ' IndexVisibleFrame = 4
    End If
              
        
   
    
End Sub

'Poorly made resizing functions just for the example
Private Sub RSRight(c As Control, Source As Object, adjust As Long, Optional LimitLeft& = -1, Optional LimitRight& = -1)
On Error Resume Next
Dim aux&
    aux& = (Source.ScaleWidth - c.Width) + adjust
    If (err.NUMBER > 0) Then aux& = (Source.Width - c.Width) + adjust
    If (aux < LimitLeft) And (LimitLeft <> -1) Then aux = LimitLeft
    If (aux > LimitRight&) And (LimitRight& <> -1) Then aux = LimitRight&
    c.Left = aux
End Sub

Private Sub RSWidth(c As Control, Source As Object, adjust As Long, Optional LimitLeft& = 0, Optional LimitRight& = -1)
Dim aux&
    aux& = Source.Width + adjust
    If (aux < LimitLeft) Then aux = LimitLeft
    If (aux > LimitRight&) And (LimitRight& <> -1) Then aux = LimitRight&
    c.Width = aux
End Sub

Private Sub RSCenter(c As Control, Source As Object, Optional adjust As Long = 0, Optional LimitLeft& = -1, Optional LimitRight& = -1)
Dim aux&
    aux& = ((Source.Width / 2) - (c.Width / 2)) + adjust
    If (aux < LimitLeft) And (LimitLeft <> -1) Then aux = LimitLeft
    If (aux > LimitRight&) And (LimitRight& <> -1) Then aux = LimitRight&
    c.Left = aux
End Sub

Private Sub RSBottom(c As Control, Source As Object, adjust As Long, Optional LimitBot& = -1)
On Error Resume Next
Dim aux&
    aux& = (Source.ScaleHeight - c.Height) + adjust
    If (err.NUMBER > 0) Then aux& = (Source.Height - c.Height) + adjust
    If (aux < LimitBot) And (LimitBot <> -1) Then aux = LimitBot
    c.Top = aux
End Sub

Private Sub RSLeft(c As Control, Source As Object, adjust As Long, Optional LimitLeft& = -1, Optional LimitRight& = -1)
Dim aux&
    aux& = Source.Left + adjust
    If (aux < LimitLeft) And (LimitLeft <> -1) Then
        aux = LimitLeft
    ElseIf (aux > LimitRight&) And (LimitRight& <> -1) Then
        aux = LimitRight&
    End If
    c.Left = aux
End Sub



Private Sub SaveSizes()
Dim i As Integer
Dim ctl As Control
' Save the controls' positions and sizes.
On Error GoTo ERR_SAVE

m_ControlGridFontSizeOld = 1
m_ControlGridColWidthOld = 1
m_ControlGridRowHeightOld = 1

ReDim m_ControlPositions(1 To Controls.Count)
i = 1
For Each ctl In Controls
    With m_ControlPositions(i)
        If TypeOf ctl Is Line Then
            .Left = ctl.x1
            .Top = ctl.y1
            .Width = ctl.x2 - ctl.x1
            .Height = ctl.y2 - ctl.y1
        ElseIf TypeOf ctl Is Menu Then
        ElseIf TypeOf ctl Is Inet Then
        ElseIf TypeOf ctl Is Timer Then
        ElseIf TypeOf ctl Is ucScrollAdd Then

        Else
            .Left = ctl.Left
           ' MsgBox (TypeName(ctl))
            .Top = ctl.Top
            .Width = ctl.Width
            .Height = ctl.Height
            On Error Resume Next
            .FontSize = ctl.Font.Size
            
            'MsgBox (TypeName(ctl))
            On Error GoTo 0
        End If
    End With
    i = i + 1
Next ctl
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
Dim ctl As Control
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

'If Not (bStazioneEsterna) Then
'm_ControlGridFontSize = 1 ' y_scale * 0.8
'm_ControlGridColWidth = 1 ' x_scale

'End If
'm_ControlGridRowHeight = 1 '1.4


For Each ctl In Controls
    With m_ControlPositions(i)
        If TypeOf ctl Is Line Then
            ctl.x1 = x_scale * .Left
            ctl.y1 = y_scale * .Top
            ctl.x2 = ctl.x1 + x_scale * .Width
            ctl.y2 = ctl.y1 + y_scale * .Height
        ElseIf TypeOf ctl Is Timer Then
        ElseIf TypeOf ctl Is Inet Then
        ElseIf TypeOf ctl Is Image Then
            ctl.Left = (x_scale * .Left) + IIf(x_scale = 1, 0, (x_scale - 1) * 200)
            ctl.Top = y_scale * .Top
        ElseIf TypeOf ctl Is ucScrollAdd Then
        ElseIf TypeOf ctl Is Grid Then
           ctl.Left = x_scale * .Left
            ctl.Top = y_scale * .Top
            ctl.Width = x_scale * .Width
            ctl.Height = y_scale * .Height

               ' ctl.DefaultFont.Size = 12 * m_ControlGridFontSize
               ' ctl.DefaultRowHeight = 30 * m_ControlGridRowHeight
           
        Else
            ctl.Left = x_scale * .Left
           ' MsgBox (TypeName(Ctl))
            ctl.Top = y_scale * .Top
            ctl.Width = x_scale * .Width
            If Not (TypeOf ctl Is ComboBox) Then
                ' Cannot change height of ComboBoxes.
                ctl.Height = y_scale * .Height
            End If
            On Error Resume Next
            ctl.Font.Size = y_scale * .FontSize
            On Error GoTo 0
        End If
    End With
    i = i + 1
Next ctl
Exit Sub
ERR_SAVE:
Resume Next
End Sub

Private Sub Form_Initialize()

SaveSizes
End Sub






Private Sub DefaultMenuLabel_Click(Index As Integer)
DefaultMenu_Click Index
End Sub



Private Sub DefaultMenu_Click(Index As Integer)
Dim MyIndex As Integer
Select Case Index
    Case 0
        If F_MsgBox.DoShow("Quit STDPreparation?") Then
            
            If Grid2.Rows > 1 Then
                If bIfDataPath Then
                    If F_MsgBox.DoShow("Save STDPreparation?") Then
                        frCommandInside_Click 4
                    Else
                    End If
                End If
            
            End If
            
            Unload Me
        End If
        
    Case 3
        ' Previous
         If IndexVisibleFrame > 1 Then
            MyIndex = IndexVisibleFrame - 1
            
            
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(MyIndex).Top - 680
        Else
            ucScrollAdd1.UCScrollV.ScrollToValue 0
         End If
    
    
    
    Case 4
        ' forward
        If IndexVisibleFrame < frInside.UBound Then
            MyIndex = IndexVisibleFrame + 1
            
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(MyIndex).Top - 680
        Else
            ucScrollAdd1.UCScrollV.ScrollToValue 0
        End If
          
End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27
        Unload Me
    Case 37
         DefaultMenu_Click 3
    Case 39
        DefaultMenu_Click 4
    Case 38
        DefaultMenu_Click 3
    Case 40
        DefaultMenu_Click 4
        
End Select
End Sub

Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To PicMenu.Count - 1
 
    If i = IndexProcedura Then
        ' lascialo stare...
    ElseIf i = Index Then
        PicMenu(i).BackColor = &H886010
    Else
        PicMenu(i).BackColor = &H644603
    End If
Next
End Sub

Private Sub PicMenu_Click(Index As Integer)

If IndexProcedura = Index Then
Else
    Call SelectProcedura(Index)
End If
End Sub


Private Function SelectProcedura(ByVal Index As Integer) As Boolean
Dim rc As Boolean
Dim i As Integer

If Index > PicMenu.UBound Then Exit Function


For i = 0 To PicMenu.UBound
    If i = Index Then
        PicMenu(i).BackColor = &HA48643
    Else
        PicMenu(i).BackColor = &H644603
    End If
Next
blTable = Label2(Index)
IndexProcedura = Index

PBContainerViewport(Index).ZOrder
PBContainerViewport(Index).Visible = True

Select Case IndexProcedura
    Case 0
        
    Case 1
 
End Select

PBFooter.ZOrder


End Function



Private Sub Label2_Click(Index As Integer)
PicMenu_Click Index
End Sub


Private Sub PicMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
PicMenu_Click Index
End Sub
Private Sub Image3_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub frCommandInside_Click(Index As Integer)
Dim rc As Boolean
m_rc = True
Dim UserHannaCode As String

    txQRCode.Visible = False


    bCancelUpdate = False
        
    Select Case Index
    Case 0
        ' add Hanna Code
        Call SelectCode
    Case 2
        ucScrollAdd1.UCScrollV.ScrollToValue 0
    Case 3
        ' STDPreparation hystory table
        frInside(1).ZOrder
        ucScrollAdd1.UCScrollV.ScrollToValue frInside(1).Top - 680
    Case 4
        ' save STDPreparation
        Call SaveSTDPreparation
    Case 5, 10
        ' add STDPreparation
        Call AddSTDPreparation
    Case 6
        ' open product classification
        Call OpenProductCalssification(HannaCode, 0)
    Case 7
        ' exit STDPreparation
        Call SetSTDPreparation(False)
    Case 8
        ' save STDPreparation
        Call SetSTDPreparation(True)
    Case 9
        ' Acquisition : open product classification
        Call OpenProductCalssification(txAcquisition(0), 0)
        
    Case 11
            ' delete acquisition
            Call DeleteAcquisition
            
    Case 12
            Call ClearRevisionForm
    Case 13
        If SettingName <> "" Then
            AddcmbRevType
            lbInside(3).ForeColor = vbWhite
            Call GetSTDPreparationNotes(Grid4, SettingName)
             frExcel2.Visible = IIf(Grid4.Rows > 1, True, False)
            Frame6.Visible = IIf(Grid4.Rows > 1, False, True)
        
            Call ClearRevisionForm
            frInside(3).Visible = True
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(3).Top - 680
        Else
            PopupMessage 3, "Please save STDPreparation first..."
        End If
            
            
    End Select
    
End Sub


Private Sub SelectCode()
Dim HannaCode As String
Dim rc As Boolean
    FormCodes.ZOrder
    rc = FormCodes.DoShow(HannaCode) ', Me.Top
    
    If HannaCode = "" Then Exit Sub
    
    If bImportata Then bImportata = False
    
    Call AddHannaCodeInSTDPreparation(HannaCode)
   
    
    

End Sub

Private Sub AddHannaCodeInSTDPreparation(ByVal HannaCode As String)
Dim i As Integer
Dim rc As Boolean
Dim VarRecipeCode() As String



On Error GoTo ERR_ADD:

   ' uRecipe = uRecipe
    

   
    Call AddCodeInSTDPreparationGrid(Grid1, HannaCode, uHannaCodes)
    
    With uSTDPreparation
        .HannaCodes = uHannaCodes
        .HannaCodesCount = UBound(uHannaCodes)
    End With
    
    '-------------------------------------------
    ' carico la Gird1
    '-------------------------------------------
    Call FillGridSTDPreparationFromFile(Grid1, uSTDPreparation, 1)
    
    
    If SettingName = "" Then
        USER_PATH = USER_STD_PREPARATION_PATH
        SetSettingName
        MessageInfoTime = 3000
        PopupMessage 2, "(1) Check and Fill Recipe for STDPreparation Details" & vbCrLf & "(2) SAVE STDPreparation!! " & vbCrLf & "(3) Add STDPreparation acquisition."
    End If
    
ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_ADD:
    MsgBox "Error Recipe  :   " & uRecipe(i).Code & vbCrLf & err.Description
    Resume Next
End Sub



Private Function SetSTDPreparation(ByVal rc As Boolean)
Dim mrc As Boolean



    If rc Then
      
        If bNoPreparationRecipe Then
        
            If txAcquisition(12) = "" Or txAcquisition(13) = "" Then
                PopupMessage 2, "Please enter Mix 1& 2 Lot...", , True
                Exit Function
            End If
        End If
      
        If txAcquisition(5) = "" Then
            ' manufacturer lot
            txAcquisition(5).SetFocus
            PopupMessage 2, "Please enter Lot..", , True, txAcquisition(5)
            Exit Function
        End If
        
        mrc = SaveAcquisition
        If mrc Then
            PopupMessage 2, "Acquisition done..."
        Else
            PopupMessage 2, "Warning : Acquisition not Saved...", , True, txAcquisition(0)
        End If
    Else
       '
    End If
    ucScrollAdd1.UCScrollV.ScrollToValue 0
    frInside(2).Visible = False
    
End Function

Private Function AddSTDPreparation()

    If STDPreparationID = 0 Then
        MessageInfoTime = 2500
        PopupMessage 3, "Please Save Info STDPreparation data...", , True, "STDPreparation"
        Exit Function
    End If
    ClearSTDPreparation
    frInside(2).ZOrder
    frInside(2).Visible = True
    txAcquisition(0).SetFocus
    ucScrollAdd1.UCScrollV.ScrollToValue frInside(2).Top - 680
    Call FillUserHannaCode(HannaCode, False)
End Function
Private Sub txAcquisition_Change(Index As Integer)
Dim rc As Boolean
    rc = IIf(Len(txAcquisition(Index)) > 0, True, False)
    
    txAcquisition(Index).BackColor = IIf(rc, vbWhite, &HF0F0F0)
    
    Select Case Index
        Case 0
            frCommandInside(9).Visible = rc
            
            If rc Then
                txAcquisition(5).BackColor = IIf(txAcquisition(5) <> "", vbWhite, &HF0F0F0)
            End If
          
        Case 10
            ' peso acquisito

            PicTolerance.Visible = False
            
            If rc Then
               
                txAcquisition(11).SetFocus
            End If

            If Len(txAcquisition(0)) > 0 Then
                frCommandInside(8).Visible = rc
            Else
                 frCommandInside(8).Visible = False
            End If
            
        Case 12, 13
            
 
    End Select
End Sub

Private Sub CheckTxAcquisition()
Dim Index As Integer
Dim rc As Boolean
For Index = txAcquisition.LBound To 5
    rc = IIf(Len(Trim(txAcquisition(Index))) > 0, True, False)
    
    txAcquisition(Index).BackColor = IIf(rc, vbWhite, vbRed)
Next
    
    rc = IIf(Len(Trim(txAcquisition(14))) > 0, True, False)
    txAcquisition(14).BackColor = IIf(rc, vbWhite, vbRed)
End Sub

Private Function ClearSTDPreparation()
Dim i As Integer
    
    For i = txAcquisition.LBound To txAcquisition.UBound
        txAcquisition(i) = ""
    Next
    
    PicTolerance.Visible = False
                           
    Frame3(2).BackColor = &H886010
    lbInside(2) = "STDPreparation Details"
    lbInside(2).ForeColor = vbWhite
    
    userAcquisition = userAcquisitionClean
    AcquisitionID = 0
    AcquisitionHannaCode = ""
    AcquisitionQty = 0
    frCommandInside(11).Visible = False
End Function


Private Sub frInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


Dim i As Integer
    For i = 0 To frCommandInside.UBound

            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 4 Or i = 8 Then
                frCommandInside(i).BackColor = &H8000&
            End If

    
    Next
 
 
End Sub

Private Sub frCommandInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
IndexDashCommInside = Index
Dim i As Integer
    For i = 0 To frCommandInside.UBound
        If i = Index Then
            ' quando ci passo sopra....
            frCommandInside(i).BackColor = &H846623
            lbCommandInside(i).ForeColor = vbWhite
            If i = 4 Or i = 8 Then
                frCommandInside(i).BackColor = &H20A020
            End If
            
        Else
            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 4 Or i = 8 Then
                frCommandInside(i).BackColor = &H8000&
            End If
        End If
    
    Next
 
 
End Sub
Private Sub lbCommandInside_Click(Index As Integer)

frCommandInside_Click Index
End Sub
Private Sub lbCommandInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
frCommandInside_MouseMove Index, Button, Shift, X, Y
End Sub
Private Sub PBTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub PBTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub PBTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
FrmMove = False
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FrmMove = True
    DragX = X
    DragY = Y
    If Me.WindowState = 2 Then
        FrmMove = False
       
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nx, ny
    If Me.WindowState = 2 Then
        FrmMove = False
        Exit Sub
    End If
    nx = Me.Left + X - DragX
    ny = Me.Top + Y - DragY
    Me.Left = nx
    Me.Top = ny
    FrmMove = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer



If Me.WindowState = 2 Then
    FrmMove = False
End If
Dim nx, ny
    If FrmMove Then
        nx = Me.Left + X - DragX
        ny = Me.Top + Y - DragY
        Me.Left = nx
        Me.Top = ny
    End If
    
For i = 0 To PicMenu.UBound
    If i = IndexProcedura Then
    Else
        PicMenu(i).BackColor = &H644603
    End If
Next

End Sub


'-----------------------------------------------------------------------------------------------
'
'
'                                   GetSTDPreparationFromFile

'
'-----------------------------------------------------------------------------------------------


Private Sub GetFileInfo()
Dim rc As Boolean
 rc = GetSTDPreparationFromFile
 
 bImportata = rc
 
End Sub

Private Function GetSTDPreparationFromFile() As Boolean
Dim i As Integer
Dim rc As Boolean

On Error GoTo ERR_GET:
   
rc = True



        
    lbWait.Visible = True
    
    ReDim uRecipe(0)
    ReDim uHannaCodes(0)
    Debug.Print USER_PATH
    
    uSTDPreparation = uSTDPreparationClean
    


    
        If FileExists(USER_STD_PREPARATION_PATH & SettingName) Then
            USER_PATH = USER_STD_PREPARATION_PATH
        ElseIf FileExists(USER_STD_PREPARATION_PATH & "data\" & SettingName) Then
            USER_PATH = USER_STD_PREPARATION_PATH & "data\"
        
        Else
            rc = False
            PopupMessage 2, "No File STDPreparation found!!", , True, SettingName
            Exit Function
            
        End If
        
    If FileExists(USER_PATH & SettingName) = False Or SettingName = "" Then
        MessageInfoTime = 2000
        PopupMessage 2, "Warning : File non found! ", , True, "STDPreparation"
        rc = False
        GoTo ERR_END
    End If

    
    If bIfDataPath Then

            bSTDPreparationClosed = False
        Else
            
            
            frCommandInside(0).Visible = False
            frCommandInside(1).Visible = False
            frCommandInside(4).Visible = False
            frCommandInside(5).Visible = False
            frCommandInside(10).Visible = False
            bSTDPreparationClosed = True
            blTable.Visible = False
            blTable = "STDPreparation : Closed"
            lbLine = "Date : " & GetSTDPreparationDate(STDPreparationID)
            
            frExcel.Move frCommandInside(4).Left, frCommandInside(4).Top, frCommandInside(4).Width, frCommandInside(4).Height
            frExcel.Visible = True
    
    End If
    
    
    Call STDPreparationGetSetting(uSTDPreparation, SettingName)
     
    
    With uSTDPreparation

        uRecipe = .Recipes
        uHannaCodes = .HannaCodes

        If .PreparationDate = "" Then
            .PreparationDate = FormatDataLAT(Now())
            
        End If

        
        'If .ExpDate = "" Then .ExpDate = SetExpDate(.PreparationDate, .Recipes(1).Exp)
       ' .Recipes(1).ExpDate = .ExpDate
            
        If .PrepWeek = "" Then .PrepWeek = PreparationWeek(Now())
        If .numPrepWeek = "" Then
            Dim strPrep As String
            PopupMessage 2, "Please enter Number Preparation Week", , , "STDPreparation"
            If F_InputBox.DoShow("Number Preparation Week", "STDPreparation", , , , strPrep) Then
                .numPrepWeek = strPrep
            End If
        End If
         
        txFormulation(1) = .PreparationDate
        txFormulation(2) = .numPrepWeek
        txFormulation(6) = .PrepWeek
       ' txFormulation(7) = .ExpDate
        
        
        
        
        txFormulation(5) = .Note
        txFormulation(3) = .PlannedPrepWeek
        txFormulation(4) = .PlanningReference
        
        txFormulation(0) = .RecipeBy
        
        

        
       Call SetComboMachine(.HannaCodes(1).Line, ComboMachine)
    End With
    
    
    Call FillGridSTDPreparationFromFile(Grid1, uSTDPreparation, 1)
    Call FillGridSTDPreparationFromFile(Grid2, uSTDPreparation, 2)
   ' Call FillGridSTDPreparationFromFile(Grid3, uSTDPreparation, 3)

    
    
    
    
    blTable.Visible = True
    lbLine.Visible = Not (bIfDataPath)
ERR_END:

    On Error GoTo 0
     uRecipe = uSTDPreparation.Recipes
    lbWait.Visible = False
    GetSTDPreparationFromFile = rc
    Exit Function
ERR_GET:
    rc = False
    Resume Next
    
End Function

Private Function FillUserHannaCode(ByVal Code As String, ByVal bValue As Boolean)
Dim i As Integer
Dim Difference As Double
Dim DiffPerc As Double
    
    
        For i = LBound(uHannaCodes) To UBound(uHannaCodes)
            With uHannaCodes(i)
                If LCase(.Code) = LCase(Code) Then
                
                    If IfNoPreparationRecipe(.Recipe) Then
                        bNoPreparationRecipe = True
                        frLotMix.Visible = True
                        
                        Call SetMix1eMix2(Grid1.Cell(lhannaCodeRow, 11).Text)
                        
                    Else
                        bNoPreparationRecipe = False
                    End If
                     
                    txAcquisition(0) = .Code
                    txAcquisition(2) = .Machine
                    txAcquisition(5) = .LotNumber
                    
                    If .ExpDate = "" Then .ExpDate = SetExpDate(txFormulation(1), GetRecipeExp(.Recipe))
                    
                    
                    
                    
                    txAcquisition(14) = .ExpDate
                    
                    
                    
                    txAcquisition(4) = MyOperatore.Name
                    txAcquisition(1) = FormatDataLAT(Now())
                    txAcquisition(3) = PreparationWeek(Now())
                    txAcquisition(6) = .QtyToProduce
                    txAcquisition(7) = .QtyProduced
                    If .QtyProduced = "" Then .QtyProduced = 0
                    If .QtyToProduce = "" Then .QtyToProduce = 0
                    Difference = CDbl(.QtyProduced) - CDbl(.QtyToProduce)
                    txAcquisition(8) = Int(Difference)
                    If CDbl(.QtyToProduce) > 0 Then
                        txAcquisition(9) = FormatNumber((Difference) / .QtyToProduce, 2) * 100 & " %"
                    Else
                        txAcquisition(9) = "0 %"
                    End If
                    
                    Exit For
                    
                    
                    
                   '
                   
                End If
            End With
        Next


End Function
Private Sub txAcquisition_Click(Index As Integer)
Dim userCode As String
Dim Answer As String
Dim Selected As String
Dim bNumber As Boolean
Dim sString As String
Dim rc As Boolean


    Selected = lbAcquisition(Index) ' "STDPreparation"
    Answer = txAcquisition(Index)
    sString = "Please Enter STDPreparation Detail"
    ComboMachine.Visible = False

    Select Case Index
        Case 0
            ' importa RMCode
            
            'OpenComponentDatabase True
            Exit Sub
        Case 2
            '
            With ComboMachine
                .Move txAcquisition(2).Left, txAcquisition(2).Top, txAcquisition(2).Width
                .ZOrder
                .Visible = True
                .SetFocus
            End With
            Exit Sub
        Case 4

    
            'If frmLogin.DoShow Then
                'txFormulation(Index) = MyOperatore.Name
                'Exit Sub
            'Else
               ' Exit Sub
          ' End If
            

        
        Case 10
             Call CheckTxAcquisition

    End Select
    
    
    If txAcquisition(Index).Locked Then Exit Sub
    If txAcquisition(0) = "" Then Exit Sub
    
    If F_InputBox.DoShow(sString, Selected, , , , Answer, , bNumber, Me.Top) Then
    
        txAcquisition(Index) = Answer
        
        Select Case Index
            Case 1
                ' isdate?
                If IsDate(Answer) Then
                     txAcquisition(Index) = FormatDataLAT(Answer)
                Else
                    PopupMessage 2, "Please enter a valid Date...", , True
                End If
            Case 12, 13
                If IsDate(Answer) Then
                     txAcquisition(Index) = FormatDataLAT(Answer)
                Else
                   ' PopupMessage 2, "Please enter a valid Date...", , True
                End If
            Case 14
                If IsDate(Answer) Then
                     txAcquisition(Index) = FormatDataExp(Answer)
                Else
                    PopupMessage 2, "Please enter a valid Exp Date ( MM/YYYY ) ...", , True
                End If
                    
        End Select
    End If
    
    
    
    
End Sub


Private Function SaveAcquisition() As Boolean
Dim rc As Boolean
    rc = True
    
    '-------------------------------------------
    ' salvo l'acquisizione in userAcquisition
    ' se non avevo il componente lo aggiungo
    ' in uHannaCodes(lhannaCodeRow)
    '-------------------------------------------
    
    Call SetNewUserAcquisition
    
    '-------------------------------------------
    ' salvo l'acquisizione in TabProdHistory
    '-------------------------------------------
    
    Call SaveAcquisitionInTabProdHistory
    '-------------------------------------------
    ' Aggiungo Row in Grid2 ' acquisitions
    '-------------------------------------------
    With Grid2
        .AutoRedraw = False
        
        Call STDPreparationAddNewRowInAcquisition(Grid2, userAcquisition)
        
        .Refresh
        .AutoRedraw = True
    End With
    '-------------------------------------------
    ' Salvo su file :
    ' cancello e risalvo? o aggiungo e basta?
    '
    '-------------------------------------------
    Dim MaxCount As Integer
    MaxCount = uHannaCodes(lhannaCodeRow).AcquisitionCount
    ReDim Preserve uHannaCodes(lhannaCodeRow).Acquisitions(MaxCount)
    
    uHannaCodes(lhannaCodeRow).Acquisitions(MaxCount) = userAcquisition
    
    '-------------------------------------------
    ' ricarico la Gird1
    '-------------------------------------------
    Call FillGridSTDPreparationFromFile(Grid1, uSTDPreparation, 1)
    
    uHannaCodes(lhannaCodeRow).QtyProduced = uSTDPreparation.HannaCodes(lhannaCodeRow).QtyProduced
    
    SaveAcquisition = rc
End Function

Private Sub SetNewUserAcquisition()
Dim AcquisitionsCount As Integer

    With userAcquisition
        .Code = txAcquisition(0)
        .AcquisitionTime = Now()
        .bDeleted = False
        .DateProd = txAcquisition(1)
        .LotNumber = txAcquisition(5)
        .Machine = txAcquisition(2)
        .WeekProd = txAcquisition(3)
        
        .Operator = txAcquisition(4)
        .Note = txAcquisition(11)
        .QtyProduced = txAcquisition(10)
        
        .Mix1Lot = txAcquisition(12)
        .Mix2Lot = txAcquisition(13)
        .ExpDate = txAcquisition(14)
        
        uHannaCodes(lhannaCodeRow).ExpDate = .ExpDate
        uHannaCodes(lhannaCodeRow).Machine = .Machine
        uHannaCodes(lhannaCodeRow).LotNumber = .LotNumber
       
        
    End With
    
    
    
    
    AcquisitionsCount = uHannaCodes(lhannaCodeRow).AcquisitionCount
    AcquisitionsCount = AcquisitionsCount + 1
    uHannaCodes(lhannaCodeRow).AcquisitionCount = AcquisitionsCount
    
    ReDim Preserve uHannaCodes(lhannaCodeRow).Acquisitions(AcquisitionsCount)
    
    uHannaCodes(lhannaCodeRow).Acquisitions(AcquisitionsCount) = userAcquisition
    
    If uHannaCodes(lhannaCodeRow).QtyProduced = "" Then uHannaCodes(lhannaCodeRow).QtyProduced = "0"
    
    uHannaCodes(lhannaCodeRow).QtyProduced = CDbl(uHannaCodes(lhannaCodeRow).QtyProduced) + CDbl(userAcquisition.QtyProduced)
    userAcquisition.Index = AcquisitionsCount
    
    
    uSTDPreparation.HannaCodes(lhannaCodeRow) = uHannaCodes(lhannaCodeRow)
    
    uSTDPreparation.WeekProd = userAcquisition.WeekProd
    
End Sub
Private Function SaveAcquisitionInTabProdHistory()

With dbTabProdHistory
    .AddNew
    !AcquisitionTime = userAcquisition.AcquisitionTime
    !Code = userAcquisition.Code
    !Index = userAcquisition.Index
    
    !DateProd = userAcquisition.DateProd
    !LotNumber = userAcquisition.LotNumber
    !Machine = userAcquisition.Machine

    !QtyProduced = userAcquisition.QtyProduced
    !Note = userAcquisition.Note
    !Operator = userAcquisition.Operator
    !WeekProd = userAcquisition.WeekProd
    !FileName = SettingName
    !STDPreparationID = STDPreparationID
    !Mix1Lot = userAcquisition.Mix1Lot
    !Mix2Lot = userAcquisition.Mix2Lot
    !ExpDate = userAcquisition.ExpDate
    .Update
    
    userAcquisition.ID = !ID

End With

End Function


Public Function GetIndexHannaCodeFromFile(ByVal FileName As String, ByVal Code As String, ByVal Path As String) As Integer
Dim Index As Integer
Dim i As Integer
Dim fileCode As String

Index = 0

    CloseSettingDataFile
    
    If FileName = "" Then Exit Function
    If Code = "" Then Exit Function
    If Path = "" Then Path = USER_STD_PREPARATION_PATH
    
    For i = LBound(uHannaCodes) To UBound(uHannaCodes)
        
        fileCode = GetSettingData(FileName, "HannaCode" & i, "Code", "", Path)
        If fileCode = Code Then
            GetIndexHannaCodeFromFile = i
            Exit Function
        End If
    
    Next
    GetIndexHannaCodeFromFile = 0
    CloseSettingDataFile

End Function



Private Function DeleteAcquisition()
Dim IndexCode As Integer

Dim i As Integer

If AcquisitionID > 0 And AcquisitionHannaCode <> "" Then

     
    If F_MsgBox.DoShow("Delete Acquisition Code : " & AcquisitionHannaCode & vbCrLf & "Qty Produced : " & Trim(PadString(AcquisitionQty)), RecipeCode, True) Then
        IndexCode = GetIndexHannaCodeFromFile(SettingName, AcquisitionHannaCode, USER_STD_PREPARATION_PATH)
    Else
        Exit Function
    End If
    '-----------------------------------
    ' sottraggo il peso inserito...
    '-----------------------------------
    With uHannaCodes(IndexCode)
        ' controllo che IndexCode sia corretto....
        If .Code <> AcquisitionHannaCode Then
            Exit Function
        End If

        .Acquisitions(lAcquisitionIndex).bDeleted = True
         .QtyProduced = .QtyProduced - AcquisitionQty
          
    End With
    

cont:
    '-----------------------------------
    ' cancello la riga
    '-----------------------------------
        
    Call DeleteRowInTabSTDPreparationAcquisition(AcquisitionID)
    
    '-----------------------------------
    ' cancello dalla tabella
    '-----------------------------------
    
    Grid2.ReadOnly = False
    Grid2.Selection.DeleteByRow
    Grid2.ReadOnly = True
    
    '-------------------------------------------
    ' ricarico la Gird1
    '-------------------------------------------
    
    uSTDPreparation.HannaCodes(IndexCode) = uHannaCodes(IndexCode)
    
    Call FillGridSTDPreparationFromFile(Grid1, uSTDPreparation, 1)
    
    uHannaCodes(IndexCode).QtyProduced = uSTDPreparation.HannaCodes(IndexCode).QtyProduced
    
    PopupMessage 2, "Acquisition Deleted....", , , AcquisitionHannaCode
    frCommandInside(11).Visible = False
End If

End Function



Private Function SaveSTDPreparation()
    
If Grid1.Rows > 1 Then
    
    With uSTDPreparation
        .HannaCodes = uHannaCodes
        .bSaved = True
        .Note = txFormulation(5)
        .numPrepWeek = txFormulation(2)
        .PlannedPrepWeek = txFormulation(3)
        .PlanningReference = txFormulation(4)
        .PreparationDate = txFormulation(1)
        .PrepWeek = txFormulation(6)
        .RecipeBy = txFormulation(0)
        .fileNameRecForProd = SettingName
    End With
    '-------------------------------------------
    ' Salva e aggiorna TabSTDPreparation
    '-------------------------------------------
    Call AggiornaTabSTDPreparation(STDPreparationID, uSTDPreparation)
    '-------------------------------------------
    ' Salva e aggiorna File
    '-------------------------------------------
    Call STDPreparationSaveSetting(uSTDPreparation, SettingName)
    
    
    PopupMessage 2, "STDPreparation correctly Saved", , , "STDPreparation"
    
Else
    
    PopupMessage 2, "Please Select HannaCode first and fill Recipe for STDPreparation Details....", , True, "STDPreparation"
    
End If
End Function


Private Sub SetSTDPreparationNoPreparation()

Dim i As Integer

For i = txFormulation.LBound To txFormulation.UBound
    
    txFormulation(i).Locked = False
    txFormulation(i).BackColor = vbRed
Next
frCommandInside(0).Visible = True

End Sub
Private Sub SetSettingName()
' LINE+DATERECIPE+PREPARATIONWEEK+PLANNEDPREPARATION
SettingName = FormatNomeFile(Trim(uHannaCodes(1).Code) & "." & Trim(uHannaCodes(1).Line) & "." & txFormulation(1) & "." & txFormulation(2) & "." & txFormulation(3)) & "." & USER_ESTENSIONE_RFP

End Sub


Private Sub SetMix1eMix2(ByVal strMix As String)
Dim sMix() As String
If strMix = "" Or InStr(strMix, ";") = 0 Then Exit Sub

sMix() = Split(strMix, ";")

If UBound(sMix) > 0 Then
    
    Grid2.Cell(0, 12).Text = "Lot " & sMix(0)
    Grid2.Cell(0, 13).Text = "Lot " & sMix(1)
    
    lbAcquisition(12) = "Lot " & sMix(0)
    lbAcquisition(13) = "Lot " & sMix(1)
End If


End Sub





Private Sub frExcel_Click()
Dim ExcelFilename As String
' export LOT Excel
If SettingName = "" Then
Else
    'ExcelFilename = "PROD_" & FormatNomeFile(Replace(SettingName, ".rfp", "")) '' FormatNomeFile(Trim(uRecipe(1).Line) & "." & Trim(uRecipe(1).code) & "." & txFormulation(1) & "." & txFormulation(2) & "." & txFormulation(3)) & ".xls"
    
    ExcelFilename = "PROD_" & FormatNomeFile(Trim(uHannaCodes(1).Recipe) & "." & Trim(uHannaCodes(1).Line) & "." & txFormulation(2) & "." & txFormulation(6)) & ".xls"

    
    
    
    If Len(ExcelFilename) > 31 Then ExcelFilename = Left$(ExcelFilename, 30)
    PopupMessage 2, "Exporting data to Excel : please wait...." & vbCrLf & ExcelFilename
    Call EsportaSTDPreparationExcel(SettingName, ExcelFilename, uSTDPreparation)
End If



End Sub
















Private Sub Grid4_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)


NotesID = 0

    With Grid4
    
        If FirstRow > 0 Then
        
            NotesID = .Cell(FirstRow, 5).Text
            txRevision(0) = .Cell(FirstRow, 1).Text
            txRevision(1) = .Cell(FirstRow, 2).Text
            txRevision(2) = .Cell(FirstRow, 4).Text
            txRevision(3) = .Cell(FirstRow, 3).Text
            
        End If
    
    End With


End Sub

Private Sub lbFunction_Click(Index As Integer)
ImCode_Click Index
End Sub

Private Sub ImCode_Click(Index As Integer)
Select Case Index
                        
                        
        Case 4
            ' aggiungi Notes specifics
            If AddNotes Then
                 Call GetSTDPreparationNotes(Grid4, SettingName)
            End If
            
            frExcel2.Visible = IIf(Grid4.Rows > 1, True, False)
            Frame6.Visible = IIf(Grid4.Rows > 1, False, True)
        Case 5
            ' delete Notes specifics
            If DeleteNotes(NotesID) Then
                 Call GetSTDPreparationNotes(Grid4, SettingName)
            End If
            
            frExcel2.Visible = IIf(Grid4.Rows > 1, True, False)
            Frame6.Visible = IIf(Grid4.Rows > 1, False, True)

End Select

End Sub




Private Sub ClearRevisionForm()


Dim i As Integer
For i = 0 To txRevision.UBound
    txRevision(i) = ""
Next
txRevision(2) = MyOperatore.Name

End Sub


Private Sub txRevision_Click(Index As Integer)
Dim userCode As String
Dim Answer As String
Dim Selected As String
Dim bNumber As Boolean
Dim sString As String
Dim rc As Boolean

    Selected = lbAcquisition(Index) ' "ScheduledSTD"
    Answer = txRevision(Index)
    sString = "Please Enter " & lbAcquisition(Index)
  
    bNumber = False
    cmbRevType.Visible = False
    If txRevision(2) = "" Then txRevision(2) = MyOperatore.Name

    Select Case Index
        Case 0
            If Answer = "" Then Answer = FormatDataLAT(Now())
        Case 1
            ' type
            cmbRevType.ZOrder
            cmbRevType.Visible = True
            Exit Sub
        Case 2
        
    End Select
    
    
    If txRevision(Index).Locked Then Exit Sub
    
    
  
    If F_InputBox.DoShow(sString, Selected, , , , Answer, , bNumber, Me.Top) Then
    
        txRevision(Index) = Answer
        
        Select Case Index
            Case 0
                ' isdate?
                If IsDate(Answer) Then
                     txRevision(Index) = FormatDataLAT(Answer)
                Else
                    PopupMessage 2, "Please enter a valid Date...", , True
                End If
        End Select
    End If
    
    
    
    
End Sub

Private Function DeleteNotes(ByVal ID As Long) As Boolean
Dim rc As Boolean
Dim i As Integer

rc = True

For i = 1 To txRevision.UBound
    If txRevision(i) = "" Then
        rc = False
        PopupMessage 2, "Please Select a Note form the table...", , True, "Delete Notes"
        DeleteNotes = rc
        Exit Function
    End If
Next

With dbTabSTDPreparationNotes
    .filter = ""
    .filter = "ID='" & ID & "'"
    If .EOF Then
        
    Else
        If F_MsgBox.DoShow("Delete Note?", "STDPreparation Notes", , "Delete", "Exit") Then
            .Delete
            .Update
        Else
            rc = False
        End If
    End If
     
        
        
        
    



End With

DeleteNotes = rc
End Function
Private Function AddNotes() As Boolean
Dim rc As Boolean
Dim i As Integer
Dim OldDate As Date

rc = True

For i = 1 To txRevision.UBound
    If txRevision(i) = "" Then
        rc = False
        PopupMessage 2, "Please enter all fields...", , True, "Notes History"
        AddNotes = rc
        Exit Function
    End If
Next





With dbTabSTDPreparationNotes
    .filter = ""
    .filter = "filename='" & SettingName & "' and NoteDate='" & txRevision(0) & "'"
    If .EOF Then
        .AddNew
    Else
        .MoveFirst
        OldDate = FormatDataLAT(Trim(!NoteDate))
        If F_MsgBox.DoShow("Note Date : " & OldDate & " already exsists.", "Add STDPreparation Note", , "Add", "Exit") Then
            .AddNew
        Else
            AddNotes = False
            Exit Function
        End If
    End If
        
        !NoteDate = txRevision(0)
        !Type = txRevision(1)
        !Description = IIf(Len(txRevision(3)) > 255, Left(txRevision(3), 255), txRevision(3))
        !Operator = txRevision(2)
        !FileName = SettingName
        .Update


End With

AddNotes = rc
End Function
Private Sub Image1_Click()
frExcel2_Click
End Sub
Private Sub Label4_Click()
frExcel2_Click
End Sub
Private Sub frExcel2_Click()
    If SettingName <> "" Then
        If Grid4.ExportToExcel(USER_DESKTOP & "\" & SettingName & "_STDPreparationNote_History.xls", True, True) Then
            MessageInfoTime = 2500
            PopupMessage 2, "File correcly created on Desktop", , , RecipeCode & "_Note_History.xls"
        End If
    End If
End Sub

Private Sub cmbRevType_Click()
txRevision(1) = cmbRevType
cmbRevType.Visible = False
End Sub


Private Sub AddcmbRevType()


    With cmbRevType
        .AddItem "Revision"
        .AddItem "Improvement"
        .AddItem "Issue"
        .ListIndex = 0
    End With

End Sub
