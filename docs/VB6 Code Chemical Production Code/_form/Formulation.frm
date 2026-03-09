VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form ReceiptForProduction 
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
   Icon            =   "Formulation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   12000
   ScaleWidth      =   19200
   Begin VB.PictureBox PBContainerViewport 
      BackColor       =   &H00F0F0F0&
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
      Index           =   1
      Left            =   0
      ScaleHeight     =   9975
      ScaleWidth      =   18975
      TabIndex        =   19
      Top             =   1080
      Width           =   18975
      Begin VB.Frame frIRequisition 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   3735
         Index           =   1
         Left            =   1800
         TabIndex        =   89
         Top             =   720
         Width           =   15255
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   4
            Left            =   6480
            TabIndex        =   102
            Top             =   1440
            Width           =   8775
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   100
            Top             =   1440
            Width           =   2175
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   2
            Left            =   13440
            TabIndex        =   98
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   1
            Left            =   6480
            TabIndex        =   96
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   93
            Top             =   960
            Width           =   2175
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00F0F0F0&
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
            Index           =   7
            Left            =   0
            TabIndex        =   90
            Top             =   0
            Width           =   15255
            Begin VB.Line Line10 
               BorderColor     =   &H00E0E0E0&
               X1              =   0
               X2              =   15240
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label lbInside 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Material Requisition Document"
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
               Index           =   7
               Left            =   0
               TabIndex        =   92
               Top             =   120
               Width           =   3510
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Material Requisition"
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
               Left            =   13395
               TabIndex        =   91
               Top             =   180
               Visible         =   0   'False
               Width           =   1755
            End
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fill Document form and save pdf for material requisition"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   10200
            TabIndex        =   109
            Top             =   3120
            Width           =   4920
         End
         Begin VB.Label lbDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reason of the request"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   4
            Left            =   4320
            TabIndex        =   101
            Top             =   1440
            Width           =   2100
         End
         Begin VB.Label lbDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "today "
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   99
            Top             =   1485
            Width           =   630
         End
         Begin VB.Label lbDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From Production line no./department"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   2
            Left            =   9600
            TabIndex        =   97
            Top             =   1005
            Width           =   3630
         End
         Begin VB.Label lbDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operator"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   1
            Left            =   5520
            TabIndex        =   95
            Top             =   1005
            Width           =   885
         End
         Begin VB.Label lbDocument 
            BackStyle       =   0  'Transparent
            Caption         =   "Document No: MR-"
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   94
            Top             =   1005
            Width           =   1935
         End
         Begin VB.Label lbpdf 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Save pdf for Material Requisition"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00644603&
            Height          =   255
            Left            =   840
            TabIndex        =   103
            Top             =   3120
            Width           =   3105
         End
         Begin VB.Image impdf 
            Height          =   480
            Left            =   240
            Picture         =   "Formulation.frx":33E2
            Top             =   3000
            Width           =   480
         End
         Begin VB.Label lbCommand 
            BackColor       =   &H00C0FFC0&
            Height          =   735
            Left            =   120
            TabIndex        =   104
            Top             =   2880
            Width           =   4095
         End
      End
      Begin VB.Frame frIRequisition 
         BackColor       =   &H00F0F0F0&
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
         Height          =   5055
         Index           =   0
         Left            =   1800
         TabIndex        =   78
         Top             =   4680
         Width           =   15255
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
            Index           =   6
            Left            =   9120
            TabIndex        =   87
            Top             =   4200
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Open pdf Folder"
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
               TabIndex        =   88
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.Frame Frame6 
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
            TabIndex        =   84
            Top             =   1200
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
               Index           =   4
               Left            =   1920
               TabIndex        =   85
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
            Index           =   5
            Left            =   12240
            TabIndex        =   82
            Top             =   4200
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Delete Record"
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
               TabIndex        =   83
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00F0F0F0&
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
            Index           =   6
            Left            =   0
            TabIndex        =   79
            Top             =   0
            Width           =   15255
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Material Requisition"
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
               Left            =   13395
               TabIndex        =   81
               Top             =   180
               Visible         =   0   'False
               Width           =   1755
            End
            Begin VB.Label lbInside 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Materials Requested Table"
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
               Index           =   6
               Left            =   0
               TabIndex        =   80
               Top             =   120
               Width           =   3015
            End
            Begin VB.Line Line9 
               BorderColor     =   &H00E0E0E0&
               X1              =   0
               X2              =   15240
               Y1              =   480
               Y2              =   480
            End
         End
         Begin FlexCell.Grid Grid6 
            Height          =   3135
            Left            =   0
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   720
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   5530
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
            DefaultFontSize =   9.75
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formulation needs (1) material requisition (2) formulation check out before Preparation"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   0
            TabIndex        =   108
            Top             =   4200
            Width           =   7710
         End
      End
   End
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
      Left            =   -240
      ScaleHeight     =   9975
      ScaleWidth      =   19245
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1080
      Width           =   19245
      Begin VB.PictureBox PBContainer 
         BackColor       =   &H00F0F0F0&
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
         Height          =   36215
         Left            =   0
         ScaleHeight     =   36210
         ScaleWidth      =   19155
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   -24360
         Width           =   19155
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
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
            Height          =   3360
            Index           =   2
            Left            =   1680
            TabIndex        =   117
            Top             =   15320
            Visible         =   0   'False
            Width           =   15255
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
               Height          =   540
               Index           =   8
               Left            =   0
               TabIndex        =   118
               Top             =   0
               Width           =   15255
               Begin VB.Line Line11 
                  BorderColor     =   &H00E0E0E0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mix Recipes"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000040C0&
                  Height          =   345
                  Index           =   8
                  Left            =   0
                  TabIndex        =   120
                  Top             =   75
                  Width           =   15225
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Formulation"
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
                  Left            =   14100
                  TabIndex        =   119
                  Top             =   180
                  Width           =   1050
               End
            End
            Begin FlexCell.Grid Grid7 
               Height          =   2415
               Left            =   0
               TabIndex        =   121
               TabStop         =   0   'False
               Top             =   600
               Width           =   15255
               _ExtentX        =   26908
               _ExtentY        =   4260
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
               DefaultFontSize =   9.75
               BoldFixedCell   =   0   'False
               DisplayRowIndex =   -1  'True
               DrawMode        =   1
               DefaultRowHeight=   20
               FixedRowColStyle=   0
               ForeColorFixed  =   8937488
               GridColor       =   15790320
               Rows            =   1
               ScrollBarStyle  =   0
               SelectionMode   =   3
               MultiSelect     =   0   'False
               DateFormat      =   0
            End
            Begin VB.Line Line12 
               BorderColor     =   &H00E0E0E0&
               X1              =   0
               X2              =   15240
               Y1              =   3120
               Y2              =   3120
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            Height          =   3615
            Index           =   6
            Left            =   600
            TabIndex        =   62
            Top             =   30960
            Width           =   18015
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   2520
               TabIndex        =   110
               Top             =   2160
               Width           =   13575
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00F0F0F0&
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
               Left            =   1080
               TabIndex        =   75
               Top             =   0
               Width           =   15255
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Formulation"
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
                  Left            =   14100
                  TabIndex        =   77
                  Top             =   180
                  Width           =   1050
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Settings"
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
                  Index           =   5
                  Left            =   0
                  TabIndex        =   76
                  Top             =   120
                  Width           =   855
               End
               Begin VB.Line Line8 
                  BorderColor     =   &H00E0E0E0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   8400
               TabIndex        =   74
               Top             =   1560
               Width           =   2535
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   4080
               TabIndex        =   72
               Top             =   1560
               Width           =   1695
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   14400
               TabIndex        =   70
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   7680
               TabIndex        =   68
               Top             =   960
               Width           =   3255
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
               Left            =   5760
               TabIndex        =   65
               Top             =   3000
               Width           =   6255
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Save Formulation"
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
                  TabIndex        =   66
                  Top             =   120
                  Width           =   6255
               End
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   2520
               TabIndex        =   64
               Top             =   960
               Width           =   3255
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Note"
               Height          =   255
               Index           =   5
               Left            =   1320
               TabIndex        =   111
               Top             =   2160
               Width           =   480
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Planning Reference"
               Height          =   255
               Index           =   4
               Left            =   6240
               TabIndex        =   73
               Top             =   1560
               Width           =   1860
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Planned Preparation Week"
               Height          =   255
               Index           =   3
               Left            =   1320
               TabIndex        =   71
               Top             =   1560
               Width           =   2595
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "# Prep Week"
               Height          =   255
               Index           =   2
               Left            =   12840
               TabIndex        =   69
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "Date Recipe"
               Height          =   255
               Index           =   1
               Left            =   6240
               TabIndex        =   67
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "Recipe by"
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   63
               Top             =   960
               Width           =   975
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
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
            Height          =   2055
            Index           =   5
            Left            =   1680
            TabIndex        =   57
            Top             =   28760
            Width           =   15255
            Begin VB.Frame Frame3 
               BackColor       =   &H00F0F0F0&
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
               Index           =   4
               Left            =   0
               TabIndex        =   58
               Top             =   0
               Width           =   15255
               Begin VB.Line Line6 
                  BorderColor     =   &H00E0E0E0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bottling"
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
                  Index           =   4
                  Left            =   0
                  TabIndex        =   60
                  Top             =   120
                  Width           =   840
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Formulation"
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
                  Left            =   14100
                  TabIndex        =   59
                  Top             =   180
                  Width           =   1050
               End
            End
            Begin FlexCell.Grid Grid5 
               Height          =   975
               Left            =   0
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   600
               Width           =   15255
               _ExtentX        =   26908
               _ExtentY        =   1720
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
               DefaultFontSize =   9.75
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
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
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
            Height          =   3735
            Index           =   4
            Left            =   1680
            TabIndex        =   50
            Top             =   24800
            Width           =   15255
            Begin VB.Frame Frame5 
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
               TabIndex        =   54
               Top             =   1080
               Visible         =   0   'False
               Width           =   5055
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Please select a code and Q.ty to produce..."
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
                  Left            =   315
                  TabIndex        =   55
                  Top             =   555
                  Width           =   4365
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00F0F0F0&
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
               Index           =   3
               Left            =   0
               TabIndex        =   51
               Top             =   0
               Width           =   15255
               Begin VB.Label Label8 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Formulation"
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
                  Left            =   14100
                  TabIndex        =   53
                  Top             =   180
                  Width           =   1050
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total Weight to Produce"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00644603&
                  Height          =   345
                  Index           =   3
                  Left            =   0
                  TabIndex        =   52
                  Top             =   75
                  Width           =   3270
               End
               Begin VB.Line Line7 
                  BorderColor     =   &H00E0E0E0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
            End
            Begin FlexCell.Grid Grid4 
               Height          =   2535
               Left            =   0
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   600
               Width           =   15255
               _ExtentX        =   26908
               _ExtentY        =   4471
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
               DefaultFontSize =   9.75
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
            Begin VB.Line Line15 
               BorderColor     =   &H00E0E0E0&
               X1              =   0
               X2              =   15240
               Y1              =   3240
               Y2              =   3240
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total weight to produce for each Recipe"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   106
               Top             =   3360
               Width           =   3615
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
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
            Height          =   4335
            Index           =   3
            Left            =   1680
            TabIndex        =   43
            Top             =   19200
            Visible         =   0   'False
            Width           =   15255
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
               Index           =   9
               Left            =   12240
               TabIndex        =   127
               Top             =   3600
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe List"
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
                  TabIndex        =   128
                  Top             =   120
                  Width           =   3015
               End
               Begin VB.Image Image4 
                  Height          =   240
                  Left            =   240
                  Picture         =   "Formulation.frx":5DD4
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   240
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
               Height          =   540
               Index           =   2
               Left            =   0
               TabIndex        =   46
               Top             =   0
               Width           =   15255
               Begin VB.Line Line5 
                  BorderColor     =   &H00E0E0E0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe Components : D002/1"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000040C0&
                  Height          =   285
                  Index           =   2
                  Left            =   0
                  TabIndex        =   48
                  Top             =   120
                  Width           =   15225
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Formulation"
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
                  Left            =   14100
                  TabIndex        =   47
                  Top             =   180
                  Width           =   1050
               End
            End
            Begin VB.Frame Frame4 
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
               TabIndex        =   44
               Top             =   1320
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
                  Index           =   2
                  Left            =   1920
                  TabIndex        =   45
                  Top             =   555
                  Width           =   1155
               End
            End
            Begin FlexCell.Grid Grid3 
               Height          =   2775
               Left            =   0
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   480
               Width           =   15255
               _ExtentX        =   26908
               _ExtentY        =   4895
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
               DefaultFontSize =   9.75
               BoldFixedCell   =   0   'False
               DisplayRowIndex =   -1  'True
               DrawMode        =   1
               DefaultRowHeight=   20
               FixedRowColStyle=   0
               ForeColorFixed  =   8937488
               GridColor       =   15790320
               Rows            =   1
               ScrollBarStyle  =   0
               SelectionMode   =   3
               MultiSelect     =   0   'False
               DateFormat      =   0
            End
            Begin VB.Line Line13 
               BorderColor     =   &H00E0E0E0&
               X1              =   0
               X2              =   15240
               Y1              =   3360
               Y2              =   3360
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
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
            Height          =   4215
            Index           =   1
            Left            =   1680
            TabIndex        =   34
            Top             =   10880
            Width           =   15255
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
               Left            =   5400
               TabIndex        =   40
               Top             =   1200
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
                  TabIndex        =   41
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
               Index           =   3
               Left            =   12240
               TabIndex        =   38
               Top             =   3480
               Width           =   3015
               Begin VB.Image Image1 
                  Height          =   240
                  Left            =   240
                  Picture         =   "Formulation.frx":87C6
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   240
               End
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Update"
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
                  TabIndex        =   39
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00F0F0F0&
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
               Index           =   1
               Left            =   0
               TabIndex        =   35
               Top             =   0
               Width           =   15255
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Formulation"
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
                  Left            =   14100
                  TabIndex        =   37
                  Top             =   180
                  Width           =   1050
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipes in Formulation"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00644603&
                  Height          =   345
                  Index           =   1
                  Left            =   0
                  TabIndex        =   36
                  Top             =   80
                  Width           =   3105
               End
               Begin VB.Line Line4 
                  BorderColor     =   &H00E0E0E0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
            End
            Begin FlexCell.Grid Grid2 
               Height          =   2655
               Left            =   0
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   600
               Width           =   15255
               _ExtentX        =   26908
               _ExtentY        =   4683
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
               DefaultFontSize =   9.75
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
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
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
            Height          =   7215
            Index           =   0
            Left            =   1680
            TabIndex        =   24
            Top             =   2640
            Width           =   15255
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
               Index           =   8
               Left            =   0
               TabIndex        =   125
               Top             =   5040
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Image Image5 
                  Height          =   240
                  Left            =   240
                  Picture         =   "Formulation.frx":B1B8
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   240
               End
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Delete Code"
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
                  TabIndex        =   126
                  Top             =   120
                  Width           =   3015
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
               Index           =   7
               Left            =   9120
               TabIndex        =   122
               Top             =   5040
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Reset Quantity"
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
                  TabIndex        =   123
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frQuantityCheck 
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   0  'None
               Caption         =   "Frame7"
               Height          =   735
               Left            =   0
               TabIndex        =   112
               Top             =   6120
               Width           =   6015
               Begin VB.PictureBox PicMin 
                  BackColor       =   &H000000C0&
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   2280
                  ScaleHeight     =   255
                  ScaleWidth      =   255
                  TabIndex        =   114
                  Top             =   240
                  Width           =   255
               End
               Begin VB.PictureBox Picture1 
                  BackColor       =   &H000000C0&
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   5040
                  ScaleHeight     =   255
                  ScaleWidth      =   255
                  TabIndex        =   113
                  Top             =   240
                  Width           =   255
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Min Quantity check"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   116
                  Top             =   240
                  Width           =   1860
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Max  Quantity check"
                  Height          =   255
                  Left            =   2880
                  TabIndex        =   115
                  Top             =   240
                  Width           =   1980
               End
               Begin VB.Shape Shape1 
                  BorderColor     =   &H00E0E0E0&
                  Height          =   735
                  Left            =   0
                  Top             =   0
                  Width           =   5535
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00F0F0F0&
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
               TabIndex        =   25
               Top             =   0
               Width           =   15255
               Begin VB.Line Line3 
                  BorderColor     =   &H00E0E0E0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hanna Codes in Formulation"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00644603&
                  Height          =   345
                  Index           =   0
                  Left            =   0
                  TabIndex        =   27
                  Top             =   80
                  Width           =   3900
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Formulation"
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
                  Left            =   14100
                  TabIndex        =   26
                  Top             =   180
                  Width           =   1050
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
               Index           =   2
               Left            =   12240
               TabIndex        =   32
               Top             =   5040
               Width           =   3015
               Begin VB.Image Image2 
                  Height          =   240
                  Left            =   240
                  Picture         =   "Formulation.frx":BBBA
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   240
               End
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Update"
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
                  TabIndex        =   33
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame Frame1 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   1335
               Left            =   5400
               TabIndex        =   30
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
                  TabIndex        =   31
                  Top             =   555
                  Width           =   1155
               End
            End
            Begin FlexCell.Grid Grid1 
               Height          =   4215
               Left            =   0
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   600
               Width           =   15255
               _ExtentX        =   26908
               _ExtentY        =   7435
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
               DefaultFontSize =   9.75
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
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Set Each Code Quantity to produce : click darker cells and set quantity"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   255
               Index           =   0
               Left            =   8880
               TabIndex        =   105
               Top             =   6360
               Width           =   6285
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
            Left            =   9840
            TabIndex        =   22
            Top             =   1440
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Select Recipe"
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
               TabIndex        =   23
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
            Index           =   0
            Left            =   6240
            TabIndex        =   20
            Top             =   1440
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Select Hanna Code"
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
               TabIndex        =   21
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.Label lbRecipes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Recipe to view component , raw materials and quantity to produce"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Left            =   360
            TabIndex        =   124
            Top             =   17400
            Width           =   11220
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Fill Settings form and Save formulation then check Material Requisition to print pdf report"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   107
            Top             =   34800
            Width           =   18945
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Select Hanna Code or Recipe to start formulation"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   18975
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
      TabIndex        =   16
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
         TabIndex        =   18
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
         TabIndex        =   17
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
      TabIndex        =   6
      Top             =   11040
      Width           =   19215
      Begin VB.Timer TimerBeginForm 
         Interval        =   10
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
         TabIndex        =   13
         Top             =   630
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
         TabIndex        =   12
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Formulation"
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
         Left            =   8880
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   630
         Width           =   1380
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   0
         Left            =   9360
         MousePointer    =   99  'Custom
         Picture         =   "Formulation.frx":E5AC
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   3
         Left            =   15600
         MousePointer    =   99  'Custom
         Picture         =   "Formulation.frx":1198E
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   4
         Left            =   18240
         MousePointer    =   99  'Custom
         Picture         =   "Formulation.frx":14D70
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
         TabIndex        =   10
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
         Index           =   2
         Left            =   0
         TabIndex        =   9
         Top             =   -120
         Visible         =   0   'False
         Width           =   1935
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
         TabIndex        =   8
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
         TabIndex        =   7
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
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   19215
      TabIndex        =   0
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
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
         Index           =   1
         Left            =   1920
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   1
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "Formulation.frx":18152
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Material Requisition"
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
            Index           =   1
            Left            =   90
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   720
            Width           =   1830
         End
      End
      Begin ChemicalProduction.ucScrollAdd ucScrollAdd1 
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
         ScaleWidth      =   1935
         TabIndex        =   3
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "Formulation.frx":1B534
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Formulation"
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
            TabIndex        =   4
            Top             =   720
            Width           =   1830
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formulation"
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
         Left            =   16605
         TabIndex        =   5
         Top             =   240
         Width           =   2325
      End
   End
End
Attribute VB_Name = "ReceiptForProduction"
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


Public Function DoShow(Optional ByVal ID As Long, Optional ByVal FileName As Long) As Boolean

    On Error GoTo ERR_SHOW
    
    m_rc = False
    mOk
    
    
    

    SettingName = FileName
    

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




Private Sub Form_Activate()
Me.WindowState = MainWindowState
End Sub

Private Sub Grid1_DblClick()
ucScrollAdd1.UCScrollV.ScrollToValue frInside(1).Top - 460
End Sub

Private Sub Grid2_DblClick()
frCommandInside_Click 1
End Sub

Private Sub Grid4_DblClick()
 ucScrollAdd1.UCScrollV.ScrollToValue 0
End Sub



Private Sub Grid7_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim Qty As String
Dim SelectedCode As String
Dim sString As String

Dim rc As Boolean

    
    SelectedCode = ""
    
    
    Call ColoraRiga(Grid7, 0)
    
    lbInside(2) = "Recipe Components "
    Grid3.Rows = 1
    
    If FirstRow < 1 Then Exit Sub
    
    SelectedCode = Grid7.Cell(FirstRow, 1).Text

    Call AddComponentInGrid3(Grid3, SelectedCode)
    
    lbInside(2) = "Recipe Components : " & SelectedCode
    
    Call ColoraRiga(Grid7, FirstRow)
    
    
    ' ucScrollAdd1.UCScrollV.ScrollToValue frInside(2).Top - 460
    


        
        
End Sub

Private Sub lbRecipes_Click()
ucScrollAdd1.UCScrollV.ScrollToValue frInside(1).Top - 460
End Sub

Private Sub TimerBeginForm_Timer()



uReceiptForProduction = uReceiptForProductionClean




Dim Grid(10) As Grid

Set Grid(0) = Grid1
Set Grid(1) = Grid2
Set Grid(2) = Grid3
Set Grid(3) = Grid4
Set Grid(4) = Grid5
Set Grid(5) = Grid6
Set Grid(6) = Grid7

Call SetAllFormulationGrid(Grid())

Dim i As Integer
For i = 3 To frInside.UBound
    frInside(i).Top = frInside(i).Top - (frInside(2).Height) * m_ControlGridRowHeight
Next
        



lbRecipes.Visible = True

TimerBeginForm.Enabled = False

End Sub


Private Sub Form_Load()


    PBContainer.Top = 0
    PBContainer.Left = 0
    PBContainer.Width = Me.Width
    lbCommand.BackColor = vbColorAzzurrino
    
    

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
    
    lbRecipes.Left = PBContainer.Width / 2 - lbRecipes.Width / 2
    

    PBContainerViewport(0).ZOrder
    PBFooter.ZOrder
    
    
    
    
    
End Sub

Private Sub Form_Resize()

    PBTitle.Width = Me.Width
    PBFooter.Top = Me.ScaleHeight - PBFooter.Height
    PBFooter.Width = Me.Width
 
    
    'Resize the container (needed to show the full bottom box on maximized state)
    'First resize our container
    ucScrollAdd1.ContainerW = Me.ScaleWidth
    'But also need to resize PBContainer wich hide the width of the bottom box

    
    
      ResizeControls

    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set ReceiptForProduction = Nothing
End Sub

Private Sub Grid1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim Qty As String

Dim sString As String

SelectedCode = ""

If FirstRow > 0 Then

    SelectedCode = Grid1.Cell(FirstRow, 1).Text
    Qty = Grid1.Cell(FirstRow, 6).Text
    sString = Grid1.Cell(0, 6).Text
    frCommandInside(8).Visible = False
    Select Case FirstCol
        Case 6
            ' Q.ty to produce
            If F_InputBox.DoShow(sString, SelectedCode, , , , Qty, , True, Me.Top) Then
            
                
                If IsNumeric(Qty) Or Qty = "" Then
                    Grid1.Cell(FirstRow, 6).Text = Qty
                    Grid1.Cell(FirstRow, 6).Alignment = cellCenterCenter
                End If
                
            End If
        Case Else
            frCommandInside(8).Visible = True
            
    
    
    End Select

End If
End Sub

Private Sub SetIntroFrame()


    
End Sub
Private Sub Grid2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim Qty As String
Dim SelectedCode As String
Dim sString As String

Dim Mix As String

Dim rc As Boolean

    frInside(3).Visible = False
    lbRecipes.Visible = True
    
    SelectedCode = ""
    
    
    Call ColoraRiga(Grid2, 0)
    Call ColoraRiga(Grid7, 0)
    
    Grid7.Rows = 1
    Grid3.Rows = 1
    
     SetMixTable False

If FirstRow < 1 Then Exit Sub

SelectedCode = Grid2.Cell(FirstRow, 1).Text
Qty = Grid2.Cell(FirstRow, 4).Text
sString = Grid2.Cell(0, 4).Text

Mix = Grid2.Cell(FirstRow, 7).Text
lbRecipes.Visible = True
        
Select Case FirstCol
    Case 4
        ' Q.ty to produce

        If F_InputBox.DoShow(sString, SelectedCode, , , , Qty, , True, Me.Top) Then
        
            
            If IsNumeric(Qty) Or Qty = "" Then
                Grid2.Cell(FirstRow, 6).Text = Qty
                Grid2.Cell(FirstRow, 6).Alignment = cellCenterCenter
            End If
            
        End If
    Case Else
    
        lbRecipes.Visible = False
        
      ' se ho Mix allora visualizzo la frInside(2)
        lbInside(8) = SelectedCode & " | Recipe Mixes"
        rc = IIf(Len(Trim(Mix)) > 0, True, False)
        frInside(3).Visible = True
        Call SetMixTable(rc)
        
        Call ColoraRiga(Grid2, FirstRow)
        
        frCommandInside(9).Visible = rc
        
        If rc Then
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(2).Top - 460
            Call AddComponentInGrid3(Grid7, SelectedCode)
            If Grid7.Rows > 1 Then
                ' se ci sono più ricette di ricette seleziono la prima....
                Call Grid7_SelChange(1, 1, 1, 1)
                Exit Sub
            End If

        Else
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(1).Top - 460
        End If
         Call AddComponentInGrid3(Grid3, SelectedCode)
        lbInside(2) = "Recipe Components : " & SelectedCode


End Select

        
        
End Sub

Private Sub ColoraRiga(ByVal Grid As Grid, ByVal lRow As Long)
Dim i As Integer
Dim t As Integer
Dim Count As Integer
Count = Grid.Rows - 1


If Count > 0 Then
    
    For i = 1 To Count
        For t = 0 To Grid.Cols - 1
            If i = lRow Then
                Grid.Cell(i, t).ForeColor = &H40C0&
                Grid.Cell(i, t).FontBold = True
            Else
                Grid.Cell(i, t).ForeColor = vbBlack
                Grid.Cell(i, t).FontBold = False
            End If
        Next
    Next
End If

End Sub



Private Sub impdf_Click()
lbCommand_Click
End Sub

Private Sub lbCommand_Click()
Call StampaMaterialRequisition

End Sub


Private Sub lbpdf_Click()
lbCommand_Click
End Sub


Private Sub txFormulation_Change(Index As Integer)
    Select Case Index
        Case 0
            ' operator...
            txDocument(1) = txFormulation(0)
        Case 1
            ' date...
            txDocument(3) = txFormulation(1)
    
    End Select
End Sub

Private Sub txFormulation_Click(Index As Integer)
Dim Answer As String
Dim Selected As String
Dim sString As String
Dim bNumber As Boolean

Selected = "Formulation"
Answer = txFormulation(Index)
sString = lbFormulation(Index)

bNumber = IIf(Index = 2, True, False)

        
        If F_InputBox.DoShow(sString, Selected, , , , Answer, , bNumber, Me.Top) Then
        
            txFormulation(Index) = Answer
            
            Select Case Index
                Case 1
                    ' isdate?
                    If IsDate(Answer) Then
                         txFormulation(Index) = CDate(Answer)
                    Else
                        PopupMessage 2, "Please insert a valid Date...", , True
                    End If
            End Select
        End If
        
        




End Sub

Private Sub txDocument_Click(Index As Integer)
Dim Answer As String
Dim Selected As String
Dim sString As String
Dim bNumber As Boolean

Selected = "Material Requisition"
Answer = txDocument(Index)
sString = lbDocument(Index)

bNumber = IIf(Index = 2, True, False)

        
        If F_InputBox.DoShow(sString, Selected, , , , Answer, , bNumber, Me.Top) Then
        
            txDocument(Index) = Answer
            
            Select Case Index
                Case 3
                    ' isdate?
                    If IsDate(Answer) Then
                         txDocument(Index) = CDate(Answer)
                    Else
                        PopupMessage 2, "Please insert a valid Date...", , True
                    End If
            End Select
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
    ElseIf ucScrollAdd1.UCScrollV.Value > frInside(3).Top And ucScrollAdd1.UCScrollV.Value <= frInside(4).Top Then
        IndexVisibleFrame = 4
    End If
              
        
   
    
End Sub

'Poorly made resizing functions just for the example
Private Sub RSRight(c As Control, Source As Object, adjust As Long, Optional LimitLeft& = -1, Optional LimitRight& = -1)
On Error Resume Next
Dim aux&
    aux& = (Source.ScaleWidth - c.Width) + adjust
    If (Err.NUMBER > 0) Then aux& = (Source.Width - c.Width) + adjust
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
    If (Err.NUMBER > 0) Then aux& = (Source.Height - c.Height) + adjust
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
        Unload Me
        
    Case 3
        ' Previous
         If IndexVisibleFrame >= 1 Then
            MyIndex = IndexVisibleFrame - 1
            If frInside(MyIndex).Visible = False Then
                MyIndex = IndexVisibleFrame - 3
            End If
            
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(MyIndex).Top - 480
        Else
            ucScrollAdd1.UCScrollV.ScrollToValue 0
         End If
    
    
    
    Case 4
        ' forward
        If IndexVisibleFrame < frInside.UBound Then
            MyIndex = IndexVisibleFrame + 1
            If frInside(MyIndex).Visible = False Then
                MyIndex = IndexVisibleFrame + 3
            End If
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(MyIndex).Top - 480
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
        DefaultMenuLabel_Click 2
    Case 39
        DefaultMenuLabel_Click 0
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
    Select Case Index
        Case 0
            ' select codes
            Call SelectCode
        Case 1
            ' select recipe
            Call SelectRecipe
        Case 2
            Call UpdateFormulation
        Case 3
            ' update Recipes Table
            Grid2.Cell(0, 0).SetFocus
            Call AddRecipeInFormulationGrid(Grid1, Grid2, Grid4)
            
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(1).Top - 480
                
        Case 6
           ' Debug.Print PathRequisition
            ApriIlReportFolder (USER_DOCUMENTI & PathRequisition)
        Case 7
            ' reset quantity Hanna Code
            If ResetQuantityHannaCode(Grid1) Then frCommandInside_Click 2
        Case 8
            ' cancella Hanna code
            If SelectedCode <> "" Then
                If F_MsgBox.DoShow("Warning : Delete Hanna Code from Table?", SelectedCode) Then
                    Grid1.ReadOnly = False
                    Grid1.Selection.DeleteByRow
                    Grid1.ReadOnly = True
                    Grid1.Refresh
                    ' update everything!!!!!!
                    frCommandInside_Click 2
                End If
            End If
        Case 9
            ucScrollAdd1.UCScrollV.ScrollToValue frInside(1).Top - 480
            

    End Select
End Sub


Private Sub frInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim i As Integer
    For i = 0 To frCommandInside.UBound

            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 4 Then
                frCommandInside(i).BackColor = &H8000&
            End If

    
    Next
 
 
End Sub

Private Sub frCommandInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
IndexDashCommInside = Index
Dim i As Integer
    For i = 0 To frCommandInside.UBound
        If i = Index Then
            frCommandInside(i).BackColor = &H846623
            lbCommandInside(i).ForeColor = vbWhite
            If i = 4 Then
                frCommandInside(i).BackColor = &H20A020
            End If
            
        Else
            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 4 Then
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

Private Sub UpdateFormulation()
    Call AddRecipeInFormulationGrid(Grid1, Grid2, Grid4)
    
End Sub
Private Sub SelectCode()
Dim HannaCode As String
Dim rc As Boolean
    FormCodes.ZOrder
    rc = FormCodes.DoShow(HannaCode) ', Me.Top
    
    If HannaCode = "" Then Exit Sub
    
    Call AddCodeAndRecipe(HannaCode)
   
    
    

End Sub
Private Sub SelectRecipe()
Dim RecipeCode As String
Dim rc As Boolean
Dim i As Integer
    FormRecipes.ZOrder
    rc = FormRecipes.DoShow(RecipeCode)
    If RecipeCode = "" Then Exit Sub
    With dbTabCode
        .filter = ""
        If .EOF Then
        Else
            .MoveFirst
            For i = 1 To .RecordCount
                If InStr(!Recipe, RecipeCode) Then
                    Call AddCodeAndRecipe(Trim(!Code), True)
                    .filter = ""
                    .Move i - 1
                End If
                .MoveNext
            Next
        
        End If
    End With
    
    
    ucScrollAdd1.UCScrollV.ScrollToValue frInside(1).Top - 480
    
    

End Sub

Private Sub AddCodeAndRecipe(ByVal HannaCode As String, Optional ByVal bSalta As Boolean)

 Call AddCodeInFormulationGrid(Grid1, HannaCode, bSalta)
    Call AddRecipeInFormulationGrid(Grid1, Grid2, Grid4)
    
End Sub
Private Sub StampaMaterialRequisition()
Dim rc As Boolean

rc = True

rc = CheckTxDocument
If rc Then rc = ReportStampato
If rc Then PopupMessage 2, "Report Succesfully Generated...", , , "Material Requisition"





End Sub

Private Function ReportStampato() As Boolean
    Dim rc As Boolean
    Dim NumReport As String
    On Error GoTo ERR_SAVE
    rc = True


    rc = OkStampa(NumReport, bSeStampa)
     
ERR_END:
    On Error GoTo 0
    ReportStampato = rc
    Exit Function
ERR_SAVE:
    rc = False
    Resume ERR_END
End Function

Private Function CheckTxDocument() As Boolean
Dim rc As Boolean
Dim i As Integer
    rc = True
    For i = txDocument.LBound To txDocument.UBound - 1
        If Len(txDocument(i)) = 0 Then
            rc = False
            PopupMessage 2, "Please Enter field : " & lbDocument(i), , True, "Formulation Document"
            txDocument(i).SetFocus
            Exit For
        End If
    Next
    CheckTxDocument = rc
End Function

Private Sub SetMixTable(ByVal bValue As Boolean)

' bValue = vero allora visualizzo e sposto le altre
Dim i As Integer
Dim frHeight As Double

frHeight = frInside(2).Height


    If bValue And frInside(2).Visible = False Then
    
        frInside(2).Visible = True
        
     
        For i = 3 To frInside.UBound
            frInside(i).Top = frInside(i).Top + (frHeight) * m_ControlGridRowHeight
        Next

    ElseIf bValue = False And frInside(2).Visible Then
        For i = 3 To frInside.UBound
            frInside(i).Top = frInside(i).Top - (frHeight) * m_ControlGridRowHeight
        Next
    ElseIf bValue = False And frInside(2).Visible = False Then
        
    End If
    
    frInside(2).Visible = bValue

End Sub
