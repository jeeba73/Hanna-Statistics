VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form frmRecipeForProduction 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Chemical Production"
   ClientHeight    =   12000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19200
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ReceiptForProduction.frx":0000
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
      Left            =   120
      ScaleHeight     =   9975
      ScaleWidth      =   19095
      TabIndex        =   18
      Top             =   960
      Width           =   19095
      Begin VB.Frame frIRequisition 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   3975
         Index           =   1
         Left            =   1800
         TabIndex        =   86
         Top             =   720
         Width           =   15615
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
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
            Height          =   375
            Index           =   5
            Left            =   1920
            TabIndex        =   132
            Top             =   1920
            Width           =   4575
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
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
            Height          =   375
            Index           =   4
            Left            =   6480
            TabIndex        =   99
            Top             =   1440
            Width           =   8775
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
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
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   97
            Top             =   1440
            Width           =   2175
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
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
            Height          =   375
            Index           =   2
            Left            =   12120
            TabIndex        =   95
            Top             =   960
            Width           =   3135
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
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
            Height          =   375
            Index           =   1
            Left            =   6480
            TabIndex        =   93
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
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
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   90
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
            TabIndex        =   87
            Top             =   0
            Width           =   15255
            Begin VB.Line Line10 
               BorderColor     =   &H00B0B0B0&
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
               TabIndex        =   89
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
               TabIndex        =   88
               Top             =   180
               Visible         =   0   'False
               Width           =   1755
            End
         End
         Begin VB.Label lbDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dep. Manager"
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
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   133
            Top             =   1965
            Width           =   1395
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
            TabIndex        =   105
            Top             =   2040
            Width           =   4920
         End
         Begin VB.Label lbDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Planning Reference"
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
            Height          =   255
            Index           =   4
            Left            =   4560
            TabIndex        =   98
            Top             =   1440
            Width           =   1860
         End
         Begin VB.Label lbDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "today "
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
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   96
            Top             =   1485
            Width           =   630
         End
         Begin VB.Label lbDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Production line no./dep."
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
            Height          =   255
            Index           =   2
            Left            =   9600
            TabIndex        =   94
            Top             =   1005
            Width           =   2370
         End
         Begin VB.Label lbDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operator"
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
            Height          =   255
            Index           =   1
            Left            =   5520
            TabIndex        =   92
            Top             =   1005
            Width           =   885
         End
         Begin VB.Label lbDocument 
            BackStyle       =   0  'Transparent
            Caption         =   "Document No: MR-"
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
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   91
            Top             =   1005
            Width           =   1935
         End
         Begin VB.Label lbpdf 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Save pdf for Material Requisition"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00644603&
            Height          =   255
            Left            =   6600
            TabIndex        =   100
            Top             =   3360
            Width           =   3105
         End
         Begin VB.Image impdf 
            Height          =   480
            Left            =   6000
            Picture         =   "ReceiptForProduction.frx":29F2
            Top             =   3240
            Width           =   480
         End
         Begin VB.Label lbCommand 
            BackColor       =   &H00C0FFC0&
            Height          =   735
            Left            =   5760
            TabIndex        =   101
            Top             =   3120
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
         TabIndex        =   75
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
            TabIndex        =   84
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
               TabIndex        =   85
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
            TabIndex        =   81
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
               TabIndex        =   82
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
            TabIndex        =   79
            Top             =   4200
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Delete Component"
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
               TabIndex        =   80
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
            TabIndex        =   76
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
               TabIndex        =   78
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
               TabIndex        =   77
               Top             =   120
               Width           =   3015
            End
            Begin VB.Line Line9 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   15240
               Y1              =   480
               Y2              =   480
            End
         End
         Begin FlexCell.Grid Grid6 
            Height          =   3135
            Left            =   0
            TabIndex        =   83
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
            DefaultFontSize =   8.25
            DisplayRowIndex =   -1  'True
            DrawMode        =   1
            DefaultRowHeight=   20
            FixedRowColStyle=   0
            ForeColorFixed  =   6571523
            GridColor       =   15790320
            Rows            =   1
            ScrollBarStyle  =   0
            SelectionMode   =   1
            MultiSelect     =   0   'False
            DateFormat      =   0
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Click on Component to change Quantity"
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
            TabIndex        =   104
            Top             =   4200
            Width           =   3645
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
      Left            =   0
      ScaleHeight     =   9975
      ScaleWidth      =   19245
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   960
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
         Height          =   52000
         Left            =   360
         ScaleHeight     =   52000
         ScaleMode       =   0  'User
         ScaleWidth      =   19155
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
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
            Height          =   3960
            Index           =   2
            Left            =   960
            TabIndex        =   113
            Top             =   16936
            Visible         =   0   'False
            Width           =   17175
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
               Index           =   10
               Left            =   0
               TabIndex        =   126
               Top             =   3360
               Width           =   6015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Material Requisition"
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
                  TabIndex        =   127
                  Top             =   120
                  Width           =   6015
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00A88030&
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
               TabIndex        =   114
               Top             =   0
               Width           =   17175
               Begin VB.Label Label6 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mixes "
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   15.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Left            =   120
                  TabIndex        =   140
                  Top             =   80
                  Width           =   1815
               End
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mix Recipes"
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
                  Index           =   8
                  Left            =   0
                  TabIndex        =   116
                  Top             =   110
                  Width           =   17115
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select Recipe to view Components"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   13920
                  TabIndex        =   115
                  Top             =   120
                  Width           =   3165
               End
            End
            Begin FlexCell.Grid Grid7 
               Height          =   2415
               Left            =   0
               TabIndex        =   117
               TabStop         =   0   'False
               Top             =   600
               Width           =   17175
               _ExtentX        =   30295
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
               DefaultFontSize =   8.25
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
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   17160
               Y1              =   3120
               Y2              =   3120
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            Height          =   4455
            Index           =   6
            Left            =   1080
            TabIndex        =   59
            Top             =   42840
            Width           =   18015
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   2520
               TabIndex        =   106
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
               TabIndex        =   72
               Top             =   240
               Width           =   15255
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Receipt for Production"
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
                  Left            =   13140
                  TabIndex        =   74
                  Top             =   180
                  Width           =   2010
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Description"
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
                  TabIndex        =   73
                  Top             =   120
                  Width           =   1275
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
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   8400
               TabIndex        =   71
               Top             =   1560
               Width           =   2535
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   4080
               TabIndex        =   69
               Top             =   1560
               Width           =   1695
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   14400
               TabIndex        =   67
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   7680
               TabIndex        =   65
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
               Left            =   5880
               TabIndex        =   62
               Top             =   3240
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
                  TabIndex        =   63
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
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   2520
               TabIndex        =   61
               Top             =   960
               Width           =   3255
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Fill Description form and Save Recipe For Production , then check Material Requisition and create pdf report"
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
               TabIndex        =   125
               Top             =   3960
               Width           =   17985
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Note"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   1320
               TabIndex        =   107
               Top             =   2160
               Width           =   480
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Planning Reference"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   6240
               TabIndex        =   70
               Top             =   1560
               Width           =   1860
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Planned Preparation Week"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   1320
               TabIndex        =   68
               Top             =   1560
               Width           =   2595
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "# Prep Week"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   12840
               TabIndex        =   66
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "Date Recipe"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   6240
               TabIndex        =   64
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "Recipe by"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   60
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
            Height          =   3015
            Index           =   5
            Left            =   2040
            TabIndex        =   54
            Top             =   38280
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
               TabIndex        =   55
               Top             =   0
               Width           =   15255
               Begin VB.Line Line6 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Packaging"
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
                  TabIndex        =   57
                  Top             =   120
                  Width           =   1290
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe for Production"
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
                  Left            =   13215
                  TabIndex        =   56
                  Top             =   180
                  Width           =   1935
               End
            End
            Begin FlexCell.Grid Grid5 
               Height          =   1815
               Left            =   0
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   720
               Width           =   15255
               _ExtentX        =   26908
               _ExtentY        =   3201
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
            Height          =   6375
            Index           =   4
            Left            =   1920
            TabIndex        =   47
            Top             =   29838
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
               Index           =   14
               Left            =   12240
               TabIndex        =   138
               Top             =   5880
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
                  Index           =   14
                  Left            =   0
                  TabIndex        =   139
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
               Index           =   13
               Left            =   0
               TabIndex        =   136
               Top             =   5880
               Width           =   6015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Material Requisition Mixes"
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
                  Width           =   6015
               End
            End
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
               Left            =   5040
               TabIndex        =   51
               Top             =   3000
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
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   3
                  Left            =   405
                  TabIndex        =   52
                  Top             =   555
                  Width           =   4185
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
               TabIndex        =   48
               Top             =   0
               Width           =   15255
               Begin VB.Label Label8 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe for Production"
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
                  Left            =   13215
                  TabIndex        =   50
                  Top             =   180
                  Width           =   1935
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total Weight to Produce ( Recipes | Mixes )"
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
                  TabIndex        =   49
                  Top             =   75
                  Width           =   5910
               End
               Begin VB.Line Line7 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
            End
            Begin FlexCell.Grid Grid4 
               Height          =   5055
               Left            =   0
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   600
               Width           =   15255
               _ExtentX        =   26908
               _ExtentY        =   8916
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
            Begin VB.Line Line15 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   15240
               Y1              =   5760
               Y2              =   5760
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total weight to produce for each Recipe / Mix"
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
               Left            =   6360
               TabIndex        =   103
               Top             =   5880
               Width           =   5640
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
            Height          =   6015
            Index           =   3
            Left            =   1920
            TabIndex        =   41
            Top             =   22155
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
               TabIndex        =   123
               Top             =   5400
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
                  TabIndex        =   124
                  Top             =   120
                  Width           =   3015
               End
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
               TabIndex        =   44
               Top             =   0
               Width           =   15255
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Components"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   15.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Left            =   120
                  TabIndex        =   141
                  Top             =   80
                  Width           =   1950
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
                  ForeColor       =   &H00E0E0E0&
                  Height          =   285
                  Index           =   2
                  Left            =   45
                  TabIndex        =   45
                  Top             =   110
                  Width           =   15135
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
               Left            =   5040
               TabIndex        =   42
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
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   2
                  Left            =   1890
                  TabIndex        =   43
                  Top             =   555
                  Width           =   1215
               End
            End
            Begin FlexCell.Grid Grid3 
               Height          =   4575
               Left            =   0
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   600
               Width           =   15255
               _ExtentX        =   26908
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
            Begin VB.Line Line4 
               BorderColor     =   &H00B0B0B0&
               X1              =   120
               X2              =   17160
               Y1              =   5280
               Y2              =   5280
            End
            Begin VB.Line Line13 
               BorderColor     =   &H00B0B0B0&
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
            Height          =   5055
            Index           =   1
            Left            =   960
            TabIndex        =   32
            Top             =   10880
            Width           =   17175
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
               Index           =   15
               Left            =   11040
               TabIndex        =   142
               Top             =   4440
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hanna Codes"
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
                  Index           =   15
                  Left            =   0
                  TabIndex        =   143
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
               Index           =   11
               Left            =   0
               TabIndex        =   130
               Top             =   4440
               Width           =   6015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Material Requisition"
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
                  TabIndex        =   131
                  Top             =   120
                  Width           =   6015
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
               Left            =   5880
               TabIndex        =   38
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
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   0
                  Left            =   1890
                  TabIndex        =   39
                  Top             =   555
                  Width           =   1215
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
               TabIndex        =   36
               Top             =   4440
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total Weights"
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
                  TabIndex        =   37
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame Frame3 
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
               Height          =   540
               Index           =   1
               Left            =   0
               TabIndex        =   33
               Top             =   0
               Width           =   17175
               Begin VB.Image ImViewRecipes 
                  Height          =   240
                  Index           =   1
                  Left            =   840
                  Picture         =   "ReceiptForProduction.frx":53E4
                  Top             =   120
                  Width           =   240
               End
               Begin VB.Image ImViewRecipes 
                  Height          =   240
                  Index           =   0
                  Left            =   240
                  Picture         =   "ReceiptForProduction.frx":5DE6
                  Top             =   120
                  Width           =   240
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select Recipe to view Components / Mixes"
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
                  Left            =   13200
                  TabIndex        =   35
                  Top             =   120
                  Width           =   3825
               End
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipes"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   15.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00644603&
                  Height          =   375
                  Index           =   1
                  Left            =   7845
                  TabIndex        =   34
                  Top             =   75
                  Width           =   1215
               End
            End
            Begin FlexCell.Grid Grid2 
               Height          =   3615
               Left            =   0
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   600
               Width           =   17175
               _ExtentX        =   30295
               _ExtentY        =   6376
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
            Begin VB.Line Line5 
               BorderColor     =   &H00D0D0D0&
               X1              =   0
               X2              =   17160
               Y1              =   4320
               Y2              =   4320
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
            TabIndex        =   23
            Top             =   3000
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
               Index           =   12
               Left            =   6000
               TabIndex        =   128
               Top             =   5040
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Delete All"
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
                  TabIndex        =   129
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
               Index           =   8
               Left            =   0
               TabIndex        =   121
               Top             =   5040
               Visible         =   0   'False
               Width           =   3015
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
                  TabIndex        =   122
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
               TabIndex        =   118
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
                  TabIndex        =   119
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
               TabIndex        =   108
               Top             =   6120
               Visible         =   0   'False
               Width           =   6015
               Begin VB.PictureBox PicMin 
                  BackColor       =   &H000000C0&
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   2280
                  ScaleHeight     =   255
                  ScaleWidth      =   255
                  TabIndex        =   110
                  Top             =   240
                  Width           =   255
               End
               Begin VB.PictureBox PicMax 
                  BackColor       =   &H000000C0&
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   5040
                  ScaleHeight     =   255
                  ScaleWidth      =   255
                  TabIndex        =   109
                  Top             =   240
                  Width           =   255
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Min Quantity check"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   112
                  Top             =   240
                  Width           =   1860
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Max  Quantity check"
                  Height          =   255
                  Left            =   2880
                  TabIndex        =   111
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
               TabIndex        =   24
               Top             =   0
               Width           =   15255
               Begin VB.Line Line3 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hanna Codes"
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
                  TabIndex        =   26
                  Top             =   75
                  Width           =   1890
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe for Production"
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
                  Left            =   13215
                  TabIndex        =   25
                  Top             =   180
                  Width           =   1935
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
               TabIndex        =   30
               Top             =   5040
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipes List"
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
                  TabIndex        =   31
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
               TabIndex        =   28
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
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   1
                  Left            =   1890
                  TabIndex        =   29
                  Top             =   555
                  Width           =   1215
               End
            End
            Begin FlexCell.Grid Grid1 
               Height          =   4095
               Left            =   0
               TabIndex        =   134
               TabStop         =   0   'False
               Top             =   600
               Width           =   15255
               _ExtentX        =   26908
               _ExtentY        =   7223
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
            Begin VB.Line Line2 
               BorderColor     =   &H00D0D0D0&
               X1              =   0
               X2              =   15240
               Y1              =   4860
               Y2              =   4860
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
               TabIndex        =   102
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
            TabIndex        =   21
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
               TabIndex        =   22
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
            TabIndex        =   19
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
               TabIndex        =   20
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
            ForeColor       =   &H00606060&
            Height          =   375
            Left            =   360
            TabIndex        =   120
            Top             =   17400
            Width           =   11220
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Select Hanna Code or Recipe to start formulation"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   27
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
      TabIndex        =   15
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         Interval        =   1
         Left            =   8400
         Top             =   120
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   2
         Left            =   0
         MousePointer    =   99  'Custom
         TabIndex        =   145
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mat.Req. folder"
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
         Index           =   5
         Left            =   300
         MousePointer    =   99  'Custom
         TabIndex        =   144
         Top             =   675
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   2
         Left            =   600
         MousePointer    =   99  'Custom
         Picture         =   "ReceiptForProduction.frx":67E8
         Top             =   120
         Visible         =   0   'False
         Width           =   480
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
         Left            =   17760
         MousePointer    =   99  'Custom
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   660
         Width           =   1230
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Recipe for Production "
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
         Left            =   8475
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   660
         Width           =   2190
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   0
         Left            =   9360
         MousePointer    =   99  'Custom
         Picture         =   "ReceiptForProduction.frx":9BCA
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   3
         Left            =   15600
         MousePointer    =   99  'Custom
         Picture         =   "ReceiptForProduction.frx":CFAC
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   4
         Left            =   18240
         MousePointer    =   99  'Custom
         Picture         =   "ReceiptForProduction.frx":1038E
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
         TabIndex        =   9
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
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
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
         Left            =   2160
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   2175
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   840
            MousePointer    =   99  'Custom
            Picture         =   "ReceiptForProduction.frx":13770
            Top             =   120
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
            Top             =   640
            Width           =   2070
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
         ScaleWidth      =   2175
         TabIndex        =   3
         Top             =   0
         Width           =   2175
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   840
            MousePointer    =   99  'Custom
            Picture         =   "ReceiptForProduction.frx":16B52
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Recipe for Production"
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
            Top             =   640
            Width           =   2070
         End
      End
      Begin VB.Label lbWait 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "Wait : Loading Data..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5760
         TabIndex        =   135
         Top             =   360
         Visible         =   0   'False
         Width           =   7575
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recipe for Production"
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
         Left            =   14655
         TabIndex        =   5
         Top             =   200
         Width           =   4275
      End
   End
   Begin VB.Line Line1 
      X1              =   9000
      X2              =   10200
      Y1              =   5760
      Y2              =   6240
   End
End
Attribute VB_Name = "frmRecipeForProduction"
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


Private ProductionWay() As ProdWay
Private uRecipeForProduction As RecipeForProduction
Private uMaterialRequisition As MaterialRequisition

Private SettingName As String
Private bImportata As Boolean
Private bIfDataPath As Boolean
Private bfrInsideMoveTop As Boolean


Private strLine As String
Private Nr As String
Private nrWeek As String
Private nrYear As String

Private bCancelUpdate As Boolean







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


End Sub


Public Function DoShow(Optional ByVal ID As Long, Optional ByVal FileName As String) As Boolean

    On Error GoTo ERR_SHOW
    
    m_rc = False
    mOk

    bIfDataPath = IIf(USER_PATH = USER_DATA_PATH, True, False)

    SettingName = FileName
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




Private Sub Command1_Click()
Dim i As Integer
Dim t As Integer

    Debug.Print
    Debug.Print
    
    For i = 1 To UBound(uRecipe)
        Debug.Print "i=" & i & "  Recipe : " & uRecipe(i).Code
        Debug.Print
        For t = 0 To UBound(uRecipe(i).RmxRecipe)
            
            Debug.Print "i=" & i & "  t=" & t & " : " & uRecipe(i).RmxRecipe(t).CHCode & " , Qty = " & uRecipe(i).RmxRecipe(t).TheoreticalWeight
        
        Next
        Debug.Print
    Next
    
    Debug.Print
    
End Sub

Private Sub Form_Activate()
Me.WindowState = MainWindowState
End Sub

Private Sub Grid1_Click()

If lRowHanna > 0 And lColHanna = 6 Then
    Call SetQtyHannaCode
End If
If bCODLine Then
    If lRowHanna > 0 And lColHanna = 10 Then
        Call SetLotNumberHannaCode
    End If
End If

End Sub

Private Sub Grid1_DblClick()
PBContainer.Top = -(frInside(1).Top - 460)
End Sub



Private Sub Grid2_Click()
If lRowRecipe > 0 And lColRecipe = 4 Then
   Call CheckRecipeQtytoProduce
End If
End Sub

Private Sub Grid2_DblClick()
 If F_MsgBox.DoShow("Open Recipes List to change Recipe in Recipe For Production?", "Recipe For Production", True) Then frCommandInside_Click 1
End Sub

Private Sub Grid3_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim rc As Boolean
Dim strMix As String
On Error GoTo ERR_GRID
If FirstRow > 0 Then
    
    
    rc = Grid3.Cell(FirstRow, 10).Text
    strMix = Trim(Grid3.Cell(FirstRow, 1).Text)
    If rc Then
    
       
    End If


End If


ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_GRID:
    MsgBox err.Description



End Sub

Private Sub Grid5_CellChange(ByVal Row As Long, ByVal Col As Long)
Select Case Col
    
    Case 2
        


End Select

End Sub

Private Sub Grid5_ComboClick(ByVal Index As Integer)

Call SetEstimatedTime(lRowCombo)

End Sub

Private Sub Grid5_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
Dim rc As Boolean
Dim i As Integer

    lRowCombo = Row
 
    With Grid5

        rc = SetProductionWay(uRecipe(Row).Line, ProductionWay())
            
       
        If rc Then
            .ComboBox(3).Clear
            .ComboBox(3).Font = "Calibri"
            .ComboBox(3).Font.Size = 12
            
            For i = LBound(ProductionWay) To UBound(ProductionWay)
                .ComboBox(3).AddItem ProductionWay(i).Production
            Next
        
        Else
           
        End If
       
    End With


End Sub

Private Sub Grid5_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim Speed As String
Dim i As Integer
With Grid5
    If FirstRow > 0 Then
    
        i = FirstRow
        
        Select Case FirstCol
            Case 5
                ' se non ho inserito la macchina allora salto
                If .Cell(i, 3).Text = "" Or .Cell(i, 4).Text = "" Then
                    
                Else
                
                    Speed = .Cell(i, 5).Text
                    If F_InputBox.DoShow("Please Enter Speed", "Evaluated Machine Speed", , , , Speed, , True) Then
                        .Cell(i, 5).Text = Speed
                    End If
                    Call SetEstimatedTime(i)
                    .Cell(i, 6).SetFocus
                End If
        
        End Select
    
    
    End If
End With
End Sub

Private Sub Grid6_Click()
Dim strNote As String
strNote = Grid6.Cell(lRowMaterialReq, 7).Text
If lColMaterialReq = 7 Then
    
    If F_InputBox.DoShow("Confirm or Change Note", "MATERIAL REQ NOTE", , , , strNote) Then
         Grid6.Cell(lRowMaterialReq, 7).Text = (strNote)
    End If

Else
    Call ChangeMaterialReqQty(Grid6, lRowMaterialReq)
End If


End Sub

Private Sub Grid6_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
lRowMaterialReq = FirstRow
lColMaterialReq = FirstCol


End Sub



Private Sub Grid7_Click()
If lRowMixes > 0 And lColMixes = 4 Then
    Call CheckMixesMultipleProduce
End If
End Sub



Private Sub Grid4_DblClick()
 PBContainer.Top = -(frInside(2).Top - 460)
End Sub

Private Sub Grid7_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim Qty As String
Dim sString As String
Dim rc As Boolean

    lRowMixes = FirstRow
    lColMixes = FirstCol
    
    SelectedMixCode = ""
    
        
    'Call ColoraRiga(Grid7, 0)
    
    lbInside(2) = "Recipe Components "
    Grid3.Rows = 1
    
    If FirstRow < 1 Then Exit Sub
    
    
    Call ColoraRiga(Grid7, FirstRow)
    
    
    SelectedMixCode = Grid7.Cell(FirstRow, 1).Text
    
    Call AddComponentInGrid3(Grid3, SelectedMixCode)
    
    lbInside(2) = SelectedMixCode & " : " & uRecipe(IndexRecipe).RmxRecipe(indexMix).Description
    
    
    
    
    indexMix = FirstRow - 1
    
    Call UpdateValue(True)
    Call AddComponentTheorethicalWeight(Grid3, SelectedMixCode, IndexRecipe, indexMix)



Select Case FirstCol
    Case 4
    
        Call CheckMixesMultipleProduce
    
        Exit Sub
    
        

End Select


End Sub

Private Sub Image4_Click(Index As Integer)

End Sub

Private Sub ImViewRecipes_Click(Index As Integer)
    Select Case Index
        Case 0
            ' ripristina tutte le Rows..
            Call ViewRecipesRFP(uRecipe, Grid1, Grid2, Grid4, Grid5, True)
        Case 1
            Call ViewRecipesRFP(uRecipe, Grid1, Grid2, Grid4, Grid5, False)
    
    End Select
End Sub

Private Sub lbRecipes_Click()
PBContainer.Top = -(frInside(1).Top - 460)
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
    
            
    For i = txDocument.LBound To txDocument.UBound
        txDocument(i) = ""
        
    Next
    
    For i = txFormulation.LBound To txFormulation.UBound
        txFormulation(i) = ""
        
    Next
    
    txDocument(2) = GetSetting(App.Title, "Workstation", "no.department", "")
    
    
    lbRecipes.Visible = True
    
    
    '--------------------------------------
    '
    '   Recipe importata
    '
    '--------------------------------------
    
    
    If bImportata Then
        GetFileInfo
    Else
       txFormulation(0) = MyOperatore.Name
    End If
    
    
    '--------------------------------------


End Sub
Private Sub InitForm()



    frCommandInside(10).Visible = bImportata
    frCommandInside(11).Visible = bImportata
    frCommandInside(13).Visible = False
    
    PicMenu(1).Visible = bImportata

        
        
    lbCommandInside(10) = "Material Requisition Mixes"
    lbCommandInside(11) = "Material Requisition Components"
    
    'frCommandInside(10).Visible = False
    'frCommandInside(11).Visible = False
    'PicMenu(1).Visible = False

    uRecipeForProduction = uRecipeForProductionClean
    
    ReDim uRecipe(0)
    ReDim ProductionWay(0)
    
    lbInside(6) = "Materials Requested Table"
    
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
    Set Grid(2) = Grid3
    Set Grid(3) = Grid4
    Set Grid(4) = Grid5
    Set Grid(5) = Grid6
    Set Grid(6) = Grid7
    
    Call SetAllRecipeForProductionGrid(Grid())
    Call SetColumnWidth
    
    Grid1.FrozenCols = 2
    Grid2.FrozenCols = 2
    Grid3.FrozenCols = 2
    Grid4.FrozenCols = 2
    Grid5.FrozenCols = 2
    Grid6.FrozenCols = 2
    Grid7.FrozenCols = 2
    
    bCODLine = IIf(InStr(UserLine, "59"), True, False)
    Grid1.Column(10).Width = IIf(bCODLine, 100, 0)
    
    
    

End Sub
Private Sub Form_Load()


    PBContainer.Top = 0
    PBContainer.Left = 0
    PBContainer.Width = Me.Width
    lbCommand.BackColor = vbColorResults
    
    
    
    

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
Set frmRecipeForProduction = Nothing
End Sub

Private Sub Grid1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim Qty As String

Dim sString As String

SelectedCode = ""
lColHanna = FirstCol
lRowHanna = FirstRow


If FirstRow > 0 Then

     

    SelectedCode = Grid1.Cell(FirstRow, 1).Text
    Qty = Grid1.Cell(FirstRow, 6).Text
    sString = Grid1.Cell(0, 6).Text
    frCommandInside(8).Visible = False
    Select Case FirstCol
        Case 6
           ' Q.ty to produce
         '  SetQtyHannaCode
        Case Else
            frCommandInside(8).Visible = True
            
    
    
    End Select

End If
End Sub

Private Sub SetQtyHannaCode()
  ' Q.ty to produce
Dim Recipe As String
Dim Qty As String

Dim sString As String
 Qty = Grid1.Cell(lRowHanna, 6).Text
 Recipe = Grid1.Cell(lRowHanna, 7).Text
   sString = Grid1.Cell(0, 6).Text
    
            If F_InputBox.DoShow(sString, SelectedCode & " | " & Recipe, , , , Qty, , True, Me.Top) Then
            
                
                If IsNumeric(Qty) Or Qty = "" Then
                
                    If bImportata Then bImportata = False
                    
                    Grid1.Cell(lRowHanna, 6).Text = Qty
                    Grid1.Cell(lRowHanna, 6).Alignment = cellCenterCenter
                    Grid1.Refresh
                    Grid1.AutoRedraw = True
                    Call UpdateValue
                    
                    If bCODLine Then SetLotNumberHannaCode
                End If
                
            End If
End Sub


Private Sub SetLotNumberHannaCode()
  ' Lot
Dim Recipe As String
Dim Lot As String

Dim sString As String
 Lot = Grid1.Cell(lRowHanna, 10).Text
 sString = "Please Enter Lot Number (es.0221)"
    
            If F_InputBox.DoShow(sString, SelectedCode & "- Production Lot Number", , , , Lot) Then
            
                
                'If Len(Lot) > 4 Or Len(Lot) < 4 Then
                    'PopupMessage 2, "Lot Number is Invalid!", "Production Lot Number", True
               ' End If
                
                Grid1.Cell(lRowHanna, 10).Text = Lot
                Grid1.Cell(lRowHanna, 10).Alignment = cellCenterCenter
                Grid1.Refresh
                Grid1.AutoRedraw = True
           
                
            End If
End Sub
Private Sub Grid2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim Qty As String

Dim sString As String

Dim Mix As String

lRowRecipe = FirstRow
lColRecipe = FirstCol

Dim rc As Boolean


    On Error GoTo ERR_SEL:


    frInside(3).Visible = False
    lbRecipes.Visible = True
    
    SelectedRecipeCode = ""
    

    Grid7.Rows = 1
    Grid3.Rows = 1
    
     SetMixTable False

    If FirstRow < 1 Then Exit Sub
    
    SelectedRecipeCode = Grid2.Cell(FirstRow, 1).Text
    


    lbCommandInside(10) = "Material Requisition Mixes : " & SelectedRecipeCode
    lbCommandInside(11) = "Material Requisition Components : " & SelectedRecipeCode
    
    
    lbInside(6) = "Materials Requested Table : " & SelectedRecipeCode
    
    Qty = Grid2.Cell(FirstRow, 4).Text
    sString = Grid2.Cell(0, 4).Text
    
    Mix = Grid2.Cell(FirstRow, 7).Text
    lbRecipes.Visible = True
    



    Select Case FirstCol
        Case 4
            Call CheckRecipeQtytoProduce
    End Select

    lbRecipes.Visible = False
    
  ' se ho Mix allora visualizzo la frInside(2)
    lbInside(8) = SelectedRecipeCode & " : " & uRecipe(FirstRow).Description
    
    'rc = IIf(Len(Trim(Mix)) > 0, True, False)
    'If rc Then
    rc = IfAllMixes(SelectedRecipeCode)
    


    frInside(3).Visible = True
    Call SetMixTable(rc)
    
    Call ColoraRiga(Grid2, FirstRow)
    
    frCommandInside(9).Visible = rc
    
     IndexRecipe = FirstRow
    
    If rc Then
    

        Call AddComponentInGrid7(Grid7, uRecipe(FirstRow).Code, Grid4, uRecipe(IndexRecipe))
        

        If Grid7.Rows > 1 Then
            ' se ci sono più ricette di ricette seleziono la prima....
            
            
            Call Grid7_SelChange(1, 1, 1, 1)
            Call UpdateValue(True)
            Exit Sub
        End If

    Else
        
        PBContainer.Top = -(frInside(1).Top - 460)
    End If
    

     Call AddComponentInGrid3(Grid3, SelectedRecipeCode)
    lbInside(2) = SelectedRecipeCode & " : " & uRecipe(IndexRecipe).Description


    Call UpdateValue(True)
    
    If rc = False Then Call AddComponentTheorethicalWeight(Grid3, SelectedRecipeCode, IndexRecipe, 0)
   
ERR_END:

    On Error GoTo 0
    Exit Sub
ERR_SEL:
    MsgBox err.Description
    Resume ERR_END
End Sub

Private Sub CheckRecipeQtytoProduce()
Dim sString As String
Dim SelectedRecipeCode As String
Dim Qty As String
Dim MinMultiple As String
        ' Q.ty to produce

    MinMultiple = Grid2.Cell(lRowRecipe, 12).Text & " " & Grid2.Cell(lRowRecipe, 13).Text
    SelectedRecipeCode = Grid2.Cell(lRowRecipe, 1).Text
    Qty = Grid2.Cell(lRowRecipe, 4).Text
    sString = "Enter " & Grid2.Cell(0, 4).Text & "  ( Min multiple = " & MinMultiple & " )"

        If F_InputBox.DoShow(sString, SelectedRecipeCode, , , , Qty, , True, Me.Top) Then
        
            
            If IsNumeric(Qty) Or Qty = "" Then
            
                If bImportata Then bImportata = False
                
                Grid2.Cell(lRowRecipe, 4).Text = Qty
                Grid2.Cell(lRowRecipe, 4).Alignment = cellCenterCenter
                
                Call UpdateValue(True)
                
            End If
            
        End If
End Sub

Private Sub CheckMixesMultipleProduce()
Dim sString As String
Dim SelectedRecipeCode As String
Dim Qty As String
Dim MinMultiple As String
        ' Q.ty to produce

MinMultiple = Grid7.Cell(lRowMixes, 13).Text & " " & Grid7.Cell(lRowMixes, 14).Text

SelectedRecipeCode = Grid7.Cell(lRowMixes, 1).Text
Qty = Grid7.Cell(lRowMixes, 4).Text
sString = "Enter " & Grid7.Cell(0, 4).Text & "  ( Min multiple = " & MinMultiple & " )"

        If F_InputBox.DoShow(sString, SelectedRecipeCode, , , , Qty, , True, Me.Top) Then
        
            
            If IsNumeric(Qty) Or Qty = "" Then
                If bImportata Then bImportata = False
                Grid7.Cell(lRowMixes, 4).Text = Qty
                Grid7.Cell(lRowMixes, 4).Alignment = cellCenterCenter
                
                Call UpdateValue(True)
                
                Call AddComponentTheorethicalWeight(Grid3, SelectedMixCode, IndexRecipe, indexMix)
            End If
            
        End If
End Sub


Private Sub ColoraRiga(ByVal Grid As Grid, ByVal lRow As Long)
Dim i As Integer
Dim t As Integer
Dim Count As Integer


With Grid
    .AutoRedraw = False
    Count = .Rows - 1
    If Count > 0 Then
        
        For i = 1 To Count
            For t = 0 To .Cols - 1
                If i = lRow Then
                    .Cell(i, t).ForeColor = &H40C0&
                    .Cell(i, t).FontBold = True
                Else
                    .Cell(i, t).ForeColor = vbBlack
                    .Cell(i, t).FontBold = False
                End If
            Next
        Next
    End If
    .Refresh
    .AutoRedraw = True
End With

End Sub



Private Sub impdf_Click()
lbCommand_Click
End Sub

Private Sub lbCommand_Click()
Dim sString As String

sString = nrWeek & "-" & nrYear
 
SaveSetting App.Title, strLine, sString, Nr
SaveSetting App.Title, "MaterialRequisition", "Dep.Manager", txDocument(5)
Call StampaMaterialRequisition

End Sub


Private Sub lbpdf_Click()
lbCommand_Click
End Sub


Private Sub txDocument_LostFocus(Index As Integer)
    Select Case Index
        Case 2
            SaveSetting App.Title, "Workstation", "no.department", txDocument(Index)
        
    End Select

End Sub

Private Sub txFormulation_Change(Index As Integer)
    Select Case Index
        Case 0
            ' operator...
            txDocument(1) = txFormulation(0)
        Case 1
            ' date...
            txDocument(3) = txFormulation(1)
        Case 4
            ' planning = reason of
            txDocument(4) = txFormulation(4)
        
    
    End Select
End Sub

Private Sub txFormulation_Click(Index As Integer)
Dim Answer As String
Dim Selected As String
Dim sString As String
Dim bNumber As Boolean

Selected = "RecipeForProduction"
Answer = txFormulation(Index)
sString = lbFormulation(Index)

bNumber = IIf(Index = 2, True, False)

If Index = 1 Then If Answer = "" Then Answer = FormatDataLAT(Now())
If Index = 3 Then If Answer = "" Then Answer = PreparationWeek(Now())
        
        If F_InputBox.DoShow(sString, Selected, , , , Answer, , bNumber, Me.Top) Then
        
            txFormulation(Index) = Answer
            
            Select Case Index
                Case 1
                    ' isdate?
                    If IsDate(Answer) Then
                         txFormulation(Index) = FormatDataLAT(Answer)
                    Else
                        PopupMessage 2, "Please enter a valid Date...", , True
                    End If
                Case 2
                    ' controllo se esiste già!
                    
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

If Index = 3 And Answer = "" Then Answer = FormatDateTime(Now())

'bNumber = IIf(Index = 2, True, False)

        
        If F_InputBox.DoShow(sString, Selected, , , , Answer, , bNumber, Me.Top) Then
        
            txDocument(Index) = Answer
            
            Select Case Index
            
                Case 3
                    ' isdate?
                    If IsDate(Answer) Then
                         txDocument(Index) = CDate(Answer)
                    Else
                        PopupMessage 2, "Please enter a valid Date...", , True
                    End If
            End Select
        End If
        
        




End Sub




Private Sub ucScrollAdd1_ScrollH(Value As Long)
    Form_Resize
End Sub
Private Sub PicHover_Click()
PBContainer.Top = 0
End Sub
Private Sub lblHoverClick_Click()
    PBContainer.Top = 0
    
End Sub
Private Sub imOver_Click()
PBContainer.Top = 0
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
        If F_MsgBox.DoShow("Quit Recipe for Production?") Then Unload Me
    Case 2
        ' Debug.Print PathRequisition
        ApriIlReportFolder (USER_DOCUMENTI & PathRequisition)
    Case 3
        ' Previous
         If IndexVisibleFrame >= 1 Then
            MyIndex = IndexVisibleFrame - 1
            If frInside(MyIndex).Visible = False Then
                MyIndex = IndexVisibleFrame - 3
            End If
            
            PBContainer.Top = -(frInside(MyIndex).Top - 480)
        Else
            PBContainer.Top = 0
         End If
    
    
    
    Case 4
        ' forward
        If IndexVisibleFrame < frInside.UBound Then
            MyIndex = IndexVisibleFrame + 1
            If frInside(MyIndex).Visible = False Then
                MyIndex = IndexVisibleFrame + 3
            End If
            PBContainer.Top = -(frInside(MyIndex).Top - 480)
        Else
            PBContainer.Top = 0
        End If
          
End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27
        DefaultMenu_Click 0
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

DefaultMenu(2).Visible = IIf(IndexProcedura = 1, True, False)
DefaultMenuLabel(2).Visible = DefaultMenu(2).Visible
Lab(5).Visible = DefaultMenu(2).Visible

Select Case IndexProcedura
    Case 0
        
    Case 1
        
        
        txDocument(1) = txFormulation(0)
        txDocument(3) = txFormulation(1)
        txDocument(4) = txFormulation(4)
        txDocument(0) = GetMaterialRequisitionNumber
        txDocument(2) = GetLineNumber(uRecipeForProduction.Recipes(1).Line)
        txDocument(5) = GetSetting(App.Title, "MaterialRequisition", "Dep.Manager", "Kis Laszlo")
End Select

PBFooter.ZOrder


End Function

Private Function GetMaterialRequisitionNumber() As String
Dim sString As String
With uRecipeForProduction
    nrWeek = Week(txFormulation(1)) ' recipe date
    nrYear = year(txFormulation(1))
    nrWeek = Format(nrWeek, "00")
    strLine = GetLineNumber(.Recipes(1).Line)
    sString = nrWeek & "-" & nrYear
    Nr = GetSetting(App.Title, strLine, sString, "00") + 1
    Nr = Format(Nr, "00")
    GetMaterialRequisitionNumber = strLine & nrWeek & Nr
    
End With
End Function




Private Function GetLineNumber(ByVal Line As String) As String
GetLineNumber = Mid(Line, 2, 2)
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


    bCancelUpdate = False
        
    Select Case Index
        Case 0
            ' select codes
            Call SelectCode
        Case 1
            ' select recipe
            Call SelectRecipe
        Case 2
            Call UpdateRecipeForProduction
            PBContainer.Top = -(frInside(1).Top - 480)
        Case 3
            ' update Recipes Table
            Call UpdateValue(True)
            Call AddComponentTheorethicalWeight(Grid3, SelectedRecipeCode, IndexRecipe, indexMix)
        
            PBContainer.Top = -(frInside(4).Top - 480)
        Case 4
            ' SaveReceipt
                Call SaveReceipt
        Case 5
            ' MATERIAL REQUISITION : delete record
            Dim SelectedComponent As String
            SelectedComponent = Grid6.Cell(lRowMaterialReq, 1).Text
            If SelectedComponent <> "" Then
                If F_MsgBox.DoShow("Warning : Delete Component from Table?", SelectedComponent) Then
                    Call MaterialRequisitionDeleteRecord(Grid6)
                End If
            
            End If
                
        Case 6
           ' Debug.Print PathRequisition
            ApriIlReportFolder (USER_DOCUMENTI & PathRequisition)
        Case 7
            ' reset quantity Hanna Code
            If ResetQuantityHannaCode(Grid1) Then Call UpdateValue
        Case 8
            ' cancella Hanna code
            Call DeleteHannaCodeButton
        Case 9
            PBContainer.Top = -(frInside(1).Top - 480)
            
        Case 10
            ' material requisition ALL RECIPE
            Call SetMaterialRequisitionComponents
        Case 11
            ' material requisition single Recipe
            Call SetMaterialRequisitionMixes
        Case 12
        
            Call DeleteAllButton
            
        Case 13
            Call SetMaterialRequisitionOnlyMixes
            
        Case 14
            PBContainer.Top = -(frInside(1).Top - 480)
        Case 15
             PBContainer.Top = 0
                
    End Select
End Sub


Private Function DeleteHannaCodeButton()

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

End Function
Private Function DeleteAllButton()

    If F_MsgBox.DoShow("Warning : Delete All Hanna Codes and Recipes in Tables?", "Recipe for Production") Then
    
        ' riparto da zero e nel dubbio metto il percorso TEMP di default
        
        SettingName = ""
        USER_PATH = USER_TEMP_PATH
        bImportata = False
        StartUpForm
        PBContainer.Top = 0
        PopupMessage 2, "All Recipes And Hanna Codes deleted..."
        
        
    End If


End Function

Private Sub frInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


Dim i As Integer
    For i = 0 To frCommandInside.UBound

            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 4 Or i = 11 Or i = 10 Or i = 13 Then
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
            If i = 4 Or i = 11 Or i = 10 Or i = 13 Then
                frCommandInside(i).BackColor = &H20A020
            End If
            
        Else
            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 4 Or i = 11 Or i = 10 Or i = 13 Then
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




Private Sub SelectCode()
Dim HannaCode As String
Dim rc As Boolean
    FormCodes.ZOrder
    rc = FormCodes.DoShow(HannaCode) ', Me.Top
    
    If HannaCode = "" Then Exit Sub
    
    If bImportata Then bImportata = False
    
    Call AddCodeAndRecipe(HannaCode)
   
    
    

End Sub

Private Sub AddRecipeAndHannaCodeuRecipe(ByVal RecipeCode As String)
Dim i As Integer
    If F_MsgBox.DoShow("Import Hanna Code for this recipe?", RecipeCode) Then
    
            With dbTabCode
                .filter = ""
               ' .filter = "Recipe='" & Trim(RecipeCode) & "'"
                
                .filter = "Recipe like '*" & Trim(RecipeCode) & "*'"
                
                If .EOF Then
                    Call AddCodeAndRecipe(RecipeCode, True, True)
                    PBContainer.Top = -(frInside(1).Top - 480)
   
                Else
                    .MoveFirst
                    For i = 1 To .RecordCount
                        If InStr(UCase(Trim(!Recipe)), UCase(Trim(RecipeCode))) Then
                            Call AddCodeAndRecipe(Trim(!Code), True)
                            '.filter = ""
                            '.Move i - 1
                            'GoTo cont
                        End If
                        
cont:
                        
                        If Not (.EOF) Then
                           .MoveNext
                        Else
                            Exit For
                        End If
                    Next
                
                End If
            End With
    Else
        Call AddCodeAndRecipe(RecipeCode, True, True)
    End If

    
End Sub
Private Sub SelectRecipe()
Dim RecipeCode As String
Dim rc As Boolean
Dim i As Integer
    FormRecipes.ZOrder
    rc = FormRecipes.DoShow(RecipeCode)
    If RecipeCode = "" Then Exit Sub
    
    If bImportata Then bImportata = False
    Call AddRecipeAndHannaCodeuRecipe(RecipeCode)

    

End Sub

Private Sub AddCodeAndRecipe(ByVal HannaCode As String, Optional ByVal bSalta As Boolean, Optional ByVal bSoloRecipe As Boolean)
Dim i As Integer
Dim rc As Boolean
Dim VarRecipeCode() As String



On Error GoTo ERR_ADD:

    
    If bSoloRecipe Then
       
        ReDim VarRecipeCode(0)
        VarRecipeCode(0) = HannaCode

        rc = AddRecipeInRecipeGrid2(Grid2, Grid4, VarRecipeCode(), uRecipe, 1)
        If rc = False Then Exit Sub

    Else
     
        Call AddCodeInRecipeForProductionGrid(Grid1, HannaCode, bSalta)
        Call AddRecipeInRecipeForProductionGrid(Grid1, Grid2, Grid4, uRecipe)

    End If
    
    ReDim Preserve uRecipe(Grid2.Rows - 1)
    Dim t As Integer
    Dim strMix As String
    For i = 1 To Grid2.Rows - 1
        If uRecipe(i).Code = Trim(Grid2.Cell(i, 1).Text) Then
        
        Else
            uRecipe(i).Code = Trim(Grid2.Cell(i, 1).Text)
            
            If IfRecipeExsists(uRecipe(i).Code) Then
                Call SetMyRecipeByCode(uRecipe(i).Code, uRecipe(i))
            End If
            
        End If
        'If uRecipe(i).RmxRecipe Is Nothing Then
       ' Else
            For t = 0 To UBound(uRecipe(i).RmxRecipe)
                If uRecipe(i).RmxRecipe(t).bMix Then
                    strMix = uRecipe(i).RmxRecipe(t).CHCode
                    If IfRecipeNotInGrid2(strMix, Grid2) Then
                        Call AddCodeAndRecipe(strMix, True, True)
                    End If
                End If
            Next
       ' End If
        
    Next
    
    
    
    
 

   ' uRecipe = uRecipe

    Call AddMachine(Grid5, Grid2)
    
ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_ADD:
    MsgBox "Error Recipe  :   " & uRecipe(i).Code & vbCrLf & err.Description
    Resume Next
End Sub


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







Private Sub UpdateRecipeForProduction()




    Call AddRecipeInRecipeForProductionGrid(Grid1, Grid2, Grid4, uRecipe)
    
    ReDim Preserve uRecipe(Grid2.Rows - 1)
    
    Call UpdateValue
    
    
    
End Sub


Private Function UpdateValue(Optional ByVal SeSoloRicetta As Boolean)
Dim i As Integer
Dim t As Integer
Dim rcMin As Boolean
Dim rcMax As Boolean
Dim RowsHannaCount As Integer
Dim RowsRecipeCount As Integer
Dim HannaQuantity As Double
Dim HannaVolume As Double
Dim UmHannaQuantity As String

Dim UmTotalWeight As String

Dim bUmMassa As Boolean
Dim bUmMultipleMassa As Boolean
Dim bUmHannaMassa As Boolean

Dim NowRecipe As Integer


On Error GoTo ERR_UPDATE



If bCancelUpdate Then Exit Function

UmTotalWeight = "Kg"

NowRecipe = 1
RowsHannaCount = Grid1.Rows - 1
RowsRecipeCount = Grid2.Rows - 1

If SeSoloRicetta Then
    NowRecipe = IndexRecipe
    RowsRecipeCount = IndexRecipe
End If

If RowsHannaCount > 0 Or RowsRecipeCount > 0 Then


    For i = NowRecipe To RowsRecipeCount
    
        uRecipe(i).TotalWeightKg = 0
        uRecipe(i).TotalMultiple = 0
        uRecipe(i).TotalWeightL = 0
        uRecipe(i).MultipleMassa = 0
        
    
        If uRecipe(i).Code = "" Then GoTo NullCode:
        ' controllo se esiste la ricetta in uRecipe....
        
        If uRecipe(i).bUpdated = False Then Call SetMyRecipeByCode(uRecipe(i).Code, uRecipe(i))
            
        If uRecipe(i).bUpdated = False Then
            uRecipe(i).bUpdated = SetRmxRecipeByRecipeCode(uRecipe(i).Code, uRecipe(i).RmxRecipe, False)
        End If
            
        If uRecipe(i).bUpdated = False Then GoTo NullCode

        '-----------------------------------------
        ' definisco massa o volume nei calcoli
        '------------------------------------------
        bUmMassa = SetbUmMassa(uRecipe(i).UmMinQty)
        bUmMultipleMassa = SetbUmMassa(uRecipe(i).UmMultiple)
        
        
            
        If IsNumeric(Grid2.Cell(i, 4).Text) Then
          
            uRecipe(i).Density = IIf(Grid2.Cell(i, 8).Text = "", 1, Grid2.Cell(i, 8).Text)
            
            uRecipe(i).MultipleToProduce = (Grid2.Cell(i, 4).Text)
            uRecipe(i).TotalRecipe = (Grid2.Cell(i, 6).Text)
            uRecipe(i).Multiple = CDbl(Grid2.Cell(i, 12).Text)
            uRecipe(i).UmMultiple = (Grid2.Cell(i, 5).Text)
         
            
            If bUmMultipleMassa Then
                uRecipe(i).MultipleMassa = uRecipe(i).MultipleToProduce * uRecipe(i).Multiple '/ uRecipe(i).Density
            Else
                uRecipe(i).MultipleMassa = uRecipe(i).MultipleToProduce * uRecipe(i).Multiple * uRecipe(i).Density
            End If
            
            
            
            uRecipe(i).TotalWeightKg = uRecipe(i).TotalWeightKg + uRecipe(i).MultipleMassa * Um(uRecipe(i).UmMultiple)
            uRecipe(i).TotalMultiple = uRecipe(i).TotalMultiple + uRecipe(i).MultipleToProduce
            uRecipe(i).TotalWeightL = uRecipe(i).TotalWeightKg / uRecipe(i).Density
            
        End If
        
ContHanna:

        For t = 1 To RowsHannaCount
            If Not (IsNumeric(Grid1.Cell(t, 4).Text)) Or Not (IsNumeric(Grid1.Cell(t, 6).Text)) Then GoTo cont
            Debug.Print Grid1.Cell(t, 7).Text
            If InStr(uRecipe(i).Code, Grid1.Cell(t, 7).Text) Then
                
                If IsNumeric(Grid1.Cell(t, 6).Text) Then
        
                    HannaQuantity = Grid1.Cell(t, 6).Text
                    UmHannaQuantity = IIf(Grid1.Cell(t, 5).Text = "", uRecipe(i).UmMultiple, Grid1.Cell(t, 5).Text)
                    HannaVolume = Grid1.Cell(t, 4).Text
                    HannaQuantity = HannaQuantity * Um(UmHannaQuantity)
                    
                    bUmHannaMassa = SetbUmMassa(UmHannaQuantity)
                    
                    HannaQuantity = (HannaQuantity * HannaVolume) * IIf(bUmHannaMassa, 1, uRecipe(i).Density)
                  
                    uRecipe(i).TotalWeightKg = uRecipe(i).TotalWeightKg + HannaQuantity
                    uRecipe(i).TotalWeightL = uRecipe(i).TotalWeightKg / uRecipe(i).Density
                    
                    If uRecipe(i).Multiple = 0 Then
                        uRecipe(i).TotalMultiple = 0
                    Else
                        uRecipe(i).TotalMultiple = uRecipe(i).TotalMultiple + ((HannaQuantity / Um(uRecipe(i).UmMultiple) / IIf(bUmHannaMassa, 1, uRecipe(i).Density) / uRecipe(i).Multiple))
                    End If
                    
                End If
            
            End If
cont:
        Next
        
        ' theoretical weight Mixes
        Dim X As Integer
        Dim MultipleToProduceMix As Double
        Dim MultipleToProduceMassa As Double
        Dim MultipleMix As Double
        Dim UmMultipleMix As String
        Dim MixDensity As Double
        
        
        If uRecipe(i).bUpdated = False Then GoTo NullCode
        
        For X = LBound(uRecipe(i).RmxRecipe) To UBound(uRecipe(i).RmxRecipe)
            With uRecipe(i).RmxRecipe(X)
            
                Debug.Print .CHCode
               
                If .bMix And isMixInGrid7(Grid7, uRecipe(i).RmxRecipe(X).CHCode) Then
                
                  If .RecipeCode <> uRecipe(i).Code Then GoTo salta
                    'If bUmMassa Then
                        .TheoreticalWeight = (.Perc * uRecipe(i).TotalWeightKg / Um(UmTotalWeight)) / 100
                        .UmTheoreticalWeight = UmTotalWeight
                    'Else
                       ' .TheoreticalWeight = (.Perc * uRecipe(i).TotalWeightKg * uRecipe(i).Density / Um(UmTotalWeight)) / 100
                        '.UmTheoreticalWeight = UmTotalWeight
                  '  End If
                    

                    
                    ' add multiple Grid7
                    With Grid7
        
                        If .Rows > 1 And X < .Rows - 1 Then
                        
                            uRecipe(i).RmxRecipe(X).MultipleInCell = Trim((.Cell(X + 1, 4).Text))
                           
                            If .Cell(X + 1, 4).Text <> "" Then
                            
                                'uRecipe(i).RmxRecipe(X).TheoreticalWeight = 0
                                MultipleToProduceMix = IIf(IsNumeric(.Cell(X + 1, 4).Text), CDbl(.Cell(X + 1, 4).Text), 0)
                                MultipleMix = IIf(IsNumeric(.Cell(X + 1, 13).Text), CDbl(.Cell(X + 1, 13).Text), 0)
                                UmMultipleMix = IIf((.Cell(X + 1, 14).Text) = "", uRecipe(i).UmMultiple, (.Cell(X + 1, 14).Text))
                                MixDensity = IIf(IsNumeric(.Cell(X + 1, 9).Text), CDbl(.Cell(X + 1, 9).Text), 1)
                                
                                If MultipleToProduceMix > 0 Then
                                    
                                    If bUmMassa Then
                                        ' in grammi
                                        MultipleToProduceMassa = (MultipleToProduceMix * MultipleMix / IIf(bUmMultipleMassa, 1, MixDensity)) * Um(UmMultipleMix)
                                    Else
                                    
                                        MultipleToProduceMassa = (MultipleToProduceMix * MultipleMix * IIf(bUmMultipleMassa, 1, MixDensity)) * Um(UmMultipleMix)
                                    End If
                                    ' in kili
                                    uRecipe(i).RmxRecipe(X).TheoreticalWeight = (uRecipe(i).RmxRecipe(X).TheoreticalWeight + MultipleToProduceMassa / Um(uRecipe(i).RmxRecipe(X).UmTheoreticalWeight))
                                End If
                            End If
                            
                        End If
                    End With
                    
                  .TheoreticalWeight = (.TheoreticalWeight)
                    
                    Call CheckMinMaxQuantity(uRecipe(i).RmxRecipe(X))


                Else
                    If .TheoreticalWeight = 0 Then .TheoreticalWeight = FormatNumber(uRecipe(i).TotalWeightKg, 3)
                End If
                
            End With
        
           
salta:
        Next
        
        Dim MinQtyKg As Double
        Dim MaxQtyKg As Double

        If SetbUmMassa(uRecipe(i).UmMax) Then
            MinQtyKg = uRecipe(i).MinQty * Um(uRecipe(i).UmMax)
            MaxQtyKg = uRecipe(i).MaxQty * Um(uRecipe(i).UmMax)
        Else
            MinQtyKg = uRecipe(i).MinQty * Um(uRecipe(i).UmMax) * uRecipe(i).Density
            MaxQtyKg = uRecipe(i).MaxQty * Um(uRecipe(i).UmMax) * uRecipe(i).Density
        End If
            
         
        ' check Min e MAx
 
            rcMin = IIf(MinQtyKg <= uRecipe(i).TotalWeightKg, True, False)
            rcMax = IIf(MaxQtyKg >= uRecipe(i).TotalWeightKg, rcMin, False)
     
        
        PicMin.BackColor = IIf(rcMin, &H8000&, &HC0&)
        PicMax.BackColor = IIf(rcMax, &H8000&, &HC0&)
        
        Dim uDecimali As Integer
        
        uDecimali = IIf((uRecipe(i).TotalWeightKg / Um(UmTotalWeight)) > 1, 3, 6)
        uRecipe(i).TotalWeightKg = (uRecipe(i).TotalWeightKg / Um(UmTotalWeight))
        uRecipe(i).TotalWeightL = (uRecipe(i).TotalWeightL / Um(UmTotalWeight))
              
        If Grid4.Rows > i Then
            Grid4.Cell(i, 3).Text = PadString(uRecipe(i).TotalWeightKg) & "  "
            Grid4.Cell(i, 4).Text = PadString(uRecipe(i).TotalWeightL) & "  "
            Grid4.Cell(i, 5).Text = FormatNumber(uRecipe(i).TotalMultiple, 1) & "  "
            Grid4.Cell(i, 8).BackColor = PicMin.BackColor
            Grid4.Cell(i, 11).BackColor = PicMax.BackColor
            
            Grid4.Refresh
            
        End If
        
        If bUmMultipleMassa Then
            Grid2.Cell(i, 6).Text = PadString(uRecipe(i).TotalWeightKg) & " Kg"
            Grid2.Cell(i, 6).Alignment = cellCenterCenter
        Else
            Grid2.Cell(i, 6).Text = PadString(uRecipe(i).TotalWeightL) & " L"
            Grid2.Cell(i, 6).Alignment = cellCenterCenter
        End If

        Call SetEstimatedTime(i)
        
NullCode:

    Next

    Call AddTheorethicalWeight(Grid7, NowRecipe, SeSoloRicetta)

End If

ERR_END:
    On Error GoTo 0
    Exit Function
ERR_UPDATE:
    MessageInfoTime = 2000
    PopupMessage 2, err.Description & vbCrLf & "Warning : Bad Recipe...Check Recipe Settings ", , True, uRecipe(i).Code
    GoTo NullCode

End Function



Private Sub CheckMinMaxQuantity(ByRef uRMxRecipe As RmxRecipe)

Dim bUmMassa As Boolean
Dim rcMin As Boolean
Dim rcMax As Boolean
Dim UmTotalWeight As String
Dim PicMinBackColor As OLE_COLOR
Dim PicMaxBackColor As OLE_COLOR
Dim MinQtyKg As Double
Dim MaxQtyKg As Double

Dim i As Integer
Dim TotalWeightKg As Double
Dim TotalWeightL As Double
Dim MultipleToProduce As Double
Dim strMixCode As String


    If uRMxRecipe.TheoreticalWeight = 0 Then Exit Sub
    
    strMixCode = uRMxRecipe.CHCode
    
    
    If uRMxRecipe.UmMax = "" And uRMxRecipe.MinQty = 0 Then
        ' probabilmente è un Mix annidato
        Call SetMyRecipeMixByCode(strMixCode, uRMxRecipe)
    
    End If
    
    bUmMassa = SetbUmMassa(uRMxRecipe.UmMultiple)
    
    If uRMxRecipe.Density = 0 Then uRMxRecipe.Density = 1
    If uRMxRecipe.Multiple = 0 Then uRMxRecipe.Multiple = 1
    If bUmMassa Then
        uRMxRecipe.MultipleToProduce = Int((((uRMxRecipe.TheoreticalWeight * 1000)) / Um(uRMxRecipe.UmMultiple)) / uRMxRecipe.Multiple)
    Else
        uRMxRecipe.MultipleToProduce = Int((((uRMxRecipe.TheoreticalWeight * 1000) / uRMxRecipe.Density) / Um(uRMxRecipe.UmMultiple)) / uRMxRecipe.Multiple)
    End If
    
    TotalWeightKg = 0
    TotalWeightL = 0
    MultipleToProduce = 0
    
    Call AddTotalWeightMixesAllRecipes(uRecipe(), strMixCode, TotalWeightKg, TotalWeightL, MultipleToProduce)
    
    
    UmTotalWeight = "kg"

     
    If bUmMassa Then
        MultipleToProduce = Int((((TotalWeightKg * 1000)) / Um(uRMxRecipe.UmMultiple)) / uRMxRecipe.Multiple)
    Else
        MultipleToProduce = Int((((TotalWeightKg * 1000) / uRMxRecipe.Density) / Um(uRMxRecipe.UmMultiple)) / uRMxRecipe.Multiple)
    End If
      
    
      If uRMxRecipe.Density = 0 Then uRMxRecipe.Density = 1
    
       
       ' SIAMO IN GRAMMI!
      If uRMxRecipe.Multiple > 0 Then
      
          If SetbUmMassa(uRMxRecipe.UmMax) Then
              MinQtyKg = uRMxRecipe.MinQty * Um(uRMxRecipe.UmMax)
              MaxQtyKg = uRMxRecipe.MaxQty * Um(uRMxRecipe.UmMax)
          Else
              MinQtyKg = uRMxRecipe.MinQty * Um(uRMxRecipe.UmMax) * uRMxRecipe.Density
              MaxQtyKg = uRMxRecipe.MaxQty * Um(uRMxRecipe.UmMax) * uRMxRecipe.Density
          End If
          
              rcMin = IIf(MinQtyKg <= TotalWeightKg * 1000, True, False)
              rcMax = IIf(MaxQtyKg >= TotalWeightKg * 1000, rcMin, False)
      End If
      
      PicMinBackColor = IIf(rcMin, &H8000&, &HC0&)
      PicMaxBackColor = IIf(rcMax, &H8000&, &HC0&)
        
               
      uRMxRecipe.TotalWeightKg = (uRMxRecipe.TheoreticalWeight)
      uRMxRecipe.TotalWeightL = (uRMxRecipe.TheoreticalWeight / IIf(uRMxRecipe.Density = 0, 1, uRMxRecipe.Density))


    With Grid4
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .Cell(i, 1).Text = strMixCode Then
                    .Cell(i, 3).Text = PadString(TotalWeightKg) & "  "
                    .Cell(i, 4).Text = PadString(TotalWeightL) & "  "
                    .Cell(i, 5).Text = FormatNumber(MultipleToProduce, 1) & "  "
                    .Cell(i, 8).BackColor = PicMinBackColor
                    .Cell(i, 11).BackColor = PicMaxBackColor

                    TotalWeightKg = Replace(LCase(TotalWeightKg), "kg", "")
                  
                    
                    'uRecipe(i).TotalWeightKg = TotalWeightKg
                End If
            Next
        End If
        .Refresh
    End With

End Sub



Public Function AddMachine(ByVal Grid5 As Grid, ByVal Grid2 As Grid)
Dim i As Integer

If Grid2.Rows > 1 Then
    With Grid5
        .Column(0).Width = 40
        .Rows = 1
        .ReadOnly = False
        For i = 1 To Grid2.Rows - 1
            .AddItem "", False
            .Cell(.Rows - 1, 1).Text = uRecipe(i).Code
            .Cell(.Rows - 1, 2).Text = uRecipe(i).Line
            If .Cell(.Rows - 1, 3).Text <> "" Then
               Call SetEstimatedTime(i)
            End If
            .Cell(.Rows - 1, 3).BackColor = vbColorIns
            .Column(1).Locked = True
            .Column(2).Locked = True
            .Column(3).Locked = False
            .Column(4).Locked = True
            .Column(5).Locked = True
            .Column(6).Locked = True

        Next
        .Refresh
    End With
End If

End Function

Private Function SetEstimatedTime(ByVal i As Integer)
Dim t As Integer
Dim ErtTimeH As Double
Dim Production As Double
Dim Speed As String
On Error GoTo ERR_SET
    If i = 0 Then Exit Function
    If i > Grid5.Rows - 1 Then Exit Function
    With Grid5
    
        ' pezzi prodotti per ogni ricetta?
        
        Production = uRecipe(i).TotalMultiple
        
        ' oppure pezzi prodotti dei codici hanna???
        
        'Production=
        
        
       
       Speed = .Cell(i, 5).Text
       
       
       
        For t = LBound(ProductionWay) To UBound(ProductionWay)
           
            If ProductionWay(t).Production = "" Then GoTo cont
            If InStr(.Cell(i, 3).Text, ProductionWay(t).Production) And InStr(.Cell(i, 1).Text, uRecipe(i).Code) Then
            
                .Column(3).AutoFit
                ProductionWay(t).Speed = CDbl(Speed)
                 
                If ProductionWay(t).Speed = 0 Then
                    If F_InputBox.DoShow("Please Enter Speed", ProductionWay(t).Production & " : Evaluated Machine Speed", , , , Speed, , True) Then
                        ProductionWay(t).Speed = CDbl(Speed)
                    Else
                        
                    End If
                End If
               
                .Cell(i, 5).Text = ProductionWay(t).Speed
                .Cell(i, 4).Text = ProductionWay(t).Head
                If ProductionWay(t).Speed > 0 Then
                    ErtTimeH = Production / ((ProductionWay(t).Speed / ProductionWay(t).Head) * 60)
                    .Cell(i, 6).Text = FormatNumber(ErtTimeH, 2)
                    .Cell(i, 7).Text = FormatNumber(ErtTimeH / 7, 2)
                End If
                
                 
                 
                uRecipe(i).ProductionWay.Speed = ProductionWay(t).Speed
                uRecipe(i).ProductionWay.Head = ProductionWay(t).Head
                uRecipe(i).ProductionWay.EsttimeD = FormatNumber(ErtTimeH / 7, 2)
                uRecipe(i).ProductionWay.EstTimeH = ErtTimeH
                uRecipe(i).ProductionWay.Production = ProductionWay(t).Production
                uRecipe(i).ProductionWay.Line = ProductionWay(t).Line
                
                GoTo ERR_END
                
            End If
cont:
        Next
        
    End With
ERR_END:
  On Error GoTo 0
  Exit Function
ERR_SET:
  Debug.Print err.Description
  Resume Next
                '.Cell(0, 1).Text = "Recipe"
                '.Cell(0, 2).Text = "Production speed ( pcs/min )"
                '.Cell(0, 3).Text = "Estimated time machine ( h )"
                '.Cell(0, 4).Text = "Estimated time machine ( d )"
End Function



Private Function AddTheorethicalWeight(ByVal Grid As Grid, Optional ByVal IndexRicetta As Integer, Optional ByVal bSeSoloRicetta As Boolean)
Dim i As Integer
Dim X As Integer
Dim t As Integer
Dim TableCode As String
Dim IndexRecipe As Integer
Dim MaxRecipes As Integer

If bSeSoloRicetta Then
    IndexRecipe = IndexRicetta
    MaxRecipes = IndexRicetta
Else
    IndexRecipe = 1
    MaxRecipes = Grid2.Rows - 1
End If

On Error GoTo ERR_ADD
For i = IndexRecipe To MaxRecipes
    
    If uRecipe(i).bUpdated And uRecipe(i).bHide = False Then
        For X = 0 To Grid.Rows - 2
            For t = 0 To UBound(uRecipe(i).RmxRecipe)
                
                TableCode = Grid.Cell(X + 1, 1).Text
                    With uRecipe(i).RmxRecipe(t)
                    If .bMix And .CHCode = TableCode And uRecipe(i).Code = SelectedRecipeCode Then
                        Grid.Cell(X + 1, 7).Text = PadString(.TheoreticalWeight)
                        Grid.Cell(X + 1, 8).Text = .UmTheoreticalWeight
                        Grid.Refresh
                        GoTo contX
                    End If
                End With
                
            Next
contX:
        Next
    End If
cont:
Next

ERR_END:
    On Error GoTo 0
    Exit Function
ERR_ADD:
    GoTo ERR_END
End Function
Private Function AddComponentTheorethicalWeight(ByVal Grid As Grid, ByVal Code As String, ByVal IndexRecipe As Integer, ByVal indexMix As Integer)
Dim i As Integer
Dim X As Integer
Dim t As Integer
Dim z As Integer
Dim TableCode As String
Dim IndexComponent As Integer
Dim Component() As RmxRecipe
Dim totalWeighCount As Double

Dim TotalWeightRecipe As Double
Dim Umth As String
Dim RecipeCode As String
Dim rCode As String
Dim haveMixes As String

Dim Count As Integer

    If Grid.Rows = 1 Then Exit Function
    If uRecipe(IndexRecipe).bUpdated = True Then
    totalWeighCount = 0
        For i = 1 To Grid.Rows - 1
            ReDim Component(Grid.Rows - 1)
            RecipeCode = Code '.CHCode
            
            If Grid.Cell(i, 1).Text = "" Then Exit For
            
            Component(i).CHCode = Grid.Cell(i, 1).Text
            
            
            
            With dbTabRMxRecipe
                .filter = ""
                .filter = "RecipeCode='" & RecipeCode & "' and CHCode='" & Component(i).CHCode & "'"
                If .EOF Then
                Else
                
                     
                
                       ' If uRecipe(IndexRecipe).RmxRecipe(indexMix).bMix And Not (uRecipe(IndexRecipe).bIsMix) Then
                        
                        If Grid7.Rows > 1 And Not (uRecipe(IndexRecipe).bIsMix) Then
                        
                            If uRecipe(IndexRecipe).RmxRecipe(indexMix).CHCode = Component(i).CHCode Then
                                
                                Umth = uRecipe(IndexRecipe).RmxRecipe(indexMix).UmTheoreticalWeight
                                TotalWeightRecipe = uRecipe(IndexRecipe).RmxRecipe(indexMix).TheoreticalWeight * Um(Umth)
                            Else
                                '---------------------------------------------------------------------------
                                ' attenzione potrebbe essere sbagliato per qualche tipologia di ricetta...
                                ' ma ho dovuto correggerla ( 31/05/21 ) per sistemare un problema sulle
                                ' ricette EPA e CP-B148---
                                '---------------------------------------------------------------------------
                                Umth = uRecipe(IndexRecipe).RmxRecipe(indexMix).UmTheoreticalWeight
                                Debug.Print Component(i).CHCode
                                TotalWeightRecipe = uRecipe(IndexRecipe).RmxRecipe(indexMix).TheoreticalWeight * Um(Umth)
                            End If
                        
                        Else
                            
                            Umth = "kg"
                            TotalWeightRecipe = uRecipe(IndexRecipe).TotalWeightKg * Um(Umth)
                        
                        End If
                        
                        
                        Component(i).Perc = CheckDot(IIf(IsNull(dbTabRMxRecipe!Perc), 0, (dbTabRMxRecipe!Perc)))
                        Component(i).TheoreticalWeight = TotalWeightRecipe * Component(i).Perc / 100
                        Grid.Cell(i, 7).Text = PadString(Component(i).TheoreticalWeight)
                        
                        totalWeighCount = totalWeighCount + Component(i).TheoreticalWeight
                        
                        Grid.Cell(i, 8).Text = "g"
                        Grid.Refresh
                        
                        Count = 0
riprova:
                             Count = Count + 1
                            If uRecipe(IndexRecipe).RmxRecipe(indexMix).bMix And uRecipe(IndexRecipe).RmxRecipe(indexMix).RecipeCode = uRecipe(IndexRecipe).Code Then
                                rCode = uRecipe(IndexRecipe).RmxRecipe(indexMix).CHCode
                                
                            Else
                                rCode = uRecipe(IndexRecipe).Code
                                
                            End If
                 
                    
                            IndexComponent = CheckRmxRecipeInRecipe(uRecipe(IndexRecipe).RmxRecipe, Component(i).CHCode, uRecipe(IndexRecipe).Code, uRecipe(IndexRecipe).bHaveMixes, RecipeCode)
                            If IndexComponent = -1 Then
                                ' non c'è
                               
                                Call SetRmxRecipeByRecipeCode(rCode, uRecipe(IndexRecipe).RmxRecipe, False, UBound(uRecipe(IndexRecipe).RmxRecipe))
                                If Count > 1000 Then
                                    MessageInfoTime = 2000
                                    PopupMessage 2, "Problem Recipe : " & rCode
                                Else
                                    GoTo riprova
                                End If
                            Else
                                 Debug.Print "Component " & IndexComponent & " Theoretical ----" & uRecipe(IndexRecipe).RmxRecipe(IndexComponent).CHCode & " == " & Component(i).TheoreticalWeight
                                uRecipe(IndexRecipe).RmxRecipe(IndexComponent).TheoreticalWeight = Component(i).TheoreticalWeight
                                uRecipe(IndexRecipe).RmxRecipe(IndexComponent).UmTheoreticalWeight = "g"
                                If uRecipe(IndexRecipe).RmxRecipe(IndexComponent).bMix Then
                                    ' sono nei component ma ho un Mix per cui voglio sapere quanto pesano i suoi componenti
                                     
                                     Debug.Print
                                     Debug.Print "sono nei component ma ho un Mix per cui voglio sapere quanto pesano i suoi componenti"
                                     Debug.Print uRecipe(IndexRecipe).RmxRecipe(IndexComponent).CHCode
                                     Debug.Print
                                     Dim TheoreticalWeight As Double
                                     Dim OriginMix As String
                                     Dim Perc As Double
                                     
                                     TheoreticalWeight = uRecipe(IndexRecipe).RmxRecipe(IndexComponent).TheoreticalWeight
                                     OriginMix = uRecipe(IndexRecipe).RmxRecipe(IndexComponent).CHCode
                                     For z = IndexComponent + 1 To UBound(uRecipe(IndexRecipe).RmxRecipe)
                                            If uRecipe(IndexRecipe).RmxRecipe(i).RecipeCode = OriginMix Then
                                                  Perc = uRecipe(IndexRecipe).RmxRecipe(z).Perc
                                                  uRecipe(IndexRecipe).RmxRecipe(z).TheoreticalWeight = TheoreticalWeight * Perc / 100
                                                  uRecipe(IndexRecipe).RmxRecipe(z).UmTheoreticalWeight = "g"
                                            
                                            
                                            End If
                                        
                                     
                                     Next
                                     
                                     
                                     
                                     
                                End If
                            End If
                            
                        
                        
                    End If
                    
                    
            
            End With

           
        Next
        '------------------------
        '           TOTALS
        '------------------------
        Call AddTotals(Grid3, i, totalWeighCount)
       
     
    End If
    
    ' ho i pesi dei component, ora controllo  in tutte le ricette il peso dei Mix e li sommo!!!
    ' devo controllare anche all'interno dei mix se trovo altri mix da sommare!!!!!
    ' lo faccio qui perchè ho appena calcolato i component di un Mix! quindi rifacio il calcolo e il check min max
    
    
    
    For i = 1 To UBound(uRecipe)
        If uRecipe(i).bHide = False And uRecipe(i).bUpdated And uRecipe(i).TotalWeightKg > 0 Then
            For X = 0 To UBound(uRecipe(i).RmxRecipe)
                If uRecipe(i).RmxRecipe(X).bMix Then
                    Debug.Print uRecipe(i).RmxRecipe(X).CHCode
                    Call CheckMinMaxQuantity(uRecipe(i).RmxRecipe(X))
                End If
            Next
        End If
    Next
    
    
End Function



Public Function AddTotals(ByVal Grid3 As Grid, ByVal i As Integer, ByVal totalWeighCount As Double)
'------------------------
'           TOTALS
'------------------------
Dim bAddRow As Boolean
 With Grid3
 
        bAddRow = False
        
        If i = .Rows Then
            .AddItem "", False
            .AddItem "", False
            bAddRow = True
        End If

        totalWeighCount = totalWeighCount / 1000

        .Cell(i + 1, 6).Text = "Totals"
        .Cell(i + 1, 6).Alignment = cellLeftCenter
        .Cell(i + 1, 7).Text = PadString(totalWeighCount)
        .Cell(i + 1, 7).Alignment = cellCenterCenter
        .Cell(i + 1, 8).Text = "Kg"
        
        .Cell(i + 1, 6).FontBold = True
        .Cell(i + 1, 7).FontBold = True
        .Cell(i + 1, 8).FontBold = True
        
        .Cell(i + 1, 6).ForeColor = vbColorDarkFont ' &H644603
        .Cell(i + 1, 7).ForeColor = vbColorDarkFont '&H644603
        .Cell(i + 1, 7).BackColor = vbColorResults
        .Cell(i + 1, 8).ForeColor = vbColorDarkFont '&H644603
        
        If uRecipe(IndexRecipe).bUmMassa Then
        Else
            
            If bAddRow Then
                .AddItem "", False
            End If
        .Cell(.Rows - 1, 6).Text = "Totals"
        .Cell(.Rows - 1, 6).Alignment = cellLeftCenter
        .Cell(.Rows - 1, 7).Text = PadString(totalWeighCount / uRecipe(IndexRecipe).Density)
        .Cell(.Rows - 1, 7).Alignment = cellCenterCenter
        .Cell(.Rows - 1, 8).Text = "L"
        
        .Cell(.Rows - 1, 6).FontBold = True
        .Cell(.Rows - 1, 7).FontBold = True
        .Cell(.Rows - 1, 8).FontBold = True
        
        .Cell(.Rows - 1, 6).ForeColor = vbColorDarkFont '&H644603
        .Cell(.Rows - 1, 7).ForeColor = vbColorDarkFont '&H644603
        .Cell(.Rows - 1, 7).BackColor = vbColorResults
        .Cell(.Rows - 1, 8).ForeColor = vbColorDarkFont '&H644603
        
        End If
        .Refresh
        .AutoRedraw = True
  End With
End Function










'-----------------------------------------------------------------------------------------------
'
'
'                                   SaveRecipeFor Production
'
'
'           FileName = SetSettingName
'
'           dbTabReceiptForProduction
'
'
'-----------------------------------------------------------------------------------------------
Private Function CheckSettingName() As Boolean
Dim rc As Boolean
On Error GoTo ERR_CHECK

    rc = True
    
    If SettingName = "" Then
        rc = False
    Else
    
        If FileExists(USER_TEMP_PATH & SettingName) Then
            rc = True ' (F_MsgBox.DoShow("This Recipe already Exsists." & vbCrLf & "Save changes?", "Recipe For production", True, "Save", "Exit"))
        
        ElseIf FileExists(USER_DATA_PATH & SettingName) Then
            rc = False
            PopupMessage 2, "This Recipe already Exsists." & vbCrLf & "Please change #PrepWeek or Planned Preparation", , True
        End If

    
    
    End If


ERR_END:
    On Error GoTo 0
    CheckSettingName = rc
    Exit Function
ERR_CHECK:
    rc = False
    Resume ERR_END
    

End Function



Private Sub SaveReceipt()
Dim i As Integer
Dim rc As Boolean
Dim NewName As String

If Grid2.Rows = 1 Then
    MessageInfoTime = 2500
    PBContainer.Top = 0
    PopupMessage 2, "Please select Recipe or Hanna Codes first...", , True
    Exit Sub
End If


With uRecipeForProduction
   
    .bAllMixes = IfRecipeForProductionHasAllMixes(uRecipe())
    .RecipeCount = Grid2.Rows - 1
    .Recipes = uRecipe
    If IsDate(txFormulation(1)) Then
        .DateRecipe = CDate(txFormulation(1))
    End If
    .Note = txFormulation(5)
    .PlannedPrepWeek = txFormulation(3)
    .PlanningReference = txFormulation(4)
    .numPrepWeek = txFormulation(2)
    .RecipeBy = txFormulation(0)

End With



rc = CheckRfPBeforeSave(uRecipeForProduction)

If rc = False Then
    
    PopupMessage 2, "Please change #Prep Week or Planned preparation , Recipe already in use!", , True
    
    Exit Sub
End If



NewName = FormatNomeFile(Trim(uRecipe(1).Code) & "." & Trim(uRecipe(1).Line) & "." & txFormulation(1) & "." & txFormulation(2) & "." & txFormulation(3)) & "." & USER_ESTENSIONE_RFP

If FileExists(USER_PATH & NewName) Then
    If F_MsgBox.DoShow("This Recipe already Exsists." & vbCrLf & "Save changes?", "Recipe For production", True, "Save", "Exit") Then
    Else
        Exit Sub
    End If

End If



CloseSettingDataFile

If SettingName = "" Then

    SettingName = NewName
    
ElseIf SettingName <> NewName Then
    rc = True
    
    If FileExists(USER_PATH & SettingName) Then
        
        FileCopy USER_PATH & SettingName, USER_PATH & NewName
        Kill USER_PATH & SettingName
    End If
    
    
    DoEvents
     SettingName = NewName
     
     
    GoTo cont:
End If



rc = ChecktxFormulation
If rc Then
    
    rc = CheckSettingName
    
End If
cont:
If Grid2.Rows < 2 Then
        PopupMessage 2, "Warning : Please add a valid Recipe and quantity to produce...", , True, "Recip for Production"

        PBContainer.Top = 0

    Exit Sub
    
End If

If rc = False Then
  
  Exit Sub
    
End If


lbWait.Caption = "Wait : Saving Data..."
lbWait.Visible = True


uRecipeForProduction = uRecipeForProductionClean




Call SetAllVariables

With uRecipeForProduction
    .fileNameRecForProd = SettingName
    .bAllMixes = IfRecipeForProductionHasAllMixes(uRecipe())
    .RecipeCount = Grid2.Rows - 1
    .Recipes = uRecipe
    .bOpen = True
    .DateRecipe = CDate(txFormulation(1))
    .Note = txFormulation(5)
    .PlannedPrepWeek = txFormulation(3)
    .PlanningReference = txFormulation(4)
    .numPrepWeek = txFormulation(2)
    .RecipeBy = txFormulation(0)
    
    
    If .RecipeCount = 0 Then Exit Sub
    
    .TotalCount = Grid4.Rows - 1
    
    Call SetTotalsFromGrid(.TotalGrid, Grid4)
    
    .PackagingCount = Grid5.Rows - 1
        
    
    Call SetPackagingFromGrid(.Packaging, Grid5)
    
End With





    rc = ReceiptSaveSetting(uRecipeForProduction, SettingName)
    
    If rc Then
        
        rc = SaveRecipeForProductionInDatabase
    
    End If
    
    uRecipeForProduction.bSaved = rc
    
    If rc Then
        ' salvata correttamente
        DoEvents
        PopupMessage 2, "Recipt correctly saved...", , , "Recip for Production"
        DoEvents
        frCommandInside(10).Visible = True
        frCommandInside(11).Visible = True
        frCommandInside(13).Visible = IfMixesInTotalGrid(Grid4)
        PicMenu(1).Visible = True
        PBContainer.Top = 0 '-(frInside(1).Top - 460)
        DoEvents
    Else
        ' errore
        PopupMessage 2, "Warning : Save Error ", , True, "Recip for Production"
    
    End If

    lbWait.Visible = False

    lbWait.Caption = "Wait : Loading Data..."
End Sub

Private Function IfMixesInTotalGrid(ByVal Grid4 As Grid) As Boolean
Dim i As Integer
IfMixesInTotalGrid = False
    With Grid4
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                If .RowHeight(i) > 0 Then
                    If .Cell(i, 15).Text <> "" Then
                            If .Cell(i, 15).Text = True Then
                                    IfMixesInTotalGrid = True
                                    Exit For
                            End If
                    End If
                End If
            Next
        End If
    End With
End Function

Private Sub SetSettingName()
' LINE+DATERECIPE+PREPARATIONWEEK+PLANNEDPREPARATION
SettingName = FormatNomeFile(Trim(uRecipe(1).Code) & "." & Trim(uRecipe(1).Line) & "." & txFormulation(1) & "." & txFormulation(2) & "." & txFormulation(3)) & "." & USER_ESTENSIONE_RFP

End Sub


Private Sub SetAllVariables()
Dim i As Integer

On Error GoTo VarSet
    With uRecipeForProduction
    
        Call GetHannaCodePerGrid(Grid1, .HannaCodes())
    
        .HannaCodesCount = Grid1.Rows - 1
        For i = 1 To Grid2.Rows - 1
            With uRecipe(i)
            
                .HannaCodesCount = UBound(.HannaCodes)
                .MaterialRequisitionCount = UBound(.MaterialRequisition)
                .RmxRecipeCount = UBound(.RmxRecipe)
            End With
        Next
        '-----------------------------------------------------------
        ' Totals Mixes
        '-----------------------------------------------------------
        Call SetTotalMixes
            '-----------------------------------------------------------
            ' MATERIAL REQUISITION
            '-----------------------------------------------------------
         
            '-----------------------------------------------------------
            ' RmxRecipe
            '-----------------------------------------------------------
        
    End With
ERR_END:
    On Error GoTo 0
    Exit Sub
VarSet:
    'MsgBox Err.Description
    Resume Next
    
End Sub


Private Function SaveRecipeForProductionInDatabase() As Boolean
Dim rc As Boolean

' se sono in Data allora la ricerca è tra i Recipe chiusi!!
On Error GoTo SaveReceipt
    rc = True
    With dbTabReceiptForProduction
        .filter = ""
        .filter = "FileName ='" & SettingName & IIf(bIfDataPath, "' and bClosed=true", "' and bClosed=false")
        If .EOF Then
                
            .AddNew
            
        Else
        
        
        End If
        
        !Recipe = GetStrRecipe(uRecipe)
        !Description = GetStrDescriptionRecipe(uRecipe)
        !Line = GetStrLineRecipe(uRecipe)
        !PlanningReference = txFormulation(4)
        !DataRecipe = txFormulation(1)
        !RecipeWeek = txFormulation(2)
        !PlannedPreparation = txFormulation(3)
        !Operator = txFormulation(0)
        !bClosed = bIfDataPath
        !Note = txFormulation(5)
        !FileName = SettingName
    
        .Update
    
    End With

ERR_END:
    On Error GoTo 0
    SaveRecipeForProductionInDatabase = rc
    Exit Function
SaveReceipt:
    rc = False
    MsgBox err.Description
    Resume Next
    
End Function
Private Function ChecktxFormulation() As Boolean
Dim rc As Boolean
Dim i As Integer
    rc = True
    For i = txFormulation.LBound To txFormulation.UBound - 1
        If Len(txFormulation(i)) = 0 Then
            rc = False
            MessageInfoTime = 1500
            PopupMessage 2, "Please Enter field : " & lbFormulation(i), , True, "Recipe For Production"
            txFormulation(i).SetFocus
            rc = False
            Exit For
        End If
    Next
    ChecktxFormulation = rc
End Function





'-----------------------------------------------------------------------------------------------
'
'
'                                   MATERIAL REQUISITION

'
'-----------------------------------------------------------------------------------------------


Private Function GetRecipeCodeFromMixes(ByRef Index As Integer) As String
Dim bMix As Boolean
Dim MixCode As String
Dim RecipeCode As String

Dim i As Integer
With Grid4
  If .Rows > 1 Then
        For i = 1 To .Rows - 1
        
            bMix = .Cell(i, 15).Text
            
            If bMix Then
            
                MixCode = .Cell(i, 1).Text
                Index = i
                GetRecipeCodeFromMixes = MixCode
                Exit Function
        
            End If
        Next
  End If

End With

End Function


Private Sub SetMaterialRequisitionOnlyMixes()
Dim rc As Boolean



     SelectedRecipeCode = GetRecipeCodeFromMixes(IndexRecipe)
     lbInside(6) = "Materials Requested Table : " & SelectedRecipeCode
     
     If bImportata Then

        ' devo trovare il RecipeCode Madre!
        
                    
        If F_MsgBox.DoShow("Material Requisition already done. Create new Material Requisition?", SelectedRecipeCode) = False Then
        
            rc = AddMaterialRequisitionFromFile(Grid6, IndexRecipe)
            If rc = False Then GoTo MrMixes:
            
            PicMenu_Click 1
        Else
            GoTo MrMixes
        End If
        
    Else
MrMixes:
        Call GotoMaterialRequisition(False, True)
    End If
    
              
        
End Sub

Private Sub SetMaterialRequisitionMixes()
Dim rc As Boolean
     If bImportata Then
                    
        If F_MsgBox.DoShow("Material Requisition already done. Create new Material Requisition?", SelectedRecipeCode) = False Then
        
            rc = AddMaterialRequisitionFromFile(Grid6, IndexRecipe)
            If rc = False Then GoTo MrMixes:
            
            PicMenu_Click 1
        Else
            GoTo MrMixes
        End If
        
    Else
MrMixes:
        If SelectedRecipeCode = "" Then
            PopupMessage 2, "Please select a Recipe first..", , True
        Else
            Call GotoMaterialRequisition(False, False)
        End If
        
    End If
    
              
        
End Sub




Private Sub SetMaterialRequisitionComponents()
Dim rc As Boolean
    If bImportata Then
        If F_MsgBox.DoShow("Material Requisition already done. Create new Material Requisition?", SelectedRecipeCode) = False Then
            rc = AddMaterialRequisitionFromFile(Grid6, IndexRecipe)
            If rc = False Then GoTo MrComponent:
            PicMenu_Click 1
        Else
            GoTo MrComponent
        End If
    Else
MrComponent:
        If SelectedRecipeCode = "" Then
            PopupMessage 2, "Please select a Recipe first..", , True
        Else
            Call GotoMaterialRequisition(True, False)
        End If
    End If
End Sub

Private Sub GotoMaterialRequisition(ByVal bValue As Boolean, ByVal bMixes As Boolean)
Dim rc As Boolean
Dim Index As Integer
Dim i As Integer
Dim MaterialReqRecipe As RecipeType
Dim MaterialReqMixes() As RecipeType
Dim ArrayMaterialReqRecipe() As RecipeType


    If IndexRecipe = 0 Then IndexRecipe = 1

        ReDim ArrayMaterialReqRecipe(1)
        Index = IndexRecipe
        
        
        If bMixes Then
            rc = SetMaterialRequisitionAllMixes(MaterialReqMixes, Grid4)
            If rc Then
            
                ArrayMaterialReqRecipe = MaterialReqMixes
                txDocument(2) = ArrayMaterialReqRecipe(1).Line
                lRowMaterialReq = 0
                lColMaterialReq = 0
                Call AddMixesToMaterialReqGrid(Grid6, ArrayMaterialReqRecipe())
                PicMenu_Click 1
            
            End If
        Else
        
            rc = SetMaterialRequisition(uRecipe, MaterialReqRecipe, IndexRecipe, bValue)
            ArrayMaterialReqRecipe(1) = MaterialReqRecipe
            
            
            txDocument(2) = ArrayMaterialReqRecipe(1).Line

    
            If rc Then
                lRowMaterialReq = 0
                lColMaterialReq = 0
            
                Call AddRecipeToMaterialReqGrid(Grid6, ArrayMaterialReqRecipe())
                PicMenu_Click 1
        
            End If
        End If
        
      
   

End Sub
Private Sub StampaMaterialRequisition()
Dim rc As Boolean
Dim FileName As String
Dim xDocument() As String
Dim strHannaCode As String

    rc = True
     
    rc = CheckTxDocument(xDocument())
    If rc = False Then Exit Sub
    
    CloseSettingDataFile
    
    lbWait.Caption = "Material Requisition PDF file : Wait while Saving Data..."
    lbWait.Visible = True



    rc = MaterialRequisitionSaveSettingsFile(Grid6, xDocument(), SettingName, IndexRecipe)
    
    Call SaveMaterialRequisitionForRecipeForProductionInDatabase
    
    If uRecipeForProduction.HannaCodesCount = 0 Then
    Else
        strHannaCode = SetNeHannaCodeQtyString(uRecipeForProduction.HannaCodes)
    End If
    
    
    
    If rc Then rc = MaterialRequisitionSaveSettingsTempFile(Grid6, xDocument(), FileName, strHannaCode, "")
    If rc Then rc = ReportStampato(FileName)
    If rc Then
        PopupMessage 2, "Document Succesfully Generated...", , , "Material Requisition : MR-" & xDocument(0)
        
        
        ' user temp / data Material Requisition  TEMP file! da cancellare! lo uso solo per stampare...
        
        If FileExists(USER_PATH & FileName) Then
        

            Kill USER_PATH & FileName

        End If
        
        
    End If
    CloseSettingDataFile
    lbWait.Visible = False



End Sub


Private Function SaveMaterialRequisitionForRecipeForProductionInDatabase() As Boolean
Dim rc As Boolean
Dim strMaterialRequisition As String
On Error GoTo SaveReceipt
    rc = True
    With dbTabReceiptForProduction
        .filter = ""
        .filter = "FileName ='" & SettingName & IIf(bIfDataPath, "' and bClosed=true", "' and bClosed=false")
        If .EOF Then
                
           ' .AddNew
                
            '!Recipe = GetStrRecipe(uRecipe)
            '!Description = GetStrDescriptionRecipe(uRecipe)
            '!Line = GetStrLineRecipe(uRecipe)
           ' !PlanningReference = txFormulation(4)
            
            '!DataRecipe = txFormulation(1)
           ' !RecipeWeek = txFormulation(2)
           ' !PlannedPreparation = txFormulation(3)
           ' !Operator = txFormulation(0)
           ' !bClosed = bIfDataPath
           ' !Note = txFormulation(5)
            '!FileName = SettingName
        Else
        
        End If
        
        If IsNull(Trim(!MaterialRequisitionNumber)) Or Trim(!MaterialRequisitionNumber) = "" Then
            strMaterialRequisition = txDocument(0)
        Else
            strMaterialRequisition = CheckStrMaterialRequisition(Trim(!MaterialRequisitionNumber), txDocument(0))
        End If
        
        !MaterialRequisitionNumber = strMaterialRequisition
        !bMaterialRequisitionPrinted = True
        .Update
    
    End With

ERR_END:
    On Error GoTo 0
    SaveMaterialRequisitionForRecipeForProductionInDatabase = rc
    Exit Function
SaveReceipt:
    rc = False
    MsgBox err.Description
    Resume Next
    
End Function



Private Function ReportStampato(ByVal FileName As String) As Boolean
    Dim rc As Boolean
    Dim NumReport As String
    On Error GoTo ERR_SAVE
    rc = True

    NumReport = FormatNomeFile(txDocument(0) & "." & txDocument(1))

    rc = OkStampa(NumReport, bSeStampa, FileName)
     
ERR_END:
    On Error GoTo 0
    ReportStampato = rc
    Exit Function
ERR_SAVE:
    rc = False
    Resume ERR_END
End Function

Private Function CheckTxDocument(ByRef xDocument() As String) As Boolean
Dim rc As Boolean
Dim i As Integer
    rc = True
    ReDim xDocument(txDocument.UBound)
    
    For i = txDocument.LBound To txDocument.UBound
        xDocument(i) = txDocument(i)
        If Len(txDocument(i)) = 0 Then
            rc = False
            PopupMessage 2, "Please Enter field : " & lbDocument(i), , True, "RecipeForProduction Document"
            txDocument(i).SetFocus
            Exit For
        End If
    Next
    CheckTxDocument = rc
End Function






Public Function AddMaterialRequisitionFromFile(ByVal Grd As Grid, ByVal Index As Integer) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim X As Integer
Dim RowsCount As Integer

    On Error GoTo ERR_ADD:

    If Index = 0 Then Index = 1
    
    rc = True
    

    If SettingName = "" Then
        rc = False
        GoTo ERR_END
     
    End If

    CloseSettingDataFile


    For i = 0 To txDocument.UBound
        txDocument(i) = GetSettingData(SettingName, "Material Requisition" & Index, "txDocument(" & i & ")", "")
    Next
    RowsCount = GetSettingData(SettingName, "Material Requisition" & Index, "Rows", 0)
    
    
    If RowsCount = 0 Then
        ' ho riaperto la Recipe ma non hpo fatto Material requisition
        rc = False
        GoTo ERR_END
    End If
    
    
    With Grd
        .AutoRedraw = False
        .Rows = 1
        For i = 1 To RowsCount
            .AddItem "", False
            For t = 1 To .Cols - 1
                .Cell(i, t).Text = GetSettingData(SettingName, "Material Requisition" & Index, "Grd(" & i & "," & t & ")", "")
                .Column(t).Alignment = cellLeftCenter
                .Column(t).Width = 150
                .Cell(0, t).FontBold = True
            
            Next
        Next
        
        .Column(2).Width = 250
        .Column(3).Width = 100
        .Column(5).Width = 100
        .Column(4).Alignment = cellRightCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Refresh
        .AutoRedraw = True
    End With

ERR_END:
    On Error GoTo 0
    AddMaterialRequisitionFromFile = rc
    Exit Function
ERR_ADD:
    rc = False
    MsgBox err.Description
    GoTo ERR_END
End Function





'-----------------------------------------------------------------------------------------------
'
'
'                                   GetReceiptFromFile

'
'-----------------------------------------------------------------------------------------------



Private Sub GetFileInfo()
Dim rc As Boolean
 rc = GetReceiptFromFile
 
 bImportata = rc
 
End Sub

Private Function GetReceiptFromFile() As Boolean
Dim i As Integer
Dim rc As Boolean

On Error GoTo ERR_GET:
   
rc = True

    CloseSettingDataFile

        
    lbWait.Visible = True

    
    ReDim uRecipe(0)
    
    Debug.Print USER_PATH
    
    uRecipeForProduction = uRecipeForProductionClean
    
    
    If SettingName = "" Then
            MessageInfoTime = 2000
            PopupMessage 2, "Warning : File non found! ", , True, "Recipe For Production"
            rc = False
            GoTo ERR_END
    End If
    
    If GetSettingData(SettingName, "iRecipeForProduction", "bOpen", True) Then
            PBFooter.BackColor = &H886010
            PBTitle.BackColor = &H644603
        Else
            PBFooter.BackColor = &H4D3B37   '&H40C0&
            blTable.Visible = False
            blTable = "Recipe for Production : Closed"
            PicMenu(1).Visible = False
            PBTitle.BackColor = &H473733
            PicMenu(0).BackColor = &H473733
    End If
    
    
    Call ReceiptGetSetting(uRecipeForProduction, SettingName)
     
    
    With uRecipeForProduction

        uRecipe = .Recipes
        

        
        
        txFormulation(1) = .DateRecipe
        txFormulation(5) = .Note
        txFormulation(3) = .PlannedPrepWeek
        txFormulation(4) = .PlanningReference
        txFormulation(2) = .numPrepWeek
        txFormulation(0) = .RecipeBy
        
        
        If .RecipeCount = 0 Then
            MessageInfoTime = 2000
            PopupMessage 2, "Warning : No recipes found! ", , True, "Recipe For Production"
            rc = False
            GoTo ERR_END
        End If
        
    End With
    
    Call FillGridRfPFromFile(Grid1, uRecipeForProduction, 1)
    Call FillGridRfPFromFile(Grid2, uRecipeForProduction, 2)
    Call FillGridRfPFromFile(Grid4, uRecipeForProduction, 4)
    Call FillGridRfPFromFile(Grid5, uRecipeForProduction, 5)
    Call FillGridRfPFromFile(Grid3, uRecipeForProduction, 3)
    
    '-------------------------------------------------------
    '
    ' nascondo i Recipes senza qty
    '
    Call ViewRecipesRFP(uRecipe, Grid1, Grid2, Grid4, Grid5, False)
    '
    '-------------------------------------------------------
    
    ' visualizzo o no il mat req per i mixes in grid4
    frCommandInside(13).Visible = IfMixesInTotalGrid(Grid4)
    
    
ERR_END:

    On Error GoTo 0
    lbWait.Visible = False

    blTable.Visible = True
        
        
    GetReceiptFromFile = rc
    Exit Function
ERR_GET:
    rc = False
    Resume Next
    
End Function





Private Function SetTotalMixes()
Dim i As Integer
Dim bMix As Boolean
Dim TotalWeightKg As String
    With Grid4
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                bMix = .Cell(i, 15).Text
                If bMix Then
                    TotalWeightKg = .Cell(i, 3).Text
                    TotalWeightKg = Replace(LCase(TotalWeightKg), "kg", "")
                    uRecipe(i).TotalWeightKg = CDbl(TotalWeightKg)
                End If
            Next
        End If
        .Refresh
    End With
End Function

Private Function isMixInGrid7(ByVal Grid7 As Grid, ByVal Code As String) As Boolean
Dim i As Integer
Dim rc As Boolean
rc = False
With Grid7
    If .Rows > 1 Then
        For i = 1 To .Rows - 1
            
            If Trim(.Cell(i, 1).Text) = Code Then
        
                rc = True
                Exit For
        
            End If
        Next
    
        
    
    End If

End With

isMixInGrid7 = rc
End Function
