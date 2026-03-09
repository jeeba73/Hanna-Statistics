VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form FormFormulation 
   BackColor       =   &H00F0F0F0&
   Caption         =   "Chemical Production"
   ClientHeight    =   12000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19200
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormFormulation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   12000
   ScaleWidth      =   19200
   Begin VB.PictureBox PicHover 
      BackColor       =   &H00886010&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   675
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   675
      Begin VB.Label lblHoverClick 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00808080&
         Height          =   570
         Left            =   60
         TabIndex        =   49
         Top             =   0
         Width           =   585
      End
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
         TabIndex        =   48
         Top             =   80
         Width           =   330
      End
   End
   Begin VB.PictureBox PBFooter 
      BackColor       =   &H00886010&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   8
      Top             =   11040
      Width           =   19215
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   8400
         Top             =   120
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   4
         Left            =   17280
         TabIndex        =   9
         Top             =   -120
         Width           =   1935
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   0
         Left            =   8760
         TabIndex        =   12
         Top             =   -120
         Width           =   1695
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   3
         Left            =   14760
         TabIndex        =   10
         Top             =   -120
         Width           =   2175
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
         TabIndex        =   75
         Top             =   660
         Width           =   1380
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
         Left            =   15360
         MousePointer    =   99  'Custom
         TabIndex        =   74
         Top             =   660
         Width           =   1230
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
         TabIndex        =   73
         Top             =   660
         Width           =   1200
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   0
         Left            =   9360
         MousePointer    =   99  'Custom
         Picture         =   "FormFormulation.frx":33E2
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   3
         Left            =   15600
         MousePointer    =   99  'Custom
         Picture         =   "FormFormulation.frx":67C4
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   4
         Left            =   18240
         Picture         =   "FormFormulation.frx":9BA6
         Top             =   120
         Width           =   480
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   1
         Left            =   3960
         TabIndex        =   11
         Top             =   -240
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.PictureBox PBTitle 
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
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   1
      Top             =   0
      Width           =   19215
      Begin ChemicalMR.ucScrollAdd ucScrollAdd1 
         Left            =   15000
         Top             =   360
         _ExtentX        =   1138
         _ExtentY        =   423
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   4
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "FormFormulation.frx":CF88
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New"
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
            TabIndex        =   5
            Top             =   640
            Width           =   1830
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   1920
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   2
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "FormFormulation.frx":1036A
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Save"
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
            TabIndex        =   3
            Top             =   640
            Width           =   1830
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Code"
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
         Left            =   9060
         TabIndex        =   6
         Top             =   240
         Width           =   9870
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   19200
      _ExtentX        =   33867
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox ctlCalendar1 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   0
      Width           =   1000
   End
   Begin VB.PictureBox PBContainerViewport 
      BorderStyle     =   0  'None
      Height          =   9735
      Left            =   0
      ScaleHeight     =   9735
      ScaleWidth      =   19215
      TabIndex        =   13
      Top             =   960
      Width           =   19215
      Begin VB.Frame PBContainer 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   29615
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   19215
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "&H00F0F0F0&"
            Height          =   7335
            Index           =   5
            Left            =   1200
            TabIndex        =   59
            Top             =   20040
            Visible         =   0   'False
            Width           =   15255
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
               Index           =   4
               Left            =   6720
               TabIndex        =   66
               Top             =   4680
               Width           =   3255
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Back To Recipe"
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
                  TabIndex        =   67
                  Top             =   120
                  Width           =   3255
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H000040C0&
               BorderStyle     =   0  'None
               Height          =   1335
               Left            =   5880
               TabIndex        =   63
               Top             =   1800
               Width           =   5055
               Begin VB.Label lbChem 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Open Recipe/Chemical to add Components..."
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
                  Left            =   315
                  TabIndex        =   68
                  Top             =   720
                  Width           =   4380
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Empty List..."
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   2
                  Left            =   1920
                  TabIndex        =   64
                  Top             =   360
                  Width           =   1155
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   0  'None
               Caption         =   "l"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   3
               Left            =   0
               TabIndex        =   60
               Top             =   0
               Width           =   15255
               Begin VB.Label Label13 
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
                  TabIndex        =   62
                  Top             =   180
                  Width           =   1050
               End
               Begin VB.Label lbInside 
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
                  Index           =   4
                  Left            =   0
                  TabIndex        =   61
                  Top             =   120
                  Width           =   3345
               End
               Begin VB.Line Line8 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
            End
            Begin FlexCell.Grid Grid3 
               Height          =   3735
               Left            =   0
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   600
               Width           =   15255
               _ExtentX        =   26908
               _ExtentY        =   6588
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
            Begin VB.Line Line9 
               BorderColor     =   &H00D0D0D0&
               X1              =   120
               X2              =   15240
               Y1              =   4440
               Y2              =   4440
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "&H00F0F0F0&"
            Height          =   6135
            Index           =   3
            Left            =   1200
            TabIndex        =   34
            Top             =   8680
            Width           =   17055
            Begin VB.Frame frHannaCode 
               Appearance      =   0  'Flat
               BackColor       =   &H000040C0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1335
               Left            =   5880
               TabIndex        =   38
               Top             =   1560
               Visible         =   0   'False
               Width           =   5055
               Begin VB.Label lbChem 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Click + to add chemicals"
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
                  Left            =   45
                  TabIndex        =   46
                  Top             =   720
                  Width           =   5010
               End
               Begin VB.Label lbChem 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Empty List..."
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   2
                  Left            =   0
                  TabIndex        =   45
                  Top             =   360
                  Width           =   4995
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   0  'None
               Caption         =   "l"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   2
               Left            =   0
               TabIndex        =   35
               Top             =   0
               Width           =   17055
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe Link with Hanna Codes"
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
                  Index           =   2
                  Left            =   14115
                  TabIndex        =   37
                  Top             =   180
                  Width           =   2655
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hanna Codes"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00606060&
                  Height          =   345
                  Index           =   3
                  Left            =   0
                  TabIndex        =   36
                  Top             =   75
                  Width           =   1890
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   0
                  X2              =   16800
                  Y1              =   480
                  Y2              =   480
               End
            End
            Begin FlexCell.Grid Grid1 
               Height          =   3615
               Left            =   0
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   600
               Width           =   16935
               _ExtentX        =   29871
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
               Caption         =   "Delete Code"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   255
               Left            =   8760
               TabIndex        =   44
               Top             =   5040
               Width           =   1245
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Add  Code"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   255
               Left            =   7080
               TabIndex        =   43
               Top             =   5040
               Width           =   1080
            End
            Begin VB.Label lbFunction 
               BackStyle       =   0  'Transparent
               Height          =   855
               Index           =   3
               Left            =   8400
               TabIndex        =   42
               Top             =   4275
               Width           =   1815
            End
            Begin VB.Label lbFunction 
               BackStyle       =   0  'Transparent
               Height          =   855
               Index           =   2
               Left            =   6840
               TabIndex        =   41
               Top             =   4275
               Width           =   1575
            End
            Begin VB.Line Line6 
               BorderColor     =   &H00D0D0D0&
               X1              =   0
               X2              =   16800
               Y1              =   4320
               Y2              =   4320
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Select Hanna Code and link to Recipe"
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
               Index           =   1
               Left            =   0
               TabIndex        =   40
               Top             =   5400
               Width           =   17055
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   2
               Left            =   7440
               Picture         =   "FormFormulation.frx":1374C
               Top             =   4680
               Width           =   240
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   3
               Left            =   9120
               Picture         =   "FormFormulation.frx":1414E
               Top             =   4680
               Width           =   240
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1455
            Index           =   4
            Left            =   1200
            TabIndex        =   29
            Top             =   18000
            Width           =   17055
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
               Index           =   1
               Left            =   8520
               TabIndex        =   32
               Top             =   600
               Width           =   3255
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Close"
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
                  TabIndex        =   33
                  Top             =   120
                  Width           =   3255
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00008000&
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
               Index           =   0
               Left            =   5040
               TabIndex        =   30
               Top             =   600
               Width           =   3255
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
                  Index           =   0
                  Left            =   0
                  TabIndex        =   31
                  Top             =   120
                  Width           =   3255
               End
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "16800"
            Height          =   7335
            Index           =   2
            Left            =   1200
            TabIndex        =   15
            Top             =   960
            Width           =   17055
            Begin VB.Frame SetPercentageLastComponent 
               BackColor       =   &H00008000&
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
               Left            =   10800
               TabIndex        =   78
               Top             =   4560
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbSetPercentageLastComponent 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Set   %   last component"
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
                  TabIndex        =   79
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00886010&
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
               Index           =   2
               Left            =   0
               TabIndex        =   76
               Top             =   4560
               Visible         =   0   'False
               Width           =   4815
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Open Component Table"
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
                  TabIndex        =   77
                  Top             =   120
                  Width           =   4815
               End
            End
            Begin VB.PictureBox PicUmComponent 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               CausesValidation=   0   'False
               Height          =   2295
               Left            =   1560
               ScaleHeight     =   2295
               ScaleWidth      =   6975
               TabIndex        =   69
               Top             =   2040
               Visible         =   0   'False
               Width           =   6975
               Begin VB.ComboBox cmbUM2 
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
                  Left            =   2160
                  Style           =   2  'Dropdown List
                  TabIndex        =   70
                  Top             =   1200
                  Width           =   2655
               End
               Begin VB.Image Image6 
                  Height          =   240
                  Left            =   6600
                  Picture         =   "FormFormulation.frx":14B50
                  Top             =   120
                  Width           =   240
               End
               Begin VB.Label lbUM2 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "label"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   3225
                  TabIndex        =   72
                  Top             =   360
                  Width           =   480
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select from list"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   0
                  Left            =   2820
                  TabIndex        =   71
                  Top             =   720
                  Width           =   1380
               End
            End
            Begin VB.Frame frQuantityCheck 
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   0  'None
               Caption         =   "Frame7"
               Height          =   615
               Left            =   1800
               TabIndex        =   50
               Top             =   6000
               Width           =   12855
               Begin VB.PictureBox PicPerc 
                  BackColor       =   &H000040C0&
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   2280
                  ScaleHeight     =   255
                  ScaleWidth      =   255
                  TabIndex        =   51
                  Top             =   240
                  Width           =   255
               End
               Begin VB.Image imPerc 
                  Height          =   240
                  Left            =   11880
                  Picture         =   "FormFormulation.frx":15552
                  Top             =   240
                  Width           =   240
               End
               Begin VB.Label lbTotalWL 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "0"
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
                  Left            =   9840
                  TabIndex        =   57
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.Label lbTotalWKg 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "0"
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
                  Left            =   5880
                  TabIndex        =   56
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total Weight ( L )"
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
                  Left            =   7920
                  TabIndex        =   55
                  Top             =   240
                  Width           =   1620
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total Weight ( Kg )"
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
                  Left            =   3720
                  TabIndex        =   54
                  Top             =   240
                  Width           =   1785
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Percentage Check"
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
                  Left            =   240
                  TabIndex        =   53
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.Label lbPerc 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00808080&
                  Height          =   255
                  Left            =   2760
                  TabIndex        =   52
                  Top             =   240
                  Width           =   60
               End
               Begin VB.Shape Shape1 
                  BorderColor     =   &H00E0E0E0&
                  Height          =   735
                  Left            =   -1320
                  Top             =   -2040
                  Width           =   11295
               End
            End
            Begin VB.Frame frChemicals 
               BackColor       =   &H00886010&
               BorderStyle     =   0  'None
               Height          =   1335
               Left            =   5880
               TabIndex        =   21
               Top             =   1800
               Visible         =   0   'False
               Width           =   5055
               Begin VB.Label lbChem 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Empty List..."
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   23
                  Top             =   360
                  Width           =   4995
               End
               Begin VB.Label lbChem 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Click + to add chemicals"
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
                  Left            =   45
                  TabIndex        =   22
                  Top             =   720
                  Width           =   5010
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
               Index           =   3
               Left            =   13920
               TabIndex        =   19
               Top             =   4560
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Delete Table"
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
                  TabIndex        =   20
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
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   0
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Width           =   17055
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe Formulation"
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
                  Left            =   15165
                  TabIndex        =   18
                  Top             =   120
                  Width           =   1725
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe : "
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00606060&
                  Height          =   345
                  Index           =   1
                  Left            =   0
                  TabIndex        =   17
                  Top             =   75
                  Width           =   13200
               End
               Begin VB.Line Line4 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   0
                  X2              =   16920
                  Y1              =   480
                  Y2              =   480
               End
            End
            Begin FlexCell.Grid GridChemicals 
               Height          =   3615
               Left            =   0
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   600
               Width           =   16935
               _ExtentX        =   29871
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
            Begin VB.Shape Shape2 
               BorderColor     =   &H00D0D0D0&
               Height          =   930
               Left            =   0
               Top             =   5880
               Width           =   16935
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Set Each Component Quantity in Recipe: click darker cells and set quantity"
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
               Index           =   0
               Left            =   5430
               TabIndex        =   58
               Top             =   7080
               Width           =   6105
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00D0D0D0&
               X1              =   0
               X2              =   16920
               Y1              =   4320
               Y2              =   4320
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   1
               Left            =   9120
               Picture         =   "FormFormulation.frx":15F54
               ToolTipText     =   "4000"
               Top             =   4605
               Width           =   240
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   0
               Left            =   7440
               Picture         =   "FormFormulation.frx":16956
               ToolTipText     =   "4000"
               Top             =   4605
               Width           =   240
            End
            Begin VB.Label lbFunction 
               BackStyle       =   0  'Transparent
               Height          =   855
               Index           =   0
               Left            =   6840
               TabIndex        =   28
               Top             =   4440
               Width           =   1575
            End
            Begin VB.Label lbFunction 
               BackStyle       =   0  'Transparent
               Height          =   855
               Index           =   1
               Left            =   8400
               TabIndex        =   27
               Top             =   4440
               Width           =   1815
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Add Comp."
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   255
               Left            =   6960
               TabIndex        =   26
               Top             =   4875
               Width           =   1155
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Delete Comp."
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   255
               Left            =   8640
               TabIndex        =   25
               Top             =   4875
               Width           =   1380
            End
         End
         Begin VB.Line Line2 
            Visible         =   0   'False
            X1              =   9600
            X2              =   9600
            Y1              =   120
            Y2              =   39600
         End
      End
   End
End
Attribute VB_Name = "FormFormulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


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

Private MyLbHelpCount As Integer

Private IndexTabella As Integer
Private MaxIndex As Integer



Private bHilight As Boolean

Private m_rc As Boolean

Private MyID As Long
Private MyIndexRecord As Integer
Private lRow As Long
Private lCol As Long
Private DatabaseIndex As Integer

Private DatabaseString As String
Private dbTab As ADODB.Recordset
Private ID_CHEMICAL As Long


Private uRecipe As RecipeType

Private IndexVisibleFrame As Integer

Private SelectedCode As String
Private UmComponent As String


Private bSetPercentageLastComponent As Boolean
Private bFlagOpenRecipe As Boolean





Private Sub cmbUM2_Click()
Dim i As Integer
'Dim Perc As String
    uRecipe.UmMultiple = cmbUM2
    
    With GridChemicals
        For i = 1 To .Rows - 1
            .Cell(i, 5).Text = cmbUM2
            
           ' Perc = .Cell(i, 6).Text
            
            'If Perc <> "" Then
                'Call SetComponentWeightByPerc(Perc, i)
          '  End If
            
             
        Next
    End With
    
    Call CheckTotalsAndPercentage
    
End Sub





Private Sub Form_Activate()
Me.WindowState = MainWindowState
End Sub

Private Sub Form_Initialize()
SaveSizes
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FormFormulation = Nothing
End Sub


Private Sub Image6_Click()
UmComponent = cmbUM2
PicUmComponent.Visible = False

End Sub



Private Sub lbSetPercentageLastComponent_Click()
SetPercentageLastComponent_Click
End Sub

Private Sub lbUM2_Click()
Image6_Click
End Sub

Private Sub PicUmComponent_Click()
Image6_Click
End Sub



Private Sub SetPercentageLastComponent_Click()


bSetPercentageLastComponent = True

Call SetComponentWeightByPerc("", 0)

End Sub



Private Sub Timer2_Timer()
CheckTotalsAndPercentage
Timer2.Enabled = False
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
   
    If PBContainerViewport.Visible = False Then Exit Sub
    
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
    
    
    If ucScrollAdd1.UCScrollV.Value <= frInside(2).Top Then
        IndexVisibleFrame = 2
    ElseIf ucScrollAdd1.UCScrollV.Value > frInside(2).Top And ucScrollAdd1.UCScrollV.Value <= frInside(3).Top Then
        IndexVisibleFrame = 3
    
    ElseIf ucScrollAdd1.UCScrollV.Value > frInside(3).Top And ucScrollAdd1.UCScrollV.Value <= frInside(4).Top Then
        IndexVisibleFrame = 4
  '  ElseIf ucScrollAdd1.UCScrollV.Value > frInside(4).Top And ucScrollAdd1.UCScrollV.Value <= frInside(5).Top Then
       ' IndexVisibleFrame = 3
   ' ElseIf ucScrollAdd1.UCScrollV.Value > frInside(5).Top And ucScrollAdd1.UCScrollV.Value <= frInside(6).Top Then
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
            'MsgBox (TypeName(ctl))
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



For Each ctl In Controls
    With m_ControlPositions(i)
        If TypeOf ctl Is Line Then
            ctl.x1 = x_scale * .Left
            ctl.y1 = y_scale * .Top
            ctl.x2 = ctl.x1 + x_scale * .Width
            ctl.y2 = ctl.y1 + y_scale * .Height
        ElseIf TypeOf ctl Is Menu Then
        ElseIf TypeOf ctl Is Inet Then
        ElseIf TypeOf ctl Is Image Then
            ctl.Left = (x_scale * .Left) + IIf(x_scale = 1, 0, (x_scale - 1) * 200)
            ctl.Top = y_scale * .Top
        ElseIf TypeOf ctl Is Timer Then
        ElseIf TypeOf ctl Is ucScrollAdd Then
        ElseIf TypeOf ctl Is Grid Then
           ctl.Left = x_scale * .Left
            ctl.Top = y_scale * .Top
            ctl.Width = x_scale * .Width
            ctl.Height = y_scale * .Height

             
           
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


Private Sub Form_Load_Scroll()

   

    PBContainer.Top = 0
    PBContainer.Left = 0
    PBContainer.Width = Me.Width
    IndexVisibleFrame = 2
    ucScrollAdd1.AddScroll PBContainerViewport
    ucScrollAdd1.TrackMouseWheel Vertical
   ucScrollAdd1.ResizeTargetOnFormResize 0, 0
   ' ucScrollAdd1.RemoveFromContainer PBFooter
    ucScrollAdd1.UCScrollV.ShowButtons = False
     ucScrollAdd1.UCScrollH.ShowButtons = False
    
    
    Dim i As Integer
    If Screen.Width - Me.Width > 1000 And bFullScreen Then
        Me.WindowState = 2
    
    End If



        PBContainerViewport.Move 0, PBTitle.Height, Me.ScaleWidth, Me.ScaleHeight - PBTitle.Height

  
    RSBottom PicHover, Me, -1350
    RSRight PicHover, Me, -450
   

    PBContainerViewport.ZOrder
    PBFooter.ZOrder
    
  
    
End Sub
Private Sub Form_Load()
Dim i As Integer
    
    
On Error GoTo ERR_LOAD:
        PBContainerViewport.Left = 0
        PBContainerViewport.Top = PBTitle.Height
        PBContainerViewport.Width = Me.ScaleWidth
        PBContainerViewport.Height = PBFooter.Top
        
        

    PBContainer.Top = 0
    PBContainer.Left = 0
    PBContainer.Width = Me.Width
    IndexVisibleFrame = 2
    ucScrollAdd1.AddScroll PBContainerViewport
    ucScrollAdd1.TrackMouseWheel Vertical
   ucScrollAdd1.ResizeTargetOnFormResize 0, 0
   ' ucScrollAdd1.RemoveFromContainer PBFooter
    ucScrollAdd1.UCScrollV.ShowButtons = False
     ucScrollAdd1.UCScrollH.ShowButtons = False
    
  
    If Screen.Width - Me.Width > 1000 And bFullScreen Then
        Me.WindowState = 2
    
    End If



    PBContainerViewport.Move 0, PBTitle.Height, Me.ScaleWidth, Me.ScaleHeight - PBTitle.Height

  
    RSBottom PicHover, Me, -1350
    RSRight PicHover, Me, -450
   

    PBContainerViewport.ZOrder
    PBFooter.ZOrder
    
     

  MyID = 0
  MyIndexRecord = 3



'Form_Load_Scroll
   
DoEvents
DoEvents
DoEvents

ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_LOAD:
    MsgBox "FOR_LOAD ERROR"
    GoTo ERR_END:

End Sub


Public Function DoShow() As Boolean
Dim i As Integer
    On Error GoTo ERR_SHOW
    
    m_rc = False
    mOk
    
    
  
    
    SetChemicalsXRecipe

    
    uRecipe = zRecipe
    
    'uRecipe.Code = Code
    
  
    blTable = "Formulation : " & uRecipe.Code
    lbInside(1) = "Recipe : " & uRecipe.Code
  
    CheckChemicalsPerRecipe
    
 
 
    DoEvents
    
  


    Me.Show vbModal
    
    

    
    If m_rc = True Then
    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    MsgBox "DOSHOW ERROR  " & err.Description
    Resume ERR_END
End Function




Private Sub frChemicals_Click()
ImCode_Click 0
End Sub

Private Sub frCommandInside_Click(Index As Integer)
Select Case Index

    Case 0
        ' save
        Call SaveRecipe
    Case 1
         Unload Me
    Case 2
    
        Call SetComponentComponentRecipe

        
    Case 3
        If F_MsgBox.DoShow("Delete all Components?", uRecipe.Code) Then
            If DeleteRecipeComponentByCode(uRecipe.Code) Then
                CheckChemicalsPerRecipe
                PopupMessage 2, "Records deleted...", , , uRecipe.Code
                
            End If
            
        End If
        
    Case 4
        frInside(5).Visible = False
         ucScrollAdd1.UCScrollV.ScrollToValue 0
       

End Select
End Sub

Private Sub frHannaCode_Click()
ImCode_Click 2
End Sub


Private Sub CheckChemicalsPerRecipe()
                
 On Error GoTo ERR_CHECK:
 
    Call GetChemicalsPerRecipe(GridChemicals, uRecipe)
    Call CheckTotalsAndPercentage
    
    
    
    frChemicals.Visible = IIf(GridChemicals.Rows > 1, False, True)
    
    Call GetHannaCodePerRecipe(Grid1, uRecipe)
    Grid1.SelectionMode = cellSelectionByRow
ERR_END:
        On Error GoTo 0
        Exit Sub
ERR_CHECK:
    MsgBox "CheckChemicalsPerRecipe ERROR"
    GoTo ERR_END:
End Sub


Private Sub GridChemicals_Click()

Dim sString As String
Dim Perc As String

If lRow > 0 And lCol = 6 Then
    
    
    sString = GridChemicals.Cell(0, lCol).Text
    Perc = GridChemicals.Cell(lRow, lCol).Text
    
    
    Call SetPercentageNumber(sString, SelectedCode, Perc, lRow)

End If


End Sub


Private Function SetPercentageNumber(ByVal sString As String, ByVal SelectedCode As String, ByVal Perc As String, ByVal FirstRow As Long)

If F_InputBox.DoShow(sString, "Set Percentage", , , , Perc, , True, Me.Top) Then
    Call SetComponentWeightByPerc(Perc, FirstRow)
    
End If

End Function
Private Sub GridChemicals_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim Qty As String
Dim Perc As String
Dim sString As String
Dim sNote As String
Dim Value As String
SelectedCode = GridChemicals.Cell(FirstRow, 1).Text
Qty = GridChemicals.Cell(FirstRow, FirstCol).Text
sString = GridChemicals.Cell(0, FirstCol).Text
sNote = GridChemicals.Cell(FirstRow, FirstCol).Text
Perc = GridChemicals.Cell(FirstRow, 6).Text

lRow = 0
frCommandInside(2).Visible = False

If FirstRow > 0 Then
frCommandInside(2).Visible = IIf(GridChemicals.Cell(FirstRow, 9).Text, True, False)
lbCommandInside(2) = SelectedCode & " - Component Table"
lCol = FirstCol
lRow = FirstRow

    Select Case FirstCol
        Case 4
            ' Q.ty to produce
            
            If F_InputBox.DoShow(sString, "Qty/multiple", , , , Qty, , True, Me.Top) Then
            
                
                If IsNumeric(Qty) Or Qty = "" Then
                    GridChemicals.Cell(FirstRow, 4).Text = Qty
                    GridChemicals.Cell(FirstRow, 4).Alignment = cellCenterCenter
                    Call CheckTotalsAndPercentage
                End If
                
            End If
        Case 5
          ' Value = GridChemicals.Cell(FirstRow, 5).Text
           ' lRow = FirstRow
            
          '  lbUM2 = lbInside(1) ' GridChemicals.Cell(FirstRow, 1).Text
             
           '  If Value <> "" Then cmbUM2 = Value
             
           '  PicUmComponent.Left = GridChemicals.Width / 2 - PicUmComponent.Width / 2
           '  PicUmComponent.Top = GridChemicals.Height / 2 - PicUmComponent.Height / 2
           '  PicUmComponent.Visible = True

        Case 6  ' perc
            
           Call SetPercentageNumber(sString, SelectedCode, Perc, lRow)
        Case 7
              ' note
            Perc = GridChemicals.Cell(FirstRow, 7).Text
            If F_InputBox.DoShow(sString, "Tolerance", , , , Perc, , True, Me.Top) Then
                GridChemicals.Cell(FirstRow, 7).Text = Perc
                GridChemicals.Cell(FirstRow, 7).Alignment = cellCenterCenter
            End If
            
    
         Case 8
            ' note
            If F_InputBox.DoShow(sString, "Note", , , , sNote, , , Me.Top) Then
                GridChemicals.Cell(FirstRow, 8).Text = sNote
                GridChemicals.Cell(FirstRow, 8).Alignment = cellLeftCenter
            End If
    
    End Select

    

End If
End Sub




Private Sub Image3_Click(Index As Integer)

    

    Select Case Index
        Case 1

                frCommandInside_Click 0
           
     
    End Select
End Sub



Private Sub Form_Resize()

ucScrollAdd1.ContainerW = Me.ScaleWidth

ResizeControls

End Sub


Private Sub DefaultMenu_Click(Index As Integer)
Dim MyIndex As Integer
Select Case Index
    Case 0
        If F_MsgBox.DoShow("Quit Formulation?", uRecipe.Code) Then
            'GridChemicals.Cell(0, 0).SetFocus
            Unload Me
      
        End If
    Case 2
        ' Open Report folder
        OpenWithDefault (USER_DOCUMENTI & PathRequisition)
      
    Case 1
        ' filtro
        
       
        
    Case 3
        

        
            ' Previous
            If IndexVisibleFrame >= 3 Then
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

    
    
    Case 5
   
    Case 6
      
    Case 7
     
    Case 8
        m_rc = True
        Unload Me
    Case 9
       
    Case 10
      
    Case 11
      
    Case 12

        
End Select
End Sub



Private Sub DefaultMenuLabel_Click(Index As Integer)
DefaultMenu_Click Index
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





Private Sub ImageTAV_Click(Index As Integer)
Select Case Index
        Case 0
            Unload Me
        
        Case 2
        

End Select
End Sub

Private Sub ImCode_Click(Index As Integer)
Dim UserChCode As String
Dim UserHannaCode As String
Dim MyCodeID As Long


    frCommandInside(2).Visible = False
    

    Select Case Index
        Case 0
            'add
            If FormChemicalMR.DoShow(UserChCode) Then
                If UserChCode <> "" Then
                
                    frChemicals.Visible = False
                    
                    Call CopyUserChCodeInGrid(GridChemicals, UserChCode, uRecipe)
                    Call CheckTotalsAndPercentage
                
                End If
                
            End If
        Case 1
            ' delete
            If lRow > 0 Then
                If F_MsgBox.DoShow("Delete Component " & GridChemicals.Cell(lRow, 1).Text & " ? ", "Recipe : " & uRecipe.Code) Then
                    GridChemicals.ReadOnly = False
                    GridChemicals.Selection.DeleteByRow
                    GridChemicals.ReadOnly = True
                    Call CheckTotalsAndPercentage
                End If
            End If
            
        Case 2
            ' add hanna code per recipe
               If FormCodes.DoShow(UserHannaCode) Then
                If UserHannaCode <> "" Then
                
                    frHannaCode.Visible = False
                    
                    Call CopyUserHannaCodeInGrid(Grid1, UserHannaCode, uRecipe)
                
                End If
                
            End If
            
        Case 3
            ' delete hanna code per recipe
            If lRow > 0 Then
                If F_MsgBox.DoShow("Delete Recipe in Hanna Code " & Grid1.Cell(lRow, 1).Text & " ? ", "Recipe : " & uRecipe.Code) Then
                    Call DeleteRecipePerCode(Grid1.Cell(lRow, 1).Text, uRecipe.Code)
                    Call GetHannaCodePerRecipe(Grid1, uRecipe)
                    'Grid1.ReadOnly = False
                    'Grid1.Selection.DeleteByRow
                    'Grid1.ReadOnly = True
                    
                End If
            End If
                        
            
            
            
            
    End Select
End Sub

Private Sub imPerc_Click()
CheckTotalsAndPercentage
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
    Case 0
        Image6_Click
    Case 1
        frChemicals_Click
End Select

End Sub

Private Sub Label2_Click(Index As Integer)
Image3_Click Index
End Sub



Private Sub lbChem_Click(Index As Integer)
Select Case Index
    Case 0, 1
        frChemicals_Click
    Case 2, 3
        frHannaCode_Click
End Select


End Sub

Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub

Private Sub lbFunction_Click(Index As Integer)
ImCode_Click Index
End Sub





Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To PicMenu.Count - 1
    If i = Index Then
        PicMenu(i).BackColor = &H886010
    Else
        PicMenu(i).BackColor = &H644603
    End If
Next
End Sub

Private Sub PicMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3_Click Index
End Sub




Private Sub SetChemicalsXRecipe()
    
    Call SetDatabaseComponentGrid(GridChemicals)
    Call SetCodeGrid(Grid1)
    
    Grid1.Column(6).Width = 0
    Grid1.ReadOnly = True

End Sub




Public Function SetUM(ByVal Cmb As ComboBox) As Boolean

    With Cmb
        .Clear
        .AddItem "L"
        .AddItem "kg"
        .AddItem "pcs"
        .AddItem "mL"
        .AddItem "g"
        .AddItem "mg"
        .ListIndex = 5
    End With
End Function

Public Function SetUMPeso(ByVal Cmb As ComboBox) As Boolean

    With Cmb
        .Clear
        .AddItem "kg"
        .AddItem "g"
        .AddItem "mg"
        .ListIndex = 1
    End With
End Function



Private Sub CheckTotalsAndPercentage()
Dim rc As Boolean
Dim strPerc As String
Dim Totali As Double
Dim bUmMassa As Boolean

On Error GoTo ERR_CHECK:

    
    If Me.Visible = False Then Exit Sub
    If uRecipe.Density = 0 Then uRecipe.Density = 1
    
    
    Select Case uRecipe.bUmMassa
        Case True
        
            If SetbUmMassa(uRecipe.UmMultiple) Then
                Totali = uRecipe.Multiple
            Else
                Totali = (uRecipe.Multiple / uRecipe.Density)
            End If
            
            uRecipe.TotalWeightKg = Totali
            uRecipe.UmTotalWeightKg = uRecipe.UmMultiple
            uRecipe.UmTotalWeightL = SetUmVolume(uRecipe.UmMultiple)
            
            
            
            
            Label6 = "Total Weight ( " & (uRecipe.UmMultiple) & " )"
            Label7 = "Total Weight ( " & SetUmVolume(uRecipe.UmMultiple) & " )"
            
    
            
    
            
        Case False
            Totali = uRecipe.Multiple * uRecipe.Density
            uRecipe.TotalWeightKg = Totali
            
            uRecipe.UmTotalWeightKg = SetUmVolume(uRecipe.UmMultiple)
            uRecipe.UmTotalWeightL = uRecipe.UmMultiple
            
            Label6 = "Total Weight ( " & uRecipe.UmTotalWeightKg & " )"
            Label7 = "Total Weight ( " & uRecipe.UmMultiple & " )"
            
            
            
            
            
    End Select
    
    'Dim Perc As Double
    Dim i As Integer
    For i = 1 To GridChemicals.Rows - 1
        strPerc = GridChemicals.Cell(i, 6).Text
        If strPerc = "" Then Exit For
        Call SetComponentWeightByPerc(CDbl(strPerc), i)
    Next
    
    'rc = CheckPercentageByWeight(GridChemicals, strPerc, Totali * Um(uRecipe.UmTotalWeightKg))
    
    
    'PicPerc.BackColor = IIf(rc, &H8000&, &H40C0&)
    'CheckPercentage


ERR_END:
    On Error GoTo 0
    Totali = FormatNumber(Totali, iVirgola(Totali))
    
        
    lbTotalWKg = Totali
    lbTotalWL = FormatNumber((Totali / uRecipe.Density), iVirgola(Totali))
    Exit Sub

ERR_CHECK:
    MsgBox "CheckTotalsAndPercentage ERROR"
    GoTo ERR_END:
End Sub

Private Function CalculateLastPercentage(ByRef Row As Long) As String
Dim rc As Boolean
Dim i As Integer
Dim TotalPerc As Double

    If lbPerc <> "" Then

        TotalPerc = CDbl(Replace(lbPerc, "%", ""))
    
    
        CalculateLastPercentage = ""
        Row = 0
        
        With GridChemicals
            If .Rows < 2 Then Exit Function
            For i = 1 To .Rows - 1
                If .Cell(i, 6).Text = "" Then

                    CalculateLastPercentage = FormatNumber(100 - TotalPerc, 4)
                    Row = i
                    Exit Function
                End If
            Next
            TotalPerc = 0
            For i = 1 To .Rows - 2
               
                    TotalPerc = TotalPerc + Replace(.Cell(i, 6).Text, "%", "")
                    CalculateLastPercentage = FormatNumber(100 - TotalPerc, 4)
                   
                
                
               
            Next
             Row = .Rows - 1
            
        End With
    End If
    
End Function

Private Function SetComponentWeightByPerc(ByVal Perc As String, ByVal Row As Long)

Dim TotalW As Double
Dim Qty As Double
Dim UmComponent As String

   
   If bSetPercentageLastComponent Then
        bSetPercentageLastComponent = False
        Perc = CalculateLastPercentage(Row)
        
     
    End If
    
    If Perc = "" Then
            
            GridChemicals.Cell(Row, 4).Text = ""
            
            GridChemicals.Cell(Row, 6).Text = ""
            Call CheckPercentage
        Exit Function
    End If
    
    

    With uRecipe
        
        UmComponent = GridChemicals.Cell(Row, 5).Text
        
        If UmComponent = "" Then
        
            UmComponent = .UmMultiple
            GridChemicals.Cell(Row, 5).Text = .UmMultiple
            
        End If
    
        TotalW = .TotalWeightKg * Um(.UmTotalWeightKg)
        
        If TotalW = 0 Then
            PopupMessage 2, "Warning : Please check Recipes Specification...", , True, .Code
            Exit Function
        End If
        
    Select Case uRecipe.bUmMassa
        Case True
            If SetbUmMassa(uRecipe.UmMultiple) Then
            Else
                TotalW = TotalW * uRecipe.Density
            End If
        
            
        Case False
        
            'If SetbUmMassa(uRecipe.UmMultiple) Then
            
           ' Else
                'TotalW = TotalW / uRecipe.Density
           ' End If
            
          
            
    End Select
        
        
        If IsNumeric(Perc) Or Perc <> "" Then
            Qty = FormatNumber((Perc * (TotalW / 100) / Um(UmComponent)), 3)
            GridChemicals.Cell(Row, 4).Text = Qty
            GridChemicals.Cell(Row, 4).Alignment = cellRightCenter
            
            GridChemicals.Cell(Row, 6).Text = FormatNumber(Perc, 4)
            GridChemicals.Cell(Row, 6).Alignment = cellCenterCenter
            
            Call CheckPercentage
        End If
        
    End With
    
    
    
End Function



Private Sub CheckPercentage()
Dim i As Integer
Dim rc As Boolean
Dim TotPerc As Double
Dim strPerc As String
strPerc = ""
rc = False
SetPercentageLastComponent.Visible = False
With GridChemicals
    If .Rows > 1 Then
        SetPercentageLastComponent.Visible = True
        For i = 1 To .Rows - 1
            If .Cell(i, 6).Text = "" Then
                rc = False
                Exit For
            End If
            TotPerc = TotPerc + CDbl(.Cell(i, 6).Text)
        Next
        rc = IIf((TotPerc) = 100, True, False)
        strPerc = TotPerc
    End If

End With

PicPerc.BackColor = IIf(rc, &H8000&, &H40C0&)

lbPerc = strPerc & " %"


End Sub



Private Sub SaveRecipe()
Dim rc As Boolean
If CheckChemicalsInRecipe(GridChemicals) Then

   
    rc = SaveFullRecipe(GridChemicals, uRecipe.Code)
    
    If rc Then
         PopupMessage 2, "Formulation Recipe correctly saved...", , , uRecipe.Code
         
        
    Else
        GoTo err
    End If
    
Else
err:
    PopupMessage 2, "Check all Chemicals and Weights before saving...", , True, uRecipe.Code
End If
End Sub

Private Sub SetComponentComponentRecipe()

    Dim strRecipe As String
    
    
        strRecipe = SelectedCode
        
        lbInside(4) = "Components : " & strRecipe

        Call SetComponentGrid(Grid3)
        Grid3.Column(7).Width = 0
        Grid3.Column(8).Width = 0
        Call AddComponentGrid(Grid3, strRecipe)
        
        Frame4.Visible = IIf(Grid3.Rows > 1, False, True)
        
        frInside(5).Top = frInside(3).Top
        frInside(5).Left = frInside(3).Left
        frInside(5).Width = frInside(3).Width
        frInside(5).ZOrder
        frInside(5).Visible = True
        ucScrollAdd1.UCScrollV.ScrollToValue frInside(5).Top - 480
         
End Sub

