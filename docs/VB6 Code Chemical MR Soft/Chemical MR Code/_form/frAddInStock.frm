VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form frAddInStock 
   BackColor       =   &H80000005&
   Caption         =   "Warehouse"
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
   Icon            =   "frAddInStock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12000
   ScaleMode       =   0  'User
   ScaleWidth      =   19320.76
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   62
      Top             =   0
      Width           =   19215
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
         TabIndex        =   63
         Top             =   0
         Width           =   2175
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Warehouse"
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
            TabIndex        =   64
            Top             =   640
            Width           =   2070
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   840
            MousePointer    =   99  'Custom
            Picture         =   "frAddInStock.frx":33E2
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.Label lbChemicalMR 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MR09"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   4440
         TabIndex        =   67
         Top             =   0
         Width           =   10455
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Control"
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
         Left            =   16275
         TabIndex        =   66
         Top             =   195
         Width           =   2655
      End
      Begin VB.Label lbWait 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "Wait : Loading Data..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5760
         TabIndex        =   65
         Top             =   360
         Visible         =   0   'False
         Width           =   7575
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
      TabIndex        =   55
      Top             =   11040
      Width           =   19215
      Begin VB.Timer TimerBeginForm 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   8400
         Top             =   120
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
         TabIndex        =   61
         Top             =   -120
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
         TabIndex        =   60
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
         Index           =   0
         Left            =   8760
         TabIndex        =   59
         Top             =   -120
         Width           =   1695
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   4
         Left            =   18240
         MousePointer    =   99  'Custom
         Picture         =   "frAddInStock.frx":5DD4
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   3
         Left            =   15600
         MousePointer    =   99  'Custom
         Picture         =   "frAddInStock.frx":91B6
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   0
         Left            =   9360
         MousePointer    =   99  'Custom
         Picture         =   "frAddInStock.frx":C598
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Stock Conrol"
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
         TabIndex        =   58
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
         Left            =   15345
         MousePointer    =   99  'Custom
         TabIndex        =   57
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
         TabIndex        =   56
         Top             =   660
         Width           =   1200
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
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   675
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
         TabIndex        =   54
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
         TabIndex        =   53
         Top             =   80
         Width           =   330
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
      Left            =   120
      ScaleHeight     =   9975
      ScaleWidth      =   19245
      TabIndex        =   12
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
         Height          =   52000
         Left            =   0
         ScaleHeight     =   52000
         ScaleMode       =   0  'User
         ScaleWidth      =   19155
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   -30000
         Width           =   19155
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   9495
            Index           =   2
            Left            =   480
            TabIndex        =   76
            Top             =   31440
            Width           =   18015
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   14
               Left            =   14640
               TabIndex        =   129
               Top             =   3720
               Width           =   1215
            End
            Begin VB.Frame frPrinQRCode 
               BackColor       =   &H00000080&
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
               Left            =   1200
               TabIndex        =   119
               Top             =   8040
               Width           =   4575
               Begin VB.Label lbPrinQRCode 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Print QRCode"
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
                  TabIndex        =   120
                  Top             =   120
                  Width           =   4575
               End
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   12
               Left            =   13920
               TabIndex        =   117
               Top             =   6120
               Width           =   1695
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   11
               Left            =   10440
               TabIndex        =   115
               Top             =   6120
               Width           =   1695
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   10
               Left            =   7200
               TabIndex        =   113
               Top             =   6120
               Width           =   1695
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   9
               Left            =   3600
               TabIndex        =   110
               Top             =   6120
               Width           =   1695
            End
            Begin VB.ComboBox cmbStatus 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   9840
               TabIndex        =   109
               Text            =   "Combo1"
               Top             =   1800
               Visible         =   0   'False
               Width           =   3255
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
               Left            =   12240
               TabIndex        =   92
               Top             =   8040
               Width           =   3855
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "MR Stock Table"
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
                  TabIndex        =   93
                  Top             =   120
                  Width           =   3855
               End
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   3
               Left            =   14640
               TabIndex        =   91
               Top             =   3000
               Width           =   1815
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   2
               Left            =   10680
               TabIndex        =   90
               Top             =   3000
               Width           =   1695
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   24
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00886010&
               Height          =   585
               Index           =   8
               Left            =   11520
               Locked          =   -1  'True
               TabIndex        =   89
               Text            =   "34F"
               Top             =   4720
               Width           =   1455
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   0
               Left            =   1800
               TabIndex        =   88
               Top             =   3000
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
               Index           =   6
               Left            =   5880
               TabIndex        =   86
               Top             =   8040
               Width           =   6255
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
                  Index           =   6
                  Left            =   0
                  TabIndex        =   87
                  Top             =   120
                  Width           =   6255
               End
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   4
               Left            =   3120
               TabIndex        =   85
               Top             =   3720
               Width           =   1455
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   5
               Left            =   8280
               TabIndex        =   84
               Top             =   3720
               Width           =   1215
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   1
               Left            =   6240
               TabIndex        =   83
               Top             =   3000
               Width           =   2895
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   7
               Left            =   6840
               TabIndex        =   82
               Top             =   4920
               Width           =   1455
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
               Left            =   1080
               TabIndex        =   79
               Top             =   2160
               Width           =   15255
               Begin VB.Line Line4 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "MR Stock Specifics"
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
                  Index           =   1
                  Left            =   0
                  TabIndex        =   81
                  Top             =   120
                  Width           =   2085
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "MR Warehouse"
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
                  Left            =   13815
                  TabIndex        =   80
                  Top             =   180
                  Width           =   1335
               End
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   13
               Left            =   1920
               TabIndex        =   78
               Top             =   7080
               Width           =   14535
            End
            Begin VB.TextBox txStock 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   6
               Left            =   11760
               TabIndex        =   77
               Top             =   3720
               Width           =   1215
            End
            Begin VB.Label lbUUnitEdit 
               BackStyle       =   0  'Transparent
               Caption         =   "mg/L"
               Height          =   255
               Left            =   15960
               TabIndex        =   131
               Top             =   3720
               Width           =   975
            End
            Begin VB.Label lbStock 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "U  ±"
               Height          =   255
               Index           =   14
               Left            =   14160
               TabIndex        =   130
               Top             =   3720
               Width           =   375
            End
            Begin VB.Label lbReduction 
               BackStyle       =   0  'Transparent
               Caption         =   "* 120 days Reduction"
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
               Left            =   10440
               TabIndex        =   123
               Top             =   6540
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.Label lbStock 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "STATUS"
               Height          =   255
               Index           =   12
               Left            =   12960
               TabIndex        =   118
               Top             =   6120
               Width           =   585
            End
            Begin VB.Label lbStock 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "MR EXP"
               Height          =   255
               Index           =   11
               Left            =   9480
               TabIndex        =   116
               Top             =   6120
               Width           =   690
            End
            Begin VB.Label lbStock 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Finished Date"
               Height          =   255
               Index           =   10
               Left            =   5760
               TabIndex        =   114
               Top             =   6120
               Width           =   1290
            End
            Begin VB.Label lbUM 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00644603&
               Caption         =   "mg/L"
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   8400
               TabIndex        =   112
               Top             =   4920
               Width           =   735
            End
            Begin VB.Label lbStock 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Open Date"
               Height          =   255
               Index           =   9
               Left            =   2280
               TabIndex        =   111
               Top             =   6120
               Width           =   1080
            End
            Begin VB.Label lbStock 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Supplier EXP ( Date )"
               Height          =   255
               Index           =   3
               Left            =   12600
               TabIndex        =   108
               Top             =   3000
               Width           =   1950
            End
            Begin VB.Label lbStock 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date Arrived"
               Height          =   255
               Index           =   2
               Left            =   9360
               TabIndex        =   107
               Top             =   3000
               Width           =   1215
            End
            Begin VB.Label lbStock 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00644603&
               Caption         =   "Bottle # "
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   8
               Left            =   10335
               TabIndex        =   106
               Top             =   4920
               Width           =   825
            End
            Begin VB.Label lbStock 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Lot"
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   105
               Top             =   3000
               Width           =   300
            End
            Begin VB.Label lbStock 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Purity"
               Height          =   255
               Index           =   4
               Left            =   2400
               TabIndex        =   104
               Top             =   3720
               Width           =   510
            End
            Begin VB.Label lbStock 
               BackStyle       =   0  'Transparent
               Caption         =   "MR value ( concentration )"
               Height          =   255
               Index           =   5
               Left            =   5520
               TabIndex        =   103
               Top             =   3720
               Width           =   2775
            End
            Begin VB.Label lbStock 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Location"
               Height          =   255
               Index           =   1
               Left            =   5280
               TabIndex        =   102
               Top             =   3000
               Width           =   855
            End
            Begin VB.Label lbStock 
               Alignment       =   2  'Center
               BackColor       =   &H00644603&
               Caption         =   "Bottle Qty ( volume )"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   7
               Left            =   4335
               TabIndex        =   101
               Top             =   4920
               Width           =   2445
            End
            Begin VB.Label lbStock 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Note"
               Height          =   255
               Index           =   13
               Left            =   1320
               TabIndex        =   100
               Top             =   7080
               Width           =   480
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fill All Specifics and Save to Add Bottles in Stock"
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
               Index           =   2
               Left            =   1920
               TabIndex        =   99
               Top             =   7560
               Width           =   4335
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               BackColor       =   &H00886010&
               Caption         =   "Edit Stock Entry"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   26.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   735
               Left            =   0
               TabIndex        =   98
               Top             =   480
               Width           =   18015
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "mg/L"
               Height          =   255
               Left            =   9720
               TabIndex        =   97
               Top             =   3720
               Width           =   975
            End
            Begin VB.Label lbStock 
               BackStyle       =   0  'Transparent
               Caption         =   "Density"
               Height          =   255
               Index           =   6
               Left            =   10800
               TabIndex        =   96
               Top             =   3720
               Width           =   855
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "mg/L"
               Height          =   255
               Left            =   13200
               TabIndex        =   95
               Top             =   3720
               Width           =   975
            End
            Begin VB.Shape Shape3 
               BorderColor     =   &H00C0C0C0&
               Height          =   855
               Left            =   4080
               Top             =   4680
               Width           =   10095
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   255
               Left            =   4800
               TabIndex        =   94
               Top             =   3720
               Width           =   150
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
            Left            =   6960
            TabIndex        =   50
            Top             =   240
            Width           =   5535
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Change Selected Chemical MR"
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
               Height          =   375
               Index           =   0
               Left            =   0
               TabIndex        =   51
               Top             =   75
               Width           =   5535
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
            Height          =   7455
            Index           =   0
            Left            =   840
            TabIndex        =   36
            Top             =   2280
            Width           =   18255
            Begin VB.Frame frReadQR 
               BackColor       =   &H00000080&
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
               Left            =   3120
               TabIndex        =   121
               Top             =   6360
               Width           =   3015
               Begin VB.Label lbReadQR 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Acquire From QRCode"
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
                  TabIndex        =   122
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
               Left            =   6360
               TabIndex        =   46
               Top             =   2760
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
                  TabIndex        =   47
                  Top             =   555
                  Width           =   1215
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
               TabIndex        =   43
               Top             =   0
               Width           =   18015
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Double Click Bottle to Edit/Modify data"
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
                  Left            =   14055
                  TabIndex        =   45
                  Top             =   180
                  Width           =   3480
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "MR Stock Table : List of all Bottles in Warehouse"
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
                  TabIndex        =   44
                  Top             =   75
                  Width           =   6375
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   0
                  X2              =   17640
                  Y1              =   480
                  Y2              =   480
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00008000&
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
               Left            =   0
               TabIndex        =   41
               Top             =   6360
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Add In Stock"
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
                  TabIndex        =   42
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
               Index           =   2
               Left            =   11520
               TabIndex        =   39
               Top             =   6360
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Delete Single Entry"
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
                  TabIndex        =   40
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
               Index           =   4
               Left            =   14640
               TabIndex        =   37
               Top             =   6360
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
                  Index           =   4
                  Left            =   0
                  TabIndex        =   38
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin FlexCell.Grid Grid1 
               Height          =   5415
               Left            =   0
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   600
               Width           =   17535
               _ExtentX        =   30930
               _ExtentY        =   9551
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
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Select a Chemical MR, add in stock or delete entries"
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
               Left            =   12960
               TabIndex        =   49
               Top             =   7080
               Width           =   4650
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00D0D0D0&
               X1              =   0
               X2              =   17640
               Y1              =   6120
               Y2              =   6120
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8175
            Index           =   1
            Left            =   360
            TabIndex        =   19
            Top             =   21600
            Width           =   18015
            Begin VB.Frame frStockLabel 
               BackColor       =   &H00000080&
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
               Left            =   1200
               TabIndex        =   132
               Top             =   6360
               Visible         =   0   'False
               Width           =   4575
               Begin VB.Label lbStockLabel 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Print Label"
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
                  TabIndex        =   133
                  Top             =   120
                  Width           =   4575
               End
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   10
               Left            =   14640
               TabIndex        =   126
               Top             =   3720
               Width           =   1455
            End
            Begin VB.Timer Timer1 
               Interval        =   60
               Left            =   600
               Top             =   1440
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   6
               Left            =   11640
               TabIndex        =   6
               Top             =   3720
               Width           =   1215
            End
            Begin VB.ComboBox cmbLocation 
               BackColor       =   &H00E0E0E0&
               Height          =   375
               Left            =   6000
               TabIndex        =   69
               Text            =   "Combo1"
               Top             =   1920
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.ComboBox cbUM 
               BackColor       =   &H00644603&
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   7800
               TabIndex        =   11
               Text            =   "cbUM"
               Top             =   4560
               Width           =   855
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   9
               Left            =   1920
               TabIndex        =   9
               Top             =   5520
               Width           =   14535
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
               TabIndex        =   23
               Top             =   2160
               Width           =   15255
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "MR Warehouse"
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
                  Left            =   13815
                  TabIndex        =   25
                  Top             =   180
                  Width           =   1335
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "MR Stock Specifics"
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
                  TabIndex        =   24
                  Top             =   120
                  Width           =   2085
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
               Height          =   300
               Index           =   7
               Left            =   6360
               TabIndex        =   7
               Top             =   4560
               Width           =   1335
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   1
               Left            =   6240
               TabIndex        =   1
               Top             =   3000
               Width           =   2895
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   5
               Left            =   8280
               TabIndex        =   5
               Top             =   3720
               Width           =   1215
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   4
               Left            =   3120
               TabIndex        =   4
               Top             =   3720
               Width           =   1455
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
               Index           =   3
               Left            =   5880
               TabIndex        =   10
               Top             =   6360
               Width           =   6255
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
                  Index           =   3
                  Left            =   0
                  TabIndex        =   22
                  Top             =   120
                  Width           =   6255
               End
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   0
               Left            =   1800
               TabIndex        =   0
               Top             =   3000
               Width           =   3255
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   8
               Left            =   12240
               TabIndex        =   8
               Top             =   4560
               Width           =   735
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   2
               Left            =   10680
               TabIndex        =   2
               Top             =   3000
               Width           =   1695
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   3
               Left            =   14640
               TabIndex        =   3
               Top             =   3000
               Width           =   1815
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
               TabIndex        =   20
               Top             =   6360
               Width           =   3855
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Goto MR Stock Table"
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
                  TabIndex        =   21
                  Top             =   120
                  Width           =   3855
               End
            End
            Begin VB.Label lbUUnit 
               BackStyle       =   0  'Transparent
               Caption         =   "mg/L"
               Height          =   255
               Left            =   16200
               TabIndex        =   128
               Top             =   3720
               Width           =   975
            End
            Begin VB.Label lbFormulation 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "U  ±"
               Height          =   255
               Index           =   10
               Left            =   13680
               TabIndex        =   127
               Top             =   3720
               Width           =   855
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               Height          =   255
               Left            =   4800
               TabIndex        =   75
               Top             =   3720
               Width           =   150
            End
            Begin VB.Shape Shape2 
               BorderColor     =   &H00C0C0C0&
               Height          =   855
               Left            =   3600
               Top             =   4320
               Width           =   10095
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "mg/L"
               Height          =   255
               Left            =   12960
               TabIndex        =   74
               Top             =   3720
               Width           =   975
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "Density"
               Height          =   255
               Index           =   6
               Left            =   10680
               TabIndex        =   73
               Top             =   3720
               Width           =   855
            End
            Begin VB.Label lbUmconcetration 
               BackStyle       =   0  'Transparent
               Caption         =   "mg/L"
               Height          =   255
               Left            =   9720
               TabIndex        =   70
               Top             =   3720
               Width           =   975
            End
            Begin VB.Label lbStoc 
               Alignment       =   2  'Center
               BackColor       =   &H00664805&
               Caption         =   "Add In Stock"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   36
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   975
               Index           =   0
               Left            =   0
               TabIndex        =   68
               Top             =   480
               Width           =   18015
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fill All Specifics and Save to Add Bottles in Stock"
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
               Left            =   1080
               TabIndex        =   35
               Top             =   6000
               Width           =   4335
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Note"
               Height          =   255
               Index           =   9
               Left            =   1320
               TabIndex        =   34
               Top             =   5520
               Width           =   480
            End
            Begin VB.Label lbFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00644603&
               Caption         =   "Bottle Qty ( volume )"
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   7
               Left            =   3960
               TabIndex        =   33
               Top             =   4560
               Width           =   2235
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Location"
               Height          =   255
               Index           =   1
               Left            =   5280
               TabIndex        =   32
               Top             =   3000
               Width           =   855
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "MR value ( concentration )"
               Height          =   255
               Index           =   5
               Left            =   5520
               TabIndex        =   31
               Top             =   3720
               Width           =   2775
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Purity"
               Height          =   255
               Index           =   4
               Left            =   2400
               TabIndex        =   30
               Top             =   3720
               Width           =   510
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "Lot"
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   29
               Top             =   3000
               Width           =   975
            End
            Begin VB.Label lbFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00644603&
               Caption         =   "# Bottles Arrived ( number )"
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   8
               Left            =   9120
               TabIndex        =   28
               Top             =   4560
               Width           =   2985
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date Arrived"
               Height          =   255
               Index           =   2
               Left            =   9360
               TabIndex        =   27
               Top             =   3000
               Width           =   1215
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Supplier EXP ( Date )"
               Height          =   255
               Index           =   3
               Left            =   12600
               TabIndex        =   26
               Top             =   3000
               Width           =   1950
            End
         End
         Begin VB.Frame frQuantityCheck 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   1680
            TabIndex        =   14
            Top             =   1200
            Width           =   15135
            Begin VB.PictureBox PicMax 
               BackColor       =   &H000000C0&
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
               Height          =   300
               Left            =   14320
               ScaleHeight     =   272.727
               ScaleMode       =   0  'User
               ScaleWidth      =   615
               TabIndex        =   15
               Top             =   225
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label lbStockMinQty 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "--.--"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00644603&
               Height          =   300
               Left            =   10440
               TabIndex        =   125
               Top             =   225
               Width           =   2055
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               BackColor       =   &H00644603&
               Caption         =   "Min Q.ty"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   9000
               TabIndex        =   124
               Top             =   225
               Width           =   1455
            End
            Begin VB.Label lbStockBottles 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "--.--"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00644603&
               Height          =   300
               Left            =   6795
               TabIndex        =   72
               Top             =   225
               Width           =   1950
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackColor       =   &H00644603&
               Caption         =   "Bottles #"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   4920
               TabIndex        =   71
               Top             =   225
               Width           =   1920
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               BackColor       =   &H00644603&
               Caption         =   "MR Stock Quantity"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   480
               TabIndex        =   18
               Top             =   225
               Width           =   2175
            End
            Begin VB.Label Label15 
               Alignment       =   2  'Center
               BackColor       =   &H00644603&
               Caption         =   "Qty Ceck"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   13080
               TabIndex        =   17
               Top             =   225
               Width           =   1215
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H00B0B0B0&
               Height          =   855
               Left            =   0
               Top             =   0
               Width           =   15135
            End
            Begin VB.Label lbStockQTY 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "--.--"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00644603&
               Height          =   300
               Left            =   2640
               TabIndex        =   16
               Top             =   240
               Width           =   2055
            End
         End
      End
   End
   Begin VB.Line Line1 
      X1              =   9056.606
      X2              =   10264.15
      Y1              =   5760
      Y2              =   6240
   End
End
Attribute VB_Name = "frAddInStock"
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


Private lRow As Long
Private lCol As Long


Private SettingName As String

Private bfrInsideMoveTop As Boolean

Private bCancelUpdate As Boolean

Private uMR As MRType
Private uWarehouseEntry As WareHouseEntry
Private uWarehouseEntryClean As WareHouseEntry
Private WarehouseEntries() As WareHouseEntry
Private WarehouseEntriesClean() As WareHouseEntry
Private UserBarcode As Barcode
Private bCopyEntry As Boolean
Private bViewWarehouseEntry As Boolean


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


Public Function DoShow(Optional ByVal ID As Long, Optional ByVal UserMRcode As String) As Boolean

    On Error GoTo ERR_SHOW
    
    m_rc = False
    mOk
    
    
    ' UserMRcode
    If UserMRcode <> "" Then
        uMR.Code = UserMRcode
    End If
    
    TimerBeginForm.Enabled = True
    

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












Private Sub Grid1_BeforeUserSort(ByVal Col As Long)
lRow = 0
End Sub

Private Sub Grid1_DblClick()

    If lRow = 0 Then Exit Sub


    Call ViewEntry
   
        
End Sub

Private Sub Grid1_OwnerDrawCell(ByVal Row As Long, ByVal Col As Long, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Handled As Boolean)

frReadQR.Visible = IIf(Grid1.Rows > 1, True, False)
frCommandInside(4).Visible = frReadQR.Visible

End Sub

Private Sub Grid1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

Dim rc As Boolean
lRow = FirstRow

uWarehouseEntry = MyWareHouseEntryClean

frCommandInside(2).Visible = False
'frCommandInside(5).Visible = False

If lRow > 0 Then
    With UserBarcode
        .Code = Grid1.Cell(lRow, 1).Text
        .Bottle = Grid1.Cell(lRow, 4).Text
        .Lot = Grid1.Cell(lRow, 6).Text
        .Date = Grid1.Cell(lRow, 13).Text
    End With
    
    With uWarehouseEntry
        .ID = Grid1.Cell(lRow, 20).Text
        .MRCode = Grid1.Cell(lRow, 1).Text
        ReDim .Bottle(0)
        .Bottle(0) = Grid1.Cell(lRow, 4).Text
        .Lot = Grid1.Cell(lRow, 6).Text
        .ArrivedTime = Grid1.Cell(lRow, 13).Text
    End With
    
    frCommandInside(2).Visible = True

    

End If

   

End Sub

Private Sub lbPrinQRCode_Click()
frPrinQRCode_Click
End Sub

Private Sub lbStockLabel_Click()
frStockLabel_Click
End Sub

Private Sub Timer1_Timer()


Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()

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
    
 
    
    For i = txFormulation.LBound To txFormulation.UBound
        txFormulation(i) = ""
        
    Next
 


End Sub
Private Sub InitForm()



 '00
 ' "00"
 On Error GoTo ERR_INIT
    
    
    SelectedCode = ""
  
    lRow = 0
    lCol = 0
   
    lbChemicalMR = ""
    
    
   ' "00.1"
    Dim Grid(10) As Grid
    
    Set Grid(0) = Grid1
   
    
      ' "00.2"
    Call SetAllRecipeForProductionGrid(Grid())
    Call SetColumnWidth
    
      ' "00.3"
    Grid1.FrozenCols = 2
   
    
      ' "00.4"
      
      ' uMR.Code
    
    If uMR.Code <> "" Then
        FillChemicalMRTable (uMR.Code)
    
    End If
    
      ' "00.5"
ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_INIT:
    MsgBox Err.Description
    Resume ERR_END:
    
    

End Sub
Private Sub Form_Load()

   Me.WindowState = MainWindowState

    PBContainer.Top = 0
    PBContainer.Left = 0
    PBContainer.Width = Me.Width
   
    
    

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
  '  ucScrollAdd1.ContainerW = Me.ScaleWidth
    'But also need to resize PBContainer wich hide the width of the bottom box

    
    
      ResizeControls

    SetColumnWidth
    
   ' MainWindowState = Me.WindowState
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERR_FORM:
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    
   
    
    Set frAddInStock = Nothing
    
   
ERR_END:
    Exit Sub
ERR_FORM:
    MsgBox Err.Description
    GoTo ERR_END
End Sub



Private Sub txFormulation_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case KeyAscii
        Case 13
            If Index < txFormulation.UBound Then
                txFormulation(Index + 1).SetFocus
            Else
                txFormulation(0).SetFocus
            End If
End Select
End Sub





Private Sub ucScrollAdd1_ScrollH(Value As Long)
    Form_Resize
End Sub
Private Sub PicHover_Click()
PBContainer.Top = -0
End Sub
Private Sub lblHoverClick_Click()
    PBContainer.Top = -0
    
End Sub
Private Sub imOver_Click()
PBContainer.Top = -0
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
        If F_MsgBox.DoShow("Quit WareHouse?") Then
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            
            Unload Me
        
        End If
        
    Case 3
        ' Previous
         If IndexVisibleFrame >= 1 Then
            MyIndex = IndexVisibleFrame - 1
            If frInside(MyIndex).Visible = False Then
                MyIndex = IndexVisibleFrame - 3
            End If
            
            PBContainer.Top = -frInside(MyIndex).Top + 480
        Else
            PBContainer.Top = -0
         End If
    
    
    
    Case 4
        ' forward
        If IndexVisibleFrame < frInside.UBound Then
            MyIndex = IndexVisibleFrame + 1
            If frInside(MyIndex).Visible = False Then
                MyIndex = IndexVisibleFrame + 3
            End If
            PBContainer.Top = -frInside(MyIndex).Top + 480
        Else
            PBContainer.Top = -0
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

Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
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



PBFooter.ZOrder


End Function




Private Sub PicMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
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
            ' select MR
            SetMRInTable
        Case 1
            ' Add In Stock
            
           
            
            Call AddInStockFrame(False)
        Case 2
            ' delete single entry
            Call DeleteSingleEntry
            
        Case 3
            ' Save
            Call SaveEntry
           
        Case 4
            ' DELETE ALL
            
              Call DeleteAll
        Case 5, 7
            ' exit specifics
            Call FillChemicalMRTable(uMR.Code)
            PBContainer.Top = -0
        Case 6
          Call ModifyEntry

        Case 8
            ' cancella Hanna code
       
        Case 9
            PBContainer.Top = -frInside(1).Top + 480
            
        Case 10
            ' material requisition ALL RECIPE
         
        Case 11
            ' material requisition single Recipe
          
        Case 12
        
          
            
        Case 13
          
            
        Case 14
            PBContainer.Top = -frInside(1).Top + 480
                
    End Select
End Sub


Private Sub frInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)


Dim i As Integer
    For i = 0 To frCommandInside.UBound

            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 3 Or i = 1 Or i = 6 Then
                frCommandInside(i).BackColor = &H8000&
            End If

    
    Next
 
 
End Sub

Private Sub frCommandInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
IndexDashCommInside = Index
Dim i As Integer
    For i = 0 To frCommandInside.UBound
        If i = Index Then
            frCommandInside(i).BackColor = &H846623
            lbCommandInside(i).ForeColor = vbWhite
            If i = 3 Or i = 1 Or i = 6 Then
                frCommandInside(i).BackColor = &H20A020
            End If
            
        Else
            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 3 Or i = 1 Or i = 6 Then
                frCommandInside(i).BackColor = &H8000&
            End If
        End If
    
    Next
 
 
End Sub
Private Sub lbCommandInside_Click(Index As Integer)

frCommandInside_Click Index
End Sub
Private Sub lbCommandInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
frCommandInside_MouseMove Index, Button, Shift, x, y
End Sub
Private Sub PBTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Form_MouseDown Button, Shift, x, y
End Sub

Private Sub PBTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Form_MouseMove Button, Shift, x, y
End Sub

Private Sub PBTitle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
FrmMove = False
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMove = True
    DragX = x
    DragY = y
    If Me.WindowState = 2 Then
        FrmMove = False
       
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nx, ny
    If Me.WindowState = 2 Then
        FrmMove = False
        Exit Sub
    End If
    nx = Me.Left + x - DragX
    ny = Me.Top + y - DragY
    Me.Left = nx
    Me.Top = ny
    FrmMove = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer



If Me.WindowState = 2 Then
    FrmMove = False
End If
Dim nx, ny
    If FrmMove Then
        nx = Me.Left + x - DragX
        ny = Me.Top + y - DragY
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












Private Sub SetMRInTable()

Dim MRCode As String
    

If FormChemicalMR.DoShow(MRCode) Then

    Call FillChemicalMRTable(MRCode)
 
End If




End Sub

Private Function SetForm(ByVal MRCode As String)

On Error GoTo ERR_SET:

     lbChemicalMR = MRCode
     PicMax.Visible = False
     lbStockQTY = "--.--"
     lbStockBottles = "--"
     lbStockMinQty = "--.--"
     uMR = MRTypeClean
     uWarehouseEntry = MyWareHouseEntryClean
     
     
     '01
     ' MRCode
     
     If SetMR(MRCode, uMR) Then
    
    '02
     ' "02"
     
         If (uMR.MinQty) <> "" Then
         
            If uMR.STOCK_QTY - uMR.MinQty > 0 Then
            
                PicMax.BackColor = vbColorGreen
            Else
                PicMax.BackColor = vbColorRed
            
            End If
        End If
        
        PicMax.Visible = True
        
        lbStockMinQty = uMR.MinQty & " " & uMR.STOCK_UNIT
        lbStockQTY = uMR.STOCK_QTY & " " & uMR.STOCK_UNIT
        lbStockBottles = Grid1.Rows - 1
     Else
     
     
     
     
     End If
 
ERR_END:
 On Error GoTo 0
 Exit Function
ERR_SET:
 MsgBox Err.Description
 Resume Next

End Function

Private Function FillChemicalMRTable(ByVal UserMRcode As String, Optional ByVal bClosed As Boolean)
lbChemicalMR = UserMRcode
Call GetStockFromDatabase(Grid1, bClosed, UserMRcode)
Call SetForm(UserMRcode)
End Function




Private Function DeleteSingleEntry()
Dim rc As Boolean
        rc = DeleteWharehouseEntry(uWarehouseEntry.ID, Grid1)
        
        
        If rc Then
            Call FillChemicalMRTable(uWarehouseEntry.MRCode)
        
            
            
            PBContainer.Top = -0
             
            PopupMessage 2, "Bottle deleted from Warehouse...", , , uWarehouseEntry.MRCode
        End If
     
     
     
End Function
Private Function DeleteAll()
Dim rc As Boolean

rc = False

    If F_MsgBox.DoShow("Delete all bottles?", uMR.Code) Then
    Else
        Exit Function
    End If
     
     
   
    dbChemicalMR.Execute "DELETE * FROM TabMRWarehouse where Code=" & uMR.Code


     'With dbTabMRWarehouse
    '
    '    .filter = ""
    '    .filter = "Code='" & uMR.Code & "'"
    '    If .EOF Then
     '   Else
     '       rc = True
     '       .MoveFirst
     '       Do
     ''           .Delete
      '          .MoveNext
     '       Loop Until .EOF
      '
      ''
      '  End If
    '
     
     'End With
        
        
        
        If rc Then
            Grid1.Rows = 1
        
            PopupMessage 2, "All Bottles deleted from Warehouse...", , , uWarehouseEntry.MRCode
        End If
     

End Function


Private Function ViewEntry()
Dim i As Integer
    For i = txStock.LBound To txStock.UBound
        txStock(i) = ""
    Next
    
    Call SetCmbStatus(cmbStatus)
    
    bViewWarehouseEntry = True

    ViewWarehouseEntry
    
    
     txFormulation(8).Locked = True
     
      PBContainer.Top = -frInside(2).Top + 480
     
     PopupMessage 2, "Entry data from Warehouse...", , , uWarehouseEntry.MRCode & " - Bottle : " & uWarehouseEntry.EntryBottle
    
End Function
Private Function AddInStockFrame(ByVal mrc As Boolean)
Dim rc As Boolean
 Dim i As Integer
 
 lbFormulation(8) = "# Bottles Arrived ( number )"
 bViewWarehouseEntry = False
  
frStockLabel.Visible = False
  
 
 bCopyEntry = mrc
 
 If uMR.Code = "" Then
    If F_MsgBox.DoShow("Please select a valid MR") Then
        
        SetMRInTable
        
    
    Else
        
    End If
    
    Exit Function
 Else
 
 
    For i = txFormulation.LBound To txFormulation.UBound
        txFormulation(i) = ""
         txFormulation(i).Locked = False
        
    Next
    
    lbUmconcetration = uMR.Unit
    lbUUnit = uMR.Unit
    
    
    Call SetLocations
    Call AddUmInCombo
    
    rc = IIf(uMR.PhysicalState = "L", True, False)
    lbFormulation(6).Visible = rc
    txFormulation(6).Visible = rc
    Label6.Visible = rc
    
    If rc = False Then txFormulation(6) = 1
    
    txFormulation(0).SetFocus
    PBContainer.Top = -frInside(1).Top + 480
    If Not (mrc) Then
        PopupMessage 2, "Enter Chemical MR specifics and Save( add in Stock )...", , , uMR.Code
    Else
        ' copy data from table....
    
    End If
    
End If

End Function




Private Sub txFormulation_Click(Index As Integer)
Dim Answer As String
Dim Selected As String
Dim sString As String
Dim bNumber As Boolean

Selected = "Stock Control"
Answer = txFormulation(Index)
sString = lbFormulation(Index)


cmbLocation.Visible = False
bNumber = False
If Index = 4 Or Index = 5 Or Index = 6 Or Index = 7 Or Index = 8 Then bNumber = True
If Index = 2 Then If Answer = "" Then Answer = FormatDataLAT(Now())
If Index = 1 Then GoTo Location
        
        If F_InputBox.DoShow(Selected & " : Enter Value", sString, , , , Answer, , bNumber, Me.Top) Then
    
            txFormulation(Index) = Answer
            
            Select Case Index
            
                Case 4 To 8
                    ' is numeric
                    If IsNumeric(Answer) Then
                         txFormulation(Index) = CDbl(Answer)
                    Else
                        PopupMessage 2, "Please enter a valid Number...", , True
                    End If
                 
                    If IsNumeric(txFormulation(Index)) Then
                        txFormulation(Index) = FormatNumber(txFormulation(Index), 2)
                    End If
        
                
                    If Index = 8 Then
                        If IsNumeric(txFormulation(Index)) Then
                            txFormulation(Index) = FormatNumber(txFormulation(Index), 0)
                        End If
                    End If
            
                Case 1
Location:
                    With cmbLocation
                        .Move txFormulation(Index).Left, txFormulation(Index).Top, txFormulation(Index).Width
                        .ZOrder
                        .Visible = True
                        .SetFocus
                    End With
                    Exit Sub
            
                    
                Case 2, 3
                    ' isdate?
                    If IsDate(Answer) Then
                         txFormulation(Index) = FormatDataLAT(Answer)
                    Else
                        PopupMessage 2, "Please enter a valid Date...", , True
                    End If
            End Select
            
            If Index < txFormulation.UBound - 1 Then txFormulation_Click Index + 1
            
        End If
End Sub

Private Sub txFormulation_LostFocus(Index As Integer)
    Select Case Index
        Case 7
            If IsNumeric(txFormulation(Index)) Then
                txFormulation(Index) = FormatNumber(txFormulation(Index), 2)
            End If
           
    
    End Select
    
    
End Sub

Private Sub txFormulation_Change(Index As Integer)
Dim rc As Boolean
    rc = IIf(Len(txFormulation(Index)) > 0, True, False)
    
    txFormulation(Index).BackColor = IIf(rc, vbWhite, &HE0E0E0)
    
    Select Case Index
        Case 0, 1
           
        Case 1
            cmbLocation = txFormulation(1)
          
        Case 13

            
    End Select
End Sub

Private Sub SetLocations()
cmbLocation.Move txFormulation(1).Left, txFormulation(1).Top, txFormulation(1).Width
cmbLocation.Visible = False
cmbLocation.ZOrder
Call AddLocationInCombo(cmbLocation)


If uMR.Location <> "" Then cmbLocation = uMR.Location
End Sub
Private Sub cmbLocation_LostFocus()
txFormulation(1) = cmbLocation
cmbLocation.Visible = False
End Sub

Private Sub cmbLocation_Change()
txFormulation(1) = cmbLocation

End Sub
Private Sub cmbLocation_Click()
If cmbLocation.ListIndex >= 0 Then
    txFormulation(1) = cmbLocation
Else
    txFormulation(1) = ""
End If
cmbLocation.Visible = False
End Sub
Private Sub cmbLocation_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            txFormulation(1) = cmbLocation
            cmbLocation.Visible = False
        Case 8
            cmbLocation.Visible = False
            If cmbLocation <> "" And F_MsgBox.DoShow("Delete record from Location Database?", cmbLocation) Then
                DeleteLocationFromDatabase (cmbLocation)
            End If
            
    End Select
End Sub


Private Sub AddUmInCombo()

Call SetUmCombo(cbUM, False, False)
If Trim(uMR.STOCK_UNIT) = "" Then
Else

    cbUM = uMR.STOCK_UNIT

End If

End Sub




Private Function SaveEntry()
Dim rc As Boolean


rc = ChecktxFormulation

If rc Then

    If cmbLocation <> "" Then Call AddLocationInDatabase(cmbLocation)
    
        If bViewWarehouseEntry Then uWarehouseEntry.NumberBottle = 1
        rc = SetuWarehouseEntry

    
    If rc Then
    
     If SaveWarehouseEntryInDatabase(uWarehouseEntry, bViewWarehouseEntry) Then
        
        
        PopupMessage 2, uWarehouseEntry.NumberBottle & " Bottles stored in Warehouse...", , , uMR.Code
        
        frStockLabel.Visible = True
      
     End If
    
    End If
End If

End Function

Private Function ChecktxFormulation() As Boolean
Dim rc As Boolean
Dim i As Integer
    rc = True
    For i = txFormulation.LBound To txFormulation.UBound - 1
        If Len(txFormulation(i)) = 0 Then
            rc = False
            MessageInfoTime = 1500
            PopupMessage 2, "Please Enter field : " & lbFormulation(i), , True, "WareHouse"
            txFormulation(i).SetFocus
            rc = False
            Exit For
        End If
    Next
    ChecktxFormulation = rc
End Function

Private Function ViewWarehouseEntry() As Boolean
Dim rc As Boolean
Dim NumBottle As Integer
Dim LastLetter As String
Dim NewLetter   As String
Dim i As Integer
rc = True
On Error GoTo ERR_SET:




WarehouseEntries() = MyWareHouseEntryCleanArray
ReDim WarehouseEntries(0)


 rc = GetDatabaseWareHouseEntry(uWarehouseEntry.ID, 0, True, WarehouseEntries())
 

uWarehouseEntry = WarehouseEntries(0)


With uWarehouseEntry
    .NumberBottle = 1
    
    uMR.Description = IIf(.Description = "", uMR.Description, .Description)
    uMR.FWParameter = IIf(.Description = "", uMR.FWParameter, .FWParameter)
    uMR.Code = IIf(.MRCode = "", uMR.Code, .MRCode)
    uMR.Parameter = IIf(.Parameter = "", uMR.Parameter, .Parameter)
    uMR.Unit = IIf(.Unit = "", uMR.Unit, .Unit)
    If IsNull(.stockUnit) Or .stockUnit = "" Then .stockUnit = "mL"
    lbUM = .stockUnit
    
    
    
    If IsDate(.SupplierEXP) And IsDate(.Open) Then
        .MREXP = CreateMRExp((.SupplierEXP), uMR.ReductionExpDays)
    End If
    
    
    lbUUnitEdit = uMR.Unit
    
    txStock(0) = .Lot
    txStock(1) = .Location
    txStock(2) = FormatDataLAT(.ArrivedTime)
    txStock(3) = FormatDataLAT(.SupplierEXP)
    txStock(4) = .Purity
    txStock(5) = .MRValueConcentration
    txStock(6) = .Density
    
    txStock(7) = .StockQTY
    txStock(8) = .EntryBottle
    txStock(9) = FormatDataLAT(.Open)
    txStock(10) = FormatDataLAT(.Finished)
    txStock(14) = .U
    
    
    txStock(11) = FormatDataLAT(.MREXP)
    
    If .Finished = "" And .MREXP <> "" Then
        If .MREXP < Now() Then
            txStock(11).BackColor = vbColorRed
            txStock(11).ForeColor = vbWhite
        Else
            txStock(11).BackColor = vbWhite
            txStock(11).ForeColor = vbBlack
            
        
        End If
    Else
         txStock(11).BackColor = &HE0E0E0
        txStock(11).ForeColor = vbBlack
    
    End If
    
    txStock(12) = GetStatus(.Status)
    txStock(13) = .Note

    cmbLocation = txStock(1)
    

End With


ERR_END:
    On Error GoTo 0
    ViewWarehouseEntry = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox Err.Description
    Resume Next
    

End Function


Private Function SetuWarehouseEntry() As Boolean
Dim rc As Boolean
Dim NumBottle As Integer
Dim LastLetter As String
Dim NewLetter   As String
Dim i As Integer
rc = True
On Error GoTo ERR_SET:

If bViewWarehouseEntry Then
    NumBottle = 1
Else

    NumBottle = CInt(txFormulation(8))
End If

With uWarehouseEntry
    .ArrivedTime = txFormulation(2)
    
    .Density = txFormulation(6)
    .Description = uMR.Description
    .Finished = ""
    .FWParameter = uMR.FWParameter
    .Location = txFormulation(1)
    .Lot = txFormulation(0)
    .MRCode = uMR.Code
    .MRValueConcentration = txFormulation(5)
    .Note = txFormulation(9)
    .U = txFormulation(10)
    .Operator = MyOperatore.Name
    .Parameter = uMR.Parameter
    .Purity = txFormulation(4)
    .Status = 0
    .StockQTY = txFormulation(7)
    .stockUnit = cbUM
    .SupplierEXP = txFormulation(3)
    .Unit = uMR.Unit
    .NumberBottle = NumBottle
    
    If bViewWarehouseEntry Then
    
    Else
    
          
          ReDim .Bottle(NumBottle - 1)
          ReDim .BarcodeText(NumBottle - 1)
          ReDim .bBarcode(NumBottle - 1)
          
          Call GetLastUMBottleLetter(uWarehouseEntry)
          
          LastLetter = IIf(Len(.LastLetter) = 1, "0" & .LastLetter, .LastLetter)
          
          For i = 0 To NumBottle - 1
        
              NewLetter = GetLetter(LastLetter)
          
              .Bottle(i) = NewLetter
              .BarcodeText(i) = .MRCode & sQRSeparator & .Lot & sQRSeparator & .Bottle(i)
              .bBarcode(i) = False
              
              LastLetter = NewLetter
              
          Next
    
    End If

End With


ERR_END:
    On Error GoTo 0
    SetuWarehouseEntry = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox Err.Description
    Resume Next
    

End Function


Private Function CopyEntry(ByVal Row As Long)

AddInStockFrame (True)

End Function




Private Function ModifyEntry()
Dim rc As Boolean


rc = ChecktxStock

If rc Then

    If txStock(1) <> "" Then Call AddLocationInDatabase(txStock(1))
    
        If bViewWarehouseEntry Then uWarehouseEntry.NumberBottle = 1
        rc = ModifyuWarehouseEntry

    
    If rc Then
    
     If SaveWarehouseEntryInDatabase(uWarehouseEntry, bViewWarehouseEntry) Then
        
        
        PopupMessage 2, uWarehouseEntry.NumberBottle & " Bottles stored in Warehouse...", , , uMR.Code
        
        Call FillChemicalMRTable(uMR.Code)
      
        
        
        PBContainer.Top = -0
         
        
     
     End If
    
    End If
End If

End Function

Private Function ChecktxStock() As Boolean
Dim rc As Boolean
Dim i As Integer
    rc = True
    'For i = txStock.LBound To txStock.UBound - 5
        'If Len(txStock(i)) = 0 Then
            'rc = False
            'MessageInfoTime = 1500
            'PopupMessage 2, "Please Enter field : " & lbStock(i), , True, "WareHouse"
           ' txStock(i).SetFocus
           ' rc = False
            'Exit For
      '  End If
   ' Next
    
    
    If Len(txStock(9)) > 0 Then
            
            txStock(12) = GetStatus(1)
    End If
      If Len(txStock(10)) > 0 Then
            
            txStock(12) = GetStatus(2)
    End If
      
    
    ChecktxStock = rc
End Function
Private Sub txStock_DblClick(Index As Integer)
    Select Case Index
        
        Case 2, 3, 9, 10, 11
            txStock(Index) = FormatDataLAT(Date)
    End Select


End Sub

Private Sub txStock_Click(Index As Integer)
Dim Answer As String
Dim Selected As String
Dim sString As String
Dim bNumber As Boolean

Selected = "Stock Control"
Answer = txStock(Index)
sString = lbStock(Index)


cmbLocation.Visible = False
bNumber = False
If Index = 8 Then
    
    PopupMessage 2, "Cannot Change Bottle Number!"
    Exit Sub
    
End If
If Index = 4 Or Index = 5 Or Index = 6 Or Index = 7 Or Index = 8 Or Index = 14 Then bNumber = True
If Index = 2 Then If Answer = "" Then Answer = FormatDataLAT(Now())
If Index = 1 Then GoTo Location

If Index = 12 Then
    cmbStatus.Move txStock(12).Left, txStock(12).Top, txStock(12).Width
    cmbStatus.Visible = True
    cmbStatus.ZOrder
    Exit Sub
End If
        
        If F_InputBox.DoShow(Selected & " : Enter Value", sString, , , , Answer, , bNumber, Me.Top) Then
    
            txStock(Index) = Answer
            
            Select Case Index
            
                Case 4 To 8
                    ' is numeric
                    If IsNumeric(Answer) Then
                         txStock(Index) = CDbl(Answer)
                    Else
                        PopupMessage 2, "Please enter a valid Number...", , True
                     
                    End If
                 
                    If IsNumeric(txStock(Index)) Then
                        txStock(Index) = FormatNumber(txStock(Index), 2)
                    End If
        
                
                    If Index = 8 Then
                        If IsNumeric(txStock(Index)) Then
                            txStock(Index) = FormatNumber(txStock(Index), 0)
                        End If
                    End If
                Case 14
                    If IsNumeric(Answer) Then
                         txStock(Index) = CDbl(Answer)
                    Else
                        PopupMessage 2, "Please enter a valid Number...", , True
                     
                    End If
            
                Case 1
Location:
                    With cmbLocation
                        .Move txStock(Index).Left, txStock(Index).Top, txStock(Index).Width
                        .ZOrder
                        .Visible = True
                        .SetFocus
                    End With
                    Exit Sub
            
                    
                Case 2, 3, 9, 10, 11
                    ' isdate?
                    If IsDate(Answer) Then
                         txStock(Index) = FormatDataLAT(Answer)
                    Else
                        PopupMessage 2, "Please enter a valid Date...", , True
                        Exit Sub
                    End If
                    
                    If Index = 9 Then
                        '-------------------------------------------------
                        ' MREXP
                        ' Supplier EXP ( Date ) - ReductionExpDays
                        '-------------------------------------------------
                        
                        If IsDate(txStock(3)) Then
                            txStock(11) = CreateMRExp(txStock(3), uMR.ReductionExpDays) ' FormatDataLAT(DateAdd("d", -CInt(uMR.ReductionExpDays), txStock(3)))
                        End If
                     ' (DateAdd("d", GIORNI_AVVISO, Date))
                    End If
                    
                Case 12
                
            
            End Select
        End If
End Sub




Private Sub txStock_LostFocus(Index As Integer)
    Select Case Index
        Case 7
            If IsNumeric(txStock(Index)) Then
                txStock(Index) = FormatNumber(txStock(Index), 2)
            End If
        Case 9

           ' ReductionExpDays
           If IsDate(txStock(3)) And IsDate(txStock(9)) Then
                txStock(11) = CreateMRExp(txStock(3), uMR.ReductionExpDays)  'txStock(11) = FormatDataLAT(DateAdd("d", -CInt(uMR.ReductionExpDays), txStock(3)))
           End If
        ' (DateAdd("d", GIORNI_AVVISO, Date))
    
        
    
    End Select
    
    
End Sub

Private Sub txStock_Change(Index As Integer)
Dim rc As Boolean
    rc = IIf(Len(txStock(Index)) > 0, True, False)
    
    
    Select Case Index
        Case 0, 1
           
        Case 1
            cmbLocation = txStock(1)
        
        Case 8
            Exit Sub
        Case 11
        lbReduction = "* " & uMR.ReductionExpDays & " Reduction Days"
            lbReduction.Visible = rc
            

            
    End Select
    
    txStock(Index).BackColor = IIf(rc, vbWhite, &HE0E0E0)
    
End Sub


Private Sub cmbStatus_LostFocus()
txStock(12) = cmbStatus
cmbLocation.Visible = False
End Sub

Private Sub cmbStatus_Change()
txStock(12) = cmbStatus
End Sub
Private Sub cmbStatus_Click()
If cmbStatus.ListIndex >= 0 Then
    txStock(12) = cmbStatus
Else
    txStock(12) = ""
End If
cmbStatus.Visible = False
End Sub
Private Sub cmbStatus_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            txStock(12) = cmbStatus
            cmbStatus.Visible = False
        Case 8
    End Select
End Sub




Private Sub frPrinQRCode_Click()


    If F_MsgBox.DoShow("Print QRCode?") Then
            
            Call DoPrintQRCodeLabel(uWarehouseEntry.MRCode, uWarehouseEntry.Lot, uWarehouseEntry.EntryBottle, uWarehouseEntry.MREXP)
    
    End If

End Sub


Private Sub frStockLabel_Click()
Dim rc As Boolean
    If F_MsgBox.DoShow("Print Stock Label?", "Print " & uWarehouseEntry.NumberBottle & " Label/s") Then
            
            rc = DoPrintQRCodeStockLabel(uWarehouseEntry)
            If rc Then
            PopupMessage 2, "All Labels were printed and stored in " & vbCrLf & USER_LABEL_PATH, , , "LABEL PRINTED"
            End If
          
    End If
End Sub


Private Function ModifyuWarehouseEntry() As Boolean
Dim rc As Boolean
Dim NumBottle As Integer
Dim LastLetter As String
Dim NewLetter   As String
Dim i As Integer
rc = True
On Error GoTo ERR_SET:


    NumBottle = 1


With uWarehouseEntry
    
    .FWParameter = uMR.FWParameter
    .MRCode = uMR.Code
    .Operator = MyOperatore.Name
    .Parameter = uMR.Parameter
    .Unit = uMR.Unit
    .NumberBottle = NumBottle

    .Lot = txStock(0)
    .Location = txStock(1)
    .ArrivedTime = txStock(2)
    .SupplierEXP = txStock(3)
    .Purity = txStock(4)
    .MRValueConcentration = txStock(5)
    .Density = txStock(6)
    .U = txStock(14)
    .StockQTY = txStock(7)
    .stockUnit = lbUM
    .EntryBottle = txStock(8)
    .Description = uMR.Description
    
    .Open = txStock(9)
    .Finished = txStock(10)
    .MREXP = txStock(11)
    
    .Status = IndexStatus(txStock(12))
    .Note = txStock(13)

End With


ERR_END:
    On Error GoTo 0
    ModifyuWarehouseEntry = rc
    Exit Function
ERR_SET:
    rc = False
    MsgBox Err.Description
    Resume Next
    

End Function

