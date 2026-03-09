VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form F_MAIN 
   BackColor       =   &H004D3B37&
   Caption         =   "Chemical MR"
   ClientHeight    =   11985
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
   ForeColor       =   &H00404040&
   Icon            =   "F_MAIN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12500
   ScaleMode       =   0  'User
   ScaleWidth      =   19320
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicMain 
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
      Height          =   9060
      Index           =   1
      Left            =   480
      ScaleHeight     =   9060
      ScaleWidth      =   19200
      TabIndex        =   36
      Top             =   1080
      Visible         =   0   'False
      Width           =   19200
      Begin VB.ComboBox cmbLine 
         BackColor       =   &H00F0F0F0&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   7440
         Visible         =   0   'False
         Width           =   3495
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
         Height          =   8655
         Index           =   0
         Left            =   840
         TabIndex        =   46
         Top             =   360
         Width           =   17655
         Begin VB.Frame frCommandInside 
            BackColor       =   &H00008080&
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
            TabIndex        =   99
            Top             =   8160
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Add Manual Preparation"
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
               Height          =   240
               Index           =   11
               Left            =   120
               TabIndex        =   100
               Top             =   90
               Width           =   2880
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
            Index           =   10
            Left            =   6240
            TabIndex        =   95
            Top             =   7560
            Visible         =   0   'False
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Open Preparation"
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
               TabIndex        =   96
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
            Index           =   9
            Left            =   14640
            TabIndex        =   93
            Top             =   7560
            Visible         =   0   'False
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Delete Preparation"
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
               TabIndex        =   94
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
            Index           =   7
            Left            =   3120
            TabIndex        =   91
            Top             =   7560
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "MR Code Table"
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
               TabIndex        =   92
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
            Index           =   8
            Left            =   0
            TabIndex        =   89
            Top             =   7560
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Add Preparation"
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
               Height          =   240
               Index           =   8
               Left            =   885
               TabIndex        =   90
               Top             =   90
               Width           =   1440
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
            TabIndex        =   49
            Top             =   0
            Width           =   17655
            Begin VB.Line Line1 
               BorderColor     =   &H00B0B0B0&
               X1              =   120
               X2              =   17640
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label lbInside 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "STD Preparation (Hanna Code)"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00644603&
               Height          =   285
               Index           =   0
               Left            =   120
               MouseIcon       =   "F_MAIN.frx":6852
               MousePointer    =   99  'Custom
               TabIndex        =   50
               Top             =   120
               Width           =   3540
            End
         End
         Begin VB.Frame frPreparationGrid 
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
            Height          =   1455
            Left            =   4440
            TabIndex        =   47
            Top             =   2520
            Visible         =   0   'False
            Width           =   8535
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "STD Preparation"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   435
               Index           =   7
               Left            =   3045
               TabIndex        =   64
               Top             =   360
               Width           =   2415
            End
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
               Left            =   30
               TabIndex        =   48
               Top             =   720
               Width           =   8415
            End
         End
         Begin FlexCell.Grid Grid2 
            Height          =   6855
            Left            =   0
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   600
            Width           =   17655
            _ExtentX        =   31141
            _ExtentY        =   12091
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
         Begin VB.Line Line11 
            BorderColor     =   &H00B0B0B0&
            X1              =   0
            X2              =   17280
            Y1              =   0
            Y2              =   0
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   285
         Left            =   15600
         MouseIcon       =   "F_MAIN.frx":6B5C
         MousePointer    =   99  'Custom
         TabIndex        =   37
         Top             =   960
         Width           =   60
      End
      Begin VB.Label lbWaitPreparation 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "Wait : Loading Data..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   8640
         Visible         =   0   'False
         Width           =   18975
      End
   End
   Begin VB.Frame frAvvio 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   975
      Left            =   3240
      TabIndex        =   74
      Top             =   8160
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Preferences...."
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
         Index           =   10
         Left            =   90
         TabIndex        =   76
         Top             =   480
         Width           =   19035
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wait"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   9
         Left            =   0
         TabIndex        =   75
         Top             =   120
         Width           =   19215
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   0
      Top             =   6360
   End
   Begin VB.Timer TimeriNTRO 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   0
      Top             =   5640
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   7320
   End
   Begin VB.Frame frSelezione 
      BackColor       =   &H00964901&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   100
      Left            =   0
      TabIndex        =   77
      Top             =   960
      Visible         =   0   'False
      Width           =   19215
   End
   Begin VB.Frame frCloseFrame 
      BackColor       =   &H004D3B37&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   8640
      TabIndex        =   58
      Top             =   10920
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Image DefaultMenu 
         DragIcon        =   "F_MAIN.frx":6E66
         Height          =   480
         Index           =   1
         Left            =   720
         MouseIcon       =   "F_MAIN.frx":A248
         MousePointer    =   99  'Custom
         Picture         =   "F_MAIN.frx":A552
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Close Table"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   9
         Left            =   480
         MouseIcon       =   "F_MAIN.frx":D934
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   795
         Width           =   975
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox PBTitle 
      BackColor       =   &H00473733&
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
      Begin VB.Frame frSearchRecipe 
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   6120
         TabIndex        =   80
         Top             =   120
         Width           =   8175
         Begin VB.Frame frCommandInside 
            BackColor       =   &H00644603&
            BorderStyle     =   0  'None
            Caption         =   "ù"
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
            Left            =   4080
            TabIndex        =   82
            Top             =   120
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "MR Search"
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
               TabIndex        =   83
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.TextBox txSearch 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00808080&
            Height          =   400
            Left            =   360
            TabIndex        =   81
            Text            =   "Search"
            Top             =   160
            Width           =   3495
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00B0B0B0&
            Height          =   735
            Left            =   0
            Top             =   0
            Width           =   8175
         End
         Begin VB.Image Image1 
            Height          =   360
            Left            =   7440
            Picture         =   "F_MAIN.frx":DC3E
            Top             =   160
            Width           =   360
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00473733&
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
         MouseIcon       =   "F_MAIN.frx":F6B0
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   9
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   720
            MouseIcon       =   "F_MAIN.frx":F9BA
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":FCC4
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STD Preparation"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   1
            Left            =   270
            MouseIcon       =   "F_MAIN.frx":126B6
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   645
            Width           =   1350
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00473733&
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
         MouseIcon       =   "F_MAIN.frx":129C0
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   7
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MR Stock Control"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   300
            MouseIcon       =   "F_MAIN.frx":12CCA
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   645
            Width           =   1410
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "F_MAIN.frx":12FD4
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":132DE
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open Batch Table"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   585
         Left            =   14730
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   4260
      End
   End
   Begin VB.Frame frDatabaseHistory 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   9735
      Left            =   -1920
      TabIndex        =   66
      Top             =   -2280
      Visible         =   0   'False
      Width           =   19095
      Begin VB.Frame frCommandInside 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Caption         =   "ù"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   4
         Left            =   9960
         TabIndex        =   69
         Top             =   3840
         Width           =   6000
         Begin VB.Image Image 
            Height          =   480
            Index           =   4
            Left            =   360
            MouseIcon       =   "F_MAIN.frx":15CD0
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":15FDA
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preparation STD"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   4
            Left            =   1620
            TabIndex        =   70
            Top             =   240
            Width           =   2700
         End
      End
      Begin VB.Frame frCommandInside 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Caption         =   "ù"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   3
         Left            =   3120
         TabIndex        =   67
         Top             =   3840
         Width           =   6000
         Begin VB.Image Image 
            Height          =   480
            Index           =   3
            Left            =   360
            MouseIcon       =   "F_MAIN.frx":189CC
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":18CD6
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chemical MR"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   420
            Index           =   3
            Left            =   1815
            TabIndex        =   68
            Top             =   240
            Width           =   2370
         End
      End
      Begin VB.Label llbExit 
         BackStyle       =   0  'Transparent
         Height          =   1455
         Left            =   7680
         TabIndex        =   73
         Top             =   8280
         Width           =   3855
      End
      Begin VB.Label lbExit 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Database Manager"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   8625
         MouseIcon       =   "F_MAIN.frx":1C0B8
         MousePointer    =   99  'Custom
         TabIndex        =   72
         Top             =   9300
         Width           =   1920
      End
      Begin VB.Image DefaultExit 
         DragIcon        =   "F_MAIN.frx":1C3C2
         Height          =   480
         Left            =   9360
         MouseIcon       =   "F_MAIN.frx":1F7A4
         MousePointer    =   99  'Custom
         Picture         =   "F_MAIN.frx":1FAAE
         Top             =   8760
         Width           =   480
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Manager"
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
         Height          =   870
         Left            =   5970
         TabIndex        =   71
         Top             =   2040
         Width           =   6960
      End
   End
   Begin VB.PictureBox PicIntro 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Height          =   9690
      Left            =   600
      MouseIcon       =   "F_MAIN.frx":22E90
      MousePointer    =   99  'Custom
      ScaleHeight     =   9690
      ScaleWidth      =   19200
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   19200
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Instruments"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644603&
         Height          =   540
         Index           =   2
         Left            =   9240
         TabIndex        =   21
         Top             =   5280
         Width           =   4050
      End
      Begin VB.Label lbProgram 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644603&
         Height          =   270
         Left            =   16440
         TabIndex        =   16
         Top             =   5400
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chemical MR"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00644603&
         Height          =   1620
         Index           =   0
         Left            =   5280
         TabIndex        =   12
         Top             =   3600
         Width           =   12660
      End
   End
   Begin VB.PictureBox PicDatabase 
      BackColor       =   &H00C0C0C0&
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
      Height          =   9615
      Left            =   1560
      ScaleHeight     =   9615
      ScaleWidth      =   19215
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Image Image6 
         Height          =   480
         Index           =   8
         Left            =   7680
         Picture         =   "F_MAIN.frx":2319A
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Go to Settings > User Account"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004D3B37&
         Height          =   1125
         Left            =   0
         MouseIcon       =   "F_MAIN.frx":2657C
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   4320
         Visible         =   0   'False
         Width           =   19140
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   4
         Left            =   6480
         Picture         =   "F_MAIN.frx":26886
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No User Account in your Database : Please enter at least 1 User to set Privilege"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   0
         MouseIcon       =   "F_MAIN.frx":29C68
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   6720
         Visible         =   0   'False
         Width           =   19185
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Operator  / User Account"
         ForeColor       =   &H00000080&
         Height          =   450
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   5520
         Visible         =   0   'False
         Width           =   19170
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Your Hanna Code Database is Empty : Please goto Settings > Database and Import Hanna Code form Excel"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   3
         Left            =   0
         MouseIcon       =   "F_MAIN.frx":29F72
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   3000
         Width           =   19185
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Go to Settings > Database"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004D3B37&
         Height          =   885
         Index           =   4
         Left            =   0
         MouseIcon       =   "F_MAIN.frx":2A27C
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   4080
         Width           =   19125
      End
      Begin VB.Image Im 
         Height          =   480
         Index           =   7
         Left            =   10920
         Picture         =   "F_MAIN.frx":2A586
         Top             =   3960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Code"
         ForeColor       =   &H00000080&
         Height          =   450
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Top             =   1680
         Width           =   19170
      End
   End
   Begin VB.PictureBox PicMain 
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
      Height          =   10020
      Index           =   2
      Left            =   360
      ScaleHeight     =   10020
      ScaleWidth      =   19200
      TabIndex        =   38
      Top             =   960
      Visible         =   0   'False
      Width           =   19200
      Begin VB.ComboBox cmbLineQC 
         BackColor       =   &H00F0F0F0&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   8760
         Visible         =   0   'False
         Width           =   3495
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
         Height          =   8055
         Index           =   1
         Left            =   840
         TabIndex        =   52
         Top             =   480
         Width           =   17655
         Begin VB.Frame Frame2 
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
            Height          =   1335
            Left            =   4800
            TabIndex        =   55
            Top             =   3000
            Visible         =   0   'False
            Width           =   8535
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Preparation QC"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   435
               Index           =   6
               Left            =   0
               TabIndex        =   63
               Top             =   360
               Width           =   8535
            End
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
               Left            =   0
               TabIndex        =   56
               Top             =   720
               Width           =   8535
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
            TabIndex        =   53
            Top             =   0
            Width           =   17655
            Begin VB.Label lbInside 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "STD Preparation"
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
               TabIndex        =   54
               Top             =   120
               Width           =   1800
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   17640
               Y1              =   480
               Y2              =   480
            End
         End
         Begin FlexCell.Grid Grid3 
            Height          =   7215
            Left            =   0
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   600
            Width           =   17655
            _ExtentX        =   31141
            _ExtentY        =   12726
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
         Begin VB.Line Line14 
            BorderColor     =   &H00B0B0B0&
            X1              =   0
            X2              =   17640
            Y1              =   8040
            Y2              =   8040
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   285
         Left            =   15600
         MouseIcon       =   "F_MAIN.frx":2D968
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   480
         Width           =   60
      End
   End
   Begin VB.PictureBox PicMain 
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
      Height          =   10020
      Index           =   0
      Left            =   360
      ScaleHeight     =   10020
      ScaleWidth      =   19200
      TabIndex        =   13
      Top             =   1080
      Width           =   19200
      Begin VB.Frame frStockControl 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   8775
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Visible         =   0   'False
         Width           =   19095
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
            Left            =   15360
            TabIndex        =   97
            Top             =   8160
            Visible         =   0   'False
            Width           =   3015
            Begin VB.Image Image 
               Height          =   480
               Index           =   0
               Left            =   120
               MousePointer    =   99  'Custom
               OLEDropMode     =   1  'Manual
               Picture         =   "F_MAIN.frx":2DC72
               Top             =   0
               Width           =   480
            End
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
               TabIndex        =   98
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
            Index           =   6
            Left            =   3840
            TabIndex        =   87
            Top             =   8160
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "MR Stock QTY"
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
         Begin VB.Frame frCommandInside 
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
            Index           =   5
            Left            =   10080
            TabIndex        =   85
            Top             =   8160
            Visible         =   0   'False
            Width           =   3015
            Begin VB.Label lbCommandInside 
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
               Index           =   5
               Left            =   0
               TabIndex        =   86
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.ComboBox cmbLineRfP 
            BackColor       =   &H00F0F0F0&
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   14040
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   7560
            Visible         =   0   'False
            Width           =   3375
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
            Height          =   7815
            Index           =   2
            Left            =   720
            TabIndex        =   40
            Top             =   360
            Width           =   17895
            Begin VB.Frame frReceiptGrid 
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
               Height          =   1335
               Left            =   5400
               TabIndex        =   43
               Top             =   3000
               Visible         =   0   'False
               Width           =   8535
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "MR Stock Control"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   18
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   435
                  Index           =   8
                  Left            =   60
                  TabIndex        =   65
                  Top             =   360
                  Width           =   8385
               End
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
                  Left            =   90
                  TabIndex        =   44
                  Top             =   720
                  Width           =   8415
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
               Height          =   495
               Index           =   2
               Left            =   0
               TabIndex        =   41
               Top             =   0
               Width           =   18255
               Begin VB.Label lbStockHistory 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "+ Finished MR"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00964901&
                  Height          =   375
                  Left            =   14880
                  TabIndex        =   84
                  Top             =   120
                  Width           =   2655
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "MR Stock Bottles"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00644603&
                  Height          =   285
                  Index           =   2
                  Left            =   0
                  MouseIcon       =   "F_MAIN.frx":31054
                  MousePointer    =   99  'Custom
                  TabIndex        =   42
                  Top             =   120
                  Width           =   1785
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   0
                  X2              =   17640
                  Y1              =   480
                  Y2              =   480
               End
            End
            Begin FlexCell.Grid Grid1 
               Height          =   7095
               Left            =   0
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   600
               Width           =   17895
               _ExtentX        =   31565
               _ExtentY        =   12515
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
            TabIndex        =   32
            Top             =   8160
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
               Index           =   0
               Left            =   0
               TabIndex        =   33
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
            Index           =   1
            Left            =   720
            TabIndex        =   30
            Top             =   8160
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Add MR in Stock"
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
               Height          =   240
               Index           =   1
               Left            =   720
               TabIndex        =   31
               Top             =   120
               Width           =   1410
            End
         End
      End
      Begin VB.Frame frPrivilege 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1455
         Left            =   5160
         TabIndex        =   34
         Top             =   4320
         Visible         =   0   'False
         Width           =   9135
         Begin VB.Image DisableImage 
            Height          =   480
            Left            =   4080
            Picture         =   "F_MAIN.frx":3135E
            Top             =   240
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Line Leader/Production Manager or Administrator Login Required"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   270
            Left            =   1320
            TabIndex        =   35
            Top             =   840
            Width           =   6045
         End
      End
      Begin VB.Label lbWait 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "Wait : Loading Data..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   9600
         Visible         =   0   'False
         Width           =   18975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   285
         Left            =   15600
         MouseIcon       =   "F_MAIN.frx":34740
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   480
         Width           =   60
      End
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
      Height          =   1455
      Index           =   1
      Left            =   8280
      MouseIcon       =   "F_MAIN.frx":34A4A
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   10560
      Width           =   2655
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
      Left            =   17640
      MouseIcon       =   "F_MAIN.frx":34D54
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   10800
      Width           =   1455
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
      Height          =   1815
      Index           =   2
      Left            =   12960
      MouseIcon       =   "F_MAIN.frx":3505E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   10200
      Width           =   2775
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
      Height          =   1455
      Index           =   3
      Left            =   3840
      MouseIcon       =   "F_MAIN.frx":35368
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   10560
      Width           =   1815
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
      Height          =   1455
      Index           =   4
      Left            =   0
      MouseIcon       =   "F_MAIN.frx":35672
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   10680
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quit Program"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   8
      Left            =   17880
      MouseIcon       =   "F_MAIN.frx":3597C
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   11715
      Width           =   1110
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   7
      Left            =   14160
      MouseIcon       =   "F_MAIN.frx":35C86
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   11715
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set Operator"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   6
      Left            =   4260
      MouseIcon       =   "F_MAIN.frx":35F90
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   11715
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   240
      MouseIcon       =   "F_MAIN.frx":3629A
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   11715
      Width           =   645
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   14160
      MouseIcon       =   "F_MAIN.frx":365A4
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":368AE
      Top             =   11160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   4560
      MouseIcon       =   "F_MAIN.frx":39C90
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":39F9A
      Top             =   11160
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   360
      MouseIcon       =   "F_MAIN.frx":3D37C
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":3D686
      Top             =   11160
      Width           =   480
   End
   Begin VB.Line BottomLine 
      BorderColor     =   &H00C0C0C0&
      Visible         =   0   'False
      X1              =   0
      X2              =   19320
      Y1              =   11514.39
      Y2              =   11514.39
   End
   Begin VB.Label lbMenuHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esci"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Index           =   1
      Visible         =   0   'False
      X1              =   4830
      X2              =   4830
      Y1              =   0
      Y2              =   12390.49
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   9660
      X2              =   9660
      Y1              =   375.469
      Y2              =   12765.96
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_MAIN.frx":40A68
      Height          =   480
      Index           =   0
      Left            =   18120
      MouseIcon       =   "F_MAIN.frx":43E4A
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":44154
      Top             =   11160
      Width           =   480
   End
End
Attribute VB_Name = "F_MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private IndexProcedura As Integer
    
Private lRowCode As Long
Private lRow As Long

Private TimerCount As Integer

Private SelectedCode As String
Private SelectedCodeID As Long
Private bManualPreparation As Boolean

Private m_rc As Boolean

Private IndexDashCommInside As Integer
Private TX_INTRO As String

Private STDPreparationID  As Long
Private STDPreparationFileName  As String

Private UserBarcode As Barcode
Private uMR As MRType
Private uEntry As WareHouseEntry

Private FileName As String






Private Sub Command1_Click()

    frmPreparation_Static.Left = Me.Left
    frmPreparation_Static.Top = Me.Top
    DoEvents
    DoEvents
    
    If frmPreparation_Static.DoShow(SelectedCode, FileName, SelectedCodeID) Then
        
        DoEvents
        DoEvents
       ' Call GlobalSearch
    
    End If
    
    DoEvents
    DoEvents
    '

End Sub







Private Sub DefaultMenu_Click(Index As Integer)
DefaultMenuLabel_Click Index
End Sub

Private Sub Form_Activate()
CloseSettingDataFile
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


Private Sub SaveSizes()
Dim i As Integer
Dim ctl As Control
' Save the controls' positions and sizes.

m_ControlGridFontSizeOld = 1
 m_ControlGridColWidthOld = 1
 m_ControlGridRowHeightOld = 1
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
        ElseIf TypeOf ctl Is Image Then
            ctl.Left = (x_scale * .Left) + IIf(x_scale = 1, 0, (x_scale - 1) * 200)
            ctl.Top = y_scale * .Top
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
Private Sub DefaultMenuLabel_Click(Index As Integer)
DefaultMenu(2).Visible = True
Label2(7).Visible = True


Select Case Index
    Case 0
            If F_MsgBox.DoShow("Quit Chemical MR?", "Exit") Then

                CloseSettingDataFile
                
                If bAddNewDatabaseRelease Then AddReleaseNumber
                   
                If F_MsgBox.DoShow("Turn off PC?", ProjectName) Then
                    ShutDownNT True
                End If
                
                
                Unload Me
                Exit Sub
            End If
    Case 1
        ' close frame MAterial requisisition
        frCommandInside_Click 16
        
            
    Case 2
    
        Call SetDatabaseHistory

    Case 3
       frmLogin.DoShow
       Label2(6) = IIf(MyOperatore.Name <> "", MyOperatore.Name, Label2(6))
    Case 4
        F_SETTING.Left = Me.Left
        F_SETTING.Top = Me.Top
        F_SETTING.WindowState = Me.WindowState
        If F_SETTING.DoShow(, , , DefaultMenu(4)) Then
            FormIntro
            Call FillStockDatabaseGrid(False)
        End If
    Case 5
        
        
    Case 6
        
    Case 7
        
    Case 8
      

    Case 9

        Exit Sub
        
End Select
'frmSTDToleranceInfo.Visible = False
End Sub

Private Sub DisableImage_Click()
Dim rc As Boolean
'PopupMessage 2, "Warning : Line Leader Only can Operate...", , True
rc = CheckPrivilege(1)
frStockControl.ZOrder
frStockControl.Visible = rc
frPrivilege.Visible = Not (rc)
frStockControl.ZOrder

End Sub

Private Sub Frame3_Click(Index As Integer)
    Select Case Index
        Case 0
            ' preparation // refresh grid
            Call GetScheduledSTDInGrid
            
    
    End Select
End Sub

Private Sub Frame4_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub frCloseFrame_Click()
DefaultMenuLabel_Click 1
End Sub


Private Sub frPrivilege_Click()
DisableImage_Click
End Sub


Private Sub Form_Initialize()
lbProgram = "Release " & App.Major & "." & App.Minor & "." & App.Revision

Call StartProcedure

SaveSizes

End Sub

Private Sub Form_Load()
Dim rc As Boolean
    PicDatabase.Top = PBTitle.Height
    PicDatabase.Left = 0
    
    
    
    If bFullScreen Then
        Me.WindowState = 2
    Else
        Me.WindowState = 0
    End If
    

    
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
        PicMenu(i).BackColor = &H473733
    End If
Next
End Sub





Private Sub Form_Resize()

On Error Resume Next

SetPicForm



frSearchRecipe.Left = PicMain(1).Width / 2 - frSearchRecipe.Width / 2


'PBContainerViewport.Move 0, PBTitle.Top + PBTitle.Height, Me.ScaleWidth, BottomLine.y1 - (PBTitle.Top + PBTitle.Height)

ResizeControls


MainWindowState = Me.WindowState

SetColumnWidth

End Sub



Private Sub Form_Unload(Cancel As Integer)

SaveSetting App.Title, "Opzioni", "Full Screen Mode", IIf(Me.WindowState = 0, False, True)

CloseSettingDataFile


Set F_MAIN = Nothing

End

End Sub



Private Sub frCommandInside_Click(Index As Integer)
    Select Case Index
        Case 0
            Call DeleteWharehouseEntry(uEntry.ID, Grid1)
        
        Case 1
            ' add MR in Stock
            Call AddMRInStock
            Exit Sub
            
        Case 2
        
            SearcMRInTable (False)
        Case 3
               Call SetChemicalMRDatabase
               
        Case 4
            
            ' database history : preparation
            Call SetPreparationDatabase
            
            Call ResetGrid2
        Case 5

            Call PrintQRCode
        Case 6
            ' Tabella MR stock qty
             FormChemicalMR.DoShow
             
        Case 7

            If FormChemicalMR.DoShow(, , SelectedCode, False) Then
                If SelectedCode <> "" Then Call AddPreparation(True)
                Exit Sub
            End If
        Case 8
            Call AddPreparation
            Exit Sub
        Case 9
            ' delete preparation
            DeletePreparation
        Case 10
            ' open preparation
            Call OpenPreparationButton
        Case 11
            Call AddPreparation(True, True)
            Exit Sub
    End Select
    
    
    Call SetFrameEmptyView
    
End Sub


Private Sub PrintQRCode()
With UserBarcode
    If lRow > 0 And .Code <> "" Then
        'If bStampaOk Then
           'UploadDownloadMessageCounter = 0
            'PopupMessage 2, "Printing code Label... ", , , .Code & " Bottle : " & .Bottle
           ' Call DoPrintLabel(.Code, .Lot, .Bottle, .Date)
          
       ' End If
   ' Else
       ' MessageInfoTime = 2000
       ' PopupMessage 2, "Please select a valid MR code from table", , , Printer.DeviceName
    'End If
    
    Dim rc As Boolean
    If F_MsgBox.DoShow("Print Stock Label?" & vbCrLf & "Bottle = " & uEntry.Bottle(0), "Print code " & uEntry.MRCode) Then
            uEntry.NumberBottle = 1
            rc = DoPrintQRCodeStockLabel(uEntry)
            If rc Then
            PopupMessage 2, "Label printed and stored in " & vbCrLf & USER_LABEL_PATH, , , "Print code " & uEntry.MRCode
            End If
          
    End If
    
    End If
    
End With
End Sub


Private Function OpenScheduledSTDButton()

    If MyOperatore.Name = "" Then
    
        If frmLogin.DoShow Then
               
        Else
            Exit Function
        End If
    
    End If

    

        USER_PATH = USER_TEMP_PATH

    Call GetScheduledSTDInGrid

End Function

Private Sub Image_Click(Index As Integer)
Select Case Index
    Case 0
        frExcel_Click
End Select

End Sub

Private Sub lbExcel_Click()
frExcel_Click
End Sub

Private Sub frExcel_Click()

    Grid1.ExportToExcel USER_DESKTOP & "\MR_Stock_Bottles.xls", True, True

DoEvents
MessageInfoTime = 2500
PopupMessage 2, "File correcly created on Desktop", , , "MR_Stock_Bottles.xls"
End Sub

Private Function SetFrameEmptyView()
    
    FileName = ""
    SelectedCode = ""

    frReceiptGrid.Visible = IIf(Grid1.Rows > 1, False, True)
    frPreparationGrid.Visible = IIf(Grid2.Rows > 1, False, True)
    frExcel.Visible = IIf(Grid1.Rows > 1, True, False)

End Function
Private Sub frCommandInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
IndexDashCommInside = Index
Dim i As Integer
    For i = 0 To frCommandInside.UBound
        If i = Index Then
            frCommandInside(i).BackColor = &H846623
            lbCommandInside(i).ForeColor = vbWhite
            If i = 1 Or i = 14 Or i = 8 Or i = 10 Then
                frCommandInside(i).BackColor = &H20A020
            End If
            If i = 5 Then
                frCommandInside(i).BackColor = &H40C0&
            End If
            
            If i = 11 Then
                frCommandInside(i).BackColor = vbColorManualPreparation
            End If
            
            
        Else
            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 1 Or i = 14 Or i = 8 Or i = 10 Then
                frCommandInside(i).BackColor = &H8000&
            End If
            
            If i = 5 Then
                frCommandInside(i).BackColor = &H80&
            End If
            
            If i = 11 Then
                frCommandInside(i).BackColor = vbColorManualPreparation
            End If
        End If
    
    Next
 
 
End Sub



Private Sub frReceiptGrid_Click()
frCommandInside_Click 1

End Sub

Private Sub Grid1_BeforeUserSort(ByVal Col As Long)
lRow = 0
End Sub

Private Sub Grid1_DblClick()

    If lRow = 0 Then Exit Sub

    
  
    'If FormChemicalMR.DoShow(uEntry.MRCode) Then
    
        OpenWharehouseStoch (uEntry.MRCode)
    
    'End If
        
End Sub

Private Sub Grid1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

Dim rc As Boolean
lRow = FirstRow

UserBarcode = UserBarcodeClean
uMR = MRTypeClean
uEntry = MyWareHouseEntryClean

frCommandInside(0).Visible = False
frCommandInside(5).Visible = False

If lRow > 0 Then
    With UserBarcode
        .Code = Grid1.Cell(lRow, 1).Text
        .Bottle = Grid1.Cell(lRow, 4).Text
        .Lot = Grid1.Cell(lRow, 6).Text
        .Date = Grid1.Cell(lRow, 13).Text
    End With
    
    With uEntry
        .ID = Grid1.Cell(lRow, 20).Text
        .MRCode = Grid1.Cell(lRow, 1).Text
        ReDim .Bottle(0)
        .Bottle(0) = Grid1.Cell(lRow, 4).Text
        .Lot = Grid1.Cell(lRow, 6).Text
        .MRValueConcentration = Grid1.Cell(lRow, 8).Text
        .ArrivedTime = Grid1.Cell(lRow, 13).Text
        .SupplierEXP = Grid1.Cell(lRow, 16).Text
        .U = Grid1.Cell(lRow, 21).Text
    End With
    
    
    
    frCommandInside(0).Visible = True
    frCommandInside(5).Visible = bStampaOk
    

End If

   

End Sub

Private Sub Grid2_DblClick()
frCommandInside_Click 10
End Sub


Private Function OpenPreparationButton()

   ' If MyOperatore.Name = "" Then
    
   '     If frmLogin.DoShow Then
                'Label2(6) = IIf(MyOperatore.Name <> "", MyOperatore.Name, Label2(6))
   '     Else
   '         Exit Function
    '    End If
    
   ' End If
   
    
    If SelectedCode <> "" Then
        USER_PATH = USER_PREPARATION_PATH
        
        If bManualPreparation Then
        
            frmPreparation_Manual.Left = Me.Left
            frmPreparation_Manual.Top = Me.Top
            frmPreparation_Manual.WindowState = Me.WindowState
            frmPreparation_Manual.DoShow SelectedCode, FileName, SelectedCodeID
        
        Else
        
            frmPreparation_Static.Left = Me.Left
            frmPreparation_Static.Top = Me.Top
            frmPreparation_Static.WindowState = Me.WindowState
            frmPreparation_Static.DoShow SelectedCode, FileName, SelectedCodeID
        
        End If

    End If

Call GetPreparationFromDatabase(Grid2, False)
frPreparationGrid.Visible = IIf(Grid2.Rows > 1, False, True)
frCommandInside(9).Visible = False

End Function

Private Sub Grid2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

frCommandInside(9).Visible = False
FileName = ""
SelectedCode = ""
SelectedCodeID = 0
bManualPreparation = False
frCommandInside(10).Visible = False
If FirstRow > 0 Then

   SelectedCode = Grid2.Cell(FirstRow, 1).Text
   FileName = Grid2.Cell(FirstRow, 17).Text
   SelectedCodeID = Grid2.Cell(FirstRow, 16).Text
   
   If IsNull(Grid2.Cell(FirstRow, 20).Text) Or Grid2.Cell(FirstRow, 20).Text = "" Then
   Else
    bManualPreparation = True
   End If
 

    frCommandInside(10).Visible = True
    frCommandInside(9).Visible = True

End If


End Sub


Private Function DeletePreparation()
Dim rc As Boolean
If SelectedCode <> "" And SelectedCodeID > 0 Then
    
    If F_MsgBox.DoShow("Delete selected Preparation?", SelectedCode) Then
        
        rc = DeleteSelectedPreparation(SelectedCodeID, FileName)
    
        If rc Then
            PopupMessage 2, "Preparation Deleted....", , , SelectedCode
        End If
        
        
    End If

End If

Call ResetGrid2

End Function
Private Sub Image1_Click()
SearcMRInTable (True)
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Image3_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub Label2_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub
Private Sub lbCommandInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
frCommandInside_MouseMove Index, Button, Shift, x, y
End Sub


Private Sub lbInside_Click(Index As Integer)
    Select Case Index
        Case 0
            ' Preparation
            Call ResetGrid2
            Call SetFrameEmptyView
            PopupMessage 2, "Refresh preparation table...", , , "STD Preparation"
            
        Case 2
        
            ' MR Stock Control : Grid1 refresh
            Call FillStockDatabaseGrid(False)
            PopupMessage 2, "Refresh Wharehouse table...", , , "Stock Bottles"
    
    End Select
End Sub


Private Function FillStockDatabaseGrid(ByVal rc As Boolean)

Call GetStockFromDatabase(Grid1, rc)
SetFrameEmptyView

End Function

Private Sub lbStockHistory_Click()
If InStr(lbStockHistory, "Finished") Then

     Call FillStockDatabaseGrid(True)

    lbStockHistory = "In Stock Only"
Else

     Call FillStockDatabaseGrid(False)
     
    lbStockHistory = "+ Finished MR"

End If

 frReceiptGrid.Visible = IIf(Grid1.Rows > 1, False, True)
End Sub

Private Sub llbExit_Click()
SetVisibleDatabaseFrame False

End Sub

Private Sub PBTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Form_MouseDown Button, Shift, x, y
End Sub

Private Sub PBTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Form_MouseMove Button, Shift, x, y
End Sub

Private Sub PBTitle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'FrmMove = False
End Sub

Private Sub PicIntro_Click()
TimeriNTRO.Enabled = False
FormIntro
End Sub

Private Sub PicMain_Click(Index As Integer)

    
    
   
    Select Case Index
        Case 0
            ' Recipe for STDPreparation
            DisableImage_Click
    
    End Select
End Sub
Private Sub SetFrameDefault()

FileName = ""
SelectedCode = ""

  
    Grid3.Rows = 1
  
    ' database
    PicDatabase.Visible = False
    
    DoEvents

End Sub
Private Sub PicMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

            
            frCommandInside(IndexDashCommInside).BackColor = &H644603
            lbCommandInside(IndexDashCommInside).ForeColor = &HE0E0E0
            
          If IndexDashCommInside = 1 Or IndexDashCommInside = 5 Or IndexDashCommInside = 8 Or IndexDashCommInside = 10 Then
              frCommandInside(IndexDashCommInside).BackColor = &H8000&
          End If
           If IndexDashCommInside = 5 Then
              frCommandInside(IndexDashCommInside).BackColor = &H80&
          End If
           If IndexDashCommInside = 11 Then
              frCommandInside(IndexDashCommInside).BackColor = vbColorManualPreparation
          End If
'
End Sub

Private Sub PicMenu_Click(Index As Integer)
SetVisibleDatabaseFrame (False)
If IndexProcedura = Index Then
Else
    Call SelectProcedura(Index)
End If
End Sub


Private Function SelectProcedura(ByVal Index As Integer) As Boolean
Dim rc As Boolean
Dim i As Integer
If Index > 3 Then Exit Function



Call SetFrameDefault


For i = 0 To PicMenu.UBound
    If i = Index Then
        PicMenu(i).BackColor = vbColorBlueProgram   '&H307030   '&H6D5B57
        'PicInfo(i).Visible = True
        'PicInfo(i).ZOrder
        
    Else
        'PicInfo(i).Visible = False
        PicMenu(i).BackColor = &H473733
    End If
Next
blTable = Label2(Index)
blTable.Visible = True

DoEvents


IndexProcedura = Index


PicMain(Index).ZOrder
PicMain(Index).Visible = True




Select Case IndexProcedura
    Case 0
    
        lbCommandInside(2) = "MR Search"
    
        Call FillStockDatabaseGrid(False)
       
        rc = True 'IIf(MyOperatore.IndexPrivilege > 0, True, False)
        Label2(6) = IIf(MyOperatore.Name <> "", MyOperatore.Name, Label2(6))
        frPrivilege.Visible = Not (rc)
        frStockControl.Visible = rc
        frStockControl.ZOrder
        USER_PATH = USER_TEMP_PATH
        
        
    Case 1
    
        

        Call ResetGrid2

        Call SetLine(cmbLine, True)
     
        frCloseFrame.Visible = False
    Case 2
    
        Call SetLine(cmbLineQC, True)
       

        frCloseFrame.Visible = False
    Case 3
        USER_PATH = USER_TEMP_PATH

        frCloseFrame.Visible = False
End Select

txSearch = "Search"

SaveSetting App.Title, "Intro", "IndexProcedura", IndexProcedura
TimeriNTRO.Enabled = False
frSelezione.ZOrder
        frSelezione.Visible = True
End Function

Private Sub ResetGrid2()

lbCommandInside(2) = "HannaCode Search"

txSearch = TX_INTRO
Call GetPreparationFromDatabase(Grid2, False)
frPreparationGrid.Visible = IIf(Grid2.Rows > 1, False, True)


frCommandInside(9).Visible = False

End Sub


Private Sub PicMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Form_MouseDown Button, Shift, x, y
End Sub

Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer

  Form_MouseMove Button, Shift, x, y
 
For i = PicMenu.LBound To PicMenu.UBound
    If i = IndexProcedura Then
        ' lascialo stare...
    ElseIf i = Index Then
        PicMenu(i).BackColor = &H5D4B47
    Else
        PicMenu(i).BackColor = &H473733
    End If
Next
End Sub

Private Sub PicMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseUp Button, Shift, x, y
End Sub


Private Sub SetPicForm()
'On Error GoTo ERR_SET:
Dim i As Integer
PicIntro.Left = 0
PicIntro.Top = PBTitle.Height
PicIntro.Width = Me.Width

BottomLine.x1 = 0
BottomLine.x2 = Me.Width



PicMain(0).Move 0, PBTitle.Top + PBTitle.Height, Me.Width, BottomLine.y1 - (PBTitle.Top + PBTitle.Height)
PicDatabase.Move 0, PBTitle.Top + PBTitle.Height, Me.Width, BottomLine.y1 - (PBTitle.Top + PBTitle.Height)
PicIntro.BackColor = &H929292

frStockControl.Move 0, 0, PicMain(0).Width ', lbWait.Top - 240




    
    For i = 1 To PicMain.UBound
        PicMain(i).Top = PicMain(0).Top
        PicMain(i).Left = PicMain(0).Left
        PicMain(i).Width = PicMain(0).Width
        PicMain(i).Height = PicMain(0).Height
    Next

lbWait.Left = 0
lbWait.Width = PicMain(0).Width
lbWait.Top = PicMain(0).Height - lbWait.Height

lbWaitPreparation.Left = 0
lbWaitPreparation.Width = PicMain(1).Width
lbWaitPreparation.Top = PicMain(1).Height - lbWaitPreparation.Height


Exit Sub
ERR_SET:
Resume Next
End Sub

Private Sub Picture4_Click(Index As Integer)
Dim MyNewIndex As Integer
Dim Frm As Form
Dim rc As Boolean
Dim StringName As String
If MyOperatore.Name = "" Then

    If frmLogin.DoShow Then
        Label2(6) = IIf(MyOperatore.Name <> "", MyOperatore.Name, Label2(6))
    Else
        Exit Sub
    End If
End If

End Sub


Private Sub CelarForm()
Dim i As Integer

  
    DisableImage.Visible = True
    Label10.Visible = DisableImage.Visible
    blTable.Visible = False
    
End Sub


Private Sub StartProcedure()
Call ClearText
Call SetPicForm
Call CelarForm


TX_INTRO = txSearch


End Sub


Private Sub ClearText()

Dim frmObj
    
   
    For Each frmObj In Me
        If TypeOf frmObj Is TextBox Then
            frmObj = ""
        End If
    Next

    DoEvents
    
End Sub






Private Sub MN_ONLINE_Click()

    CreateVerFile
    

    If FileExists(App.PATH & "\smartupdate.exe") Then
   
        ApriEseguibile App.PATH & "\smartupdate.exe"
        SaveSetting App.Title, "Opzioni", "Avvisa Update", True
    Else
       ' MessageCenterInfoTime = 2500
        PopupMessage 2, ("Attenzione impossibile trovare SmartUpdate.exe, Si consiglia di Reinstallare il programma."), , True
        SaveSetting App.Title, "Opzioni", "Avvisa Update", False
    End If



End Sub

Public Function SmartUpdate() As Boolean

Dim bAvvisami As Boolean

    bAvvisami = GetSetting(App.Title, "Opzioni", "Avvisa Update", True)
    If Not (bAvvisami) Then Exit Function

  
  
  If Not (CheckInternetConn) Then Exit Function
    
    With Inet1
        .url = "ftp://ftp.bilsoft.it"
        .UserName = "9620198@aruba.it"
        .Password = "qjE4Kb7NGhUF"
    End With
    



    StgVersione = App.Major & "." & App.Minor & "." & App.Revision
    
    
   
    If GetFileFromUrl(Inet1, "http://www.bilsoft.it/Download/" & PROGRAM_NAME & "/Update/", PROGRAM_NAME & ".txt") Then
      '-----------------------------------------------------------------------
      ' ok esiste il file con la versione
      '-----------------------------------------------------------------------
      StgVersioneftp = GetVerSoft(PROGRAM_NAME & ".txt")


      
      
      If EsistonoAggiornamenti(StgVersione, StgVersioneftp) Then
          '-----------------------------------
          ' esistono gli aggiornamenti
          '-----------------------------------

          Call GetSpecificheAggiornamento
          
          
          If F_MsgBox.DoShow(("New release available : ") & StgVersioneftp & vbCrLf & ("Current Release  : ") & StgVersione & vbCrLf & ("Update now?"), PROGRAM_NAME & " r." & StgVersioneftp, , ("Update"), (("Exit"))) Then
            
            MN_ONLINE_Click
            SaveSetting "Update " & App.EXEName, "UPDATE", "AVVISAMI", True
            'Call UploadMyInfo(Inet1)
          Else
               
                If F_MsgBox.DoShow(("Skip this version") & vbCrLf & ("Update software with Search Online Update funcion in Settings"), PROGRAM_NAME & " r." & StgVersioneftp, , ("Notify me"), ("Don't")) Then
                    
                    
                    SaveSetting App.Title, "Opzioni", "Avvisa Update", True
                    
                Else
                      SaveSetting App.Title, "Opzioni", "Avvisa Update", False
                      
                End If
                
          End If
      Else
          '-----------------------------------
          ' NO aggiornamenti
          '-----------------------------------
      End If
    End If
    
        
        
End Function




Public Function GetSpecificheAggiornamento() As String
   If GetFileFromUrl(Inet1, "http://www.bilsoft.it/Download/" & PROGRAM_NAME & "/Update/", PROGRAM_NAME & "_info.rtf") Then
   End If
End Function


Private Sub Timer1_Timer()

'
' qui inizia il programma , unica istanza, ritardata all'avvio
'
 frAvvio.ZOrder
frAvvio.Move 0, Me.ScaleHeight - frAvvio.Height, Me.ScaleWidth
frAvvio.Visible = True

      
    ' controllo aggiornamenti
    DoEvents
    If GetSetting(App.Title, "Opzioni", "Avvisa Update", True) Then Call SmartUpdate
    
     
 
    Call InitStrumenti
                
    
    mOk

 If SetWorkStation Then
 
 Else
    'MessageInfoTime = 2000
    'PopupMessage 2, "Please Enter Laboratory Info in Settings", , , , DefaultMenu(4)
    'F_SETTING.DoShow (3)
 End If
 
 
'--------------------------------
' qui finisce di caricare
'--------------------------------
frAvvio.Visible = True
Call SetPicMenu
frAvvio.Visible = True
frAvvio.Visible = False
Timer1.Enabled = False


End Sub


Private Function InitStrumenti()

    
    On Error Resume Next
   
  
    SearchInfoLabelPrinter
   

    
    'Call SetFormPrinter
     
    '------------------------------------------------
    ' BioPFD Printer
    '------------------------------------------------
    Dim ErrStr As String
    Dim DestStr As String
    If SetPDFPrinter(ErrStr, DestStr) Then
      
        bStampanteOK = True
       
    Else
        'non trovo la stampante virtuale BioPDF
        PopupMessage 2, "Errore stampante PDF :" & vbCrLf & ErrStr, , True, "Printer PDF"
        Frame2.Visible = True
        bStampanteOK = False
    End If

    On Error GoTo 0
End Function




Private Sub Timer2_Timer()
PicIntro.Visible = False
Timer2.Enabled = False
End Sub

Private Sub TimeriNTRO_Timer()
FormIntro
TimeriNTRO.Enabled = False
End Sub

Public Sub FormIntro()

Dim rc As Boolean
Dim Grid(10) As Grid
    
    
    rc = CheckPrimoAvvio
    
    Set Grid(0) = Grid1
    Set Grid(1) = Grid2
    Set Grid(2) = Grid3
  
 
    Call SetAllMainGrid(Grid())
    
    Call SetColumnWidth
  
    USER_PATH = USER_TEMP_PATH

    If rc Then
        SelectProcedura (GetSetting(App.Title, "Intro", "IndexProcedura", 0))
    End If
       
  SetFrameEmptyView
  
  
  

End Sub

Private Sub GetScheduledSTDInGrid()

lbWaitPreparation.Visible = True
'Call GetDataPreparationInGrid(Grid2, cmbLine, txSearch, bPreparationDetails)
If IndexProcedura = 1 Then blTable = "STD Schedule"

lbWaitPreparation.Visible = False
SetFrameEmptyView

End Sub





Private Function CheckPrimoAvvio() As Boolean
Dim rc As Boolean

    rc = True
    With dbTabCode
    
        .filter = ""
        If .EOF Then
            rc = False
            PicIntro.ZOrder
            PicDatabase.Visible = True
            IndexProcedura = 99
        Else
            PicDatabase.Visible = False
        End If
        
    End With
    
    With dbTabUserAccount
        .filter = ""
        bExistAccount = Not (.EOF)
    End With
    
    If bExistAccount Then
        
    Else
    
    End If
    Label14.Visible = Not (bExistAccount)
    Image6(4).Visible = Not (bExistAccount)
    Label9(1).Visible = Not (bExistAccount)
    Lab(0).Visible = Not (bExistAccount)
    
    
    CheckPrimoAvvio = rc
End Function




Private Sub txSearch_Change()
Dim rc As Boolean
rc = False


If txSearch = TX_INTRO Then Exit Sub

If txSearch = "" Then
    txSearch = TX_INTRO
    rc = True
End If


SearcMRInTable (rc)


End Sub


Private Sub SearcMRInTable(ByVal rc As Boolean)

Select Case IndexProcedura
    Case 0
        Call MRSearchInGrid(Grid1, txSearch, rc)
    Case 1
        Call MRSearchInGrid(Grid2, txSearch, rc)
    Case 2
        Call MRSearchInGrid(Grid3, txSearch, rc)
End Select


End Sub

Private Sub txSearch_Click()
If txSearch = TX_INTRO Then txSearch = " "
End Sub

Private Sub txSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then frCommandInside_Click 18
End Sub

Private Sub txSearch_LostFocus()
If Trim(txSearch) = "" Then txSearch = TX_INTRO
End Sub


Private Sub txSearSTDNumber_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then frCommandInside_Click 20
End Sub




Private Sub SetDatabaseHistory()



frDatabaseHistory.ZOrder
frDatabaseHistory.Move 0, PicMain(0).Top, PicMain(0).Width + PicMain(0).Left, PicMain(0).Height

SetVisibleDatabaseFrame True


End Sub

Private Function SetVisibleDatabaseFrame(ByVal rc As Boolean)
frDatabaseHistory.Visible = rc

DefaultMenu(2).Visible = Not (rc)
Label2(7).Visible = Not (rc)
End Function



Private Sub SetSTDPreparationDatabase()

    
        DoEvents
        DoEvents
        
        'FormProductionDatabaseHistory.Top = Me.Top
        'FormProductionDatabaseHistory.Left = Me.Left
        'FormProductionDatabaseHistory.WindowState = Me.WindowState
        'FormProductionDatabaseHistory.DoShow
        
End Sub


Private Sub SetPicMenu()

  DefaultMenu(2).Visible = True
  Label2(7).Visible = True

End Sub



Private Sub AddMRInStock()

Dim MRCode As String
    

If FormChemicalMR.DoShow(MRCode) Then
Else
    Exit Sub
End If

OpenWharehouseStoch (MRCode)

End Sub

Private Function OpenWharehouseStoch(ByVal MRCode As String)

If MRCode <> "" Then

DoEvents
DoEvents


    frAddInStock.Top = Me.Top
    frAddInStock.Left = Me.Left
    frAddInStock.Height = Me.Height
    frAddInStock.WindowState = Me.WindowState
    DoEvents
    If frAddInStock.DoShow(, MRCode) Then
    
    End If

    DoEvents

End If


End Function



Private Sub AddPreparation(Optional ByVal bSaltaSelect As Boolean, Optional ByVal bManual As Boolean)

On Error GoTo ERR_ADD:

Dim HannaCode As String
    

    If bSaltaSelect Then
    
         HannaCode = SelectedCode
         
    Else
        
        SelectedCode = ""
            If FormCodes.DoShow(HannaCode) Then
            
            
            Else
                Exit Sub
            End If
        
    
    End If


    Dim iHannaCode As HannaCode

   Call SetHannaCodeByCode(HannaCode, iHannaCode)
   DoEvents
   
    If bManual Then
        USER_PATH = USER_PREPARATION_PATH
        frmPreparation_Manual.Left = Me.Left
        frmPreparation_Manual.Top = Me.Top
        frmPreparation_Manual.WindowState = Me.WindowState
        frmPreparation_Manual.DoShow HannaCode, "", iHannaCode.ID
    Else

        If HannaCode <> "" Then
            USER_PATH = USER_PREPARATION_PATH
            
    
            
                frmPreparation_Static.Left = Me.Left
                frmPreparation_Static.Top = Me.Top
                frmPreparation_Static.WindowState = Me.WindowState
                frmPreparation_Static.DoShow HannaCode, "", iHannaCode.ID
            
            
            
    
        End If
    End If
    Call ResetGrid2

    
    
    
    

ERR_END:
    On Error GoTo 0
    
   ' Call ResetGrid2
    
    Exit Sub
ERR_ADD:
    MsgBox Err.Description
    Resume Next

End Sub

Private Sub SetPreparationDatabase()

       ' SetVisibleDatabaseFrame False
        DoEvents
        DoEvents
        
        FormPreparationDatabaseHistory.Top = Me.Top
        FormPreparationDatabaseHistory.Left = Me.Left
        FormPreparationDatabaseHistory.WindowState = Me.WindowState
        FormPreparationDatabaseHistory.DoShow
        
End Sub
Private Sub SetChemicalMRDatabase()
 
        FormChemicalMRDatabase.Top = Me.Top
        FormChemicalMRDatabase.Left = Me.Left
        FormChemicalMRDatabase.Width = Me.Width
        FormChemicalMRDatabase.Height = Me.Height
        FormChemicalMRDatabase.WindowState = Me.WindowState
        FormChemicalMRDatabase.DoShow
        
        
End Sub

