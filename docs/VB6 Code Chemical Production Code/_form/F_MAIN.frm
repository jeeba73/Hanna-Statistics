VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form F_MAIN 
   BackColor       =   &H004D3B37&
   Caption         =   "Chemical Production"
   ClientHeight    =   12045
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
   ScaleHeight     =   12555
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
      Height          =   10020
      Index           =   2
      Left            =   1680
      ScaleHeight     =   10020
      ScaleWidth      =   19200
      TabIndex        =   50
      Top             =   1080
      Visible         =   0   'False
      Width           =   19200
      Begin VB.Frame frCloseQC 
         BackColor       =   &H000040C0&
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
         Left            =   15480
         TabIndex        =   172
         Top             =   8760
         Visible         =   0   'False
         Width           =   3015
         Begin VB.Label lbClosedQC 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Close QC"
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
            TabIndex        =   173
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.Frame frInside 
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
         Height          =   8055
         Index           =   1
         Left            =   840
         TabIndex        =   77
         Top             =   480
         Width           =   17655
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
            Index           =   27
            Left            =   0
            TabIndex        =   209
            Top             =   4200
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Scan QRCode"
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
               Index           =   27
               Left            =   0
               TabIndex        =   210
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
            Index           =   26
            Left            =   3120
            TabIndex        =   207
            Top             =   4200
            Visible         =   0   'False
            Width           =   3015
            Begin VB.Label lbCommandInside 
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
               Index           =   26
               Left            =   0
               TabIndex        =   208
               Top             =   120
               Width           =   3015
            End
         End
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
            Left            =   14160
            Style           =   2  'Dropdown List
            TabIndex        =   176
            Top             =   4320
            Width           =   3495
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
            Left            =   0
            TabIndex        =   138
            Top             =   4920
            Width           =   17655
            Begin VB.Line Line12 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   17640
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label lbInside 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "QC per Recipe"
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
               TabIndex        =   139
               Top             =   120
               Width           =   3735
            End
         End
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
            TabIndex        =   80
            Top             =   1920
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
               TabIndex        =   154
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
               TabIndex        =   81
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
            TabIndex        =   78
            Top             =   0
            Width           =   17655
            Begin VB.Image ImMR 
               Height          =   240
               Index           =   3
               Left            =   17280
               Picture         =   "F_MAIN.frx":5E62
               Top             =   120
               Width           =   240
            End
            Begin VB.Image ImMR 
               Height          =   240
               Index           =   2
               Left            =   16680
               Picture         =   "F_MAIN.frx":6864
               Top             =   120
               Width           =   240
            End
            Begin VB.Label lbMR 
               BackStyle       =   0  'Transparent
               Height          =   495
               Index           =   3
               Left            =   17160
               TabIndex        =   180
               Top             =   0
               Width           =   2775
            End
            Begin VB.Label lbMR 
               BackStyle       =   0  'Transparent
               Height          =   495
               Index           =   2
               Left            =   14400
               TabIndex        =   179
               Top             =   0
               Width           =   2655
            End
            Begin VB.Label lbInside 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Recipes  in QC"
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
               TabIndex        =   79
               Top             =   120
               Width           =   4065
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
            Height          =   3495
            Left            =   0
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   600
            Width           =   17655
            _ExtentX        =   31141
            _ExtentY        =   6165
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
         Begin FlexCell.Grid Grid7 
            Height          =   2415
            Left            =   0
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   5520
            Width           =   17655
            _ExtentX        =   31141
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
         Left            =   840
         TabIndex        =   53
         Top             =   8760
         Width           =   3015
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Closed QC Table"
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
            TabIndex        =   54
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
         Index           =   8
         Left            =   6840
         TabIndex        =   55
         Top             =   8760
         Visible         =   0   'False
         Width           =   6255
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Recipe in Production"
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
            TabIndex        =   56
            Top             =   120
            Width           =   6255
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
         Index           =   6
         Left            =   6840
         TabIndex        =   51
         Top             =   8760
         Visible         =   0   'False
         Width           =   6255
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Set New QC"
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
            TabIndex        =   52
            Top             =   120
            Width           =   6255
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
         MouseIcon       =   "F_MAIN.frx":7266
         MousePointer    =   99  'Custom
         TabIndex        =   57
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
      Left            =   2160
      ScaleHeight     =   10020
      ScaleWidth      =   19200
      TabIndex        =   17
      Top             =   960
      Width           =   19200
      Begin VB.Frame frFromulation 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   8295
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   19095
         Begin VB.Frame frMaterialRequisition 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            Height          =   9735
            Left            =   13560
            TabIndex        =   89
            Top             =   -1200
            Visible         =   0   'False
            Width           =   19215
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
               Height          =   2535
               Index           =   4
               Left            =   1680
               TabIndex        =   118
               Top             =   240
               Width           =   15615
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
                  TabIndex        =   119
                  Top             =   0
                  Width           =   15615
                  Begin VB.Label lbMR 
                     BackStyle       =   0  'Transparent
                     Height          =   495
                     Index           =   1
                     Left            =   12120
                     TabIndex        =   127
                     Top             =   0
                     Width           =   2655
                  End
                  Begin VB.Label lbMR 
                     BackStyle       =   0  'Transparent
                     Height          =   495
                     Index           =   0
                     Left            =   14880
                     TabIndex        =   126
                     Top             =   0
                     Width           =   2775
                  End
                  Begin VB.Image ImMR 
                     Height          =   240
                     Index           =   1
                     Left            =   14400
                     Picture         =   "F_MAIN.frx":7570
                     Top             =   120
                     Width           =   240
                  End
                  Begin VB.Image ImMR 
                     Height          =   240
                     Index           =   0
                     Left            =   15000
                     Picture         =   "F_MAIN.frx":7F72
                     Top             =   120
                     Width           =   240
                  End
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
                     Index           =   4
                     Left            =   0
                     TabIndex        =   120
                     Top             =   75
                     Width           =   1185
                  End
               End
               Begin FlexCell.Grid Grid5 
                  Height          =   1815
                  Left            =   0
                  TabIndex        =   121
                  TabStop         =   0   'False
                  Top             =   600
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
               Height          =   4095
               Index           =   0
               Left            =   1680
               TabIndex        =   106
               Top             =   5400
               Width           =   15615
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
                  Index           =   16
                  Left            =   12240
                  TabIndex        =   122
                  Top             =   3600
                  Width           =   3135
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
                     Index           =   16
                     Left            =   0
                     TabIndex        =   123
                     Top             =   120
                     Width           =   3135
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
                  TabIndex        =   113
                  Top             =   0
                  Width           =   15255
                  Begin VB.Line Line9 
                     BorderColor     =   &H00B0B0B0&
                     X1              =   0
                     X2              =   15240
                     Y1              =   480
                     Y2              =   480
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
                     TabIndex        =   115
                     Top             =   120
                     Width           =   3015
                  End
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
                     Index           =   2
                     Left            =   13395
                     TabIndex        =   114
                     Top             =   180
                     Visible         =   0   'False
                     Width           =   1755
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
                  Index           =   15
                  Left            =   9120
                  TabIndex        =   111
                  Top             =   3600
                  Visible         =   0   'False
                  Width           =   3015
                  Begin VB.Label lbCommandInside 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "Delete"
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
                     TabIndex        =   112
                     Top             =   120
                     Width           =   3015
                  End
               End
               Begin VB.Frame Frame7 
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
                  TabIndex        =   109
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
                     TabIndex        =   110
                     Top             =   555
                     Width           =   1155
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
                  Index           =   14
                  Left            =   4920
                  TabIndex        =   107
                  Top             =   3600
                  Visible         =   0   'False
                  Width           =   4095
                  Begin VB.Label lbCommandInside 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "Check Out : Recipe in Preparation"
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
                     TabIndex        =   108
                     Top             =   120
                     Width           =   4095
                  End
               End
               Begin FlexCell.Grid Grid6 
                  Height          =   2535
                  Left            =   0
                  TabIndex        =   116
                  TabStop         =   0   'False
                  Top             =   720
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
               Begin VB.Line Line7 
                  BorderColor     =   &H00E0E0E0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   3360
                  Y2              =   3360
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Check out and go to Preparation"
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
                  Left            =   240
                  TabIndex        =   117
                  Top             =   3720
                  Width           =   2925
               End
            End
            Begin VB.Frame frIRequisition 
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   0  'None
               Caption         =   "Frame7"
               Height          =   2535
               Index           =   1
               Left            =   1680
               TabIndex        =   90
               Top             =   2760
               Width           =   15615
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
                  TabIndex        =   97
                  Top             =   0
                  Width           =   15255
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
                     TabIndex        =   99
                     Top             =   180
                     Visible         =   0   'False
                     Width           =   1755
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
                     TabIndex        =   98
                     Top             =   120
                     Width           =   3510
                  End
                  Begin VB.Line Line10 
                     BorderColor     =   &H00B0B0B0&
                     X1              =   0
                     X2              =   15240
                     Y1              =   480
                     Y2              =   480
                  End
               End
               Begin VB.TextBox txDocument 
                  BackColor       =   &H00E0E0E0&
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
                  ForeColor       =   &H00404040&
                  Height          =   300
                  Index           =   0
                  Left            =   4560
                  Locked          =   -1  'True
                  TabIndex        =   96
                  Top             =   960
                  Width           =   3255
               End
               Begin VB.TextBox txDocument 
                  BackColor       =   &H00E0E0E0&
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
                  ForeColor       =   &H00404040&
                  Height          =   300
                  Index           =   1
                  Left            =   9600
                  Locked          =   -1  'True
                  TabIndex        =   95
                  Top             =   1320
                  Width           =   2655
               End
               Begin VB.TextBox txDocument 
                  BackColor       =   &H00E0E0E0&
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
                  ForeColor       =   &H00404040&
                  Height          =   300
                  Index           =   2
                  Left            =   4560
                  Locked          =   -1  'True
                  TabIndex        =   94
                  Top             =   1680
                  Width           =   3255
               End
               Begin VB.TextBox txDocument 
                  BackColor       =   &H00E0E0E0&
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
                  ForeColor       =   &H00404040&
                  Height          =   300
                  Index           =   3
                  Left            =   4560
                  Locked          =   -1  'True
                  TabIndex        =   93
                  Top             =   1320
                  Width           =   3255
               End
               Begin VB.TextBox txDocument 
                  BackColor       =   &H00E0E0E0&
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
                  ForeColor       =   &H00404040&
                  Height          =   300
                  Index           =   4
                  Left            =   4560
                  Locked          =   -1  'True
                  TabIndex        =   92
                  Top             =   2040
                  Width           =   7695
               End
               Begin VB.TextBox txDocument 
                  BackColor       =   &H00E0E0E0&
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
                  ForeColor       =   &H00404040&
                  Height          =   300
                  Index           =   5
                  Left            =   9600
                  Locked          =   -1  'True
                  TabIndex        =   91
                  Top             =   1680
                  Width           =   2655
               End
               Begin VB.Label lbDocument 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Document No: MR-"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   375
                  Index           =   0
                  Left            =   2280
                  TabIndex        =   105
                  Top             =   960
                  Width           =   2130
               End
               Begin VB.Label lbDocument 
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
                  ForeColor       =   &H00404040&
                  Height          =   255
                  Index           =   1
                  Left            =   8160
                  TabIndex        =   104
                  Top             =   1320
                  Width           =   1275
               End
               Begin VB.Label lbDocument 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Production line no./dep."
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   255
                  Index           =   2
                  Left            =   2280
                  TabIndex        =   103
                  Top             =   1680
                  Width           =   2130
               End
               Begin VB.Label lbDocument 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Request Date"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   255
                  Index           =   3
                  Left            =   2280
                  TabIndex        =   102
                  Top             =   1320
                  Width           =   2130
               End
               Begin VB.Label lbDocument 
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
                  ForeColor       =   &H00404040&
                  Height          =   255
                  Index           =   4
                  Left            =   3960
                  TabIndex        =   101
                  Top             =   2040
                  Width           =   450
               End
               Begin VB.Label lbDocument 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Dep. Manager"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00404040&
                  Height          =   255
                  Index           =   5
                  Left            =   8160
                  TabIndex        =   100
                  Top             =   1680
                  Width           =   1275
               End
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   175
            Top             =   7320
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
            Height          =   6015
            Index           =   2
            Left            =   1200
            TabIndex        =   65
            Top             =   360
            Width           =   16935
            Begin VB.Timer Timer2 
               Interval        =   10
               Left            =   720
               Top             =   5040
            End
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
               Left            =   3840
               TabIndex        =   68
               Top             =   2640
               Visible         =   0   'False
               Width           =   8535
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe for Production"
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
                  TabIndex        =   156
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
                  TabIndex        =   69
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
               TabIndex        =   66
               Top             =   0
               Width           =   18255
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Active Recipe for production"
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
                  Index           =   2
                  Left            =   0
                  TabIndex        =   67
                  Top             =   120
                  Width           =   3285
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
               Height          =   5415
               Left            =   0
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   600
               Width           =   16815
               _ExtentX        =   29660
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
            Index           =   13
            Left            =   14880
            TabIndex        =   63
            Top             =   6720
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Closed Recipe Table"
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
               TabIndex        =   64
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
            Index           =   12
            Left            =   11640
            TabIndex        =   61
            Top             =   6720
            Visible         =   0   'False
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Delete Recipe"
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
               TabIndex        =   62
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
            Left            =   5160
            TabIndex        =   38
            Top             =   6720
            Visible         =   0   'False
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Open Recipe"
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
               TabIndex        =   39
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
            Left            =   6600
            TabIndex        =   36
            Top             =   7800
            Visible         =   0   'False
            Width           =   6255
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Material Requisition Check Out"
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
               TabIndex        =   37
               Top             =   120
               Width           =   6255
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
            Left            =   1560
            TabIndex        =   34
            Top             =   6720
            Width           =   3375
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "New Recipe For Production"
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
               Left            =   15
               TabIndex        =   35
               Top             =   120
               Width           =   3240
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
            Index           =   17
            Left            =   8400
            TabIndex        =   131
            Top             =   6720
            Visible         =   0   'False
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Close Recipe"
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
               Index           =   17
               Left            =   0
               TabIndex        =   132
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00B0B0B0&
            X1              =   1200
            X2              =   18120
            Y1              =   6480
            Y2              =   6480
         End
         Begin VB.Label lbRecipeForProductionInfo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Recipe Check Out : After Warehouse Approval, Laboratory Manager must Check Out before Preparation"
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   90
            TabIndex        =   60
            Top             =   7440
            Width           =   18960
         End
      End
      Begin VB.Frame frPrivilege 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1455
         Left            =   5160
         TabIndex        =   40
         Top             =   4320
         Visible         =   0   'False
         Width           =   9135
         Begin VB.Image DisableImage 
            Height          =   480
            Left            =   4080
            Picture         =   "F_MAIN.frx":8974
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
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   270
            Left            =   1320
            TabIndex        =   41
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
         TabIndex        =   128
         Top             =   9600
         Visible         =   0   'False
         Width           =   18975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recipe for Production"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   870
         Left            =   5715
         TabIndex        =   174
         Top             =   8400
         Width           =   7500
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
         MouseIcon       =   "F_MAIN.frx":BD56
         MousePointer    =   99  'Custom
         TabIndex        =   18
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
      Height          =   9420
      Index           =   1
      Left            =   1080
      ScaleHeight     =   9420
      ScaleWidth      =   19200
      TabIndex        =   42
      Top             =   720
      Visible         =   0   'False
      Width           =   19200
      Begin VB.Frame frClassification 
         BackColor       =   &H00000080&
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
         Height          =   615
         Index           =   0
         Left            =   1080
         TabIndex        =   181
         Top             =   8280
         Width           =   17535
         Begin VB.Label lbClassification 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "RM Classification"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
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
            TabIndex        =   182
            Top             =   120
            Width           =   17535
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
         Height          =   495
         Index           =   19
         Left            =   6240
         TabIndex        =   141
         Top             =   7440
         Visible         =   0   'False
         Width           =   3015
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Pass To Production"
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
            Index           =   19
            Left            =   0
            TabIndex        =   142
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.Frame frSearchRecipe 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   6000
         TabIndex        =   133
         Top             =   360
         Width           =   8175
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
            TabIndex        =   136
            Text            =   "Search"
            Top             =   160
            Width           =   3495
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
            Height          =   495
            Index           =   18
            Left            =   4080
            TabIndex        =   134
            Top             =   120
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Search Recipe"
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
               Index           =   18
               Left            =   0
               TabIndex        =   135
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.Image Image1 
            Height          =   360
            Left            =   7440
            Picture         =   "F_MAIN.frx":C060
            Top             =   160
            Width           =   360
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00B0B0B0&
            Height          =   735
            Left            =   0
            Top             =   0
            Width           =   8175
         End
      End
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
         TabIndex        =   130
         Top             =   7440
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
         Height          =   5895
         Index           =   0
         Left            =   840
         TabIndex        =   71
         Top             =   1200
         Width           =   17655
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
            TabIndex        =   74
            Top             =   0
            Width           =   17655
            Begin VB.Label lbPreparation 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "goto Preparation columns"
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
               Height          =   495
               Left            =   12720
               TabIndex        =   137
               Top             =   120
               Width           =   4935
            End
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
               Caption         =   "Preparation Table"
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
               Left            =   120
               TabIndex        =   75
               Top             =   120
               Width           =   2055
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
            TabIndex        =   72
            Top             =   2400
            Visible         =   0   'False
            Width           =   8535
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Preparation"
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
               Left            =   60
               TabIndex        =   155
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
               Index           =   1
               Left            =   30
               TabIndex        =   73
               Top             =   720
               Width           =   8415
            End
         End
         Begin FlexCell.Grid Grid2 
            Height          =   5295
            Left            =   120
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   600
            Width           =   17415
            _ExtentX        =   30718
            _ExtentY        =   9340
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
         Index           =   5
         Left            =   12480
         TabIndex        =   48
         Top             =   7440
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
            Index           =   5
            Left            =   0
            TabIndex        =   49
            Top             =   120
            Width           =   3015
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
         Height          =   495
         Index           =   4
         Left            =   9360
         TabIndex        =   46
         Top             =   7440
         Visible         =   0   'False
         Width           =   3015
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Pass To QC"
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
            TabIndex        =   47
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
         Index           =   3
         Left            =   15600
         TabIndex        =   44
         Top             =   7440
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
            Index           =   3
            Left            =   0
            TabIndex        =   45
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00B0B0B0&
         X1              =   1080
         X2              =   18600
         Y1              =   7200
         Y2              =   7200
      End
      Begin VB.Label lbWaitPreparation 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "Wait : Loading Data..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   129
         Top             =   9120
         Visible         =   0   'False
         Width           =   18975
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00B0B0B0&
         X1              =   1080
         X2              =   18360
         Y1              =   7200
         Y2              =   7200
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
         Index           =   0
         Left            =   15600
         MouseIcon       =   "F_MAIN.frx":DAD2
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   960
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
      Height          =   9420
      Index           =   3
      Left            =   240
      ScaleHeight     =   9420
      ScaleWidth      =   19200
      TabIndex        =   58
      Top             =   1200
      Visible         =   0   'False
      Width           =   19200
      Begin VB.Frame frClassification 
         BackColor       =   &H00000080&
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
         Height          =   615
         Index           =   2
         Left            =   600
         TabIndex        =   185
         Top             =   8520
         Width           =   8895
         Begin VB.Label lbClassification 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Recipes Classification"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   186
            Top             =   120
            Width           =   8895
         End
      End
      Begin VB.Frame frClassification 
         BackColor       =   &H00000080&
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
         Height          =   615
         Index           =   1
         Left            =   9720
         TabIndex        =   183
         Top             =   8520
         Width           =   8895
         Begin VB.Label lbClassification 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hanna Codes Classification"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   184
            Top             =   120
            Width           =   8895
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   6000
         TabIndex        =   143
         Top             =   360
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
            Index           =   20
            Left            =   4080
            TabIndex        =   145
            Top             =   120
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Search Hanna Code"
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
               Index           =   20
               Left            =   0
               TabIndex        =   146
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.TextBox txSearchCode 
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
            TabIndex        =   144
            Text            =   "Search"
            Top             =   160
            Width           =   3495
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00B0B0B0&
            Height          =   735
            Left            =   0
            Top             =   0
            Width           =   8175
         End
         Begin VB.Image Image2 
            Height          =   360
            Left            =   7440
            Picture         =   "F_MAIN.frx":DDDC
            Top             =   160
            Width           =   360
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
         Index           =   3
         Left            =   840
         TabIndex        =   83
         Top             =   1200
         Width           =   17775
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
            Index           =   23
            Left            =   14520
            TabIndex        =   170
            Top             =   6480
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
               Index           =   23
               Left            =   0
               TabIndex        =   171
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.ComboBox cmbLineProduction 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   157
            Top             =   6480
            Width           =   2775
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
            Index           =   10
            Left            =   3000
            TabIndex        =   151
            Top             =   6480
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "New Production"
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
               TabIndex        =   152
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
            Index           =   11
            Left            =   11400
            TabIndex        =   147
            Top             =   6480
            Visible         =   0   'False
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Close Production"
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
               TabIndex        =   148
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
            Index           =   3
            Left            =   0
            TabIndex        =   86
            Top             =   0
            Width           =   17775
            Begin VB.Line Line5 
               BorderColor     =   &H00E0E0E0&
               X1              =   0
               X2              =   17280
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label lbInside 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Production"
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
               Index           =   3
               Left            =   0
               TabIndex        =   87
               Top             =   120
               Width           =   1245
            End
         End
         Begin VB.Frame Frame5 
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
            Left            =   4320
            TabIndex        =   84
            Top             =   2880
            Visible         =   0   'False
            Width           =   8535
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Production"
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
               Index           =   5
               Left            =   105
               TabIndex        =   153
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
               Index           =   3
               Left            =   90
               TabIndex        =   85
               Top             =   720
               Width           =   8415
            End
         End
         Begin FlexCell.Grid Grid4 
            Height          =   5535
            Left            =   0
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   600
            Width           =   17295
            _ExtentX        =   30506
            _ExtentY        =   9763
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
            Index           =   9
            Left            =   6120
            TabIndex        =   149
            Top             =   6480
            Visible         =   0   'False
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Open Production"
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
               TabIndex        =   150
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.Line Line15 
            BorderColor     =   &H00B0B0B0&
            X1              =   120
            X2              =   17520
            Y1              =   6240
            Y2              =   6240
         End
      End
      Begin VB.Label Label7 
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
         MouseIcon       =   "F_MAIN.frx":F84E
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   480
         Width           =   60
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   9735
      Index           =   4
      Left            =   1320
      ScaleHeight     =   9735
      ScaleWidth      =   18135
      TabIndex        =   189
      Top             =   1320
      Width           =   18135
      Begin VB.Frame Frame4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   6000
         TabIndex        =   201
         Top             =   360
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
            Index           =   25
            Left            =   4080
            TabIndex        =   203
            Top             =   120
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Search Recipe"
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
               Index           =   25
               Left            =   0
               TabIndex        =   204
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.TextBox txSearchRecipe 
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
            TabIndex        =   202
            Text            =   "Search"
            Top             =   160
            Width           =   3495
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00B0B0B0&
            Height          =   735
            Left            =   0
            Top             =   0
            Width           =   8175
         End
         Begin VB.Image Image4 
            Height          =   360
            Left            =   7440
            Picture         =   "F_MAIN.frx":FB58
            Top             =   160
            Width           =   360
         End
      End
      Begin VB.Frame frInside 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "16800"
         Height          =   7935
         Index           =   6
         Left            =   1320
         TabIndex        =   190
         Top             =   1320
         Width           =   17055
         Begin VB.Frame Frame6 
            BackColor       =   &H00886010&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   5880
            TabIndex        =   196
            Top             =   2760
            Width           =   5055
            Begin VB.Label lbChem 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Empty List..."
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
               Index           =   5
               Left            =   1920
               TabIndex        =   198
               Top             =   360
               Width           =   1155
            End
            Begin VB.Label lbChem 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Enter Recipe in Search field."
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
               Left            =   1185
               TabIndex        =   197
               Top             =   720
               Width           =   2730
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
            Index           =   8
            Left            =   0
            TabIndex        =   193
            Top             =   0
            Width           =   17055
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Recipe history"
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
               Left            =   15555
               TabIndex        =   195
               Top             =   120
               Width           =   1245
            End
            Begin VB.Label lbInside 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Formulation History"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00644603&
               Height          =   420
               Index           =   8
               Left            =   0
               TabIndex        =   194
               Top             =   0
               Width           =   3300
            End
            Begin VB.Line Line16 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   16920
               Y1              =   480
               Y2              =   480
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
            Left            =   0
            TabIndex        =   191
            Top             =   6960
            Width           =   3015
            Begin VB.Image Image 
               Height          =   480
               Index           =   1
               Left            =   120
               MousePointer    =   99  'Custom
               OLEDropMode     =   1  'Manual
               Picture         =   "F_MAIN.frx":115CA
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
               TabIndex        =   192
               Top             =   120
               Width           =   3015
            End
         End
         Begin FlexCell.Grid Grid8 
            Height          =   6015
            Left            =   0
            TabIndex        =   199
            TabStop         =   0   'False
            Top             =   600
            Width           =   16935
            _ExtentX        =   29871
            _ExtentY        =   10610
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
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formulation History : Search per Recipe or select from Preparation Table"
            ForeColor       =   &H00808080&
            Height          =   240
            Index           =   2
            Left            =   5640
            TabIndex        =   200
            Top             =   7560
            Width           =   5745
         End
         Begin VB.Line Line17 
            BorderColor     =   &H00D0D0D0&
            X1              =   0
            X2              =   16920
            Y1              =   6720
            Y2              =   6720
         End
      End
      Begin VB.Image IconFH 
         Height          =   480
         Left            =   720
         OLEDropMode     =   1  'Manual
         Picture         =   "F_MAIN.frx":149AC
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Timer TimeriNTRO 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1080
      Top             =   6960
   End
   Begin VB.PictureBox PicIntro 
      AutoSize        =   -1  'True
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
      Height          =   9690
      Left            =   3120
      MouseIcon       =   "F_MAIN.frx":17D8E
      MousePointer    =   99  'Custom
      ScaleHeight     =   9690
      ScaleWidth      =   19200
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   19200
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "Wait : Loading Data..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   205
         Top             =   6600
         Width           =   19215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         Left            =   0
         TabIndex        =   25
         Top             =   5160
         Width           =   19170
      End
      Begin VB.Label lbProgram 
         Alignment       =   2  'Center
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
         Left            =   0
         TabIndex        =   20
         Top             =   5880
         Width           =   19125
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chemical Production"
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
         Left            =   0
         TabIndex        =   16
         Top             =   3600
         Width           =   19140
      End
   End
   Begin VB.Frame frAvvio 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   975
      Left            =   0
      TabIndex        =   166
      Top             =   10080
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
         TabIndex        =   168
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
         TabIndex        =   167
         Top             =   120
         Width           =   19215
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame frSelezione 
      BackColor       =   &H00307030&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   100
      Left            =   0
      TabIndex        =   169
      Top             =   960
      Visible         =   0   'False
      Width           =   19215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   7800
   End
   Begin VB.Frame frCloseFrame 
      BackColor       =   &H004D3B37&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   8640
      TabIndex        =   124
      Top             =   10920
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Image DefaultMenu 
         DragIcon        =   "F_MAIN.frx":18098
         Height          =   480
         Index           =   1
         Left            =   720
         MouseIcon       =   "F_MAIN.frx":1B47A
         MousePointer    =   99  'Custom
         Picture         =   "F_MAIN.frx":1B784
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
            Size            =   9.75
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
         MouseIcon       =   "F_MAIN.frx":1EB66
         MousePointer    =   99  'Custom
         TabIndex        =   125
         Top             =   795
         Width           =   975
      End
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
         Index           =   4
         Left            =   7800
         MouseIcon       =   "F_MAIN.frx":1EE70
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   187
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Formulation history"
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
            Index           =   4
            Left            =   0
            MouseIcon       =   "F_MAIN.frx":1F17A
            MousePointer    =   99  'Custom
            TabIndex        =   188
            Top             =   640
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   4
            Left            =   720
            MouseIcon       =   "F_MAIN.frx":1F484
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":1F78E
            Top             =   120
            Width           =   480
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
         Index           =   3
         Left            =   5760
         MouseIcon       =   "F_MAIN.frx":22B70
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   13
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   720
            MouseIcon       =   "F_MAIN.frx":22E7A
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":23184
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
            Index           =   3
            Left            =   0
            MouseIcon       =   "F_MAIN.frx":25B76
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   640
            Width           =   1890
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
         Index           =   2
         Left            =   3840
         MouseIcon       =   "F_MAIN.frx":25E80
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   11
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   720
            MouseIcon       =   "F_MAIN.frx":2618A
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":26494
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "QC"
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
            Index           =   2
            Left            =   0
            MouseIcon       =   "F_MAIN.frx":28E86
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   640
            Width           =   1890
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
         MouseIcon       =   "F_MAIN.frx":29190
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
            MouseIcon       =   "F_MAIN.frx":2949A
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":297A4
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preparation"
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
            Left            =   450
            MouseIcon       =   "F_MAIN.frx":2C196
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   640
            Width           =   990
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
         MouseIcon       =   "F_MAIN.frx":2C4A0
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
            Caption         =   "Recipe for Prod."
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
            Left            =   345
            MouseIcon       =   "F_MAIN.frx":2C7AA
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   645
            Width           =   1320
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   735
            MouseIcon       =   "F_MAIN.frx":2CAB4
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":2CDBE
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
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   4260
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
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Image Image6 
         Height          =   480
         Index           =   8
         Left            =   7680
         Picture         =   "F_MAIN.frx":2F7B0
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
         MouseIcon       =   "F_MAIN.frx":32B92
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   4320
         Visible         =   0   'False
         Width           =   19140
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   4
         Left            =   6480
         Picture         =   "F_MAIN.frx":32E9C
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
         MouseIcon       =   "F_MAIN.frx":3627E
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   6720
         Visible         =   0   'False
         Width           =   19185
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Operator  / User Account"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   450
         Index           =   1
         Left            =   120
         TabIndex        =   30
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
         MouseIcon       =   "F_MAIN.frx":36588
         MousePointer    =   99  'Custom
         TabIndex        =   29
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
         MouseIcon       =   "F_MAIN.frx":36892
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   4080
         Width           =   19125
      End
      Begin VB.Image Im 
         Height          =   480
         Index           =   7
         Left            =   10920
         Picture         =   "F_MAIN.frx":36B9C
         Top             =   3960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Code"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   450
         Index           =   0
         Left            =   0
         TabIndex        =   27
         Top             =   1680
         Width           =   19170
      End
   End
   Begin VB.TextBox txQRCode 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   211
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   14535
   End
   Begin VB.Frame frDatabaseHistory 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   9735
      Left            =   240
      TabIndex        =   158
      Top             =   840
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
         Index           =   24
         Left            =   3120
         TabIndex        =   177
         Top             =   5040
         Width           =   6000
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chemical RM"
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
            Index           =   24
            Left            =   1815
            TabIndex        =   178
            Top             =   240
            Width           =   2370
         End
         Begin VB.Image Image 
            Height          =   480
            Index           =   0
            Left            =   360
            MouseIcon       =   "F_MAIN.frx":39F7E
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":3A288
            Top             =   240
            Width           =   480
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
         Index           =   22
         Left            =   9960
         TabIndex        =   161
         Top             =   3840
         Width           =   6000
         Begin VB.Image Image 
            Height          =   480
            Index           =   22
            Left            =   360
            MouseIcon       =   "F_MAIN.frx":3D66A
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":3D974
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Production"
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
            Index           =   22
            Left            =   0
            TabIndex        =   162
            Top             =   240
            Width           =   5940
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
         Index           =   21
         Left            =   3120
         TabIndex        =   159
         Top             =   3840
         Width           =   6000
         Begin VB.Image Image 
            Height          =   480
            Index           =   21
            Left            =   360
            MouseIcon       =   "F_MAIN.frx":40366
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":40670
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preparation"
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
            Index           =   21
            Left            =   75
            TabIndex        =   160
            Top             =   240
            Width           =   5850
         End
      End
      Begin VB.Label llbExit 
         BackStyle       =   0  'Transparent
         Height          =   1455
         Left            =   7680
         TabIndex        =   165
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   8625
         MouseIcon       =   "F_MAIN.frx":43062
         MousePointer    =   99  'Custom
         TabIndex        =   164
         Top             =   9300
         Width           =   1920
      End
      Begin VB.Image DefaultExit 
         DragIcon        =   "F_MAIN.frx":4336C
         Height          =   480
         Left            =   9360
         MouseIcon       =   "F_MAIN.frx":4674E
         MousePointer    =   99  'Custom
         Picture         =   "F_MAIN.frx":46A58
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
         TabIndex        =   163
         Top             =   2040
         Width           =   6960
      End
   End
   Begin VB.Label lbWaitMain 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "Wait : Loading Data..."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   206
      Top             =   11400
      Visible         =   0   'False
      Width           =   8535
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
      MouseIcon       =   "F_MAIN.frx":49E3A
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
      MouseIcon       =   "F_MAIN.frx":4A144
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
      Left            =   13080
      MouseIcon       =   "F_MAIN.frx":4A44E
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
      MouseIcon       =   "F_MAIN.frx":4A758
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
      MouseIcon       =   "F_MAIN.frx":4AA62
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
         Size            =   9.75
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
      MouseIcon       =   "F_MAIN.frx":4AD6C
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   11715
      Width           =   1110
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Databases"
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
      Left            =   13965
      MouseIcon       =   "F_MAIN.frx":4B076
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   11715
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set Operator"
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
      Index           =   6
      Left            =   4260
      MouseIcon       =   "F_MAIN.frx":4B380
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   11715
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
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
      Left            =   240
      MouseIcon       =   "F_MAIN.frx":4B68A
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   11715
      Width           =   645
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   14160
      MouseIcon       =   "F_MAIN.frx":4B994
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":4BC9E
      Top             =   11160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   4560
      MouseIcon       =   "F_MAIN.frx":4F080
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":4F38A
      Top             =   11160
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   360
      MouseIcon       =   "F_MAIN.frx":5276C
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":52A76
      Top             =   11160
      Width           =   480
   End
   Begin VB.Line BottomLine 
      BorderColor     =   &H00C0C0C0&
      Visible         =   0   'False
      X1              =   0
      X2              =   19320
      Y1              =   11507.45
      Y2              =   11507.45
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
      Y2              =   12383.01
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   9660
      X2              =   9660
      Y1              =   375.243
      Y2              =   12758.26
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_MAIN.frx":55E58
      Height          =   480
      Index           =   0
      Left            =   18120
      MouseIcon       =   "F_MAIN.frx":5923A
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":59544
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
Private MyLot As String
Private MyCode As String
    
Private lRowCode As Long
Private lRow As Long

Private TimerCount As Integer

Private SelectedCode As String
Private SelectedCodeID As Long
Private m_rc As Boolean

Private ReceiptFileName As String
Private PreparationFileName As String
Private PreparationQC As String
Private bNoPreparation   As Boolean
Private ReceiptID As Long
Private IndexDashCommInside As Integer
Private IndexRecipeMaterialRequisition As Integer
Private lRowGridMaterialRequisition As Integer
Private bReceiptHistoryView As Boolean
Private PreparationID As Long
Private PreparationCode As String
Private TX_INTRO As String
Private bPreparationDetails As Boolean
Private bClosedQc As Boolean
Private ProductionID  As Long
Private ProductionFileName  As String
Private HannacodeProduction As String
Private RecipeProduction As String
Private RevisionID As Long
Private RecipeCode As String
Private UserQrCode As QRCodeType


Private Sub cmbLine_Click()
GetPreparationInGrid
End Sub

Private Sub Command1_Click()
F_MAIN_CLP.Show
End Sub

Private Sub cmbLineProduction_Click()
GetProductionInGrid
End Sub



Private Sub cmbLineQC_Click()
Grid7.Rows = 1
Call GetQcInGrid(Grid3, cmbLineQC, , bClosedQc, Frame2)

End Sub

Private Sub cmbLineRfP_Click()
Call GetReceiptFromDatabase(Grid1, bReceiptHistoryView, , cmbLineRfP)
bCODLine = IIf(InStr(cmbLineRfP, "59"), True, False)
End Sub

Private Sub DefaultMenu_Click(Index As Integer)
DefaultMenuLabel_Click Index
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
            If F_MsgBox.DoShow("Quit Chemical Production?", "Exit") Then

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
        F_SETTING.DoShow , , , DefaultMenu(4)
        'FormIntro
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
frFromulation.ZOrder
frFromulation.Visible = rc
frPrivilege.Visible = Not (rc)
frFromulation.ZOrder

End Sub

Private Sub Frame3_Click(Index As Integer)
    
    Select Case Index
        Case 0
            ' preparation // refresh grid
            Call GetPreparationInGrid
            
    
    End Select
End Sub

Private Sub frClassification_Click(Index As Integer)
     Dim MyID As Long
     
   Select Case Index
        Case 0


            Call FormChemicalRM.DoShow(PreparationCode, , , 1)
     
        Case 1
          If HannacodeProduction <> "" Then
                
                If F_MsgBox.DoShow("Open " & HannacodeProduction & " Hanna Code Classification?", "Classification", , "Open", "View Hanna Codes") Then
                    
              
                    MyID = GetHannaCodeID(HannacodeProduction)
                    If MyID = 0 Then GoTo HannaCode
                    Call F_PICTOGRAM.DoShow(MyID, 0, HannacodeProduction, True)
                Else
                    GoTo HannaCode:
                End If
            
            Else
HannaCode:
                Call FormCodes.DoShow(HannacodeProduction, , , 1)
            
            End If
        Case 2
        
        
        
                  If RecipeProduction <> "" Then
                
                If F_MsgBox.DoShow("Open " & RecipeProduction & " Recipe Classification?", "Classification", , "Open", "View Recipes") Then
                    
                  
                    MyID = GetRecipeIdByName(RecipeProduction)
                    If MyID = 0 Then GoTo Recipe
                    Call F_PICTOGRAM.DoShow(MyID, 2, RecipeProduction, True)
                Else
                    GoTo Recipe:
                End If
            
            Else
Recipe:

                Call FormRecipes.DoShow(RecipeProduction, , 1)
            
            End If
    End Select
End Sub

Private Sub frCloseFrame_Click()
DefaultMenuLabel_Click 1
End Sub


Private Sub frCloseQC_Click()
Dim rc As Boolean

If MyOperatore.Name = "" Then

    If frmLogin.DoShow Then
            'Label2(6) = IIf(MyOperatore.Name <> "", MyOperatore.Name, Label2(6))
    Else
        Exit Sub
    End If

End If
    
If F_MsgBox.DoShow("Close QC Recipe/Mix?", PreparationCode) Then
    
        rc = RecipeCloseQC(PreparationFileName, PreparationID)
    
        If rc Then
            Call GetQcInGrid(Grid3, cmbLineQC, , bClosedQc, Frame2)
            Call GetProductionInGrid4(Grid4, False, txSearchCode, Frame5, cmbLineProduction)
            Grid7.Rows = 1
            frCommandInside(6).Visible = False
            frCommandInside(8).Visible = False
            frCloseQC.Visible = False
            PopupMessage 2, "QC Closed....", , , PreparationCode
        End If
    End If
End Sub

Private Sub frPrivilege_Click()
DisableImage_Click
End Sub
Private Sub Form_Activate()
CloseSettingDataFile
    Label2(6) = IIf(MyOperatore.Name <> "", MyOperatore.Name, Label2(6))

Dim rc As Boolean

        rc = IIf(MyOperatore.IndexPrivilege >= 1, True, False)
        frPrivilege.Visible = Not (rc)
        frFromulation.Visible = rc
        frFromulation.ZOrder
      
'DefaultMenuLabel(8).Move DefaultMenuLabel(9).Left, DefaultMenuLabel(9).Top, DefaultMenuLabel(9).Width, DefaultMenuLabel(9).Height
'Image3(8).Left = Image4(1).Left
'Image3(8).Top = Image4(1).Top

bCODLine = IIf(InStr(UserLine, "59"), True, False)

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
    Dim svar
    svar = EnumWindows(AddressOf getalltopwindows, 0)
    End
    
    
End Sub



Private Sub frCommandInside_Click(Index As Integer)
    Select Case Index
        Case 0
            ' Recipe check out
            Call RecipeCheckOutButton
        Case 1
            ' new RecipeForProduction
            Call NewRecipeForProductionButton
            
        Case 2
            ' open RecipeForProduction
            Grid1_DblClick
        Case 3
            ' delete preparation
            Call DeletePreparation
        Case 4
            ' pass to QC
            Call PassToQC
            
        Case 5
            Call OpenPreparationButton
        Case 6
            ' add qc to recipe
            Call AddQcToRecipe
        Case 7
            ' closed preparation Recipe
            Call ClosedQCTable
        Case 8
            ' move recipe in production
            Call MoveRecipeInProduction(False, False)
        Case 9
            ' open production
            Grid4_DblClick
        Case 10
            Call NewProductionButton
        Case 11
            ' production CheckOut
            Call ProductionCheckOut
        Case 12
            ' delete Recipe for Production
            Call DeleteRfPFunction
            
        Case 13
            ' Recipe for production History!!!!
            Call ReceiptHistory
        Case 14
            ' MaterialRequisition get CheckOut!
            Call MaterialRequisitionCheckOutButton
        Case 16
            ' exit frMaterialRequisition
            Call ExitMaterialRequisitionButton
        Case 17
            Call SetCloseRecipeForProduction
        Case 18
            ' search recipe
            Call GetPreparationInGrid
        Case 19
            ' pass to production
            Call MoveRecipeInProduction(False, False)
        Case 20
            ' search Hanna Code
            Call GetProductionInGrid
        Case 21
            ' database history : preparation
            Call SetPreparationDatabase
        Case 22
            ' database history  production
            Call SetProductionDatabase
        Case 23
            ' Delete production
            Call ProductionDelete
        Case 24
            Call SetChemicalRMDatabase
        Case 25
             Call OpenRevisionHistory
        Case 26
            If F_MsgBox.DoShow("Print Recipe Labels to QC ?", RecipeCode) Then
                Call PrintLabelPreparation(PreparationFileName, PreparationID, RecipeCode, bCODLine)
            End If
             
        Case 27
            ' SCANN BARCODE
            Call ScanQRCodeQC
            
            
    End Select
    
    
    Call SetFrameEmptyView
    
End Sub


Private Function OpenPreparationButton()

    If MyOperatore.Name = "" Then
    
        If frmLogin.DoShow Then
                'Label2(6) = IIf(MyOperatore.Name <> "", MyOperatore.Name, Label2(6))
        Else
            Exit Function
        End If
    
    End If
    
    If bNoPreparation Then
            
        PopupMessage 2, "Recipe doesn't need Preparation." & vbCrLf & "Move to Production?", , , PreparationCode
        
        frCommandInside_Click 19
    
        Exit Function

    End If
    
    If PreparationCode <> "" Then
        USER_PATH = USER_PREPARATION_PATH
        frmPreparation.Left = Me.Left
        frmPreparation.Top = Me.Top
        frmPreparation.WindowState = Me.WindowState
        frmPreparation.DoShow PreparationCode, PreparationFileName, PreparationID
        'Call GetReceiptFromDatabase(Grid1, bReceiptHistoryView)
    End If
    
    Call GetPreparationInGrid

End Function
Private Function NewRecipeForProductionButton()

    USER_PATH = USER_TEMP_PATH
    frmRecipeForProduction.Left = Me.Left
    frmRecipeForProduction.Top = Me.Top
    frmRecipeForProduction.WindowState = Me.WindowState
    frmRecipeForProduction.DoShow
    Call GetReceiptFromDatabase(Grid1, bReceiptHistoryView, , cmbLineRfP)
End Function
Private Function NewProductionButton()

    If F_MsgBox.DoShow("Open new Production without Recipe for Production?", "Production") Then
        USER_PATH = USER_PRODUCTION_PATH
        frmProduction.Left = Me.Left
        frmProduction.Top = Me.Top
        frmProduction.WindowState = Me.WindowState
        frmProduction.DoShow
        'Call GetReceiptFromDatabase(Grid1, bReceiptHistoryView)
    End If
    
     GetProductionInGrid
End Function


Private Function ExitMaterialRequisitionButton()
 
    frMaterialRequisition.Visible = False
    frCloseFrame.Visible = False
    frCommandInside(15).Visible = False
    frCommandInside(14).Visible = False
    Call GetReceiptFromDatabase(Grid1, bReceiptHistoryView, , cmbLineRfP)

End Function

Private Function RecipeCheckOutButton()

    lbWaitMain.ZOrder
    lbWaitMain.Visible = True
    lbWait = "Wait : Loading Data..."
    lbWait.Visible = True
    DoEvents
    frCommandInside(0).Visible = False
    USER_PATH = IIf(bReceiptHistoryView, USER_DATA_PATH, USER_TEMP_PATH)
    Call SetRecipeGrid5(Grid5, ReceiptFileName)
    frMaterialRequisition.ZOrder
    frFromulation.Height = PicMain(1).Height
    frMaterialRequisition.Height = PicMain(1).Height
    
    frMaterialRequisition.Visible = True
    
    Label2(9) = "Close Material Req."
    frCloseFrame.Visible = True
    lbWait.Visible = False
    lbWaitMain.Visible = False
End Function
            
          
Private Function MaterialRequisitionCheckOutButton()

Dim strRecipe As String

    strRecipe = Grid5.Cell(lRowGridMaterialRequisition, 1).Text
    If strRecipe = "" Then
        PopupMessage 2, "Select a Recipe first..."
    Else
        If F_MsgBox.DoShow("Check Out Recipe :  " & strRecipe & " ? ", "Recipe for prod.") Then
            USER_PATH = IIf(bReceiptHistoryView, USER_DATA_PATH, USER_TEMP_PATH)
            lbWait = "Wait : Saving Data | Moving to Preparation..."
            lbWait.Visible = True
            Call CheckOutMaterialRequisition(IndexRecipeMaterialRequisition)
            Call SetRecipeGrid5(Grid5, ReceiptFileName)
            lbWait.Visible = False
        End If
    End If

End Function


Private Function SetFrameEmptyView()
    
    frReceiptGrid.Visible = IIf(Grid1.Rows > 1, False, True)
    frPreparationGrid.Visible = IIf(Grid2.Rows > 1, False, True)
    
    
   ' cmbLine.Visible = Not (frPreparationGrid.Visible)
    frCommandInside(5).Visible = False
    frCommandInside(3).Visible = False
    
    
    


End Function
Private Sub frCommandInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
IndexDashCommInside = Index
Dim i As Integer
    For i = 0 To frCommandInside.UBound
        If i = Index Then
            frCommandInside(i).BackColor = &H846623
            lbCommandInside(i).ForeColor = vbWhite
            If i = 1 Or i = 5 Or i = 6 Or i = 14 Or i = 9 Then
                frCommandInside(i).BackColor = &H20A020
            End If
            
        Else
            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 1 Or i = 5 Or i = 6 Or i = 14 Or i = 9 Then
                frCommandInside(i).BackColor = &H8000&
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
    

    

    If ReceiptFileName <> "" Then
        frmRecipeForProduction.Left = Me.Left
        frmRecipeForProduction.Top = Me.Top
        frmRecipeForProduction.WindowState = Me.WindowState
        
        Call frmRecipeForProduction.DoShow(, ReceiptFileName)
        Call GetReceiptFromDatabase(Grid1, bReceiptHistoryView, , cmbLineRfP)
        
    End If
    
    SetFrameEmptyView
        
End Sub

Private Sub Grid1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
'PicMenu(4).Visible = False
Dim rc As Boolean
lRow = FirstRow
    USER_PATH = IIf(bReceiptHistoryView, USER_DATA_PATH, USER_TEMP_PATH)
    
    ReceiptFileName = ""
    ReceiptID = 0
    rc = False
    frCommandInside(0).Visible = False
    frCommandInside(2).Visible = False
    frCommandInside(12).Visible = False
    frCommandInside(17).Visible = False
     
    If FirstRow > 0 Then
        rc = Grid1.Cell(FirstRow, 9).Text
        ReceiptFileName = Grid1.Cell(FirstRow, 12).Text
       ' ReceiptFileName = Replace(ReceiptFileName, ".prep", ".rfp")
        ReceiptID = Grid1.Cell(FirstRow, 13).Text
        frCommandInside(2).Visible = True
        frCommandInside(12).Visible = True
         frCommandInside(0).Visible = True 'rc
         frCommandInside(17).Visible = True
    End If
    
    
   
   

End Sub

Private Sub Grid2_BeforeUserSort(ByVal Col As Long)
lRow = 0
'PicMenu(4).Visible = False
End Sub

Private Sub Grid2_DblClick()
If lRow = 0 Then Exit Sub
frCommandInside_Click 5
'PicMenu(4).Visible = False
End Sub


Private Sub Grid2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
lRow = FirstRow
    'PicMenu(4).Visible = False
    RevisionID = 0
    RecipeCode = ""
    PreparationQC = ""
    PreparationCode = ""
    bNoPreparation = False
    frCommandInside(4).Visible = False
    frCommandInside(19).Visible = False
    bNoPreparation = False
    
    frCommandInside(5).Visible = False
    frCommandInside(3).Visible = False
    
    If FirstRow > 0 Then
        'PicMenu(4).Visible = True
        PreparationID = Grid2.Cell(FirstRow, 13).Text
        PreparationCode = Grid2.Cell(FirstRow, 2).Text
        RecipeCode = PreparationCode
        PreparationFileName = Grid2.Cell(FirstRow, 9).Text
        bNoPreparation = Grid2.Cell(FirstRow, 16).Text
        If PreparationFileName = "" Then PreparationFileName = Grid2.Cell(FirstRow, 9).Text
        PreparationQC = Trim(Grid2.Cell(FirstRow, 14).Text)
        frCommandInside(4).Visible = IIf(Len(PreparationQC) > 0, False, True)
        
        frCommandInside(19).Visible = IIf(frCommandInside(4).Visible, (bNoPreparation), False)
        frCommandInside(4).Visible = IIf(frCommandInside(19).Visible, False, frCommandInside(4).Visible)
 
        frCommandInside(5).Visible = True
        frCommandInside(3).Visible = True
        
        Call SetNoPreparationTableViewMixes(Grid2, bNoPreparation, PreparationID)
    Else
        PreparationID = 0
        
    End If
   
    
    

End Sub



Public Function SetNoPreparationTableViewMixes(ByRef Grid As Grid, ByVal bNoPreparation As Boolean, ByVal PreparationID As Long) As Boolean
Dim i As Integer
Dim t As Integer
Dim RecipeWeek As String
Dim PlannedPreparation As String
Dim DataRecipe As String
Dim RecipeCode As String

Dim splitString As Variant



Dim Mix1 As String
Dim Mix2 As String
Dim strMix As String

Dim vBackColor As OLE_COLOR


If bNoPreparation Then

    With dbTabPreparation
        .filter = ""
        .filter = "ID='" & PreparationID & "'"
        RecipeWeek = IIf(IsNull(Trim(!RecipeWeek)), "", Trim(!RecipeWeek))
        PlannedPreparation = IIf(IsNull(Trim(!PlannedPreparation)), "", Trim(!PlannedPreparation))
        DataRecipe = IIf(IsNull(Trim(!DataRecipe)), "", Trim(!DataRecipe))
        RecipeCode = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
        If RecipeCode <> "" Then
            With dbTabRecipe
                .filter = ""
                .filter = "Code='" & RecipeCode & "'"
                If .EOF Then
                Else
                
                    strMix = IIf(IsNull(Trim(!Mix)), "", Trim(!Mix))
                    
                    If strMix <> "" Then
                        splitString = Split(strMix, ";")
                            
                        If UBound(splitString) > 0 Then
                            Mix2 = splitString(1)
                        Else
                            Mix2 = ""
                        End If
                        Mix1 = splitString(0)
                    Else


                        If bNoPreparation Then
                                ' non ci sono i mix
                            If F_MsgBox.DoShow("Move Recipe to Production?", RecipeCode) Then
                                
                                lbCommandInside_Click 19
                            End If
                        End If
                        
                        Exit Function
    
                    End If
                  
                
                
                End If
            End With
        End If
    End With
    
End If

Dim CountMix As Integer
    CountMix = 0
    
    With Grid
        If Grid.Rows > 1 Then
        
            For i = 1 To .Rows - 1
                vBackColor = &HF0F0F0
                If bNoPreparation Then
                
                    If (Grid.Cell(i, 4).Text = RecipeWeek And Grid.Cell(i, 5).Text = PlannedPreparation And Grid.Cell(i, 6).Text = DataRecipe) Then
                            
                            If Mix1 = Grid.Cell(i, 2).Text Then
                                vBackColor = vbColorAzzurrino
                                CountMix = CountMix + 1
                            ElseIf Mix2 = Grid.Cell(i, 2).Text Then
                                vBackColor = vbColorAzzurrino
                                CountMix = CountMix + 1
                            End If
                    End If
                    
                    
                End If

                    For t = 1 To 7 ' perchè altrimenti mi "rovina" le altre colonne
                    
                        .Cell(i, t).BackColor = vBackColor
                    
                    Next
           
            Next
        End If
    End With
    
cont:
    
    If bNoPreparation Then
        
        If CountMix = UBound(splitString) + 1 Then
        
        
        Else
            ' non ci sono i mix
            If F_MsgBox.DoShow("One or more Mixes already prepared." & vbCrLf & "Move Recipe to Production?", RecipeCode) Then
                
                lbCommandInside_Click 19
                
            End If
            'PopupMessage 2, "One or more Mixes already prepared ( QC passed > moved to Production ) ", , , RecipeCode
            
        End If
    
    End If

End Function
            
                    
Private Sub Grid3_BeforeUserSort(ByVal Col As Long)
lRow = 0
End Sub

Private Sub Grid3_DblClick()
If lRow = 0 Then Exit Sub
If bClosedQc = False Then frCommandInside_Click 6
End Sub

Private Sub Grid3_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
lRow = FirstRow

frCommandInside(6).Visible = False
RevisionID = 0
PreparationQC = ""
PreparationCode = ""
bNoPreparation = False
Grid7.Rows = 1
frCommandInside(8).Visible = False
frCommandInside(26).Visible = False

'PicMenu(4).Visible = False
   
    If FirstRow > 0 Then
        PreparationID = Grid3.Cell(FirstRow, 12).Text
        PreparationCode = Grid3.Cell(FirstRow, 2).Text
        RecipeCode = PreparationCode
        PreparationFileName = Grid3.Cell(FirstRow, 9).Text
        If PreparationFileName = "" Then PreparationFileName = Grid3.Cell(FirstRow, 9).Text
        PreparationQC = Trim(Grid3.Cell(FirstRow, 7).Text)
        
        If bClosedQc = False Then
            frCommandInside(6).Visible = IIf(PreparationQC <> "Passed", True, False)
            frCommandInside(8).Visible = IIf(PreparationQC <> "Passed", False, True)
            frCloseQC.Visible = True ' IIf(PreparationQC <> "", True, False)
            
        Else
            
            frCommandInside(8).Visible = True
            
        End If
        
        frCommandInside(26).Visible = True
        
        Call GetQCPerRecipeInGrid7(Grid7, PreparationFileName)
    Else
        PreparationID = 0
        
    End If

End Sub

Private Sub Grid4_BeforeUserSort(ByVal Col As Long)
lRow = 0
End Sub

Private Sub Grid4_DblClick()
If lRow = 0 Then Exit Sub
    If ProductionFileName <> "" Then
        frmProduction.Left = Me.Left
        frmProduction.Top = Me.Top
        frmProduction.WindowState = Me.WindowState
        If frmProduction.DoShow(ProductionFileName, ProductionID) Then
        
            GetProductionInGrid
        
        End If
        frCommandInside(9).Visible = False
        Grid4.Cell(0, 1).SetFocus
        frCommandInside(11).Visible = False
        frCommandInside(23).Visible = False
        frCommandInside(9).Visible = False
    End If
    
End Sub

Private Sub Grid4_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
lRow = FirstRow
ProductionID = 0
ProductionFileName = ""
HannacodeProduction = ""
RecipeProduction = ""
RecipeCode = ""
'PicMenu(4).Visible = False
frCommandInside(11).Visible = False
frCommandInside(23).Visible = False
frCommandInside(9).Visible = False

If FirstRow > 0 Then
    
    
    ProductionFileName = Grid4.Cell(FirstRow, 10).Text
    ProductionID = Grid4.Cell(FirstRow, 11).Text
    HannacodeProduction = Trim(Grid4.Cell(FirstRow, 1).Text)
    RecipeProduction = Grid4.Cell(FirstRow, 4).Text
    RecipeCode = RecipeProduction
    frCommandInside(9).Visible = True
    frCommandInside(23).Visible = True
     frCommandInside(11).Visible = IsProductionStarted(ProductionID)
     
End If

End Sub

Private Sub Grid5_BeforeUserSort(ByVal Col As Long)
lRow = 0
End Sub

Private Sub Grid5_DblClick()
If lRow = 0 Then Exit Sub
ImMR_Click 1
End Sub

Private Sub Grid5_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

Dim bChecOut As Boolean
Dim bMRPrinted As Boolean
'PicMenu(4).Visible = False
lRow = FirstRow
    Call ClearMaterialRequisition
    IndexRecipeMaterialRequisition = 0
    lRowGridMaterialRequisition = 0
    bChecOut = False
    frCommandInside(15).Visible = False

    If FirstRow > 0 Then
    
        frCommandInside(15).Visible = True
        bChecOut = Grid5.Cell(FirstRow, 7).Text
        bMRPrinted = Grid5.Cell(FirstRow, 5).Text
        frCommandInside(14).Visible = Not (bChecOut) ' And bMRPrinted
        lRowGridMaterialRequisition = FirstRow
        IndexRecipeMaterialRequisition = IIf(Grid5.Cell(FirstRow, 10).Text = 0, 1, Grid5.Cell(FirstRow, 10).Text)
        Call AddRequisitionInTable(IndexRecipeMaterialRequisition)
    
    End If


End Sub

Private Sub Grid6_BeforeUserSort(ByVal Col As Long)
lRow = 0
End Sub

Private Sub Grid6_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
lRow = FirstRow
'PicMenu(4).Visible = False
End Sub

Private Sub Grid7_BeforeUserSort(ByVal Col As Long)
lRow = 0
End Sub

Private Sub Grid7_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
lRow = FirstRow
'PicMenu(4).Visible = False
End Sub



Private Sub Image_Click(Index As Integer)
frCommandInside_Click Index
End Sub

Private Sub Image1_Click()
txSearch = ""
frCommandInside_Click 18
End Sub

Private Sub Image2_Click()

txSearchCode = ""
Call GetProductionInGrid4(Grid4, False, txSearchCode, Frame5, cmbLineProduction)
frCommandInside(11).Visible = False
frCommandInside(23).Visible = False
frCommandInside(9).Visible = False
End Sub

Private Sub Image3_Click(Index As Integer)
PicMenu_Click Index
End Sub



Private Sub ImMR_Click(Index As Integer)
Grid5.ZOrder
Select Case Index
    Case 0
    
        frInside(4).Height = frMaterialRequisition.Height - frInside(4).Top
        Grid5.Height = frInside(4).Height - Grid5.Top - 240
     
    Case 1
    
        frInside(4).Height = frIRequisition(1).Top - frInside(4).Top
        Grid5.Height = frInside(4).Height - Grid5.Top - 60
    Case 2
        Grid3.Height = frInside(1).Height - cmbLineQC.Top - 60
    Case 3
        Grid3.ZOrder
        Grid3.Height = frInside(1).Height - 60

End Select
End Sub

Private Sub Lab_Click(Index As Integer)
If Index = 4 Then
    F_SETTING.Top = Me.Top
    F_SETTING.Left = Me.Left
    F_SETTING.WindowState = Me.WindowState
    F_SETTING.DoShow (1)
   ' FormIntro
End If
End Sub

Private Sub Label1_Click(Index As Integer)
If Index = 0 Then frCommandInside_Click 1
End Sub

Private Sub Label10_Click()
DisableImage_Click
End Sub

Private Sub Label14_Click()

    F_SETTING.DoShow (0)
    '
End Sub

Private Sub Label2_Click(Index As Integer)
PicMenu_Click Index
End Sub


Private Sub lbClassification_Click(Index As Integer)
frClassification_Click Index
End Sub

Private Sub lbClosedQC_Click()
frCloseQC_Click
End Sub

Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub
Private Sub lbCommandInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
frCommandInside_MouseMove Index, Button, Shift, X, Y
End Sub

Private Sub lbMR_Click(Index As Integer)
ImMR_Click Index
End Sub

Private Sub lbSearchRecipe_Click()

End Sub

Private Sub lbPreparation_Click()

    bPreparationDetails = Not (bPreparationDetails)
    If bPreparationDetails Then
       lbPreparation = "goto Recipe for prod. columns"
    Else
        lbPreparation = "goto Preparation columns"
    End If

    Call GetPreparationInGrid
End Sub

Private Sub llbExit_Click()
SetVisibleDatabaseFrame False

End Sub

Private Sub PBTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub PBTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub PBTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'FrmMove = False
End Sub

Private Sub PicIntro_Click()
TimeriNTRO.Enabled = False
FormIntro
End Sub

Private Sub PicMain_Click(Index As Integer)

    
    
   
    Select Case Index
        Case 0
            ' Recipe for production
            DisableImage_Click
    
    End Select
End Sub
Private Sub SetFrameDefault()




    ReceiptID = 0
    
    lRowGridMaterialRequisition = 0
    
    
    ' preparation
    frCommandInside(4).Visible = False
    frCommandInside(19).Visible = False
    PreparationID = 0
    PreparationCode = ""
    PreparationFileName = ""
    PreparationQC = ""
    bNoPreparation = False
    'QC
    
    frCommandInside(6).Visible = False
    Grid3.Rows = 1
    Grid7.Rows = 1
    
    ' database
    PicDatabase.Visible = False
    
    DoEvents

End Sub
Private Sub PicMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

            frCommandInside(IndexDashCommInside).BackColor = &H644603
            lbCommandInside(IndexDashCommInside).ForeColor = &HE0E0E0
            
            If IndexDashCommInside = 1 Or IndexDashCommInside = 5 Or IndexDashCommInside = 6 Or IndexDashCommInside = 14 Or IndexDashCommInside = 9 Then
                frCommandInside(IndexDashCommInside).BackColor = &H8000&
            End If
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


HannacodeProduction = ""
RecipeProduction = ""
ProductionFileName = ""
frCommandInside(26).Visible = False


Call SetFrameDefault


For i = 0 To PicMenu.UBound
    If i = Index Then
        PicMenu(i).BackColor = &H307030   '&H6D5B57
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


If PicIntro.Visible = False Then PicMain(Index).ZOrder
PicMain(Index).Visible = True

Select Case IndexProcedura
    Case 0
        USER_PATH = USER_PREPARATION_PATH
        Call SetLine(cmbLineRfP, True)
        rc = IIf(MyOperatore.IndexPrivilege > 0, True, False)
        Label2(6) = IIf(MyOperatore.Name <> "", MyOperatore.Name, Label2(6))
        frPrivilege.Visible = Not (rc)
        frFromulation.Visible = rc
        frFromulation.ZOrder
        'PicMenu(4).Visible = False
        RecipeCode = ""

    Case 1
        USER_PATH = USER_PREPARATION_PATH
        Call SetLine(cmbLine, True)
        
        frCloseFrame.Visible = False
        'PicMenu(4).Visible = False
        RecipeCode = ""
    Case 2
    
        
        Call SetLine(cmbLineQC, True)
       
        'Call GetQcInGrid(Grid3, cmbLineQC, , bClosedQc, Frame2)
        Grid7.Rows = 1
        
        frCloseFrame.Visible = False
        'PicMenu(4).Visible = False
        RecipeCode = ""
    Case 3
        USER_PATH = USER_PRODUCTION_PATH
        Call SetLine(cmbLineProduction, True)
        
       ' Call GetProductionInGrid4(Grid4, False, txSearchCode, Frame5, cmbLineProduction)
        frCloseFrame.Visible = False
        RecipeCode = ""
    Case 4
            ' rev history table
            If RecipeCode <> "" Then Call OpenRevisionHistory
            
        
        
        
End Select



SaveSetting App.Title, "Intro", "IndexProcedura", IndexProcedura
TimeriNTRO.Enabled = False
frSelezione.ZOrder
frSelezione.Visible = True



End Function



Private Sub PicMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'  Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

  Form_MouseMove Button, Shift, X, Y
 
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

Private Sub PicMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 '   Form_MouseUp Button, Shift, X, Y
End Sub


Private Sub SetPicForm()
'On Error GoTo ERR_SET:
Dim i As Integer
PicIntro.Left = 0
PicIntro.Top = 0 ' PBTitle.Height
PicIntro.Width = Me.Width
PicIntro.Height = Me.Height - PicIntro.Top

BottomLine.x1 = 0
BottomLine.x2 = Me.Width



PicMain(0).Move 0, PBTitle.Top + PBTitle.Height, Me.Width, BottomLine.y1 - (PBTitle.Top + PBTitle.Height)
PicDatabase.Move 0, PBTitle.Top + PBTitle.Height, Me.Width, BottomLine.y1 - (PBTitle.Top + PBTitle.Height)
'PicIntro.BackColor = &H929292

frFromulation.Move 0, 0, PicMain(0).Width ', lbWait.Top - 240


frMaterialRequisition.Move 0, 0, frFromulation.Width, frFromulation.Height
'PicInfo(3).BackColor = vbTimBlue






    
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
    

    If FileExists(App.Path & "\smartupdate.exe") Then
   
        ApriEseguibile App.Path & "\smartupdate.exe"
        SaveSetting App.Title, "Opzioni", "Avvisa Update", True
    Else
       ' MessageInfoTime = 2500
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
' frAvvio.ZOrder
'frAvvio.Move Me.Width / 2 - frAvvio.Width / 2, Me.Height / 2 - frAvvio.Height / 2
frAvvio.Visible = True
      
    ' controllo aggiornamenti
    DoEvents
    If GetSetting(App.Title, "Opzioni", "Avvisa Update", True) Then Call SmartUpdate
    
     
 
    If PrinterExist Then
       
    Else
        PopupMessage 2, "PDF Printer problem! Please check BioPdf setup files...", , True, "PDF Printer"
    End If
                
    
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

Private Sub Timer2_Timer()
PicIntro.Visible = False
Timer2.Enabled = False
End Sub

Private Sub TimeriNTRO_Timer()
FormIntro
TimeriNTRO.Enabled = False
End Sub

Private Sub FormIntro()

Dim rc As Boolean
Dim Grid(10) As Grid
    PicIntro.ZOrder
    PicIntro.Visible = True
    
    F_MAIN.BackColor = vbColorRed
    lbWait.ZOrder
    lbWait.Visible = True
    
    rc = CheckPrimoAvvio
    
    Set Grid(0) = Grid1
    Set Grid(1) = Grid2
    Set Grid(2) = Grid3
    Set Grid(3) = Grid7
    Set Grid(4) = Grid5
    Set Grid(5) = Grid6
    Set Grid(6) = Grid4
    
    Call SetAllMainGrid(Grid())
    '-----------------------------------
    ' revision history
    '-----------------------------------
    Call SetGridRecipeRevision(Grid8)
    
    '-----------------------------------
    
    Call SetColumnWidth
    Call GetReceiptFromDatabase(Grid1, bReceiptHistoryView, , cmbLineRfP)
    
    
        USER_PATH = IIf(bReceiptHistoryView, USER_DATA_PATH, USER_TEMP_PATH)
        
    
    
    
    If rc Then
        SelectProcedura (GetSetting(App.Title, "Intro", "IndexProcedura", 0))
    End If
       
  SetFrameEmptyView
  
  If GetSetting(App.Title, "Classification", "SetAllClassificationByRecipe", False) Then SetAllClassificationByRecipe
  

 lbWait.Visible = False
 F_MAIN.BackColor = &H4D3B37
 PicIntro.Visible = False
 PicMain(IndexProcedura).ZOrder
 
 
End Sub

Private Sub GetPreparationInGrid()

lbWaitPreparation.Visible = True
Call GetDataPreparationInGrid(Grid2, cmbLine, txSearch, bPreparationDetails)
If IndexProcedura = 1 Then blTable = cmbLine & " : Preparations"

lbWaitPreparation.Visible = False
SetFrameEmptyView

End Sub


Private Sub GetProductionInGrid()

lbWaitPreparation.Visible = True

Call GetProductionInGrid4(Grid4, False, txSearchCode, Frame5, cmbLineProduction)

If IndexProcedura = 3 Then blTable = cmbLineProduction & " : Production"

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
            IndexProcedura = 0
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


Private Sub DeleteRfPFunction()
Dim rc As Boolean

    If F_MsgBox.DoShow("DeleteRecipe?") = False Then Exit Sub
    
    rc = DeleteRecipeForProduction(ReceiptID, ReceiptFileName)
    If rc Then
        PopupMessage 2, "Recipe Correctly Deleted...", , , "Recipe For Production"
    Else
        PopupMessage 2, "Warning : NO Recipe Deleted...", , True, "Recipe For Production"
        
    End If

    Call GetReceiptFromDatabase(Grid1, bReceiptHistoryView, , cmbLineRfP)
    
    
End Sub


Private Sub ClearMaterialRequisition()
Dim i As Integer
Grid6.Rows = 1
For i = txDocument.LBound To txDocument.UBound
    txDocument(i) = ""
Next
End Sub



Private Sub AddRequisitionInTable(ByVal Index As Integer)
Dim i As Integer
Dim xDocument() As String

    Call ClearMaterialRequisition
    ReDim xDocument(txDocument.UBound)
    Call AddMaterialRequisitionFromFile(Grid6, Index, xDocument, ReceiptFileName)
    
    For i = txDocument.LBound To txDocument.UBound
        txDocument(i) = xDocument(i)
    Next
End Sub


Private Function CheckOutMaterialRequisition(ByVal Index As Integer) As Boolean
Dim rc As Boolean
Dim RecipeName As String
Dim bIfAllChecked As Boolean

CloseSettingDataFile

    CheckOutMaterialRequisition = CheckOutMaterialRequisitionInFile(Index, ReceiptFileName, RecipeName)
    
    
    If CheckOutMaterialRequisition Then
        ' attivo la preparation per questa ricetta...
        ' 1) copio i dati su TabPreparation
        ' 2) copio il file ReceiptFileName in \Preparation : nel nome del file deve comparire la ricetta!!
        Call SetTabPreparationPerRecipe(ReceiptFileName, RecipeName)
        Call SetRecipeGrid5(Grid5, ReceiptFileName)
        PopupMessage 2, "Material Requisition Check Out Done...", , , "Recipe : " & RecipeName
        
    End If
    
    bIfAllChecked = CheckOutAllRequisitionInFile(ReceiptFileName)
    CloseSettingDataFile
    
    If bIfAllChecked Then
        ' chiudo Recipe for production
        ' 1) sposto il file ReceiptFileName in \Data
        ' 2) aggiorno Main > sparisce il record
        If F_MsgBox.DoShow("No Material requisition left for this Recipe for production." & vbCrLf & "Close Recipe for production?", "Material Requisition finished") Then
            
            Call SetCloseRecipeForProduction
            
           
        End If
    End If
    
        
    CloseSettingDataFile
    
    Call GetPreparationInGrid
    
    CloseSettingDataFile

End Function

Private Function SetCloseRecipeForProduction()

Dim Path As String

    Path = USER_TEMP_PATH
    
    If bReceiptHistoryView Then Path = USER_DATA_PATH
    
    Call CloseRecipeForProduction(ReceiptFileName, Path)
    
    PopupMessage 2, "Recipe For Production Closed...", , , "Marerial Req. Check Out"
    frCommandInside_Click 16
End Function


Private Function ReceiptHistory()
Dim strOpen As String
Dim strClosed As String

    strOpen = "Recipe Check Out : After Warehouse Approval, Line Leader must Check Out before Preparation"
    strClosed = "Closed Recipe : After Line Leader Check Out, a Recipe is Closed and already in Preparation "
    
    bReceiptHistoryView = Not (bReceiptHistoryView)
    'If bReceiptHistoryView Then
    '    bReceiptHistoryView = False
    'Else
    '    bReceiptHistoryView = True

    'End If
    
    lbRecipeForProductionInfo = IIf(bReceiptHistoryView, strClosed, strOpen)
            
    USER_PATH = IIf(bReceiptHistoryView, USER_DATA_PATH, USER_TEMP_PATH)
    
    
    Call GetReceiptFromDatabase(Grid1, bReceiptHistoryView, , cmbLineRfP)
    

    Grid1.ForeColorFixed = IIf(Not (bReceiptHistoryView), &H745613, &H40C0&)
   
    
    blTable.Caption = IIf(Not (bReceiptHistoryView), "Active Recipe For Prod.", "Closed Recipe For Prod.")
  
    lbInside(2).ForeColor = IIf(Not (bReceiptHistoryView), &H745613, &H40C0&)
    frReceiptGrid.BackColor = IIf(Not (bReceiptHistoryView), &H745613, &H40C0&)
    lbInside(2).Caption = IIf(Not (bReceiptHistoryView), "Active Recipe", "Closed Recipe")
    lbCommandInside(13).Caption = IIf(bReceiptHistoryView, "Active Recipe Table", "Closed Recipe Table")
    
    frCommandInside(1).Visible = Not (bReceiptHistoryView)
    
    

    SetFrameEmptyView
    
    PopupMessage 2, IIf(bReceiptHistoryView, "Closed", "Acrive") & " Recipe for Production Table"

End Function

Private Function DeletePreparation()
Dim RecipeName As String
Dim strPreparation As String
Dim ID_PREP As Long
Dim RfpFileName As String
Dim rc As Boolean

If PreparationID > 0 Then
    ID_PREP = PreparationID
    With dbTabPreparation
        .filter = ""
        .filter = "ID='" & PreparationID & "'"
        If .EOF Then
        Else
            RecipeName = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
            RfpFileName = IIf(IsNull(Trim(!RfpFileName)), "", Trim(!RfpFileName))
            
            strPreparation = "Date Recipe : " & IIf(IsNull(Trim(!DataRecipe)), "", Trim(!DataRecipe))
            strPreparation = strPreparation & "  |  Recipe Week  : " & IIf(IsNull(Trim(!RecipeWeek)), "", Trim(!RecipeWeek))
            strPreparation = strPreparation & "  |  Planned Prep.  : " & IIf(IsNull(Trim(!PlannedPreparation)), "", Trim(!PlannedPreparation))
            strPreparation = strPreparation & vbCrLf & "Line  : " & IIf(IsNull(Trim(!Line)), "", Trim(!Line))
            strPreparation = strPreparation & vbCrLf & "Note  : " & IIf(IsNull(Trim(!Note)), "", Trim(!Note))
            
            If F_MsgBox.DoShow("Delete Preparation?" & vbCrLf & strPreparation, RecipeName) Then
                
                .Delete
                .Update
                
                  
                
                Dim i As Integer
                With dbTabAcquisition
                    .filter = ""
                    .filter = "PreparationID='" & ID_PREP & "'"
                    If .EOF Then
                    Else
                        .MoveFirst
                        For i = 1 To .RecordCount
                            .Delete
                            .MoveNext
                        Next
            
                    End If
                End With
                
                .Close
                .Open "SELECT *  FROM TabPreparation order by id ", dbChemicalProduction, adOpenKeyset, adLockOptimistic, adCmdText
                
                DoEvents
                
                If PreparationFileName <> "" Then
                        
                    If FileExists(USER_PREPARATION_PATH & PreparationFileName) Then
                        Kill USER_PREPARATION_PATH & PreparationFileName
                    ElseIf FileExists(USER_PREPARATION_PATH & "data\" & PreparationFileName) Then
                        Kill USER_PREPARATION_PATH & "data\" & PreparationFileName
                    End If
                    
                End If
                
                PopupMessage 2, "Preparation Correctly Deleted!", , , RecipeName
                
                If F_MsgBox.DoShow("Preparation Deleted." & vbCrLf & "Delete Recipe for Production ?", "Recipe for Production") Then
                    If RfpFileName <> "" Then
                        If FileExists(USER_TEMP_PATH & RfpFileName) Then
                            Kill USER_TEMP_PATH & RfpFileName
                        ElseIf FileExists(USER_DATA_PATH & RfpFileName) Then
                            Kill USER_DATA_PATH & RfpFileName
                        End If
                    End If
                    
                    
                    rc = DeleteRecipeForProduction(0, RfpFileName)
                
                End If
                
                Call GetPreparationInGrid
            End If
        End If
    
  End With

 
End If
End Function


Private Sub txSearch_Change()
Dim rc As Boolean
rc = False


If txSearch = TX_INTRO Then Exit Sub

If txSearch = "" Then
    txSearch = TX_INTRO
    rc = True
End If


SearcInTable (rc)
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

Private Sub txSearchCode_Change()

Dim rc As Boolean
rc = False


If txSearchCode = TX_INTRO Then Exit Sub

If txSearchCode = "" Then
    txSearchCode = TX_INTRO
    rc = True
End If


SearcInTable (rc)


End Sub

Private Sub txSearchCode_Click()
If txSearchCode = TX_INTRO Then txSearchCode = " "
End Sub

Private Sub txSearchCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then frCommandInside_Click 20
End Sub

Private Sub txSearchCode_LostFocus()
If Trim(txSearchCode) = "" Then txSearchCode = TX_INTRO
End Sub




Private Sub SearcInTable(ByVal rc As Boolean)

Select Case IndexProcedura
    Case 0
       ' Call SearchInGrid(Grid1, txSearch, rc)
    Case 1
        Call SearchInGrid(Grid2, txSearch, rc, 2)
    Case 3
        Call SearchInGrid(Grid4, txSearchCode, rc)
End Select


End Sub


Private Function PassToQC()

Dim RecipeQCType As QCType
Dim QtyProduced As String
RecipeQCType = RecipeQCClean


Dim rc As Boolean

If PreparationID > 0 Then

    With dbTabPreparation
        .filter = ""
        .filter = "ID='" & PreparationID & "'"
        If .EOF Then
        
        Else
            QtyProduced = IIf(IsNull(Trim(!QtyProduced)), "0", Trim(!QtyProduced))
        End If
    End With
    
    With RecipeQCType
        .Date = Now()
        .Note = "Pass To QC"
        .Operator = MyOperatore.Name
        .RecipeCode = PreparationCode
        .ID = PreparationID
        .Status = "Waiting"
        .SettingName = PreparationFileName
    End With
    
    
    If F_MsgBox.DoShow("Pass to QC?" & vbCrLf & "Qty. Produced = " & QtyProduced & " kg", PreparationCode) Then
    
        rc = PassToQCInfile(RecipeQCType)
        
        If rc Then
        
            Call GetPreparationInGrid
         
         
         
            PopupMessage 2, "Recipe passed to QC....", , , PreparationCode
        Else
            PopupMessage 2, "Warning : Data Problem ....", , True, PreparationCode
        End If
        
        
        RecipeCode = PreparationCode
       ' PreparationFileName, PreparationID, RecipeCode
      '  PreparationID = Grid2.Cell(FirstRow, 13).Text
       ' PreparationCode = Grid2.Cell(FirstRow, 2).Text
       ' RecipeCode = PreparationCode
       ' PreparationFileName = Grid2.Cell(FirstRow, 9).Text
        
        '-----------------------
        ' creo i QR Code!
        '-----------------------
        'frCommandInside_Click 26
        
        
    End If
    
End If



End Function



Private Function AddQcToRecipe(Optional bValue As Boolean)
Dim strPrepQC As String
Dim strFileName As String
Dim strCode As String
Dim strID As Long
Dim bIsMix As Boolean

' bValue = true se arrivo da QRCode Label!!!!
If bValue Then GoTo isPassed:
    If FormQC.DoShow(PreparationCode, PreparationFileName, PreparationID, PreparationQC) Then
    
        
        strPrepQC = PreparationQC
        strFileName = PreparationFileName
        strCode = PreparationCode
        strID = PreparationID
        
        bIsMix = IfRecipeIsMixString(PreparationCode)
        
        Call GetQcInGrid(Grid3, cmbLineQC, , bClosedQc, Frame2)
        Call GetQCPerRecipeInGrid7(Grid7, PreparationFileName)
        
        frCommandInside(6).Visible = IIf(strPrepQC <> "Passed", True, False)
        frCommandInside(8).Visible = IIf(strPrepQC <> "Passed", False, True)
        frCloseQC.Visible = IIf(strPrepQC <> "", True, False)
        
        
        PreparationQC = strPrepQC
        PreparationFileName = strFileName
        PreparationCode = strCode
        PreparationID = strID
        
isPassed:
        
        If strPrepQC = "Passed" Then
        
        
        
            If bIsMix Then
            
            
                MoveRecipeInProduction True, True
                
                
            Else
            
                If F_MsgBox.DoShow("Move preparation to Production?", PreparationCode) Then
                    
                    
                    
                    
                    MoveRecipeInProduction True, False
                    
                End If
            End If
        
        End If
    End If



End Function

Private Function MoveRecipeInProduction(ByVal bSalta As Boolean, ByVal bIsMix As Boolean) As Boolean
Dim rc As Boolean
Dim strLot As String
Dim rfpSettingName As String
If bSalta Then GoTo cont:

If CheckPrivilege(1) Then

    If F_MsgBox.DoShow("Move " & PreparationCode & " to Production?", "Move to Production") Then
cont:
 
        With dbTabPreparation
            .filter = "Recipe='" & PreparationCode & "' and FileName='" & PreparationFileName & "'"
        If .EOF Then
        Else
             rfpSettingName = IIf(IsNull(Trim(!RfpFileName)), "", Trim(!RfpFileName))
             
            If IsNull(!Lot) Then
                !Lot = Trim(strLot)
                .Update
            End If
        End If
        
    End With
    
    Call ExportPreparationAfterQC(PreparationFileName, PreparationCode, strLot, rfpSettingName)
       
        rc = MoveRecipeInProductionDatabase(PreparationFileName, PreparationID, bIsMix)
    
        If rc Then
        
            'Call ExportPreparationAfterQC(PreparationFileName, PreparationCode)
            Call GetQcInGrid(Grid3, cmbLineQC, , bClosedQc, Frame2)
            Call GetQCPerRecipeInGrid7(Grid7, PreparationFileName)
            Call GetProductionInGrid4(Grid4, False, txSearchCode, Frame5, cmbLineProduction)
            Call GetPreparationInGrid
            frCommandInside(6).Visible = False
            frCommandInside(8).Visible = False
            frCloseQC.Visible = False
            If Not (bIsMix) Then PopupMessage 2, "Recipe in Production....", , , PreparationCode
        End If
    End If
Else

    
End If

End Function



Private Function ClosedQCTable() As Boolean

bClosedQc = Not (bClosedQc)
If bClosedQc Then
    lbCommandInside(7) = "Open QC Table"
Else
    lbCommandInside(7) = "Closed QC Table"
End If


    lbInside(1).ForeColor = IIf(Not (bClosedQc), &H745613, &H40C0&)
    lbInside(5).ForeColor = IIf(Not (bClosedQc), &H745613, &H40C0&)
    frReceiptGrid.BackColor = IIf(Not (bClosedQc), &H745613, &H40C0&)
    lbInside(1).Caption = IIf(Not (bClosedQc), "Open Recipes in QC", "Closed QC")
    blTable.Caption = IIf(bClosedQc, "Closed QC", "Open QC")
    
    
    

Call GetQcInGrid(Grid3, cmbLineQC, , bClosedQc, Frame2)

Grid3.ForeColorFixed = IIf(Not (bClosedQc), &H745613, &H40C0&)
Grid7.ForeColorFixed = IIf(Not (bClosedQc), &H745613, &H40C0&)
Grid7.Rows = 1

frCommandInside(6).Visible = False
frCommandInside(8).Visible = False
frCloseQC.Visible = False

Dim sString As String
sString = IIf(bClosedQc, "Closed", "Open")
PopupMessage 2, sString & " Preparation's QC Table...", , , sString & " QC"

            
            

End Function

Private Function ProductionCheckOut()
Dim rc As Boolean

    'ProductionFileName = Grid4.Cell(FirstRow, 10).Text
    'productionID = Grid4.Cell(FirstRow, 11).Text
    
    If CheckPrivilege(1) Then
    
        If F_MsgBox.DoShow("Production Check Out ? ") Then
        
            rc = SetProductionCheckOut(ProductionFileName, ProductionID)
            
            
            If F_MsgBox.DoShow("Print Final QC Production QRCode?", HannacodeProduction) Then
                Call PrintLabelProduction(ProductionFileName, ProductionID, HannacodeProduction, True)
            End If

            If rc Then
            
                Call GetProductionInGrid4(Grid4, False, txSearchCode, Frame5, cmbLineProduction)
                frCommandInside(11).Visible = False
                frCommandInside(23).Visible = False
                frCommandInside(9).Visible = False
    
                PopupMessage 2, "Production Closed!"
            End If
        
        End If
    
    End If
    

End Function
Private Function ProductionDelete()
Dim rc As Boolean
    
      
    If F_MsgBox.DoShow("Delete Production?") Then
        rc = SetProductionDelete(ProductionFileName, ProductionID)
        If rc Then
            Call GetProductionInGrid4(Grid4, False, txSearchCode, Frame5, cmbLineProduction)
            frCommandInside(11).Visible = False
            frCommandInside(23).Visible = False
            frCommandInside(9).Visible = False
            PopupMessage 2, "Production deleted..."
        End If
    End If
    

End Function


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



Private Sub SetProductionDatabase()

       ' SetVisibleDatabaseFrame False

        DoEvents
        DoEvents
        
        FormProductionDatabaseHistory.Top = Me.Top
        FormProductionDatabaseHistory.Left = Me.Left
        FormProductionDatabaseHistory.WindowState = Me.WindowState
        FormProductionDatabaseHistory.DoShow
        
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

Private Sub SetPicMenu()

  DefaultMenu(2).Visible = True
  Label2(7).Visible = True

End Sub

Private Sub SetChemicalRMDatabase()
 
        FormChemicalRMDatabase.Top = Me.Top
        FormChemicalRMDatabase.Left = Me.Left
        FormChemicalRMDatabase.Width = Me.Width
        FormChemicalRMDatabase.Height = Me.Height
        FormChemicalRMDatabase.WindowState = Me.WindowState
        FormChemicalRMDatabase.DoShow
        
        
End Sub







'-----------------------------------------
'
'
'           Revision History
'
'
'-----------------------------------------






Private Sub lbExcel_Click()
frExcel_Click
End Sub


Private Sub frExcel_Click()

    Grid2.ExportToExcel USER_DESKTOP & "\" & FormatNomeFile(RecipeCode) & "_RevHistory.xls", True, True
    MessageInfoTime = 2500
    PopupMessage 2, "File correcly created on Desktop", , , FormatNomeFile(RecipeCode) & "_RevHistory.xls"
End Sub


Private Sub lbInside_Click(Index As Integer)
    Select Case Index
        
        Case 5
            ' rev history table
            If RecipeCode <> "" Then Call OpenRevisionHistory
    
    End Select
End Sub



Private Sub OpenRevisionHistory()
Dim Description As String

RecipeCode = Trim(RecipeCode)
Call GetRecipeRevision(Grid8, RecipeCode, Description)

If Description <> "" Then
    RecipeCode = UCase(RecipeCode)
IconFH.Visible = True
lbInside(8) = RecipeCode & " " & Description
Else
    IconFH.Visible = False
    lbInside(8) = "Formulation History"
End If



frExcel.Visible = IIf(Grid8.Rows > 1, True, False)
Frame6.Visible = IIf(Grid8.Rows > 1, False, True)



End Sub




Private Sub txSearchRecipe_Change()


' rev history table

If txSearchRecipe = TX_INTRO Then Exit Sub

If txSearchRecipe = "" Then
    txSearchRecipe = TX_INTRO
  
End If

RecipeCode = Trim(txSearchRecipe)
Call OpenRevisionHistory
            
End Sub


Private Sub txSearchRecipe_Click()
If Trim(txSearchRecipe) = TX_INTRO Then txSearchRecipe = " "
End Sub

Private Sub txSearchRecipe_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then frCommandInside_Click 26
End Sub

Private Sub txSearchRecipe_LostFocus()
If Trim(txSearchRecipe) = "" Then txSearchRecipe = TX_INTRO
End Sub
Private Sub Image4_Click()
txSearchRecipe = ""
End Sub









Private Function ScanQRCodeQC() As Boolean

MessageInfoTime = 2000

txQRCode = ""
txQRCode.Top = 0
txQRCode.Visible = True
txQRCode.SetFocus

UserQrCode = iQRCodeTypeClean
iQRCodeType = iQRCodeTypeClean

PopupMessage 2, "Scan QRCode....", , , "Preparation | QC"


'CP-D23:HI96780V-0:100:9999:04/2050:vianello:2025/04/26:07.43:Waiting
'CP-D23:HI96780V-0:100:9999:04/2050:vianello:2025/04/26:07.43:Waiting
'CP-D23:HI93754C-0:100:9998:04/2050:vianello:2025/04/26:07.43:Waiting
'CP-D23:HI93754C-0:100:9998:04/2050:vianello:2025/04/26:07.43:Waiting
'CP-D23:HI94754C-0:1000:9997:04/2050:vianello:2025/04/26:07.43:Waiting
'CP-D23:HI94754C-0:1000:9997:04/2050:vianello:2025/04/26:07.43:Waiting

'txQRCode = "CP-D23:HI96780V-0:100:9999:04/2050:vianello:2025/04/26:07.43:Waiting"
'txQRCode_KeyPress
End Function



Private Sub txQRCode_KeyPress(KeyAscii As Integer)
Dim rc As Boolean

On Error GoTo ERR_QR:



If KeyAscii = 13 Then

    
    DoEvents
    If txQRCode = "" Then Exit Sub
   
    If GetQRCodeFromString(Trim(txQRCode), UserQrCode) Then
    
        iQRCodeType = UserQrCode
        If QRCodeQCToTabPreparation(UserQrCode, False) Then
            
            ' ho Giò il QC in tabella....
            'If SearchQCInTab(UserQrCode, Grid3) Then
            
                PreparationQC = UserQrCode.QC
                Call AddQcToRecipe
            '
            'End If
            
        Else
            ' apro QC nuovo ?
            MessageInfoTime = 3000
            With UserQrCode
                PopupMessage 2, "Recipe, Code or Lot not in QC Table" & vbCrLf & "Please Check QRCode or Scan Again...", , True, .Recipe & " | " & .Code & " | " & .Lot
            End With
        End If

        
       
       
    Else
        PopupMessage 2, "Please Check QRCode or Scan Again...", , True, "QRCode"
    End If
    

End If

ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_QR:
    MessageInfoTime = 2000
    PopupMessage 2, err.Description & vbCrLf & "Please repeat reading...", , , "QR Code Reader"
    Resume ERR_END:


End Sub

