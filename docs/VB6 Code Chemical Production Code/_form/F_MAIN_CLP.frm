VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form F_MAIN_CLP 
   BackColor       =   &H004D3B37&
   Caption         =   "Chemical CLP Classification Software for Production"
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
   Icon            =   "F_MAIN_CLP.frx":0000
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
      Height          =   9900
      Index           =   1
      Left            =   2880
      ScaleHeight     =   9900
      ScaleWidth      =   19200
      TabIndex        =   34
      Top             =   1080
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
         Height          =   1095
         Index           =   1
         Left            =   960
         TabIndex        =   42
         Top             =   8640
         Width           =   17535
         Begin VB.Label lbClassification 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recipe Classification"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   26.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   630
            Index           =   1
            Left            =   6600
            TabIndex        =   43
            Top             =   240
            Width           =   4365
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
         Height          =   7815
         Index           =   1
         Left            =   840
         TabIndex        =   35
         Top             =   360
         Width           =   17655
         Begin VB.Frame Frame1 
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
            TabIndex        =   38
            Top             =   2400
            Visible         =   0   'False
            Width           =   8535
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
               Left            =   30
               TabIndex        =   40
               Top             =   720
               Width           =   8415
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Recipe Table"
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
               Index           =   0
               Left            =   3300
               TabIndex        =   39
               Top             =   360
               Width           =   1905
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
            TabIndex        =   36
            Top             =   0
            Width           =   17655
            Begin VB.Label lbInside 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Recipe Table"
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
               Index           =   1
               Left            =   120
               TabIndex        =   37
               Top             =   120
               Width           =   1515
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00B0B0B0&
               X1              =   120
               X2              =   17640
               Y1              =   480
               Y2              =   480
            End
         End
         Begin FlexCell.Grid Grid2 
            Height          =   7215
            Left            =   120
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   600
            Width           =   17535
            _ExtentX        =   30930
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
      End
      Begin VB.Label Label3 
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
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   960
         Width           =   60
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   9735
      Index           =   3
      Left            =   1680
      ScaleHeight     =   9735
      ScaleWidth      =   18135
      TabIndex        =   62
      Top             =   1920
      Width           =   18135
      Begin VB.Frame frInside 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "16800"
         Height          =   7935
         Index           =   6
         Left            =   1320
         TabIndex        =   67
         Top             =   1320
         Width           =   17055
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
            TabIndex        =   74
            Top             =   6960
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
               TabIndex        =   75
               Top             =   120
               Width           =   3015
            End
            Begin VB.Image Image 
               Height          =   480
               Index           =   1
               Left            =   120
               MousePointer    =   99  'Custom
               OLEDropMode     =   1  'Manual
               Picture         =   "F_MAIN_CLP.frx":0A02
               Top             =   0
               Width           =   480
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
            TabIndex        =   71
            Top             =   0
            Width           =   17055
            Begin VB.Line Line16 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   16920
               Y1              =   480
               Y2              =   480
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
               TabIndex        =   73
               Top             =   0
               Width           =   3300
            End
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
               TabIndex        =   72
               Top             =   120
               Width           =   1245
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00886010&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   5880
            TabIndex        =   68
            Top             =   2760
            Width           =   5055
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
               TabIndex        =   70
               Top             =   720
               Width           =   2730
            End
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
               TabIndex        =   69
               Top             =   360
               Width           =   1155
            End
         End
         Begin FlexCell.Grid Grid8 
            Height          =   6015
            Left            =   0
            TabIndex        =   76
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
         Begin VB.Line Line17 
            BorderColor     =   &H00D0D0D0&
            X1              =   0
            X2              =   16920
            Y1              =   6720
            Y2              =   6720
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
            TabIndex        =   77
            Top             =   7560
            Width           =   5745
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   6000
         TabIndex        =   63
         Top             =   360
         Width           =   8175
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
            TabIndex        =   66
            Text            =   "Search"
            Top             =   160
            Width           =   3495
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
            Left            =   4080
            TabIndex        =   64
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
               Index           =   1
               Left            =   0
               TabIndex        =   65
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.Image Image4 
            Height          =   360
            Left            =   7440
            Picture         =   "F_MAIN_CLP.frx":3DE4
            Top             =   160
            Width           =   360
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00B0B0B0&
            Height          =   735
            Left            =   0
            Top             =   0
            Width           =   8175
         End
      End
      Begin VB.Image IconFH 
         Height          =   480
         Left            =   720
         OLEDropMode     =   1  'Manual
         Picture         =   "F_MAIN_CLP.frx":5856
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
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
      Height          =   9900
      Index           =   2
      Left            =   4320
      ScaleHeight     =   9900
      ScaleWidth      =   19200
      TabIndex        =   45
      Top             =   1080
      Visible         =   0   'False
      Width           =   19200
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
         Left            =   840
         TabIndex        =   48
         Top             =   360
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
            Index           =   2
            Left            =   0
            TabIndex        =   52
            Top             =   0
            Width           =   17655
            Begin VB.Line Line4 
               BorderColor     =   &H00B0B0B0&
               X1              =   120
               X2              =   17640
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label lbInside 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hanna Code Table"
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
               Left            =   120
               TabIndex        =   53
               Top             =   120
               Width           =   2205
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
            Height          =   1455
            Left            =   4440
            TabIndex        =   49
            Top             =   2400
            Visible         =   0   'False
            Width           =   8535
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hanna Code Table"
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
               Index           =   4
               Left            =   2880
               TabIndex        =   51
               Top             =   360
               Width           =   2745
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
               Left            =   30
               TabIndex        =   50
               Top             =   720
               Width           =   8415
            End
         End
         Begin FlexCell.Grid Grid3 
            Height          =   7215
            Left            =   120
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   600
            Width           =   17535
            _ExtentX        =   30930
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
         Height          =   1095
         Index           =   2
         Left            =   960
         TabIndex        =   46
         Top             =   8640
         Width           =   17535
         Begin VB.Label lbClassification 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hanna Code Classification"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   26.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   630
            Index           =   2
            Left            =   6000
            TabIndex        =   47
            Top             =   240
            Width           =   5535
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
         MousePointer    =   99  'Custom
         TabIndex        =   55
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
      Height          =   9900
      Index           =   0
      Left            =   360
      ScaleHeight     =   9900
      ScaleWidth      =   19200
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   19200
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
         Index           =   0
         Left            =   840
         TabIndex        =   10
         Top             =   360
         Width           =   17655
         Begin FlexCell.Grid Grid1 
            Height          =   7215
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   600
            Width           =   17535
            _ExtentX        =   30930
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
            TabIndex        =   13
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
               Caption         =   "Chemical RM Table"
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
               TabIndex        =   14
               Top             =   120
               Width           =   2265
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
            TabIndex        =   11
            Top             =   2400
            Visible         =   0   'False
            Width           =   8535
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Chemical MR"
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
               Left            =   3270
               TabIndex        =   16
               Top             =   360
               Width           =   1965
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
               TabIndex        =   12
               Top             =   720
               Width           =   8415
            End
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
         Height          =   1095
         Index           =   0
         Left            =   960
         TabIndex        =   21
         Top             =   8640
         Width           =   17535
         Begin VB.Label lbClassification 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chemical RM Classification"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   26.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   630
            Index           =   0
            Left            =   5880
            TabIndex        =   22
            Top             =   240
            Width           =   5745
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
         Index           =   0
         Left            =   15600
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   960
         Width           =   60
      End
   End
   Begin VB.Frame frCloseFrame 
      BackColor       =   &H004D3B37&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   8640
      TabIndex        =   23
      Top             =   10920
      Visible         =   0   'False
      Width           =   1935
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
         MouseIcon       =   "F_MAIN_CLP.frx":8C38
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   795
         Width           =   975
      End
      Begin VB.Image DefaultMenu 
         DragIcon        =   "F_MAIN_CLP.frx":8F42
         Height          =   480
         Index           =   1
         Left            =   720
         MouseIcon       =   "F_MAIN_CLP.frx":C324
         MousePointer    =   99  'Custom
         Picture         =   "F_MAIN_CLP.frx":C62E
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame frSelezione 
      BackColor       =   &H00307030&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   100
      Left            =   0
      TabIndex        =   20
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
   Begin VB.Timer TimeriNTRO 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1080
      Top             =   6960
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox PBTitle 
      BackColor       =   &H004D3B37&
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
         Index           =   3
         Left            =   5760
         MouseIcon       =   "F_MAIN_CLP.frx":FA10
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   60
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   720
            MouseIcon       =   "F_MAIN_CLP.frx":FD1A
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN_CLP.frx":10024
            Top             =   120
            Width           =   480
         End
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
            Index           =   3
            Left            =   0
            MouseIcon       =   "F_MAIN_CLP.frx":13406
            MousePointer    =   99  'Custom
            TabIndex        =   61
            Top             =   640
            Width           =   1890
         End
      End
      Begin VB.Frame frSearchRecipe 
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   8760
         TabIndex        =   56
         Top             =   120
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
            TabIndex        =   59
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
            Index           =   0
            Left            =   4080
            TabIndex        =   57
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
               Index           =   0
               Left            =   0
               TabIndex        =   58
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.Image Image1 
            Height          =   360
            Left            =   7440
            Picture         =   "F_MAIN_CLP.frx":13710
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
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   6
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN_CLP.frx":15182
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hanna Code"
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
            MousePointer    =   99  'Custom
            TabIndex        =   7
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
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   4
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN_CLP.frx":17B74
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recipe"
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
            Left            =   675
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   645
            Width           =   540
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
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   2
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chemical RM"
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
            Left            =   465
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   645
            Width           =   1080
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN_CLP.frx":1A566
            Top             =   120
            Width           =   480
         End
      End
   End
   Begin VB.Frame frAvvio 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   975
      Left            =   0
      TabIndex        =   17
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   120
         Width           =   19215
      End
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_MAIN_CLP.frx":1CF58
      Height          =   480
      Index           =   0
      Left            =   18120
      MouseIcon       =   "F_MAIN_CLP.frx":2033A
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN_CLP.frx":20644
      Top             =   11160
      Width           =   480
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
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   4830
      X2              =   4830
      Y1              =   0
      Y2              =   12383.01
   End
   Begin VB.Image DefaultMenu 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   360
      MouseIcon       =   "F_MAIN_CLP.frx":23A26
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN_CLP.frx":23D30
      Top             =   11160
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   4560
      MouseIcon       =   "F_MAIN_CLP.frx":27112
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN_CLP.frx":2741C
      Top             =   11160
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   14160
      MouseIcon       =   "F_MAIN_CLP.frx":2A7FE
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN_CLP.frx":2AB08
      Top             =   11160
      Visible         =   0   'False
      Width           =   480
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
      MouseIcon       =   "F_MAIN_CLP.frx":2DEEA
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   11715
      Width           =   645
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
      MouseIcon       =   "F_MAIN_CLP.frx":2E1F4
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   11715
      Width           =   1050
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
      MouseIcon       =   "F_MAIN_CLP.frx":2E4FE
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   11715
      Visible         =   0   'False
      Width           =   870
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
      MouseIcon       =   "F_MAIN_CLP.frx":2E808
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   11715
      Width           =   1110
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
      MouseIcon       =   "F_MAIN_CLP.frx":2EB12
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   10680
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
      Height          =   1455
      Index           =   3
      Left            =   3840
      MouseIcon       =   "F_MAIN_CLP.frx":2EE1C
      MousePointer    =   99  'Custom
      TabIndex        =   28
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
      Height          =   1815
      Index           =   2
      Left            =   13080
      MouseIcon       =   "F_MAIN_CLP.frx":2F126
      MousePointer    =   99  'Custom
      TabIndex        =   27
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
      Height          =   1215
      Index           =   0
      Left            =   17640
      MouseIcon       =   "F_MAIN_CLP.frx":2F430
      MousePointer    =   99  'Custom
      TabIndex        =   26
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
      Height          =   1455
      Index           =   1
      Left            =   8280
      MouseIcon       =   "F_MAIN_CLP.frx":2F73A
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   10560
      Width           =   2655
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
End
Attribute VB_Name = "F_MAIN_CLP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private IndexProcedura As Integer

Private TimerCount As Integer
Private IndexDashCommInside As Integer
Private m_rc As Boolean
Private TX_INTRO As String

Private CodeChemical As String
Private CodeRecipe As String
Private CodeHannaCode As String

Private ChemicalID As Long
Private RecipeID As Long
Private HannaCodeID As Long
Private RecipeCode As String


Private Sub DefaultMenu_Click(Index As Integer)
DefaultMenuLabel_Click Index
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

Select Case Index
    Case 0
            If F_MsgBox.DoShow("Quit Chemical CLP Classification?", "Exit") Then

                CloseSettingDataFile
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
    
        

    Case 3
       frmLogin.DoShow
       Label2(6) = IIf(MyOperatore.Name <> "", MyOperatore.Name, Label2(6))
    Case 4
        F_SETTING.Left = Me.Left
        F_SETTING.Top = Me.Top
        F_SETTING.WindowState = Me.WindowState
        F_SETTING.DoShow , , , DefaultMenu(4)
       ' FormIntro
    Case 5
        
        
    Case 6
        
    Case 7
        
    Case 8
      

    Case 9

        Exit Sub
        
End Select
'frmSTDToleranceInfo.Visible = False
End Sub



Private Sub frClassification_Click(Index As Integer)
     Dim MyID As Long
     
     
     
 'CodeChemical
 'CodeRecipe
 'CodeHannaCode

 'ChemicalID
 'RecipeID
 'HannaCodeID
 
   Select Case Index
        Case 0

            If CodeChemical = "" Then
                Call FormChemicalRM.DoShow(, , , 1)
            Else
                Call F_PICTOGRAM.DoShow(ChemicalID, 1, CodeChemical, True)
            End If
     
        Case 1
            If CodeRecipe = "" Then
                Call FormRecipes.DoShow
            Else
                Call F_PICTOGRAM.DoShow(RecipeID, 2, CodeRecipe, True)
            End If

        Case 2
        
             If CodeHannaCode = "" Then
                Call FormCodes.DoShow
            Else
                Call F_PICTOGRAM.DoShow(HannaCodeID, 0, CodeHannaCode, True)
            End If
       
        

    End Select
End Sub



Private Sub Form_Activate()
CloseSettingDataFile
    Label2(6) = IIf(MyOperatore.Name <> "", MyOperatore.Name, Label2(6))

Dim rc As Boolean

        rc = IIf(MyOperatore.IndexPrivilege >= 1, True, False)




End Sub

Private Sub Form_Initialize()
' lbProgram = "Release " & App.Major & "." & App.Minor & "." & App.Revision

Call StartProcedure

SaveSizes

End Sub

Private Sub Form_Load()
Dim rc As Boolean
   
    If bFullScreen Then
        Me.WindowState = 2
    Else
        Me.WindowState = 0
    End If
    

    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer


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



Set F_MAIN_CLP = Nothing

    Dim svar
    svar = EnumWindows(AddressOf getalltopwindows, 0)
    End
    
End Sub



Private Sub frCommandInside_Click(Index As Integer)
    Select Case Index
        Case 0
        Case 1
            Call OpenRevisionHistory
    
    End Select
    
   
End Sub




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


Private Sub SetFrameDefault()
'PicMenu(3).Visible = False

 CodeChemical = ""
 CodeRecipe = ""
 CodeHannaCode = ""

 ChemicalID = 0
 RecipeID = 0
 HannaCodeID = 0

End Sub

Private Sub Grid1_DblClick()

    If CodeChemical <> "" Then
        frClassification_Click 0
    End If

End Sub
Private Sub Grid2_DblClick()

    If CodeRecipe <> "" Then
        frClassification_Click 1
    End If
    
    'PicMenu(3).Visible = False

End Sub
Private Sub Grid1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)


SetFrameDefault

If FirstRow > 0 Then
    
    CodeChemical = Grid1.Cell(FirstRow, 1).Text
    ChemicalID = Grid1.Cell(FirstRow, 11).Text


End If


End Sub

Private Sub Grid2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
SetFrameDefault

If FirstRow > 0 Then
    RecipeCode = Grid2.Cell(FirstRow, 1).Text
    CodeRecipe = Grid2.Cell(FirstRow, 1).Text
    RecipeID = Grid2.Cell(FirstRow, 5).Text
    'PicMenu(3).Visible = True

End If

End Sub

Private Sub Grid3_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
'PicMenu(3).Visible = False
If FirstRow > 0 Then
    
    CodeHannaCode = Grid3.Cell(FirstRow, 1).Text
    HannaCodeID = Grid3.Cell(FirstRow, 7).Text


End If
End Sub

Private Sub Image1_Click()
txSearch = TX_INTRO
SearcMRInTable (True)
End Sub

Private Sub Image3_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub Label2_Click(Index As Integer)
PicMenu_Click Index
End Sub



Private Sub lbClassification_Click(Index As Integer)
frClassification_Click Index
End Sub

Private Sub lbCommandInside_Click(Index As Integer)
    Select Case Index
        Case 0
            SearcMRInTable (False)
            
    
    End Select
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
   
  
    
    IndexProcedura = Index
    
    lbCommandInside(0).Caption = " Search " + Label2(Index)
    
    PicMain(Index).ZOrder
    PicMain(Index).Visible = True
    
    Select Case IndexProcedura
        Case 0
          'PicMenu(3).Visible = False
            
        Case 1
          'PicMenu(3).Visible = False
        Case 2
            'PicMenu(3).Visible = False
        
        Case 3
            ' revision history
            OpenRevisionHistory
    End Select
    
    
    
    SaveSetting App.Title, "Intro", "IndexProcedura", IndexProcedura
    TimeriNTRO.Enabled = False

End Function




Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer


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




Private Sub SetPicForm()
'On Error GoTo ERR_SET:
Dim i As Integer


BottomLine.x1 = 0
BottomLine.x2 = Me.Width



PicMain(0).Move 0, PBTitle.Top + PBTitle.Height, Me.Width, BottomLine.y1 - (PBTitle.Top + PBTitle.Height)


    
    For i = 1 To PicMain.UBound
        PicMain(i).Top = PicMain(0).Top
        PicMain(i).Left = PicMain(0).Left
        PicMain(i).Width = PicMain(0).Width
        PicMain(i).Height = PicMain(0).Height
    Next


Exit Sub
ERR_SET:
Resume Next
End Sub




Private Sub StartProcedure()
Call ClearText
Call SetPicForm



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

frAvvio.Visible = True
frAvvio.Visible = False
Timer1.Enabled = False
End Sub


Private Sub TimeriNTRO_Timer()
FormIntro
TimeriNTRO.Enabled = False
End Sub

Private Sub FormIntro()

Dim rc As Boolean
Dim Grid(10) As Grid
    
        
    F_MAIN_CLP.BackColor = vbColorRed
    'lbWait.ZOrder
    'lbWait.Visible = True
    
    
    rc = CheckPrimoAvvio
    
    Set Grid(0) = Grid1
    Set Grid(1) = Grid2
    Set Grid(2) = Grid3
    'Set Grid(3) = Grid7
    'Set Grid(4) = Grid5
    'Set Grid(5) = Grid6
    'Set Grid(6) = Grid4
    
    Call SetAllMainCLPGrid(Grid())
      
    '-----------------------------------
    ' revision history
    '-----------------------------------
    Call SetGridRecipeRevision(Grid8)
    
  
  
    Call SetColumnWidth
  
    Call AddChemicalRMinGrid(Grid1)
    Call AddRecipeinGrid(Grid2)
    Call AddHannaCodeinGrid(Grid3)
    
    If rc Then
        SelectProcedura (GetSetting(App.Title, "Intro", "IndexProcedura", 0))
    End If
       

    SetAllClassificationByRecipe
  

 'lbWait.Visible = False
 F_MAIN_CLP.BackColor = &H4D3B37
 
End Sub

Private Sub lbInside_Click(Index As Integer)
    Select Case Index
        Case 0
            
            Call AddChemicalRMinGrid(Grid1)
        
            PopupMessage 2, "Reload Table....", , , "Chemical RM"
        
        Case 1
        
            Call AddRecipeinGrid(Grid2)
            
            PopupMessage 2, "Reload Table....", , , "Recipe"
        
        Case 2
        
            PopupMessage 2, "Reload Table....", , , "Hanna Code"
        
    
    End Select
End Sub


Private Function CheckPrimoAvvio() As Boolean
Dim rc As Boolean

    rc = True
    With dbTabCode
    
        .filter = ""
        If .EOF Then
            rc = False
          
            IndexProcedura = 99
        Else
           
        End If
        
    End With
    
    With dbTabUserAccount
        .filter = ""
        bExistAccount = Not (.EOF)
    End With
    
    If bExistAccount Then
        
    Else
    
    End If
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

Private Sub txSearch_LostFocus()
If Trim(txSearch) = "" Then txSearch = TX_INTRO
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

RecipeCode = txSearchRecipe
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

