VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form F_MAIN 
   BackColor       =   &H00808080&
   Caption         =   "Chemical QC"
   ClientHeight    =   12045
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
   ForeColor       =   &H00404040&
   Icon            =   "F_MAIN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12602.08
   ScaleMode       =   0  'User
   ScaleWidth      =   19320
   StartUpPosition =   2  'CenterScreen
   Begin FlexCell.Grid GrdCode 
      Height          =   3480
      Left            =   3600
      TabIndex        =   41
      Top             =   4920
      Visible         =   0   'False
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   6138
      AllowUserReorderColumn=   -1  'True
      AllowUserSort   =   -1  'True
      Appearance      =   0
      BackColor1      =   14737632
      BackColor2      =   14737632
      BackColorActiveCellSel=   12632256
      BackColorBkg    =   14737632
      BackColorFixed  =   12632256
      BackColorFixedSel=   12632256
      BackColorScrollBar=   -2147483635
      BackColorSel    =   8421504
      BorderColor     =   9849089
      CellBorderColor =   16512
      CellBorderColorFixed=   16777215
      Cols            =   10
      DefaultFontName =   "Calibri"
      DefaultFontSize =   12
      BoldFixedCell   =   0   'False
      DisplayDateTimeMask=   -1  'True
      FixedRowColStyle=   0
      ForeColorFixed  =   4210752
      GridColor       =   16777215
      ReadOnly        =   -1  'True
      Rows            =   10
      SelectionMode   =   3
      MultiSelect     =   0   'False
      DateFormat      =   2
      EnterKeyMoveTo  =   1
      BackColorComment=   -2147483635
      AllowUserPaste  =   2
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   3
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   19215
      TabIndex        =   21
      Top             =   5640
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Graph QC : Select Lot Number , Hanna Code and go to Graph"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   3
         Left            =   480
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   6210
      End
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   2
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   19215
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Evaluation QC : Select Lot Number , Hanna Code and Start Evaluation"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   7110
      End
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H000080DF&
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   1
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   19215
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reading QC : Select Lot Number , Hanna Code and Start Tests"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   6255
      End
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      DrawWidth       =   7
      Height          =   420
      Index           =   0
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   19215
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Information QC : Enter Lot Number , Hanna Code and fill information QC"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   7410
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3360
      MouseIcon       =   "F_MAIN.frx":429A
      ScaleHeight     =   855
      ScaleWidth      =   3015
      TabIndex        =   72
      Top             =   10960
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Image Default 
         Height          =   480
         Index           =   8
         Left            =   1320
         MouseIcon       =   "F_MAIN.frx":45A4
         MousePointer    =   99  'Custom
         Picture         =   "F_MAIN.frx":48AE
         Top             =   195
         Width           =   480
      End
   End
   Begin FlexCell.Grid GrdBatch 
      Height          =   3480
      Left            =   9000
      TabIndex        =   40
      Top             =   5880
      Visible         =   0   'False
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   6138
      AllowUserReorderColumn=   -1  'True
      AllowUserSort   =   -1  'True
      Appearance      =   0
      BackColor1      =   14737632
      BackColor2      =   14737632
      BackColorActiveCellSel=   12632256
      BackColorBkg    =   14737632
      BackColorFixed  =   12632256
      BackColorFixedSel=   12632256
      BackColorScrollBar=   -2147483635
      BackColorSel    =   8421504
      BorderColor     =   9849089
      CellBorderColor =   16512
      CellBorderColorFixed=   16777215
      Cols            =   10
      DefaultFontName =   "Calibri"
      DefaultFontSize =   8.25
      BoldFixedCell   =   0   'False
      ButtonLocked    =   -1  'True
      DisplayDateTimeMask=   -1  'True
      FixedRowColStyle=   0
      ForeColorFixed  =   4210752
      GridColor       =   16777215
      ReadOnly        =   -1  'True
      Rows            =   10
      SelectionMode   =   3
      MultiSelect     =   0   'False
      DateFormat      =   2
      EnterKeyMoveTo  =   1
      BackColorComment=   -2147483635
      AllowUserPaste  =   2
   End
   Begin VB.PictureBox PicIntro 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   9690
      Left            =   10800
      MouseIcon       =   "F_MAIN.frx":7C90
      MousePointer    =   99  'Custom
      ScaleHeight     =   9690
      ScaleWidth      =   19200
      TabIndex        =   15
      Top             =   1200
      Width           =   19200
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7680
         Left            =   840
         Picture         =   "F_MAIN.frx":7F9A
         ScaleHeight     =   7680
         ScaleWidth      =   7680
         TabIndex        =   16
         Top             =   720
         Width           =   7680
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Instruments"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Index           =   2
         Left            =   9240
         TabIndex        =   63
         Top             =   5160
         Width           =   4635
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
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   16440
         TabIndex        =   34
         Top             =   5400
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chemical QC"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1620
         Index           =   0
         Left            =   8880
         TabIndex        =   17
         Top             =   3600
         Width           =   8085
      End
   End
   Begin VB.Timer TimeriNTRO 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   6960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   7800
   End
   Begin VB.Frame frmSTDToleranceInfo 
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Left            =   2280
      TabIndex        =   52
      Top             =   2640
      Visible         =   0   'False
      Width           =   17415
      Begin FlexCell.Grid Grd3 
         Height          =   2760
         Left            =   480
         TabIndex        =   53
         Top             =   1320
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   4868
         AllowUserReorderColumn=   -1  'True
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColor1      =   14737632
         BackColor2      =   14737632
         BackColorActiveCellSel=   12632256
         BackColorBkg    =   14737632
         BackColorFixed  =   12632256
         BackColorFixedSel=   12632256
         BackColorScrollBar=   -2147483635
         BackColorSel    =   8421504
         BorderColor     =   -2147483635
         CellBorderColor =   16512
         CellBorderColorFixed=   9849089
         Cols            =   10
         DefaultFontName =   "Calibri"
         DefaultFontSize =   12
         DefaultFontBold =   -1  'True
         DisplayDateTimeMask=   -1  'True
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   9849089
         ReadOnly        =   -1  'True
         Rows            =   10
         SelectionMode   =   1
         MultiSelect     =   0   'False
         DateFormat      =   2
         EnterKeyMoveTo  =   1
         BackColorComment=   -2147483635
         AllowUserPaste  =   2
      End
      Begin VB.Label lbSpecification 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOT_NUMBER"
         ForeColor       =   &H00C0C0C0&
         Height          =   540
         Left            =   9600
         MouseIcon       =   "F_MAIN.frx":26A9B
         MousePointer    =   99  'Custom
         TabIndex        =   55
         Top             =   600
         Width           =   7215
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   2
         Left            =   480
         Picture         =   "F_MAIN.frx":26DA5
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Standards : Tolerance Information Table"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1080
         MouseIcon       =   "F_MAIN.frx":2A187
         MousePointer    =   99  'Custom
         TabIndex        =   54
         Top             =   720
         Width           =   4125
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox PicMenu 
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
      Index           =   4
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   19215
      TabIndex        =   0
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   1920
         MouseIcon       =   "F_MAIN.frx":2A491
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   9
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   735
            MouseIcon       =   "F_MAIN.frx":2A79B
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":2AAA5
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reading QC"
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
            Index           =   1
            Left            =   345
            MouseIcon       =   "F_MAIN.frx":2DE87
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   720
            Width           =   1200
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   3
         Left            =   5760
         MouseIcon       =   "F_MAIN.frx":2E191
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   13
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   735
            MouseIcon       =   "F_MAIN.frx":2E49B
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":2E7A5
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Graph QC"
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
            Height          =   225
            Index           =   3
            Left            =   0
            MouseIcon       =   "F_MAIN.frx":31B87
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   720
            Width           =   1890
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   3840
         MouseIcon       =   "F_MAIN.frx":31E91
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   11
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   735
            MouseIcon       =   "F_MAIN.frx":3219B
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":324A5
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Evaluation QC"
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
            Height          =   225
            Index           =   2
            Left            =   0
            MouseIcon       =   "F_MAIN.frx":35887
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   720
            Width           =   1890
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MouseIcon       =   "F_MAIN.frx":35B91
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
            Caption         =   "Lot Information QC"
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
            Left            =   60
            MouseIcon       =   "F_MAIN.frx":35E9B
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   735
            MouseIcon       =   "F_MAIN.frx":361A5
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":364AF
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open Batch Table"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   16110
         TabIndex        =   32
         Top             =   300
         Visible         =   0   'False
         Width           =   2580
      End
   End
   Begin ChemicalQC.ctlCalendar ctlCalendar1 
      Height          =   6960
      Left            =   2160
      TabIndex        =   59
      Top             =   240
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
   Begin VB.PictureBox PicIntro2 
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   8880
      Picture         =   "F_MAIN.frx":39891
      ScaleHeight     =   9615
      ScaleWidth      =   19215
      TabIndex        =   64
      Top             =   1320
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Image Image6 
         Height          =   480
         Index           =   8
         Left            =   7680
         Picture         =   "F_MAIN.frx":5776A
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
         ForeColor       =   &H00964901&
         Height          =   1125
         Left            =   -480
         MouseIcon       =   "F_MAIN.frx":5AB4C
         MousePointer    =   99  'Custom
         TabIndex        =   71
         Top             =   4320
         Visible         =   0   'False
         Width           =   19140
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   4
         Left            =   6480
         Picture         =   "F_MAIN.frx":5AE56
         Top             =   5520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No User Account in your Database : Please enter at least 1 User to set Privilege"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   0
         MouseIcon       =   "F_MAIN.frx":5E238
         MousePointer    =   99  'Custom
         TabIndex        =   70
         Top             =   6720
         Visible         =   0   'False
         Width           =   19185
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Operator  / User Account"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   450
         Index           =   1
         Left            =   0
         TabIndex        =   69
         Top             =   5520
         Visible         =   0   'False
         Width           =   19170
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chemical QC : Database Setup"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   480
         TabIndex        =   68
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Your Hanna Code Database is Empty : Please goto Settings > Database and Import Hanna Code form Excel"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   3
         Left            =   0
         MouseIcon       =   "F_MAIN.frx":5E542
         MousePointer    =   99  'Custom
         TabIndex        =   67
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
         ForeColor       =   &H00964901&
         Height          =   885
         Index           =   4
         Left            =   0
         MouseIcon       =   "F_MAIN.frx":5E84C
         MousePointer    =   99  'Custom
         TabIndex        =   66
         Top             =   4080
         Width           =   19125
      End
      Begin VB.Image Im 
         Height          =   480
         Index           =   7
         Left            =   10920
         Picture         =   "F_MAIN.frx":5EB56
         Top             =   3960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Code"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   450
         Index           =   0
         Left            =   0
         TabIndex        =   65
         Top             =   1680
         Width           =   19170
      End
   End
   Begin VB.TextBox txQRCode 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   76
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   14535
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'None
      Height          =   9060
      Index           =   0
      Left            =   120
      ScaleHeight     =   9060
      ScaleWidth      =   21000
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   21000
      Begin VB.Frame frCommandInside 
         BackColor       =   &H000040C0&
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
         Index           =   1
         Left            =   4800
         TabIndex        =   77
         Top             =   6360
         Width           =   3615
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Import Data from File"
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
            TabIndex        =   78
            Top             =   240
            Width           =   3615
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
         Height          =   735
         Index           =   0
         Left            =   8520
         TabIndex        =   74
         Top             =   6360
         Width           =   5895
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Scan QRCode"
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
            Index           =   0
            Left            =   0
            TabIndex        =   75
            Top             =   240
            Width           =   5895
         End
      End
      Begin VB.ComboBox ComboRecipe 
         BackColor       =   &H00404040&
         ForeColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   2
         Left            =   3000
         TabIndex        =   42
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   8160
         TabIndex        =   37
         Top             =   3000
         Width           =   4335
      End
      Begin VB.ComboBox CmbVisual 
         BackColor       =   &H00404040&
         ForeColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   11760
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   1
         Left            =   8160
         TabIndex        =   28
         Top             =   3600
         Width           =   4335
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   0
         Left            =   4800
         MouseIcon       =   "F_MAIN.frx":61F38
         MousePointer    =   99  'Custom
         ScaleHeight     =   1335
         ScaleWidth      =   9615
         TabIndex        =   35
         Top             =   4920
         Width           =   9615
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start New Lot"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   4
            Left            =   3990
            MouseIcon       =   "F_MAIN.frx":62242
            MousePointer    =   99  'Custom
            TabIndex        =   36
            Top             =   840
            Width           =   1605
         End
         Begin VB.Image Image4 
            Height          =   480
            Index           =   0
            Left            =   4560
            Picture         =   "F_MAIN.frx":6254C
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   5
         Left            =   16800
         MouseIcon       =   "F_MAIN.frx":6592E
         MousePointer    =   99  'Custom
         TabIndex        =   49
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   6
         Left            =   14400
         MouseIcon       =   "F_MAIN.frx":65C38
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   7
         Left            =   11280
         MouseIcon       =   "F_MAIN.frx":65F42
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh Table"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   10
         Left            =   1200
         MouseIcon       =   "F_MAIN.frx":6624C
         MousePointer    =   99  'Custom
         TabIndex        =   61
         Top             =   1560
         Width           =   1380
      End
      Begin VB.Image Image3 
         Height          =   480
         Index           =   8
         Left            =   675
         MouseIcon       =   "F_MAIN.frx":66556
         MousePointer    =   99  'Custom
         Picture         =   "F_MAIN.frx":66860
         Top             =   1440
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Code"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   58
         Top             =   360
         Width           =   2070
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   9
         Left            =   120
         MouseIcon       =   "F_MAIN.frx":69C42
         MousePointer    =   99  'Custom
         TabIndex        =   57
         Top             =   0
         Width           =   3615
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   1
         Left            =   480
         Picture         =   "F_MAIN.frx":69F4C
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STD tolerance Info"
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
         Left            =   1080
         MouseIcon       =   "F_MAIN.frx":6D32E
         MousePointer    =   99  'Custom
         TabIndex        =   56
         Top             =   480
         Width           =   1890
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Open Batch Table"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   825
         Left            =   4440
         TabIndex        =   44
         Top             =   1440
         Width           =   10230
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Lot Number + Hanna Code"
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   7800
         TabIndex        =   43
         Top             =   6000
         Width           =   3315
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   5
         Left            =   15000
         Picture         =   "F_MAIN.frx":6D638
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   4
         Left            =   17040
         Picture         =   "F_MAIN.frx":70A1A
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   3
         Left            =   12480
         Picture         =   "F_MAIN.frx":73DFC
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Code"
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
         Left            =   13080
         MouseIcon       =   "F_MAIN.frx":771DE
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Number"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   345
         Index           =   1
         Left            =   5760
         TabIndex        =   38
         Top             =   3000
         Width           =   1710
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Left            =   9360
         Picture         =   "F_MAIN.frx":774E8
         Top             =   5400
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closed Lots"
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
         Left            =   17640
         MouseIcon       =   "F_MAIN.frx":7A8CA
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open Lots"
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
         MouseIcon       =   "F_MAIN.frx":7ABD4
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   480
         Width           =   1035
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   13200
         MouseIcon       =   "F_MAIN.frx":7AEDE
         MousePointer    =   99  'Custom
         Picture         =   "F_MAIN.frx":7B1E8
         Top             =   3240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Login Operator !"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   8745
         TabIndex        =   29
         Top             =   8160
         Width           =   1725
      End
      Begin VB.Image ImMain 
         Height          =   480
         Left            =   9360
         Picture         =   "F_MAIN.frx":7E5CA
         Top             =   7680
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Code"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   345
         Index           =   0
         Left            =   5760
         TabIndex        =   27
         Top             =   3600
         Width           =   1770
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   1095
         Left            =   4440
         Top             =   1560
         Visible         =   0   'False
         Width           =   10335
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   1815
         Left            =   4800
         Top             =   2640
         Visible         =   0   'False
         Width           =   9615
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   8
         Left            =   120
         MouseIcon       =   "F_MAIN.frx":819AC
         MousePointer    =   99  'Custom
         TabIndex        =   62
         Top             =   1080
         Visible         =   0   'False
         Width           =   3615
      End
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Index           =   1
      Left            =   8520
      MouseIcon       =   "F_MAIN.frx":81CB6
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   10920
      Width           =   2655
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   4
      Left            =   0
      MouseIcon       =   "F_MAIN.frx":81FC0
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   10680
      Width           =   1695
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1815
      Index           =   2
      Left            =   13080
      MouseIcon       =   "F_MAIN.frx":822CA
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   10200
      Width           =   2775
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   975
      Index           =   0
      Left            =   17640
      MouseIcon       =   "F_MAIN.frx":825D4
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   11040
      Width           =   1455
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   3
      Left            =   3840
      MouseIcon       =   "F_MAIN.frx":828DE
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   10560
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close Table"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   9
      Left            =   9120
      MouseIcon       =   "F_MAIN.frx":82BE8
      MousePointer    =   99  'Custom
      TabIndex        =   60
      Top             =   11595
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quit Program"
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
      Index           =   8
      Left            =   17880
      MouseIcon       =   "F_MAIN.frx":82EF2
      MousePointer    =   99  'Custom
      TabIndex        =   48
      Top             =   11600
      Width           =   1110
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database Lot"
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
      Index           =   7
      Left            =   13800
      MouseIcon       =   "F_MAIN.frx":831FC
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Top             =   11595
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set Operator"
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
      Left            =   4275
      MouseIcon       =   "F_MAIN.frx":83506
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Top             =   11595
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
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
      Left            =   400
      MouseIcon       =   "F_MAIN.frx":83810
      MousePointer    =   99  'Custom
      TabIndex        =   45
      Top             =   11600
      Width           =   660
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   14160
      MouseIcon       =   "F_MAIN.frx":83B1A
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":83E24
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_MAIN.frx":87206
      Height          =   480
      Index           =   1
      Left            =   9360
      MouseIcon       =   "F_MAIN.frx":8A5E8
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":8A8F2
      Top             =   11040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   4560
      MouseIcon       =   "F_MAIN.frx":8DCD4
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":8DFDE
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   4
      Left            =   480
      MouseIcon       =   "F_MAIN.frx":913C0
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":916CA
      Top             =   11040
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   483
      X2              =   18837
      Y1              =   11173.95
      Y2              =   11173.95
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
      X1              =   14490
      X2              =   14490
      Y1              =   0
      Y2              =   12429.45
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   4830
      X2              =   4830
      Y1              =   0
      Y2              =   12429.45
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   9660
      X2              =   9660
      Y1              =   376.65
      Y2              =   12806.1
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_MAIN.frx":94AAC
      Height          =   480
      Index           =   0
      Left            =   18240
      MouseIcon       =   "F_MAIN.frx":97E8E
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":98198
      Top             =   11040
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
Private lRowLot As Long
Private lRow As Long

Private TimerCount As Integer

Private SelectedCode As String
Private SelectedLot As String
Private SelectedCodeID As Long
Private m_rc As Boolean
Private bNotSearchRecipe As Boolean
Private OrderByIndex As Integer
Private OrderCell As Integer
Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

Private m_ControlPositions() As ControlPositionType

Private UserQrCode As QRCodeType


Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long



Private Sub Form_Activate()
  



    
If MyOperatore.Name <> "" Then
    Label5 = MyOperatore.Name
End If
DefaultMenuLabel(8).Move DefaultMenuLabel(9).Left, DefaultMenuLabel(9).Top, DefaultMenuLabel(9).Width, DefaultMenuLabel(9).Height
Image3(8).Left = Image4(1).Left
Image3(8).Top = Image4(1).Top
Label2(10).Left = Label12.Left
Label2(10).Top = Label12.Top





End Sub


Private Sub ComboRecipe_Click()

'bNotSearchRecipe = IIf(ComboRecipe.ListIndex > 0, True, False)
'Call CheckLots(False, bNotSearchRecipe)

Call CheckInTable(ComboRecipe)

End Sub

Private Sub ComboRecipe_GotFocus()
 AddComboRecipe
End Sub

Private Sub Default_Click(Index As Integer)
Picture1_Click
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    FrmMove = True
    DragX = x
    DragY = Y
    If Me.WindowState = 2 Then
        FrmMove = False
       
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim nx, ny
    If Me.WindowState = 2 Then
        FrmMove = False
        Exit Sub
    End If
    nx = Me.Left + x - DragX
    ny = Me.Top + Y - DragY
    Me.Left = nx
    Me.Top = ny
    FrmMove = False
End Sub

Public Function DoShow() As Boolean

    On Error GoTo ERR_SHOW
    mOk
    m_rc = False
    
    
    'Call CheckPrimoAvvio

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


Private Sub GetSettings()








End Sub





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
On Error Resume Next
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

                Ctl.DefaultFont.Size = 12 * m_ControlGridFontSize
                Ctl.DefaultRowHeight = 30 * m_ControlGridRowHeight
           
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

Private Sub CmbVisual_Click()
SaveSetting App.Title, "Settings Filtro", "Filtro Data", CmbVisual.ListIndex
Call CheckLots(bSearchClosedLot, bNotSearchRecipe)
End Sub

Private Sub DefaultMenuLabel_Click(Index As Integer)

CloseSettingDataFile

Select Case Index
    Case 0
            If F_MsgBox.DoShow("Quit Chemical QC?", "Exit") Then

                CloseSettingDataFile
                
                If bAddNewDatabaseRelease Then AddReleaseNumber
                
                If F_MsgBox.DoShow("Turn off PC?", ProjectName) Then
                    ShutDownNT True
                End If
                
                
                Unload Me
                
                Exit Sub
            End If
    Case 1
        ' close table
        CleanformLot False
        
            
    Case 2
        FormGrid.Top = Me.Top
        FormGrid.Left = Me.Left
        
        If FormGrid.DoShow Then
        
            DefaultMenuLabel_Click 8
        
        End If
        
    Case 3
       frmLogin.DoShow
       Label5 = MyOperatore.Name
    Case 4
        F_SETTING.WindowState = Me.WindowState
        F_SETTING.Left = Me.Left
        F_SETTING.Top = Me.Top
        F_SETTING.DoShow , , , DefaultMenu(4)
        FormIntro
    Case 5
        Label7_Click
        
    Case 6
        Label6_Click
    Case 7
        Label8_Click
    Case 8
        Text1(0) = ""
        Text1(1) = ""
        CloseSettingDataFile
        Call CheckLots(bSearchClosedLot, bNotSearchRecipe)
        'ComboRecipe.ListIndex = -1
        DoEvents
        Text1(0).SetFocus
        'PopupMessage 2, "Table Reloaded..."

    Case 9
        frmSTDToleranceInfo.ZOrder
        frmSTDToleranceInfo.Visible = Not (frmSTDToleranceInfo.Visible)
        GrdBatch.Visible = False
        'frCommandInside(2).Visible = False
        Exit Sub
        
End Select
frmSTDToleranceInfo.Visible = False
End Sub

Private Sub DisableImage_Click()
PopupMessage 2, "Warning : Administrator Only can Operate...", , True
End Sub

Private Sub Form_Initialize()
 lbProgram = "Release " & App.Major & "." & App.Minor & "." & App.Revision

Call StartProcedure

SaveSizes


 ' Call ShowWindow(Me.hWnd, vbHide)
    Me.Caption = Me.Caption
   ' Call ShowWindow(Me.hWnd, vbNormalNoFocus)
    
    
    
    
    
    
 If bFullScreen Then Me.WindowState = 2
   


End Sub
Private Sub ResizeTab()
Dim rc As Boolean

    With GrdCode

      .AutoRedraw = False
        
        .DefaultRowHeight = 30 * m_ControlGridRowHeight
        .Column(0).Width = 30 * m_ControlGridColWidth
        .Column(1).Width = 200 * m_ControlGridColWidth
        .Column(2).Width = 350 * m_ControlGridColWidth
        .Column(3).Width = 100 * m_ControlGridColWidth
        .Column(4).Width = 100 * m_ControlGridColWidth
        
        .DefaultFont.Size = 12 ' * m_ControlGridFontSize
        .DefaultFont.Bold = False
        
        
      .AutoRedraw = True
      .Refresh
      '.ReadOnly = True
      
    End With
    
    With GrdBatch

      .AutoRedraw = False
      
       ' .Cell(0, 0).Text = "n."
       ' .Cell(0, 1).Text = "Lot Number"
       ' .Cell(0, 2).Text = "Code SFG"
       ' .Cell(0, 3).Text = "Description"
       ' .Cell(0, 4).Text = "Recipe"
       ' .Cell(0, 5).Text = "Prep. Week"
       ' .Cell(0, 6).Text = "Range Min"
       ' .Cell(0, 7).Text = "Range Max"
       ' .Cell(0, 8).Text = "Date"
       ' .Cell(0, 9).Text = "Exp.Date"
       ' .Cell(0, 10).Text = "# Test" ' quanti test ho fatto
       ' .Cell(0, 11).Text = "Mean Value" ' se ho fatto calcolo medie
       ' .Cell(0, 12).Text = "Finalise" ' se ho finalizzat ( solo Laboratory Manager )
       ' .Cell(0, 13).Text = "QC Operator"
       ' .Cell(0, 14).Text = "QC Note"
       ' .Cell(0, 15).Text = "ID"
       ' .Cell(0, 16).Text = "FileName"
       ' .Cell(0, 17).Text = "NomeFileReport"
       ' .Cell(0, 18).Text = "NomeFileExcel"
       ' .Cell(0, 19).Text = "CODE_ID"

        
        .DefaultRowHeight = 30 * m_ControlGridRowHeight
        .Column(0).Width = 30 * m_ControlGridColWidth
        .Column(1).Width = 150 * m_ControlGridColWidth
        .Column(2).Width = 150 * m_ControlGridColWidth
        .Column(3).Width = 300 * m_ControlGridColWidth
        .Column(4).Width = 100 * m_ControlGridColWidth
        .Column(5).Width = 100 * m_ControlGridColWidth
        .Column(6).Width = 100 * m_ControlGridColWidth
        .Column(7).Width = 100 * m_ControlGridColWidth
        .Column(8).Width = 100 * m_ControlGridColWidth
        .Column(9).Width = 100 * m_ControlGridColWidth
        .Column(10).Width = 100 * m_ControlGridColWidth
        .Column(11).Width = 0 * m_ControlGridColWidth
        .Column(12).Width = 0
        .Column(13).Width = 0
        .Column(14).Width = 0
        ' closed
        '.Column(5).Width = 100 * m_ControlGridColWidth
        '.Column(7).Width = 100 * m_ControlGridColWidth
        '.Column(8).Width = 100 * m_ControlGridColWidth
        '.Column(9).Width = 100 * m_ControlGridColWidth
      
        
        .DefaultFont.Size = 11 * m_ControlGridFontSize
        .DefaultFont.Bold = False
        
        
      .AutoRedraw = True
      .Refresh
      '.ReadOnly = True
      
    End With
    
    
    
  
End Sub


Private Sub Form_Load()
PicIntro2.Top = PicMenu(4).Height
PicIntro2.Left = 0




'If Screen.Width - Me.Width > 1000 And bFullScreen Then
    'Me.WindowState = 2
    
    'If Screen.Width - Me.Width > 2000 And bFullScreen Then Me.Picture = LoadPicture(PictureMaxScreen)
    PicMain(0).Picture = LoadPicture(PictureMaxScreen)
    PicIntro.Picture = LoadPicture(PictureMaxScreen)
    If Screen.Width - Me.Width > 2000 And bFullScreen Then
        Picture3.Visible = False
       ' Label3(0).Left = PicIntro.Width / 2 - Label3(0).Width / 2
      '  Label3(0).Top = PicIntro.Height / 2 - Label3(0).Height / 2
    

    End If
    Picture3.Visible = False
    
  DropShadow Me.hWnd
'Else
   ' Me.WindowState = 0
'End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Integer
If Me.WindowState = 2 Then
    FrmMove = False
End If
Dim nx, ny
    If FrmMove Then
        nx = Me.Left + x - DragX
        ny = Me.Top + Y - DragY
        Me.Left = nx
        Me.Top = ny
    End If
    
For i = 0 To 3
    If i = IndexProcedura Then
    Else
        PicMenu(i).BackColor = vbColorDarkFont
    End If
Next
End Sub





Private Sub Form_Resize()
'SetPicForm

ResizeControls
If Picture3.Visible = False Then
    Label3(0).Left = PicIntro.Width / 2 - Label3(0).Width / 2
    Label3(0).Top = PicIntro.Height / 2 - Label3(0).Height / 2
     Label3(2).Left = Label3(0).Left
    Label3(2).Top = Label3(2).Top + 500
    lbProgram.Top = Label3(2).Top + 400
    lbProgram.Left = Label3(0).Left + Label3(0).Width - lbProgram.Width
End If
ResizeTab

End Sub

Private Sub Image2_Click(Index As Integer)
Select Case Index
    Case 0
        PopupMessage 2, "Add New Lot..."
    Case 1
        PopupMessage 2, "Delete Lot..."
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set F_MAIN = Nothing
    Dim svar
    svar = EnumWindows(AddressOf getalltopwindows, 0)
    End
    
    
End Sub

Private Sub frCommandInside_Click(Index As Integer)
    Select Case Index
        Case 0
          ' SCANN BARCODE
            Call ScanQRCodeQC
        Case 1
            ' import data from file...
            Call ImportDataFromFile
        
    End Select
    
    
  
    
End Sub

Private Sub frmSTDToleranceInfo_Click()
frmSTDToleranceInfo.Visible = False
End Sub

Private Sub frmSTDToleranceInfo_DragDrop(Source As Control, x As Single, Y As Single)
frmSTDToleranceInfo.Visible = False
End Sub

Private Sub GrdBatch_BeforeUserSort(ByVal Col As Long)
lRow = 0
OrderByIndex = Col

OrderCell = GrdBatch.Column(OrderByIndex).UserSortIndicator


End Sub

Private Sub GrdBatch_DblClick()
lRowCode = 0
CloseSettingDataFile
If lRow > 0 Then
    OpenAnyLots 0
    lRow = 0
    
     DefaultMenuLabel_Click 8
End If

End Sub


Private Sub GrdBatch_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
SettingName = 0
SelectedCodeID = 0
SelectedCode = ""
lRow = FirstRow
'Picture1.Visible = False
'frCommandInside(2).Visible = False
If lRow > 0 Then

    Text1(0) = Trim(GrdBatch.Cell(lRow, 1).Text)
    Text1(1) = Trim(GrdBatch.Cell(lRow, 2).Text)
    SettingName = Trim(GrdBatch.Cell(lRow, 16).Text)
    SelectedCodeID = GetCodeID(Trim(GrdBatch.Cell(lRow, 2).Text), Trim(GrdBatch.Cell(lRow, 6).Text), Trim(GrdBatch.Cell(lRow, 7).Text))
    SelectedCode = Trim(GrdBatch.Cell(lRow, 2).Text)
    SelectedLot = Trim(GrdBatch.Cell(lRow, 1).Text)
    'frCommandInside(2).Visible = True
    Picture1.Visible = True
    lRowLot = lRow
Else

End If

End Sub

Private Sub GrdCode_BeforeUserSort(ByVal Col As Long)
lRow = 0
lRowCode = 0
End Sub

Private Sub GrdCode_DblClick()
lRow = 0
If lRowCode > 0 Then
    Text1(0) = ""
    Text1(1) = SelectedCode
    Text1(0).SetFocus
    lRowCode = 0
    Label8_Click
End If
SelectedCode = ""
End Sub

Private Sub GrdCode_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
lRowCode = FirstRow

If FirstRow > 0 Then
    SelectedCode = Trim(GrdCode.Cell(FirstRow, 1).Text)
    SelectedCodeID = Trim(GrdCode.Cell(FirstRow, 7).Text)
    MyCode = SelectedCode
Else
    SelectedCode = ""
    SelectedCodeID = 0

End If
End Sub

Private Sub Image1_Click()
CelarForm
Text1(0).SetFocus
End Sub

Private Sub Image3_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub Image4_Click(Index As Integer)
If Index = 0 Then OpenAnyLots 0
End Sub

Private Sub Lab_Click(Index As Integer)
If Index = 4 Then

    F_SETTING.DoShow (1)
    FormIntro
End If
End Sub

Private Sub Label14_Click()

    F_SETTING.DoShow (0)
    FormIntro
End Sub

Private Sub Label2_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub Label6_Click()
Dim rc As Boolean
rc = True

blTable = "Open Lot List"
blTable.Visible = rc
ShowSTDInfo Not (rc)
bSearchClosedLot = False
Call CheckLots(bSearchClosedLot, bNotSearchRecipe)
CleanformLot (True)
End Sub


Private Sub Label7_Click()
Dim rc As Boolean
rc = True
blTable = "Closed Lot List"
bSearchClosedLot = True
blTable.Visible = rc
 ShowSTDInfo Not (rc)
Call CheckLots(bSearchClosedLot, bNotSearchRecipe)
Call CleanformLot(True, 1)
End Sub

Private Sub Label8_Click()
Dim rc As Boolean
  bSearchClosedLot = False
rc = Not (GrdCode.Visible)
blTable = "Select or search Hanna SFG Code"
Call CleanformCode(Not (GrdCode.Visible), 1)
blTable.Visible = rc
ShowSTDInfo Not (rc)
End Sub

Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub

Private Sub lbSpecification_Click()
frmSTDToleranceInfo.Visible = False
End Sub

Private Sub PicIntro_Click()
TimerIntro.Enabled = False
FormIntro
End Sub

Private Sub PicMenu_Click(Index As Integer)
Text1(0).SetFocus

If IndexProcedura = Index Then
Else
    Call SelectProcedura(Index)
End If


End Sub


Private Function SelectProcedura(ByVal Index As Integer) As Boolean
Dim rc As Boolean
Dim i As Integer
If Index > 3 Then Exit Function
For i = 0 To 3
    If i = Index Then
        PicMenu(i).BackColor = vbColorForeFixed
        PicInfo(i).Visible = True
    Else
        PicInfo(i).Visible = False
        PicMenu(i).BackColor = vbColorDarkFont
    End If
Next
Set Image4(0) = Image3(Index)
Picture4(0).BackColor = PicInfo(Index).BackColor
Label2(4) = "Start : " & Label2(Index)
Label11 = Label2(Index)
Label11.ForeColor = PicInfo(Index).BackColor
'Shape2.BackColor = PicInfo(Index).BackColor
'Shape1.BackColor = PicInfo(Index).BackColor
Picture1.BackColor = PicInfo(Index).BackColor
IndexProcedura = Index

Select Case IndexProcedura
    Case 0
        FormSelected True
        
    Case 1
        FormSelected False
    Case 2
        FormSelected False
    Case 3
        FormSelected False
    
End Select
'CheckLots bSearchClosedLot
PicIntro.Visible = False
frmSTDToleranceInfo.Visible = False
PicMain(0).Visible = True
'CleanformLot (False)
SaveSetting App.Title, "Intro", "IndexProcedura", IndexProcedura
TimerIntro.Enabled = False
End Function

Private Sub CleanformLot(ByVal bValue As Boolean, Optional ByVal Index As Integer = 0)
DefaultMenu(1).Visible = bValue
Label2(9).Visible = bValue
GrdBatch.Visible = bValue
'frCommandInside(2).Visible = False
ComboRecipe.Visible = bValue
RefreshTableForm (bValue)
'GrdCode.Visible = False
CmbVisual.Visible = IIf(Index = 1, GrdBatch.Visible, False)
If bValue Then CleanformCode False
If Me.Visible Then If Not (bValue) Then Image1_Click

End Sub



Private Sub CleanformCode(ByVal bValue As Boolean, Optional ByVal Index As Integer = 0)
  '  ShowSTDInfo Not (bValue)
    GrdCode.Visible = bValue
    Text1(2).Visible = bValue
    Label4(2).Visible = bValue
    
 
    
    
    If bValue Then
        CleanformLot False
        Text1(2) = ""
        Call FillGridCode(GrdCode, , True)
        If GrdCode.Rows > 1 Then GrdCode.Cell(lRowCode, 1).SetFocus
        Text1(2).SetFocus
    End If

End Sub
Private Sub ShowSTDInfo(ByVal bValue As Boolean)

If bValue Then If Text1(1) = "" Then bValue = False
If GrdBatch.Visible Then bValue = False

    Label12.Visible = (bValue)
    Image4(1).Visible = (bValue)
    DefaultMenuLabel(9).Visible = (bValue)

End Sub
Private Sub FormSelected(ByVal bValue As Boolean)
        DefaultMenuLabel(5).Visible = bValue
        Label8.Visible = bValue
        Image4(3).Visible = bValue
End Sub



Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Integer

  
 
For i = 0 To 3
    If i = IndexProcedura Then
        ' lascialo stare...
    ElseIf i = Index Then
        PicMenu(i).BackColor = &H505050
    Else
        PicMenu(i).BackColor = vbColorDarkFont
    End If
Next
End Sub






Private Sub SetPicForm()
On Error GoTo ERR_SET:
ctlCalendar1.Left = Me.Width / 2 - ctlCalendar1.Width / 2
ctlCalendar1.Top = Me.Height / 2 - ctlCalendar1.Height / 2
PicMain(0).Left = 0
PicMain(0).Top = PicMenu(4).Height + PicInfo(0).Height
PicMain(0).Width = Me.Width
PicMain(0).Height = Line1.Y1 - PicMain(0).Top
PicIntro.Left = 0
PicIntro.Top = PicMenu(4).Height
PicIntro.Width = Me.Width
PicIntro.Height = Line1.Y1 - PicIntro.Top
PicIntro.BackColor = &H929292

PicInfo(3).BackColor = vbTimBlue
Label1(3).ForeColor = vbWhite 'vbTimBlue 'vbColorTextBlue
'Picture1(3).BackColor = vbTimBlue 'vbColorTextBlue
'Line3(3).BorderColor = vbTimBlue 'vbColorTextBlue
'Picture4(2).BackColor = vbTimBlue 'vbColorTextBlue

GrdBatch.ReadOnly = True
GrdBatch.Left = 360
GrdBatch.Top = PicMain(0).Top + ComboRecipe.Height + ComboRecipe.Top + 300
GrdBatch.Width = Me.Width - GrdBatch.Left * 2
GrdBatch.Height = Line1.Y1 - GrdBatch.Top - 180
GrdCode.Top = GrdBatch.Top
GrdCode.Left = GrdBatch.Left
GrdCode.Width = GrdBatch.Width
GrdCode.Height = GrdBatch.Height
frmSTDToleranceInfo.Left = Me.Width / 2 - frmSTDToleranceInfo.Width / 2
frmSTDToleranceInfo.Top = Me.Height / 2 - frmSTDToleranceInfo.Height / 2
GrdCode.ZOrder
GrdBatch.ZOrder
Dim i As Integer
For i = 0 To 3
    
    PicInfo(i).Left = 0
    PicInfo(i).Top = PicMenu(4).Height
 
Next
Exit Sub
ERR_SET:
Resume Next
End Sub

Private Sub Picture1_Click()
GrdBatch_DblClick
End Sub

Private Sub Picture4_Click(Index As Integer)
   ' controllo il CODICE e il lotto
    Text1(0) = Trim(Text1(0))
    Text1(1) = Trim(Text1(1))
    
    SelectedCode = Text1(1)
    SelectedLot = Text1(0)
    
    
OpenAnyLots (Index)
End Sub


Private Sub OpenAnyLots(Index As Integer, Optional ByVal FileName As String)
Dim MyNewIndex As Integer
Dim Frm As Form
Dim rc As Boolean
Dim StringName As String
If MyOperatore.Name = "" Then

    If frmLogin.DoShow Then
        Label5 = MyOperatore.Name
    Else
        Exit Sub
    End If
End If


   
    USER_PATH = IIf(bSearchClosedLot, USER_DATA_PATH, USER_TEMP_PATH)
    
 
    
    If SelectedCode = "" And Text1(1) <> "" Then SelectedCode = Text1(1)
    
    If FileName <> "" Then
        CloseSettingDataFile
        SettingName = FileName
        GoTo cont:
    End If
    
    With dbTabCode
        .filter = ""
        If SelectedCode <> "" Then
            .filter = "Code='" & SelectedCode & "'"
            
        ElseIf SelectedCodeID > 0 Then
            .filter = "ID='" & SelectedCodeID & "'"
        Else
            PopupMessage 2, "Code not found....", , True
            Exit Sub
        End If
        
        If .EOF Then
            MessageInfoTime = 2500
            PopupMessage 2, SelectedCode & " : Invalid code..." & vbCrLf & "Check and Select Hanna Code form Table", , True
            DefaultMenuLabel_Click 7
            Exit Sub
            
        Else
            StringName = SelectedLot & " " & SelectedCode & " " & IIf(IsNull(Trim(!RangeMin)), "", Trim(!RangeMin)) & " " & USER_ESTENSIONE
            SettingName = FormatNomeFile(StringName)
            
            
            MyLot = SelectedLot
            MyCode = SelectedCode
            
            
            
            
        End If

    
    End With
cont:
    If SettingName <> "" Then
       If FileExists(USER_TEMP_PATH & SettingName) Then
       ElseIf FileExists(USER_DATA_PATH & SettingName) Then
            PopupMessage 2, "Lot : " & SelectedLot & vbCrLf & "Code : " & SelectedCode & vbCrLf & "This Lot Is Closed..."
            bSearchClosedLot = True
       Else
       
            GoTo warning
       End If
       
       
        MyLot = SelectedLot
        MyCode = SelectedCode
            
    Else
warning:


       
        
        If IndexProcedura > 0 Then
            MessageInfoTime = 2100
            PopupMessage 2, "Lot : " & SelectedLot & vbCrLf & "Code : " & SelectedCode & vbCrLf & "WARNING : NO INFORMATION QC...", , True
            Exit Sub
        End If
    
    End If

    Select Case Index
        Case 0
            '-------------------------------------
            ' apro la procedura selezionata
            '-------------------------------------
         
            Select Case IndexProcedura
                Case 0
                    Set Frm = F_INFORMATION
                   
                Case 1
                    Set Frm = F_READING
                Case 2
                    F_EVALUATION.WindowState = Me.WindowState
                    F_EVALUATION.Top = Me.Top
                    F_EVALUATION.Left = Me.Left
                    F_EVALUATION.DoShow MyNewIndex, MyLot, MyCode, SelectedCodeID, Image3(IndexProcedura), SettingName
                
                    Exit Sub
                    
                Case 3
                    F_GRAPH.WindowState = Me.WindowState
                    F_GRAPH.Top = Me.Top
                    F_GRAPH.Left = Me.Left
                    F_GRAPH.DoShow MyNewIndex, MyLot, MyCode, SelectedCodeID, Image3(IndexProcedura), SettingName
                    Exit Sub
            End Select
                   Frm.WindowState = Me.WindowState
                   Frm.Top = Me.Top
                   Frm.Left = Me.Left
                   MyNewIndex = IndexProcedura
                   rc = Frm.DoShow(MyNewIndex, MyLot, MyCode, SelectedCodeID, Image3(IndexProcedura), SettingName)
                  
                   
                   GoSub CheckNewProcedura
        Case 1
            ' tabella lotti chiusi
            Label7_Click
        Case 2
            GrdCode.Visible = Not (GrdCode.Visible)
        Case 4

 
    End Select
    
  
    
    Exit Sub
    
CheckNewProcedura:




 
            If rc Then
            
               ' Text1(0) = MyLot
               ' Text1(1) = MyCode
               
              
                
                If MyNewIndex <> IndexProcedura Then
                    '-------------------------------------
                    ' apro la nuova procedura
                    '-------------------------------------
                    IndexProcedura = MyNewIndex
                    Call SelectProcedura(IndexProcedura)
                    SelectedCode = MyCode
                    OpenAnyLots 0
                End If
            Else
                 Call CheckLots(bSearchClosedLot, bNotSearchRecipe)
            End If
            
            
            
    
    Return
End Sub


Private Sub CelarForm()
Dim i As Integer
    For i = 0 To Text1.Count - 1
        Text1(i) = ""
    Next
    Picture4(0).Visible = False
    DisableImage.Visible = True
    Label10.Visible = DisableImage.Visible
    blTable.Visible = False
    
End Sub

Private Sub Text1_Change(Index As Integer)
Dim rc As Boolean
Dim mrc As Boolean
Dim sString As String

'bSearchClosedLot = False

rc = IIf(Len(Text1(1)) > 0, True, False)
If rc = False Then
    SelectedCode = ""
    SelectedCodeID = 0
End If
rc = IIf(Len(Text1(0)) > 0, rc, False)

DisableImage.Visible = Not (rc)
Label10.Visible = DisableImage.Visible
Picture4(0).Visible = rc

Image1.Visible = IIf(Len(Text1(0)) > 0 Or Len(Text1(1)) > 0, True, False)




Select Case Index
    Case 0
        MyLot = Text1(0)
        
        'Call LotSelected
        
    Case 1
        MyCode = Text1(1)
        If rc = False Then STDToleranceLabel (False)
        lbSpecification = ""
        mrc = FillSTDToleranceGrid(Text1(1), Grd3, lbSpecification, SelectedCodeID)
        STDToleranceLabel (mrc)
        ShowSTDInfo IIf(Len(Text1(Index)) > 0, True, False)
    Case 2
        ' Text1(1) = Text1(2)
         
        If InStr(UCase(Text1(2)), UCase("code")) Then
        
        Else
            
            Call FillGridCode(GrdCode, Trim(Text1(2)), True)
            If GrdCode.Rows > 1 Then GrdCode.Cell(lRowCode, 1).SetFocus
        End If

End Select



End Sub



Private Sub StartProcedure()
Call SetPicForm
Call CelarForm
Call CleanformLot(False)
Call CleanformCode(False)
Call SetGrid(GrdBatch)
Call SetGridCode(GrdCode)
Call SetGridStandardTolerance(Grd3)

STDToleranceLabel False
    
With CmbVisual
    .Clear
    .AddItem "Day"
    .AddItem "Month"
    .AddItem "Year"
    .AddItem "Archive"
     .ListIndex = GetSetting(App.Title, "Settings Filtro", "Filtro Data", 0)
End With
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = Text1.Count - 1 Then
            Text1(0).SetFocus
        Else
            
            On Error Resume Next
            
            
            If Text1(Index + 1).Enabled Then Text1(Index + 1).SetFocus
        End If
        
    
    End If
End Sub





Private Sub MN_ONLINE_Click()

    CreateVerFile
    

    If FileExists(App.Path & "\smartupdate.exe") Then
   
        ApriEseguibile App.Path & "\smartupdate.exe"
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

If TimerCount > 1 Then

   
    Call GetSettings


    TimerCount = 0
    Timer1.Enabled = False
            
           
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
 
 
'lbOperatore = Operatore
        
 

 
End If
TimerCount = TimerCount + 1


End Sub

Private Sub TimerIntro_Timer()
FormIntro
TimerIntro.Enabled = False
End Sub

Private Sub FormIntro()
PicIntro.Visible = False
If CheckPrimoAvvio Then

    SelectProcedura (GetSetting(App.Title, "Intro", "IndexProcedura", 0))
End If


    
End Sub


Private Sub STDToleranceLabel(ByVal bValue As Boolean)
DefaultMenuLabel(9).Visible = bValue
Image4(1).Visible = bValue
Label12.Visible = bValue

End Sub


Private Function CheckInTable(ByVal strCombo As String)
Dim bTutti As Boolean
Dim i As Integer

If InStr(LCase(strCombo), "select") Or strCombo = "" Then bTutti = True


With GrdBatch
    
    .AutoRedraw = False

    For i = 1 To .Rows - 1
        
        If InStr(UCase(.Cell(i, 4).Text), UCase(strCombo)) Or bTutti Then
        
            .RowHeight(i) = 35
        
        Else
        
            .RowHeight(i) = 0
        End If
    
    
    
    Next


    .Refresh
    .AutoRedraw = True

End With

End Function




Private Function CheckLots(ByVal bChiuso As Boolean, ByVal bSearchRecipe As Boolean)
Dim rc As Boolean
Dim i As Integer
Dim Periodo As String

On Error GoTo ERR_CHECK

    With GrdBatch
    
        .AutoRedraw = False
        
        If bChiuso Then
            Periodo = CmbVisual
            rc = True
        Else
            Periodo = ""

        End If

        CloseSettingDataFile
        
        Dim StartTime As Date
        Dim EndTime As Date
        StartTime = Now()
            Call FillGridBatch(bChiuso, Periodo)
        EndTime = Now()
        
        CloseSettingDataFile
        
      
        .Column(1).AutoFit
        .Column(2).AutoFit
        .Column(3).AutoFit
        .Column(4).AutoFit
        .Column(5).AutoFit
        .Column(6).AutoFit

        ' carica la ComboRecipe
        
        'If bSearchRecipe = False Then AddComboRecipe
        If OrderByIndex > 0 Then
            .Column(OrderByIndex).Sort OrderCell
        End If
        
        .Refresh
        .AutoRedraw = True
        
         If .Rows > 1 And lRowLot > 0 Then
         
            If lRowLot > .Rows Then GoTo ERR_END
            
            .Cell(lRowLot, 1).EnsureVisible
            .Cell(lRowLot, 1).SetFocus
            Call GrdBatch_SelChange(lRowLot, 1, lRowLot, 1)
        End If
    End With
    
ERR_END:
    Exit Function
ERR_CHECK:
    MsgBox Err.Description
    Resume Next
End Function

Private Function AddComboRecipe()

Dim strRecipe() As String
Dim Recipe As String
Dim i As Integer
Dim t As Integer
t = 0
 With GrdBatch
    ComboRecipe.Clear
    If .Rows > 1 Then
        
        ComboRecipe.AddItem "Select Recipe"
        For i = 1 To .Rows - 1
            
            Recipe = Trim(.Cell(i, 4).Text)
            If t > 0 Then
                If GetIndexArStrSingle(strRecipe(), Recipe) = -1 Then
Aggiungi:
                        
                    t = t + 1
                    ReDim Preserve strRecipe(t)
                    strRecipe(t) = Recipe
                   
                    
                    
                
                End If
            Else
                GoTo Aggiungi
            End If
            
                        
        
        Next
    
    
    End If
    
 End With
 
 If t > 0 Then
    Call OrdinaArray(strRecipe)
    For i = 1 To UBound(strRecipe)
        ComboRecipe.AddItem strRecipe(i)
    Next
    ComboRecipe.ListIndex = 0
 End If
End Function

Private Function GetDataFromTabReport(ByVal Periodo As String)
Dim StartDate As Date
Dim i As Integer
Dim t As Integer
 With GrdBatch
        .Rows = 1
        DoEvents
       .AutoRedraw = False
       
       With dbTabReport
        .filter = ""
        .filter = "Visible=true and Finished=true"
        If .EOF Then
        
        Else
            For t = 1 To .RecordCount
            
            
                StartDate = IIf(IsNull(Trim(!StartDate)), "", Trim(!StartDate))
                
                If Periodo > StartDate Then GoTo cont:
                        
                With GrdBatch
                
                    
                    .Column(22).Width = 0
                    .Column(23).Width = 0
                    .Column(24).Width = 0
                    .Column(25).Width = 0
                    
                    If IsNull(Trim(dbTabReport!Lot)) Or IsNull(IsNull(Trim(dbTabReport!Code))) Then
                        GoTo cont:
                    End If
                    
                    .AddItem "", False
                
                     .Cell(.Rows - 1, 0).Text = .Rows - 1
                    
                           
                    
                
                   
                   .Cell(.Rows - 1, 1).Text = "  " & IIf(IsNull(Trim(dbTabReport!Lot)), "", Trim(dbTabReport!Lot))
                   .Cell(.Rows - 1, 2).Text = "  " & IIf(IsNull(Trim(dbTabReport!Code)), "", Trim(dbTabReport!Code))
                   .Cell(.Rows - 1, 3).Text = "  " & IIf(IsNull(Trim(dbTabReport!Description)), "", Trim(dbTabReport!Description))
                   .Cell(.Rows - 1, 4).Text = "  " & IIf(IsNull(Trim(dbTabReport!Recipe)), "", Trim(dbTabReport!Recipe))
                   .Cell(.Rows - 1, 5).Text = "  " & IIf(IsNull(Trim(dbTabReport!PREPWK)), "", Trim(dbTabReport!PREPWK))
                   .Cell(.Rows - 1, 6).Text = "  " & IIf(IsNull(Trim(dbTabReport!RangeMin)), "", Trim(dbTabReport!RangeMin))
                   .Cell(.Rows - 1, 7).Text = "  " & IIf(IsNull(Trim(dbTabReport!RangeMax)), "", Trim(dbTabReport!RangeMax))
                   .Cell(.Rows - 1, 8).Text = IIf(IsNull(Trim(dbTabReport!StartDate)), "", Trim(dbTabReport!StartDate))
                   .Cell(.Rows - 1, 9).Text = IIf(IsNull(Trim(dbTabReport!Exp)), "", Trim(dbTabReport!Exp))
                   .Cell(.Rows - 1, 10).Text = IIf(IsNull(Trim(dbTabReport!TestNumber)), "", Trim(dbTabReport!TestNumber))
                   .Cell(.Rows - 1, 11).Text = dbTabReport!Evaluation
                   .Cell(.Rows - 1, 12).Text = dbTabReport!Finished
                   
                   .Cell(.Rows - 1, 13).Text = IIf(IsNull(Trim(dbTabReport!Operator)), "", Trim(dbTabReport!Operator))
                   .Cell(.Rows - 1, 14).Text = IIf(IsNull(Trim(dbTabReport!Note)), "", Trim(dbTabReport!Note))
                   .Cell(.Rows - 1, 15).Text = dbTabReport!ID
                   .Cell(.Rows - 1, 16).Text = IIf(IsNull(Trim(dbTabReport!Nomefile)), "", Trim(dbTabReport!Nomefile))
                   
                   .Cell(.Rows - 1, 17).Text = IIf(IsNull(Trim(dbTabReport!NomeFileReport)), "", Trim(dbTabReport!NomeFileReport))
                   .Cell(.Rows - 1, 18).Text = IIf(IsNull(Trim(dbTabReport!NomeFileExcel)), "", Trim(dbTabReport!NomeFileExcel))
                    
                 
          
                    For i = 1 To .Cols - 1
                        .Cell(.Rows - 1, i).Alignment = cellCenterCenter
                        .Cell(.Rows - 1, i).ForeColor = vbColorDarkFont ' &H963D01
                       
                            .Cell(.Rows - 1, i).FontBold = True
                    
                        
                    Next
                
                    
                    .Cell(.Rows - 1, 1).Alignment = cellLeftCenter
                    .Cell(.Rows - 1, 2).Alignment = cellLeftCenter
                    .Cell(.Rows - 1, 3).Alignment = cellLeftCenter
                        
                        
                        
                      If dbTabReport!ExcelDone Then
                        .Cell(.Rows - 1, 21).BackColor = vbColorGreen
                        .Cell(.Rows - 1, 21).Text = "OK"
                        .Cell(.Rows - 1, 21).ForeColor = vbWhite
                        
                    End If
                    
                End With
            
            
            
            
            
            
            
cont:
                .MoveNext
            Next
            
              GrdBatch.Column(21).Width = 100
        End If
       
       
       End With
       
       
       
       
End With
End Function

Private Function FillGridBatch(ByVal bChiuso As Boolean, Optional Periodo As String) As Boolean
Dim i As Integer
Dim FileName As String
Dim rc As Boolean
Dim sString As String
Dim bClosedReadings As Boolean
Dim strQC As String

Dim StartDate As Date
Dim MyPeriodo As Date


    On Error GoTo ERR_FILL:
    
  MyPeriodo = SetData(Periodo)


Dim t As Integer

    USER_PATH = USER_TEMP_PATH
    If bChiuso Then
        USER_PATH = USER_DATA_PATH
    End If
    
    CloseSettingDataFile
 With GrdBatch
        .Rows = 1
        DoEvents
       .AutoRedraw = False
       
       If bChiuso Then
        sString = "and Finished=true"
        
       Else
            sString = "and Finished=false"
            
       End If
       
     
       
       With dbTabReport
        .filter = ""
        .filter = "Visible=true" & sString '
        If .EOF Then
        
        Else
            For t = 1 To .RecordCount
            
                FileName = IIf(IsNull(Trim(dbTabReport!Nomefile)), "", Trim(dbTabReport!Nomefile))
            

                StartDate = IIf(IsNull(Trim(!StartDate)), "", Trim(!StartDate))
                
                If MyPeriodo > StartDate Then GoTo cont:
                        
                With GrdBatch
                
                
                   If dbTabReport!Lot = "" Or dbTabReport!Code = "" Then
                        GoTo cont:
                    End If
                    
                    
                  
                    .AddItem "", False
                
                     .Cell(.Rows - 1, 0).Text = .Rows - 1
                    
                           
                    
                
                   
                   .Cell(.Rows - 1, 1).Text = "  " & IIf(IsNull(Trim(dbTabReport!Lot)), "", Trim(dbTabReport!Lot))
                   .Cell(.Rows - 1, 2).Text = "  " & IIf(IsNull(Trim(dbTabReport!Code)), "", Trim(dbTabReport!Code))
                   .Cell(.Rows - 1, 3).Text = "  " & IIf(IsNull(Trim(dbTabReport!Description)), "", Trim(dbTabReport!Description))
                   .Cell(.Rows - 1, 4).Text = "  " & IIf(IsNull(Trim(dbTabReport!Recipe)), "", Trim(dbTabReport!Recipe))
                   .Cell(.Rows - 1, 5).Text = "  " & IIf(IsNull(Trim(dbTabReport!PREPWK)), "", Trim(dbTabReport!PREPWK))
                   .Cell(.Rows - 1, 6).Text = "  " & IIf(IsNull(Trim(dbTabReport!RangeMin)), "", Trim(dbTabReport!RangeMin))
                   .Cell(.Rows - 1, 7).Text = "  " & IIf(IsNull(Trim(dbTabReport!RangeMax)), "", Trim(dbTabReport!RangeMax))
                   .Cell(.Rows - 1, 8).Text = IIf(IsNull(Trim(dbTabReport!StartDate)), "", Trim(dbTabReport!StartDate))
                   .Cell(.Rows - 1, 9).Text = IIf(IsNull(Trim(dbTabReport!Exp)), "", Trim(dbTabReport!Exp))
                   .Cell(.Rows - 1, 10).Text = IIf(IsNull(Trim(dbTabReport!TestNumber)), "", Trim(dbTabReport!TestNumber))
                   .Cell(.Rows - 1, 11).Text = dbTabReport!Evaluation
                   .Cell(.Rows - 1, 12).Text = dbTabReport!Finished
                   
                               
                    If bChiuso Then
                    Else
                        CloseSettingDataFile
                        bClosedReadings = GetSettingData(FileName, "Reading", "Closed", False)
                        
                        If bClosedReadings Then
                        .Cell(.Rows - 1, 10).BackColor = &HFBD9AB
                        End If

                        strQC = GetSettingData(FileName, "Evaluation QC", "ResultQC", "")
                        
                        Select Case strQC
                        
                        
                           
                            Case "Waiting"
                                  .Cell(.Rows - 1, 20).BackColor = &HA88030
                            Case "Failed"
                                .Cell(.Rows - 1, 20).BackColor = &H40C0&
                           Case "Passed"
                               .Cell(.Rows - 1, 20).BackColor = &H208040
                        End Select
                    End If
                  

                   .Cell(.Rows - 1, 13).Text = IIf(IsNull(Trim(dbTabReport!Operator)), "", Trim(dbTabReport!Operator))
                   .Cell(.Rows - 1, 14).Text = IIf(IsNull(Trim(dbTabReport!Note)), "", Trim(dbTabReport!Note))
                   .Cell(.Rows - 1, 15).Text = dbTabReport!ID
                   .Cell(.Rows - 1, 16).Text = IIf(IsNull(Trim(dbTabReport!Nomefile)), "", Trim(dbTabReport!Nomefile))
                   
                   .Cell(.Rows - 1, 17).Text = IIf(IsNull(Trim(dbTabReport!NomeFileReport)), "", Trim(dbTabReport!NomeFileReport))
                   .Cell(.Rows - 1, 18).Text = IIf(IsNull(Trim(dbTabReport!NomeFileExcel)), "", Trim(dbTabReport!NomeFileExcel))
                    
                    .Cell(.Rows - 1, 22).Text = IIf(IsNull(Trim(dbTabReport!ReagentLot)), "", Trim(dbTabReport!ReagentLot))
                    .Cell(.Rows - 1, 23).Text = IIf(IsNull(Trim(dbTabReport!ReagentCode)), "", Trim(dbTabReport!ReagentCode))
                    .Cell(.Rows - 1, 24).Text = IIf(IsNull(Trim(dbTabReport!ReagentLot2)), "", Trim(dbTabReport!ReagentLot2))
                    .Cell(.Rows - 1, 25).Text = IIf(IsNull(Trim(dbTabReport!ReagentCode2)), "", Trim(dbTabReport!ReagentCode2))
          
                    
                    For i = 1 To .Cols - 1
                        .Cell(.Rows - 1, i).Alignment = cellCenterCenter
                        .Cell(.Rows - 1, i).ForeColor = vbColorDarkFont ' &H963D01
                         If bClosedReadings Then
                            .Cell(.Rows - 1, i).FontBold = True
                        Else
                            .Cell(.Rows - 1, i).FontBold = False
                         End If
                        
                    Next
                
                
                                         
                    If dbTabReport!ExcelDone Then
                        .Cell(.Rows - 1, 21).BackColor = vbColorGreen
                        .Cell(.Rows - 1, 21).Text = "OK"
                        .Cell(.Rows - 1, 21).ForeColor = vbWhite
                        
                    End If
                    

                        
                                
                    .Cell(.Rows - 1, 1).Alignment = cellLeftCenter
                    .Cell(.Rows - 1, 2).Alignment = cellLeftCenter
                    .Cell(.Rows - 1, 3).Alignment = cellLeftCenter
                

                End With
            
            
cont:
                .MoveNext
            Next
            
                    
            GrdBatch.Column(21).Width = 100
            GrdBatch.Column(22).AutoFit
            GrdBatch.Column(23).AutoFit
            GrdBatch.Column(24).AutoFit
            GrdBatch.Column(25).AutoFit
            
            GrdBatch.AllowUserSort = True
            GrdBatch.AutoRedraw = True
            GrdBatch.Refresh
        
        End If
       
       
       End With
       
       
       
       
End With
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_FILL:
    MsgBox Err.Description
    
    Resume ERR_END
End Function

Private Function FillGridBatch_OLD(ByVal bChiuso As Boolean, Optional Periodo As String) As Boolean

Dim rc As Boolean
Dim Path As String
Dim MyPeriodo As String
Dim FSO As New Scripting.FileSystemObject
Dim Cartella As Folder
Dim FileGenerico As File
Dim DataFile As Date


Dim bFileOk As Boolean

Dim bFileCodeOk As Boolean
Dim bFileLotOk As Boolean

    rc = False
    
    
    MyPeriodo = SetData(Periodo)
    
    Path = IIf(bChiuso, USER_DATA_PATH, USER_TEMP_PATH)
    
    
    
       
       
    If bChiuso Then
        GetDataFromTabReport (MyPeriodo)
        GoTo ERR_END:
    End If
     
     
 
    
    Set Cartella = FSO.GetFolder(Path)
    
    With GrdBatch
        .Rows = 1
        DoEvents
       .AutoRedraw = False
       
        .Column(21).Width = 0

        For Each FileGenerico In Cartella.Files
            If InStr(FileGenerico.Name, USER_ESTENSIONE) Then
            
                ' controllo Lot e Code se li ho inseriti in Text1 allora filtro...
                Dim strCode As String
                strCode = UCase(FormatNomeFile(Text1(1)))
                'If Len(Text1(1)) > 0 Then
                   ' bFileCodeOk = InStr(UCase(FileGenerico.Name), strCode)
                'Else
                    bFileCodeOk = True
                'End If
                
               'If Len(Text1(0)) > 0 Then
                    'Dim strLot As String
                    ''strLot = UCase(FormatNomeFile(Text1(0)))
                   ' bFileLotOk = InStr(UCase(FileGenerico.Name), strLot)
               ' Else
                    bFileLotOk = True
               ' End If
                
                bFileOk = bFileLotOk And bFileCodeOk
                
                If bFileOk Then
                    If Periodo <> "" And Periodo <> "0" Then
                        ' filtro per data...
                        DataFile = FormatDateTime(FileGenerico.DateLastModified, vbShortDate)
                       
                        If DateDiff("d", DataFile, MyPeriodo) <= 0 Then
                            Call RiempiGridBatch(FileGenerico.Name, bChiuso)
                        End If
                        
                    Else
                        Call RiempiGridBatch(FileGenerico.Name, bChiuso)
                    End If
                    
                End If
                CloseSettingDataFile
            End If
        Next
        
        .Column(22).AutoFit
        .Column(23).AutoFit
        .Column(24).AutoFit
        .Column(25).AutoFit
        
        .AllowUserSort = True
        .AutoRedraw = True
        .Refresh
        
    End With
   
ERR_END:
    On Error GoTo 0
   

End Function

Private Function RiempiGridBatch(ByVal FileName As String, Optional ByVal bChiuso As Boolean)
Dim i As Integer
Dim Path As String
Dim NumTest As String
Dim Rows As Long
Dim sLot As String
Dim sCode As String
Dim sRecipe As String
Dim bClosedReadings As Boolean
Dim strQC As String
Dim sRangeMin As String
Dim sRangeMax As String
On Error GoTo ERR_FILL

    Path = IIf(bChiuso, USER_DATA_PATH, USER_TEMP_PATH)
    
       

    NumTest = GetSettingData(FileName, "Reading QC", "Grd2 Rows", 1, Path) - 1
    
    If NumTest = 0 Then
        Rows = 1
       Call CheckRows(Rows, FileName, Path)
       NumTest = Rows - 1
    End If
     
    If IndexProcedura = 2 Or IndexProcedura = 3 Then
        If NumTest = "" Or NumTest = "0" Then
            Exit Function
        End If
    End If
                           
                           
        sLot = Trim(GetSettingData(FileName, "Information QC", "Text10", "", Path))
       sCode = Trim(GetSettingData(FileName, "Information QC", "Text11", "", Path))
       sRecipe = Trim(GetSettingData(FileName, "Information QC", "Text15", "", Path))
       sRangeMin = GetSettingData(FileName, "Code Information", "RangeMin", "", Path)
       sRangeMax = GetSettingData(FileName, "Code Information", "RangeMax", "", Path)
       
       If sLot = "" And sCode = "" Then
        'Kill Path & FileName
        DoEvents
        Exit Function
       End If
                      
       If ComboRecipe.ListIndex > 0 Then
        
            If ComboRecipe <> sRecipe Then Exit Function
       
       End If
                           
                           
           
                           
    With GrdBatch
         .AutoRedraw = False
        .Rows = .Rows + 1
        .Cell(.Rows - 1, 0).Text = .Rows - 1
        

        .Cell(.Rows - 1, 1).Text = "  " & sLot
        .Cell(.Rows - 1, 2).Text = "  " & sCode
        .Cell(.Rows - 1, 3).Text = "  " & GetSettingData(FileName, "Information QC", "Text12", "", Path)
        
        .Cell(.Rows - 1, 4).Text = "  " & GetSettingData(FileName, "Information QC", "Text15", "", Path)
        .Cell(.Rows - 1, 5).Text = "  " & GetSettingData(FileName, "Information QC", "Text121", "", Path)
              
        
        .Cell(.Rows - 1, 6).Text = "  " & sRangeMin
        .Cell(.Rows - 1, 7).Text = "  " & sRangeMax

        
        .Cell(.Rows - 1, 8).Text = GetSettingData(FileName, "File Information", "Creation Date", "", Path)
        
        .Cell(.Rows - 1, 9).Text = GetSettingData(FileName, "Information QC", "Text13", "", Path)
        .Cell(.Rows - 1, 10).Text = GetSettingData(FileName, "Reading QC", "Grd2 Rows", 1, Path) - 1
        
        If bChiuso Then
        Else
            bClosedReadings = GetSettingData(FileName, "Reading", "Closed", False)
            
            If bClosedReadings Then
            .Cell(.Rows - 1, 10).BackColor = &HFBD9AB
            End If
            
          
            
        
        End If
            
            strQC = GetSettingData(FileName, "Evaluation QC", "ResultQC", "")
    
            Select Case strQC
               
                Case "Waiting"
                    .Cell(.Rows - 1, 20).BackColor = &HA88030
                Case "Failed"
                    .Cell(.Rows - 1, 20).BackColor = &H40C0&
                Case "Passed"
                    .Cell(.Rows - 1, 20).BackColor = &H208040
            End Select
            
        .Cell(.Rows - 1, 13).Text = GetSettingData(FileName, "Information QC", "Operator", "", Path)
        .Cell(.Rows - 1, 14).Text = GetSettingData(FileName, "Information QC", "Text130", "", Path)
        .Cell(.Rows - 1, 15).Text = "" ' GetSettingData(Filename, "Information QC", "Creation Date", "",Path)
        .Cell(.Rows - 1, 16).Text = FileName
      '  .Cell(.Rows - 1, 19).Text = GetCodeID(sCode, sRangeMin, sRangeMax)
        
        .Cell(.Rows - 1, 22).Text = GetSettingData(SettingName, "Information QC", "strReagentLot", "")
        .Cell(.Rows - 1, 23).Text = GetSettingData(SettingName, "Information QC", "strReagentCode", "")
        .Cell(.Rows - 1, 24).Text = GetSettingData(SettingName, "Information QC", "strReagentLot2", "")
        .Cell(.Rows - 1, 25).Text = GetSettingData(SettingName, "Information QC", "strReagentCode2", "")

        
    For i = 1 To .Cols - 1
        .Cell(.Rows - 1, i).Alignment = cellCenterCenter
        .Cell(.Rows - 1, i).ForeColor = vbColorDarkFont ' &H963D01
         If bClosedReadings Then
            .Cell(.Rows - 1, i).FontBold = True
         End If
        
    Next

    
    .Cell(.Rows - 1, 1).Alignment = cellLeftCenter
    .Cell(.Rows - 1, 2).Alignment = cellLeftCenter
    .Cell(.Rows - 1, 3).Alignment = cellLeftCenter
    

    End With
ERR_END:
 
    Exit Function
ERR_FILL:
    MsgBox Err.Description
    Resume Next
    
End Function

Public Function GetCodeID(ByVal strCode As String, ByVal RangeMin As String, ByVal RangeMax As String) As Long
   GetCodeID = 0
   With dbTabCode
       
        .filter = ""
        .filter = "Code='" & strCode & "' and RangeMin='" & RangeMin & "' and RangeMax='" & RangeMax & "'"
        If .EOF Then
        Else
            GetCodeID = !ID
        End If
    End With
End Function

Private Function RefreshTableForm(ByVal bValue As Boolean)
DefaultMenuLabel(8).Visible = bValue
DefaultMenuLabel(8).ZOrder
Image3(8).Visible = bValue
Label2(10).Visible = bValue
If bValue Then ShowSTDInfo False
End Function


Private Function CheckPrimoAvvio() As Boolean
Dim rc As Boolean

    rc = True
    With dbTabCode
    
        .filter = ""
        If .EOF Then
            rc = False
            PicIntro.ZOrder
            PicIntro2.Visible = True
        Else
            PicIntro2.Visible = False
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








Private Function ScanQRCodeQC() As Boolean

MessageInfoTime = 2000

txQRCode = ""
txQRCode.Top = 0
txQRCode.Visible = True
txQRCode.SetFocus

UserQrCode = iQRCodeTypeClean
iQRCodeType = iQRCodeTypeClean

PopupMessage 2, "Scan QRCode....", , , "Preparation | QC"


End Function



Private Sub txQRCode_KeyPress(KeyAscii As Integer)
Dim rc As Boolean

On Error GoTo ERR_QR:



If KeyAscii = 13 Then


    Image1_Click
    
    DoEvents
    If txQRCode = "" Then Exit Sub
   
    If GetQRCodeFromString(Trim(txQRCode), UserQrCode) Then
    
        iQRCodeType = UserQrCode
        
    
        If UserQrCode.Code = "" Or UserQrCode.Lot = "" Then
            PopupMessage 2, "Please Rescan QRCode...", , , "QRCode Reader"
            txQRCode = ""
            
            Exit Sub
        End If
        
        Text1(0) = UserQrCode.Lot
        Text1(1) = UserQrCode.Code
        
        
        
        If QRCodeToTabReport(UserQrCode, False) Then
            
            ' ho Giò il QC in tabella....
            
            
            
           ' If SearchQCInTab(UserQrCode, GrdBatch) Then
           
           
                PicMenu_Click 1
              
                SelectedCode = Text1(1)
                SelectedLot = Text1(0)
                
                
                Call OpenAnyLots(0, UserQrCode.FileName)

               'Call OpenReadingQC
               
            
           ' End If
            
        Else
            ' apro QC nuovo !

            PicMenu_Click 0
            If F_MsgBox.DoShow("Open new Information Lot?", UserQrCode.Code & " | " & UserQrCode.Lot) Then

                PicMenu_Click 0
            
            End If
            
        
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
    PopupMessage 2, Err.Description & vbCrLf & "Please repeat reading...", , , "QR Code Reader"
    Resume ERR_END:


End Sub



Private Function OpenReadingQC()

' vai a Readings...
PicMenu_Click 1

With UserQrCode
    SettingName = .FileName
    SelectedCodeID = 0
    SelectedCode = .Code
    SelectedLot = .Lot
End With

    If FileExists(USER_TEMP_PATH & SettingName) Then
        USER_PATH = USER_TEMP_PATH
    ElseIf FileExists(USER_DATA_PATH & SettingName) Then
        USER_PATH = USER_DATA_PATH
         PopupMessage 2, "Lot : " & SelectedLot & vbCrLf & "Code : " & SelectedCode & vbCrLf & "This Lot Is Closed..."
        bSearchClosedLot = True
    Else

    End If


 GrdBatch_DblClick

End Function


