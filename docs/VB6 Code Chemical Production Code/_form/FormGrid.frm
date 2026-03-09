VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form FormProductionDatabaseHistory 
   BackColor       =   &H00E0E0E0&
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
   Icon            =   "FormGrid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   Picture         =   "FormGrid.frx":33E2
   ScaleHeight     =   12000
   ScaleWidth      =   19200
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
      Left            =   13920
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   9840
      Width           =   2775
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H004D3B37&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   19215
      TabIndex        =   34
      Top             =   10800
      Width           =   19215
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00846623&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   10680
         MouseIcon       =   "FormGrid.frx":3724
         ScaleHeight     =   855
         ScaleWidth      =   4455
         TabIndex        =   35
         Top             =   160
         Visible         =   0   'False
         Width           =   4455
         Begin VB.Image DefaultMenu 
            Height          =   480
            Index           =   8
            Left            =   2040
            MouseIcon       =   "FormGrid.frx":3A2E
            MousePointer    =   99  'Custom
            Picture         =   "FormGrid.frx":3D38
            Top             =   200
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   8
         Left            =   11640
         TabIndex        =   46
         Top             =   -240
         Width           =   2655
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   0
         Left            =   9360
         MouseIcon       =   "FormGrid.frx":711A
         MousePointer    =   99  'Custom
         Picture         =   "FormGrid.frx":7424
         Top             =   240
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   1
         Left            =   4560
         MouseIcon       =   "FormGrid.frx":A806
         MousePointer    =   99  'Custom
         Picture         =   "FormGrid.frx":AB10
         Top             =   240
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   2
         Left            =   600
         MouseIcon       =   "FormGrid.frx":DEF2
         MousePointer    =   99  'Custom
         Picture         =   "FormGrid.frx":E1FC
         Top             =   240
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   3
         Left            =   15600
         MouseIcon       =   "FormGrid.frx":115DE
         MousePointer    =   99  'Custom
         Picture         =   "FormGrid.frx":118E8
         Top             =   240
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   4
         Left            =   18240
         MouseIcon       =   "FormGrid.frx":14CCA
         MousePointer    =   99  'Custom
         Picture         =   "FormGrid.frx":14FD4
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excel folder"
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
         Left            =   450
         MouseIcon       =   "FormGrid.frx":183B6
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   795
         Width           =   960
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apply filter"
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
         Left            =   4335
         MouseIcon       =   "FormGrid.frx":186C0
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   795
         Width           =   930
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Database"
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
         Left            =   9015
         MouseIcon       =   "FormGrid.frx":189CA
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   795
         Width           =   1140
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
         MouseIcon       =   "FormGrid.frx":18CD4
         MousePointer    =   99  'Custom
         TabIndex        =   42
         Top             =   795
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
         Left            =   17760
         MouseIcon       =   "FormGrid.frx":18FDE
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   795
         Width           =   1200
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   4
         Left            =   17400
         MouseIcon       =   "FormGrid.frx":192E8
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   -120
         Width           =   1815
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   3
         Left            =   15000
         MouseIcon       =   "FormGrid.frx":195F2
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   -120
         Width           =   2055
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   0
         Left            =   8640
         MouseIcon       =   "FormGrid.frx":198FC
         MousePointer    =   99  'Custom
         TabIndex        =   38
         Top             =   -120
         Width           =   2055
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   1
         Left            =   3720
         MouseIcon       =   "FormGrid.frx":19C06
         MousePointer    =   99  'Custom
         TabIndex        =   37
         Top             =   10680
         Width           =   2055
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   2
         Left            =   0
         MouseIcon       =   "FormGrid.frx":19F10
         MousePointer    =   99  'Custom
         TabIndex        =   36
         Top             =   -120
         Width           =   2055
      End
   End
   Begin VB.PictureBox PicMainMenu 
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
      Height          =   1095
      Index           =   4
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   19215
      TabIndex        =   16
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   5
         Left            =   9600
         MouseIcon       =   "FormGrid.frx":1A21A
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   5
            Left            =   720
            MouseIcon       =   "FormGrid.frx":1A524
            MousePointer    =   99  'Custom
            Picture         =   "FormGrid.frx":1A82E
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Production Specifics"
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
            Left            =   0
            MouseIcon       =   "FormGrid.frx":1D220
            MousePointer    =   99  'Custom
            TabIndex        =   49
            Top             =   720
            Width           =   1890
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   4
         Left            =   5760
         MouseIcon       =   "FormGrid.frx":1D52A
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   27
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Export Table"
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
            MouseIcon       =   "FormGrid.frx":1D834
            MousePointer    =   99  'Custom
            TabIndex        =   28
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   4
            Left            =   720
            MouseIcon       =   "FormGrid.frx":1DB3E
            MousePointer    =   99  'Custom
            Picture         =   "FormGrid.frx":1DE48
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MouseIcon       =   "FormGrid.frx":2122A
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   23
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "FormGrid.frx":21534
            MousePointer    =   99  'Custom
            Picture         =   "FormGrid.frx":2183E
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delete filter"
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
            Left            =   525
            MouseIcon       =   "FormGrid.frx":24C20
            MousePointer    =   99  'Custom
            TabIndex        =   24
            Top             =   720
            Width           =   960
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   1920
         MouseIcon       =   "FormGrid.frx":24F2A
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   21
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   720
            MouseIcon       =   "FormGrid.frx":25234
            MousePointer    =   99  'Custom
            Picture         =   "FormGrid.frx":2553E
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Export Prod"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0070B070&
            Height          =   225
            Index           =   1
            Left            =   420
            MouseIcon       =   "FormGrid.frx":28920
            MousePointer    =   99  'Custom
            TabIndex        =   22
            Top             =   720
            Width           =   990
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   3840
         MouseIcon       =   "FormGrid.frx":28C2A
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   19
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   720
            MouseIcon       =   "FormGrid.frx":28F34
            MousePointer    =   99  'Custom
            Picture         =   "FormGrid.frx":2923E
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Date filter"
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
            MouseIcon       =   "FormGrid.frx":2C620
            MousePointer    =   99  'Custom
            TabIndex        =   20
            Top             =   720
            Width           =   1890
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   3
         Left            =   7680
         MouseIcon       =   "FormGrid.frx":2C92A
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   17
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   720
            MouseIcon       =   "FormGrid.frx":2CC34
            MousePointer    =   99  'Custom
            Picture         =   "FormGrid.frx":2CF3E
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Delete"
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
            MouseIcon       =   "FormGrid.frx":30320
            MousePointer    =   99  'Custom
            TabIndex        =   18
            Top             =   720
            Width           =   1890
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Production Database"
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
         Left            =   14475
         TabIndex        =   25
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   1920
      Top             =   10200
   End
   Begin VB.Timer Timer3 
      Interval        =   250
      Left            =   960
      Top             =   10320
   End
   Begin VB.ComboBox CmbVisual 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   14520
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1860
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1860
      Width           =   4095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00008000&
      Caption         =   "Frame3"
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   17400
      TabIndex        =   2
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   1440
      Top             =   10200
   End
   Begin FlexCell.Grid GridDatabase 
      Height          =   6975
      Left            =   480
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2640
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   12303
      AllowUserSort   =   -1  'True
      Appearance      =   0
      BackColor1      =   15790320
      BackColor2      =   15790320
      BackColorActiveCellSel=   15790320
      BackColorBkg    =   15790320
      BackColorFixed  =   15790320
      BackColorFixedSel=   15790320
      BackColorScrollBar=   15592423
      BorderColor     =   15790320
      CellBorderColor =   15790320
      CellBorderColorFixed=   15790320
      Cols            =   5
      DefaultFontName =   "Calibri"
      DefaultFontSize =   9.75
      BoldFixedCell   =   0   'False
      DisplayRowIndex =   -1  'True
      DrawMode        =   1
      DefaultRowHeight=   20
      FixedRowColStyle=   0
      ForeColorFixed  =   6571523
      GridColor       =   15790320
      ReadOnly        =   -1  'True
      Rows            =   1
      ScrollBarStyle  =   0
      SelectionMode   =   1
      MultiSelect     =   0   'False
   End
   Begin ChemicalProduction.ctlCalendar ctlCalendar1 
      Height          =   6960
      Left            =   11760
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   12277
      ShowLastMonthButton=   -1  'True
      ShowNextMonthButton=   -1  'True
      ShowLastMonthDays=   -1  'True
      ShowNextMonthDays=   -1  'True
      ShowTodayLabel  =   -1  'True
      ColorBackgroundHeader=   6571523
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
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDay {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   20.25
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
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FlexCell.Grid Grid2 
      Height          =   6975
      Left            =   480
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   12303
      AllowUserSort   =   -1  'True
      Appearance      =   0
      BackColor1      =   15790320
      BackColor2      =   15790320
      BackColorActiveCellSel=   15790320
      BackColorBkg    =   15790320
      BackColorFixed  =   15790320
      BackColorFixedSel=   15790320
      BackColorScrollBar=   15592423
      BorderColor     =   15790320
      CellBorderColor =   15790320
      CellBorderColorFixed=   15790320
      Cols            =   5
      DefaultFontName =   "Calibri"
      DefaultFontSize =   9.75
      BoldFixedCell   =   0   'False
      DisplayRowIndex =   -1  'True
      DrawMode        =   1
      DefaultRowHeight=   20
      FixedRowColStyle=   0
      ForeColorFixed  =   6571523
      GridColor       =   15790320
      ReadOnly        =   -1  'True
      Rows            =   1
      ScrollBarStyle  =   0
      SelectionMode   =   1
      MultiSelect     =   0   'False
   End
   Begin VB.Frame frSpecifics 
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   480
      TabIndex        =   51
      Top             =   1320
      Visible         =   0   'False
      Width           =   18495
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Production Specifics"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   0
         TabIndex        =   52
         Top             =   240
         Width           =   18495
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00606060&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2040
      ScaleHeight     =   2895
      ScaleWidth      =   13455
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   13455
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel filter"
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
         Left            =   12195
         MouseIcon       =   "FormGrid.frx":3062A
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   2580
         Width           =   1020
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apply filter"
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
         Left            =   300
         MouseIcon       =   "FormGrid.frx":30934
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   2580
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Date Range and Apply Filter"
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
         Height          =   270
         Index           =   2
         Left            =   5040
         TabIndex        =   26
         Top             =   2160
         Width           =   3450
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   11
         Left            =   480
         MouseIcon       =   "FormGrid.frx":30C3E
         MousePointer    =   99  'Custom
         Picture         =   "FormGrid.frx":30F48
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   10
         Left            =   12480
         MouseIcon       =   "FormGrid.frx":3432A
         MousePointer    =   99  'Custom
         Picture         =   "FormGrid.frx":34634
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lbDataFiltro 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   600
         Index           =   1
         Left            =   7560
         MouseIcon       =   "FormGrid.frx":37A16
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label lbDataFiltro 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   600
         Index           =   0
         Left            =   1560
         MouseIcon       =   "FormGrid.frx":37D20
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "END DATE"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   9030
         TabIndex        =   9
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "START DATE"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   2895
         TabIndex        =   8
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   11
         Left            =   0
         TabIndex        =   13
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1095
         Index           =   10
         Left            =   11760
         TabIndex        =   12
         Top             =   1800
         Width           =   1695
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Line"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   3
      Left            =   12510
      MouseIcon       =   "FormGrid.frx":3802A
      MousePointer    =   99  'Custom
      TabIndex        =   54
      Top             =   9840
      Width           =   1230
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Production"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   2
      Left            =   5280
      MouseIcon       =   "FormGrid.frx":38334
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   9840
      Width           =   1620
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Closed Production"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   1
      Left            =   2760
      MouseIcon       =   "FormGrid.frx":3863E
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   9840
      Width           =   2130
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Production"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   0
      Left            =   360
      MouseIcon       =   "FormGrid.frx":38948
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   9840
      Width           =   2010
   End
   Begin VB.Label Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hi340504333"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   0
      Left            =   4680
      MouseIcon       =   "FormGrid.frx":38C52
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1875
      Width           =   9585
   End
   Begin VB.Label lbColonneGrid 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show Less"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   17430
      MouseIcon       =   "FormGrid.frx":38F5C
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   9840
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIOD"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   1
      Left            =   14520
      TabIndex        =   6
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODE"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   0
      Left            =   4695
      TabIndex        =   5
      Top             =   1440
      Width           =   9570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FILTER"
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   5
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   915
   End
End
Attribute VB_Name = "FormProductionDatabaseHistory"
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
Private ExcelFilename As String

Private IndexTabella As Integer
Private MaxIndex As Integer
Private dIndexProcedura As Integer
Private m_Procedura As Boolean
Private SelectedDBCode As String

Private bHilight As Boolean
Private bFiltroIntervalloDate As Boolean
Private DataIndex As Integer
Private m_rc As Boolean

Private MyDa As String
Private MyA As String
Private MyPeriodo As String
Private ProductionID As Long
Private ProductionFileName As String
Private bMeanValue As Boolean
Private bClosedProduction As Boolean
Private MyIndexRecord As Integer
Private lRow As Long
Private IndexOpenClosedLot As Integer


Private Sub cmbLineProduction_Click()
GlobalSearch
End Sub

Private Sub Form_Load()
If Screen.Width - Me.Width > 1000 And bFullScreen Then
    Me.WindowState = 2
    'Me.Picture = LoadPicture()
   
End If
IndexOpenClosedLot = 2
ChangeLabelLots

Call SetDatabaseProductionTable(GridDatabase)
Call SetDatabaseProductionHistory(Grid2)
Call SetLine(cmbLineProduction, True)

    

    ProductionID = 0
    ProductionFileName = ""
    MyIndexRecord = 3
    bMeanValue = False
    bClosedProduction = False
    
    RiempiCombo
    
    SaveSizes
End Sub


Public Function DoShow(Optional ByVal IndexProcedura As Integer = 0, Optional ByVal Code As String = "") As Boolean
Dim i As Integer
    On Error GoTo ERR_SHOW
    
    m_rc = False
    mOk
    dIndexProcedura = IndexProcedura
    
    SelectedDBCode = Code
    

    

    Me.Show vbModal
    
    

    
    If m_rc = True Then
        Code = SelectedDBCode
    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    Resume ERR_END
End Function



Private Sub GridDatabase_BeforeUserSort(ByVal Col As Long)
lRow = 0
End Sub

Private Sub GridDatabase_Click()
ctlCalendar1.Visible = False

End Sub


Private Sub GridDatabase_DblClick()
If lRow = 0 Then Exit Sub
    If ProductionFileName <> "" Then
    
        If FileExists(USER_PRODUCTION_PATH & ProductionFileName) Then
            USER_PATH = USER_PRODUCTION_PATH
        ElseIf FileExists(USER_PRODUCTION_PATH & "data\" & ProductionFileName) Then
            USER_PATH = USER_PRODUCTION_PATH & "data\"
        
        Else
        
            Exit Sub
            
        End If
        
        frmProduction.Left = Me.Left
        frmProduction.Top = Me.Top
        frmProduction.WindowState = Me.WindowState
        Call frmProduction.DoShow(ProductionFileName, ProductionID)
    End If


End Sub


Private Sub Image3_Click(Index As Integer)
Dim MyexcelName As String


    Select Case Index
        Case 0
            'cmbLineProduction.ListIndex = 0
            Combo1.ListIndex = 0
            Combo1_Click
             IntervalloDate False
            GlobalSearch
        Case 1
              ' stampa lista
            Dim ExcelFilename As String
            Dim ProdWeek As String
           
            Dim ProdWeekFileArray() As ProdFileArray
            Dim ProdWeekFileArrayClean() As ProdFileArray

            Dim i As Integer
            Dim ProdCount As Integer
            
           
             ProdWeekFileArray = ProdWeekFileArrayClean
            ' export LOT Excel
            If GridDatabase.Rows > 1 Then 'And ProdWeek <> "" Then
            
                ReDim ProdWeekFileArray(GridDatabase.Rows - 1)
                ProdCount = GridDatabase.Rows - 1
                
                For i = 1 To GridDatabase.Rows - 1
                    
                    ProdWeekFileArray(i).Text = Trim(GridDatabase.Cell(i, 13).Text)
                    ProdWeekFileArray(i).ProdDate = Trim(GridDatabase.Cell(i, 2).Text)
                Next
         
                
                 iProductionExportExcel = ProductionExportExcelClean
                
                With iProductionExportExcel
                    .FirstDate = GridDatabase.Cell(1, 2).Text
                    .LastDate = GridDatabase.Cell(GridDatabase.Rows - 1, 2).Text
                    .ProdCount = ProdCount
                    .ProdLine = cmbLineProduction
                    
                    ProdWeek = PreparationWeek(CDate(.FirstDate))
                    
                    .WeekProd = ProdWeek
                End With
         
            
                 ExcelFilename = "PROD_WK_" & FormatNomeFile(Trim(ProdWeek) & ".LINE_" & Trim(cmbLineProduction)) & ".xls"
               
               
               
                If Len(ExcelFilename) > 28 Then ExcelFilename = Left$(ExcelFilename, 27)
                PopupMessage 2, "Exporting data to Excel : please wait...." & vbCrLf & ExcelFilename
                Call EsportaProductionWeekExcel(ProdWeekFileArray, ExcelFilename, iProductionExportExcel)
            End If
            




        Case 2
            Call OpenIntervalloDate
        Case 3
             ' cancella campione
             If CheckPrivilege(2) Then
                If InStr(Text1(0), Combo1) Then
                    GoTo err_canc:
                Else
                    If Text1(0) = "" Then
err_canc:
                    Else
                        Call CancellaTab
                    End If
                End If
            End If
        Case 4
            Dim ExcelName As String
            ExcelName = FormatNomeFile("Production List." & FormatDataLAT(Now()))
            
            If F_InputBox.DoShow("Please Set Excel Name", "Production To Excel", , , , ExcelName) Then
    
                ExcelName = USER_DESKTOP & "\" & FormatNomeFile(ExcelName) & ".xls"
                
                
                GridDatabase.ExportToExcel ExcelName, True, True
                MessageInfoTime = 2500
                PopupMessage 2, "Excel Done..." & vbCrLf & ExcelName
            End If
        Case 5
            ' View Production History
            Call ViewProductionHistory

    End Select
End Sub

Private Sub ViewProductionHistory()

    Grid2.ZOrder

    If Grid2.Visible = True Then
        frSpecifics.Visible = False
        Me.BackColor = &HE0E0E0
        Label2(5) = "Production Specifics"
    Else
    
        frSpecifics.Visible = True
        Me.BackColor = &HA0A0A0
        Call FillProductionSpecifics
        Label2(5) = "Close Specifics"
    
    End If
    
    Grid2.Visible = Not (Grid2.Visible)

End Sub

Private Sub Label2_Click(Index As Integer)
Image3_Click Index
End Sub

Private Sub Label4_Click(Index As Integer)
IndexOpenClosedLot = Index
ChangeLabelLots
GlobalSearch
End Sub

Private Sub lbColonneGrid_Click()
Dim rc As Boolean
rc = IIf(GridDatabase.Column(10).Width = 0, True, False)
SaveSetting App.Title, Me.Name, "Visualizza Colonne", Not (rc)
VisualizzaColonne rc

End Sub

Private Sub VisualizzaColonne(ByVal rc As Boolean)
With GridDatabase
    .AutoRedraw = False
    '.Column(4).Width = IIf(rc, 120, 0)
   ' .Column(5).Width = IIf(rc, 120, 0)
    .Column(6).Width = IIf(rc, 120, 0)
    .Column(7).Width = IIf(rc, 120, 0)
    .Column(8).Width = IIf(rc, 120, 0)
    .Column(9).Width = IIf(rc, 120, 0)
    .Column(10).Width = IIf(rc, 120, 0)
    .Column(11).Width = IIf(rc, 120, 0)
    
    
    .Refresh
    .AutoRedraw = True
End With
lbColonneGrid.Caption = IIf(rc, ("Hide Columns"), ("Show Columns"))
End Sub
Private Sub lbDataFiltro_Change(Index As Integer)
    MessageInfoTime = 2000
    Select Case Index
        Case 0
            If Len(lbDataFiltro(1)) > 0 And Len(lbDataFiltro(0)) > 0 Then
                ctlCalendar1.Visible = False
                If CDate(lbDataFiltro(0)) > CDate(lbDataFiltro(1)) Then
                    PopupMessage 2, ("Warning: Check Date..."), , True, ("Filter"), DefaultMenu(9)
                    lbDataFiltro(0) = ""
                End If
                
            End If
        Case 1
            If Len(lbDataFiltro(0)) > 0 And Len(lbDataFiltro(1)) > 0 Then
                ctlCalendar1.Visible = False
                If CDate(lbDataFiltro(0)) > CDate(lbDataFiltro(1)) Then
                    
                    PopupMessage 2, ("Warning: Check Date..."), , True, ("Filter"), DefaultMenu(9)
                    lbDataFiltro(1) = ""
                End If
                                
                
            End If
    
    
    End Select
    If Index = 0 Or Index = 1 Then
        lbDataFiltro(Index).BackColor = IIf(Len(lbDataFiltro(Index)) > 0, vbWhite, &HE0E0E0)
    End If
End Sub


Private Sub Form_Resize()
ResizeControls
ResizeTab
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
        ElseIf TypeOf ctl Is Timer Then
        

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
        
        ElseIf TypeOf ctl Is Grid Then
           ctl.Left = x_scale * .Left
            ctl.Top = y_scale * .Top
            ctl.Width = x_scale * .Width
            ctl.Height = y_scale * .Height

              
        ElseIf TypeOf ctl Is Menu Then
        ElseIf TypeOf ctl Is Timer Then
        Else
            ctl.Left = x_scale * .Left
            'MsgBox (TypeName(ctl))
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
End Sub

Private Sub CmbVisual_Click()
SaveSetting App.Title, Me.Name, "Filtro Data", CmbVisual.ListIndex
PicMenu(5).Visible = False
MyDa = ""
MyA = ""
If CmbVisual.ListIndex < 0 Then
    MyPeriodo = 1
Else
    MyPeriodo = GetDateCombo(CmbVisual.ListIndex)

End If


If Me.Visible Then GlobalSearch
End Sub


Private Sub Combo1_Click()
'If Me.Visible ThenPicMenu
PicMenu(5).Visible = False
GridDatabase.Cell(0, 0).SetFocus
    
    Text1(0) = (" - Search") & Combo1 & " - "
    Label1(0) = UCase(Combo1)
    'If Me.Visible Then Text1(0).SetFocus
    SaveSetting App.Title, Me.Name, "Filtro Combo", Combo1.ListIndex
    
   ' PicMenu(1).Visible = IIf(InStr(LCase(Combo1), "week"), True, False)
    
'End If
End Sub



Private Sub ctlCalendar1_DateClicked(inputDate As Date)
lbDataFiltro(DataIndex) = inputDate
ctlCalendar1.Visible = False
End Sub

Private Sub lbDataFiltro_Click(Index As Integer)
If Index = 0 Or Index = 1 Then
    ctlCalendar1.Left = Picture2.Left + lbDataFiltro(Index).Left + (lbDataFiltro(Index).Width / 2 - ctlCalendar1.Width / 2)
    ctlCalendar1.ZOrder
    ctlCalendar1.Visible = True
    DataIndex = Index
Else
    OpenIntervalloDate
End If

End Sub

Private Sub DefaultMenu_Click(Index As Integer)
Select Case Index
    Case 0
        Unload Me
    Case 2
        ' Open Report folder
        OpenWithDefault (USER_EXCEL_PATH)
      
    Case 1
        ' filtro
        
            Call GlobalSearch
    Case 4
        ' avanti di 10
        Call ScorriTabella(True)
    Case 3
        ' indietro di 10
        Call ScorriTabella(False)
    
    
    
    Case 5
   
    Case 6
      
    Case 7
        ' aggiungi campione
        GetFormDatiCampione
        
    Case 8
        ' play
        'PopupMessage 2, "Campione Importato in procedura", , , FormName
        SelectedDBCode = Text1(0)
        m_Procedura = True
        m_rc = True
        Unload Me
    Case 9
       
    Case 10
    
        IntervalloDate False
      
    Case 11
        If lbDataFiltro(0) = "" Or lbDataFiltro(1) = "" Then
            PopupMessage 2, ("Select Start and End Filter Date"), , , ("Filter"), DefaultMenu(9)
        Else
        
            IntervalloDate True
        End If
    Case 12

        
End Select
End Sub



Private Sub DefaultMenuLabel_Click(Index As Integer)
DefaultMenu_Click Index
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub



Private Sub RiempiCombo()

    With CmbVisual
        .Clear
        .AddItem "Day"
        .AddItem "Month"
        .AddItem "Year"
        .AddItem "Archive"
        .ListIndex = GetSetting(App.Title, Me.Name, "Filtro Data", 0)
    End With

    With Combo1
        .Clear
        '.AddItem " " & ("Line")
        .AddItem " " & ("Hanna Code")
        .AddItem " " & ("Recipe")
        .AddItem " " & ("Week")
        .AddItem " " & ("Prep.Lot")
        .AddItem " " & ("SFG Lot")
        
        Combo1.ListIndex = 0 ' GetSetting(App.Title, Me.Name, "Filtro Combo", 0)
    End With
    
End Sub



Private Function RiempiGrid(ByRef Grd As Grid)
Dim i As Integer
Dim t As Integer

    ' --------------------------------------
    '
    '  filtra TabReport e riempi Tabella
    '
    ' --------------------------------------
    MyDa = lbDataFiltro(0)
    MyA = lbDataFiltro(1)
    Call FillTabellaTutte(Grd, MyPeriodo, Combo1, Text1(0), MyDa, MyA, IndexOpenClosedLot, cmbLineProduction)
    IndexTabella = 1
    MaxIndex = IIf(Int((GridDatabase.Rows - 1) / 10) < (GridDatabase.Rows - 1) / 10, (Int((GridDatabase.Rows - 1) / 10)) + 1, Int((GridDatabase.Rows - 1) / 10))
    If MaxIndex = 0 Then MaxIndex = 1
    

End Function







Private Sub ImageTAV_Click(Index As Integer)
Select Case Index
        Case 0
            Unload Me
        
        Case 2
        

End Select
End Sub

Private Sub ResizeTab()
Dim rc As Boolean
  rc = (GetSetting(App.Title, Me.Name, "Visualizza Colonne", True))
  VisualizzaColonne Not (rc)
  Grid2.Move GridDatabase.Left, GridDatabase.Top, GridDatabase.Width, GridDatabase.Height
End Sub



Private Sub GridDatabase_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim NumCol As Integer

ProductionID = 0
ProductionFileName = ""
bMeanValue = False
bClosedProduction = False
lRow = FirstRow
ExcelFilename = ""
PicMenu(4).Visible = False

If FirstRow > 0 Then

    MyLbHelpCount = 0
  
    ' RECIPE + CODE + RANGE + LOT + PREPWK
    ExcelFilename = Trim(GridDatabase.Cell(lRow, 5).Text) & "_" & Trim(GridDatabase.Cell(lRow, 2).Text) & "_" & Trim(GridDatabase.Cell(lRow, 7).Text) & "_LOT" & Trim(GridDatabase.Cell(lRow, 1).Text & "_PW" & Trim(GridDatabase.Cell(lRow, 6).Text))
    
    
    NumCol = SetNumCol(Combo1)
    'Text1(0) = Trim(GridDatabase.Cell(FirstRow, NumCol).Text)
    ProductionID = GridDatabase.Cell(FirstRow, 14).Text
    ProductionFileName = GridDatabase.Cell(FirstRow, 13).Text
    bClosedProduction = GridDatabase.Cell(FirstRow, 15).Text
    PicMenu(4).Visible = True
    
    USER_PATH = IIf(bClosedProduction, USER_PRODUCTION_PATH & "Data\", USER_PRODUCTION_PATH)
   


    PicMenu(5).Visible = IsProductionStarted(ProductionID)
Else
End If
End Sub



Private Sub PicMenu_Click(Index As Integer)
Image3_Click Index
End Sub

Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

 
For i = PicMenu.LBound To PicMenu.UBound

    If i = Index Then
        PicMenu(i).BackColor = &H5D4B47
    Else
        PicMenu(i).BackColor = &H473733
    End If
Next
End Sub


Private Sub Picture1_Click()
'Picture1.BackColor = vbColorTextBlue ' &H8000&
DefaultMenu_Click 8

End Sub



Private Sub Text1_Change(Index As Integer)
If Me.Visible Then
    Text1(0).Top = Label1(0).Top + Label1(0).Height + Text1(0).Height / 20
    'Text1(Index).ForeColor = vbWhite
    If dIndexProcedura > 0 Then
       
        If Len(Text1(Index)) > 0 Then
            If (InStr(Text1(Index), Combo1)) Then
                Picture1.Visible = False
                DefaultMenu(8).Visible = False
            Else
                If Combo1.ListIndex = 0 Then
                    Picture1.Visible = True
                    DefaultMenu(8).Visible = True
                Else
                    Picture1.Visible = False
                    DefaultMenu(8).Visible = False
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub Text1_Click(Index As Integer)
Dim sString As String
Dim rc As Boolean

Select Case Index
    Case 0
        
        sString = Text1(0)
        If InStr(sString, Combo1) Then sString = ""
        If F_InputBox.DoShow(("Enter ") & Combo1, ("Filter : ") & Combo1, True, ("Apply"), ("Exit"), sString) Then
            Text1(0) = sString
           
            GlobalSearch
            Call ExistCampioneInTabella
            
            rc = IIf(Len(Text1(0)) > 0, True, False)
            'DefaultMenu(5).Visible = rc
            
            If dIndexProcedura > 0 Then
            
            
                If Combo1.ListIndex = 0 Then
                    Picture1.Visible = rc
                    DefaultMenu(8).Visible = rc
                Else
                    Picture1.Visible = False
                    DefaultMenu(8).Visible = False
                End If
            End If

            
            
            
            
        End If
End Select
End Sub

Private Sub Timer2_Timer()

Dim i As Integer
    '
    ' start form
    '
      bHilight = True

    If SelectedDBCode <> "" Then
        Text1(0) = SelectedDBCode
        ExistCampioneInTabella
    End If
     RiempiGrid GridDatabase
    Timer2.Enabled = False
    
    
End Sub

Private Sub Timer3_Timer()
    If MyLbHelpCount > 1 Then
      '  lbMenuHelp(0).Visible = False
      '  lbMenuHelp(1).Visible = False
        
        Picture1.BackColor = vbColorTextDarkBlue
        MyLbHelpCount = 0
    End If
    MyLbHelpCount = MyLbHelpCount + 1
End Sub


Private Sub GetFormDatiCampione()
Dim sString As String
Dim rc As Boolean
    sString = Text1(0)
    If InStr(Text1(0), Combo1) Then sString = ""
   ' rc = F_AnagraficaCampione.DoShow(sString)
    If rc Then
        PopupMessage 2, ("Campione salvato in Archivio"), , , ("Anagrafica Campione")
    End If
                
End Sub


Private Sub ScorriTabella(ByVal bValue As Boolean)

Dim MyRow As Integer
If GridDatabase.Rows > 1 Then
    MyRow = IIf(bValue, (IndexTabella * 10) + 10, (IndexTabella * 10) - 19)
    IndexTabella = IIf(bValue, IndexTabella + 1, IndexTabella - 1)
    If IndexTabella < 1 Then
        IndexTabella = 1
        GridDatabase.Cell(1, 1).EnsureVisible
    ElseIf MyRow >= GridDatabase.Rows Then
        GridDatabase.Cell(GridDatabase.Rows - 1, 1).EnsureVisible
        IndexTabella = MaxIndex
    'ElseIf IndexTabella >= MaxIndex - 1 And Not (bValue) Then
        'GridDatabase.Cell((IndexTabella) * 10, 1).EnsureVisible
    Else
         
        GridDatabase.Cell(MyRow, 1).EnsureVisible
    
    End If
End If

End Sub




Private Function ExistCampioneInTabella() As Boolean
    Dim NumCol As Integer
    Dim rc As Boolean
    Dim i As Integer
    
    rc = True
    If InStr(Text1(0), Combo1) Then Exit Function
    
    rc = IIf(GridDatabase.Rows < 2, False, True)
    
  '  NumCol = SetNumCol(Combo1)

    
    If Not (rc) And Combo1.ListIndex = 0 Then
        'If ExistCampione(Text1(0)) = False Then
        '    If F_MsgBox.DoShow(("Campione nuovo, inserire le specifiche in Archivio?"), Text1(0), True, ("SI"), ("NO"), Image1) Then
        '        DefaultMenu_Click 7
        '    End If
            
       ' End If
    End If
  
    ExistCampioneInTabella = rc
    
    
End Function
Private Function SetNumCol(ByVal sString As String) As Integer

Select Case Trim(UCase(sString))
    Case UCase(("Hanna Code"))
        SetNumCol = 4
    Case UCase(("Line"))
        SetNumCol = 1
    Case UCase(("Recipe"))
        SetNumCol = 6
    Case UCase(("Week"))
        SetNumCol = 3
    Case UCase(("Prep.Lot"))
        SetNumCol = 3
    Case UCase(("SFG Lot"))
        SetNumCol = 18
    
    End Select

End Function
Private Function GlobalSearch()
    
    PicMenu(5).Visible = False
    
    If Trim(Combo1) = Trim("SFG Lot") Then
        Call SearchSFGLot
    Else
    
        RiempiGrid GridDatabase
    
    End If
    '

End Function
Private Sub SearchSFGLot()
Dim i As Integer
Dim t As Integer
Dim rc As Boolean
Dim StrSearch As String
Dim strProdLots As String

StrSearch = UCase(Text1(0))

With GridDatabase
    .AutoRedraw = False
    If StrSearch <> "" And .Rows > 0 Then
        For i = 1 To .Rows - 1
            strProdLots = .Cell(i, 18).Text
            If InStr(strProdLots, StrSearch) Then
            Else
                .RowHeight(i) = 0
            End If
        Next
    End If
    .Refresh
    .AutoRedraw = True
End With

End Sub
Private Function OpenIntervalloDate()
Picture2.ZOrder
Picture2.Visible = True

End Function





Private Sub IntervalloDate(ByVal bValue As Boolean)
    
    
    MyPeriodo = IIf(bValue, "", MyPeriodo)
    ' se falso allora ripristino il periodo di filtro
    

    
    bFiltroIntervalloDate = bValue
    ctlCalendar1.Visible = False
    Picture2.Visible = False
     If bValue = False Then
        lbDataFiltro(0) = ""
        lbDataFiltro(1) = ""
        CmbVisual_Click
    Else
        GlobalSearch
    End If
End Sub


Public Function FillTabellaTutte(ByVal Grd As Grid, Optional ByVal Periodo As String, Optional ByVal StringaFiltro As String, Optional ByVal stringa As String, Optional MyDa As String, Optional MyA As String, Optional ByVal ChangefilterLots As Integer = 2, Optional ByVal strLine As String) As Boolean
    
    Dim i As Integer
    Dim t As Integer
    Dim rc As Boolean
   
    Dim sName As String
    Dim sString As String
    Dim dMyDA As Date
    Dim dMyA As Date
    Dim strSFGLotti As String
    Dim NumSFG As Integer
    Dim ProductionFileName As String
    Dim NowRighe As Integer
    On Error GoTo ERR_FILL
    rc = True
    

    stringa = Trim(stringa)
    
    If Len(Trim(MyDa)) > 0 Then
        dMyDA = FormatDateTime(MyDa, vbShortDate)
        dMyA = FormatDateTime(MyA, vbShortDate)
    End If
            
    
    
    Grd.Rows = 1
    Grd.AutoRedraw = False
    
    If StringaFiltro = "" Then
        sString = ""
    Else
    
        If InStr(UCase(stringa), UCase(("Search"))) Then
            sString = ""
        Else
            Select Case Trim(StringaFiltro)
                Case ("Hanna Code")
                    If stringa = "" Then
                    Else
                        sString = " and HannaCode like '*" & Replace(Trim(stringa), "'", "''") & "*'"
                    End If
                Case ("Line")
                    If stringa = "" Then
                    Else
                        sString = " and Line like '*" & Replace(Trim(stringa), "'", "''") & "*'"
                    End If
                Case ("Recipe")
                    If stringa = "" Then
                    Else
                        sString = " and Recipe like '*" & Replace(Trim(stringa), "'", "''") & "*'"
                    End If
                Case ("Week")
                    If stringa <> "" Then
                        rc = DateWeek(stringa, dMyDA, dMyA)
                        
                        If rc = False Then
                            PopupMessage 2, "Please enter valid week : es. 2/" & year(Now())
                        Else
                            Periodo = ""
                        End If
                    End If

                Case ("Prep.Lot")
                    If stringa = "" Then
                    Else
                        sString = " and Lot like '*" & Replace(Trim(stringa), "'", "''") & "*'"
                    End If
                    
            End Select
        End If
    End If
    
    If LCase(strLine) = "all lines" Or strLine = "" Then
       ' sString = ""
       ' Grid4.Column(2).Width = 150
    Else
    
        sString = sString & " and line like '*" & Replace(Trim(strLine), "'", "''") & "*'"
        
        
        
       ' Grid4.Column(2).Width = 0
    End If
    
    
    Select Case ChangefilterLots
        Case 0
           ' If Len(sString) > 0 Then
                sString = " and bClosed=FALSE"
           ' Else
               'sString = "Finished=FALSE"
           ' End If
        Case 1
            'If Len(sString) > 0 Then
                sString = " and bClosed=TRUE"
            'Else
               ' sString = "Finished=TRUE"
           ' End If
    
    End Select
    With dbTabProduction
    
        .filter = ""
        If Periodo <> "" Then
            Periodo = FormatDateTime(Periodo, vbShortDate)
            .filter = "StartDate>=#" & Periodo & "# " & sString
        Else
            .filter = "StartDate>#" & dMyDA & "# AND StartDate<=#" & dMyA & "# " & sString
            
        End If
        
        If .EOF Then
            GoTo ERR_END:
        Else
            '  trovato qualcosa....
        End If
            .MoveFirst
            Debug.Print .RecordCount
            Do
                With Grd
                    .AddItem "", False
                    
                    
                    
                    '.Cell(0, 1).Text = "Line"
                    '.Cell(0, 2).Text = "Date"
                    '.Cell(0, 3).Text = "Hanna Code"
                    '.Cell(0, 4).Text = "# Lot"
                    '.Cell(0, 5).Text = "Rcipe"e
                    '.Cell(0, 6).Text = "Mix"
                    '.Cell(0, 7).Text = "Pl. Ref."
                   '
                   ' .Cell(0, 8).Text = "Preparation Date"
                   ' .Cell(0, 9).Text = "Preparation Week"
                   ' .Cell(0, 10).Text = "# Prep. Week"
                   ' .Cell(0, 11).Text = "FileName"
                   ' .Cell(0, 12).Text = "ID"
                   ' .Cell(0, 13).Text = ""
        
                    .Cell(.Rows - 1, 0).Text = .Rows - 1
                    .Cell(.Rows - 1, 1).Text = IIf(IsNull(Trim(dbTabProduction!Line)), "", Trim(dbTabProduction!Line))
                    .Cell(.Rows - 1, 2).Text = IIf(IsNull(Trim(dbTabProduction!startDate)), "", FormatDataLAT(Trim(dbTabProduction!startDate)))
                    .Cell(.Rows - 1, 3).Text = PreparationWeek(dbTabProduction!startDate)
                    .Cell(.Rows - 1, 4).Text = IIf(IsNull(Trim(dbTabProduction!HannaCode)), "", Trim(dbTabProduction!HannaCode))
                    .Cell(.Rows - 1, 5).Text = IIf(IsNull(Trim(dbTabProduction!Lot)), "", Trim(dbTabProduction!Lot))
                    .Cell(.Rows - 1, 6).Text = IIf(IsNull(Trim(dbTabProduction!Recipe)), "", Trim(dbTabProduction!Recipe))
                    .Cell(.Rows - 1, 8).Text = IIf(IsNull(Trim(dbTabProduction!PlanningReference)), "", Trim(dbTabProduction!PlanningReference))
                    .Cell(.Rows - 1, 9).Text = IIf(IsNull(Trim(dbTabProduction!PrepDate)), "", Trim(dbTabProduction!PrepDate))
                    .Cell(.Rows - 1, 10).Text = IIf(IsNull(Trim(dbTabProduction!PrepWeek)), "", (Trim(dbTabProduction!PrepWeek)))
                    .Cell(.Rows - 1, 11).Text = IIf(IsNull(Trim(dbTabProduction!numPrepWeek)), "", Trim(dbTabProduction!numPrepWeek))
                    .Cell(.Rows - 1, 12).Text = IIf(IsNull(Trim(dbTabProduction!Note)), "", Trim(dbTabProduction!Note))
                    .Cell(.Rows - 1, 13).Text = IIf(IsNull(Trim(dbTabProduction!FileName)), "", Trim(dbTabProduction!FileName))
                    .Cell(.Rows - 1, 14).Text = dbTabProduction!ID
                    .Cell(.Rows - 1, 15).Text = dbTabProduction!bClosed
                    If dbTabProduction!bClosed Then
                        .Cell(.Rows - 1, 16).Text = IIf(IsNull(Trim(dbTabProduction!CloseDate)), .Cell(.Rows - 1, 2).Text, Trim(dbTabProduction!CloseDate))
                    End If
                    If dbTabProduction!bClosed Then
                        
                        For t = 1 To .Cols - 1
                        
                            .Cell(.Rows - 1, t).FontBold = True
                            .Cell(.Rows - 1, t).ForeColor = &H4D3B37   '&H644603
                        
                        
                        Next
                    
                    End If
                    
                    If dbTabProduction!ExcelDone Then
                        .Cell(.Rows - 1, 17).BackColor = vbColorGreen
                        .Cell(.Rows - 1, 17).Text = "OK"
                        .Cell(.Rows - 1, 17).ForeColor = vbWhite
                        
                    End If
                    
                    
                   ' .Cell(.Rows - 1, 18).Text = IIf(IsNull(Trim(dbTabProduction!Lot)), "", Trim(dbTabProduction!Lot))
                    ProductionFileName = IIf(IsNull(Trim(dbTabProduction!FileName)), "", Trim(dbTabProduction!FileName))
                    If ProductionFileName <> "" Then
                    If FileExists(USER_TEMP_PATH & ProductionFileName) Then
                        USER_PATH = USER_TEMP_PATH
                    ElseIf FileExists(USER_DATA_PATH & ProductionFileName) Then
                        USER_PATH = USER_DATA_PATH
                    Else
                        GoTo cont:
                    End If
                    
                    CloseSettingDataFile
                    
                    NumSFG = GetSettingData(ProductionFileName, "HannaCodes", "HannaCodesCount", 0)
                    strSFGLotti = ""
                    For i = 1 To NumSFG
                        strSFGLotti = strSFGLotti & " ; " & GetSettingData(ProductionFileName, "HannaCode" & i, "LotNumber", "")
                    Next
                    
                    
                    CloseSettingDataFile
                    
                    .Cell(.Rows - 1, 18).Text = strSFGLotti
                    
                    
                    
                   End If
                   
         
                End With
cont:
                If Not .EOF Then .MoveNext
            Loop Until .EOF
    End With

    
ERR_END:
    On Error GoTo 0
    Grd.Column(17).Alignment = cellCenterCenter
    Grd.Refresh
    Grd.AutoRedraw = True
    
    FillTabellaTutte = rc
    Exit Function
ERR_FILL:
    rc = False
    MsgBox err.Description
    Resume Next
End Function

Private Function CancellaTab() As Boolean
    
    If ProductionID > 0 Then
        If F_MsgBox.DoShow(("Delete Selected Record ?"), "Database", , ("Delete"), ("Exit")) Then
            
            If CancellaRecord(ProductionID) Then
                Text1(0) = ""
                GlobalSearch
                UploadDownloadMessageCounter = 0
                PopupMessage 2, ("Record Deleted..."), , , PROGRAM_NAME
               
            Else
            End If
        End If
    End If
End Function

Private Function CancellaRecord(ByVal ProductionID As Long) As Boolean
Dim rc As Boolean

    On Error GoTo ERR_CAN
    rc = True
    With dbTabProduction
        .filter = ""
        .filter = "ID='" & ProductionID & "'"
        If .EOF Then
        Else
            .Delete
            .Update
        
            ' cancello anche il file.....
            If ProductionFileName <> "" Then
                If FileExists(USER_PATH & ProductionFileName) Then Kill USER_PATH & ProductionFileName
            End If
        End If
    
    End With

ERR_END:
    On Error GoTo 0
    CancellaRecord = rc
    Exit Function
ERR_CAN:
    rc = False
    MsgBox err.Description
    Resume Next
End Function


Private Function ChangeLabelLots()
    Dim i As Integer
    For i = 0 To 2
        If i = IndexOpenClosedLot Then
            Label4(i).ForeColor = vbColorTextDarkBlue 'vbColorOrange
        Else
             Label4(i).ForeColor = &H404040
        End If
    Next
End Function


Private Sub FillProductionSpecifics()
Dim i As Integer

With Grid2
    .AutoRedraw = False
    .Rows = 1
    
 
    With dbTabProdHistory
        .filter = ""
        .filter = "ProductionID='" & ProductionID & "'"
        .MoveFirst
        For i = 1 To .RecordCount
        
            Grid2.AddItem "", False

            Grid2.Cell(i, 1).Text = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
            Grid2.Cell(i, 2).Text = IIf(IsNull(Trim(!QtyProduced)), "", Trim(!QtyProduced))
            Grid2.Cell(i, 3).Text = IIf(IsNull(Trim(!LotNumber)), "", Trim(!LotNumber))
            Grid2.Cell(i, 4).Text = IIf(IsNull(Trim(!Operator)), "", Trim(!Operator))
            Grid2.Cell(i, 5).Text = IIf(IsNull(Trim(!DateProd)), "", Trim(!DateProd))
            Grid2.Cell(i, 6).Text = IIf(IsNull(Trim(!WeekProd)), "", Trim(!WeekProd))
            Grid2.Cell(i, 7).Text = IIf(IsNull(Trim(!Machine)), "", Trim(!Machine))
            Grid2.Cell(i, 8).Text = IIf(IsNull(Trim(!Note)), "", Trim(!Note))
            
            Grid2.Cell(i, 9).Text = IIf(IsNull(Trim(!AcquisitionTime)), "", Trim(!AcquisitionTime))
            Grid2.Cell(i, 10).Text = !ID
            Grid2.Cell(i, 11).Text = IIf(IsNull(Trim(!Index)), 0, Trim(!Index))
            
            Grid2.Cell(i, 2).BackColor = vbColorResults
            Grid2.Cell(i, 2).Alignment = cellRightCenter
        
            .MoveNext
        Next
    End With
        
    .Refresh
    .AutoRedraw = True
End With



End Sub
