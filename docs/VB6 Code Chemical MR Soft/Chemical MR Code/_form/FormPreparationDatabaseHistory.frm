VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form FormPreparationDatabaseHistory 
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
   Icon            =   "FormPreparationDatabaseHistory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   Picture         =   "FormPreparationDatabaseHistory.frx":33E2
   ScaleHeight     =   12000
   ScaleWidth      =   19200
   Begin FlexCell.Grid Grid2 
      Height          =   6975
      Left            =   5520
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2160
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
   Begin VB.PictureBox ctlCalendar1 
      Height          =   6960
      Left            =   11760
      ScaleHeight     =   6900
      ScaleWidth      =   8055
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   8115
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00307030&
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":3724
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
            MouseIcon       =   "FormPreparationDatabaseHistory.frx":3A2E
            MousePointer    =   99  'Custom
            Picture         =   "FormPreparationDatabaseHistory.frx":3D38
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":711A
         MousePointer    =   99  'Custom
         Picture         =   "FormPreparationDatabaseHistory.frx":7424
         Top             =   240
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   1
         Left            =   4560
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":A806
         MousePointer    =   99  'Custom
         Picture         =   "FormPreparationDatabaseHistory.frx":AB10
         Top             =   240
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   2
         Left            =   600
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":DEF2
         MousePointer    =   99  'Custom
         Picture         =   "FormPreparationDatabaseHistory.frx":E1FC
         Top             =   240
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   3
         Left            =   15600
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":115DE
         MousePointer    =   99  'Custom
         Picture         =   "FormPreparationDatabaseHistory.frx":118E8
         Top             =   240
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   4
         Left            =   18240
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":14CCA
         MousePointer    =   99  'Custom
         Picture         =   "FormPreparationDatabaseHistory.frx":14FD4
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":183B6
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":186C0
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":189CA
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":18CD4
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":18FDE
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":192E8
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":195F2
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":198FC
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":19C06
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":19F10
         MousePointer    =   99  'Custom
         TabIndex        =   36
         Top             =   -120
         Width           =   2055
      End
   End
   Begin VB.PictureBox PicMainMenu 
      BackColor       =   &H00105010&
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
         BackColor       =   &H00105010&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   5
         Left            =   9600
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":1A21A
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
            MouseIcon       =   "FormPreparationDatabaseHistory.frx":1A524
            MousePointer    =   99  'Custom
            Picture         =   "FormPreparationDatabaseHistory.frx":1A82E
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "STD Specifics"
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
            MouseIcon       =   "FormPreparationDatabaseHistory.frx":1D220
            MousePointer    =   99  'Custom
            TabIndex        =   49
            Top             =   720
            Width           =   1890
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00105010&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   4
         Left            =   5760
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":1D52A
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
            MouseIcon       =   "FormPreparationDatabaseHistory.frx":1D834
            MousePointer    =   99  'Custom
            TabIndex        =   28
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   4
            Left            =   720
            MouseIcon       =   "FormPreparationDatabaseHistory.frx":1DB3E
            MousePointer    =   99  'Custom
            Picture         =   "FormPreparationDatabaseHistory.frx":1DE48
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00105010&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":2122A
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
            MouseIcon       =   "FormPreparationDatabaseHistory.frx":21534
            MousePointer    =   99  'Custom
            Picture         =   "FormPreparationDatabaseHistory.frx":2183E
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
            MouseIcon       =   "FormPreparationDatabaseHistory.frx":24C20
            MousePointer    =   99  'Custom
            TabIndex        =   24
            Top             =   720
            Width           =   960
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00105010&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   1920
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":24F2A
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   720
            MouseIcon       =   "FormPreparationDatabaseHistory.frx":25234
            MousePointer    =   99  'Custom
            Picture         =   "FormPreparationDatabaseHistory.frx":2553E
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Close Preparation"
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
            Left            =   195
            MouseIcon       =   "FormPreparationDatabaseHistory.frx":27F30
            MousePointer    =   99  'Custom
            TabIndex        =   22
            Top             =   720
            Width           =   1500
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00105010&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   3840
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":2823A
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
            MouseIcon       =   "FormPreparationDatabaseHistory.frx":28544
            MousePointer    =   99  'Custom
            Picture         =   "FormPreparationDatabaseHistory.frx":2884E
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
            MouseIcon       =   "FormPreparationDatabaseHistory.frx":2BC30
            MousePointer    =   99  'Custom
            TabIndex        =   20
            Top             =   720
            Width           =   1890
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00105010&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   3
         Left            =   7680
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":2BF3A
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
            MouseIcon       =   "FormPreparationDatabaseHistory.frx":2C244
            MousePointer    =   99  'Custom
            Picture         =   "FormPreparationDatabaseHistory.frx":2C54E
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
            MouseIcon       =   "FormPreparationDatabaseHistory.frx":2F930
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
         Caption         =   "Preparation Database"
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
         Left            =   14325
         TabIndex        =   25
         Top             =   240
         Width           =   4365
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00307030&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   3360
      ScaleHeight     =   2895
      ScaleWidth      =   13455
      TabIndex        =   7
      Top             =   1440
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":2FC3A
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":2FF44
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
            Size            =   14.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   4560
         TabIndex        =   26
         Top             =   2160
         Width           =   3990
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   11
         Left            =   480
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":3024E
         MousePointer    =   99  'Custom
         Picture         =   "FormPreparationDatabaseHistory.frx":30558
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   10
         Left            =   12480
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":3393A
         MousePointer    =   99  'Custom
         Picture         =   "FormPreparationDatabaseHistory.frx":33C44
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":37026
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
         MouseIcon       =   "FormPreparationDatabaseHistory.frx":37330
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
      Left            =   360
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
   Begin VB.Frame frSpecifics 
      BackColor       =   &H00A0A0A0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   360
      TabIndex        =   51
      Top             =   1200
      Visible         =   0   'False
      Width           =   18495
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "QC Specifics : "
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
         Top             =   360
         Width           =   18495
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Preparation"
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
      Left            =   5175
      MouseIcon       =   "FormPreparationDatabaseHistory.frx":3763A
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   9840
      Width           =   1725
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Closed Preparation"
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
      Left            =   2655
      MouseIcon       =   "FormPreparationDatabaseHistory.frx":37944
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   9840
      Width           =   2235
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Preparation"
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
      Left            =   255
      MouseIcon       =   "FormPreparationDatabaseHistory.frx":37C4E
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   9840
      Width           =   2115
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
      MouseIcon       =   "FormPreparationDatabaseHistory.frx":37F58
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
      MouseIcon       =   "FormPreparationDatabaseHistory.frx":38262
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
Attribute VB_Name = "FormPreparationDatabaseHistory"
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
Private PreparationID As Long
Private PreparationFileName As String
Private bMeanValue As Boolean
Private bClosedPreparation As Boolean
Private MyIndexRecord As Integer
Private lRow As Long
Private IndexOpenClosedLot As Integer
Private HannaCode As String
Private bManualPreparation As Boolean


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


Call SetPreparationGrid(GridDatabase)

Call SetPreparationComponentGrid(Grid2)

    PreparationID = 0
    PreparationFileName = ""
    MyIndexRecord = 3
    bMeanValue = False
    bClosedPreparation = False
    
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



Private Sub Form_Unload(Cancel As Integer)
Set FormPreparationDatabaseHistory = Nothing
End Sub

Private Sub GridDatabase_BeforeUserSort(ByVal Col As Long)
lRow = 0
End Sub

Private Sub GridDatabase_Click()
ctlCalendar1.Visible = False

End Sub


Private Sub GridDatabase_DblClick()
If lRow = 0 Then Exit Sub

Call OpenPreparation(False)

End Sub








Private Function OpenPreparation(ByVal bClosePreparation As Boolean)
    If PreparationFileName <> "" Then
    
        If FileExists(USER_TEMP_PATH & PreparationFileName) Then
            USER_PATH = USER_TEMP_PATH
        ElseIf FileExists(USER_DATA_PATH & PreparationFileName) Then
            USER_PATH = USER_DATA_PATH
        
        Else
            PopupMessage 3, "Cannot find Preparation file!", , True
            Exit Function
            
        End If
        
        If bManualPreparation Then
        
            frmPreparation_Manual.Left = Me.Left
            frmPreparation_Manual.Top = Me.Top
            frmPreparation_Manual.WindowState = Me.WindowState
            frmPreparation_Manual.DoShow HannaCode, PreparationFileName, PreparationID, , bClosePreparation
        
        Else
        
        
            frmPreparation_Static.Left = Me.Left
            frmPreparation_Static.Top = Me.Top
            DoEvents
            DoEvents
            If frmPreparation_Static.DoShow(HannaCode, PreparationFileName, PreparationID, , bClosePreparation) Then
                
              
            
            End If
            
        End If
        
        DoEvents
        DoEvents
        Call GlobalSearch
        
        
    End If


End Function



Private Sub Image3_Click(Index As Integer)
Dim MyexcelName As String

    Select Case Index
        Case 0
            Combo1.ListIndex = 0
            Combo1_Click
             IntervalloDate False
            GlobalSearch
        Case 1
            OpenPreparation (True)

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
            ExcelName = FormatNomeFile("Preparation List." & FormatDataLAT(Now()))
            
            If F_InputBox.DoShow("Please Set Excel Name", "Preparation To Excel", , , , ExcelName) Then
    
                ExcelName = USER_DESKTOP & "\" & FormatNomeFile(ExcelName) & ".xls"
                
                
                GridDatabase.ExportToExcel ExcelName, True, True
                MessageInfoTime = 2500
                PopupMessage 2, "Excel Done..." & vbCrLf & ExcelName
                
            End If
            
        Case 5
            ' View Preparation History
            Call ViewPreparationQC

    End Select
End Sub








Private Sub ViewPreparationQC()

    Grid2.ZOrder
    frSpecifics.ZOrder
    If Grid2.Visible = True Then
        GridDatabase.Visible = True
        frSpecifics.Visible = False
        Me.BackColor = &HE0E0E0
        Label2(5) = "QC Specifics"
    Else
        frSpecifics.Visible = True
        Me.BackColor = &HA0A0A0
        Call FillPreparationSpecifics
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
rc = IIf(GridDatabase.Column(6).Width = 0, True, False)
SaveSetting App.Title, Me.Name, "Visualizza Colonne", Not (rc)
VisulaizzaColonne rc

End Sub

Private Sub VisulaizzaColonne(ByVal rc As Boolean)
With GridDatabase
    .AutoRedraw = False
    .Column(5).Width = IIf(rc, 120, 0)
    .Column(6).Width = IIf(rc, 120, 0)
    .Column(7).Width = IIf(rc, 120, 0)
    .Column(8).Width = IIf(rc, 120, 0)
    .Column(9).Width = IIf(rc, 120, 0)
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
Dim rc As Boolean

rc = IIf(InStr(Combo1, "Correction"), False, True)
Text1(0).Visible = rc
Label1(0).Visible = rc
    
PicMenu(5).Visible = False
GridDatabase.Cell(0, 0).SetFocus
    
    Text1(0) = (" - Search") & Combo1 & " - "
    Label1(0) = UCase(Combo1)
    'If Me.Visible Then Text1(0).SetFocus
    SaveSetting App.Title, Me.Name, "Filtro Combo", Combo1.ListIndex
'End If
 
If Not (rc) Then GlobalSearch


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
        .AddItem " " & ("Hanna Code")
        .AddItem " " & ("MR Code")
        .AddItem " " & ("Week")
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
    Call FillTabellaTutte(Grd, MyPeriodo, Combo1, Text1(0), MyDa, MyA, IndexOpenClosedLot)
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
  VisulaizzaColonne Not (rc)
  Grid2.Move GridDatabase.Left, GridDatabase.Top, GridDatabase.Width, GridDatabase.Height
End Sub



Private Sub GridDatabase_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim NumCol As Integer

PreparationID = 0
PreparationFileName = ""
bMeanValue = False
bClosedPreparation = False
lRow = FirstRow
ExcelFilename = ""
PicMenu(4).Visible = False
PicMenu(5).Visible = False


   bManualPreparation = False
  
If FirstRow > 0 Then

    MyLbHelpCount = 0
  
    ' RECIPE + CODE + RANGE + LOT + PREPWK
    ExcelFilename = Trim(GridDatabase.Cell(lRow, 5).Text) & "_" & Trim(GridDatabase.Cell(lRow, 2).Text) & "_" & Trim(GridDatabase.Cell(lRow, 7).Text) & "_LOT" & Trim(GridDatabase.Cell(lRow, 1).Text & "_PW" & Trim(GridDatabase.Cell(lRow, 6).Text))
    
    
    NumCol = SetNumCol(Combo1)
    'Text1(0) = Trim(GridDatabase.Cell(FirstRow, NumCol).Text)
    PreparationID = GridDatabase.Cell(FirstRow, 16).Text
    PreparationFileName = GridDatabase.Cell(FirstRow, 17).Text
    bClosedPreparation = GridDatabase.Cell(FirstRow, 15).Text
    HannaCode = GridDatabase.Cell(FirstRow, 1).Text
    PicMenu(4).Visible = True


    USER_PATH = IIf(bClosedPreparation, USER_DATA_PATH, USER_TEMP_PATH)
    
    
    If IsNull(GridDatabase.Cell(FirstRow, 20).Text) Or GridDatabase.Cell(FirstRow, 20).Text = "" Then
   Else
    bManualPreparation = True
   End If
   
    PicMenu(1).Visible = Not (bClosedPreparation)


    'PicMenu(5).Visible = IsPreparationStarted(PreparationID)
Else
End If
End Sub



Private Sub PicMenu_Click(Index As Integer)
Image3_Click Index
End Sub

Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer

 
For i = PicMenu.LBound To PicMenu.UBound

    If i = Index Then
        PicMenu(i).BackColor = &H307030
    Else
        PicMenu(i).BackColor = &H105010
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
    
    End Select

End Function
Private Function GlobalSearch()
    
    PicMenu(5).Visible = False
    
    RiempiGrid GridDatabase
    '

End Function

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

    Dim NowRighe As Integer
    On Error GoTo ERR_FILL
    rc = True
    

    stringa = Trim(stringa)
    
    
    
    If Len(Trim(MyDa)) > 0 Then
        dMyDA = FormatDateTime(MyDa, vbShortDate)
        dMyA = FormatDateTime(MyA, vbShortDate)
    End If
            
    
    Grd.AutoRedraw = False
    Grd.Rows = 1
    
    
    If StringaFiltro = "" Then
        sString = ""
    Else
    
    

        
        If InStr(UCase(stringa), UCase(("Search"))) Then
   
        Else
            Select Case Trim(StringaFiltro)
                Case ("Hanna Code")
                    If stringa = "" Then
                    Else
                        sString = " and HannaCode like '*" & Replace(Trim(stringa), "'", "''") & "*'"
                    End If
              
                Case ("MR Code")
                    If stringa = "" Then
                    Else
                        sString = " and MRCode like '*" & Replace(Trim(stringa), "'", "''") & "*'"
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
                
            End Select
        End If
    End If
    If Trim(StringaFiltro) = ("With Correction") Then
        sString = sString & " and bCorrection=true"
        GoTo cont:
    End If
    
    If LCase(strLine) = "all lines" Or strLine = "" Then
       ' sString = ""
       ' Grid4.Column(2).Width = 150
    Else
    
        sString = sString & " and line like '*" & Replace(Trim(strLine), "'", "''") & "*'"
        
        
        
       ' Grid4.Column(2).Width = 0
    End If
cont:
    
    
    Select Case ChangefilterLots
        Case 0
           ' If Len(sString) > 0 Then
                sString = sString & " and bClosed=FALSE"
           ' Else
               'sString = "Finished=FALSE"
           ' End If
        Case 1
            'If Len(sString) > 0 Then
                sString = sString & " and bClosed=TRUE"
            'Else
               ' sString = "Finished=TRUE"
           ' End If
    
    End Select
    With dbTabPreparation
        .filter = ""
        If Periodo <> "" Then
            Periodo = FormatDateTime(Periodo, vbShortDate)
            .filter = "DataPrep>=#" & FormatDataLAT(Periodo) & "# " & sString
        Else
            .filter = "DataPrep>=#" & FormatDataLAT(CDate(dMyDA)) & "# AND DataPrep<=#" & FormatDataLAT(CDate(dMyA)) & "# " & sString
            
        End If
        
        If .EOF Then
            GoTo ERR_END:
        Else
            '  trovato qualcosa....
        End If
            .MoveFirst
            Do
                With Grd
                    .AddItem "", False

        
        
        
        
                
                    .Cell(.Rows - 1, 0).Text = .Rows - 1
                         
                    
        
                    .Cell(.Rows - 1, 1).Text = IIf(IsNull(Trim(dbTabPreparation!HannaCode)), "", Trim(dbTabPreparation!HannaCode))
                    .Cell(.Rows - 1, 2).Text = IIf(IsNull(Trim(dbTabPreparation!Description)), "", Trim(dbTabPreparation!Description))
                   
                   
                   .Cell(.Rows - 1, 3).Text = IIf(IsNull(Trim(dbTabPreparation!MRCode)), "", Trim(dbTabPreparation!MRCode))
                   .Cell(.Rows - 1, 4).Text = FormatDataLAT(IIf(IsNull(Trim(dbTabPreparation!DataPrep)), "", Trim(dbTabPreparation!DataPrep)))
                   
                   .Cell(.Rows - 1, 5).Text = IIf(IsNull(Trim(dbTabPreparation!HourPrep)), "", Trim(dbTabPreparation!HourPrep))
                   .Cell(.Rows - 1, 6).Text = IIf(IsNull(Trim(dbTabPreparation!PrepWeek)), "", Trim(dbTabPreparation!PrepWeek))
                   .Cell(.Rows - 1, 7).Text = IIf(IsNull(Trim(dbTabPreparation!Operator)), "", Trim(dbTabPreparation!Operator))
                   .Cell(.Rows - 1, 8).Text = PadString(IIf(IsNull(Trim(dbTabPreparation!QtyToProduce)), "", Trim(dbTabPreparation!QtyToProduce)))
                   
                   .Cell(.Rows - 1, 9).Text = IIf(IsNull(Trim(dbTabPreparation!Unit)), "", Trim(dbTabPreparation!Unit))
                   .Cell(.Rows - 1, 10).Text = (IIf(IsNull(Trim(dbTabPreparation!STDMatrix)), "", Trim(dbTabPreparation!STDMatrix)))
                   .Cell(.Rows - 1, 11).Text = IIf(IsNull(Trim(dbTabPreparation!STDExp)), "", Trim(dbTabPreparation!STDExp))
                   .Cell(.Rows - 1, 12).Text = IIf(IsNull(Trim(dbTabPreparation!STDStorage)), "", Trim(dbTabPreparation!STDStorage))
                   
                   .Cell(.Rows - 1, 13).Text = IIf(IsNull(Trim(dbTabPreparation!Note)), "", Trim(dbTabPreparation!Note))
                   .Cell(.Rows - 1, 14).Text = IIf(IsNull(Trim(dbTabPreparation!MsType)), False, Trim(dbTabPreparation!MsType))
                   .Cell(.Rows - 1, 15).Text = IIf(IsNull(Trim(dbTabPreparation!bClosed)), False, Trim(dbTabPreparation!bClosed))
                   .Cell(.Rows - 1, 16).Text = dbTabPreparation!ID
                   .Cell(.Rows - 1, 17).Text = IIf(IsNull(Trim(dbTabPreparation!FileName)), "", Trim(dbTabPreparation!FileName))
        
                    .Cell(.Rows - 1, 1).FontBold = True
                    .Cell(.Rows - 1, 1).ForeColor = &H473733
                    .Cell(.Rows - 1, 7).FontBold = True
                    .Cell(.Rows - 1, 7).ForeColor = &H473733
                    
                    .Cell(.Rows - 1, 8).FontBold = True
                    .Cell(.Rows - 1, 8).ForeColor = &H473733
                    .Cell(.Rows - 1, 9).FontBold = True
                    .Cell(.Rows - 1, 9).ForeColor = &H473733
                             
            
                    
                    Select Case dbTabPreparation!MsType
                        
                        Case "1"
                            .Cell(.Rows - 1, 14).BackColor = vbColorAzzurrino
                             .Cell(.Rows - 1, 3).BackColor = vbColorAzzurrino
                        Case "2"
                            .Cell(.Rows - 1, 14).BackColor = vbColorRosaTabella
                            .Cell(.Rows - 1, 3).BackColor = vbColorRosaTabella
                        Case Else
                            .Cell(.Rows - 1, 14).BackColor = vbColorResults
                            .Cell(.Rows - 1, 3).BackColor = vbColorResults
                    
                    End Select
                    
                    .Cell(.Rows - 1, 3).Alignment = cellCenterCenter
                    
                         
                    
                    If dbTabPreparation!bClosed Then
                        .Cell(.Rows - 1, 18).Text = FormatDataLAT(IIf(IsNull(Trim(dbTabPreparation!CloseDate)), "", Trim(dbTabPreparation!CloseDate)))
                     End If

                    If dbTabPreparation!bClosed Then
                        
                        For t = 1 To .Cols - 1
                        
                            .Cell(.Rows - 1, t).FontBold = True
                            .Cell(.Rows - 1, t).ForeColor = &H4D3B37   '&H644603
                        
                        
                        Next
                    
                    End If

                    If dbTabPreparation!ExcelDone Then
                        .Cell(.Rows - 1, 19).BackColor = vbColorGreen
                        .Cell(.Rows - 1, 19).Text = "OK"
                         .Cell(.Rows - 1, 19).Alignment = cellCenterCenter
                        .Cell(.Rows - 1, 19).ForeColor = vbWhite
                        
                    End If
                    
                    
                    
                        
            If dbTabPreparation!bManuale Then
                .Cell(.Rows - 1, 20).Text = "true"
                .Cell(.Rows - 1, 20).ForeColor = vbColorManualPreparation
                .Cell(.Rows - 1, 20).BackColor = vbColorManualPreparation
                .Cell(.Rows - 1, 1).BackColor = vbColorManualPreparation
                .Cell(.Rows - 1, 3).BackColor = vbColorManualPreparation
                .Cell(.Rows - 1, 1).ForeColor = vbWhite
                .Cell(.Rows - 1, 3).ForeColor = vbWhite
           Else
            
           End If
                    
                    
                    
                    
                    
         
                End With
                .MoveNext
            Loop Until .EOF
            
            For i = 1 To Grd.Cols - 1
                Grd.Column(i).AutoFit
                Grd.Column(i).Width = Grd.Column(i).Width * 1.1
            Next
            
             Grd.Column(14).Width = 0
             Grd.Column(15).Width = 0
             Grd.Column(16).Width = 0
             Grd.Column(17).Width = 0
             
             
    End With

    
ERR_END:
    On Error GoTo 0
     
    Grd.Column(18).Alignment = cellCenterCenter
    Grd.AutoRedraw = True
    Grd.Refresh
    FillTabellaTutte = rc
    Exit Function
ERR_FILL:
    rc = False
    'MsgBox err.Description
    Resume Next
End Function

Private Function CancellaTab() As Boolean
    
    If PreparationID > 0 Then
        If F_MsgBox.DoShow(("Delete Selected Record ?"), "Database", , ("Delete"), ("Exit")) Then
            
            If CancellaRecord(PreparationID) Then
                Text1(0) = ""
                GlobalSearch
                UploadDownloadMessageCounter = 0
                PopupMessage 2, ("Record Deleted..."), , , PROGRAM_NAME
               
            Else
            End If
        End If
    End If
End Function

Private Function CancellaRecord(ByVal PreparationID As Long) As Boolean
Dim rc As Boolean

    On Error GoTo ERR_CAN
    rc = True
    With dbTabPreparation
        .filter = ""
        .filter = "ID='" & PreparationID & "'"
        If .EOF Then
        Else
            .Delete
            .Update
        
            ' cancello anche il file.....
            If PreparationFileName <> "" Then
                If FileExists(USER_PATH & PreparationFileName) Then Kill USER_PATH & PreparationFileName
            End If
        End If
    
    End With

ERR_END:
    On Error GoTo 0
    CancellaRecord = rc
    Exit Function
ERR_CAN:
    rc = False
    MsgBox Err.Description
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


Private Sub FillPreparationSpecifics()


    CloseSettingDataFile
    Label3 = "STD Specifics : " & HannaCode
    Grid2.Rows = 1
    Grid2.ZOrder
    GridDatabase.Visible = False
   
    CloseSettingDataFile




End Sub
