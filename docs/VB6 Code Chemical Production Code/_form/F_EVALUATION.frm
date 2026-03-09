VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form F_EVALUATION 
   BackColor       =   &H00303030&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12000
   ClientLeft      =   0
   ClientTop       =   0
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
   LinkTopic       =   "Form1"
   Picture         =   "F_EVALUATION.frx":0000
   ScaleHeight     =   12000
   ScaleMode       =   0  'User
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   0
      Left            =   0
      MouseIcon       =   "F_EVALUATION.frx":1DED9
      MousePointer    =   99  'Custom
      ScaleHeight     =   1815
      ScaleWidth      =   2775
      TabIndex        =   17
      Top             =   1080
      Width           =   2775
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recipe / Reagent"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   525
         MouseIcon       =   "F_EVALUATION.frx":1E1E3
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   0
         Left            =   1200
         Picture         =   "F_EVALUATION.frx":1E4ED
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00606060&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   2760
      TabIndex        =   16
      Top             =   1080
      Width           =   16455
      Begin VB.Frame Frame4 
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   1695
         Left            =   5880
         TabIndex        =   80
         Top             =   0
         Visible         =   0   'False
         Width           =   8895
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   12
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   83
            Top             =   1080
            Width           =   3495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   11
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   81
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Label lb 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "Closing Information"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   1
            Left            =   360
            TabIndex        =   85
            Top             =   320
            Width           =   7095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "by"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   12
            Left            =   3960
            TabIndex        =   84
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "Validation Date"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   11
            Left            =   360
            TabIndex        =   82
            Top             =   720
            Width           =   3495
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1455
         Left            =   6000
         TabIndex        =   45
         Top             =   240
         Visible         =   0   'False
         Width           =   8895
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   10
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   49
            Text            =   "3.23"
            Top             =   840
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   7
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   48
            Text            =   "0.4"
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   8
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   47
            Text            =   "2.4"
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   9
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "3.09"
            Top             =   840
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "MAX"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   10
            Left            =   6000
            TabIndex        =   55
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "MIN"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   9
            Left            =   4080
            TabIndex        =   54
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "MAX ( ppm )"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   8
            Left            =   1920
            TabIndex        =   53
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "MIN ( ppm )"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   7
            Left            =   0
            TabIndex        =   52
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lb 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "STD Range"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   7
            Left            =   0
            TabIndex        =   51
            Top             =   80
            Width           =   3735
         End
         Begin VB.Label lb 
            Alignment       =   2  'Center
            BackColor       =   &H00004080&
            Caption         =   "pH Range"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   9
            Left            =   4080
            TabIndex        =   50
            Top             =   80
            Visible         =   0   'False
            Width           =   3735
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "hjkhkj"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "hhhh"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lb 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "Lot"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   0
         Left            =   1200
         TabIndex        =   44
         Top             =   320
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "Code SFG"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   1
         Left            =   3240
         TabIndex        =   21
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "Lot Number"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   0
         Left            =   1200
         TabIndex        =   19
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.PictureBox PicMain 
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   1
      Left            =   0
      Picture         =   "F_EVALUATION.frx":218CF
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   22
      Top             =   2880
      Width           =   19215
      Begin VB.PictureBox PicInformation 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   12840
         MouseIcon       =   "F_EVALUATION.frx":3F7A8
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   5295
         TabIndex        =   78
         Top             =   5760
         Visible         =   0   'False
         Width           =   5295
         Begin VB.Label Lab 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Goto Lot Information QC"
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
            Left            =   0
            MouseIcon       =   "F_EVALUATION.frx":3FAB2
            MousePointer    =   99  'Custom
            TabIndex        =   79
            Top             =   660
            Width           =   5280
         End
         Begin VB.Image Im 
            Height          =   480
            Left            =   2400
            MouseIcon       =   "F_EVALUATION.frx":3FDBC
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":400C6
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         FillColor       =   &H00004000&
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   7080
         ScaleHeight     =   855
         ScaleWidth      =   5295
         TabIndex        =   41
         Top             =   5880
         Visible         =   0   'False
         Width           =   5295
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "CLOSE LOT"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   57
            Top             =   240
            Width           =   5295
         End
         Begin VB.Image ImageTAV 
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "F_EVALUATION.frx":434A8
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":437B2
            Top             =   180
            Width           =   480
         End
      End
      Begin FlexCell.Grid Grd3 
         Height          =   4080
         Left            =   2160
         TabIndex        =   43
         Top             =   1440
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   7197
         AllowUserReorderColumn=   -1  'True
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColor1      =   14737632
         BackColor2      =   14737632
         BackColorActiveCellSel=   12632256
         BackColorBkg    =   16777215
         BackColorFixed  =   12632256
         BackColorFixedSel=   12632256
         BackColorScrollBar=   -2147483635
         BackColorSel    =   8421504
         BorderColor     =   9849089
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
      Begin VB.Label lbClose 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"F_EVALUATION.frx":46B94
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   0
         TabIndex        =   77
         Top             =   480
         Visible         =   0   'False
         Width           =   19125
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   9360
         MouseIcon       =   "F_EVALUATION.frx":46C26
         MousePointer    =   99  'Custom
         Picture         =   "F_EVALUATION.frx":46F30
         Top             =   5880
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Only Laboratory Manager can Close Lots"
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   7560
         TabIndex        =   56
         Top             =   6360
         Width           =   4140
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Specifications Table"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   2160
         MouseIcon       =   "F_EVALUATION.frx":4A312
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   2
         Left            =   2760
         Picture         =   "F_EVALUATION.frx":4A61C
         Top             =   1200
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox PicMenuBar 
      BackColor       =   &H00303030&
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
      Index           =   0
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   19215
      TabIndex        =   10
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   3
         Left            =   9600
         MouseIcon       =   "F_EVALUATION.frx":4D9FE
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   70
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Graph QC"
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
            Index           =   3
            Left            =   0
            MouseIcon       =   "F_EVALUATION.frx":4DD08
            MousePointer    =   99  'Custom
            TabIndex        =   71
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   735
            MouseIcon       =   "F_EVALUATION.frx":4E012
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":4E31C
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   7680
         MouseIcon       =   "F_EVALUATION.frx":516FE
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   68
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   735
            MouseIcon       =   "F_EVALUATION.frx":51A08
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":51D12
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reading QC"
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
            Index           =   2
            Left            =   0
            MouseIcon       =   "F_EVALUATION.frx":550F4
            MousePointer    =   99  'Custom
            TabIndex        =   69
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
         MouseIcon       =   "F_EVALUATION.frx":553FE
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   13
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "F_EVALUATION.frx":55708
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":55A12
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "All Readings"
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
            Index           =   0
            Left            =   60
            MouseIcon       =   "F_EVALUATION.frx":58DF4
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   720
            Width           =   1875
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   1920
         MouseIcon       =   "F_EVALUATION.frx":590FE
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   11
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Mean value"
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
            Index           =   1
            Left            =   60
            MouseIcon       =   "F_EVALUATION.frx":59408
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   720
            Width           =   1875
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   735
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":59712
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Evaluation QC"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   12150
         TabIndex        =   15
         Top             =   360
         Width           =   6540
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   0
      Left            =   0
      Picture         =   "F_EVALUATION.frx":5CAF4
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   19215
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   13
         Left            =   16800
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   3720
         Width           =   1935
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         FillColor       =   &H00004000&
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   14760
         ScaleHeight     =   855
         ScaleWidth      =   3975
         TabIndex        =   64
         Top             =   960
         Width           =   3975
         Begin VB.Label lbSTDNumber 
            Alignment       =   2  'Center
            BackColor       =   &H00004000&
            BackStyle       =   0  'Transparent
            Caption         =   "STD Number"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   65
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   6
         Left            =   16800
         TabIndex        =   38
         Top             =   5760
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   5
         Left            =   14760
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2535
         Left            =   3480
         TabIndex        =   24
         Top             =   2520
         Visible         =   0   'False
         Width           =   12255
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "6 - Laboratory Manager Only can Close Lots"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   1560
            TabIndex        =   61
            Top             =   2040
            Width           =   4455
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "4 - Save STD Mean Value"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   1560
            TabIndex        =   60
            Top             =   1320
            Width           =   2520
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "5 - Check Results"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   1560
            TabIndex        =   59
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H000080DF&
            BorderColor     =   &H000060BF&
            Height          =   2535
            Left            =   0
            Top             =   0
            Width           =   12255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3 - Check ph Range : Select ph to view pH number and Range"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   27
            Top             =   960
            Width           =   6240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2 - Select Value from Readings Table : Select/Deselect CheckBoxes ( at least 80% of all value )"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1560
            TabIndex        =   26
            Top             =   600
            Width           =   9450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1 - Select Standard from SFG Standard Table to show all STD Readings"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   25
            Top             =   240
            Width           =   7155
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   600
            Picture         =   "F_EVALUATION.frx":7A9CD
            Top             =   960
            Width           =   480
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   4
         Left            =   14760
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   4800
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   14760
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   14760
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2640
         Width           =   3975
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         FillColor       =   &H00004000&
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   14760
         ScaleHeight     =   855
         ScaleWidth      =   3975
         TabIndex        =   29
         Top             =   6720
         Visible         =   0   'False
         Width           =   3975
         Begin VB.Image ImageTAV 
            Height          =   480
            Index           =   5
            Left            =   1755
            MouseIcon       =   "F_EVALUATION.frx":7DDAF
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":7E0B9
            Top             =   170
            Width           =   480
         End
      End
      Begin FlexCell.Grid Grd1 
         Height          =   6600
         Left            =   240
         TabIndex        =   0
         Top             =   960
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   11642
         AllowUserReorderColumn=   -1  'True
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColor1      =   14737632
         BackColor2      =   14737632
         BackColorBkg    =   16777215
         BackColorFixed  =   16512
         BackColorFixedSel=   16512
         BackColorScrollBar=   -2147483635
         BackColorSel    =   16777215
         BorderColor     =   9849089
         CellBorderColor =   16512
         CellBorderColorFixed=   -2147483635
         Cols            =   10
         DefaultFontName =   "Calibri"
         DefaultFontSize =   12
         DefaultFontBold =   -1  'True
         DisplayDateTimeMask=   -1  'True
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   -2147483635
         ReadOnly        =   -1  'True
         Rows            =   10
         SelectionMode   =   1
         MultiSelect     =   0   'False
         DateFormat      =   2
         EnterKeyMoveTo  =   1
         BackColorComment=   -2147483635
         AllowUserPaste  =   2
      End
      Begin FlexCell.Grid Grd2 
         Height          =   6600
         Left            =   4560
         TabIndex        =   42
         Top             =   960
         Width           =   9840
         _ExtentX        =   17357
         _ExtentY        =   11642
         AllowUserReorderColumn=   -1  'True
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColor1      =   14737632
         BackColor2      =   14737632
         BackColorActiveCellSel=   12632256
         BackColorBkg    =   16777215
         BackColorFixed  =   14737632
         BackColorFixedSel=   14737632
         BackColorScrollBar=   -2147483635
         BackColorSel    =   8421504
         BorderColor     =   9849089
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
         SelectionMode   =   3
         MultiSelect     =   0   'False
         DateFormat      =   2
         EnterKeyMoveTo  =   1
         BackColorComment=   -2147483635
         AllowUserPaste  =   2
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "# Tests"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   13
         Left            =   16800
         TabIndex        =   67
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label lb80Pec 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select at least 70% of alla value"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6405
         TabIndex        =   58
         Top             =   480
         Width           =   8010
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "# Tests ( S )"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   6
         Left            =   16800
         TabIndex        =   39
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "# Reading ( S )"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   5
         Left            =   14760
         TabIndex        =   37
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Mean Value ( Selected )"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   4
         Left            =   14760
         TabIndex        =   35
         Top             =   4440
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "# Readings"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   3
         Left            =   14760
         TabIndex        =   33
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Total Average"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   2
         Left            =   14760
         TabIndex        =   31
         Top             =   2280
         Width           =   3975
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   1
         Left            =   4560
         Picture         =   "F_EVALUATION.frx":8149B
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Readings"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   5160
         MouseIcon       =   "F_EVALUATION.frx":8487D
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   480
         Width           =   1995
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SFG Standard"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   840
         MouseIcon       =   "F_EVALUATION.frx":84B87
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   480
         Width           =   2100
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   3
         Left            =   240
         Picture         =   "F_EVALUATION.frx":84E91
         Top             =   360
         Width           =   480
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Left            =   9360
         Picture         =   "F_EVALUATION.frx":88273
         Top             =   8880
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operator"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Top             =   9120
         Width           =   975
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   8
         Left            =   120
         TabIndex        =   9
         Top             =   7440
         Width           =   2655
      End
   End
   Begin ChemicalQC.ctlCalendar ctlCalendar1 
      Height          =   6960
      Left            =   7080
      TabIndex        =   76
      Top             =   3240
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
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   4
      Left            =   0
      MouseIcon       =   "F_EVALUATION.frx":8B655
      MousePointer    =   99  'Custom
      TabIndex        =   62
      Top             =   10680
      Width           =   1935
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Index           =   1
      Left            =   8280
      MouseIcon       =   "F_EVALUATION.frx":8B95F
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   10560
      Width           =   2655
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   3
      Left            =   3960
      MouseIcon       =   "F_EVALUATION.frx":8BC69
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   10680
      Width           =   1815
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   0
      Left            =   17640
      MouseIcon       =   "F_EVALUATION.frx":8BF73
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   10680
      Width           =   1575
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   2
      Left            =   15240
      MouseIcon       =   "F_EVALUATION.frx":8C27D
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   10680
      Width           =   2175
   End
   Begin VB.Label La 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move Forward"
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
      Index           =   2
      Left            =   17655
      MouseIcon       =   "F_EVALUATION.frx":8C587
      MousePointer    =   99  'Custom
      TabIndex        =   75
      Top             =   11600
      Width           =   1200
   End
   Begin VB.Label La 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move Previous"
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
      Index           =   1
      Left            =   15630
      MouseIcon       =   "F_EVALUATION.frx":8C891
      MousePointer    =   99  'Custom
      TabIndex        =   74
      Top             =   11600
      Width           =   1230
   End
   Begin VB.Label La 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Procedure"
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
      Index           =   0
      Left            =   9000
      MouseIcon       =   "F_EVALUATION.frx":8CB9B
      MousePointer    =   99  'Custom
      TabIndex        =   73
      Top             =   11600
      Width           =   1200
   End
   Begin VB.Label La 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reading QC Info"
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
      Left            =   4140
      MouseIcon       =   "F_EVALUATION.frx":8CEA5
      MousePointer    =   99  'Custom
      TabIndex        =   72
      Top             =   11600
      Width           =   1290
   End
   Begin VB.Label lbOperator 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
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
      Left            =   300
      TabIndex        =   63
      Top             =   11600
      Width           =   2445
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   4
      Left            =   360
      MouseIcon       =   "F_EVALUATION.frx":8D1AF
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":8D4B9
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   15960
      MouseIcon       =   "F_EVALUATION.frx":9089B
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":90BA5
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_EVALUATION.frx":93F87
      Height          =   480
      Index           =   1
      Left            =   9360
      MouseIcon       =   "F_EVALUATION.frx":97369
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":97673
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   4560
      MouseIcon       =   "F_EVALUATION.frx":9AA55
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":9AD5F
      Top             =   11040
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   480
      X2              =   18720
      Y1              =   10680
      Y2              =   10680
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
      TabIndex        =   2
      Top             =   10200
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   14400
      X2              =   14400
      Y1              =   0
      Y2              =   11880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   4800
      X2              =   4800
      Y1              =   0
      Y2              =   11880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   9600
      X2              =   9600
      Y1              =   120
      Y2              =   12000
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_EVALUATION.frx":9E141
      Height          =   480
      Index           =   0
      Left            =   18240
      MouseIcon       =   "F_EVALUATION.frx":A1523
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":A182D
      Top             =   11040
      Width           =   480
   End
End
Attribute VB_Name = "F_EVALUATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IndexFormProcedura As Integer
Private IndexMainProcedura As Integer
Private IndexTextSelected As Integer
Private IndexText As Integer
Private MyLot As String
Private MyCode As String
Private m_rc As Boolean
Private bFormSaved As Boolean
Private CodeID As Long
Private MeasurementUnit As String
Private UserDecimal As String
Private intDecimal As Integer
Private STD() As String
Private MeterNumber As Integer
Private pHNumber As Integer
Private pHMin(3) As String
Private pHMax(3) As String
Private MinNumber80Pec(10) As Integer
Private SelectedSTDNumber As Integer
Private strPassed As String
Private numSelectedStandard As String
Private AndOr As String
Private ReadinglngID As Long
Private lCol As Long
Private lRow As Long
Private bHoChiusoilLotto As Boolean
Private bAnotherFormCalled As Boolean
Private STDCount As Integer
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
Public Function DoShow(ByRef Index As Integer, Optional ByRef sLot As String, Optional ByRef sCode As String, Optional ByVal lngID As Long, Optional MyImage As Image, Optional FileName As String) As Boolean

    On Error GoTo ERR_SHOW
    
    'Set DefaultMenu(4) = MyImage
    IndexMainProcedura = Index
    m_rc = False
    bFormSaved = False
    
    
    


    
    
    SettingName = FileName
    ReadinglngID = lngID
    MinNumberSelecterPerc = GetSetting(App.Title, "Options", "MinNumberSelecterPerc", 0.7)
    
    
    Call GrdRisultati(Grd3)
    
    CheckUser 1
    FormPulisciTutto

            

    
    If sLot <> "" And sCode <> "" Then
        Call GetCodeInformation(sLot, sCode, lngID)
    Else
        PopupMessage 2, "Please select a valid Code/Lot..."
        Unload Me
    End If
    
    SelectProcedura 0
    mOk
    
    
    If MyOperatore.IndexPrivilege >= 1 Then
            
        Text1(11).Locked = False
        Text1(12).Locked = False
        Text1(12) = MyOperatore.Name
    
    End If
    
    
    
    Me.Show vbModal
    
    If m_rc = True Then
        Index = IndexMainProcedura
        sLot = MyLot
        sCode = MyCode
    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    MsgBox err.Description
    Resume ERR_END
End Function






Private Sub DefaultMenuLabel_Click(Index As Integer)
Dim MyIndex As Integer
Select Case Index
    Case 0
        ' vai avanti
        If IndexFormProcedura = 1 Then
            MyIndex = 0
        Else
            MyIndex = IndexFormProcedura + 1
        End If
        PicMenu_Click MyIndex
    Case 1
        'If bFormSaved Then
        If Not (bHoChiusoilLotto) Then
            SaveSelectedTest
            
            SaveResultsTable
        End If
            USER_PATH = USER_TEMP_PATH
            
            Unload Me
       ' Else
            
       ' End If
    Case 2
        ' torna indietro
        If IndexFormProcedura = 0 Then
            MyIndex = 1
        Else
            MyIndex = IndexFormProcedura - 1
        End If
        PicMenu_Click MyIndex
    Case 3
        Frame2.Visible = Not (Frame2.Visible)
    Case 4
         frmLogin.DoShow 1
        CheckUser 1
    Case 5
       ' Label7_Click
    Case 6
       ' Label6_Click
End Select

End Sub


Private Function CheckUser(ByVal Index As Integer)
Dim rc As Boolean
     
    rc = True
    Text1(11).Locked = True
    Text1(12).Locked = True
            
    If MyOperatore.IndexPrivilege < 1 Then rc = False
    If MyOperatore.Name = "" Then rc = False
    lbOperator = MyOperatore.Name
    Picture2.Visible = rc
    
    
    If Not (Grd3.Rows > 1) Then
        Picture2.Visible = False
    End If
    
    
    If rc Then
        If Grd3.Rows > 1 Then
            Text1(12) = MyOperatore.Name
            If Text1(11) = "" Then
                Text1(11) = FormatDataLAT(Now)
            Else
            
            End If
            Text1(11).Locked = False
            Text1(12).Locked = False
            
            
        End If
    Else
         Text1(12) = ""
    End If
    
    
End Function




Private Sub DisableImage_Click()
PopupMessage 2, "Warning : Administrator Only can Operate...", , True
End Sub

Private Sub Form_Initialize()

Call SetPicForm
Call SetGrid
SaveSizes
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37
        DefaultMenuLabel_Click 2
    Case 39
        DefaultMenuLabel_Click 0
End Select
End Sub
Private Sub FormPulisciTutto()
Dim Ctl As Control
    For Each Ctl In Controls
        If TypeOf Ctl Is TextBox Then
            Ctl = ""
        ElseIf TypeOf Ctl Is Label Then
            If InStr(Ctl.Caption, "SET") Then
            Else
            Ctl.BackColor = vbColorLabelUnabled
            End If

        ElseIf TypeOf Ctl Is Grid Then
            Ctl.Rows = 1
        End If
    Next Ctl
    
   ' Picture3.BackColor = vbColorLabelUnabled
End Sub

Private Sub Form_Load()
IndexFormProcedura = 99
Dim i As Integer
If Screen.Width - Me.Width > 1000 And bFullScreen Then
    Me.WindowState = 2
    For i = 0 To PicMain.Count - 1
        PicMain(i).Picture = LoadPicture(PictureMaxScreen)
        
    Next '
    'Me.Picture = LoadPicture(PictureMaxScreen)
End If
End Sub

Private Sub Form_Resize()

ResizeControls
End Sub

Private Sub Frame2_Click()
Frame2.Visible = False
End Sub

Private Sub grd1_Click()
Frame2.Visible = False
End Sub

Private Sub Grd1_LostFocus()
'Grd1.Cell(0, 0).SetFocus
End Sub

Private Sub Grd1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim t As Integer
    SelectedSTDNumber = 0
    strPassed = ""
    Picture1.Visible = False
    PicMenu(3).Visible = False
If FirstRow > 0 Then
    Call SaveSelectedTest
    SelectedSTDNumber = CInt((Grd1.Cell(FirstRow, 1).Text))
    t = CInt((Grd1.Cell(FirstRow, 3).Text))
    numSelectedStandard = t 'STD(t, 1)
    Text1(7) = STD(t, 2)
    Text1(8) = STD(t, 3)
    Text1(9) = ""
    Text1(10) = ""
    DoEvents
    
    Call FillReadingsTablePerStandard(MeterNumber, SelectedSTDNumber)
    
    PicMenu(3).Visible = True
    Picture1.Visible = True
End If

End Sub

Private Sub Grd2_Click()
With Grd2
   Select Case lCol
        Case 1, 3, 5, 7, 9, 11, 13, 15, 17
            SetMeanValue (SelectedSTDNumber)
    End Select
End With
End Sub

Private Sub CheckpHTrue(ByVal t As Long)

Dim i As Integer
Dim rc1 As Boolean
Dim rc2 As Boolean
Dim rc3 As Boolean
Dim bValue As OLE_COLOR

    ' se è pH-----------------------------
    
    rc1 = True
    rc2 = True
    rc3 = True
    
    With Grd2
        .AutoRedraw = False
        rc1 = .Cell(t, MeterNumber * 2 + 1).Text
        rc2 = .Cell(t, MeterNumber * 2 + 3).Text
        rc3 = .Cell(t, MeterNumber * 2 + 5).Text
    
        If rc1 And rc2 And rc3 Then
    
            Else
            
            For i = 1 To MeterNumber
                .Cell(t, i * 2).BackColor = vbColorLabelUnabled
                .Cell(t, i * 2 - 1).BackColor = vbColorLabelUnabled
                .Cell(t, 0).BackColor = vbColorLabelUnabled
            Next
            DoEvents
            
        End If
        .AutoRedraw = True
        .Refresh
    End With

End Sub

Private Sub Grd2_LeaveCell(ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
    'Select Case Col
    '    Case 1, 3, 5, 7, 9, 11
    '        SetMeanValue (SelectedSTDNumber)
    'End Select
End Sub

Private Sub Grd2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
lCol = 0
lRow = 0
If FirstCol > 0 Then lCol = FirstCol
If FirstRow > 0 Then lRow = FirstRow
With Grd2
    Select Case FirstCol
        Case MeterNumber * 2 + 1, MeterNumber * 2 + 2
           ' lb(9) = "pH " & .Cell(FirstRow, MeterNumber * 2 + 5).Text & " Range"
           ' Text1(9) = .Cell(FirstRow, MeterNumber * 2 + 6).Text
           ' Text1(10) = .Cell(FirstRow, MeterNumber * 2 + 7).Text
        Case Else
    End Select

End With



End Sub

Private Sub Grd2_SetCellText(ByVal Row As Long, ByVal Col As Long, ByVal Text As String, Cancel As Boolean)
   ' Select Case Col
   '     Case 1, 3, 5, 7, 9, 11
   '         SetMeanValue (SelectedSTDNumber)
   ' End Select
End Sub

Private Sub Image2_Click()
   CheckPrivilege 1
   CheckUser 1
End Sub

Private Sub ImageTAV_Click(Index As Integer)
Select Case Index
    Case 5
        Picture1_Click
End Select
End Sub

Private Sub Label3_Click()

    Picture2_Click

End Sub

Private Sub PicMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Len(Text1(IndexText)) = 0 Then Text1(IndexText).BackColor = vbColorUnabled
IndexText = 0
Picture1.BackColor = &H4000&
Picture2.BackColor = &H4000&
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To PicMenu.Count - 1
    If i = IndexFormProcedura Then
    Else
        PicMenu(i).BackColor = vbColorDarkFont
    End If
Next
Picture1.BackColor = &H4000&
Picture2.BackColor = &H4000&
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set F_EVALUATION = Nothing
End Sub

Private Sub Image3_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub Label1_Click(Index As Integer)
Dim rc As Boolean

    Select Case Index
        Case 14
            rc = True
        Case 15
            rc = False
        Case Else
            Exit Sub
    End Select
    
    Label1(14).BackColor = IIf(rc, Picture4(0).BackColor, &H808080)
    Label1(15).BackColor = IIf(Not (rc), Picture4(0).BackColor, &H808080)
    
End Sub

Private Sub Label2_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub PicMenu_Click(Index As Integer)
    If IndexFormProcedura = Index Then
    Else
    
        Call SelectProcedura(Index)
    
        
    End If
End Sub


Private Function SelectProcedura(ByVal Index As Integer) As Boolean
Dim rc As Boolean
Dim i As Integer

    'If Index > 3 Then Exit Function
    For i = 0 To PicMenu.Count - 1
        If i = Index Then
            PicMenu(i).BackColor = vbColorForeFixed
        Else
            PicMenu(i).BackColor = vbColorDarkFont
            
        End If
    Next
    Set Image4(0) = Image3(Index)
    Select Case Index
        Case 0
            Picture4(0).BackColor = &H4080&
           ' rc = False
           Frame3.Visible = True
           Frame4.Visible = False
        Case 1
           ' rc = True
            
            Frame4.Visible = True
            Picture4(0).BackColor = &H4060&
            Frame3.Visible = False
            lb(1).BackColor = Picture4(0).BackColor
            If CheckRequestedFiles Then
            
            End If
        Case 2
           ' rc = False
           ' Picture4(0).BackColor = &H60DF&
           ' reading QC
           'PopupMessage 2, "Open Reading form..."
           OpenReadingQC
           Exit Function
           
        Case 3
            ' GRAPH QC
            If Grd2.Rows > 1 Then
                 bAnotherFormCalled = True
                SaveSelectedTest
                
                SaveResultsTable
                
                F_GRAPH.Top = Me.Top
                F_GRAPH.Left = Me.Left
                F_GRAPH.DoShow IndexFormProcedura, Text1(0), Text1(1), ReadinglngID, , SettingName, CStr(SelectedSTDNumber)
                PicMenu_Click 0
            
            Else
                PopupMessage 2, "Please Select a valid Standard with at least 1 Test to open Graph QC...", , True
            End If
            bAnotherFormCalled = False
            Exit Function
            
    End Select
    Label2(4) = Label2(Index)
    IndexFormProcedura = Index
    PicMain(Index).Visible = True
    PicMain(Index).ZOrder
   ' blTable = Label2(IndexFormProcedura)
    Cleanform (False)
End Function

Private Sub Cleanform(ByVal bValue As Boolean, Optional ByVal Index As Integer = 0)

End Sub

Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To PicMenu.Count - 1
    If i = IndexFormProcedura Then
        ' lascialo stare...
    ElseIf i = Index Then
        PicMenu(i).BackColor = &H505050
    Else
        PicMenu(i).BackColor = vbColorDarkFont
    End If
Next
End Sub


Private Sub SetPicForm()
Dim i As Integer


For i = 0 To PicMain.Count - 1
    PicMain(i).Left = 0
    PicMain(i).Top = PicMenuBar(0).Height + Frame1.Height
    PicMain(i).Width = Me.Width
    PicMain(i).Height = Line1.Y1 - PicMain(i).Top
Next


For i = 0 To Text1.Count - 1
    Text1(i).BackColor = IIf(Len(Text1(i)) > 0, vbWhite, vbColorUnabled)
Next
End Sub

Private Sub Picture1_Click()
' salva i risultati calcolati
Call SaveResults
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BackColor = &H8000&
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SaveProcedure
End Sub

Private Sub Picture2_Click()
' chiudi il lotto
If Grd3.Rows < 2 Then
    MessageInfoTime = 2000
    PopupMessage 2, "This Lot cannot be closed : Please save at least 1 Mean Value..."
Else
    Call CloseLot
End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.BackColor = &H8000&
End Sub

Private Sub Text1_Change(Index As Integer)
Dim rc As Boolean
rc = IIf(Len(Text1(Index)) > 0, True, False)
Text1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbWhite, vbColorUnabled)
Label1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextBlue, vbColorLabelUnabled)

Select Case Index
    Case 0
        lb(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextDarkBlue, vbColorLabelUnabled)
    Case 2, 4
        ' SELECTED MEAN!
        
        
        
    Case 7
       lb(7).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextDarkBlue, vbColorLabelUnabled)

    Case 9
       lb(9).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextDarkBlue, vbColorLabelUnabled)

End Select

End Sub
Private Sub Text1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = IndexText Then
   '
Else
    If Len(Text1(Index)) = 0 Then Text1(Index).BackColor = vbColorgotFocus
    If Len(Text1(IndexText)) = 0 Then Text1(IndexText).BackColor = vbColorUnabled

End If
IndexText = Index
End Sub
Private Sub Text1_Click(Index As Integer)
Text1(Index).BackColor = vbWhite
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index < Text1.Count - 1 Then Text1(Index + 1).SetFocus
End If
End Sub
Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).BackColor = vbWhite
ctlCalendar1.Visible = False
Select Case Index
    Case 11
        ctlCalendar1.ZOrder
        ctlCalendar1.Visible = True
End Select
End Sub

Private Sub ctlCalendar1_DateClicked(inputDate As Date)
'Select Case IndexTextSelected
  'Case 11
    Text1(11) = FormatDataLAT(CStr(inputDate))
'Case Else
'End Select
ctlCalendar1.Visible = False
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbWhite, vbColorUnabled)
End Sub


Private Function SaveProcedure()
Dim rc As Boolean

    rc = SaveForm
    'PicMenu(3).Visible = rc
    If rc Then
        MyLot = Text1(0)
        MyCode = Text1(1)
        PopupMessage 2, blTable & " Saved..."
    Else
    
    End If
    
    m_rc = rc
    bFormSaved = rc

End Function
Private Function SaveForm() As Boolean

Dim rc As Boolean
On Error GoTo ERR_SAVE:

    rc = True
    
ERR_END:
    On Error GoTo 0
    SaveForm = rc
    Exit Function
ERR_SAVE:
    rc = False
    GoTo ERR_END:
End Function

Private Function SetGrid()

       '------------------------------------------------
        '       SET TABELLA STD
        '------------------------------------------------
    With Grd1
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        .DefaultFont.Size = 12 '* m_ControlGridFontSize
        .DefaultRowHeight = 40
        .Cols = 4
        .Cell(0, 0).Text = "n."
        .ReadOnly = False
        .Column(0).Width = 0
        .Range(0, 1, 0, 2).Merge
        .Cell(0, 1).Text = "Standard"
        .Column(1).Width = 129
        '.Cell(0, 2).Text = "STD Value"
        .Column(2).Width = 150
        .Cell(0, 3).Text = "t"
        .Column(3).Width = 0
        .Cell(0, 1).BackColor = vbColorTextLightBlue
        '.Cell(0, 1).BackColor = vbColorTextLightBlue
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With

End Function

Private Sub SetGrd2(ByVal Index As Integer)
 With Grd2
      .Rows = 1

        .AutoRedraw = False
        .ReadOnly = False
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .DefaultFont.Size = 12
        .DefaultRowHeight = 40
        .Cols = 20
        .FixedCols = 0
        
        .Cell(0, 0).Text = "# Test"
        .Column(0).Width = 65
        .Cell(0, 0).BackColor = vbColorTextLightBlue
        .Range(0, 1, 0, 2).Merge
        .Cell(0, 1).Text = "Meter 1"
        .Column(1).CellType = cellCheckBox
        .Column(1).Width = 30
        .Column(2).Width = 100
        .Cell(0, 1).BackColor = vbColorTextLightBlue

        If Index > 1 Then
            .Range(0, 3, 0, 4).Merge
            .Cell(0, 3).Text = "Meter 2"
            .Column(3).CellType = cellCheckBox
            .Column(3).Width = 30
            .Column(4).Width = 100
            .Cell(0, 3).BackColor = vbColorTextLightBlue

             If Index > 2 Then
                .Range(0, 5, 0, 6).Merge
                .Cell(0, 5).Text = "Meter 3"
                .Column(5).CellType = cellCheckBox
                .Column(5).Width = 30
                .Column(6).Width = 100
                .Cell(0, 5).BackColor = vbColorTextLightBlue
                 If Index > 3 Then
                    .Range(0, 7, 0, 8).Merge
                    .Cell(0, 7).Text = "Meter 4"
                    .Column(7).CellType = cellCheckBox
                    .Column(7).Width = 30
                    .Column(8).Width = 100
                    .Cell(0, 7).BackColor = vbColorTextLightBlue
                End If
            End If
        End If
    
    
    ' pH1
    
        If pHMin(1) <> "" Then
            .Range(0, Index * 2 + 1, 0, Index * 2 + 2).Merge
            .Cell(0, Index * 2 + 1).Text = "pH 1"
            .Column(Index * 2 + 1).CellType = cellCheckBox
            .Column(Index * 2 + 1).Width = 30
            .Column(Index * 2 + 2).Width = 80
        Else
            .Column(Index * 2 + 1).Width = 0
            .Column(Index * 2 + 2).Width = 0
        End If
        
    ' pH2
            
        If pHMin(2) <> "" Then
            .Range(0, Index * 2 + 3, 0, Index * 2 + 4).Merge
            .Cell(0, Index * 2 + 3).Text = "pH 2"
            .Column(Index * 2 + 3).CellType = cellCheckBox
            .Column(Index * 2 + 3).Width = 30
            .Column(Index * 2 + 4).Width = 80
        Else
            .Column(Index * 2 + 3).Width = 0
            .Column(Index * 2 + 4).Width = 0
        End If
        
    ' pH3
              
        If pHMin(3) <> "" Then
            .Range(0, Index * 2 + 5, 0, Index * 2 + 6).Merge
            .Cell(0, Index * 2 + 5).Text = "pH 3"
            .Column(Index * 2 + 5).CellType = cellCheckBox
            .Column(Index * 2 + 5).Width = 30
            .Column(Index * 2 + 6).Width = 80
        Else
            .Column(Index * 2 + 5).Width = 0
            .Column(Index * 2 + 6).Width = 0
        End If

              
        
        
        
        .Cell(0, Index * 2 + 3 + (pHNumber - 1) * 2).Text = "Index"
        .Column(Index * 2 + 3 + (pHNumber - 1) * 2).Width = 0
        .Cell(0, Index * 2 + 4 + (pHNumber - 1) * 2).Text = "Type"
        .Column(Index * 2 + 4 + (pHNumber - 1) * 2).Width = 0
        
        .Cols = Index * 2 + 5 + (pHNumber - 1) * 2
        
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .AutoRedraw = True
        .Refresh
    End With
    
End Sub


Private Function GetFormSettingName()
Dim i As Integer
Dim t As Integer

    ' doppio check... si sa mai...
    
   If FileExists(USER_TEMP_PATH & SettingName) Then
   ElseIf FileExists(USER_DATA_PATH & SettingName) Then
        'PopupMessage 2, "Lot : " & Text1(0) & vbCrLf & "Code : " & Text1(1) & vbCrLf & "Is Closed..."
        USER_PATH = USER_DATA_PATH
   Else
        PopupMessage 2, "Lot : " & Text1(0) & vbCrLf & "Code : " & Text1(1) & vbCrLf & "Warning : Information QC not found...", , True

   End If
   
   
    CloseSettingDataFile
    
    CodeID = GetSettingData(SettingName, "Code Information", "ID", 0)
    
    SelectedSTDNumber = 0
    
    ' get variables
    
    UserDecimal = FormatDecimal(GetSettingData(SettingName, "Code Information", "Decimal", 0))
    intDecimal = GetSettingData(SettingName, "Code Information", "Decimal", 0)
    MeasurementUnit = IIf(IsNull(Trim(dbTabCode!MeasurementUnit)), "", Trim(dbTabCode!MeasurementUnit))
    MeterNumber = GetSettingData(SettingName, "Information QC", "MeterNumber", 1)
    pHNumber = GetSettingData(SettingName, "Information QC", "pHNumber", 1)
    
    Text1(11) = GetSettingData(SettingName, "Close QC", "Validation Date", "")
    Text1(12) = GetSettingData(SettingName, "Close QC", "Operator", "")
    
    Call GetphValue
    
    ' Riempi tabella STANDARD
    
     Call GetSTDTable
     
    ' imposta tabella Meter
    
     SetGrd2 (MeterNumber)
    
    GetMeanTable Grd3, SettingName

    CloseSettingDataFile
End Function
Private Sub GetCodeInformation(ByVal sLot As String, ByVal sCode As String, Optional ByVal MyID As Long)
Dim MeasurementUnit As String
    ' attenzione , se ho un file allora lo importo , altrimenti prendo i dati del Code
    
     MyID = GetSettingData(SettingName, "Code Information", "ID", 0)
     
    With dbTabCode
        .filter = ""
        If MyID > 0 Then
            .filter = "ID='" & MyID & "'"
        Else
        
            .filter = "Code='" & sCode & "'"
        End If
    
        Select Case UCase(Trim(!AndOr))
        Case "&"
            AndOr = Chr$(177)
        Case UCase("or")
            AndOr = Chr$(247)
        Case Else
            AndOr = ""
        End Select
            
        If .EOF Then
            MessageInfoTime = 2000
            PopupMessage 2, "Cannot find Hanna Code  :  " & sCode & vbCrLf & "Please Enter Code Information..."
            'Unload Me
        Else
        
            MeasurementUnit = IIf(IsNull(Trim(!MeasurementUnit)), "", " " & Trim(!MeasurementUnit))
            Text1(0) = sLot
            Text1(1) = sCode
          
            
            GetFormSettingName


    
        End If
    
    End With




End Sub


Private Sub GetSTDTable()
    Dim rc As Boolean
    Dim t As Integer
    
    Dim STDNumber As String
    Dim STDValue As String
    Dim STDMin As String
    Dim STDMax As String
    
    
    On Error GoTo ERR_GET
    
    
    rc = True
    
    STDCount = CInt(GetSettingData(SettingName, "Graph QC", "STDCount", 0))
    
    If STDCount = 0 Then
        rc = False
        GoTo ERR_END
    End If
    
    ReDim STD(STDCount, 4) As String
        
    For t = 1 To STDCount

        STD(t, 0) = GetSettingData(SettingName, "Graph QC", "STDNumber" & t, "")
        STD(t, 1) = GetSettingData(SettingName, "Graph QC", "STDValue" & t, "")
        STD(t, 2) = GetSettingData(SettingName, "Graph QC", "STDMin" & t, "")
        STD(t, 3) = GetSettingData(SettingName, "Graph QC", "STDMax" & t, "")
    Next
    
    
ERR_END:
    On Error GoTo 0
    
    CloseSettingDataFile
    
    If rc Then
        Call FillSTDGrid
    End If
    Exit Sub
ERR_GET:
    rc = False
    MsgBox err.Description
    Resume Next
End Sub


Private Function FillSTDGrid()

Dim i As Integer
Dim t As Integer
With Grd1
    .Rows = 1
    .ReadOnly = True
    .AutoRedraw = False
    t = 1
    For t = 1 To STDCount
        .AddItem "", False
        .Cell(.Rows - 1, 1).Text = STD(t, 0)
        .Cell(.Rows - 1, 2).Text = STD(t, 1)
        .Cell(.Rows - 1, 3).Text = t
        .Cell(.Rows - 1, 1).ForeColor = vbColorDarkUnabled
        .Cell(.Rows - 1, 2).ForeColor = vbColorDarkUnabled ' vbColorForeFixed
       
    Next
    .Column(1).Sort cellAscending
    .AutoRedraw = True
    .Refresh
End With

End Function




Private Function FillReadingsTablePerStandard(ByVal NumMeter As Integer, ByVal STDNumber As Integer) As Boolean
Dim rc As Boolean
Dim AllTests As String
Dim MyCols As String
Dim MeterValue As String
Dim sString As String
Dim ReadingSelected As Boolean
Dim pHValue As String
Dim MinValue As String
Dim MaxValue As String
Dim MyColor As OLE_COLOR
Dim i As Integer
Dim t As Integer
Dim RowCount As Integer
Dim ReadingCount As Integer
Dim STDNumberGrid As String
Dim NumberTest As String
Dim TestType As String
On Error GoTo ERR_GET
    rc = True
    
    Grd2.Rows = 1

    AllTests = CInt(GetSettingData(SettingName, "Graph QC", "Standard " & STDNumber & " Total Tests", 1))

    MinValue = Text1(7)
    MaxValue = Text1(8)
            
    If AllTests >= 1 Then
      
        With Grd2
            .DefaultFont.Size = 12 ' * m_ControlGridFontSize
            .AutoRedraw = False
            
            RowCount = 0
            ReadingCount = 0
            
            For i = 1 To AllTests
            
                    .AddItem "", False
                    '.Cell(.Rows - 1, t * 2 - 1).Text =
                    
                       NumberTest = GetSettingData(SettingName, "Graph QC", "Standard " & STDNumber & " Test " & i & " Real Test", "")
                       SaveSettingData SettingName, "Reading QC", "Grd2 Row" & NumberTest & " STD Test", i
                        .Cell(.Rows - 1, 0).Text = NumberTest
                    For t = 1 To NumMeter
                         
                         MeterValue = GetSettingData(SettingName, "Graph QC", "Standard " & STDNumber & " Test " & i & " Meter " & t & " Value", "")
                         ReadingSelected = GetSettingData(SettingName, "Graph QC", "Standard " & STDNumber & " Test " & i & " Meter " & t & " Selected", "TRUE")
                         
                         Call CheckValue(MeterValue, MinValue, MaxValue, MyColor)
                         
                         If MeterValue <> "" Then
                            ReadingCount = ReadingCount + 1
                            .Cell(.Rows - 1, t * 2 - 1).Text = ReadingSelected
                        Else
                            .Cell(.Rows - 1, t * 2 - 1).Text = False
                            .Cell(.Rows - 1, t * 2 - 1).Locked = True
                        End If
                        
                        .Cell(.Rows - 1, t * 2).Text = MeterValue       ' GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & 10 + t, "")
                        .Cell(.Rows - 1, t * 2).ForeColor = MyColor     ' GetSettingData(SettingName, "Reading QC", "Grd2 Fore" & i & 10 + t, vbBlack)
                        .Cell(.Rows - 1, t * 2).Locked = True
                        If IsNumeric(NumberTest) Then
                            SaveSettingData SettingName, "Reading QC", "Grd2 Fore Row" & NumberTest & " Col" & 10 + t, MyColor
                            
                        End If
                    Next
                    
                    
                    For t = 0 To 2
                      
                      pHValue = GetSettingData(SettingName, "Graph QC", "Standard " & STDNumber & " Test " & i & " pH" & t + 1 & " " & " Value", "")
                      
                      Call CheckpHValue(pHValue, pHMin(t + 1), pHMax(t + 1), MyColor)
                      
                      .Cell(.Rows - 1, NumMeter * 2 + 1 + t * 2).Text = True
                      .Cell(.Rows - 1, NumMeter * 2 + 2 + t * 2).Text = pHValue
                      .Cell(.Rows - 1, NumMeter * 2 + 2 + t * 2).ForeColor = MyColor
                      .Cell(.Rows - 1, NumMeter * 2 + 2 + t * 2).Locked = True
                    Next
            Next
            
            MinNumber80Pec(STDNumber) = Int(ReadingCount * MinNumberSelecterPerc) '  Int((.Rows - 1) * NumMeter * MinNumberSelecterPerc)  ' STDNumber=0 80% dei totali
            
            .AutoRedraw = True
            .ReadOnly = False
            .Refresh
        End With
     
    End If
    
    

    
    
ERR_END:
    On Error GoTo 0
    
    lb80Pec = "Select at least " & MinNumberSelecterPerc * 100 & "% of all value : " & MinNumber80Pec(STDNumber)
    
    If rc Then
        If STDNumber > 0 Then Call SetMeanValue(STDNumber)
    End If
    FillReadingsTablePerStandard = rc
    Exit Function
ERR_GET:
    rc = False
  
    MsgBox err.Description
    Resume Next
    

    
End Function
Public Function CheckpHValue(ByVal Value As String, ByVal MinValue As String, ByVal MaxValue As String, ByRef MyColor As OLE_COLOR)

    If IsNumeric(Value) Then
        If IsNumeric(MinValue) And IsNumeric(MaxValue) Then
            If CDbl(Value) > 0 Then
                If CDbl(Value) >= CDbl(MinValue) And CDbl(Value) <= CDbl(MaxValue) Then
                    MyColor = vbBlack
                    
                Else
                    MyColor = vbRed
                   
                End If
            End If
        Else
            MyColor = vbBlack
        End If
    End If
End Function

Public Function CheckValue(ByVal MeterValue As String, ByVal MinValue As String, ByVal MaxValue As String, ByRef MyColor As OLE_COLOR)

    If MinValue = 0 And MaxValue = 0 Then GoTo OK:
    If IsNumeric(MeterValue) Then
        If IsNumeric(MinValue) And IsNumeric(MaxValue) Then
            If CDbl(MeterValue) > 0 Then
                If CDbl(MeterValue) >= CDbl(MinValue) And CDbl(MeterValue) <= CDbl(MaxValue) Then
                    MyColor = vbBlack
                    
                Else
                    MyColor = vbRed
                   
                End If
            End If
        Else
OK:
            MyColor = vbBlack
        End If
    End If
End Function

Private Function SetMeanValue(ByVal STDNumber As Integer) As Boolean
Dim sSTDNumerString As String
Dim ReadingSelected As Boolean
Dim SelectedAverage As Double
Dim SelectedValue As Double
Dim ReadValue As Double
Dim SelectedCount As Integer
Dim TotalAverage As Double
Dim TotalReadingsCount As Integer
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim pHTrue As Boolean
Dim pHTrue1 As Boolean
Dim pHTrue2 As Boolean
Dim pHTrue3 As Boolean
Dim SelRows() As Integer
    On Error GoTo ERR_SET
    
    If Me.Visible = False Then Exit Function

    rc = True
    
    ReDim SelRows(MeterNumber) As Integer
    
    
    If STDNumber = 0 Then ' totali
        sSTDNumerString = "STD All"
        lbSTDNumber.BackColor = vbColorDarkUnabled
      '  Picture3.BackColor = lbSTDNumber.BackColor
        Picture1.Visible = False
    Else
        sSTDNumerString = "STD Number : " & STDNumber
        lbSTDNumber.BackColor = &H4080&
       ' Picture3.BackColor = lbSTDNumber.BackColor
       Picture1.Visible = True
    End If


   ' SelectedAverage
    SelectedCount = 0
    SelectedAverage = 0
    SelectedValue = 0
    With Grd2
        .AutoRedraw = False
        If .Rows = 1 Then
            rc = False
            GoTo ERR_END
        End If
        For i = 1 To .Rows - 1
            
            For t = 1 To MeterNumber
                If Trim(.Cell(i, t * 2).Text) = "" Or IsNull(.Cell(i, t * 2).Text) Then
                    GoTo salta
                Else
                    ReadValue = CDbl(.Cell(i, t * 2).Text)
                End If
                
                TotalAverage = TotalAverage + ReadValue
                TotalReadingsCount = TotalReadingsCount + 1
                
                

                pHTrue1 = (.Cell(i, MeterNumber * 2 + 1).Text)
                pHTrue2 = (.Cell(i, MeterNumber * 2 + 3).Text)
                pHTrue3 = (.Cell(i, MeterNumber * 2 + 5).Text)
                pHTrue = IIf(pHTrue1 And pHTrue2 And pHTrue3, True, False)

                
                ReadingSelected = .Cell(i, t * 2 - 1).Text = True
                
                If ReadingSelected Then ' se ho la spunta lo conto
                    If .Cell(i, t * 2).Text <> "" Then
                        If pHTrue Then ' controllo il PH! se ho la spunta conto tutto....
                            SelRows(t) = SelRows(t) + 1
                            SelectedValue = CDbl(.Cell(i, t * 2).Text)
                            SelectedCount = SelectedCount + 1
                            SelectedAverage = SelectedAverage + SelectedValue
                        End If
                    End If
                    .Cell(i, t * 2 - 1).BackColor = vbWhite '&HF0F0F0 '
                    .Cell(i, t * 2).BackColor = vbWhite '&HF0F0F0 '
                    .Cell(i, 0).BackColor = vbWhite '&HF0F0F0 '
                Else
                    .Cell(i, t * 2 - 1).BackColor = vbColorMediumFixed
                    .Cell(i, t * 2).BackColor = vbColorMediumFixed
                    .Cell(i, 0).BackColor = vbColorMediumFixed '&HF0F0F0 '
                End If
salta:
                .Cell(0, t * 2).BackColor = vbColorLightFixed
            Next
            CheckpHTrue i
        Next
        .AutoRedraw = True
        .Refresh
    End With
    If SelectedCount = 0 Then
        SelectedAverage = 0
    Else
        SelectedAverage = SelectedAverage / SelectedCount
    End If
    TotalAverage = TotalAverage / TotalReadingsCount '((Grd2.Rows - 1) * MeterNumber)
    
    
    If MinNumber80Pec(STDNumber) <= SelectedCount Then ' OK
    
    Else
        rc = False
    End If
    Dim Max As Integer
    For i = 0 To UBound(SelRows)
        If SelRows(i) > Max Then
            Max = SelRows(i)
        End If
    Next
ERR_END:
    On Error GoTo 0
    lbSTDNumber = sSTDNumerString
    Text1(2) = FormatNumber(TotalAverage, intDecimal)
    Text1(2) = Format$(Text1(2), UserDecimal)
    
    Text1(3) = TotalReadingsCount '* MeterNumber
    Text1(13) = (Grd2.Rows - 1)
    Text1(4) = FormatNumber(SelectedAverage, intDecimal)
    Text1(4) = Format$(Text1(4), UserDecimal)
    Text1(5) = SelectedCount
    Text1(5).ForeColor = IIf(rc, vbBlack, vbRed)
    Text1(6) = Max
    
    Call CheckValuePassed(2)
    Call CheckValuePassed(4)
   

        
    Picture1.Visible = rc
    SetMeanValue = rc
    Exit Function
ERR_SET:
    rc = False
  
    MsgBox err.Description
    Resume Next
End Function

Private Sub CheckValuePassed(ByVal Index As Integer)

'Debug.Print CDbl(Text1(Index)), CDbl(Text1(7)), CDbl(Text1(8))
    If Text1(7) <> "" And Text1(8) <> "" And Text1(7) <> "/" And Text1(8) <> "/" Then
        If CDbl(Text1(Index)) >= CDbl(Text1(7)) And CDbl(Text1(Index)) <= CDbl(Text1(8)) Then
            Text1(Index).ForeColor = vbBlack
            strPassed = "YES"
        Else
            Text1(Index).ForeColor = vbRed
            strPassed = "NO"
        End If
    End If
    
End Sub

Private Function SaveSelectedTest()
Dim i As Integer
Dim t As Integer
Dim rc As Boolean

    CloseSettingDataFile
    
    If SelectedSTDNumber = 0 Then Exit Function
    
    
    With Grd2
        For t = 1 To .Rows - 1
            For i = 2 To MeterNumber * 2 Step 2
                If .Cell(t, i).BackColor = vbColorLabelUnabled Then
                    SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Test " & t & " Meter " & i / 2 & " Selected", "False"
                Else
                    Debug.Print .Cell(t, i - 1).Text
                    SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Test " & t & " Meter " & i / 2 & " Selected", .Cell(t, i - 1).Text
                End If
            Next
        Next
    End With
    
    rc = CheckAverage
    
    If rc Then
    
salva:
        If Text1(2) <> "" Then SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Total Average", Text1(2)
        If Text1(3) <> "" Then SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Total Readings", Text1(3)
        If Text1(13) <> "" Then SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Total Tests", Text1(13)
        If Text1(4) <> "" Then SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Selected Average", Text1(4)
        If Text1(5) <> "" Then SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Selected Readings", Text1(5)
        If Text1(6) <> "" Then SaveSettingData SettingName, "Graph QC", "Standard " & SelectedSTDNumber & " Selected Tests", Text1(6)
        'Debug.Print Text1(6)
    Else
    
        SetMeanValue (SelectedSTDNumber)
        GoTo salva:
    
    End If
    CloseSettingDataFile
    
End Function

Private Function CheckAverage() As Boolean
Dim rc As Boolean
Dim i As Integer

    rc = True
    If Text1(2) = "" Then rc = False
    If Text1(3) = "" Then rc = False
    If Text1(13) = "" Then rc = False
    If Text1(4) = "" Then rc = False
    If Text1(5) = "" Then rc = False
    If Text1(6) = "" Then rc = False
    CheckAverage = rc
    
    
End Function

Private Function SaveResults() As Boolean

Dim rc As Boolean
    
Dim i As Integer
Dim t As Integer

On Error GoTo ERR_SAVE
rc = True
    
    
    If SelectedSTDNumber = 0 Then Exit Function
    
    Call SaveSelectedTest
        ' salva tabella Test SelectedSTDNumber
    Call FillGrd3

    CloseSettingDataFile
    
    PutLotInDatabase


ERR_END:
    On Error GoTo 0
    SaveResults = rc
    If rc Then
        PopupMessage 2, "Mean Value for Standard n " & STD(numSelectedStandard, 0) & " (" & STD(numSelectedStandard, 1) & "ppm ) " & vbCrLf & "Saved...", , , "Mean Results"
    End If
    Exit Function
ERR_SAVE:
    rc = False
    MsgBox err.Description
    Resume Next
    
End Function

Private Function PutLotInDatabase() As Boolean

Dim rc As Boolean

On Error GoTo ERR_SAVE:
rc = True
With dbTabReport
    .filter = ""
    .filter = "Lot='" & Trim(Text1(0)) & "' and Code='" & Trim(Text1(1)) & "' and NomeFile='" & SettingName & "'"
    
    If .EOF Then
        ' come è possibile????
        Exit Function
    Else
    End If
        !Evaluation = IIf(Grd3.Rows > 1, True, False)
        .Update
End With

ERR_END:
    On Error GoTo 0
    PutLotInDatabase = rc
    Exit Function
ERR_SAVE:
    MsgBox err.Description
    rc = False
    Resume ERR_END
End Function

Private Function SaveResultsTable()
    
Dim i As Integer
Dim t As Integer

    ' salva tabella Results

    With Grd3
        If .Rows > 1 Then
            SaveSettingData SettingName, "Evaluation QC", "Results Grid Rows", .Rows
            SaveSettingData SettingName, "Evaluation QC", "Results Grid Cols", .Cols
            For i = 0 To .Rows - 1
                For t = 1 To .Cols - 1
                    SaveSettingData SettingName, "Evaluation QC", "Results Grid Standard (" & i & ")  Column " & t, .Cell(i, t).Text
                    'If .Cell(i, t).ForeColor <> 0 Then
                        SaveSettingData SettingName, "Evaluation QC", "Results Grid Standard (" & i & ") Forecolor " & t, .Cell(i, t).ForeColor
                    'End If
                Next
            Next
        End If
    End With
    
    
    CloseSettingDataFile
End Function

Private Function FillGrd3()
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim STDValue As String
Dim STDNum As Integer

t = numSelectedStandard


STDNum = CInt(STD(t, 0))
STDValue = STD(t, 1)

If STDNum = 0 Then Exit Function
With Grd3
    .ReadOnly = True
    .AutoRedraw = False
    If .Rows <= 1 Then GoTo Aggiungi:
    For i = 1 To .Rows - 1
    
        If STDNum = Trim(.Cell(i, 5).Text) And STDValue = Trim(.Cell(i, 6).Text) Then
            'c'è già
            .Cell(i, 1).Text = STDValue & " ppm"
            .Cell(i, 2).Text = STD(t, 2) & " " & AndOr & " " & STD(t, 3)
            .Cell(i, 3).Text = Text1(4)
            .Cell(i, 4).Text = Text1(2)
            .Cell(i, 3).ForeColor = Text1(4).ForeColor
            .Cell(i, 4).ForeColor = Text1(2).ForeColor
            .Cell(i, 7).Text = strPassed
            GoTo fine:
        End If
    Next
Aggiungi:
            ' lo aggiungo
            .AddItem "", False
            .Cell(.Rows - 1, 1).Text = STDValue & " ppm"
            .Cell(.Rows - 1, 2).Text = STD(t, 2) & " " & AndOr & " " & STD(t, 3)
            .Cell(.Rows - 1, 3).Text = Text1(4)
            .Cell(.Rows - 1, 4).Text = Text1(2)
            .Cell(i, 3).ForeColor = Text1(4).ForeColor
            .Cell(i, 4).ForeColor = Text1(2).ForeColor
            .Cell(.Rows - 1, 5).Text = STDNum
            .Cell(.Rows - 1, 6).Text = STDValue
            .Cell(.Rows - 1, 7).Text = strPassed
fine:
    .Column(6).Sort cellAscending
    .AutoRedraw = True
    .Refresh
End With


End Function


Private Function GetphValue()
Dim i As Integer

    'For i = 1 To pHNumber
    '    pHMin(i) = GetSettingData(SettingName, "Information QC", "pHMin" & i, 0)
    '    pHMax(i) = GetSettingData(SettingName, "Information QC", "pHMax" & i, 0)
    'Next

Dim Num As Integer
Dim sString As String

    Num = 0
    
    With dbTabCode
        For i = 1 To 3
            Num = Num + 1
            sString = GetSettingData(SettingName, "Information QC", "pHValue" & Num, "0")
            
            
            
            If (sString <> "" Or sString <> "/") And IsNumeric(sString) And sString <> "0" Then
                'ph(Num, 0) = sString
    
                pHMin(Num) = GetSettingData(SettingName, "Information QC", "pHMin" & Num, "0")
                pHMax(Num) = GetSettingData(SettingName, "Information QC", "pHMax" & Num, "0")
            Else
                pHMin(Num) = ""
                pHMax(Num) = ""
            End If
                   
        Next
    End With

End Function

Private Function CloseLot()

Dim rc As Boolean

' chiudiamo il Lotto!!!!

rc = True

    ' salvo gli ultimi cambiamenti.....
    
    
    
    
    
    If bHoChiusoilLotto Then
        PopupMessage 2, "Lot already Closed..."
        Exit Function
    End If
    
    
    If Text1(11) = "" Then
        Text1(11) = FormatDataLAT(Now)
    End If

    SaveSelectedTest
    
    SaveResultsTable
    
    
    SaveSettingData SettingName, "Close QC", "Date", FormatDateTime(Now, vbShortDate)
    SaveSettingData SettingName, "Close QC", "Validation Date", Text1(11)
    SaveSettingData SettingName, "Close QC", "Operator", Text1(12)
    
    CloseSettingDataFile
    


On Error GoTo ERR_SAVE:
rc = True
With dbTabReport
    .filter = ""
    .filter = "Lot='" & Trim(Text1(0)) & "' and Code='" & Trim(Text1(1)) & "' and NomeFile='" & SettingName & "'"
    
    If .EOF Then
        ' come è possibile????
        Exit Function
    Else
    End If
        !Finished = IIf(Grd3.Rows > 1, True, False)
        .Update
End With

ERR_END:
    On Error GoTo 0
    
    If rc Then
    
        If USER_PATH = USER_DATA_PATH Then
            ' non c'è bisogno di spostarlo/cancellarlo----
        Else
            FileCopy USER_PATH & SettingName, USER_DATA_PATH & SettingName
            Kill USER_PATH & SettingName
        End If
        PopupMessage 2, "Lot closed..." & vbCrLf & "Lot : " & Trim(Text1(0)) & "  Code : " & Trim(Text1(1))
      
        USER_PATH = USER_DATA_PATH
    End If
    bHoChiusoilLotto = rc
    CloseLot = rc
    Exit Function
ERR_SAVE:
    MsgBox err.Description
    rc = False
    Resume ERR_END

End Function


Private Sub OpenReadingQC()

F_READING.Left = Me.Left
F_READING.Top = Me.Top

If F_READING.DoShow(IndexFormProcedura, Text1(0), Text1(1), ReadinglngID, , SettingName) Then
    
  GetFormSettingName

End If


End Sub

Private Function CheckRequestedFiles() As Boolean
Dim rc As Boolean
Dim sString As String
Dim FiledString As String
On Error GoTo ERR_CHECK



    ' campi obbligatori

    rc = True

    FiledString = "Lot Expiration"
    sString = GetSettingData(SettingName, "Information QC", "Text13", 1)
    GoSub CheckValue
    
    
    FiledString = "Preparation Week"
    sString = GetSettingData(SettingName, "Information QC", "Text121", 1)
    GoSub CheckValue
    
     
 
    FiledString = "Prep. Operator"
    sString = GetSettingData(SettingName, "Information QC", "Text122", 1)
    GoSub CheckValue
    
    
    FiledString = "First Day Prod."
    sString = GetSettingData(SettingName, "Information QC", "Text123", 1)
    GoSub CheckValue
    

    FiledString = "Last Day Prod."
    sString = GetSettingData(SettingName, "Information QC", "Text124", 1)
    GoSub CheckValue
    
    
    FiledString = "Machine"
    sString = GetSettingData(SettingName, "Information QC", "Text125", 1)
    GoSub CheckValue
    
             
    




ERR_END:
    On Error GoTo 0
    PicInformation.Left = Picture2.Left
    If Not (rc) Then
        'MessageInfoTime = 2000
        'PopupMessage 2, "This Lot cannot be closed : please fill all Requested fields" & vbCrLf & FiledString & " is missing." & FiledString & "Goto Lot Information QC"
        lbClose.Visible = True
        lbClose.ForeColor = &H40C0&     ' vbred
        PicInformation.Visible = True
        lbClose = "[Close Lot]  Please, fill all the required fields : Lot Expiration , Preparation Week, Prep. Operator, First day Prod. Last Day Prod., Machine"
    Else
        lbClose.Visible = True
        lbClose.ForeColor = vbColorGreen
        lbClose = "[Close Lot] This Lot has all the Information required..."
        If Grd3.Rows > 1 Then
        Else
            lbClose.ForeColor = &H40C0& ' vbred
            lbClose = "[Close Lot]  This Lot needs at least 1 Mean Value"
            
        End If
        
        PicInformation.Visible = False
    End If
    CheckRequestedFiles = rc
    Exit Function
ERR_CHECK:
    MsgBox err.Description
    rc = False
    Resume ERR_END

CheckValue:
    If sString = "" Then
    rc = False
    GoTo ERR_END:
    Else
        Return
    End If

End Function




Private Sub PicInformation_Click()
MyLot = Text1(0)
MyCode = Text1(1)
bAnotherFormCalled = True

F_INFORMATION.Left = Me.Left
F_INFORMATION.Top = Me.Top

If F_INFORMATION.DoShow(IndexFormProcedura, Text1(0), Text1(1), ReadinglngID, , SettingName) Then
    Form_Initialize

    lbOperator = MyOperatore.Name
  
End If
CheckRequestedFiles
bAnotherFormCalled = False

End Sub


