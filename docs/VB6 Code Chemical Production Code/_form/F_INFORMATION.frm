VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form F_INFORMATION 
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
   Picture         =   "F_INFORMATION.frx":0000
   ScaleHeight     =   12000
   ScaleMode       =   0  'User
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      FillColor       =   &H00004000&
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   2880
      ScaleHeight     =   855
      ScaleWidth      =   3975
      TabIndex        =   17
      Top             =   10920
      Width           =   3975
      Begin VB.Image ImageTAV 
         Height          =   480
         Index           =   5
         Left            =   1720
         MouseIcon       =   "F_INFORMATION.frx":1DED9
         MousePointer    =   99  'Custom
         Picture         =   "F_INFORMATION.frx":1E1E3
         Top             =   160
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   0
      Left            =   0
      MouseIcon       =   "F_INFORMATION.frx":215C5
      MousePointer    =   99  'Custom
      ScaleHeight     =   1815
      ScaleWidth      =   2775
      TabIndex        =   19
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
         MouseIcon       =   "F_INFORMATION.frx":218CF
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   0
         Left            =   1200
         Picture         =   "F_INFORMATION.frx":21BD9
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
      TabIndex        =   18
      Top             =   1080
      Width           =   16455
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   4
         Left            =   13560
         TabIndex        =   28
         Top             =   825
         Width           =   2535
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
         Left            =   11160
         TabIndex        =   26
         Top             =   825
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   2
         Left            =   4920
         TabIndex        =   24
         Top             =   825
         Width           =   6015
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
         Left            =   2400
         TabIndex        =   22
         Top             =   840
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
         Left            =   360
         TabIndex        =   0
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Line"
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
         Index           =   4
         Left            =   13560
         TabIndex        =   29
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Expiration"
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
         Index           =   3
         Left            =   11160
         TabIndex        =   27
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Description"
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
         Index           =   2
         Left            =   4920
         TabIndex        =   25
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
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
         Height          =   345
         Index           =   1
         Left            =   2400
         TabIndex        =   23
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
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
         Height          =   345
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   480
         Width           =   1815
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
      TabIndex        =   9
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   3
         Left            =   8640
         MouseIcon       =   "F_INFORMATION.frx":24FBB
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
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
            Index           =   3
            Left            =   0
            MouseIcon       =   "F_INFORMATION.frx":252C5
            MousePointer    =   99  'Custom
            TabIndex        =   68
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   735
            MouseIcon       =   "F_INFORMATION.frx":255CF
            MousePointer    =   99  'Custom
            Picture         =   "F_INFORMATION.frx":258D9
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MouseIcon       =   "F_INFORMATION.frx":28CBB
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   14
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "F_INFORMATION.frx":28FC5
            MousePointer    =   99  'Custom
            Picture         =   "F_INFORMATION.frx":292CF
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Recipe / Reagent"
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
            MouseIcon       =   "F_INFORMATION.frx":2C6B1
            MousePointer    =   99  'Custom
            TabIndex        =   15
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
         MouseIcon       =   "F_INFORMATION.frx":2C9BB
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   12
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Prod. / QC"
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
            MouseIcon       =   "F_INFORMATION.frx":2CCC5
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   720
            Width           =   1875
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   735
            MousePointer    =   99  'Custom
            Picture         =   "F_INFORMATION.frx":2CFCF
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   3840
         MouseIcon       =   "F_INFORMATION.frx":303B1
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   10
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Instruments"
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
            MouseIcon       =   "F_INFORMATION.frx":306BB
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   735
            MouseIcon       =   "F_INFORMATION.frx":309C5
            MousePointer    =   99  'Custom
            Picture         =   "F_INFORMATION.frx":30CCF
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Label LaInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This Lot is Closed : The information on this page cannot be Changed"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   13200
         MouseIcon       =   "F_INFORMATION.frx":340B1
         MousePointer    =   99  'Custom
         TabIndex        =   165
         Top             =   795
         Visible         =   0   'False
         Width           =   5490
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Information QC"
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
         Left            =   12750
         TabIndex        =   16
         Top             =   360
         Width           =   5940
      End
   End
   Begin VB.PictureBox PicMain 
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   1
      Left            =   0
      Picture         =   "F_INFORMATION.frx":343BB
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   65
      Top             =   2880
      Width           =   19215
      Begin FlexCell.Grid GrdDepart 
         Height          =   3840
         Left            =   8160
         TabIndex        =   164
         Top             =   3720
         Visible         =   0   'False
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   6773
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
         Index           =   32
         Left            =   10800
         TabIndex        =   81
         Top             =   3105
         Width           =   2535
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
         Index           =   31
         Left            =   8160
         TabIndex        =   80
         Top             =   3105
         Width           =   2535
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
         Index           =   30
         Left            =   5520
         TabIndex        =   79
         Top             =   3105
         Width           =   2535
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
         Index           =   25
         Left            =   13440
         TabIndex        =   77
         Top             =   1665
         Width           =   2535
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
         Index           =   24
         Left            =   10800
         TabIndex        =   75
         Top             =   1665
         Width           =   2535
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
         Index           =   23
         Left            =   8160
         TabIndex        =   73
         Top             =   1665
         Width           =   2535
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
         Index           =   22
         Left            =   5520
         TabIndex        =   71
         Top             =   1665
         Width           =   2535
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
         Index           =   21
         Left            =   2880
         TabIndex        =   69
         Top             =   1665
         Width           =   2535
      End
      Begin FlexCell.Grid GrdQCType 
         Height          =   3960
         Left            =   5520
         TabIndex        =   175
         Top             =   3720
         Visible         =   0   'False
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   6985
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Registration Book"
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
         Index           =   32
         Left            =   10800
         TabIndex        =   99
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "QC Department"
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
         Index           =   31
         Left            =   8160
         TabIndex        =   98
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "QC TYPE"
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
         Index           =   30
         Left            =   5520
         TabIndex        =   97
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Machine"
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
         Index           =   25
         Left            =   13440
         TabIndex        =   78
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Last day Prod."
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
         Index           =   24
         Left            =   10800
         TabIndex        =   76
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "First day Prod."
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
         Index           =   23
         Left            =   8160
         TabIndex        =   74
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Prep. Operator"
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
         Index           =   22
         Left            =   5520
         TabIndex        =   72
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Preparation Week"
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
         Index           =   21
         Left            =   2880
         TabIndex        =   70
         Top             =   1320
         Width           =   2535
      End
   End
   Begin VB.PictureBox PicMain 
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   2
      Left            =   0
      Picture         =   "F_INFORMATION.frx":52294
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   66
      Top             =   2880
      Width           =   19215
      Begin FlexCell.Grid GrdPH 
         Height          =   2040
         Left            =   2400
         TabIndex        =   154
         Top             =   4680
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   3598
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
      Begin FlexCell.Grid GrdTURBID 
         Height          =   2040
         Left            =   7200
         TabIndex        =   155
         Top             =   4680
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   3598
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
         Index           =   57
         Left            =   14400
         TabIndex        =   95
         Top             =   4080
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
         Index           =   56
         Left            =   12000
         TabIndex        =   94
         Top             =   4080
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
         Index           =   55
         Left            =   7200
         TabIndex        =   96
         Top             =   5400
         Visible         =   0   'False
         Width           =   4695
      End
      Begin FlexCell.Grid GrdMeter 
         Height          =   2040
         Left            =   960
         TabIndex        =   153
         Top             =   2520
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   3598
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
         Index           =   40
         Left            =   16080
         TabIndex        =   89
         Top             =   1920
         Width           =   2055
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
         Index           =   39
         Left            =   13920
         TabIndex        =   88
         Top             =   1920
         Width           =   2055
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
         Index           =   38
         Left            =   11760
         TabIndex        =   87
         Top             =   1920
         Width           =   2055
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
         Index           =   37
         Left            =   9600
         TabIndex        =   86
         Top             =   1920
         Width           =   2055
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
         Index           =   36
         Left            =   7440
         TabIndex        =   85
         Top             =   1920
         Width           =   2055
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
         Index           =   35
         Left            =   5280
         TabIndex        =   84
         Top             =   1920
         Width           =   2055
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
         Index           =   44
         Left            =   9600
         TabIndex        =   93
         Top             =   4080
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
         Index           =   43
         Left            =   7200
         TabIndex        =   92
         Top             =   4080
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
         Index           =   42
         Left            =   4800
         TabIndex        =   91
         Top             =   4080
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
         Index           =   41
         Left            =   2400
         TabIndex        =   90
         Top             =   4080
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
         Index           =   34
         Left            =   3120
         TabIndex        =   83
         Top             =   1920
         Width           =   2055
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
         Index           =   33
         Left            =   960
         TabIndex        =   82
         Top             =   1920
         Width           =   2055
      End
      Begin FlexCell.Grid GrdSPECTR 
         Height          =   2040
         Left            =   12000
         TabIndex        =   163
         Top             =   4680
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   3598
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
      Begin VB.Image ReloadMeter 
         Height          =   480
         Left            =   9360
         MouseIcon       =   "F_INFORMATION.frx":7016D
         MousePointer    =   99  'Custom
         Picture         =   "F_INFORMATION.frx":70477
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lbMeter 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "SPECTR. Meter"
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
         Index           =   56
         Left            =   12000
         TabIndex        =   162
         Top             =   3300
         Width           =   4695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Description"
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
         Index           =   57
         Left            =   14400
         TabIndex        =   161
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Code"
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
         Index           =   56
         Left            =   12000
         TabIndex        =   160
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "QC Operator"
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
         Index           =   55
         Left            =   7200
         TabIndex        =   157
         Top             =   5040
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Family"
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
         Index           =   39
         Left            =   13920
         TabIndex        =   130
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Code"
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
         Index           =   40
         Left            =   16080
         TabIndex        =   129
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lbMeter 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "METER 4"
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
         Index           =   39
         Left            =   13920
         TabIndex        =   128
         Top             =   1140
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Family"
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
         Index           =   37
         Left            =   9600
         TabIndex        =   127
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Code"
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
         Index           =   38
         Left            =   11760
         TabIndex        =   126
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lbMeter 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "METER 3"
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
         Index           =   37
         Left            =   9600
         TabIndex        =   125
         Top             =   1140
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Family"
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
         Index           =   35
         Left            =   5280
         TabIndex        =   124
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Code"
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
         Index           =   36
         Left            =   7440
         TabIndex        =   123
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lbMeter 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "METER 2"
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
         Index           =   35
         Left            =   5280
         TabIndex        =   122
         Top             =   1140
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Code"
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
         Index           =   43
         Left            =   7200
         TabIndex        =   121
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Description"
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
         Index           =   44
         Left            =   9600
         TabIndex        =   120
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label lbMeter 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "pH Meter"
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
         Index           =   41
         Left            =   2400
         TabIndex        =   119
         Top             =   3300
         Width           =   4695
      End
      Begin VB.Label lbMeter 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "TURBID. Meter"
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
         Index           =   43
         Left            =   7200
         TabIndex        =   111
         Top             =   3300
         Width           =   4695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Description"
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
         Index           =   42
         Left            =   4800
         TabIndex        =   110
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Code"
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
         Index           =   41
         Left            =   2400
         TabIndex        =   109
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label lbMeter 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "METER 1"
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
         Index           =   33
         Left            =   960
         TabIndex        =   108
         Top             =   1140
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Code"
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
         Index           =   34
         Left            =   3120
         TabIndex        =   107
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Family"
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
         Index           =   33
         Left            =   960
         TabIndex        =   106
         Top             =   1560
         Width           =   2055
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   0
      Left            =   0
      Picture         =   "F_INFORMATION.frx":73859
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Frame Frame2 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2175
         Left            =   17760
         TabIndex        =   112
         Top             =   5640
         Width           =   15495
         Begin VB.Image Image1 
            Height          =   480
            Left            =   4200
            Picture         =   "F_INFORMATION.frx":91732
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1 - Fill Lot / Reagent Information "
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   5400
            TabIndex        =   115
            Top             =   360
            Width           =   3405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2 - Enter Preparation "
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   5400
            TabIndex        =   114
            Top             =   720
            Width           =   2220
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3 - Select Instrument : Meter , pH ..."
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   5400
            TabIndex        =   113
            Top             =   1080
            Width           =   3645
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H000080DF&
            BorderColor     =   &H00006000&
            Height          =   2175
            Index           =   2
            Left            =   0
            Top             =   0
            Width           =   15495
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
         Index           =   26
         Left            =   1680
         TabIndex        =   100
         Top             =   5985
         Width           =   2535
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
         Index           =   27
         Left            =   4320
         TabIndex        =   101
         Top             =   5985
         Width           =   2535
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
         Index           =   28
         Left            =   6960
         TabIndex        =   102
         Top             =   5985
         Width           =   2535
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
         Index           =   29
         Left            =   9600
         TabIndex        =   103
         Top             =   5985
         Width           =   2535
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
         Index           =   58
         Left            =   12240
         TabIndex        =   104
         Top             =   5985
         Width           =   2535
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
         Index           =   59
         Left            =   14880
         TabIndex        =   105
         Top             =   5985
         Width           =   2535
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00A0A0A0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2175
         Left            =   1800
         TabIndex        =   118
         Top             =   3000
         Width           =   15495
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00606060&
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
            Index           =   54
            Left            =   11880
            TabIndex        =   140
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00606060&
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
            Index           =   53
            Left            =   9120
            TabIndex        =   139
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00606060&
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
            Index           =   52
            Left            =   6360
            TabIndex        =   138
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00606060&
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
            Index           =   51
            Left            =   3600
            TabIndex        =   137
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00606060&
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
            Index           =   50
            Left            =   720
            TabIndex        =   136
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00606060&
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
            Index           =   49
            Left            =   11880
            TabIndex        =   135
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00606060&
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
            Index           =   48
            Left            =   9120
            TabIndex        =   134
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00606060&
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
            Index           =   47
            Left            =   6360
            TabIndex        =   133
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00606060&
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
            Index           =   46
            Left            =   3600
            TabIndex        =   132
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00606060&
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
            Index           =   45
            Left            =   720
            TabIndex        =   131
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   15000
            TabIndex        =   151
            Top             =   1680
            Width           =   180
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00964901&
            Caption         =   "Reagent A Lot"
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
            Index           =   45
            Left            =   720
            TabIndex        =   150
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00964901&
            Caption         =   "Reagent B Lot"
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
            Index           =   46
            Left            =   3600
            TabIndex        =   149
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00964901&
            Caption         =   "Reagent C Lot"
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
            Index           =   47
            Left            =   6360
            TabIndex        =   148
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00964901&
            Caption         =   "Reagent D Lot"
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
            Index           =   48
            Left            =   9120
            TabIndex        =   147
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00964901&
            Caption         =   "Reagent E Lot"
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
            Index           =   49
            Left            =   11880
            TabIndex        =   146
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Expiration A"
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
            Index           =   50
            Left            =   720
            TabIndex        =   145
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Expiration B"
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
            Index           =   51
            Left            =   3600
            TabIndex        =   144
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Expiration C"
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
            Index           =   52
            Left            =   6360
            TabIndex        =   143
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Expiration D"
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
            Index           =   53
            Left            =   9120
            TabIndex        =   142
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Expiration E"
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
            Index           =   54
            Left            =   11880
            TabIndex        =   141
            Top             =   1200
            Width           =   2535
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
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
         Index           =   20
         Left            =   13680
         TabIndex        =   61
         Top             =   4545
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
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
         Index           =   19
         Left            =   10920
         TabIndex        =   59
         Top             =   4560
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
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
         Index           =   18
         Left            =   8160
         TabIndex        =   57
         Top             =   4560
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
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
         Index           =   17
         Left            =   5400
         TabIndex        =   54
         Top             =   4545
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
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
         Index           =   16
         Left            =   2520
         TabIndex        =   52
         Top             =   4545
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
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
         Index           =   15
         Left            =   13680
         TabIndex        =   50
         Top             =   3600
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
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
         Index           =   14
         Left            =   10920
         TabIndex        =   48
         Top             =   3585
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
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
         Left            =   8160
         TabIndex        =   46
         Top             =   3585
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
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
         Left            =   5400
         TabIndex        =   44
         Top             =   3600
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
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
         Left            =   2520
         TabIndex        =   42
         Top             =   3600
         Width           =   2535
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
         Index           =   10
         Left            =   14520
         TabIndex        =   40
         Top             =   1305
         Width           =   2535
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
         Left            =   11760
         TabIndex        =   38
         Top             =   1320
         Width           =   2535
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
         Left            =   8760
         TabIndex        =   36
         Top             =   945
         Width           =   1695
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
         Left            =   6840
         TabIndex        =   34
         Top             =   945
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
         Index           =   6
         Left            =   4920
         TabIndex        =   32
         Top             =   945
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
         Index           =   5
         Left            =   2040
         TabIndex        =   30
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Old Lot A"
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
         Index           =   26
         Left            =   1680
         TabIndex        =   174
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Expiration A"
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
         Index           =   27
         Left            =   4320
         TabIndex        =   173
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Old Lot B"
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
         Index           =   28
         Left            =   6960
         TabIndex        =   172
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Expiration B"
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
         Index           =   29
         Left            =   9600
         TabIndex        =   171
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Old Lot C"
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
         Index           =   58
         Left            =   12240
         TabIndex        =   170
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Expiration C"
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
         Index           =   59
         Left            =   14880
         TabIndex        =   169
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   4
         Left            =   120
         MouseIcon       =   "F_INFORMATION.frx":94B14
         MousePointer    =   99  'Custom
         TabIndex        =   156
         Top             =   6600
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   16800
         TabIndex        =   152
         Top             =   4680
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lbReagent 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "REAGENT SET 2"
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
         Index           =   51
         Left            =   5280
         MouseIcon       =   "F_INFORMATION.frx":94E1E
         MousePointer    =   99  'Custom
         TabIndex        =   117
         Top             =   2040
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label lbReagent 
         Alignment       =   2  'Center
         BackColor       =   &H000080DF&
         Caption         =   "REAGENT SET 1"
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
         Index           =   50
         Left            =   1680
         MouseIcon       =   "F_INFORMATION.frx":95128
         MousePointer    =   99  'Custom
         TabIndex        =   116
         Top             =   2040
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   6
         Left            =   480
         Picture         =   "F_INFORMATION.frx":95432
         Top             =   6840
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Get Last Lot Information : H156"
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
         MouseIcon       =   "F_INFORMATION.frx":98814
         MousePointer    =   99  'Custom
         TabIndex        =   64
         Top             =   6960
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lb 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         Caption         =   "Reagent "
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
         Height          =   405
         Index           =   22
         Left            =   1680
         TabIndex        =   63
         Top             =   2520
         Visible         =   0   'False
         Width           =   15735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Expiration E"
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
         Index           =   20
         Left            =   13680
         TabIndex        =   62
         Top             =   4200
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Expiration D"
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
         Index           =   19
         Left            =   10920
         TabIndex        =   60
         Top             =   4200
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Expiration C"
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
         Index           =   18
         Left            =   8160
         TabIndex        =   58
         Top             =   4200
         Width           =   2535
      End
      Begin VB.Label lb 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         Caption         =   "Reagent Range"
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
         Height          =   405
         Index           =   18
         Left            =   11400
         TabIndex        =   55
         Top             =   360
         Width           =   6015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Expiration B"
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
         Index           =   17
         Left            =   5400
         TabIndex        =   56
         Top             =   4200
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Expiration A"
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
         Index           =   16
         Left            =   2520
         TabIndex        =   53
         Top             =   4200
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Reagent E Lot"
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
         Index           =   15
         Left            =   13680
         TabIndex        =   51
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Reagent D Lot"
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
         Index           =   14
         Left            =   10920
         TabIndex        =   49
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Reagent C Lot"
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
         Index           =   13
         Left            =   8160
         TabIndex        =   47
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Reagent B Lot"
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
         Left            =   5400
         TabIndex        =   45
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Reagent A Lot"
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
         Left            =   2520
         TabIndex        =   43
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Max"
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
         Index           =   10
         Left            =   14520
         TabIndex        =   41
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Min"
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
         Index           =   9
         Left            =   11760
         TabIndex        =   39
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Max (mg)"
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
         Index           =   8
         Left            =   8760
         TabIndex        =   37
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Min (mg)"
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
         Index           =   7
         Left            =   6840
         TabIndex        =   35
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Ref. Weight"
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
         Left            =   4920
         TabIndex        =   33
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Recipe Number"
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
         Left            =   2040
         TabIndex        =   31
         Top             =   600
         Width           =   2535
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Left            =   9360
         Picture         =   "F_INFORMATION.frx":98B1E
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
         TabIndex        =   7
         Top             =   9120
         Width           =   975
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   7080
         Width           =   2655
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00606060&
         Height          =   2655
         Index           =   0
         Left            =   1680
         Top             =   2640
         Visible         =   0   'False
         Width           =   15735
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00606060&
         Height          =   1575
         Index           =   1
         Left            =   11400
         Top             =   480
         Width           =   6015
      End
   End
   Begin ChemicalQC.ctlCalendar ctlCalendar1 
      Height          =   6960
      Left            =   9600
      TabIndex        =   158
      Top             =   3600
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
      Height          =   1455
      Index           =   3
      Left            =   0
      MouseIcon       =   "F_INFORMATION.frx":9BF00
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   10680
      Width           =   1815
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   0
      Left            =   17640
      MouseIcon       =   "F_INFORMATION.frx":9C20A
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   10800
      Width           =   1575
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   2
      Left            =   15240
      MouseIcon       =   "F_INFORMATION.frx":9C514
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   10680
      Width           =   2175
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Index           =   1
      Left            =   8280
      MouseIcon       =   "F_INFORMATION.frx":9C81E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   10560
      Width           =   2655
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
      MouseIcon       =   "F_INFORMATION.frx":9CB28
      MousePointer    =   99  'Custom
      TabIndex        =   168
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
      MouseIcon       =   "F_INFORMATION.frx":9CE32
      MousePointer    =   99  'Custom
      TabIndex        =   167
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
      MouseIcon       =   "F_INFORMATION.frx":9D13C
      MousePointer    =   99  'Custom
      TabIndex        =   166
      Top             =   11600
      Width           =   1200
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
      TabIndex        =   159
      Top             =   11600
      Width           =   2325
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   4
      Left            =   2040
      MouseIcon       =   "F_INFORMATION.frx":9D446
      MousePointer    =   99  'Custom
      Picture         =   "F_INFORMATION.frx":9D750
      Top             =   11040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   15960
      MouseIcon       =   "F_INFORMATION.frx":A0B32
      MousePointer    =   99  'Custom
      Picture         =   "F_INFORMATION.frx":A0E3C
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_INFORMATION.frx":A421E
      Height          =   480
      Index           =   1
      Left            =   9360
      MouseIcon       =   "F_INFORMATION.frx":A7600
      MousePointer    =   99  'Custom
      Picture         =   "F_INFORMATION.frx":A790A
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   360
      MouseIcon       =   "F_INFORMATION.frx":AACEC
      MousePointer    =   99  'Custom
      Picture         =   "F_INFORMATION.frx":AAFF6
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
      TabIndex        =   1
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
      DragIcon        =   "F_INFORMATION.frx":AE3D8
      Height          =   480
      Index           =   0
      Left            =   18240
      MouseIcon       =   "F_INFORMATION.frx":B17BA
      MousePointer    =   99  'Custom
      Picture         =   "F_INFORMATION.frx":B1AC4
      Top             =   11040
      Width           =   480
   End
End
Attribute VB_Name = "F_INFORMATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IndexFormProcedura As Integer
Private IndexMainProcedura As Integer
Private IndexText As Integer
Private IndexTextSelected As Integer
Private MyLot As String
Private MyCode As String
Private m_rc As Boolean
Private bFormSaved As Boolean
Private strLot As String
Private strLastLot As String
Private MyMeterFamily As String
Private MypHMeter As String
Private MyDepartment As String
Private MyTurbidMeter As String
Private MySpectrMeter As String
Private MyQCType As String
Private MyIndex As Integer
Private bAnotherFormCalled As Boolean
Private MeterNumber As Integer
Private bReloadMeter As Boolean


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
    
    Set DefaultMenu(4) = MyImage
    IndexMainProcedura = Index
    strLastLot = ""
    strLot = sLot
    m_rc = False
    bFormSaved = False
    
    Picture1.Visible = Not (bSearchClosedLot)
    LaInfo.Visible = bSearchClosedLot
    lbOperator = MyOperatore.Name
    
    FormPulisciTutto
    Text1(55) = MyOperatore.Name
    If sLot <> "" And sCode <> "" Then
        Call GetCodeInformation(sLot, sCode, lngID)
    Else
        PopupMessage 2, "Please select a valid Code/Lot..."
        Unload Me
    End If
    
    FillAllGrid
    
    SelectProcedura 0
    mOk

    
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



Private Sub ctlCalendar1_DateClicked(inputDate As Date)
Select Case IndexTextSelected
  Case 3, 16 To 20, 27, 29, 50 To 54, 59
    Text1(IndexTextSelected) = FormatDataExp(CStr(inputDate))
Case Else

    Text1(IndexTextSelected) = FormatDataLAT(CStr(inputDate))
End Select
ctlCalendar1.Visible = False
End Sub

Private Sub DefaultMenuLabel_Click(Index As Integer)
Dim MyIndex As Integer
Select Case Index
    Case 0
        ' vai avanti
        If IndexFormProcedura = 2 Then
            MyIndex = 0
        Else
            MyIndex = IndexFormProcedura + 1
        End If
        PicMenu_Click MyIndex
    Case 1
        'If bFormSaved Then
            USER_PATH = USER_TEMP_PATH
            Unload Me
       ' Else
            
       ' End If
    Case 2
        ' torna indietro
        If IndexFormProcedura = 0 Then
            MyIndex = 2
        Else
            MyIndex = IndexFormProcedura - 1
        End If
        PicMenu_Click MyIndex
    Case 3
        frmLogin.DoShow
        lbOperator = MyOperatore.Name
    Case 4
        ' get last lot information
        Call GetLastLotInformation
        
    Case 5
       ' Label7_Click
    Case 6
       ' Label6_Click
End Select

End Sub


Private Sub DisableImage_Click()
PopupMessage 2, "Warning : Administrator Only can Operate...", , True
End Sub

Private Sub Form_Initialize()
LaInfo = "This Lot is Closed : The information on this page cannot be Changed"
Call SetPicForm
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
FrameIntroVisible False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 3
    If i = IndexFormProcedura Then
    Else
        PicMenu(i).BackColor = vbColorDarkFont
    End If
Next
Picture1.BackColor = &H4000&
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set F_INFORMATION = Nothing
End Sub



Private Sub GrdDepart_DblClick()

If MyDepartment <> "" Then
    
    Text1(31) = MyDepartment
    
    MyDepartment = ""
    GrdDepart.Visible = False
    
    Text1(32).SetFocus
End If
GrdDepart.Cell(0, 0).SetFocus


End Sub

Private Sub GrdDepart_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    If F_MsgBox.DoShow("Delete Department : " & MyDepartment & " ?") Then
        With dbTabDepartment
            .filter = ""
            .filter = "Code='" & Trim(MyDepartment) & "'"
            If .EOF Then
            Else
                .Delete
                .Update
            End If
        End With
        
        FillAllGrid
    End If
End If

End Sub

Private Sub GrdDepart_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

If FirstRow > 0 Then

    MyDepartment = GrdDepart.Cell(FirstRow, 1).Text

Else

    GrdDepart.Cell(0, 0).SetFocus
    
End If

End Sub


Private Sub GrdMeter_DblClick()

If MyIndex <> 0 And MyMeterFamily <> "" Then
    
    Text1(MyIndex) = MyMeterFamily
    
    MyMeterFamily = ""
    MyIndex = 0
    GrdMeter.Visible = False
    
    Text1(MyIndex + 1).SetFocus
End If
GrdMeter.Cell(0, 0).SetFocus
End Sub

Private Sub GrdMeter_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

If FirstRow > 0 Then

    MyMeterFamily = GrdMeter.Cell(FirstRow, 1).Text

Else

    GrdMeter.Cell(0, 0).SetFocus
    
End If

End Sub


Private Sub grdph_DblClick()

If MypHMeter <> "" Then
    
    Text1(41) = MypHMeter
    
    MypHMeter = ""
    GrdPH.Visible = False
    
    Text1(42).SetFocus
End If
GrdPH.Cell(0, 0).SetFocus
End Sub



Private Sub GrdPH_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    If F_MsgBox.DoShow("Delete ph Meter : " & MypHMeter & " ?") Then
        With dbTabPHMeter
            .filter = ""
            .filter = "Code='" & Trim(MypHMeter) & "'"
            If .EOF Then
            Else
                .Delete
                .Update
            End If
        End With
        
        FillAllGrid
    End If
End If
End Sub

Private Sub grdph_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

If FirstRow > 0 Then

    MypHMeter = GrdPH.Cell(FirstRow, 1).Text

Else

    GrdPH.Cell(0, 0).SetFocus
    
End If

End Sub

Private Sub GrdQCType_DblClick()

If MyQCType <> "" Then
    
    Text1(30) = MyQCType
    
    MyQCType = ""
    GrdQCType.Visible = False
    
   ' Text1(44).SetFocus
End If
GrdQCType.Cell(0, 0).SetFocus
End Sub

Private Sub GrdQCType_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    If F_MsgBox.DoShow("Delete QC type : " & MyQCType & " ?") Then
        With dbTabQCType
            .filter = ""
            .filter = "Type='" & MyQCType & "'"
            If .EOF Then
            Else
                .Delete
                .Update
            End If
        End With
        
        FillAllGrid
    End If
End If
End Sub

Private Sub GrdQCType_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

If FirstRow > 0 Then

    MyQCType = Trim(GrdQCType.Cell(FirstRow, 1).Text)

Else

    GrdQCType.Cell(0, 0).SetFocus
    
End If

End Sub

Private Sub GrdSPECTR_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    If F_MsgBox.DoShow("Delete Spectrophotometer Meter : " & MySpectrMeter & " ?") Then
        With dbTabSpectMeter
            .filter = ""
            .filter = "Code='" & Trim(MySpectrMeter) & "'"
            If .EOF Then
            Else
                .Delete
                .Update
            End If
        End With
        
        FillAllGrid
    End If
End If
End Sub

Private Sub grdTurbid_DblClick()

If MyTurbidMeter <> "" Then
    
    Text1(43) = MyTurbidMeter
    
    MyTurbidMeter = ""
    GrdTURBID.Visible = False
    
    Text1(44).SetFocus
End If
GrdTURBID.Cell(0, 0).SetFocus
End Sub

Private Sub GrdSPECTR_DblClick()

If MySpectrMeter <> "" Then
    
    Text1(56) = MySpectrMeter
    
    MySpectrMeter = ""
    GrdSPECTR.Visible = False
    
    Text1(44).SetFocus
End If
GrdSPECTR.Cell(0, 0).SetFocus
End Sub

Private Sub GrdTURBID_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    If F_MsgBox.DoShow("Delete Turbid Meter : " & MyTurbidMeter & " ?") Then
        With dbTabTurbMeter
            .filter = ""
            .filter = "Code='" & Trim(MyTurbidMeter) & "'"
            If .EOF Then
            Else
                .Delete
                .Update
            End If
        End With
        
        FillAllGrid
    End If
End If
End Sub

Private Sub grdTurbid_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

If FirstRow > 0 Then

    MyTurbidMeter = GrdTURBID.Cell(FirstRow, 1).Text

Else

    GrdTURBID.Cell(0, 0).SetFocus
    
End If

End Sub
Private Sub GrdSPECTR_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

If FirstRow > 0 Then

    MySpectrMeter = GrdSPECTR.Cell(FirstRow, 1).Text

Else

    GrdSPECTR.Cell(0, 0).SetFocus
    
End If

End Sub



Private Sub Image3_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub ImageTAV_Click(Index As Integer)
Select Case Index
    Case 5
        Picture1_Click
End Select
End Sub

Private Sub lbReagent_Click(Index As Integer)
Dim rc As Boolean

    Select Case Index
        Case 50
            rc = True
        Case 51
            rc = False
        Case Else
            Exit Sub
    End Select
    
    lbReagent(50).BackColor = IIf(rc, &H80DF&, &H808080)
    lbReagent(51).BackColor = IIf(Not (rc), &H80DF&, &H808080)
    
    Call ReagentSet(rc)
    
End Sub

Private Sub Label2_Click(Index As Integer)
PicMenu_Click Index
End Sub




Private Sub PicMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Len(Text1(IndexText)) = 0 Then Text1(IndexText).BackColor = vbColorUnabled
End Sub

Private Sub PicMain_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
GrdMeter.Visible = False
GrdPH.Visible = False
GrdTURBID.Visible = False
GrdSPECTR.Visible = False
GrdDepart.Visible = False
End Sub

Private Sub PicMenu_Click(Index As Integer)
If IndexFormProcedura = Index Then
ElseIf Index = PicMenu.Count - 1 Then
    IndexMainProcedura = IndexMainProcedura + 1
'    F_INFORMATION.DoShow 1, MyLot, MyCode
    Unload Me
Else

    Call SelectProcedura(Index)

    
End If
End Sub


Private Function SelectProcedura(ByVal Index As Integer) As Boolean
Dim rc As Boolean
Dim mrc As Boolean
Dim i As Integer
If Index > 3 Then Exit Function
For i = 0 To 3
    If i = Index Then
        PicMenu(i).BackColor = vbColorForeFixed
    Else
        PicMenu(i).BackColor = vbColorDarkFont
    End If
Next
Set Image4(0) = Image3(Index)
Select Case Index
    Case 0
        Picture4(0).BackColor = &H8000&
        rc = True
    Case 1
        Picture4(0).BackColor = &H6000&
        rc = False
    Case 2
        GetTabInstrument
        
        Picture4(0).BackColor = &H5000&
        rc = False
        
        Dim t As Integer
    
        For t = 33 To 37 Step 2
            mrc = IIf(Len(Text1(t)) > 0, True, False)
           ' Text1(t + 1).Enabled = mrc
            Text1(t + 2).Enabled = mrc
            Text1(t + 3).Enabled = mrc
            
           ' Text1(t + 1).BackColor = IIf(mrc, Text1(t + 1).BackColor, vbColorDarkFont)
           ' Text1(t + 2).BackColor = IIf(mrc, Text1(t + 2).BackColor, vbColorDarkFont)
           ' Text1(t + 3).BackColor = IIf(mrc, Text1(t + 3).BackColor, vbColorDarkFont)
            
        Next
    
        
End Select
For i = 0 To 4
    Text1(i).Locked = True
Next

Text1(3).Locked = False
  
GrdMeter.Visible = False
GrdPH.Visible = False
GrdTURBID.Visible = False
GrdSPECTR.Visible = False
GrdDepart.Visible = False

Label2(4) = Label2(Index)
IndexFormProcedura = Index
PicMain(Index).Visible = True
PicMain(Index).Enabled = Not (bSearchClosedLot)
PicMain(Index).ZOrder
'blTable = Label2(IndexFormProcedura)
Cleanform (False)
End Function

Private Sub Cleanform(ByVal bValue As Boolean, Optional ByVal Index As Integer = 0)

End Sub

Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 3
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
ctlCalendar1.Left = Me.Width / 2 - ctlCalendar1.Width / 2
ctlCalendar1.Top = Me.Height / 2 - ctlCalendar1.Height / 2

Frame2.Top = 3000

Frame2.Left = 1800

Frame3.Top = 3000
Frame3.Left = 1800
For i = 0 To PicMain.Count - 1
    PicMain(i).Left = 0
    PicMain(i).Top = PicMenuBar(0).Height + Frame1.Height
    PicMain(i).Width = Me.Width
    PicMain(i).Height = Line1.Y1 - PicMain(i).Top
Next


For i = 0 To Text1.Count - 1
    Text1(i).BackColor = IIf(Len(Text1(i)) > 0, vbWhite, vbColorUnabled)
Next

' impost ail primo set di reagent

ReagentSet True

Call SetAllGrid



End Sub

Private Sub PicMenuBar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
PicMenu_MouseMove 20, Button, Shift, X, Y
End Sub

Private Sub Picture1_Click()

If CheckCodeMeter Then

    Call SaveProcedure
    
End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BackColor = &H8000&
End Sub



Private Sub ReloadMeter_Click()
bReloadMeter = True
PopupMessage 2, "Reload Meter Information..."
GetAllMeterInformation
bReloadMeter = False
End Sub

Private Sub Text1_Change(Index As Integer)
Dim rc As Boolean

    
    rc = IIf(Len(Text1(Index)) > 0, True, False)
    Text1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbWhite, vbColorUnabled)
    Label1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextBlue, vbColorLabelUnabled)

    Select Case Index
        Case 2
        Case 3
             FrameIntroVisible Not (rc)
             
        Case 6
            'Text1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbColorGreen, vbColorUnabled)
            'Text1(Index).ForeColor = IIf(Len(Text1(Index)) > 0, vbWhite, vbBlack)
        Case 33, 35, 37, 39, 41, 43, 56
        
            lbMeter(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextDarkBlue, vbColorLabelUnabled)
        Case 9, 10
            lb(18).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextDarkBlue, vbColorLabelUnabled)
    
            
    End Select
    
    Dim t As Integer
    If bReloadMeter Then Exit Sub
    
    t = Index
    Dim mrc As Boolean
    Select Case t
        Case 33
            mrc = IIf(Len(Text1(t)) > 0, True, False)
            Text1(t + 1).Enabled = mrc
            Text1(t + 2).Enabled = mrc
            Text1(t + 3).Enabled = mrc

        Case 35, 37
            mrc = IIf(Len(Text1(t)) > 0, True, False)
            Text1(t + 1).Enabled = mrc
            Text1(t + 2).Enabled = mrc
            Text1(t + 3).Enabled = mrc
            
            mrc = IIf(Len(Text1(t - 1)) > 0, True, False)
            If Not (mrc) Then
                PopupMessage 2, "Please Enter Meter " & ((t - 1) - 32) / 2 & " Code"
               ' Text1(t) = ""
               If Text1(t - 1).Enabled Then Text1(t - 1).SetFocus
            End If
        Case 39
            mrc = IIf(Len(Text1(t - 1)) > 0, True, False)
            If Not (mrc) Then
                PopupMessage 2, "Please Enter Meter " & ((t - 1) - 32) / 2 & " Code"
               ' Text1(t) = ""
                If Text1(t - 1).Enabled Then Text1(t - 1).SetFocus
            End If
    End Select
    
End Sub


Private Function CheckCodeMeter() As Boolean

 Dim i      As Integer
 Dim rc     As Boolean
 Dim mrc    As Boolean
 
 
    rc = True
    For i = 0 To 3
        mrc = IIf(Len(Text1(i * 2 + 33)) > 0, True, False)
        If mrc Then
            
            mrc = IIf(Len(Text1(i * 2 + 34)) > 0, True, False)
        
            If mrc = False Then
                PopupMessage 2, "Please Enter Meter " & i + 1 & " Code"
                Text1(i * 2 + 34).SetFocus
                
                rc = False
                Exit For
            End If
        
        End If
    
    Next
    
    CheckCodeMeter = rc

End Function

Private Sub FrameIntroVisible(ByVal rc As Boolean)

        Frame2.Visible = (rc)
        lb(22).Visible = Not (Frame2.Visible)
        Shape1(0).Visible = lb(22).Visible
        Label6.Visible = Shape1(0).Visible
        lbReagent(50).Visible = Shape1(0).Visible
        lbReagent(51).Visible = Shape1(0).Visible
End Sub

Private Sub Text1_DblClick(Index As Integer)

Select Case Index

    Case 35, 37, 39
    If Len(Text1(Index - 2)) > 0 Then
    
        Text1(Index) = Text1(Index - 2)
    End If
        
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
On Error Resume Next
If KeyAscii = 13 Then
    

    
    If Index < Text1.Count - 1 Then Text1(Index + 1).SetFocus
    ctlCalendar1.Visible = False
End If
End Sub


Private Sub CheckDepartment(ByVal sString As String)
sString = Trim(sString)
If (sString) = "" Then Exit Sub

    With dbTabDepartment
        .filter = ""
        .filter = "Code='" & sString & "'"
        If .EOF Then
            If F_MsgBox.DoShow("New Departmenr : Add in Database? ") Then
                .AddNew
                !code = (sString)
                .Update
                FillAllGrid
            End If
        End If
    End With
End Sub

Private Sub CheckQCType(ByVal sString As String)
sString = Trim(sString)
If (sString) = "" Then Exit Sub

    With dbTabQCType
        .filter = ""
        .filter = "Type='" & sString & "'"
        If .EOF Then
            If F_MsgBox.DoShow("New QC Type found : Add in Database? ") Then
                .AddNew
                !type = (sString)
                .Update
                FillAllGrid
            End If
        
        End If
    
    End With
End Sub

Private Sub CheckPH(ByVal sString As String)
sString = Trim(sString)
If (sString) = "" Then Exit Sub

    With dbTabPHMeter
        .filter = ""
        .filter = "Code='" & sString & "'"
        If .EOF Then
            If F_MsgBox.DoShow("New pH Meter found : Add in Database? ") Then
                .AddNew
                !code = (sString)
                .Update
                FillAllGrid
            End If
        
        End If
    
    End With
End Sub


Private Sub CheckTurbid(ByVal sString As String)
sString = Trim(sString)
If (sString) = "" Then Exit Sub

    With dbTabTurbMeter
        .filter = ""
        .filter = "Code='" & sString & "'"
        If .EOF Then
            If F_MsgBox.DoShow("New Turbid Meter found : Add in Database? ") Then
                .AddNew
                !code = (sString)
                .Update
                FillAllGrid
            End If
        
        End If
    
    End With
End Sub


Private Sub CheckSpectr(ByVal sString As String)
sString = Trim(sString)
If (sString) = "" Then Exit Sub

    With dbTabSpectMeter
        .filter = ""
        .filter = "Code='" & sString & "'"
        If .EOF Then
            If F_MsgBox.DoShow("New Spectrophotometer Meter found : Add in Database? ") Then
                .AddNew
                !code = (sString)
                .Update
                FillAllGrid
            End If
        
        End If
    
    End With
End Sub





Private Sub Text1_GotFocus(Index As Integer)

    Text1(Index).BackColor = vbWhite
    ctlCalendar1.ZOrder
    IndexTextSelected = Index
    ctlCalendar1.Visible = False
    GrdMeter.Visible = False
    GrdPH.Visible = False
    GrdTURBID.Visible = False
    GrdSPECTR.Visible = False
    GrdDepart.Visible = False
    GrdQCType.Visible = False
    MyIndex = 0
    If Text1(Index).Locked Or Not (Text1(Index).Enabled) Then Exit Sub

    Select Case Index
        Case 3
            ' expiration
            ctlCalendar1.Left = Text1(Index).Left - 500
            ctlCalendar1.Top = Frame1.Top + Text1(Index).Top + Text1(Index).Height + 120
            ctlCalendar1.Visible = True
            
        Case 31
            GrdDepart.Visible = True
        Case 33, 35, 37, 39
            MyIndex = Index
            GrdMeter.Top = Text1(Index).Top + Text1(Index).Height + 120
            GrdMeter.Left = Text1(Index).Left
            GrdMeter.Visible = True
        Case 41
            GrdPH.Visible = True
        Case 43
            GrdTURBID.Visible = True
        Case 56
            GrdSPECTR.Visible = True
        Case 3
            
            ctlCalendar1.Left = Text1(Index).Left
            ctlCalendar1.Top = Frame1.Top + Text1(Index).Top + Text1(Index).Height + 120
            ctlCalendar1.Visible = True
        Case 16 To 19
            ctlCalendar1.Left = Text1(Index).Left - ctlCalendar1.Width / 3 + 500
            ctlCalendar1.Top = PicMain(0).Top + Text1(Index).Top - ctlCalendar1.Height - 520
            ctlCalendar1.Visible = True
        Case 20
            ctlCalendar1.Left = Text1(Index).Left - ctlCalendar1.Width / 3
            ctlCalendar1.Top = PicMain(0).Top + Text1(Index).Top - ctlCalendar1.Height - 520
            ctlCalendar1.Visible = True
        Case 50 To 53
            ctlCalendar1.Left = Text1(Index - 34).Left - ctlCalendar1.Width / 3 + 500
            ctlCalendar1.Top = PicMain(0).Top + Text1(Index - 34).Top - ctlCalendar1.Height - 520
            ctlCalendar1.Visible = True
        Case 54
            ctlCalendar1.Left = Text1(Index - 34).Left - ctlCalendar1.Width / 3
            ctlCalendar1.Top = PicMain(0).Top + Text1(Index - 34).Top - ctlCalendar1.Height - 520
            ctlCalendar1.Visible = True
        
        Case 23, 24
            ' preparation
            ctlCalendar1.Left = Text1(Index).Left - ctlCalendar1.Width - 200
            ctlCalendar1.Top = PicMain(0).Top + Text1(Index).Top - ctlCalendar1.Height / 2
            ctlCalendar1.Visible = True
        Case 27
            ' preparation
            ctlCalendar1.Left = Text1(Index).Left + Text1(Index).Width + 200
            ctlCalendar1.Top = PicMain(0).Top + Text1(Index).Top - ctlCalendar1.Height '/ 2 - 600
            ctlCalendar1.Visible = True
        Case 29, 59
            ' preparation
            ctlCalendar1.Left = Text1(Index).Left - ctlCalendar1.Width - 200
            ctlCalendar1.Top = PicMain(0).Top + Text1(Index).Top - ctlCalendar1.Height '/ 2 - 600
            ctlCalendar1.Visible = True
        Case 30
            GrdQCType.Visible = True
        Case Else
            ctlCalendar1.Visible = False
            IndexTextSelected = -1
    End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim i As Integer
Text1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbWhite, vbColorUnabled)

    Select Case Index
    
        Case 30 ' QC TYPE
        
            CheckQCType (Text1(Index))
            FillAllGrid
        Case 31 ' Department
        
            CheckDepartment (Text1(Index))
            FillAllGrid
            
        Case 33, 35, 37, 39
        
            i = (Index - 31) / 2
            
            ' si sta cercando di modificare un meter....
            If Len(Text1(Index)) = 0 Then
                ' devo sontrollare se l'operatore lo ha cancellato ma aveva già fatto dei readings
                If i > MeterNumber Then
                
                
                Else
                    MessageInfoTime = 2000
                    PopupMessage 2, "A Meter cannot be Deleted or changed!" & vbCrLf & "You already saved Information QC..."
                    GetAllMeterInformation
                    
                    
                   ' Picture1.Visible = False
                   ' If F_MsgBox.DoShow("Please , Exit Procedure and come back to make other entries...", , , "Exit", "No") Then
                   '     Unload Me
                   ' End If
                   
                   
                End If
            Else
                ' check Meter Information
                If CheckReadingQC Then
                
                    
                Else
                    ' ci sono delle differenze. non è possibile
                    MessageInfoTime = 2000
                    PopupMessage 2, "A Meter cannot be Changed after Readings"
                    GetAllMeterInformation
                End If
            
            End If
        
          
        Case 41 ' PH
        
            CheckPH (Text1(Index))
            FillAllGrid
        
        Case 43 ' TURBID
        
            CheckTurbid (Text1(Index))
            FillAllGrid
        Case 56 ' SPECTR
        
            CheckSpectr (Text1(Index))
            FillAllGrid
            
    End Select
    
    
End Sub

Private Function SaveProcedure()
Dim rc As Boolean

    If Text1(55) <> lbOperator Then
        If Text1(55) = "" Then
            Text1(55) = lbOperator
        Else
            If F_MsgBox.DoShow("Information QC : Save data... " & vbCrLf & "Change QC Operator?" & vbCrLf & "OLD : " & Text1(55) & vbCrLf & " NEW : " & lbOperator) Then
                Text1(55) = lbOperator
            End If
        End If
    End If

    rc = SaveForm
    If PicMenu(3).Visible = False Then PicMenu(3).Visible = rc
    If rc Then
        MyLot = Text1(0)
        MyCode = Text1(1)
        PopupMessage 2, "Information Saved...", , , "Information QC"
    Else
    
    End If
    
    m_rc = rc
    bFormSaved = rc

End Function
Private Function SaveForm() As Boolean

Dim rc As Boolean
On Error GoTo ERR_SAVE:

    rc = True
    
    rc = PutLotInDatabase
    
    SalvaFormSettingName
    
ERR_END:
    On Error GoTo 0
    SaveForm = rc
    Exit Function
ERR_SAVE:
    rc = False
    GoTo ERR_END:
End Function

Private Function PutLotInDatabase() As Boolean

Dim rc As Boolean

On Error GoTo ERR_SAVE:
rc = True
With dbTabReport
    .filter = ""
    .filter = "Lot='" & Trim(Text1(0)) & "' and Code='" & Trim(Text1(1)) & "' and NomeFile='" & SettingName & "'"
    
    If .EOF Then
        .AddNew
        !StartDate = FormatDateTime(Now, vbShortDate)
    Else
        '  devo essere un Administrator????
    End If
        !Lot = Trim(Text1(0))
        !code = Trim(Text1(1))
        !Description = Trim(Text1(2))
        !Exp = Trim(Text1(3))
        !PREPWK = Trim(Text1(21))
        !Line = Trim(Text1(4))
        !Recipe = Trim(Text1(5))
        !RangeMin = Trim(Text1(9))
        !RangeMax = Trim(Text1(10))
        !Operator = MyOperatore.Name
        !Note = Trim(Text1(30))
        !Department = Trim(Text1(31))
        !Visible = True
        !NomeFile = SettingName
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


Private Sub ReagentSet(ByVal bValue As Boolean)
Dim rc As Boolean
rc = bValue
lb(22).Caption = IIf(rc, "REAGENT SET 1", "REAGENT SET 2")
Frame3.ZOrder
Frame3.Visible = Not (rc)


End Sub

Private Sub SetAllGrid()
    With GrdMeter
    
        .Rows = 1
        .ZOrder
        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        .Cols = 3
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Meter (Family)"
        .Column(1).Width = 190
        .Cell(0, 2).Text = "ID"
        .Column(2).Width = 0

        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .DefaultFont.Size = 12 ' * m_ControlGridFontSize
        .DefaultFont.Bold = False
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
    End With
    
    With GrdPH
        .ZOrder
        .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        .Cols = 3
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "pH Meter (Code)"
        .Column(1).Width = 190
        .Cell(0, 2).Text = "ID"
        .Column(2).Width = 0


        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .DefaultFont.Size = 12 '* m_ControlGridFontSize
        .DefaultFont.Bold = False
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
    End With
       
       
    With GrdTURBID
     .ZOrder
     .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        .Cols = 3
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Turbid Meter (Code)"
        .Column(1).Width = 190
        .Cell(0, 2).Text = "ID"
        .Column(2).Width = 0


        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .DefaultFont.Size = 12 ' * m_ControlGridFontSize
        .DefaultFont.Bold = False
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
    End With
    
    With GrdSPECTR
        .ZOrder
        .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        .Cols = 3
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Spectr. Meter (Code)"
        .Column(1).Width = 190
        .Cell(0, 2).Text = "ID"
        .Column(2).Width = 0


        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .DefaultFont.Size = 12 ' * m_ControlGridFontSize
        .DefaultFont.Bold = False
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
    End With
    
      With GrdDepart
        .ZOrder
        .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        .Cols = 3
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Department"
        .Column(1).Width = 150
        .Cell(0, 2).Text = "ID"
        .Column(2).Width = 0


        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .DefaultFont.Size = 12 ' * m_ControlGridFontSize
        .DefaultFont.Bold = False
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
    End With
    
    
    With GrdQCType
        .ZOrder
        .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 25
        .Cols = 3
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "QC Type"
        .Cell(0, 1).Alignment = cellCenterCenter
        .Column(1).Width = 150
        .Cell(0, 2).Text = "ID"
        .Column(2).Width = 0
        '

        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
        Next
        .DefaultFont.Size = 12 * m_ControlGridFontSize
        .DefaultFont.Bold = False
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
    End With
      
End Sub
Private Sub FillAllGrid()

Dim i As Integer




       
GrdQCType.Rows = 1


With dbTabQCType
    .filter = ""
    
    If .EOF Then
    
    Else
        GrdQCType.AutoRedraw = False
        .MoveFirst
        For i = 1 To .RecordCount
            
            With GrdQCType
                .AddItem "", False
                .Cell(.Rows - 1, 1).Text = "  " & IIf(IsNull(Trim(dbTabQCType!type)), "", Trim(dbTabQCType!type))
                .Cell(.Rows - 1, 1).FontBold = True
            End With
        
       
            .MoveNext
       Next
        GrdQCType.ReadOnly = True
        GrdQCType.AutoRedraw = True
        GrdQCType.Refresh
    End If
    
End With
    
    
    
GrdPH.Rows = 1

With dbTabPHMeter
    .filter = ""
    
    If .EOF Then
    
    Else
        GrdPH.AutoRedraw = False
        .MoveFirst
        For i = 1 To .RecordCount
            
            With GrdPH
                .AddItem "", False
                .Cell(.Rows - 1, 1).Text = IIf(IsNull(Trim(dbTabPHMeter!code)), "", Trim(dbTabPHMeter!code))
                .Cell(.Rows - 1, 1).FontBold = True

            End With
        
       
            .MoveNext
       Next
        
        GrdPH.ReadOnly = True
        GrdPH.AutoRedraw = True
        GrdPH.Refresh
    End If

    
End With

       
GrdTURBID.Rows = 1


With dbTabTurbMeter
    .filter = ""
    
    If .EOF Then
    
    Else
        GrdTURBID.AutoRedraw = False
        .MoveFirst
        For i = 1 To .RecordCount
            
            With GrdTURBID
                .AddItem "", False
                .Cell(.Rows - 1, 1).Text = IIf(IsNull(Trim(dbTabTurbMeter!code)), "", Trim(dbTabTurbMeter!code))
                .Cell(.Rows - 1, 1).FontBold = True
            End With
        
       
            .MoveNext
       Next
        GrdTURBID.ReadOnly = True
        GrdTURBID.AutoRedraw = True
        GrdTURBID.Refresh
    End If
    
End With
           
GrdSPECTR.Rows = 1


With dbTabSpectMeter
    .filter = ""
    
    If .EOF Then
    
    Else
        GrdSPECTR.AutoRedraw = False
        .MoveFirst
        For i = 1 To .RecordCount
            
            With GrdSPECTR
                .AddItem "", False
                .Cell(.Rows - 1, 1).Text = IIf(IsNull(Trim(dbTabSpectMeter!code)), "", Trim(dbTabSpectMeter!code))
                .Cell(.Rows - 1, 1).FontBold = True
            End With
        
       
            .MoveNext
       Next
        GrdSPECTR.ReadOnly = True
        GrdSPECTR.AutoRedraw = True
        GrdSPECTR.Refresh
    End If
    
End With


GrdDepart.Rows = 1


With dbTabDepartment
    .filter = ""
    
    If .EOF Then
    
    Else
        GrdDepart.AutoRedraw = False
        .MoveFirst
        For i = 1 To .RecordCount
            
            With GrdDepart
                .AddItem "", False
                .Cell(.Rows - 1, 1).Text = IIf(IsNull(Trim(dbTabDepartment!code)), "", Trim(dbTabDepartment!code))
                .Cell(.Rows - 1, 1).FontBold = True
                .Cell(.Rows - 1, 1).Alignment = cellLeftCenter
            End With
        
       
            .MoveNext
       Next
        GrdDepart.ReadOnly = True
        GrdDepart.AutoRedraw = True
        GrdDepart.Refresh
    End If
    
End With

End Sub

Private Sub FormPulisciTutto()
Dim Ctl As Control
    For Each Ctl In Controls
        If TypeOf Ctl Is TextBox Then
            Ctl = ""
        ElseIf TypeOf Ctl Is Label Then
            If InStr(Ctl.Caption, "SET") Then
           ' ElseIf InStr(Ctl.Caption, "METER") Then
         
          '  ElseIf InStr(Ctl.Caption, "Meter") Then
            Else
          '  Debug.Print Ctl.Caption
            Ctl.BackColor = vbColorLabelUnabled
            End If

        ElseIf TypeOf Ctl Is Grid Then
            Ctl.Rows = 1
        End If
    Next Ctl
End Sub

Private Sub GetCodeInformation(ByVal sLot As String, ByVal sCode As String, Optional ByVal MyID As Long)
Dim MeasurementUnit As String
    ' attenzione , se ho un file allora lo importo , altrimenti prendo i dati del Code
    
    With dbTabCode
        .filter = ""
        If MyID > 0 Then
            .filter = "ID='" & MyID & "'"
        Else
        
            .filter = "Code='" & sCode & "'"
        End If
    
        If .EOF Then
            MessageInfoTime = 2000
            PopupMessage 2, "Cannot find Hanna Code  :  " & sCode & vbCrLf & "Please Enter Code Information..."
            'Unload Me
        Else
        
            MeasurementUnit = IIf(IsNull(Trim(!MeasurementUnit)), "", " " & Trim(!MeasurementUnit))
            Text1(0) = sLot
            Text1(1) = sCode
            Text1(2) = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
            Text1(4) = IIf(IsNull(Trim(!Line)), "", Trim(!Line))
            Text1(5) = IIf(IsNull(Trim(!Recipe)), "", Trim(!Recipe))
            Text1(6) = IIf(IsNull(Trim(!WeightValue)), "", Trim(!WeightValue))
            Text1(7) = IIf(IsNull(Trim(!WeightMin)), "", Trim(!WeightMin))
            Text1(8) = IIf(IsNull(Trim(!WeightMax)), "", Trim(!WeightMax))
            
            Text1(9) = IIf(IsNull(Trim(!RangeMin)), "", Trim(!RangeMin))
            Text1(10) = IIf(IsNull(Trim(!RangeMax)), "", Trim(!RangeMax))
            
            lb(18) = "Reagent Range (" & MeasurementUnit & ")"
            strLastLot = IIf(IsNull(Trim(!LastLot)), "", Trim(!LastLot))
            SetLastLot (strLastLot)
            SettingName = FormatNomeFile(Text1(0) & " " & Text1(1) & " " & Text1(9) & " " & USER_ESTENSIONE)
            
            
            
            GetFormSettingName


    
        End If
    
    End With




End Sub

Private Sub SetLastLot(ByVal LastLot As String)

Dim rc As Boolean

rc = IIf(Len(LastLot) > 0, True, False)
rc = IIf(UCase(LastLot) = UCase(Text1(0)), False, rc)
Label9.Visible = rc
Image4(6).Visible = rc
DefaultMenuLabel(4).Visible = rc
Label9.Caption = "Get Last Lot Information : " & LastLot


 ' FrameIntroVisible (rc)

End Sub
Private Function SalvaFormSettingName()
Dim CreationDate As String
Dim i As Integer

    SettingName = FormatNomeFile(Text1(0) & " " & Text1(1) & " " & Text1(9) & " " & USER_ESTENSIONE)
    
    CreationDate = GetSettingData(SettingName, "File Information", "Creation Date", "")
    
    CloseSettingDataFile
    
    SaveSettingData SettingName, "File Information", "Program Name", PROGRAM_NAME
    SaveSettingData SettingName, "File Information", "Major", App.Major
    SaveSettingData SettingName, "File Information", "Minor", App.Minor
    SaveSettingData SettingName, "File Information", "Revision", App.Revision
    
    If CreationDate = "" Then
        SaveSettingData SettingName, "File Information", "Creation Date", FormatDataLAT(date)
    End If
    
    ' Set Last Lot
    With dbTabCode
        
        !LastLot = Text1(0)
        .Update
    End With
    
    With dbTabLaboratorio
        .filter = ""
        If .EOF Then
        Else
            For i = 1 To .fields.Count - 1
                SaveSettingData SettingName, "WorkStation", .fields(i).Name, Trim(.fields(i))
            Next
        End If
    End With
    
    With dbTabCode
        For i = 0 To .fields.Count - 1
            SaveSettingData SettingName, "Code Information", .fields(i).Name, Trim(.fields(i))
        Next
    End With
    SaveSettingData SettingName, "Information QC", "MeterNumber", CheckMeterNumber
    SaveSettingData SettingName, "Information QC", "pHNumber", CheckpHNumber
    SaveSettingData SettingName, "Information QC", "Operator", MyOperatore.Name
    SaveSettingData SettingName, "Information QC", "Modification Date", FormatDataLAT(date)
    For i = 0 To Text1.Count - 1
        SaveSettingData SettingName, "Information QC", "Text1" & i, Text1(i).Text
    Next
    
    Dim rc As Boolean
    rc = False
    For i = 11 To 15
        If Trim(Text1(i)) <> "" Then rc = True
    Next
    SaveSettingData SettingName, "Information QC", "Reagent Set 1", rc
    
    rc = False
    For i = 45 To 49
        If Trim(Text1(i)) <> "" Then rc = True
    Next
    SaveSettingData SettingName, "Information QC", "Reagent Set 2", rc
        

    CloseSettingDataFile
End Function
Private Function CheckpHNumber() As Integer
Dim i As Integer
Dim Num As Integer
Num = 0
With dbTabCode
    For i = 38 To 46 Step 3
       ' If (.fields(i) <> "" Or .fields(i) <> "/") And IsNumeric(.fields(i)) Then
            Num = Num + 1
            SaveSettingData SettingName, "Information QC", "pHValue" & Num, Trim(.fields(i))
            SaveSettingData SettingName, "Information QC", "pHMin" & Num, Trim(.fields(i + 1))
            SaveSettingData SettingName, "Information QC", "pHMax" & Num, Trim(.fields(i + 2))
       ' End If
    Next
End With
CheckpHNumber = Num
End Function
Private Function CheckMeterNumber() As Integer
Dim i As Integer
Dim Num As Integer
Num = 0
For i = 33 To 39 Step 2
    If Text1(i) <> "" Then
        Num = Num + 1
    End If
Next
CheckMeterNumber = Num
End Function
Private Function GetFormSettingName()
Dim i As Integer

    USER_PATH = USER_TEMP_PATH
  
    If bSearchClosedLot Then
    
       If FileExists(USER_TEMP_PATH & SettingName) Then
       ElseIf FileExists(USER_DATA_PATH & SettingName) Then
           
            USER_PATH = USER_DATA_PATH
       Else
            Exit Function
       End If
   
    End If
   
   
   
    CloseSettingDataFile
    Debug.Print USER_PATH
    For i = 0 To Text1.Count - 1
       Text1(i) = GetSettingData(SettingName, "Information QC", "Text1" & i, Text1(i))
    Next
    
    ' ora so quanti meter ho inserito...
    GetReadingInformation
    
    
    FrameIntroVisible False
    CloseSettingDataFile
End Function

Private Function GetLastLotInformation()


    
    SettingName = FormatNomeFile(strLastLot & " " & Text1(1) & " " & Text1(9) & " " & USER_ESTENSIONE)
    GetFormSettingName
    Text1(0) = strLot
End Function


Private Function GetTabInstrument()
Dim rc As Boolean
Dim i As Integer
Dim MeterFamily(2) As Variant


    With dbTabCode
         MeterFamily(0) = IIf(IsNull(Trim(!MeterFamily1)), "", Trim(!MeterFamily1))
         MeterFamily(1) = IIf(IsNull(Trim(!MeterFamily2)), "", Trim(!MeterFamily2))
         If InStr(MeterFamily(0), "/") Then MeterFamily(0) = ""
         If InStr(MeterFamily(1), "/") Then MeterFamily(1) = ""

    End With
    
    With GrdMeter
        .AutoRedraw = False
        .Rows = 1
        For i = 0 To UBound(MeterFamily)
            If MeterFamily(i) <> "" Then
                .AddItem "", False
                .Cell(.Rows - 1, 1).Text = MeterFamily(i)
                .Cell(.Rows - 1, 1).FontBold = True
            End If
        Next
        .AutoRedraw = True
        .Refresh
    End With

End Function

Private Sub GetReadingInformation()

    MeterNumber = GetSettingData(SettingName, "Information QC", "MeterNumber", 0)

    ReloadMeter.Visible = IIf(MeterNumber > 0, True, False)

End Sub


Private Sub GetAllMeterInformation()

Dim i As Integer

    
    For i = 1 To 8
    Text1(i + 32) = GetSettingData(SettingName, "Information QC", "Text1" & i + 32, "")
    Next

    CloseSettingDataFile

End Sub



Private Function CheckReadingQC() As Boolean

Dim rc              As Boolean
Dim MeterInfo(8)    As String
Dim i               As Integer
Dim NumReadings     As Integer

    rc = True
    
    For i = 1 To 8
        MeterInfo(i) = GetSettingData(SettingName, "Information QC", "Text1" & i + 32, "")
    Next
    
    ' controllo se sono compatibili con TEXT1
    NumReadings = GetSettingData(SettingName, "Reading QC", "Grd2 Rows", 0)
    CloseSettingDataFile
    If NumReadings > 0 Then
        ' ho già fatto almeno 1 readings
        '  rc = False
        
        For i = 1 To MeterNumber * 2 Step 2
            If Text1(i + 32) <> MeterInfo(i) Then
                ' ci sono differenze tra inseriti e Text1....
                rc = False
            End If
        Next
    
    
    End If
    
    CheckReadingQC = rc

End Function
