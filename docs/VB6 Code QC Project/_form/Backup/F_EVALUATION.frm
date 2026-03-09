VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form F_EVALUATION 
   BackColor       =   &H00C0C0C0&
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
   StartUpPosition =   2  'CenterScreen
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
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "3.09"
         Top             =   1080
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
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "2.4"
         Top             =   1080
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
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "0.4"
         Top             =   1080
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
         Index           =   10
         Left            =   11880
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "3.23"
         Top             =   1080
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
      Begin VB.Label Label1 
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
         Index           =   8
         Left            =   1200
         TabIndex        =   54
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "pH"
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
         Left            =   9960
         TabIndex        =   50
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00004080&
         Caption         =   "STD Target Value "
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
         Index           =   33
         Left            =   5880
         TabIndex        =   49
         Top             =   240
         Width           =   3735
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
         Index           =   10
         Left            =   5880
         TabIndex        =   48
         Top             =   720
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
         Index           =   11
         Left            =   7800
         TabIndex        =   47
         Top             =   720
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
         Index           =   12
         Left            =   9960
         TabIndex        =   46
         Top             =   720
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
         Index           =   13
         Left            =   11880
         TabIndex        =   45
         Top             =   720
         Width           =   1815
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
   Begin VB.PictureBox PicMenuBar 
      BackColor       =   &H00404040&
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
         Index           =   0
         Left            =   0
         MouseIcon       =   "F_EVALUATION.frx":218CF
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
            MouseIcon       =   "F_EVALUATION.frx":21BD9
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":21EE3
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Mean Value"
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
            MouseIcon       =   "F_EVALUATION.frx":252C5
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
         MouseIcon       =   "F_EVALUATION.frx":255CF
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   11
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Results"
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
            MouseIcon       =   "F_EVALUATION.frx":258D9
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
            Picture         =   "F_EVALUATION.frx":25BE3
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Information QC"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   15870
         TabIndex        =   15
         Top             =   360
         Width           =   2820
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   0
      Left            =   0
      Picture         =   "F_EVALUATION.frx":28FC5
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
         Index           =   6
         Left            =   14760
         TabIndex        =   38
         Top             =   5640
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
         Index           =   5
         Left            =   14760
         TabIndex        =   36
         Top             =   4560
         Width           =   3975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1575
         Left            =   4080
         TabIndex        =   24
         Top             =   2520
         Width           =   11295
         Begin VB.Shape Shape1 
            BackColor       =   &H000080DF&
            BorderColor     =   &H000060BF&
            Height          =   1575
            Left            =   0
            Top             =   0
            Width           =   11295
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3 - Check Results"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   2640
            TabIndex        =   27
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2 - Select Value from Readings Table"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   2640
            TabIndex        =   26
            Top             =   600
            Width           =   3705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1 - Select Standard from SFG Standard Table"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   2640
            TabIndex        =   25
            Top             =   240
            Width           =   4545
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   1440
            Picture         =   "F_EVALUATION.frx":46E9E
            Top             =   480
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
         TabIndex        =   34
         Top             =   3480
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
         TabIndex        =   32
         Top             =   2400
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
         Index           =   2
         Left            =   14760
         TabIndex        =   30
         Top             =   1320
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
         Top             =   6600
         Width           =   3975
         Begin VB.Image ImageTAV 
            Height          =   480
            Index           =   5
            Left            =   1760
            MouseIcon       =   "F_EVALUATION.frx":4A280
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":4A58A
            Top             =   160
            Width           =   480
         End
      End
      Begin FlexCell.Grid Grd1 
         Height          =   6600
         Left            =   240
         TabIndex        =   0
         Top             =   960
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   11642
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
         DefaultFontBold =   -1  'True
         DisplayDateTimeMask=   -1  'True
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   16777215
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
         Left            =   5760
         TabIndex        =   42
         Top             =   960
         Width           =   8640
         _ExtentX        =   15240
         _ExtentY        =   11642
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
         DefaultFontBold =   -1  'True
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Selected Tests"
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
         Left            =   14760
         TabIndex        =   39
         Top             =   5280
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Tests Count"
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
         TabIndex        =   37
         Top             =   4200
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Mean Value STD"
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
         TabIndex        =   35
         Top             =   3120
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "STD Value"
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
         TabIndex        =   33
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Mean Value ALL"
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
         TabIndex        =   31
         Top             =   960
         Width           =   3975
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   1
         Left            =   5760
         Picture         =   "F_EVALUATION.frx":4D96C
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
         Left            =   6360
         MouseIcon       =   "F_EVALUATION.frx":50D4E
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
         MouseIcon       =   "F_EVALUATION.frx":51058
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   480
         Width           =   2100
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   3
         Left            =   240
         Picture         =   "F_EVALUATION.frx":51362
         Top             =   360
         Width           =   480
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Left            =   9360
         Picture         =   "F_EVALUATION.frx":54744
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
   Begin VB.PictureBox PicMain 
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   1
      Left            =   0
      Picture         =   "F_EVALUATION.frx":57B26
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   22
      Top             =   2880
      Width           =   19215
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         FillColor       =   &H00004000&
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   7560
         ScaleHeight     =   855
         ScaleWidth      =   3975
         TabIndex        =   41
         Top             =   6240
         Visible         =   0   'False
         Width           =   3975
         Begin VB.Image ImageTAV 
            Height          =   480
            Index           =   0
            Left            =   1720
            MouseIcon       =   "F_EVALUATION.frx":759FF
            MousePointer    =   99  'Custom
            Picture         =   "F_EVALUATION.frx":75D09
            Top             =   160
            Width           =   480
         End
      End
      Begin FlexCell.Grid Grd3 
         Height          =   3600
         Left            =   4680
         TabIndex        =   43
         Top             =   1800
         Width           =   9840
         _ExtentX        =   17357
         _ExtentY        =   6350
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
         DefaultFontBold =   -1  'True
         DisplayDateTimeMask=   -1  'True
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   16777215
         ReadOnly        =   -1  'True
         Rows            =   10
         SelectionMode   =   1
         MultiSelect     =   0   'False
         DateFormat      =   2
         EnterKeyMoveTo  =   1
         BackColorComment=   -2147483635
         AllowUserPaste  =   2
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Results Table"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   5280
         MouseIcon       =   "F_EVALUATION.frx":790EB
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   2
         Left            =   4680
         Picture         =   "F_EVALUATION.frx":793F5
         Top             =   1200
         Width           =   480
      End
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   4
      Left            =   2040
      MouseIcon       =   "F_EVALUATION.frx":7C7D7
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":7CAE1
      Top             =   11040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   0
      Left            =   17640
      MouseIcon       =   "F_EVALUATION.frx":7FEC3
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
      MouseIcon       =   "F_EVALUATION.frx":801CD
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   10680
      Width           =   2175
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Index           =   1
      Left            =   8280
      MouseIcon       =   "F_EVALUATION.frx":804D7
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   10560
      Width           =   2655
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   15960
      MouseIcon       =   "F_EVALUATION.frx":807E1
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":80AEB
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_EVALUATION.frx":83ECD
      Height          =   480
      Index           =   1
      Left            =   9360
      MouseIcon       =   "F_EVALUATION.frx":872AF
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":875B9
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   480
      MouseIcon       =   "F_EVALUATION.frx":8A99B
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":8ACA5
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
      DragIcon        =   "F_EVALUATION.frx":8E087
      Height          =   480
      Index           =   0
      Left            =   18240
      MouseIcon       =   "F_EVALUATION.frx":91469
      MousePointer    =   99  'Custom
      Picture         =   "F_EVALUATION.frx":91773
      Top             =   11040
      Width           =   480
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   3
      Left            =   0
      MouseIcon       =   "F_EVALUATION.frx":94B55
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   10560
      Width           =   1815
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
Private IndexDate As Integer
Private IndexText As Integer
Private MyLot As String
Private MyCode As String
Private m_rc As Boolean
Private bFormSaved As Boolean

Public Function DoShow(ByRef Index As Integer, Optional ByRef sLot As String, Optional ByRef sCode As String, Optional MyImage As Image) As Boolean

    On Error GoTo ERR_SHOW
    
    Set DefaultMenu(4) = MyImage
    IndexMainProcedura = Index
    m_rc = False
    bFormSaved = False
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
        'If F_InputBox.DoShow("Set Operator QC", , , , , MyOperatore.Name) Then
        '    Label5 = MyOperatore.Name
        'End If
    Case 4
 
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

Call SetPicForm
Call SetGrid
End Sub

Private Sub Form_Load()
IndexFormProcedura = 99
End Sub

Private Sub Frame2_Click()
Frame2.Visible = False
End Sub

Private Sub Grd1_Click()
Frame2.Visible = False
End Sub

Private Sub Grd1_LostFocus()
'Grd1.Cell(0, 0).SetFocus
End Sub

Private Sub Grd2_Click()
Frame2.Visible = False
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
        PicMenu(i).BackColor = &H404040
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

    If Index > 3 Then Exit Function
    For i = 0 To PicMenu.Count - 1
        If i = Index Then
            PicMenu(i).BackColor = &H606060
        Else
            PicMenu(i).BackColor = &H404040
            
        End If
    Next
    Set Image4(0) = Image3(Index)
    Select Case Index
        Case 0
            Picture4(0).BackColor = &H4080&
           ' rc = False
        Case 1
           ' rc = True
            Picture4(0).BackColor = &H4060&
        Case 2
           ' rc = False
           ' Picture4(0).BackColor = &H60DF&
    End Select
    Label2(4) = Label2(Index)
    IndexFormProcedura = Index
    PicMain(Index).Visible = True
    PicMain(Index).ZOrder
    blTable = "Evaluation QC : " & Label2(IndexFormProcedura)
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
        PicMenu(i).BackColor = &H404040
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

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BackColor = &H8000&
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SaveProcedure
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.BackColor = &H8000&
End Sub

Private Sub Text1_Change(Index As Integer)
Dim rc As Boolean
rc = IIf(Len(Text1(Index)) > 0, True, False)
Text1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbWhite, vbColorUnabled)
Select Case Index
    Case 9
        Frame2.Visible = Not (rc)

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

Select Case Index

End Select
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
        PopupMessage 2, blTable & " Saved...", , , , DefaultMenu(4)
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
        '       SET TABELLA SEDI
        '------------------------------------------------
    With Grd1
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        .DefaultFont.Size = 14
        .DefaultRowHeight = 40
        .Cols = 3
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Standard"
        .Column(1).Width = 150
        .Cell(0, 2).Text = "STD Value"
        .Column(2).Width = 200

        
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
    
    With Grd2
      .Rows = 1

        .AutoRedraw = False
        .ReadOnly = False
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .DefaultFont.Size = 14
        .DefaultRowHeight = 40
        .Cols = 9
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        
        .Range(0, 1, 0, 2).Merge
        .Cell(0, 1).Text = "Meter 1"
        .Column(1).CellType = cellCheckBox
        .Column(1).Width = 30
        .Column(2).Width = 110
        
        .Range(0, 3, 0, 4).Merge
        .Cell(0, 3).Text = "Meter 2"
        .Column(3).CellType = cellCheckBox
        .Column(3).Width = 30
        .Column(4).Width = 110
         
        .Range(0, 5, 0, 6).Merge
        .Cell(0, 5).Text = "Meter 3"
        .Column(5).CellType = cellCheckBox
        .Column(5).Width = 30
        .Column(6).Width = 120
        
        
        .Range(0, 7, 0, 8).Merge
        .Cell(0, 7).Text = "pH"
        .Column(7).CellType = cellCheckBox
        .Column(7).Width = 30
        .Column(8).Width = 110
        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
            
        Next
       ' .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
    
    With Grd3
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .DefaultFont.Size = 14
        .DefaultRowHeight = 40
        .Cols = 4
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Standard Value"
        .Column(1).Width = 200
        .Cell(0, 2).Text = "Target Value " & Chr$(177) & " U [ppm]"
        .Column(2).Width = 200
        .Cell(0, 3).Text = "Mean Value [ppm]"
        .Column(3).Width = 200
        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With

    Dim t As Integer
    With Grd1
        .Rows = 4
        .Cell(1, 1).Text = "#2"
        .Cell(1, 2).Text = "0.2"
        .Cell(2, 1).Text = "#3"
        .Cell(2, 2).Text = "0.55"
        .Cell(3, 1).Text = "#4"
        .Cell(3, 2).Text = "0.75"
        For i = 0 To .Rows - 1
            For t = 0 To .Cols - 1
                .Cell(i, t).ForeColor = &H404040
                .Cell(i, t).Alignment = cellCenterCenter
                
            Next
        Next
    End With
    
    With Grd2
        .Rows = 4
        .Cell(1, 1).Text = True
        .Cell(1, 2).Text = "0.21"
        .Cell(1, 3).Text = True
        .Cell(1, 4).Text = "0.2"
        
        .Cell(1, 6).Text = ""
        .Cell(1, 7).Text = True
        .Cell(1, 8).Text = "3.72"
        .Cell(2, 1).Text = True
        .Cell(2, 2).Text = "0.51"
        .Cell(2, 3).Text = True
        .Cell(2, 4).Text = "0.28"
        
        .Cell(2, 6).Text = ""
        .Cell(2, 8).Text = "3.33"
        .Cell(2, 7).Text = True
        .Cell(3, 1).Text = True
        .Cell(3, 2).Text = "0.28"
        .Cell(3, 3).Text = True
        .Cell(3, 4).Text = "0.62"
        
        .Cell(3, 6).Text = ""
        .Cell(3, 7).Text = True
        .Cell(3, 8).Text = "3.62"


        For i = 0 To .Rows - 1
            For t = 0 To .Cols - 1
                .Cell(i, t).ForeColor = &H404040
                .Cell(i, t).Alignment = cellCenterCenter
                
            Next
        Next
    End With
    
    With Grd3
        .Rows = 4
        .Cell(1, 1).Text = "#2"
        .Cell(1, 2).Text = "0.2"
        .Cell(2, 1).Text = "#3"
        .Cell(2, 2).Text = "0.55"
        .Cell(3, 1).Text = "#4"
        .Cell(3, 2).Text = "0.75"
        For i = 0 To .Rows - 1
            For t = 0 To .Cols - 1
                .Cell(i, t).ForeColor = &H404040
                .Cell(i, t).Alignment = cellCenterCenter
                
            Next
        Next
    End With
End Function

