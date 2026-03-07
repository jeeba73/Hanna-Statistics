VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form F_MAIN 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Chemical QC"
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
   Picture         =   "F_MAIN.frx":0000
   ScaleHeight     =   12000
   ScaleMode       =   0  'User
   ScaleWidth      =   19200
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicIntro 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   9690
      Left            =   18720
      ScaleHeight     =   9690
      ScaleWidth      =   19200
      TabIndex        =   15
      Top             =   2760
      Width           =   19200
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7680
         Left            =   840
         Picture         =   "F_MAIN.frx":1DED9
         ScaleHeight     =   7680
         ScaleWidth      =   7680
         TabIndex        =   16
         Top             =   720
         Width           =   7680
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
         Left            =   16080
         TabIndex        =   40
         Top             =   5520
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
         Left            =   8880
         TabIndex        =   17
         Top             =   3600
         Width           =   8085
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8655
      Index           =   0
      Left            =   1560
      Picture         =   "F_MAIN.frx":3C9DA
      ScaleHeight     =   8655
      ScaleWidth      =   19215
      TabIndex        =   30
      Top             =   1680
      Visible         =   0   'False
      Width           =   19215
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   2
         Left            =   3000
         TabIndex        =   52
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   9480
         TabIndex        =   44
         Top             =   3000
         Width           =   2895
      End
      Begin VB.ComboBox CmbVisual 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   360
         Width           =   7695
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   1
         Left            =   9480
         TabIndex        =   32
         Top             =   3600
         Width           =   2895
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   0
         Left            =   4800
         MouseIcon       =   "F_MAIN.frx":5A8B3
         MousePointer    =   99  'Custom
         ScaleHeight     =   1335
         ScaleWidth      =   9615
         TabIndex        =   41
         Top             =   4920
         Width           =   9615
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start New Lot"
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
            Index           =   4
            Left            =   0
            MouseIcon       =   "F_MAIN.frx":5ABBD
            MousePointer    =   99  'Custom
            TabIndex        =   42
            Top             =   840
            Width           =   9585
         End
         Begin VB.Image Image4 
            Height          =   480
            Index           =   0
            Left            =   4560
            Picture         =   "F_MAIN.frx":5AEC7
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Code"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   53
         Top             =   360
         Width           =   2070
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Lot : H156"
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
         Left            =   960
         MouseIcon       =   "F_MAIN.frx":5E2A9
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   480
         Width           =   1485
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   6
         Left            =   360
         Picture         =   "F_MAIN.frx":5E5B3
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   5
         Left            =   16800
         Picture         =   "F_MAIN.frx":61995
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   4
         Left            =   13920
         Picture         =   "F_MAIN.frx":64D77
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   3
         Left            =   11040
         Picture         =   "F_MAIN.frx":68159
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   11640
         MouseIcon       =   "F_MAIN.frx":6B53B
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Number"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   495
         Index           =   1
         Left            =   6840
         TabIndex        =   45
         Top             =   3000
         Width           =   1995
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Left            =   9360
         Picture         =   "F_MAIN.frx":6B845
         Top             =   5400
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   1815
         Left            =   4800
         Top             =   2640
         Width           =   9615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closed Lot"
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
         Left            =   14520
         MouseIcon       =   "F_MAIN.frx":6EC27
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Open Lot"
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
         Height          =   255
         Left            =   17400
         MouseIcon       =   "F_MAIN.frx":6EF31
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   480
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   13200
         MouseIcon       =   "F_MAIN.frx":6F23B
         MousePointer    =   99  'Custom
         Picture         =   "F_MAIN.frx":6F545
         Top             =   3240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operator"
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   9120
         TabIndex        =   33
         Top             =   8160
         Width           =   975
      End
      Begin VB.Image ImMain 
         Height          =   480
         Left            =   9360
         Picture         =   "F_MAIN.frx":72927
         Top             =   7680
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Code"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   495
         Index           =   0
         Left            =   6840
         TabIndex        =   31
         Top             =   3600
         Width           =   2010
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   5
         Left            =   13800
         TabIndex        =   37
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   6
         Left            =   16680
         TabIndex        =   38
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   8
         Left            =   120
         TabIndex        =   49
         Top             =   7440
         Width           =   2655
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   7
         Left            =   10800
         TabIndex        =   47
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   3
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   21
      Top             =   5640
      Visible         =   0   'False
      Width           =   19215
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   975
         Index           =   3
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   255
         TabIndex        =   29
         Top             =   0
         Width           =   255
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   0
         X2              =   19440
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         Index           =   3
         X1              =   -240
         X2              =   19200
         Y1              =   960
         Y2              =   960
      End
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
         Width           =   6210
      End
      Begin VB.Image Image3 
         Height          =   480
         Index           =   7
         Left            =   240
         MouseIcon       =   "F_MAIN.frx":75D09
         MousePointer    =   99  'Custom
         Picture         =   "F_MAIN.frx":76013
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H0070B0F0&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   2
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   19215
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         ForeColor       =   &H00008000&
         Height          =   975
         Index           =   2
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   255
         TabIndex        =   28
         Top             =   0
         Width           =   255
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   0
         X2              =   19440
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00004080&
         Index           =   2
         X1              =   0
         X2              =   19440
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Evaluation QC : Select Lot Number , Hanna Code and Start Evaluation"
         ForeColor       =   &H00004080&
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   24
         Top             =   360
         Width           =   7110
      End
      Begin VB.Image Image3 
         Height          =   480
         Index           =   6
         Left            =   240
         MouseIcon       =   "F_MAIN.frx":793F5
         MousePointer    =   99  'Custom
         Picture         =   "F_MAIN.frx":796FF
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H008FC9FA&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   1
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   19215
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000080DF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00008000&
         Height          =   975
         Index           =   1
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   255
         TabIndex        =   27
         Top             =   0
         Width           =   255
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   0
         X2              =   19440
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000080DF&
         Index           =   1
         X1              =   0
         X2              =   19440
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reading QC : Select Lot Number , Hanna Code and Start Tests"
         ForeColor       =   &H000040C0&
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   23
         Top             =   360
         Width           =   6255
      End
      Begin VB.Image Image3 
         Height          =   480
         Index           =   5
         Left            =   240
         MouseIcon       =   "F_MAIN.frx":7CAE1
         MousePointer    =   99  'Custom
         Picture         =   "F_MAIN.frx":7CDEB
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H00A0DFA0&
      BorderStyle     =   0  'None
      DrawWidth       =   7
      Height          =   975
      Index           =   0
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   19215
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00008000&
         Height          =   975
         Index           =   0
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   255
         TabIndex        =   26
         Top             =   0
         Width           =   255
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   0
         X2              =   19440
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00008000&
         Index           =   0
         X1              =   0
         X2              =   19440
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Information QC : Enter Lot Number , Hanna Code and fill information QC"
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   22
         Top             =   360
         Width           =   7410
      End
      Begin VB.Image Image3 
         Height          =   480
         Index           =   4
         Left            =   240
         MouseIcon       =   "F_MAIN.frx":801CD
         MousePointer    =   99  'Custom
         Picture         =   "F_MAIN.frx":804D7
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox PicMenu 
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
         Index           =   3
         Left            =   5760
         MouseIcon       =   "F_MAIN.frx":838B9
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
            MouseIcon       =   "F_MAIN.frx":83BC3
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":83ECD
            Top             =   180
            Width           =   480
         End
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
            MouseIcon       =   "F_MAIN.frx":872AF
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
         MouseIcon       =   "F_MAIN.frx":875B9
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
            MouseIcon       =   "F_MAIN.frx":878C3
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":87BCD
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Evaluation QC"
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
            MouseIcon       =   "F_MAIN.frx":8AFAF
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
         Index           =   1
         Left            =   1920
         MouseIcon       =   "F_MAIN.frx":8B2B9
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
            MouseIcon       =   "F_MAIN.frx":8B5C3
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":8B8CD
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
            Index           =   1
            Left            =   0
            MouseIcon       =   "F_MAIN.frx":8ECAF
            MousePointer    =   99  'Custom
            TabIndex        =   10
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
         MouseIcon       =   "F_MAIN.frx":8EFB9
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
            Caption         =   "Information QC"
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
            Left            =   75
            MouseIcon       =   "F_MAIN.frx":8F2C3
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   720
            Width           =   1860
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   735
            MouseIcon       =   "F_MAIN.frx":8F5CD
            MousePointer    =   99  'Custom
            Picture         =   "F_MAIN.frx":8F8D7
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
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   13920
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   4770
      End
   End
   Begin ChemicalQC.ctlCalendar ctlCalendar1 
      Height          =   6960
      Left            =   18000
      TabIndex        =   43
      Top             =   1200
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
   Begin FlexCell.Grid GrdCode 
      Height          =   3480
      Left            =   2040
      TabIndex        =   51
      Top             =   1080
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
      DefaultFontBold =   -1  'True
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
   Begin FlexCell.Grid GrdBatch 
      Height          =   3480
      Left            =   120
      TabIndex        =   50
      Top             =   120
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
      DefaultFontSize =   9.75
      DefaultFontBold =   -1  'True
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
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1815
      Index           =   2
      Left            =   13080
      MouseIcon       =   "F_MAIN.frx":92CB9
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
      MouseIcon       =   "F_MAIN.frx":92FC3
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   10800
      Width           =   1455
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Index           =   1
      Left            =   8280
      MouseIcon       =   "F_MAIN.frx":932CD
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   10560
      Width           =   2655
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   3
      Left            =   3840
      MouseIcon       =   "F_MAIN.frx":935D7
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   10560
      Width           =   1815
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   4
      Left            =   0
      MouseIcon       =   "F_MAIN.frx":938E1
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   10680
      Width           =   1695
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   14160
      MouseIcon       =   "F_MAIN.frx":93BEB
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":93EF5
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_MAIN.frx":972D7
      Height          =   480
      Index           =   1
      Left            =   9360
      MouseIcon       =   "F_MAIN.frx":9A6B9
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":9A9C3
      Top             =   11040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   4560
      MouseIcon       =   "F_MAIN.frx":9DDA5
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":9E0AF
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   4
      Left            =   480
      MouseIcon       =   "F_MAIN.frx":A1491
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":A179B
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
      Y1              =   360
      Y2              =   12240
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_MAIN.frx":A4B7D
      Height          =   480
      Index           =   0
      Left            =   18240
      MouseIcon       =   "F_MAIN.frx":A7F5F
      MousePointer    =   99  'Custom
      Picture         =   "F_MAIN.frx":A8269
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

Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type

Private m_ControlGridFontSize As Double
Private m_ControlGridRowHeight As Double
Private m_ControlGridColWidth As Double
Private m_ControlPositions() As ControlPositionType
Private m_FormWid As Single
Private m_FormHgt As Single

Private lRowCode As Long
Private lRow As Long

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

m_ControlGridFontSize = y_scale * 0.95
m_ControlGridColWidth = x_scale * 0.9
m_ControlGridRowHeight = y_scale * 0.95

'If Not (bStazioneEsterna) Then
'm_ControlGridFontSize = 1 ' y_scale * 0.8
'm_ControlGridColWidth = 1 ' x_scale * 0.9

'End If
'm_ControlGridRowHeight = 1 '1.4


For Each Ctl In Controls
    With m_ControlPositions(i)
        If TypeOf Ctl Is Line Then
            Ctl.X1 = x_scale * .Left
            Ctl.Y1 = y_scale * .Top
            Ctl.X2 = Ctl.X1 + x_scale * .Width
            Ctl.Y2 = Ctl.Y1 + y_scale * .Height
        
        ElseIf TypeOf Ctl Is Grid Then
           Ctl.Left = x_scale * .Left
            Ctl.Top = y_scale * .Top
            Ctl.Width = x_scale * .Width
            Ctl.Height = y_scale * .Height

                Ctl.DefaultFont.Size = 16 * m_ControlGridFontSize
                Ctl.DefaultRowHeight = 30 * m_ControlGridRowHeight
           
        Else
            Ctl.Left = x_scale * .Left
            'MsgBox (TypeName(ctl))
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
Private Sub DefaultMenuLabel_Click(Index As Integer)
Select Case Index
    Case 0
        Unload Me
    Case 2
        MessageInfoTime = 3000
        PopupMessage 2, "Messaggio in popup Messaggio in popup Messaggio in popupssaggio in popup Messaggio in popup" & vbCrLf & "essaggio in popup Messaggio in popup Messaggio in popup "
    Case 3
        If F_InputBox.DoShow("Set Operator QC", , , , , MyOperatore.Name) Then
            Label5 = MyOperatore.Name
        End If
     Case 4
 
    Case 5
        Label7_Click
    Case 6
        Label6_Click
    Case 7
        Label8_Click
End Select

End Sub

Private Sub DisableImage_Click()
PopupMessage 2, "Warning : Administrator Only can Operate...", , True
End Sub

Private Sub Form_Initialize()
lbProgram = "Release " & App.Major & "." & App.Minor & "." & App.Revision

Call StartProcedure

'SaveSizes
End Sub

Private Sub Form_Load()
IndexProcedura = 99

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 3
    If i = IndexProcedura Then
    Else
        PicMenu(i).BackColor = &H404040
    End If
Next
End Sub

Private Sub Form_Resize()
'SetPicForm
'ResizeControls
End Sub

Private Sub Image2_Click(Index As Integer)
Select Case Index
    Case 0
        PopupMessage 2, "Add New Lot..."
    Case 1
        PopupMessage 2, "Delete Lot..."
End Select

End Sub

Private Sub GrdBatch_DblClick()
lRowCode = 0
If lRow > 0 Then
    
    lRow = 0
    CleanformLot (False)
    blTable.Visible = False
End If
' copia il codice in Text1(1)

End Sub



Private Sub GrdBatch_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

lRow = FirstRow
If lRow > 0 Then

    Text1(0) = GrdBatch.Cell(lRow, 1).Text
    Text1(1) = GrdBatch.Cell(lRow, 2).Text
Else

End If

End Sub

Private Sub GrdCode_DblClick()
lRow = 0
If lRowCode > 0 Then
    lRowCode = 0
    Label8_Click
End If
End Sub

Private Sub GrdCode_LostFocus()
GrdCode.Cell(0, 0).SetFocus
End Sub

Private Sub GrdCode_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
lRowCode = FirstRow
If FirstRow > 0 Then
    Text1(2) = GrdCode.Cell(FirstRow, 1).Text
Else

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
If Index = 0 Then Picture4_Click 0
End Sub

Private Sub Label2_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub Label6_Click()
Dim rc As Boolean
rc = Not (GrdBatch.Visible)
CleanformLot (Not (GrdBatch.Visible))
blTable = "Open Lot"
blTable.Visible = rc
End Sub


Private Sub Label7_Click()
Dim rc As Boolean
rc = Not (GrdBatch.Visible)
blTable = "Closed Lot"
Call CleanformLot(Not (GrdBatch.Visible), 1)
blTable.Visible = rc
End Sub

Private Sub Label8_Click()
Dim rc As Boolean
rc = Not (GrdCode.Visible)
blTable = "Hanna SFG Code"
Call CleanformCode(Not (GrdCode.Visible), 1)
blTable.Visible = rc
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
If Index > 3 Then Exit Function
For i = 0 To 3
    If i = Index Then
        PicMenu(i).BackColor = &H606060
        PicInfo(i).Visible = True
    Else
        PicInfo(i).Visible = False
        PicMenu(i).BackColor = &H404040
    End If
Next
Set Image4(0) = Image3(Index)
Picture4(0).BackColor = Picture1(Index).BackColor
Label2(4) = "Start : " & Label2(Index)
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
PicIntro.Visible = False
PicMain(0).Visible = True
CleanformLot (False)
'blTable.ForeColor = Picture1(Index).BackColor
End Function

Private Sub CleanformLot(ByVal bValue As Boolean, Optional ByVal Index As Integer = 0)

GrdBatch.Visible = bValue
'GrdCode.Visible = False
CmbVisual.Visible = IIf(Index = 1, GrdBatch.Visible, False)
If bValue Then CleanformCode False

End Sub



Private Sub CleanformCode(ByVal bValue As Boolean, Optional ByVal Index As Integer = 0)

    Label9.Visible = Not (bValue)
    Image4(6).Visible = Not (bValue)

    GrdCode.Visible = bValue
   ' GrdBatch.Visible = False

    Text1(2).Visible = bValue
    Label4(2).Visible = bValue
    
    If bValue Then
    
       CleanformLot False
        
       ' If Text1(1) <> "" Then
            Text1(2) = Text1(1)
       ' End If
        Text1(2).SetFocus
    End If


End Sub

Private Sub FormSelected(ByVal bValue As Boolean)
        DefaultMenuLabel(5).Visible = bValue
        Label8.Visible = bValue
        Image4(3).Visible = bValue
End Sub
Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 3
    If i = IndexProcedura Then
        ' lascialo stare...
    ElseIf i = Index Then
        PicMenu(i).BackColor = &H505050
    Else
        PicMenu(i).BackColor = &H404040
    End If
Next
End Sub


Private Sub SetPicForm()
On Error GoTo ERR_SET:
ctlCalendar1.Left = Me.Width / 2 - ctlCalendar1.Width / 2
ctlCalendar1.Top = Me.Height / 2 - ctlCalendar1.Height / 2
PicMain(0).Left = 0
PicMain(0).Top = 2063
PicMain(0).Width = Me.Width
PicMain(0).Height = Line1.Y1 - PicMain(0).Top
PicIntro.Left = 0
PicIntro.Top = PicMenu(4).Height
PicIntro.Width = Me.Width
PicIntro.Height = Line1.Y1 - PicIntro.Top
PicIntro.BackColor = &H929292
PicInfo(3).BackColor = vbColorTextLightBlue
Label1(3).ForeColor = vbColorTextBlue
Picture1(3).BackColor = vbColorTextBlue
Line3(3).BorderColor = vbColorTextBlue
Picture4(2).BackColor = vbColorTextBlue
GrdBatch.Left = 360
GrdBatch.Top = 3060
GrdBatch.Width = Me.Width - GrdBatch.Left * 2
GrdBatch.Height = Line1.Y1 - GrdBatch.Top - 180
GrdCode.Top = GrdBatch.Top
GrdCode.Left = GrdBatch.Left
GrdCode.Width = GrdBatch.Width
GrdCode.Height = GrdBatch.Height

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

Private Sub Picture4_Click(Index As Integer)
Dim MyNewIndex As Integer
Dim Frm As Form

Dim rc As Boolean
    Select Case Index
        Case 0
            '-------------------------------------
            ' apro la procedura selezionata
            '-------------------------------------
            
            'MessageInfoTime = 500
            
           ' PopupMessage 2, Label2(4), , , Label2(IndexProcedura), Image4(0)
 
 
            Select Case IndexProcedura
                Case 0
                    Set Frm = F_INFORMATION
                   
                Case 1
                    Set Frm = F_READING
                Case 2
                    Set Frm = F_EVALUATION
                
            End Select
            
    

                   Frm.Top = Me.Top
                   Frm.Left = Me.Left
                   MyNewIndex = IndexProcedura
                   rc = Frm.DoShow(MyNewIndex, MyLot, MyCode, Image3(IndexProcedura))
                   GoSub CheckNewProcedura
                      
            
        Case 1
            ' tabello lotti chiusi
            Label7_Click
        Case 2
            GrdCode.Visible = Not (GrdCode.Visible)
        Case 4

 
    End Select
    
    Exit Sub
    
CheckNewProcedura:

            If rc Then
            
                Text1(0) = MyLot
                Text1(1) = MyCode
                
                If MyNewIndex <> IndexProcedura Then
                    '-------------------------------------
                    ' apro la nuova procedura
                    '-------------------------------------
                    IndexProcedura = MyNewIndex
                    Call SelectProcedura(IndexProcedura)
                    Picture4_Click 0
                End If
            Else
                CelarForm
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
    blTable.Visible = False
    
End Sub

Private Sub Text1_Change(Index As Integer)
Dim rc As Boolean

rc = IIf(Len(Text1(0)) > 0, True, False)
rc = IIf(Len(Text1(1)) > 0, rc, False)
DisableImage.Visible = Not (rc)
Picture4(0).Visible = rc

Image1.Visible = rc


Select Case Index
    Case 0
        MyLot = Text1(0)
    Case 1
        MyCode = Text1(1)
    Case 2
         Text1(1) = Text1(2)
End Select
End Sub



Private Sub StartProcedure()
Call SetPicForm
Call CelarForm
Call CleanformLot(False)
Call CleanformCode(False)
Call SetGrid(GrdBatch, GrdCode)
With CmbVisual
    .Clear
    .AddItem "Day"
    .AddItem "Month"
    .AddItem "Year"
    .AddItem "Archive"

End With
End Sub


