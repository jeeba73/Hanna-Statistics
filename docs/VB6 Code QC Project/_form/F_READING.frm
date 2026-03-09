VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form F_READING 
   BackColor       =   &H00808080&
   Caption         =   "Reading"
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
   ForeColor       =   &H00404040&
   Icon            =   "F_READING.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12000
   ScaleMode       =   0  'User
   ScaleWidth      =   19200
   Begin VB.Frame frLoading 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   5655
      Left            =   1920
      TabIndex        =   247
      Top             =   3960
      Visible         =   0   'False
      Width           =   13575
      Begin VB.Shape Shape3 
         BackColor       =   &H00964901&
         BorderColor     =   &H00964901&
         Height          =   5655
         Left            =   0
         Top             =   0
         Width           =   13575
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait"
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
         Height          =   975
         Left            =   0
         TabIndex        =   249
         Top             =   2760
         Width           =   13575
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
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
         Height          =   975
         Left            =   0
         TabIndex        =   248
         Top             =   1800
         Width           =   13575
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   2
      Left            =   0
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   228
      Top             =   2880
      Width           =   19215
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
         Height          =   495
         Index           =   0
         Left            =   12120
         TabIndex        =   236
         Top             =   7080
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
            Index           =   0
            Left            =   0
            TabIndex        =   237
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.ComboBox Combo4 
         Height          =   510
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   230
         Top             =   7080
         Width           =   2055
      End
      Begin VB.TextBox txSearch 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5520
         TabIndex        =   229
         Text            =   "Search"
         Top             =   7080
         Width           =   2175
      End
      Begin FlexCell.Grid Grd2 
         Height          =   6600
         Left            =   480
         TabIndex        =   231
         Top             =   360
         Width           =   18240
         _ExtentX        =   32173
         _ExtentY        =   11642
         AllowUserReorderColumn=   -1  'True
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColor1      =   14737632
         BackColor2      =   15790320
         BackColorActiveCellSel=   12632256
         BackColorBkg    =   14737632
         BackColorFixed  =   14737632
         BackColorFixedSel=   14737632
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
      Begin VB.Label lbReload 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clear Filter"
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
         Left            =   16365
         MouseIcon       =   "F_READING.frx":33E2
         MousePointer    =   99  'Custom
         TabIndex        =   253
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lbEditTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edit Table"
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
         Left            =   17655
         MouseIcon       =   "F_READING.frx":36EC
         MousePointer    =   99  'Custom
         TabIndex        =   244
         Top             =   0
         Width           =   1020
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Less"
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
         Left            =   16680
         MouseIcon       =   "F_READING.frx":39F6
         MousePointer    =   99  'Custom
         TabIndex        =   235
         Top             =   7200
         Width           =   1995
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   15720
         MouseIcon       =   "F_READING.frx":3D00
         MousePointer    =   99  'Custom
         TabIndex        =   234
         Top             =   7080
         Width           =   3015
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Production Value"
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
         Left            =   480
         MouseIcon       =   "F_READING.frx":400A
         MousePointer    =   99  'Custom
         TabIndex        =   233
         Top             =   7200
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   360
         MouseIcon       =   "F_READING.frx":4314
         MousePointer    =   99  'Custom
         TabIndex        =   232
         Top             =   6960
         Width           =   2655
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   0
         Left            =   8400
         MouseIcon       =   "F_READING.frx":461E
         MousePointer    =   99  'Custom
         Picture         =   "F_READING.frx":4928
         Top             =   7080
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   1
         Left            =   9360
         MouseIcon       =   "F_READING.frx":7D0A
         MousePointer    =   99  'Custom
         Picture         =   "F_READING.frx":8014
         Top             =   7080
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   2
         Left            =   10440
         MouseIcon       =   "F_READING.frx":B3F6
         MousePointer    =   99  'Custom
         Picture         =   "F_READING.frx":B700
         Top             =   7080
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00606060&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   2760
      TabIndex        =   40
      Top             =   1080
      Width           =   16455
      Begin VB.Frame Frame7 
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1575
         Left            =   120
         TabIndex        =   217
         Top             =   120
         Visible         =   0   'False
         Width           =   5655
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   219
            Top             =   360
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.ComboBox cmbTestType 
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   218
            Top             =   360
            Visible         =   0   'False
            Width           =   4215
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   1
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   0
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   840
         Width           =   2415
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1455
         Left            =   7680
         TabIndex        =   82
         Top             =   240
         Visible         =   0   'False
         Width           =   8655
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   36
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   86
            Top             =   840
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   35
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   85
            Top             =   840
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   34
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   84
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   33
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   83
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000080DF&
            Caption         =   "MAX"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   36
            Left            =   6000
            TabIndex        =   92
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000080DF&
            Caption         =   "MIN"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   35
            Left            =   4080
            TabIndex        =   91
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000080DF&
            Caption         =   "MAX"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   34
            Left            =   1920
            TabIndex        =   90
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000080DF&
            Caption         =   "MIN"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   33
            Left            =   0
            TabIndex        =   89
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lbTestTable 
            Alignment       =   2  'Center
            BackColor       =   &H000080DF&
            Caption         =   "STD Target Value "
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   35
            Left            =   0
            TabIndex        =   88
            Top             =   80
            Width           =   3735
         End
         Begin VB.Label lbTestTable 
            Alignment       =   2  'Center
            BackColor       =   &H000080DF&
            Caption         =   "pH"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   34
            Left            =   4080
            TabIndex        =   87
            Top             =   80
            Visible         =   0   'False
            Width           =   3735
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   4
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   3
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   2
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   840
         Width           =   1815
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   1455
         Left            =   10200
         TabIndex        =   207
         Top             =   320
         Visible         =   0   'False
         Width           =   6135
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   12
            Left            =   0
            TabIndex        =   210
            Top             =   760
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   14
            Left            =   3840
            TabIndex        =   209
            Top             =   760
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   13
            Left            =   1920
            TabIndex        =   208
            Top             =   760
            Width           =   1815
         End
         Begin VB.Label lbCommand 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "pH3"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   78
            Left            =   3840
            MouseIcon       =   "F_READING.frx":EAE2
            MousePointer    =   99  'Custom
            TabIndex        =   216
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00A65911&
            Caption         =   "pH Value"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   12
            Left            =   0
            TabIndex        =   215
            Top             =   400
            Width           =   1815
         End
         Begin VB.Label lbCommand 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "pH2"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   15
            Left            =   1920
            MouseIcon       =   "F_READING.frx":EDEC
            MousePointer    =   99  'Custom
            TabIndex        =   214
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label lbCommand 
            Alignment       =   2  'Center
            BackColor       =   &H000080DF&
            Caption         =   "pH1"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   14
            Left            =   0
            MouseIcon       =   "F_READING.frx":F0F6
            MousePointer    =   99  'Custom
            TabIndex        =   213
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00A65911&
            Caption         =   "ph max"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   14
            Left            =   3840
            TabIndex        =   212
            Top             =   400
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00A65911&
            Caption         =   "pH min"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   13
            Left            =   1920
            TabIndex        =   211
            Top             =   400
            Width           =   1815
         End
      End
      Begin VB.Label lbSettingName 
         BackStyle       =   0  'Transparent
         Caption         =   "Label17"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   360
         TabIndex        =   243
         Top             =   1440
         Width           =   6975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000080DF&
         Caption         =   "Code SFG"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   1
         Left            =   3000
         TabIndex        =   45
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000080DF&
         Caption         =   "Lot Number"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   0
         Left            =   360
         TabIndex        =   43
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "STD Value"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   4
         Left            =   10440
         TabIndex        =   52
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "STD Number"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   3
         Left            =   8520
         TabIndex        =   50
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "# Test"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   2
         Left            =   6600
         TabIndex        =   48
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      FillColor       =   &H00004000&
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   3120
      ScaleHeight     =   735
      ScaleWidth      =   3255
      TabIndex        =   220
      Top             =   10980
      Width           =   3255
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reading finished"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   360
         TabIndex        =   221
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image ImageClose 
         Height          =   480
         Left            =   2400
         MouseIcon       =   "F_READING.frx":F400
         MousePointer    =   99  'Custom
         Picture         =   "F_READING.frx":F70A
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000080DF&
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   0
      Left            =   0
      MouseIcon       =   "F_READING.frx":120FC
      MousePointer    =   99  'Custom
      ScaleHeight     =   1815
      ScaleWidth      =   2775
      TabIndex        =   41
      Top             =   1080
      Width           =   2775
      Begin VB.Label LabStandard 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Standard"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   0
         MouseIcon       =   "F_READING.frx":12406
         MousePointer    =   99  'Custom
         TabIndex        =   42
         Top             =   1080
         Width           =   2685
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   0
         Left            =   1200
         Picture         =   "F_READING.frx":12710
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.PictureBox PicMenuBar 
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
      Index           =   0
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   19215
      TabIndex        =   32
      Top             =   0
      Width           =   19215
      Begin VB.Frame frFilter 
         BackColor       =   &H00303030&
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   1095
         Left            =   11280
         TabIndex        =   250
         Top             =   0
         Visible         =   0   'False
         Width           =   7935
         Begin VB.ComboBox cbColFilter 
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   251
            Top             =   480
            Width           =   7215
         End
         Begin VB.Label lbFlter 
            BackStyle       =   0  'Transparent
            Caption         =   "Filter"
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
            Height          =   345
            Left            =   480
            MouseIcon       =   "F_READING.frx":15AF2
            MousePointer    =   99  'Custom
            TabIndex        =   252
            Top             =   120
            Width           =   7170
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   5
         Left            =   7680
         MouseIcon       =   "F_READING.frx":15DFC
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   222
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Evaluation QC"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   5
            Left            =   0
            MouseIcon       =   "F_READING.frx":16106
            MousePointer    =   99  'Custom
            TabIndex        =   223
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   5
            Left            =   735
            MouseIcon       =   "F_READING.frx":16410
            MousePointer    =   99  'Custom
            Picture         =   "F_READING.frx":1671A
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   4
         Left            =   9360
         MouseIcon       =   "F_READING.frx":19AFC
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   188
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Graph QC"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   4
            Left            =   0
            MouseIcon       =   "F_READING.frx":19E06
            MousePointer    =   99  'Custom
            TabIndex        =   189
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   4
            Left            =   735
            MouseIcon       =   "F_READING.frx":1A110
            MousePointer    =   99  'Custom
            Picture         =   "F_READING.frx":1A41A
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   3
         Left            =   5760
         MouseIcon       =   "F_READING.frx":1D7FC
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   99
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   735
            MouseIcon       =   "F_READING.frx":1DB06
            MousePointer    =   99  'Custom
            Picture         =   "F_READING.frx":1DE10
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lot Information"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   3
            Left            =   0
            MouseIcon       =   "F_READING.frx":211F2
            MousePointer    =   99  'Custom
            TabIndex        =   100
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
         MouseIcon       =   "F_READING.frx":214FC
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   37
         Top             =   0
         Width           =   1935
         Begin VB.Timer TimerIntro 
            Enabled         =   0   'False
            Interval        =   300
            Left            =   0
            Top             =   0
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "F_READING.frx":21806
            MousePointer    =   99  'Custom
            Picture         =   "F_READING.frx":21B10
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Set Standard"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   60
            MouseIcon       =   "F_READING.frx":24EF2
            MousePointer    =   99  'Custom
            TabIndex        =   38
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
         MouseIcon       =   "F_READING.frx":251FC
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   35
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Reading"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   1
            Left            =   60
            MouseIcon       =   "F_READING.frx":25506
            MousePointer    =   99  'Custom
            TabIndex        =   36
            Top             =   720
            Width           =   1875
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   735
            MousePointer    =   99  'Custom
            Picture         =   "F_READING.frx":25810
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
         MouseIcon       =   "F_READING.frx":28BF2
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   33
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Readings Table "
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   2
            Left            =   0
            MouseIcon       =   "F_READING.frx":28EFC
            MousePointer    =   99  'Custom
            TabIndex        =   34
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   735
            MouseIcon       =   "F_READING.frx":29206
            MousePointer    =   99  'Custom
            Picture         =   "F_READING.frx":29510
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Label LaInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The information on this page cannot be Edited : goto Lot Information to change / add fileds..."
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   11385
         MouseIcon       =   "F_READING.frx":2C8F2
         MousePointer    =   99  'Custom
         TabIndex        =   187
         Top             =   795
         Width           =   7575
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reading QC"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   12360
         TabIndex        =   39
         Top             =   360
         Width           =   6420
      End
   End
   Begin ChemicalQC.ctlCalendar ctlCalendar1 
      Height          =   6960
      Left            =   10680
      TabIndex        =   30
      Top             =   4800
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
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   3
      Left            =   0
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   101
      Top             =   2880
      Visible         =   0   'False
      Width           =   19215
      Begin VB.PictureBox PicInformation 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   7920
         MouseIcon       =   "F_READING.frx":2CBFC
         MousePointer    =   99  'Custom
         ScaleHeight     =   975
         ScaleWidth      =   3375
         TabIndex        =   185
         Top             =   6240
         Width           =   3375
         Begin VB.Label Lab 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Goto Lot Information QC"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   630
            MouseIcon       =   "F_READING.frx":2CF06
            MousePointer    =   99  'Custom
            TabIndex        =   186
            Top             =   660
            Width           =   2040
         End
         Begin VB.Image Im 
            Height          =   480
            Left            =   1440
            MouseIcon       =   "F_READING.frx":2D210
            MousePointer    =   99  'Custom
            Picture         =   "F_READING.frx":2D51A
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   52
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   162
         Top             =   5160
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   51
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   161
         Top             =   5160
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   50
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   160
         Top             =   5160
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   49
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   159
         Top             =   5160
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   48
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   158
         Top             =   5160
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   47
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   157
         Top             =   4200
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   46
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   150
         Top             =   4200
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   45
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   125
         Top             =   4200
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   44
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   124
         Top             =   4200
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   43
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   123
         Top             =   4200
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   42
         Left            =   14520
         Locked          =   -1  'True
         TabIndex        =   122
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   41
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   121
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   40
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   120
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   39
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   119
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   38
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   118
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   37
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   117
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00A0A0A0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2175
         Left            =   2040
         TabIndex        =   102
         Top             =   3480
         Visible         =   0   'False
         Width           =   15495
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   57
            Left            =   11880
            Locked          =   -1  'True
            TabIndex        =   168
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   58
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   167
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   59
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   166
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   60
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   165
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   61
            Left            =   9120
            Locked          =   -1  'True
            TabIndex        =   164
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   62
            Left            =   11880
            Locked          =   -1  'True
            TabIndex        =   163
            Top             =   1560
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00606060&
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   56
            Left            =   9120
            Locked          =   -1  'True
            TabIndex        =   149
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00606060&
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   53
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   105
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00606060&
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   54
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   104
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00606060&
            BorderStyle     =   0  'None
            Height          =   435
            Index           =   55
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   103
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Expiration E"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   55
            Left            =   11880
            TabIndex        =   116
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Expiration D"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   56
            Left            =   9120
            TabIndex        =   115
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Expiration C"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   57
            Left            =   6360
            TabIndex        =   114
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Expiration B"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   58
            Left            =   3600
            TabIndex        =   113
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            Caption         =   "Expiration A"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   59
            Left            =   720
            TabIndex        =   112
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00964901&
            Caption         =   "Reagent E Lot"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   60
            Left            =   11880
            TabIndex        =   111
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00964901&
            Caption         =   "Reagent D Lot"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   96
            Left            =   9120
            TabIndex        =   110
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00964901&
            Caption         =   "Reagent C Lot"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   97
            Left            =   6360
            TabIndex        =   109
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00964901&
            Caption         =   "Reagent B Lot"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   98
            Left            =   3600
            TabIndex        =   108
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00964901&
            Caption         =   "Reagent A Lot"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   99
            Left            =   720
            TabIndex        =   107
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label9 
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
            TabIndex        =   106
            Top             =   1680
            Width           =   180
         End
      End
      Begin VB.Label lb 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         Caption         =   "Reagent "
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   43
         Left            =   1680
         TabIndex        =   129
         Top             =   3120
         Width           =   15735
      End
      Begin VB.Label lb 
         Alignment       =   2  'Center
         BackColor       =   &H00606060&
         Caption         =   "Reagent Range"
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   47
         Left            =   11400
         TabIndex        =   133
         Top             =   1200
         Width           =   6015
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00606060&
         Height          =   1575
         Index           =   3
         Left            =   11400
         Top             =   1320
         Width           =   6015
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00606060&
         Height          =   2655
         Index           =   1
         Left            =   1680
         Top             =   3240
         Width           =   15735
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   4
         Left            =   120
         TabIndex        =   148
         Top             =   7440
         Width           =   2655
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operator"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   360
         TabIndex        =   147
         Top             =   9120
         Width           =   975
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   9360
         Picture         =   "F_READING.frx":308FC
         Top             =   8880
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Recipe Number"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   37
         Left            =   2040
         TabIndex        =   146
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Ref. Weight"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   38
         Left            =   4920
         TabIndex        =   145
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Min (mg)"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   39
         Left            =   6840
         TabIndex        =   144
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Max (mg)"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   40
         Left            =   8760
         TabIndex        =   143
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Min"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   41
         Left            =   11760
         TabIndex        =   142
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Max"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   42
         Left            =   14520
         TabIndex        =   141
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Reagent A Lot"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   43
         Left            =   2520
         TabIndex        =   140
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Reagent B Lot"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   70
         Left            =   5400
         TabIndex        =   139
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Reagent C Lot"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   54
         Left            =   8160
         TabIndex        =   138
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Reagent D Lot"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   53
         Left            =   10920
         TabIndex        =   137
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Reagent E Lot"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   52
         Left            =   13680
         TabIndex        =   136
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Expiration A"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   49
         Left            =   2520
         TabIndex        =   135
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Expiration B"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   48
         Left            =   5400
         TabIndex        =   134
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Expiration C"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   46
         Left            =   8160
         TabIndex        =   132
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Expiration D"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   45
         Left            =   10920
         TabIndex        =   131
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Expiration E"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   44
         Left            =   13680
         TabIndex        =   130
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Label lbCommand 
         Alignment       =   2  'Center
         BackColor       =   &H000080DF&
         Caption         =   "REAGENT SET 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   50
         Left            =   1680
         MouseIcon       =   "F_READING.frx":33CDE
         MousePointer    =   99  'Custom
         TabIndex        =   128
         Top             =   2640
         Width           =   3495
      End
      Begin VB.Label lbCommand 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "REAGENT SET 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   51
         Left            =   5280
         MouseIcon       =   "F_READING.frx":33FE8
         MousePointer    =   99  'Custom
         TabIndex        =   127
         Top             =   2640
         Width           =   3495
      End
      Begin VB.Label Label11 
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
         TabIndex        =   126
         Top             =   5280
         Width           =   180
      End
   End
   Begin VB.TextBox txQRCode 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   240
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   14535
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   0
      Left            =   0
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   28
      Top             =   2880
      Visible         =   0   'False
      Width           =   19215
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   67
         Left            =   11880
         TabIndex        =   241
         Top             =   2640
         Width           =   5775
      End
      Begin VB.PictureBox ReadSTDQRCode 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         FillColor       =   &H00004000&
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   11880
         ScaleHeight     =   735
         ScaleWidth      =   5775
         TabIndex        =   238
         Top             =   1200
         Width           =   5775
         Begin VB.Label lbReadSTDQRCode 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Read STD QRCode"
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   15
            TabIndex        =   239
            Top             =   195
            Width           =   5625
         End
      End
      Begin FlexCell.Grid Grd1 
         Height          =   2640
         Left            =   2520
         TabIndex        =   81
         Top             =   4440
         Visible         =   0   'False
         Width           =   15120
         _ExtentX        =   26670
         _ExtentY        =   4657
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
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   30
         Left            =   11880
         TabIndex        =   200
         Top             =   3720
         Width           =   5775
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   18
         Left            =   3120
         TabIndex        =   195
         Top             =   3720
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   20
         Left            =   5640
         TabIndex        =   194
         Top             =   3720
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Height          =   510
         Left            =   8880
         Style           =   2  'Dropdown List
         TabIndex        =   193
         Top             =   3720
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H000080DF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1575
         Left            =   3960
         TabIndex        =   76
         Top             =   4800
         Width           =   11295
         Begin VB.Shape Shape1 
            BackColor       =   &H000080DF&
            BorderColor     =   &H000060BF&
            Height          =   1575
            Index           =   0
            Left            =   0
            Top             =   0
            Width           =   11295
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3 - Enter Test Information  and Save Readings"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   2640
            TabIndex        =   79
            Top             =   960
            Width           =   4665
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2 - Check Reagent range in Lot Information"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   2640
            TabIndex        =   78
            Top             =   600
            Width           =   4395
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1 - Enter STD number or select form List : Click ""SFG Standard"""
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   2640
            TabIndex        =   77
            Top             =   240
            Width           =   6375
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   1440
            Picture         =   "F_READING.frx":342F2
            Top             =   480
            Width           =   480
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   11
         Left            =   8880
         TabIndex        =   6
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   10
         Left            =   6960
         TabIndex        =   5
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   9
         Left            =   5040
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   7
         Left            =   5640
         TabIndex        =   2
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   6
         Left            =   3120
         TabIndex        =   1
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   5
         Left            =   3120
         TabIndex        =   0
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   21
         Left            =   8880
         TabIndex        =   196
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   8
         Left            =   8880
         TabIndex        =   3
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "STD ID"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   67
         Left            =   11880
         TabIndex        =   242
         Top             =   2280
         Width           =   5775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "+/- % "
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   8
         Left            =   8880
         TabIndex        =   56
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "QC OPERATOR"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   30
         Left            =   11880
         TabIndex        =   201
         Top             =   3360
         Width           =   5775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "PROD. DATE"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   18
         Left            =   3120
         TabIndex        =   199
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "PROD. OPERATOR"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   20
         Left            =   5640
         TabIndex        =   198
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "HEAD"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   21
         Left            =   8880
         TabIndex        =   197
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Height          =   1455
         Left            =   1680
         MouseIcon       =   "F_READING.frx":376D4
         MousePointer    =   99  'Custom
         TabIndex        =   179
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lll 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clear"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   5
         Left            =   2160
         MouseIcon       =   "F_READING.frx":379DE
         MousePointer    =   99  'Custom
         TabIndex        =   184
         Top             =   1800
         Width           =   435
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   1
         Left            =   2160
         Picture         =   "F_READING.frx":37CE8
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   855
         Left            =   15480
         MouseIcon       =   "F_READING.frx":3B0CA
         MousePointer    =   99  'Custom
         TabIndex        =   61
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lbStandard 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SFG Standard Only"
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
         Height          =   525
         Left            =   2520
         MouseIcon       =   "F_READING.frx":3B3D4
         MousePointer    =   99  'Custom
         TabIndex        =   152
         Top             =   7200
         Visible         =   0   'False
         Width           =   3720
      End
      Begin VB.Label lbTarget 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Standard"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   9
         Left            =   3120
         TabIndex        =   80
         Top             =   600
         Width           =   7575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "SFG Standard"
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
         Left            =   16440
         MouseIcon       =   "F_READING.frx":3B6DE
         MousePointer    =   99  'Custom
         TabIndex        =   60
         Top             =   480
         Width           =   1815
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   3
         Left            =   15840
         Picture         =   "F_READING.frx":3B9E8
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "MAX"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   11
         Left            =   8880
         TabIndex        =   59
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "MIN"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   10
         Left            =   6960
         TabIndex        =   58
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "VALUE"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   9
         Left            =   5040
         TabIndex        =   57
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Fixed"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   7
         Left            =   5640
         TabIndex        =   55
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "MR STD"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   6
         Left            =   3120
         TabIndex        =   54
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "STD Number"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   5
         Left            =   3120
         TabIndex        =   53
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Left            =   9360
         Picture         =   "F_READING.frx":3EDCA
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
         TabIndex        =   29
         Top             =   9120
         Width           =   975
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   8
         Left            =   120
         TabIndex        =   31
         Top             =   7440
         Width           =   2655
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   1
      Left            =   0
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   46
      Top             =   2880
      Width           =   19215
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
         Height          =   495
         Index           =   1
         Left            =   3360
         TabIndex        =   254
         Top             =   6360
         Width           =   3015
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Set Readings Number"
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
            TabIndex        =   255
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   68
         Left            =   9600
         TabIndex        =   245
         Top             =   5760
         Width           =   6735
      End
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   1920
         Top             =   5160
      End
      Begin VB.ComboBox Combo2 
         Height          =   510
         Left            =   12000
         Style           =   2  'Dropdown List
         TabIndex        =   206
         Top             =   720
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   66
         Left            =   14400
         TabIndex        =   203
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   65
         Left            =   12000
         TabIndex        =   202
         Top             =   720
         Width           =   2295
      End
      Begin FlexCell.Grid GrdTestType 
         Height          =   3600
         Left            =   2400
         TabIndex        =   177
         Top             =   1200
         Visible         =   0   'False
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   6350
         AllowUserReorderColumn=   -1  'True
         AllowUserSort   =   -1  'True
         Appearance      =   0
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
         Height          =   435
         Index           =   64
         Left            =   9600
         TabIndex        =   17
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   63
         Left            =   7560
         TabIndex        =   16
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3495
         Left            =   16680
         TabIndex        =   93
         Top             =   1440
         Width           =   16815
         Begin VB.Image Image5 
            Height          =   480
            Left            =   4440
            Picture         =   "F_READING.frx":421AC
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Please Enter Valid Standard ( STD Value and pH )"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   5
            Left            =   5280
            TabIndex        =   94
            Top             =   1500
            Width           =   5700
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H000080DF&
            BorderColor     =   &H00303030&
            Height          =   3495
            Left            =   0
            Top             =   0
            Width           =   16815
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   25
         Left            =   12240
         TabIndex        =   13
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   24
         Left            =   9600
         TabIndex        =   12
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   23
         Left            =   6960
         TabIndex        =   11
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   22
         Left            =   4320
         TabIndex        =   10
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   19
         Left            =   9600
         TabIndex        =   9
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   435
         Index           =   32
         Left            =   16320
         TabIndex        =   21
         Top             =   7080
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   31
         Left            =   3360
         TabIndex        =   20
         Top             =   5760
         Width           =   6135
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   29
         Left            =   14040
         TabIndex        =   19
         Top             =   4440
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   28
         Left            =   11640
         TabIndex        =   18
         Top             =   4440
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   27
         Left            =   5520
         TabIndex        =   15
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   26
         Left            =   3360
         TabIndex        =   14
         Top             =   4440
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   17
         Left            =   7200
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   16
         Left            =   4800
         TabIndex        =   7
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   15
         Left            =   2400
         TabIndex        =   63
         Top             =   720
         Width           =   2295
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         FillColor       =   &H00004000&
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   7600
         ScaleHeight     =   855
         ScaleWidth      =   3975
         TabIndex        =   62
         Top             =   6600
         Visible         =   0   'False
         Width           =   3975
         Begin VB.Image ImageTAV 
            Height          =   480
            Index           =   5
            Left            =   1720
            MouseIcon       =   "F_READING.frx":4558E
            MousePointer    =   99  'Custom
            Picture         =   "F_READING.frx":45898
            Top             =   160
            Width           =   480
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "CORRECTION"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   47
         Left            =   9600
         TabIndex        =   246
         Top             =   5400
         Width           =   6735
      End
      Begin VB.Label lbFamily 
         Caption         =   "Label17"
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   227
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbFamily 
         Caption         =   "Label17"
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   226
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbFamily 
         Caption         =   "Label17"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   225
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbFamily 
         Caption         =   "Label17"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   224
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "OTHER Code SFG"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   65
         Left            =   12000
         TabIndex        =   205
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "LOT"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   66
         Left            =   14400
         TabIndex        =   204
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "pH 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   64
         Left            =   9600
         TabIndex        =   183
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label lbColor 
         Caption         =   "Label12"
         Height          =   255
         Index           =   64
         Left            =   9360
         TabIndex        =   182
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "pH 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   63
         Left            =   7560
         TabIndex        =   181
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label lbColor 
         Caption         =   "Label12"
         Height          =   255
         Index           =   63
         Left            =   7320
         TabIndex        =   180
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbColor 
         Caption         =   "Label12"
         Height          =   255
         Index           =   29
         Left            =   13800
         TabIndex        =   176
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbColor 
         Caption         =   "Label12"
         Height          =   255
         Index           =   28
         Left            =   11400
         TabIndex        =   175
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbColor 
         Caption         =   "Label12"
         Height          =   255
         Index           =   27
         Left            =   5280
         TabIndex        =   174
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbColor 
         Caption         =   "Label12"
         Height          =   255
         Index           =   26
         Left            =   3120
         TabIndex        =   173
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbColor 
         Caption         =   "Label12"
         Height          =   255
         Index           =   25
         Left            =   12000
         TabIndex        =   172
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbColor 
         Caption         =   "Label12"
         Height          =   255
         Index           =   24
         Left            =   9600
         TabIndex        =   171
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbColor 
         Caption         =   "Label12"
         Height          =   255
         Index           =   23
         Left            =   7200
         TabIndex        =   170
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbColor 
         Caption         =   "Label12"
         Height          =   255
         Index           =   22
         Left            =   4800
         TabIndex        =   169
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Code"
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   87
         Left            =   12240
         TabIndex        =   156
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Code2"
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   86
         Left            =   9600
         TabIndex        =   155
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Code3"
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   85
         Left            =   6960
         TabIndex        =   154
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Code4"
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   84
         Left            =   4320
         TabIndex        =   153
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "METER 4"
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   25
         Left            =   12240
         TabIndex        =   151
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1170
         Left            =   480
         TabIndex        =   98
         Top             =   0
         Width           =   1440
      End
      Begin VB.Label lbCommand 
         Alignment       =   2  'Center
         BackColor       =   &H000080DF&
         Caption         =   "REAGENT SET 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   42
         Left            =   6960
         MouseIcon       =   "F_READING.frx":48C7A
         MousePointer    =   99  'Custom
         TabIndex        =   97
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lbCommand 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "REAGENT SET 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   41
         Left            =   9600
         MouseIcon       =   "F_READING.frx":48F84
         MousePointer    =   99  'Custom
         TabIndex        =   96
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "PROD. TIME"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   19
         Left            =   9600
         TabIndex        =   95
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "QC TYPE"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   32
         Left            =   16320
         TabIndex        =   75
         Top             =   6720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "QC Note"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   31
         Left            =   3360
         TabIndex        =   74
         Top             =   5400
         Width           =   6135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "WEIGHT (mg)"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   29
         Left            =   14040
         TabIndex        =   73
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "TURBID."
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   28
         Left            =   11640
         TabIndex        =   72
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "pH 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   27
         Left            =   5520
         TabIndex        =   71
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "SPECTR. (ABS)"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   26
         Left            =   3360
         TabIndex        =   70
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "METER 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   24
         Left            =   9600
         TabIndex        =   69
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "METER 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   23
         Left            =   6960
         TabIndex        =   68
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "METER 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   22
         Left            =   4320
         TabIndex        =   67
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "QC TIME"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   17
         Left            =   7200
         TabIndex        =   66
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "QC DATE"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   16
         Left            =   4800
         TabIndex        =   65
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "TEST TYPE"
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Index           =   15
         Left            =   2400
         TabIndex        =   64
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   0
      Left            =   17640
      MouseIcon       =   "F_READING.frx":4928E
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   10680
      Width           =   1575
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   2
      Left            =   15240
      MouseIcon       =   "F_READING.frx":49598
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   10680
      Width           =   2175
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Index           =   1
      Left            =   8280
      MouseIcon       =   "F_READING.frx":498A2
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   10560
      Width           =   2655
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   3
      Left            =   0
      MouseIcon       =   "F_READING.frx":49BAC
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   10560
      Width           =   1815
   End
   Begin VB.Label La 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move Forward"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   17655
      MouseIcon       =   "F_READING.frx":49EB6
      MousePointer    =   99  'Custom
      TabIndex        =   192
      Top             =   11600
      Width           =   1200
   End
   Begin VB.Label La 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move Previous"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   15630
      MouseIcon       =   "F_READING.frx":4A1C0
      MousePointer    =   99  'Custom
      TabIndex        =   191
      Top             =   11600
      Width           =   1230
   End
   Begin VB.Label La 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Procedure"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   9000
      MouseIcon       =   "F_READING.frx":4A4CA
      MousePointer    =   99  'Custom
      TabIndex        =   190
      Top             =   11600
      Width           =   1200
   End
   Begin VB.Label lbOperator 
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   178
      Top             =   11600
      Width           =   2775
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   15960
      MouseIcon       =   "F_READING.frx":4A7D4
      MousePointer    =   99  'Custom
      Picture         =   "F_READING.frx":4AADE
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_READING.frx":4DEC0
      Height          =   480
      Index           =   1
      Left            =   9360
      MouseIcon       =   "F_READING.frx":512A2
      MousePointer    =   99  'Custom
      Picture         =   "F_READING.frx":515AC
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   360
      MouseIcon       =   "F_READING.frx":5498E
      MousePointer    =   99  'Custom
      Picture         =   "F_READING.frx":54C98
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
      TabIndex        =   23
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
      DragIcon        =   "F_READING.frx":5807A
      Height          =   480
      Index           =   0
      Left            =   18240
      MouseIcon       =   "F_READING.frx":5B45C
      MousePointer    =   99  'Custom
      Picture         =   "F_READING.frx":5B766
      Top             =   11040
      Width           =   480
   End
End
Attribute VB_Name = "F_READING"
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
Private IndexpH As Integer
Private IndexReagentSet As Integer
Private MeterNumber As Integer
Private pHNumber As Integer
Private MyTestType As String
Private MyQCType As String
Private ph(3, 3) As String
Private MeasurementUnit As String
Private UserDecimal As String
Private STD() As String
Private STDCount As Integer
Private ReadingCodeID As Long
Private STDSelectedNumber As String
Private bAnotherFormCalled As Boolean
Private ExsistReagentSet As Boolean
Private lRow As Long
Private lCol As Long
Private bEnterTable As Boolean
Private bFinishedReadings As Boolean

Private UserQrCode As QRCodeType

Private bClosedLot As Boolean
Private strValidDate As String
Private fCol As Long
Private fRow As Long


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
            'Ctl.DefaultFont.Size = 12 * m_ControlGridFontSize
            Ctl.DefaultRowHeight = 30 * m_ControlGridRowHeight
            'For i = 0 To Ctl.Cols - 1
               ' Ctl.Column(i).Width = Ctl.Column(i).Width * m_ControlGridColWidth
           ' Next
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
            Ctl.BackColor = vbColorLabelUnabled
            End If

        ElseIf TypeOf Ctl Is Grid Then
            Ctl.Rows = 1
        End If
    Next Ctl
End Sub


Public Function DoShow(ByRef Index As Integer, Optional ByRef sLot As String, Optional ByRef sCode As String, Optional ByVal lngID As Long, Optional MyImage As Image, Optional FileName As String) As Boolean

    On Error GoTo ERR_SHOW
    
    CloseSettingDataFile
    
   ' Set DefaultMenu(4) = MyImage
    IndexMainProcedura = Index
                
    Text1(16) = FormatDataLAT(date)
    Text1(17) = Replace(FormatDateTime(Now, vbShortTime), ".", ":")
    Text1(0) = sLot
    Text1(1) = sCode
    MyCode = sCode
    
    SettingName = FileName
    lbSettingName = SettingName
    lbOperator = MyOperatore.Name
    ReadingCodeID = lngID
    
    
    TimerIntro.Enabled = True
    
   
    
    
    m_rc = False
    bFormSaved = False
 
    

    mOk

    Me.Show vbModal
    
    If m_rc = True Then
        CloseSettingDataFile
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
    MsgBox Err.Description
    Resume ERR_END
End Function




Private Sub cbColFilter_Click()
Call CheckInTable(cbColFilter, fCol)
End Sub

Private Sub cmbTestType_Click()
If InStr(LCase(cmbTestType), "select") Then
Else
    Grd2.Cell(lRow, 4).Text = cmbTestType
    
End If
End Sub

Private Sub Combo2_Click()
Text1(65).Locked = False
If Combo2 = "Enter" Then
    Text1(65).SetFocus
    
Else
    Text1(65) = Combo2
    Text1(66).SetFocus
End If
Combo2.Visible = False
End Sub

Private Sub Combo1_Click()
If Combo1 = "Enter" Then
    Text1(21).SetFocus
Else
    Text1(21) = Combo1
End If

Combo1.Visible = False
End Sub

Private Sub Combo3_Click()
If InStr(LCase(Combo3), "select") Then
ElseIf InStr(LCase(Combo3), "enter") Then
    Combo3.Visible = False
    Frame7.Visible = False
    Combo3.ListIndex = 0
    Grd2.Cell(lRow, 34).Text = ""
    Grd2.Cell(lRow, 34).SetFocus
    
Else

    Grd2.Cell(lRow, 34).Text = Combo3
End If

End Sub

Private Function SearchInTab()

Dim r As Integer
Dim i As Integer
Dim t As Integer


    Select Case Combo4
        Case "STD"
        
            r = 1
        Case "Test"
        
            r = 4
    
    
    End Select
    
 With Grd2
    .ReadOnly = True
    .AutoRedraw = False
    If .Rows > 1 Then
        For t = 1 To .Rows - 1
            If InStr(txSearch, "Search") Or txSearch = "" Then
                .RowHeight(t) = 30
            Else
            
                If InStr(UCase(.Cell(t, r).Text), UCase(txSearch)) Then
                    .RowHeight(t) = 30
                Else
                    .RowHeight(t) = 0
                End If
            End If
            
        Next
        
    End If
    .Refresh
    .AutoRedraw = True
 End With
 
 
End Function

Private Sub Combo4_Click()

If Me.Visible Then
    txSearch = ""
    txSearch.SetFocus
End If

End Sub

Private Sub ctlCalendar1_DateClicked(inputDate As Date)
Text1(IndexTextSelected) = FormatDataLAT(CStr(inputDate))
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
            SalvaFormSettingName
          
            
            DoEvents
            DoEvents
            DoEvents
            
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
        Text1(30) = MyOperatore.Name
        
        Call CheckClosedLot
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
'IndexFormProcedura = GetSetting(App.Title, "Information QC", "IndexFormProcedura", 0)
frLoading.ZOrder
frLoading.Move Me.Width / 2 - frLoading.Width / 2, Me.Height / 2 - frLoading.Height / 2

frLoading.Visible = True
If IndexFormProcedura = 4 Then IndexFormProcedura = 2
lbSettingName = ""
FormPulisciTutto

CloseSettingDataFile


Call SetPicForm
Call SetGridStandardTolerance(Grd1)
Call SetAllGrid


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
Dim i As Integer
'If Screen.Width - Me.Width > 1000 And bFullScreen Then
    'Me.WindowState = 2
    For i = 0 To PicMain.Count - 1
        PicMain(i).Picture = LoadPicture(PictureMaxScreen)
       
    Next '
     'Me.Picture = LoadPicture(PictureMaxScreen)
'End If


SelectProcedura IndexFormProcedura

End Sub

Private Sub Form_Resize()


ResizeControls
End Sub

Private Sub Frame2_Click()
Frame2.Visible = False
Label3_Click
End Sub

Private Sub Frame4_Click()
SelectProcedura (0)
End Sub

Private Sub frCommandInside_Click(Index As Integer)
Dim iTest As String
MyLot = Trim(Text1(0))
    Select Case Index
        Case 0
            If F_MsgBox.DoShow("Print QC Label to Preparation?", MyCode) Then
                Call PrintLabelReadings(MyLot, MyCode)
            End If
            
        Case 1
            If F_MsgBox.DoShow("Readings number = " & Grd2.Rows - 1 & vbCrLf & "Do you want to set readings number manually (e.g. database error)?", "READINGS NUMBER") Then
                iTest = Grd2.Rows - 1
                If F_InputBox.DoShow("Enter readings number...", "READINGS NUMBER", , , , iTest, , True) Then
                    iTest = Int(iTest) + 1
                    CloseSettingDataFile
                    SaveSettingData SettingName, "Reading QC", "Grd2 Rows", iTest
                    CloseSettingDataFile
                    Call GetFormSettingName(True)
                    Call PutLotInDatabase
                End If
            End If
            
    End Select
    
End Sub



Private Sub Grd1_DblClick()
SelectProcedura (1)
Text1(15).SetFocus
Text1_GotFocus 15
End Sub

Private Sub Grd1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim i As Integer
i = FirstCol
If i > 0 Then
    With Grd1
        Select Case i
            Case 6, 9, 12, 15, 18, 21 ', 38, 41, 44
            ' ho selezionato uno standard
            Text1(5) = Int((i - 6) / 3) + 1
            Text1(9) = IIf(Trim(.Cell(2, i).Text) <> "/", Trim(.Cell(2, i).Text), "")
            
            Text1_LostFocus 9
            
            
            Text1(10) = IIf(Trim(.Cell(2, i + 1).Text) <> "/", Trim(.Cell(2, i + 1).Text), "")
            Text1(11) = IIf(Trim(.Cell(2, i + 1).Text) <> "/", Trim(.Cell(2, i + 2).Text), "")
            
            Text1(33) = Text1(10)
            Text1(34) = Text1(11)
            Text1(10).Locked = True
            Text1(11).Locked = True

        End Select
    End With
End If
                    
End Sub





Private Sub Grd2_Click()
If bClosedLot Then
Else
    If Grd2.ReadOnly = False Then
        InsertDataInGrid
    End If
End If

End Sub

Private Sub Grd2_DblClick()

     frFilter.Visible = True
    lbFlter = Grd2.Cell(0, lCol).Text
     Call fillcbColFilter(lCol, "Select " & Grd2.Cell(0, lCol).Text)
     
     blTable.Visible = False
     fCol = lCol

End Sub


Private Sub Grd2_LostFocus()

'ReadOnlyGrid
End Sub

Private Sub Grd2_OwnerDrawCell(ByVal Row As Long, ByVal Col As Long, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Handled As Boolean)

If bFinishedReadings = False Then
    Picture2.Visible = IIf(Grd2.Rows > 1, True, False)
End If

End Sub

Private Sub Grd2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim i As Integer
Dim a As Integer
'STDSelectedNumber = ""
a = FirstRow

cmbTestType.Visible = False
Combo3.Visible = False
Frame7.Visible = False
frCommandInside(0).Visible = False

i = 0




If a > 0 Then
    'frFilter.visible = False
    lRow = FirstRow
    lCol = FirstCol

    With Grd2
    
        If .ReadOnly = False Then
            
            ' inserisco i dati
            If bEnterTable Then InsertDataInGrid
        Else
           
        End If
        Debug.Print .Cell(lRow, 4).Text
        frCommandInside(0).Visible = IIf(InStr(LCase(.Cell(lRow, 4).Text), "valid") Or InStr(LCase(.Cell(lRow, 4).Text), "final"), True, False)
        
        ShowStandardMinMaxValue (a)
          
  
    
    
         
    
    'Select Case lCol
    '    Case 20 ' weight
    '        frFilter.visible = True
    '        Call fillcbColFilter(lCol, "Select Weight")
    '        blTable.Visible = False
    '        fCol = lCol
    '    Case 23  ' Correction
    '        frFilter.visible = True
    '        Call fillcbColFilter(lCol, "Select Correction")
    '        blTable.Visible = False
    '        fCol = lCol
    '    Case 37  ' Note
    '        frFilter.visible = True
    '        Call fillcbColFilter(lCol, "Select Note")
    '        blTable.Visible = False
    '        fCol = lCol
    'End Select
    
      End With
Else

    lRow = 0
    lCol = 0
    
            
            
    lbTestTable(35) = "STD Target Value"
    lbTestTable(34) = "pH"
    
    Text1(33) = ""
    Text1(34) = ""
    Text1(35) = ""
    Text1(36) = ""
End If

End Sub

Private Function fillcbColFilter(ByVal Col As Long, ByVal filter As String)

Dim strColumn() As String
Dim Column As String
Dim i As Integer
Dim t As Integer
t = 0
 With Grd2
    cbColFilter.Clear
    If .Rows > 1 Then
       
          
        cbColFilter.AddItem filter
        For i = 1 To .Rows - 1
            If Trim(UCase(Trim(.Cell(i, Col).Text))) = "" Then GoTo cont
            Column = Trim(UCase(Trim(.Cell(i, Col).Text)))
            If t > 0 Then
                If GetIndexArStrSingle(strColumn(), Column) = -1 Then
Aggiungi:
                        
                    t = t + 1
                    ReDim Preserve strColumn(t)
                    strColumn(t) = Column
                   
                    
                    
                
                End If
            Else
                GoTo Aggiungi
            End If
            
                        
cont:
        Next
    
    
    End If
    
 End With
 
 If t > 0 Then
    Call OrdinaArray(strColumn)
    For i = 1 To UBound(strColumn)
        cbColFilter.AddItem strColumn(i)
    Next
    cbColFilter.ListIndex = 0
 End If
 
End Function



Private Function CheckInTable(ByVal strCombo As String, ByVal Col As Long)
Dim bTutti As Boolean
Dim i As Integer



ReadOnlyGrid
If InStr(LCase(strCombo), "select") Or strCombo = "" Then bTutti = True


With Grd2
    
    .AutoRedraw = False

    For i = 1 To .Rows - 1
        
        If InStr(UCase(.Cell(i, Col).Text), UCase(strCombo)) Or bTutti Then
        
            .RowHeight(i) = 35
        
        Else
        
            .RowHeight(i) = 0
        End If
    
    
    
    Next


    .Refresh
    .AutoRedraw = True

End With

End Function




Private Function ShowStandardMinMaxValue(ByVal a As Long)
    With Grd2
        
        STDSelectedNumber = Trim(.Cell(a, 1).Text)
        
        If STDSelectedNumber <> "" Then

            If IsNull(Trim(.Cell(a, 27).Text)) Or Trim(.Cell(a, 27).Text) = "" Then
            
                CheckSTD
            
            End If
            
            Text1(33).Text = Trim(.Cell(a, 27).Text)
            Text1(34).Text = Trim(.Cell(a, 28).Text)
            Text1(5) = STDSelectedNumber
                    
            Text1(9) = .Cell(a, 2).Text
            Text1(10) = .Cell(a, 27).Text
            Text1(11) = .Cell(a, 28).Text
                    
            
            'Text1_LostFocus 5
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            
            lbTestTable(35) = "STD" & .Cell(a, 1).Text & " Target Value"
        End If
        
     End With
End Function
Private Sub GrdTestType_DblClick()

If MyTestType <> "" Then
    
    Text1(15) = MyTestType
    
    MyTestType = ""
    GrdTestType.Visible = False
    
    Text1(16).SetFocus
End If
GrdTestType.Cell(0, 0).SetFocus
End Sub

Private Sub GrdTestType_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    If F_MsgBox.DoShow("Delete Test Type : " & MyTestType & " ?") Then
        With dbTabTestType
            .filter = ""
            .filter = "Type='" & Trim(MyTestType) & "'"
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

Private Sub GrdTestType_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

If FirstRow > 0 Then

    MyTestType = Trim(GrdTestType.Cell(FirstRow, 1).Text)

Else

    GrdTestType.Cell(0, 0).SetFocus
    
End If

End Sub




Private Sub Im_Click()
PicInformation_Click
End Sub

Private Sub Image2_Click(Index As Integer)
    Select Case Index
        Case 0
            Call AddLine
        Case 1
            Call DeleteLine
        Case 2
            Call SaveLine
            ReadOnlyGrid
            frFilter.Visible = False
        
    End Select
    
End Sub

Private Sub ImageClose_Click()
Picture2_Click
End Sub

Private Sub ImageTAV_Click(Index As Integer)
Select Case Index
    Case 5
        Picture1_Click
End Select
End Sub

Private Sub Lab_Click()
PicInformation_Click
End Sub

Private Sub Label10_Click()
Dim rc As Boolean
frFilter.Visible = False
ReadOnlyGrid

If Label7.Caption = "Show More" Then
    Label7.Caption = "Show Less"
    rc = True
Else
    Label7.Caption = "Show More"
    rc = False

End If
Call SetViewGrid2(rc)
SaveSetting App.Title, "Reading QC", "Visualizza Colonne", (rc)
End Sub

Private Sub Label12_Click()
    Dim sString As String
    Dim vbColor As OLE_COLOR
    Dim i As Integer
    Dim r As Integer
    If Label12 = "Show Production Value" Then
        Label12 = "Show All Value"
        vbColor = vbColorTextLightBlue
setColor:
        With Grd2
            For i = 1 To .Rows - 1
                sString = .Cell(i, 4).Text
                
                If InStr(sString, "P") Then
                   For r = 1 To .Cols - 1
                        .Cell(i, r).BackColor = vbColor
                    Next
                Else
                   ' .RowHeight(i) = 0
                End If
             Next
        End With
    Else
        vbColor = vbColorLightFixed
        Label12 = "Show Production Value"
        GoTo setColor:
    End If
End Sub

Private Sub Label14_Click()
Label12_Click
End Sub

Private Sub LabelPH_Click(Index As Integer)

End Sub

Private Sub Label15_Click()
' CLEAR STD
Text1(5) = ""
CleanSTDText
Text1(5).SetFocus
End Sub

Private Sub CleanSTDText()
    
    Text1(9) = ""
    Text1(10) = ""
    Text1(11) = ""
    
    Text1(67) = ""
    
    Text1(33) = ""
    Text1(34) = ""
    
    Text1(18) = ""
    
    Text1(19) = ""
    Text1(20) = ""
    Text1(21) = ""
    Text1(30) = ""
    Text1(10).Locked = False
    Text1(11).Locked = False
    
    CancellaFormTestAfterReading
    
    Grd1.Cell(1, 1).SetFocus

    Text1(30) = MyOperatore.Name
End Sub

Private Sub Label16_Click()
Picture2_Click
End Sub

Private Sub Label7_Click()
Label10_Click
End Sub

Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub

Private Sub lbEditTable_Click()
  If Grd2.ReadOnly Then
        
        Call ActivateGrd2
        lbEditTable.ForeColor = vbColorOrange
  Else
    Grd2.ReadOnly = True
       lbEditTable.ForeColor = &H964901
        
  End If
    
     
End Sub

Private Sub lbReadSTDQRCode_Click()
ReadSTDQRCode_Click
End Sub

Private Sub lbReload_Click()
 frFilter.Visible = False
cbColFilter.ListIndex = -1

End Sub

Private Sub lbStandard_Click()
Dim rc As Boolean
rc = IIf(Grd1.Column(1).Width = 0, False, True)
STDToleranceGridColumn Grd1, rc
End Sub

Private Sub PicInformation_Click()
MyLot = Text1(0)
MyCode = Text1(1)
bAnotherFormCalled = True

F_INFORMATION.WindowState = Me.WindowState
F_INFORMATION.Left = Me.Left
F_INFORMATION.Top = Me.Top


If F_INFORMATION.DoShow(IndexFormProcedura, Text1(0), Text1(1), , Im, SettingName) Then
   ' Form_Initialize
    Text1(0) = MyLot
    Text1(1) = MyCode
    lbOperator = MyOperatore.Name
  
End If
bAnotherFormCalled = False

End Sub

Private Sub PicMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Len(Text1(IndexText)) = 0 Then Text1(IndexText).BackColor = vbColorUnabled
IndexText = 0
Picture1.BackColor = &H4000&
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Integer
For i = 0 To PicMenu.Count - 1
    If i = IndexFormProcedura Then
    Else
        PicMenu(i).BackColor = vbColorDarkFont
    End If
Next
Picture1.BackColor = &H4000&
End Sub

Private Sub Form_Unload(Cancel As Integer)

DoEvents
DoEvents
DoEvents

SaveSetting App.Title, "Information QC", "IndexFormProcedura", IndexFormProcedura
CloseSettingDataFile

DoEvents
DoEvents
DoEvents

Set F_READING = Nothing
End Sub

Private Sub Image3_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub lbCommand_Click(Index As Integer)
Dim rc As Boolean

    Select Case Index
        Case 14
            
            GoSub ph1
        Case 15
          
            GoSub ph2
        Case 78
            GoSub ph3
        Case 41
            rc = True
            GoSub SET_REAGENT
        Case 42
            rc = False
            GoSub SET_REAGENT
        Case 50
            rc = True
            GoSub REAGENT
        Case 51
            rc = False
            GoSub REAGENT
        Case Else
            Exit Sub
    End Select
    
Exit Sub

ph1:
    lbCommand(14).BackColor = &H80DF&
    lbCommand(15).BackColor = &H808080
    lbCommand(78).BackColor = &H808080
    IndexpH = 0
    Call SetPH(0)
Exit Sub
ph2:
    lbCommand(14).BackColor = &H808080
    lbCommand(15).BackColor = &H80DF&
    lbCommand(78).BackColor = &H808080
    IndexpH = 1
     Call SetPH(1)
Exit Sub
ph3:
    lbCommand(14).BackColor = &H808080
    lbCommand(15).BackColor = &H808080
    lbCommand(78).BackColor = &H80DF&
    IndexpH = 2
     Call SetPH(2)
Exit Sub
REAGENT:
  
    
    lbCommand(50).BackColor = IIf(rc, &H80DF&, &H808080)
    lbCommand(51).BackColor = IIf(Not (rc), &H80DF&, &H808080)
    Call ReagentSet(rc)
Exit Sub

SET_REAGENT:
    If rc Then
        lbCommand_Click 51
        IndexReagentSet = 1
    Else
        lbCommand_Click 50
        IndexReagentSet = 0
    End If
    
    lbCommand(41).BackColor = IIf(rc, &H80DF&, &H808080)
    lbCommand(42).BackColor = IIf(Not (rc), &H80DF&, &H808080)
  
Exit Sub
 
End Sub

Private Sub SetPH(ByVal Index As Integer)
Dim i As Integer

'checkit
With dbTabCode
    ' c'è  qualosa di sbagliato o da ricontrollare, questi numeri non li ho toccati
    If 40 + (Index * 3) < .Fields.Count - 10 Then
        For i = 0 To 2
            Text1(12 + i) = Trim(.Fields(55 + (Index * 3) + i))
        Next
    End If
End With
End Sub



Private Sub ReagentSet(ByVal bValue As Boolean)
Dim rc As Boolean
rc = bValue
'Label1(22).Caption = IIf(rc, "REAGENT SET 1", "REAGENT SET 2")
Frame5.ZOrder
Frame5.Visible = Not (rc)

IndexReagentSet = IIf(rc, 0, 1)




End Sub

Private Sub Label2_Click(Index As Integer)
PicMenu_Click Index
End Sub


Private Sub Label3_Click()
Grd1.ZOrder
Grd1.Visible = Not (Grd1.Visible)
lbStandard.Visible = Grd1.Visible
End Sub

Private Sub PicMenu_Click(Index As Integer)
frFilter.Visible = False
ReadOnlyGrid
If IndexFormProcedura = Index Then

Else
    Grd1.Cell(0, 0).SetFocus
    Grd2.Cell(0, 0).SetFocus
    Call SelectProcedura(Index)

    
End If
End Sub


Private Function SelectProcedura(ByVal Index As Integer) As Boolean
Dim rc As Boolean
Dim i As Integer

CloseSettingDataFile

    'If Index > 3 Then Exit Function
    For i = 0 To PicMenu.Count - 1
        If i = Index Then
            PicMenu(i).BackColor = vbColorForeFixed
        Else
            PicMenu(i).BackColor = vbColorDarkFont
        End If
    Next
    Set Image4(0) = Image3(Index)
    Frame3.Visible = False
    LaInfo.Visible = False
    Frame6.Visible = False
    Select Case Index
        Case 0
            Picture4(0).BackColor = &H80DF&
            rc = False
        Case 1
            FillAllGrid
           
            If CheckStandard Then
            Else
                
            End If
            rc = True
            Picture4(0).BackColor = &H70DF&

        Case 2
            rc = GetSetting(App.Title, "Reading QC", "Visualizza Colonne", False)
            Label7.Caption = IIf(rc, "Show Less", "Show More")
            Call SetViewGrid2(rc)
            Frame6.Visible = True
            Frame7.Visible = False
            
            rc = False
            Picture4(0).BackColor = &H60DF&
            Frame3.Visible = True
            
        Case 3
            
            rc = False
            LaInfo.Visible = True
            Picture4(0).BackColor = &H8000&
            'Frame3.Visible = True
            
            
        Case 4
            ' GRAPH QC
            If Grd2.Rows > 1 Then
                bAnotherFormCalled = True
                SalvaFormSettingName
                F_GRAPH.WindowState = Me.WindowState
                F_GRAPH.Top = Me.Top
                F_GRAPH.Left = Me.Left
                F_GRAPH.DoShow IndexFormProcedura, Text1(0), Text1(1), , Im, SettingName, STDSelectedNumber
            Else
                PopupMessage 2, "Please Enter at least 1 Test to open Graph QC for this Lot...", , True
            End If
            bAnotherFormCalled = False
            CloseSettingDataFile
            Exit Function
            
        Case 5
        
            ' evaluation
            If Grd2.Rows > 1 Then
                bAnotherFormCalled = True
                SalvaFormSettingName
                DoEvents
                DoEvents
                DoEvents
                F_EVALUATION.WindowState = Me.WindowState
                F_EVALUATION.Top = Me.Top
                F_EVALUATION.Left = Me.Left
               If F_EVALUATION.DoShow(IndexFormProcedura, Text1(0), Text1(1), , Im, SettingName) Then
               
               
               End If
               
            Else
                PopupMessage 2, "Please Enter at least 1 Test to open Evaluation QC for this Lot...", , True
            End If
            
            bAnotherFormCalled = False
            CloseSettingDataFile
            Exit Function
            
            
            
    End Select
    CloseSettingDataFile
    For i = 2 To 4
        Label1(i).Visible = rc
        Text1(i).Visible = rc
    Next
    LabStandard = Label2(Index)
    IndexFormProcedura = Index
    PicMain(Index).Visible = True
    PicMain(Index).ZOrder
  '  blTable = Label2(IndexFormProcedura)
   Call SetFinishedReadings
   
End Function



Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
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
ctlCalendar1.Left = Me.Width / 2 - ctlCalendar1.Width / 2
ctlCalendar1.Top = Me.Height / 2 - ctlCalendar1.Height / 2

Frame4.Left = 1200
Frame4.Top = 1680
Frame5.Top = 3600
Frame5.Left = 1800
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
    Call SaveTest
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Picture1.BackColor = &H8000&
End Sub


Private Sub Picture2_Click()
    
    If F_MsgBox.DoShow("Finished reading?", "Readings") Then
    
        m_rc = True
        
        CloseSettingDataFile
        SaveSettingData SettingName, "Reading", "Closed", True
        
        bFinishedReadings = True
        
        CloseSettingDataFile
        PopupMessage 2, "Reading closed..."
        
        SetFinishedReadings
        
    End If
    
    
End Sub


Private Function SetFinishedReadings() As Boolean
If bFinishedReadings Then
    Picture4(0).BackColor = Picture2.BackColor
    LabStandard = "Reading closed"
    Picture2.Visible = False
End If
End Function

Private Sub ReadSTDQRCode_Click()
          ' SCANN BARCODE
            Label15_Click
            Call ScanQRCodeQC
        
End Sub

Private Sub Text1_Change(Index As Integer)
Dim rc As Boolean
Dim mrc As Boolean
Dim MinValue As String
Dim MaxValue As String
Dim pHMin As String
Dim pHMax As String
Dim WeightMin As String
Dim WeightMax As String
On Error Resume Next
ctlCalendar1.Visible = False
rc = IIf(Len(Text1(Index)) > 0, True, False)
Text1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbWhite, vbColorUnabled)
If Index < 37 Then Label1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextBlue, vbColorLabelUnabled)
If Index > 64 Then Label1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextBlue, vbColorLabelUnabled)
'If Index = 63 Or Index = 64 Then Label1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextBlue, vbColorLabelUnabled)

Select Case Index
    Case 9
        lbTarget(9).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextDarkBlue, vbColorLabelUnabled)
    Case 33
        lbTestTable(35).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextDarkBlue, vbColorLabelUnabled)
    Case 35
        lbTestTable(34).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextDarkBlue, vbColorLabelUnabled)
        
End Select


Select Case Index
    Case 1

    Case 5
        Text1(3) = Text1(5)
        
        
        Text1(2) = Label6 - 1
    Case 9
        Text1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextLightBlue, vbColorUnabled)

        Text1(4) = Text1(9)
        Frame2.Visible = Not (rc)
       ' CancellaFormSTD

    Case 12
        Text1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbColorTextLightBlue, vbColorUnabled)
    Case 22 To 25
        ' meter
        MinValue = Text1(10)
        MaxValue = Text1(11)
    
        If Text1(Index) <> "" Then
            GoSub CheckValue
        Else
            Text1(Index).ForeColor = vbBlack
            lbColor(Index).ForeColor = vbBlack
        End If
    Case 27 ', 63, 64
        ' pH
        MinValue = ph(1, 1)
        MaxValue = ph(1, 2)
        If Text1(Index) <> "" Then
            GoSub CheckValue
        Else
            Text1(Index).ForeColor = vbBlack
            lbColor(Index).ForeColor = vbBlack
        End If
      Case 63, 64
        ' pH 2 e 3
        MinValue = ph(Index - 61, 1)
        MaxValue = ph(Index - 61, 2)
        If Text1(Index) <> "" Then
            GoSub CheckValue
        Else
            Text1(Index).ForeColor = vbBlack
            lbColor(Index).ForeColor = vbBlack
        End If
              
        
    Case 29
        ' Weight
        MinValue = Text1(39)
        MaxValue = Text1(40)
        If Text1(Index) <> "" Then
            GoSub CheckValue
        Else
            Text1(Index).ForeColor = vbBlack
            lbColor(Index).ForeColor = vbBlack
        End If


End Select

        
        If bEnterTable Then
            Grd2.Cell(lRow, lCol).ForeColor = lbColor(Index).ForeColor
        End If
    
    Exit Sub
    
CheckValue:

    If IsNumeric(Text1(Index)) Then
        If IsNumeric(MinValue) And IsNumeric(MaxValue) Then
            If CDbl(Text1(Index)) >= 0 Then
                If CDbl(Text1(Index)) >= CDbl(MinValue) And CDbl(Text1(Index)) <= CDbl(MaxValue) Then
                    Text1(Index).ForeColor = vbBlack
                    lbColor(Index).ForeColor = vbBlack
                Else
                    lbColor(Index).ForeColor = vbRed
                    Text1(Index).ForeColor = vbRed
                    
                  
                End If
            End If
        Else
            lbColor(Index).ForeColor = vbBlack
            Text1(Index).ForeColor = vbBlack
        End If
        

    End If
    
    Return
    
    
    
End Sub

Private Sub Text1_DblClick(Index As Integer)
    Select Case Index
    
        Case 9
        Case 16, 17
            Text1(16) = FormatDataLAT(date)
            Text1(17) = Replace(FormatDateTime(Now, vbShortTime), ".", ":")
        Case 18, 19, 20
            Text1(Index) = GetSetting(App.Title, "Information QC", "Text1" & Index, "")
        Case 22, 23, 24, 25

    End Select
End Sub

Private Sub Text1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
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
Select Case Index
    Case 30
        ' operatore
         If frmLogin.DoShow Then
            Text1(30) = MyOperatore.Name
         End If
End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    Select Case Index
        Case 5
            KeyAscii = TxtToNumber(KeyAscii)
        Case 9 To 14
            KeyAscii = TxtToNumber(KeyAscii)

        Case 22 To 29
            KeyAscii = TxtToNumber(KeyAscii)
    End Select
    
    If KeyAscii = 13 Then
        
        If Index < Text1.Count - 1 Then
            If Text1(Index + 1).Visible Or Not (Text1(Index + 1).Locked) Then Text1(Index + 1).SetFocus
        Else
            If IndexFormProcedura = 0 Then Text1(0).SetFocus
        End If
    End If
End Sub

Private Sub CheckTestType(ByVal sString As String)
sString = Trim(sString)
If (sString) = "" Then Exit Sub

    With dbTabTestType
        .filter = ""
        .filter = "Type='" & sString & "'"
        If .EOF Then
            If F_MsgBox.DoShow("New Test Type found : Add in Database? ") Then
                .AddNew
                !type = (sString)
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

GrdTestType.Visible = False

Combo1.Visible = False
Combo2.Visible = False

If Text1(Index).Locked Or Not (Text1(Index).Enabled) Then Exit Sub

Select Case Index
    Case 15
        GrdTestType.Visible = True
    Case 21
        If Combo1 = "Enter" Then Exit Sub
        If Text1(Index) <> "" Then
            If IsNumeric(Text1(Index)) And Text1(Index) > 0 And Text1(Index) < 4 Then
                Combo1 = Text1(Index)
            End If
        End If
        Combo1.ZOrder
        Combo1.Visible = True
    Case 18
        ctlCalendar1.Left = Text1(Index).Left + ctlCalendar1.Width / 2
        ctlCalendar1.Top = Frame1.Top + Text1(Index).Top + Text1(Index).Height - ctlCalendar1.Height / 2
        ctlCalendar1.Visible = True
        IndexTextSelected = Index
    Case 65
        If Combo2 = "Enter" Then Exit Sub
       ' If Text1(Index) <> "" Then
           ' If IsNumeric(Text1(Index)) And Text1(Index) > 0 And Text1(Index) < 4 Then
               ' Combo2 = Text1(Index)
            'End If
       ' End If
        Combo2.ZOrder
        Combo2.Visible = True
        
    Case Else
        ctlCalendar1.Visible = False
        IndexTextSelected = -1
End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbWhite, vbColorUnabled)

If bAnotherFormCalled Then Exit Sub

    Select Case Index
    
        Case 9
           If Text1(5) = "" Then Exit Sub
            ' CONTROLLO SE SONO IN TOLLERANZA TEXT1(41) E 42
            Call CheckToleranceSTDValue
            If Text1(5) <= STDCount Then
                Text1(10).Locked = True
                Text1(11).Locked = True
            Else
                Text1(10).Locked = False
                Text1(11).Locked = False
            End If
        Case 18, 19, 20
            SaveSetting App.Title, "Information QC", "Text1" & Index, Text1(Index)
        Case 22, 23, 24, 25
            Text1(Index) = Format$(Text1(Index), UserDecimal)

        Case 5
            
            If Text1(Index) <> "" And IsNumeric(Text1(Index)) Then
                 Text1(Index) = Int(Text1(Index))
                If Text1(Index) > 15 Or Text1(Index) < 0 Then
                    MessageInfoTime = 2000
                    PopupMessage 2, "Standard Number must be lower then 15 and greater then 0. Please Enter valid STD Number", , True, "Standard Number"
                    Text1(Index) = ""
                    Text1(Index).SetFocus
                    Exit Sub
                End If
                Call CheckTable(Text1(Index))
            Else
                'pulisco...
                If IndexFormProcedura <> 0 Then Exit Sub
                Text1(5) = ""
                CleanSTDText
                If bAnotherFormCalled Then
                Else
                If Text1(5).Enabled Then Text1(5).SetFocus
                End If
            End If
        Case 12, 13, 14 ' pH
        
            If Text1(Index) <> "" And IsNumeric(Text1(Index)) Then
                If Text1(Index) > 14 Or Text1(Index) < 0 Then
                    MessageInfoTime = 2000
                    PopupMessage 2, "pH must be lower then 14 and greater then 0. Please Enter valid pH Value", , True, "pH Value"
                    Text1(Index) = ""
                    Text1(Index).SetFocus
                    Exit Sub
                End If
               
            Else
            
            End If
            
        Case 15
            CheckTestType (Text1(Index))
    End Select
    
    
    
    
    
    
    
Call CheckProcedura(Index)
End Sub


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

Private Sub SetCombo()
Dim i As Integer

    Combo1.Clear
    For i = 1 To 4
        Combo1.AddItem i
    Next
    Combo1.AddItem "Enter"
    
    Label1(8) = Chr$(177) & " / " & Chr$(247)
    
    
    Combo2.Clear
    Combo3.Clear
    Combo3.AddItem "Select Hanna Code..."
    With dbTabCode
        .filter = ""
        .filter = "Recipe='" & Text1(37) & "'"
        If .EOF Then
        Else
            .MoveFirst
            For i = 1 To .RecordCount
                If Trim(!Code) = Text1(1) Then
                Else
                Combo2.AddItem Trim(!Code)
                Combo3.AddItem Trim(!Code)
                End If
                .MoveNext
            Next
            Combo2.AddItem "Enter"
            Combo3.AddItem "Enter"
    
        End If
        Combo3.ListIndex = 0
        .filter = ""
        .filter = "Code='" & MyCode & "'"
    
    End With
    
    
    
    
Combo4.AddItem "Select..."
Combo4.AddItem "STD"
Combo4.AddItem "Test"
Combo4.ListIndex = 0

    
End Sub



Private Sub SetTestinGrid()
Dim sValue As String
Dim dValue As Double
Dim bValue As Boolean

'UserDecimal

    
    With Grd2
        .AutoRedraw = False
        '.DefaultFont.Size = 12 * m_ControlGridFontSize
        .DefaultRowHeight = 40
        .Column(21).Width = IIf(ExsistReagentSet, .Column(21).Width, 0)
        .DefaultFont.Bold = False
        .AddItem "", False
        .Cell(.Rows - 1, 0).Text = .Rows - 1
        .Cell(.Rows - 1, 1).Text = Trim(Text1(3))
        .Cell(.Rows - 1, 2).Text = Trim(Text1(4))
        .Cell(.Rows - 1, 3).Text = .Rows - 1
        .Cell(.Rows - 1, 4).Text = Trim(Text1(15))
        .Cell(.Rows - 1, 5).Text = Trim(Text1(16)) '"QC DATE"
        .Cell(.Rows - 1, 6).Text = Trim(Text1(17)) '"QC TIME"
        .Cell(.Rows - 1, 7).Text = Trim(Text1(18)) '"PROD. DATE"
        .Cell(.Rows - 1, 8).Text = Trim(Text1(19)) '"PROD. TIME"
        .Cell(.Rows - 1, 9).Text = Trim(Text1(20)) '"MACHINE OPERATOR"
        .Cell(.Rows - 1, 10).Text = Trim(Text1(21)) '"HEAD"
        
        

        .Cell(.Rows - 1, 11).Text = Format$(Text1(22), UserDecimal) 'Trim(Text1(22)) ' "METER 1"
        .Cell(.Rows - 1, 12).Text = Format$(Text1(23), UserDecimal) 'Trim(Text1(23)) '"METER 2"
        .Cell(.Rows - 1, 13).Text = Format$(Text1(24), UserDecimal) 'Trim(Text1(24)) '"METER 3"
        .Cell(.Rows - 1, 14).Text = Format$(Text1(25), UserDecimal) 'Trim(Text1(25)) '"METER 4"
        
        .Cell(.Rows - 1, 15).Text = Trim(Text1(26)) '"SPECTR. [ABS]"
        .Cell(.Rows - 1, 16).Text = Trim(Text1(27)) ' "pH 1"
        .Cell(.Rows - 1, 17).Text = Trim(Text1(63)) ' "pH 2"
        .Cell(.Rows - 1, 18).Text = Trim(Text1(64)) ' "pH 3"
        
        .Cell(.Rows - 1, 19).Text = Trim(Text1(28)) '"TURB."
        .Cell(.Rows - 1, 20).Text = Trim(Text1(29)) '"WEIGHT [mg]"
        If ExsistReagentSet Then
            .Cell(.Rows - 1, 21).Text = IndexReagentSet + 1 '"REAGENT SET"
        End If
        .Cell(.Rows - 1, 22).Text = Trim(Text1(30)) ' "QC OPERATOR"
        ' "CORRECTION"
        .Cell(.Rows - 1, 23).Text = Trim(Text1(68))
        
        .Cell(.Rows - 1, 24).Text = 3 ' pHNumber
        
        .Cell(.Rows - 1, 25).Text = Trim(Text1(5)) '"STD"
        .Cell(.Rows - 1, 26).Text = Trim(Text1(9)) '"STD Value"
        .Cell(.Rows - 1, 27).Text = Trim(Text1(10)) '"STD Min"
        .Cell(.Rows - 1, 28).Text = Trim(Text1(11)) '"STD Max"
        
        
        .Cell(.Rows - 1, 29).Text = Trim(Text1(38)) '"Weight Value"
        .Cell(.Rows - 1, 30).Text = Trim(Text1(39)) '"Weight Min"
        .Cell(.Rows - 1, 31).Text = Trim(Text1(40)) '"Weight Max"

        .Cell(.Rows - 1, 32).Text = Trim(Text1(41)) ' "Range STD Min"
        .Cell(.Rows - 1, 33).Text = Trim(Text1(42)) '"Range STD Max"
        
        .Cell(.Rows - 1, 34).Text = Trim(Text1(65)) ' OTHER CODE
        .Cell(.Rows - 1, 35).Text = Trim(Text1(66)) 'LOT
        .Cell(.Rows - 1, 36).Text = Trim(Text1(67))
        .Cell(.Rows - 1, 37).Text = Trim(Text1(31)) 'NOTE
      
        Dim i As Integer
        
        For i = 22 To 29
            Debug.Print MeterNumber
            .Cell(.Rows - 1, i - 11).ForeColor = lbColor(i).ForeColor ' IIf(CheckMeterValue(Trim(Text1(3)), CDbl(.Cell(.Rows - 1, i - 11).Text), CDbl((Text1(10))), CDbl(Trim(Text1(11)))), vbBlack, vbRed)
        Next
        
        If MeterNumber > 0 Then
        
        For i = 22 To MeterNumber
            .Cell(.Rows - 1, i - 11).ForeColor = IIf(CheckMeterValue(Trim(Text1(3)), CDbl(.Cell(.Rows - 1, i - 11).Text), CDbl((Text1(10))), CDbl(Trim(Text1(11)))), vbBlack, vbRed)
        Next
        
        End If
        
        
        ' i due nuovi ph!!!!!
        .Cell(.Rows - 1, 17).ForeColor = lbColor(63).ForeColor
        .Cell(.Rows - 1, 18).ForeColor = lbColor(64).ForeColor
        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        
        ReadOnlyGrid
        .AutoRedraw = True
        .Refresh
        
    End With

End Sub

Private Function CheckMeterValue(ByVal Standard As Integer, ByVal MeterValue As Double, ByVal Min As Double, ByVal Max As Double) As Boolean
Dim rc As Boolean

        
    
                
            
    If Min = 0 And IsNumeric(Text1(33)) Then Min = CDbl(Text1(33))
    If Max = 0 And IsNumeric(Text1(34)) Then Min = CDbl(Text1(34))

    If MeterValue = 0 Then
        rc = False
        GoTo EXIT_CHECK
    End If
    
    If Max = 0 Then
        rc = True
        GoTo EXIT_CHECK
    End If
    
    rc = IIf(MeterValue >= Min And MeterValue <= Max, True, False)
    
    
EXIT_CHECK:
    CheckMeterValue = rc
End Function



Private Sub SetViewGrid2(ByVal bValue As Boolean)
Dim rc As Boolean
    With Grd2
    
    
        .Cell(0, 11).Text = "M. 1  " & lbFamily(0) & "  "
        .Cell(0, 12).Text = "M. 2  " & lbFamily(1) & "  "
        .Cell(0, 13).Text = "M. 3  " & lbFamily(2) & "  "
        .Cell(0, 14).Text = "M. 4  " & lbFamily(3) & "  "
         
         
        .Column(7).Width = IIf(bValue, 120, 0)
        .Column(8).Width = IIf(bValue, 150, 0)
        .Column(9).Width = IIf(bValue, 200, 0)
        .Column(10).Width = IIf(bValue, 80, 0)
       ' .Column(15).Width = IIf(bValue, 150, 0) ' spectr
        
         rc = IIf(ph(1, 0) <> "", True, False)
        
        .Column(16).Width = IIf(rc, 80, 0)
        
         rc = IIf(ph(2, 0) <> "", True, False)

        .Column(17).Width = IIf(rc, 80, 0)
        
         rc = IIf(ph(3, 0) <> "", True, False)
         
         
        .Column(18).Width = IIf(rc, 80, 0)
        
        
       ' .Column(19).Width = IIf(bValue, 120, 0)
        '.Column(20).Width = IIf(bValue, 150, 0)
        
        If ExsistReagentSet Then
            .Column(21).Width = IIf(bValue, 200, 0)
        Else
            .Column(21).Width = 0
        End If
        
        .Column(22).Width = IIf(bValue, 300, 0)
        
    End With
    
   
    
End Sub

Private Function CheckStandard() As Boolean
Dim rc As Boolean

    If Text1(9) = "" Or Text1(10) = "" Or Text1(11) = "" Then
    
        rc = False
    
    Else
    
        rc = True
    
    End If
    Frame4.ZOrder
    Frame4.Visible = Not (rc)
    CheckStandard = rc
End Function

Private Sub CheckProcedura(ByVal Index As Integer)
Dim rc As Boolean


Select Case IndexFormProcedura
    Case 0
    
    Case 1
        ' qualsiasi se non completo METER non vedo il compando Salva
        If Text1(22) = "" Or Text1(15) = "" Then
        
            Picture1.Visible = False
        Else
        
            Picture1.Visible = True
        End If
            
        
    Case 2
    
    



End Select




End Sub


Private Function SalvaFormSettingName()

    CloseSettingDataFile
    SaveSettingData SettingName, "Reading QC", "Operator", MyOperatore.Name
    SaveSettingData SettingName, "Reading QC", "Modification Date", FormatDataLAT(date)
    
    Call SaveGrd2Table(True)
    
    Call SaveGrapHData
    
    Call PutLotInDatabase
    
    CloseSettingDataFile
    
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
        !TestNumber = Grd2.Rows - 1
        .Update
End With

ERR_END:
    On Error GoTo 0
    PutLotInDatabase = rc
    Exit Function
ERR_SAVE:
    MsgBox Err.Description
    rc = False
    Resume ERR_END
End Function

Private Function SaveGrd2Table(ByVal bExit As Boolean)

    ' salva tabella
    

Dim i As Integer
Dim t As Integer
Dim iTest As String

    With Grd2
    
         iTest = .Rows
    
        If .Rows = 1 And Not bExit Then
            
           
            If F_InputBox.DoShow("Please enter number of Tests:", "Reading", , , , iTest, , True) Then
                iTest = iTest + 1
                
            End If
        
        
        End If
    
    
        SaveSettingData SettingName, "Reading QC", "Grd2 Rows", iTest
        SaveSettingData SettingName, "Reading QC", "Grd2 Cols", .Cols
        
        For i = 0 To .Rows - 1
            For t = 1 To .Cols - 1
            
                If .Cell(i, t).ForeColor = vbColorTextDarkBlue Then
                    .Cell(i, t).ForeColor = vbBlack
               End If
                SaveSettingData SettingName, "Reading QC", "Grd2 Row" & i & " Col" & t, .Cell(i, t).Text
                SaveSettingData SettingName, "Reading QC", "Grd2 Fore Row" & i & " Col" & t, .Cell(i, t).ForeColor
              
            Next
        Next
    End With
    
    CloseSettingDataFile
  ' PopupMessage "#" & iTest - 1 & " Readings saved...", SettingName
    
End Function


Private Function SaveTest()



    Call SetTestinGrid
    Call SaveAllTests
   
        
End Function


Private Function SaveAllTests()

Dim i As Integer
Dim t As Integer


    Call SaveGrd2Table(False)
    
    t = Grd2.Rows - 1
    Label6 = t + 1
    Text1(2) = Label6 - 1
    
    
    Call CancellaFormTestAfterReading
    
    PopupMessage 2, "Reading n." & t & " Saved...", , , , DefaultMenu(4)
        
End Function





Private Sub CancellaFormTestAfterReading()
Dim i As Integer
    For i = 15 To 32 ' gli altri sono  INFORMATION QC e STD
        If i = 18 Or i = 19 Or i = 20 Or i = 21 Or i = 30 Or i = 16 Or i = 17 Then
        Else
        Text1(i).Text = ""
        End If
    Next
    ' i 2 nuovi pH
    Text1(63).Text = ""
    Text1(64).Text = ""
   ' Text1(30) = MyOperatore.Name
    
    For i = 22 To 29
        lbColor(i).ForeColor = vbBlack
    Next
                

    Picture1.Visible = False
End Sub

Private Sub CancellaFormSTD()
Dim i As Integer
    For i = 3 To 36 ' gli altri sono  INFORMATION QC
        Text1(i).Text = ""
    Next
   ' Text1(30) = MyOperatore.Name
    'imposto pH0
    IndexpH = 0
    lbCommand_Click 14
End Sub




Private Function GetFormSettingName(Optional ByVal bCheckReadingsNumber As Boolean = False)
Dim i As Integer
Dim t As Integer


    On Error GoTo ERR_GET:
    ' doppio check... si sa mai...
    
   If FileExists(USER_TEMP_PATH & SettingName) Then
      USER_PATH = USER_TEMP_PATH
   ElseIf FileExists(USER_DATA_PATH & SettingName) Then
        'PopupMessage 2, "Lot : " & Text1(0) & vbCrLf & "Code : " & Text1(1) & vbCrLf & "Is Closed..."
        USER_PATH = USER_DATA_PATH
   Else
        PopupMessage 2, "Lot : " & Text1(0) & vbCrLf & "Code : " & Text1(1) & vbCrLf & "Warning : Information QC not found...", , True

   End If
   
   
    CloseSettingDataFile
    
      bFinishedReadings = GetSettingData(SettingName, "Reading", "Closed", False)
        
      
        
    

    
    With dbTabCode
        .filter = ""
        .filter = "ID='" & ReadingCodeID & "'"
        Debug.Print !Code
        If .EOF Then
            PopupMessage 2, "Please Enter a Vaid Hanna Code...", True
            Exit Function
        End If
    End With
    
    For i = 5 To 20
       Text1(i + 32) = GetSettingData(SettingName, "Information QC", "Text1" & i, Text1(i + 32))
       Text1(i + 32).Locked = True
    Next
     For i = 45 To 54
       Text1(i + 8) = GetSettingData(SettingName, "Information QC", "Text1" & i, Text1(i + 32))
       Text1(i + 8).Locked = True
    Next
    MeasurementUnit = IIf(IsNull(Trim(dbTabCode!MeasurementUnit)), "", Trim(dbTabCode!MeasurementUnit))
    If InStr(Text1(41), MeasurementUnit) Then Text1(41) = Trim(Left$(Text1(41), Len(Text1(41)) - Len(MeasurementUnit)))
    If InStr(Text1(42), MeasurementUnit) Then Text1(42) = Trim(Left$(Text1(42), Len(Text1(42)) - Len(MeasurementUnit)))
    
    lb(47) = "Reagent Range (" & MeasurementUnit & ")"
    
    ' tabella
    Dim MyRows As String
    Dim MyCols As String
    Dim Rows As Long
    
    Call SetGridtest(Grd2)
    
    
    MyRows = GetSettingData(SettingName, "Reading QC", "Grd2 Rows", 1)
    Rows = CLng(MyRows)
   
  ' If bCheckReadingsNumber = False Then
      '  Call CheckRows(Rows, SettingName, USER_PATH)
       ' MyRows = Rows
   ' End If
    
    MyCols = GetSettingData(SettingName, "Reading QC", "Grd2 Cols", 1)
    
    Dim bExit As Boolean
    Dim RowCount As Long
    Dim FileValue As String
    bExit = False
            
    If MyRows > 1 Then
        If MyCols > 1 Then
            With Grd2
              
                .DefaultFont.Size = 12 ' * m_ControlGridFontSize
                RowCount = 1
                i = 0
                Do
                    i = i + 1
                    FileValue = GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col1", "")
                    If FileValue <> "" Then
                        RowCount = RowCount + 1
                        .Rows = RowCount
                        For t = 1 To MyCols - 1
                            FileValue = GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & t, "")
                            .Cell(RowCount - 1, t).Text = FileValue
                            .Cell(RowCount - 1, t).ForeColor = GetSettingData(SettingName, "Reading QC", "Grd2 Fore Row" & i & " Col" & t, 0)
                        Next

                    Else
                      ' MsgBox "Result (ppm)"
                    End If
                    bExit = IIf(MyRows = RowCount Or i > 500, True, False)
                Loop Until bExit
                
                .Column(1).AutoFit
                .Column(2).AutoFit
                .Column(3).AutoFit
               
                .Column(5).AutoFit
                .Column(6).AutoFit
                .Column(7).AutoFit
                .Column(8).AutoFit
                .Column(9).AutoFit
                .Column(10).AutoFit
                .Column(11).AutoFit
                .Column(12).AutoFit
                .Column(13).AutoFit
                .Column(14).AutoFit
                
                
                .Column(27).AutoFit
                .Column(28).AutoFit
                
                
                
                
                .Cell(.Rows - 1, 1).EnsureVisible
            End With
        End If
    End If
    
    
    If MyRows = 0 Then
        Label6 = 1
    Else
    
        Label6 = MyRows
    End If
    
    Text1(2) = Label6 - 1
    
    MeterNumber = GetSettingData(SettingName, "Information QC", "MeterNumber", 0)
    pHNumber = 3 'GetSettingData(SettingName, "Information QC", "pHNumber", 0)
    
    GetpH
    
    
    ' Meter --------------
    
    For i = 0 To 3
    
        If i <= MeterNumber - 1 Then
            Label1(84 + i) = GetSettingData(SettingName, "Information QC", "Text1" & i * 2 + 34, 0)
            lbFamily(i) = GetSettingData(SettingName, "Information QC", "Text1" & i * 2 + 33, 0)
            Label1(22 + i).BackColor = vbColorTextDarkBlue
            
            Label1(84 + i).BackColor = vbColorTextBlue
            Text1(22 + i).Enabled = True
            Grd2.Column(i + 11).Width = 170
        Else
            Label1(84 + i) = "/"
            Label1(22 + i).BackColor = vbColorForeFixed
            Label1(84 + i).BackColor = vbColorForeFixed
            Text1(22 + i).Enabled = False
            Grd2.Column(i + 11).Width = 0
        End If
    
    Next


    Text1(30) = MyOperatore.Name
    
    UserDecimal = FormatDecimal(GetSettingData(SettingName, "Code Information", "Decimal", 0))
    
    
    
    ' REAGENT SET
    lbCommand(41).Visible = GetSettingData(SettingName, "Information QC", "Reagent Set 2", True)
    lbCommand(42).Visible = GetSettingData(SettingName, "Information QC", "Reagent Set 1", True)

    ExsistReagentSet = IIf(lbCommand(41).Visible Or lbCommand(42).Visible, True, False)
    
    ' imposto pH1
    lbCommand_Click 14
    

    SetSTDbyTable
    
    
    
    strValidDate = GetSettingData(SettingName, "Close QC", "Validation Date", "")
    bClosedLot = IIf(Len(Trim(strValidDate)) > 0, True, False)
    
    Call CheckClosedLot
    
     

ERR_END:
    On Error GoTo 0
    CloseSettingDataFile
    Exit Function
ERR_GET:
    MsgBox Err.Description
    Resume Next
End Function
Private Sub SetAllGrid()
    With GrdTestType
    
        .Rows = 1
        .ZOrder
        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        
        .DefaultRowHeight = 34
        .Cols = 3
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Test Type"
        .Cell(0, 1).Alignment = cellCenterCenter
        .Column(1).Width = 120
        .Cell(0, 2).Text = "ID"
        .Column(2).Width = 0

        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellLeftCenter
        Next
        '
        .DefaultFont.Size = 12 * m_ControlGridFontSize
        .DefaultFont.Bold = True
        ReadOnlyGrid
        .AutoRedraw = True
        .Refresh
    End With
    

End Sub
Private Sub FillAllGrid()

Dim i As Integer
GrdTestType.Rows = 1
cmbTestType.Clear
Frame7.Visible = False

With dbTabTestType
    .filter = ""
    
    If .EOF Then
    
    Else
        GrdTestType.AutoRedraw = False
        .MoveFirst
         cmbTestType.AddItem "Select Test Type..."
        For i = 1 To .RecordCount
            
            With GrdTestType
                .AddItem "", False
                .Cell(.Rows - 1, 1).Text = "  " & IIf(IsNull(Trim(dbTabTestType!type)), "", Trim(dbTabTestType!type))
                .Cell(.Rows - 1, 1).FontBold = True
            End With
            cmbTestType.AddItem IIf(IsNull(Trim(dbTabTestType!type)), "", Trim(dbTabTestType!type))
         
       
            .MoveNext
       Next
       
        cmbTestType.ListIndex = 0
        
        GrdTestType.ReadOnly = True
        GrdTestType.AutoRedraw = True
        GrdTestType.Refresh
    End If

    
End With


End Sub

Private Function CheckToleranceSTDValue()
Dim MinValue As String
Dim MaxValue As String
Dim Perc As Double
Dim Restr As Double
Dim Fixed As Double

On Error GoTo ERR_CHECK

    If Text1(9) <> "" Then
        Text1(9) = Format$(Text1(9), UserDecimal)
        If Trim(Text1(42)) = "" Or Trim(Text1(41)) = "" Then
            GoTo okcheck
        End If
        If (Trim(Text1(42)) <> "" Or Trim(Text1(42)) <> "0") And (Trim(Text1(41)) <> "" Or Trim(Text1(41)) <> "0") Then
            If CDbl(Text1(9)) <= CDbl(Text1(42)) And CDbl(Text1(9)) >= CDbl(Text1(41)) Then
                ' ok!!!
okcheck:
               ' MsgBox Grd1.Cell(2, 3).Text
                If Grd1.Cell(2, 3).Text <> "" Then
                    If IsNumeric(Replace((Grd1.Cell(2, 3).Text), "%", "")) Then
                        
                        Perc = Val(Replace((Grd1.Cell(2, 3).Text), "%", ""))
                        
                        
                        If Perc > 0 Then Perc = Perc / 100
                    Else
                        Perc = 0
                    End If
                End If
                If IsNumeric(Val(Grd1.Cell(2, 4).Text)) Then
                    Restr = Val(Grd1.Cell(2, 4).Text) / 100
                Else
                    Restr = 0
                End If
                If IsNumeric(Val(Grd1.Cell(2, 1).Text)) Then
                    Fixed = Val(Grd1.Cell(2, 1).Text)
                Else
                    Fixed = 0
                End If
                
                If StandardCal(Text1(9), Fixed, Grd1.Cell(2, 2).Text, Perc, Restr, UserDecimal, MinValue, MaxValue) Then
                    Text1(10) = MinValue
                    Text1(11) = MaxValue
                End If
                Text1(10).Locked = True
                Text1(11).Locked = True
                
                Text1(33) = Text1(10)
                Text1(34) = Text1(11)
            Else
                Text1(10) = ""
                Text1(11) = ""
                Text1(33) = ""
                Text1(34) = ""
                Text1(10).Locked = False
                Text1(11).Locked = False
                Text1(10).SetFocus
            End If
        End If
    End If
    
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_CHECK:
    MsgBox Err.Description
    Resume Next
End Function

Private Sub CheckTable(ByVal STDNumber As Integer)
Dim rc As Boolean
If STDNumber < 7 And STDNumber > 0 Then
    If IsNumeric(Grd1.Cell(2, STDNumber * 3 + 3).Text) Then
     Grd1.Cell(1, STDNumber * 3 + 3).SetFocus
     Grd1_SelChange 1, STDNumber * 3 + 3, 1, STDNumber * 3 + 3
     Else
        GoTo cont:
     End If
Else
cont:
    rc = CheckSTDArrayAndCopyValue(STDNumber)
    
    If rc = False Then
        ' pulisco ....
        CleanSTDText
        Text1(9).SetFocus
    End If
End If
End Sub
Private Function CheckSTDArrayAndCopyValue(ByVal STDNumber As Integer) As Boolean
Dim rc As Boolean
Dim i As Integer
    rc = False
    If STDNumber = 0 Then Exit Function
    ReDim Preserve STD(10, 3) As String
    For i = 1 To UBound(STD)
        If STD(i, 0) = "" Then GoTo exitfunction
        If CInt(STD(i, 0)) = STDNumber Then
            ' ho selezionato uno standard

            Text1(9) = STD(i, 1)
            
            Text1_LostFocus 9
            
            'STD(t, 0) = STDNumber
            'STD(t, 1) = STDValue
            'STD(t, 2) = STDMin
            'STD(t, 3) = STDMax
                                    
            Text1(10) = STD(i, 2)
            Text1(11) = STD(i, 3)
            
            Text1(33) = Text1(10)
            Text1(34) = Text1(11)
            Text1(10).Locked = True
            Text1(11).Locked = True
            
            
            rc = True
            Exit For
        End If
    Next
exitfunction:
    
 CheckSTDArrayAndCopyValue = rc
End Function

Private Function GetpH()
Dim i As Integer
Dim rc As Boolean
Dim Num As Integer
Dim sString As String
    Num = 0
    
    CloseSettingDataFile
    With dbTabCode
        For i = 1 To 3
            Num = Num + 1
            rc = True
            sString = GetSettingData(SettingName, "Information QC", "pHValue" & Num, "0")
            
            If (sString <> "" Or sString <> "/") And IsNumeric(sString) Then
                ph(Num, 0) = sString
            Else
                ph(Num, 0) = ""
                rc = False
            End If
        
            ph(Num, 1) = GetSettingData(SettingName, "Information QC", "pHMin" & Num, "0")
            ph(Num, 2) = GetSettingData(SettingName, "Information QC", "pHMax" & Num, "0")
            
        
            Select Case i
                Case 1
                    Text1(27).Enabled = IIf(rc, True, False)
                    Text1(27) = IIf(rc, "", "/")
                    Label1(27).BackColor = IIf(rc, vbColorTextBlue, vbColorForeFixed)
                   
                Case 2
                    Text1(63).Enabled = IIf(rc, True, False)
                    Text1(63) = IIf(rc, "", "/")
                    Label1(63).BackColor = IIf(rc, vbColorTextBlue, vbColorForeFixed)
                    
                Case 3
                    Text1(64).Enabled = IIf(rc, True, False)
                    Text1(64) = IIf(rc, "", "/")
                    Label1(64).BackColor = IIf(rc, vbColorTextBlue, vbColorForeFixed)
            End Select
        Next
    End With
    
            ' pH ----------------
    
    CloseSettingDataFile

    
    
    
    
    
End Function



'--------------------------------
' salvo i dati per Graph QC
'--------------------------------



Private Function SaveGrapHData()
    Dim rc As Boolean

    
    On Error GoTo ERR_GET

    rc = SetSTDbyTable
ERR_END:
    On Error GoTo 0
    If rc Then
        
        If STDCount > 0 Then
            Call SaveValueForGraph(STDCount)
            Call SaveGrapHTest(STDCount)
        End If
    End If
    Exit Function
ERR_GET:
    rc = False
    MsgBox Err.Description
    Resume Next
End Function

Private Function SetSTDbyTable() As Boolean
    Dim MyRows As String
    Dim MyCols As String
    Dim i As Integer
    Dim t As Integer
    Dim j As Integer
    Dim rc As Boolean
    Dim STDNumber As String
    Dim STDValue As String
    Dim STDMin As String
    Dim STDMax As String
    Dim bControlloTabellaReadings As Boolean
    
    On Error GoTo ERR_SET:
    
    rc = False
    
        MyRows = Grd2.Rows
    MyCols = Grd2.Cols
    
    If MyRows < 2 Then
        Exit Function
    End If

  
    t = 0
    ReDim STD(10, 3) As String
    bControlloTabellaReadings = False
    With Grd2
        If MyRows > 1 Then
            If MyCols > 1 Then
    
                
                For i = 1 To MyRows - 1
                    
                        STDNumber = .Cell(i, 1).Text ' GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "1", "")
                        
                        STDValue = .Cell(i, 2).Text 'GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "2", "")
                        STDMin = .Cell(i, 27).Text ' GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "27", "")
                        STDMax = .Cell(i, 28).Text ' GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "28", "")
                        
                         If STDNumber = "" Then GoTo cont:
                        
                        If t > 0 Then
                            ' faccio subito un controllo su STDs e se lo trovo prendo i valori e vado avanti
                            For j = 1 To t
                                If STD(j, 0) = STDNumber Then
                                    
                                    
                                    
                                    STDValue = STD(j, 1)
                                    STDMin = STD(j, 2)
                                    STDMax = STD(j, 3)
                                    
                                  
                                    GoTo setTableCell
                                    
                                End If
                            
                            Next
                        
                        
                        End If
                        
                        
                        If STDValue = "" Or STDValue = "/" Then GoTo cont:
                        If STDMin = "" Or STDMax = "" Then GoTo check
                        
                        If CDbl(STDValue) < CDbl(STDMin) Or CDbl(STDValue) > CDbl(STDMax) Then
                            ' ci sono degli errori
check:
                            bControlloTabellaReadings = True
                            
                            If t = 0 Then
                            t = 1
                                STD(t, 0) = STDNumber
                                STD(t, 1) = STDValue
                                STD(t, 2) = STDMin
                                STD(t, 3) = STDMax
                            End If
                            
                            If GetSTDFormDatabase(t, STDNumber, STDValue, STDMin, STDMax) Then GoTo vaiavanti
                            
                            
                            GoTo cont:
                        End If
                        
vaiavanti:
                        
                        If STDNumber <> "" And STDMin <> "" And STDMax <> "" Then
                        
                            If STDNumber = "/" Then STDNumber = "1"
                            If STDMin = "/" Then STDMin = "0"
                            If STDMax = "/" Then STDMax = "0"
                        
                            If t > 0 Then
                                If GetIndexArStr(STD(), STDNumber) = -1 Then
Aggiungi:
                                    
                                    t = t + 1
                        
                                    STD(t, 0) = STDNumber
                                    STD(t, 1) = STDValue
                                    STD(t, 2) = STDMin
                                    STD(t, 3) = STDMax
                                    If CDbl(STDValue) < CDbl(STDMin) Or CDbl(STDValue) > CDbl(STDMax) Then Call GetSTDFormDatabase(t, STDNumber, STDValue, STDMin, STDMax)
                                 
                                    rc = True
                                End If
                            Else
                                GoTo Aggiungi
                            End If
                        End If
                        
setTableCell:
               
                        
                        .Cell(i, 1).Text = STDNumber 'STD(t, 0)  ' GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "1", "")
                        
                        .Cell(i, 2).Text = STDValue 'STD(t, 1) 'GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "2", "")
                        .Cell(i, 27).Text = STDMin 'STD(t, 2)  ' GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "27", "")
                        .Cell(i, 28).Text = STDMax 'STD(t, 3)
                        
                        
                      
                        
                        
                        
cont:
                Next
            End If
        End If
    End With
ERR_END:
    On Error GoTo 0
    
    STDCount = t
    
    SaveValueForGraph (STDCount)
      
    Call CheckAndCorrectReadingsTable
   
    SetSTDbyTable = rc
    Exit Function
    
ERR_SET:
    rc = False
    MsgBox Err.Description
    Resume Next
End Function

Private Function CheckSTDAfterreadingTable()
Dim t As Integer
    Dim STDNumber As String
    Dim STDValue As String
    Dim STDMin As String
    Dim STDMax As String
    
    For t = 1 To UBound(STD)
        
        If STD(t, 0) = "" Then
        
        Else
            STDNumber = STD(t, 0)
            STDValue = STD(t, 1)
            STDMin = STD(t, 2)
            STDMax = STD(t, 3)
                        
            If CDbl(STDValue) < CDbl(STDMin) Or CDbl(STDValue) > CDbl(STDMax) Then
                            ' ci sono degli errori
                
                If IsNumeric(STD(t, 0)) Then
                    
                        If GetSTDFormDatabase(t, STDNumber, STDValue, STDMin, STDMax) = False Then
                            GoTo cont:
                        Else
                            
                            STD(t, 0) = STDNumber
                            STD(t, 1) = STDValue
                            STD(t, 2) = STDMin
                            STD(t, 3) = STDMax
                            
                            
                        
                        
                        End If
                End If
            End If
                        
        
        End If
    
cont:
    Next

End Function

Private Function GetSTDFormDatabase(ByVal t As Integer, ByVal STDNumber As String, ByRef STDValue As String, ByRef STDMin As String, ByRef STDMax As String) As Boolean
Dim rc As Boolean
Dim Index As Integer

'If (STD(t, 0)) = "" Then
    If STDNumber <> "" Then
        Index = CInt(STDNumber)
    End If
'Else
    'Index = CInt(STD(t, 0))
'End If

           

rc = False
    Select Case Index
        Case 1
            If IsNumeric(dbTabCode!STD1Value) = False Then GoTo checkValueTolerance
            STDValue = IIf(IsNull(Trim(dbTabCode!STD1Value)), "", Trim(dbTabCode!STD1Value))      '& "Value"
            STDMin = IIf(IsNull(Trim(dbTabCode!STD1Min)), "", Trim(dbTabCode!STD1Min))      '& "Min"
            STDMax = IIf(IsNull(Trim(dbTabCode!STD1Max)), "", Trim(dbTabCode!STD1Max))       '& "Max"
        
        Case 2
            If IsNumeric(dbTabCode!STD2Value) = False Then GoTo checkValueTolerance
            STDValue = IIf(IsNull(Trim(dbTabCode!STD2Value)), "", Trim(dbTabCode!STD2Value))      '& "Value"
            STDMin = IIf(IsNull(Trim(dbTabCode!STD2Min)), "", Trim(dbTabCode!STD2Min))      '& "Min"
            STDMax = IIf(IsNull(Trim(dbTabCode!STD2Max)), "", Trim(dbTabCode!STD2Max))       '& "Max"
            
        Case 3
            If IsNumeric(dbTabCode!STD3Value) = False Then GoTo checkValueTolerance
            STDValue = IIf(IsNull(Trim(dbTabCode!STD3Value)), "", Trim(dbTabCode!STD3Value))      '& "Value"
            STDMin = IIf(IsNull(Trim(dbTabCode!STD3Min)), "", Trim(dbTabCode!STD3Min))      '& "Min"
            STDMax = IIf(IsNull(Trim(dbTabCode!STD3Max)), "", Trim(dbTabCode!STD3Max))       '& "Max"
        
        Case 4
            If IsNumeric(dbTabCode!STD4Value) = False Then GoTo checkValueTolerance
            STDValue = IIf(IsNull(Trim(dbTabCode!STD4Value)), "", Trim(dbTabCode!STD4Value))      '& "Value"
            STDMin = IIf(IsNull(Trim(dbTabCode!STD4Min)), "", Trim(dbTabCode!STD4Min))      '& "Min"
            STDMax = IIf(IsNull(Trim(dbTabCode!STD4Max)), "", Trim(dbTabCode!STD4Max))       '& "Max"
        Case 5
            If IsNumeric(dbTabCode!STD5Value) = False Then GoTo checkValueTolerance
            STDValue = IIf(IsNull(Trim(dbTabCode!STD5Value)), "", Trim(dbTabCode!STD5Value))      '& "Value"
            STDMin = IIf(IsNull(Trim(dbTabCode!STD5Min)), "", Trim(dbTabCode!STD5Min))      '& "Min"
            STDMax = IIf(IsNull(Trim(dbTabCode!STD5Max)), "", Trim(dbTabCode!STD5Max))       '& "Max"
        
        Case 6
            If IsNumeric(dbTabCode!STD6Value) = False Then GoTo checkValueTolerance
            STDValue = IIf(IsNull(Trim(dbTabCode!STD6Value)), "", Trim(dbTabCode!STD6Value))      '& "Value"
            STDMin = IIf(IsNull(Trim(dbTabCode!STD6Min)), "", Trim(dbTabCode!STD6Min))      '& "Min"
            STDMax = IIf(IsNull(Trim(dbTabCode!STD6Max)), "", Trim(dbTabCode!STD6Max))
        Case Else
       
    
    End Select
    rc = True
cont:
    GetSTDFormDatabase = rc
    Exit Function
    
checkValueTolerance:

    Dim Fixed As Double
    Dim AndOr As String
    Dim Perc As Double
    Dim Restr As Double
  
    
    AndOr = UCase(Trim(Grd1.Cell(2, 2).Text))
      
    
    If IsNumeric(Val(Grd1.Cell(2, 3).Text)) Then
        
        Perc = Val(Grd1.Cell(2, 3).Text)
        If Perc > 0 Then Perc = Perc / 100
    Else
        Perc = 0
    End If
    If IsNumeric(Val(Grd1.Cell(2, 4).Text)) Then
        Restr = Val(Grd1.Cell(2, 4).Text) / 100
    Else
        Restr = 0
    End If
    If IsNumeric(Val(Grd1.Cell(2, 1).Text)) Then
        Fixed = (Grd1.Cell(2, 1).Text)
    Else
        Fixed = 0
    End If
                

    rc = True
    Call StandardCal(STDValue, Fixed, AndOr, Perc, Restr, UserDecimal, STDMin, STDMax)
    GoTo cont:
    
End Function

Private Function CheckAndCorrectReadingsTable()
    Dim MyRows As String
    Dim MyCols As String
    Dim i As Integer
    Dim t As Integer
    Dim rc As Boolean
    Dim STDNumber As String
    Dim STDValue As String
    Dim STDMin As String
    Dim STDMax As String
    Dim MeterValue As String
    Dim ResultColor As OLE_COLOR
    
    On Error GoTo ERR_CHECK:
    Call CheckSTDAfterreadingTable
    
    With Grd2
        If .Rows > 0 Then
            For i = 1 To .Rows - 1
                STDNumber = .Cell(i, 1).Text
                
                If IsNumeric(STDNumber) Then
                
                    
                    Call CheckSTDArray(STDNumber, STDValue, STDMin, STDMax)
                    
                    If STDMin = "" Or STDMax = "" Then GoTo cont
                    .Cell(i, 2).Text = STDValue
                    .Cell(i, 27).Text = STDMin
                    .Cell(i, 28).Text = STDMax
                    
                    ' check values.... 12,13,14
                    For t = 11 To 14
                         ResultColor = vbBlack
                        MeterValue = .Cell(i, t).Text
                        If IsNumeric(MeterValue) Then
                            
                            If CDbl(MeterValue) >= CDbl(STDMin) And CDbl(MeterValue) <= CDbl(STDMax) Then
                                ResultColor = vbBlack
                            Else
                                ResultColor = vbRed
                            End If
                        End If
                        
                        .Cell(i, t).ForeColor = ResultColor
                        
                        
                    Next
                    
                    
                    
                    
                    
                End If
                
cont:
            Next
        End If
            
            'STDNumber = .Cell(i, 1).Text ' GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "1", "")
            'STDValue = .Cell(i, 2).Text 'GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "2", "")
            'STDMin = .Cell(i, 27).Text ' GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "27", "")
            'STDMax = .Cell(i, 28).Text ' GetSettingData(SettingName, "Reading QC", "Grd2 Row" & i & " Col" & "28", "")
            'STD(t, 0) = STDNumber
            'STD(t, 1) = STDValue
            'STD(t, 2) = STDMin
            'STD(t, 3) = STDMax


    End With
    
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_CHECK:
    MsgBox Err.Description
    Resume Next
    


End Function




Private Function CheckSTDArray(ByVal STDNumber As Integer, ByRef STDValue As String, ByRef STDMin As String, ByRef STDMax As String)
Dim rc As Boolean
Dim i As Integer
    rc = False
    For i = 1 To UBound(STD)
        If STD(i, 0) = "" Then GoTo exitfunction
        If CInt(STD(i, 0)) = STDNumber Then
            STDValue = STD(i, 1)
            STDMin = STD(i, 2)
            STDMax = STD(i, 3)
            rc = True
            Exit For
        End If
    Next
exitfunction:
    
 CheckSTDArray = rc
End Function




Private Function CheckSTDInTable(ByVal numSTD As Integer, ByRef Index As Integer) As Boolean
Dim rc As Boolean
rc = True

If GetIndexArStr(STD(), numSTD, Index) = -1 Then
    rc = False
Else
    rc = IIf(STD(Index, 1) <> "", True, False)
End If
CheckSTDInTable = rc
End Function

Private Function SaveValueForGraph(ByVal STDCount As Integer)
 Dim t As Integer
 If STDCount = 0 Then Exit Function
        SaveSettingData SettingName, "Graph QC", "STDCount", STDCount
        For t = 1 To STDCount

            SaveSettingData SettingName, "Graph QC", "STDNumber" & t, STD(t, 0)
            SaveSettingData SettingName, "Graph QC", "STDValue" & t, STD(t, 1)
            SaveSettingData SettingName, "Graph QC", "STDMin" & t, STD(t, 2)
            SaveSettingData SettingName, "Graph QC", "STDMax" & t, STD(t, 3)
        Next
        
End Function


Private Function SaveGrapHTest(ByVal STDCount As Integer)
Dim i As Integer
Dim t As Integer
Dim k As Integer
Dim rc As Boolean
Dim lRows As Long
Dim lCols As Long
Dim ReadingStandard(99) As String
Dim numSelectedStandard As Integer
Dim STDtest(99, 6) As String
Dim STDTestCount As Integer
Dim ReadingsCount As Integer
Dim stdType As String
Dim mrc As Boolean
    CloseSettingDataFile
    
    If STDCount = 0 Then Exit Function
    
    On Error GoTo ERR_GET:
    
    If MeterNumber = 0 Or Grd2.Rows < 2 Then
        Exit Function
    End If
    
    
    ' salvo per grafico
        
        
    With Grd2
        lRows = .Rows
        lCols = .Cols
        For k = 1 To STDCount
            STDTestCount = 0
            ReadingsCount = 0
            For t = 1 To lRows - 1
                If STD(k, 0) = Trim(.Cell(t, 1).Text) Then
                
                    '
                    ' salvo solo i Test Type = "P"
                    '
                    
                    stdType = Trim(.Cell(t, 4).Text)
                    
                    If InStr(stdType, "P") Then
                    
                        STDTestCount = STDTestCount + 1
                        
                        For i = 1 To MeterNumber
                            SaveSettingData SettingName, "Graph QC", "Standard " & STD(k, 0) & " Test " & STDTestCount & " Meter " & i & " Value", .Cell(t, i + 10).Text
                            
                            If .Cell(t, i + 10).Text <> "" Then ReadingsCount = ReadingsCount + 1
                        Next
                        SaveSettingData SettingName, "Graph QC", "Standard " & STD(k, 0) & " Test " & STDTestCount & " ABS", .Cell(t, 15).Text
    
                        SaveSettingData SettingName, "Graph QC", "Standard " & STD(k, 0) & " Test " & STDTestCount & " Real Test", .Cell(t, 3).Text
    
                        Debug.Print .Cell(t, 15).Text
                        
                        '
                        ' salvo solo i pH presenti nel Code
                        '
                        
                        mrc = IIf(ph(1, 0) <> "", True, False)
                        If mrc Then SaveSettingData SettingName, "Graph QC", "Standard " & STD(k, 0) & " Test " & STDTestCount & " pH1 " & " Value", .Cell(t, 16).Text
                        mrc = IIf(ph(2, 0) <> "", True, False)
                        If mrc Then SaveSettingData SettingName, "Graph QC", "Standard " & STD(k, 0) & " Test " & STDTestCount & " pH2 " & " Value", .Cell(t, 17).Text
                        mrc = IIf(ph(3, 0) <> "", True, False)
                        
                        If mrc Then SaveSettingData SettingName, "Graph QC", "Standard " & STD(k, 0) & " Test " & STDTestCount & " pH3 " & " Value", .Cell(t, 18).Text
                    
                    End If
                End If
            Next
            SaveSettingData SettingName, "Graph QC", "Standard " & STD(k, 0) & " Total Readings", ReadingsCount
            SaveSettingData SettingName, "Graph QC", "Standard " & STD(k, 0) & " Total Tests", STDTestCount
        Next
        
    End With
    
ERR_END:
    
    CloseSettingDataFile
    Exit Function
ERR_GET:
    MsgBox Err.Description
    Resume Next
End Function


Private Function AddLine()

    Call SetSTDbyTable
    
    Call FillAllGrid
    
    
        Call AddLastLine
    

End Function

Private Sub AddLastLine()
Dim i As Integer

        Text1(16) = FormatDataLAT(date)
        Text1(17) = Replace(FormatDateTime(Now, vbShortTime), ".", ":")
    
    With Grd2
        .AutoRedraw = False
        .ReadOnly = False
        .SelectionMode = cellSelectionNone
        .AddItem "", False
        If Grd2.Rows > 2 Then
            For i = 1 To .Cols - 1
                .Cell(.Rows - 1, i).Text = .Cell(.Rows - 2, i).Text
                .Cell(.Rows - 1, i).ForeColor = vbColorTextDarkBlue
            Next
        Else
            .Cell(.Rows - 1, 7).Text = Trim(Text1(18)) '"PROD. DATE"
            .Cell(.Rows - 1, 8).Text = Trim(Text1(19)) '"PROD. TIME"
            .Cell(.Rows - 1, 9).Text = Trim(Text1(20)) '"MACHINE OPERATOR"
            
        End If
       
        .Cell(.Rows - 1, 5).Text = FormatDataLAT(date)
        .Cell(.Rows - 1, 6).Text = FormatTimeLAT(FormatDateTime(Now, vbShortTime))
        
        .Cell(.Rows - 1, 3).Text = .Rows - 1
    
        .Cell(.Rows - 1, 11).Text = ""
        .Cell(.Rows - 1, 12).Text = ""
        .Cell(.Rows - 1, 13).Text = ""
        .Cell(.Rows - 1, 14).Text = ""
        
        .Cell(.Rows - 1, 16).Text = ""
        .Cell(.Rows - 1, 17).Text = ""
        .Cell(.Rows - 1, 18).Text = ""
        .Cell(.Rows - 1, 19).Text = ""
        .Cell(.Rows - 1, 20).Text = ""
        
        If ExsistReagentSet Then
            .Cell(.Rows - 1, 21).Text = IndexReagentSet + 1 '"REAGENT SET"
        End If
        .Cell(.Rows - 1, 22).Text = Text1(30)
        .Cell(.Rows - 1, 23).Text = ""
        .Cell(.Rows - 1, 24).Text = 3 ' pHNumber

        .Cell(.Rows - 1, 29).Text = Trim(Text1(38)) '"Weight Value"
        .Cell(.Rows - 1, 30).Text = Trim(Text1(39)) '"Weight Min"
        .Cell(.Rows - 1, 31).Text = Trim(Text1(40)) '"Weight Max"
        .Cell(.Rows - 1, 32).Text = Trim(Text1(41)) ' "Range STD Min"
        .Cell(.Rows - 1, 33).Text = Trim(Text1(42)) '"Range STD Max"

        
        ActivateGrd2
        
        .Refresh
        .AutoRedraw = True
        .Cell(.Rows - 1, 1).EnsureVisible
       ' .Cell(.Rows - 1, 1).BackColor = vbColorTextLightBlue
       ' .Cell(.Rows - 1, 1).ForeColor = vbWhite
        

    End With

End Sub


Private Sub Grd2_CellChange(ByVal Row As Long, ByVal Col As Long)
Dim strModello As String
 Dim Intid As Long

 On Local Error GoTo ERR_LIST

    With Grd2
        
        If Row < 1 Then Exit Sub
        If Me.Visible = False Then Exit Sub
        
        Select Case Col
            
            Case 1
            
        
        
        
        End Select
    End With
ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_LIST:
    'MsgBox Err.Description
    Resume ERR_END
End Sub
Private Sub Grd2_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 46
        ' delete line
        DeleteLine
End Select
End Sub



Private Sub InsertDataInGrid()
Dim sString As String
Dim iSTD As Integer
Dim Index As Integer
Dim printString As String
Dim rc As Boolean


start:
    printString = ""
    bEnterTable = True

    rc = False
    

With Grd2
    Select Case lCol
        Case 1
            ' STD
            
            Index = 5
            sString = .Cell(lRow, lCol).Text
            If F_InputBox.DoShow("Enter numeric value...", "#" & lRow & " - " & .Cell(0, lCol).Text, , , , sString, , True) Then
                iSTD = CInt(sString)
                CheckSTD
                rc = True
            End If
        Case 2
            Index = 9
            If CheckSTD = False Then Exit Sub

            If .Cell(lRow, 1).Text = "" Then
                PopupMessage 2, "Please enter STD nuber first..."
                lCol = 1
                GoTo start:
            End If
            sString = .Cell(lRow, lCol).Text
            If F_InputBox.DoShow("Enter numeric value...", "#" & lRow & " - " & .Cell(0, lCol).Text, , , , sString, , True) Then
                If Text1(5) = "" Then
                    Text1(5) = .Cell(lRow, 1).Text
                End If
               Text1(9) = sString
               
               
                rc = True
            End If
        Case 4
            ' TEST
            cmbTestType.ZOrder
            cmbTestType.Visible = True
            Frame7.Visible = True
            rc = False
            Exit Sub
            
        Case 11, 12, 13, 14
                ' METER!!!!
            CheckSTD
            
            DoEvents
            
            If Text1(10) = "" Or Text1(9) = "" Then
                ' devo impostare lo standard...
                SetSTDTextbyTable
            End If
            
            CheckToleranceSTDValue
                
            Index = lCol + 11
            sString = .Cell(lRow, lCol).Text
            printString = "#" & lRow & " - " & .Cell(0, lCol).Text & " - STD " & .Cell(lRow, 1).Text
            
            If F_InputBox.DoShow("Enter numeric value - Min = " & Text1(33) & " - Max = " & Text1(34), printString, , , , sString, , True) Then

                rc = True
            End If
            
            
            
        Case 16
            If Text1(10) = "" Or Text1(9) = "" Then
                ' devo impostare lo standard...
                SetSTDTextbyTable
            End If
            Index = lCol + 11
            sString = .Cell(lRow, lCol).Text
            If F_InputBox.DoShow("Enter numeric value - Min = " & ph(1, 1) & " - Max = " & ph(1, 2), "#" & lRow & " - " & .Cell(0, lCol).Text & " - STD " & .Cell(lRow, 1).Text, , , , sString, , True) Then

                rc = True
            End If
         Case 17, 18
            If Text1(10) = "" Or Text1(9) = "" Then
                ' devo impostare lo standard...
                SetSTDTextbyTable
            End If
            Index = lCol + 46
            sString = .Cell(lRow, lCol).Text
            If F_InputBox.DoShow("Enter numeric value - Min = " & ph(lCol - 15, 1) & " - Max = " & ph(lCol - 15, 2), "#" & lRow & " - " & .Cell(0, lCol).Text & " - STD " & .Cell(lRow, 1).Text, , , , sString, , True) Then

                rc = True
            End If
        Case 34
        
            ' TEST
            Combo3.ZOrder
            Combo3.Visible = True
            Frame7.Visible = True
            rc = False
            Exit Sub
        Case Else
        
        
            sString = .Cell(lRow, lCol).Text
            If F_InputBox.DoShow("Enter value...", "#" & .Cell(lRow, 3).Text & " - " & .Cell(0, lCol).Text, , , , sString) Then
              
               
               .Cell(lRow, lCol).Text = sString
                rc = True
                GoTo SetFormatText
            End If
    
    End Select
    
    If rc Then
        Text1(Index) = ""
         Text1(Index) = sString
         Text1_LostFocus Index
         .Cell(lRow, lCol).Text = sString
         .Cell(lRow, 5).Text = IIf(.Cell(lRow, 5).Text = "", FormatDataLAT(date), .Cell(lRow, 5).Text)
         .Cell(lRow, 6).Text = IIf(.Cell(lRow, 6).Text = "", FormatTimeLAT(FormatDateTime(Now, vbShortTime)), .Cell(lRow, 6).Text)
SetFormatText:
        Select Case lCol
            Case 1
                .Cell(lRow, 2).Text = Text1(9)
                .Cell(lRow, 26).Text = Text1(9)
                .Cell(lRow, 27).Text = Text1(10)
                .Cell(lRow, 28).Text = Text1(11)
        
                STDSelectedNumber = Trim(.Cell(lRow, 1).Text)
                'Text1(33) = Trim(.Cell(lRow, 27).Text)
                'Text1(34) = Trim(.Cell(lRow, 28).Text)
                
                lbTestTable(35) = "STD" & .Cell(lRow, 1).Text & " Target Value"
                
                If Text1(9) = "" Then
                    ' devo inserire lo standard...
                    lCol = 2
                    GoTo start:
                End If
            Case 5, 7
                .Cell(lRow, lCol).Text = FormatDataLAT(sString)
                Exit Sub
            Case 11, 12, 13, 14
                .Cell(lRow, lCol).Text = Format$(sString, UserDecimal)
                Exit Sub
            Case Else
                Exit Sub
        End Select
    
    End If
    
    STDSelectedNumber = Trim(.Cell(lRow, 1).Text)
    .Cell(lRow, 27).Text = Text1(33)
    .Cell(lRow, 28).Text = Text1(34)
    
    lbTestTable(35) = "STD" & .Cell(lRow, 1).Text & " Target Value"
          
    
End With

    
     

End Sub


Private Function CheckSTD() As Boolean
Dim rc As Boolean
Dim Index As Integer
rc = True

If Grd2.Cell(lRow, 1).Text <> "" Then
    SetSTDbyTable
    SetSTDTextbyTable
    If CheckSTDInTable(Grd2.Cell(lRow, 1).Text, Index) Then
        Grd2.Cell(lRow, 2).Locked = True
        Grd2.Cell(lRow, 2).Text = STD(Index, 1)
        Grd2.Cell(lRow, 26).Text = STD(Index, 1)
        Grd2.Cell(lRow, 27).Text = STD(Index, 2)
        Grd2.Cell(lRow, 28).Text = STD(Index, 3)
        rc = False
    End If
End If

    CheckSTD = rc
End Function

Private Sub SetSTDTextbyTable()
With Grd2
    Text1(5) = .Cell(lRow, 1).Text
    Text1(9) = .Cell(lRow, 2).Text
    Text1(10) = .Cell(lRow, 27).Text
    Text1(11) = .Cell(lRow, 28).Text
    
    
    Text1(33) = Trim(.Cell(lRow, 27).Text)
    Text1(34) = Trim(.Cell(lRow, 28).Text)
    STDSelectedNumber = Trim(.Cell(lRow, 1).Text)
    .Cell(lRow, 27).Text = Text1(33)
    .Cell(lRow, 28).Text = Text1(34)
    
    lbTestTable(35) = "STD" & .Cell(lRow, 1).Text & " Target Value"
End With

End Sub

Private Function ReadOnlyGrid()
With Grd2
    .ReadOnly = True
    .SelectionMode = cellSelectionNone
     
End With
End Function

Private Function ActivateGrd2()

    With Grd2
       
        .ReadOnly = False
        .SelectionMode = cellSelectionNone
        .Column(2).Locked = True
        .Column(3).Locked = True
        Frame3.Visible = True
        Frame7.Visible = True
        FillAllGrid
        bEnterTable = True
    End With

End Function


Private Function DeleteLine()

If Grd2.Rows > 1 Then
    If F_MsgBox.DoShow("Delete Reading #" & Grd2.Rows - 1 & " ?") Then
    
        With Grd2
            .Cell(.Rows - 1, 1).SetFocus
            .AutoRedraw = False
            .ReadOnly = False
            .Selection.DeleteByRow
            .ReadOnly = True
            .Refresh
            .AutoRedraw = True
            bEnterTable = False
        End With
       ' PopupMessage 2, "Line correctly deleted..."
       
       Call SetSTDbyTable
       
    End If
End If
End Function


Private Function SaveLine()

    Call SaveAllTests
    Call SalvaFormSettingName
End Function



Private Sub SetSTDFromDatabase()
Dim mrc As Boolean
Dim AndOr As String
Dim NewTolerance As Double
Dim Restriction As Double

On Error GoTo ERR_SET:

        mrc = FillSTDToleranceGrid(Text1(1), Grd1, , ReadingCodeID, True)
        
        
            If Grd1.Cols > 1 Then
                Debug.Print Grd1.Cell(2, 5).Text
                Text1(6) = Grd1.Cell(2, 5).Text
                Text1(7) = Grd1.Cell(2, 1).Text
                AndOr = ""
                
                Select Case UCase(Trim(Grd1.Cell(2, 2).Text))
                    Case "&"
                        AndOr = Chr$(177)
                    Case UCase("or")
                        AndOr = Chr$(247)
                    Case Else
                    
                End Select
               
                Debug.Print Grd1.Cell(2, 3).Text
                Debug.Print Grd1.Cell(2, 4).Text
                
                NewTolerance = Replace(Grd1.Cell(2, 3).Text, "%", "")
                Restriction = Replace(Grd1.Cell(2, 4).Text, "%", "")
                
                
                If Grd1.Cell(2, 4).Text <> "/" And IsNumeric(Restriction) Then
                    NewTolerance = FormatNumber(NewTolerance * Restriction / 100, 2)
                    Debug.Print (Grd1.Cell(2, 4).Text)
                End If
                Text1(8) = AndOr & " " & NewTolerance
                
            If mrc Then
                Grd1.ZOrder
                Grd1.Visible = True
                lbStandard.Visible = True
    
            End If
        End If
ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_SET:
    MsgBox Err.Description
    Resume Next

End Sub



Private Sub Timer1_Timer()
Text1(17) = Replace(FormatDateTime(Now, vbShortTime), ".", ":")
End Sub



Private Sub TimerIntro_Timer()
   
     SetSTDFromDatabase
    GetFormSettingName
    
    
    
    SetCombo
    
    frLoading.Visible = False
    TimerIntro.Enabled = False
End Sub

Private Sub txSearch_Change()
SearchInTab
End Sub



'---------------------------------------
'
'           QR Code Reader
'
'---------------------------------------


Private Function ScanQRCodeQC() As Boolean

MessageInfoTime = 2000

txQRCode = ""
txQRCode.Top = 0
txQRCode.Visible = True
txQRCode.SetFocus

UserQrCode = iQRCodeTypeClean

PopupMessage 2, "Scan QRCode....", , , "Preparation | QC"


End Function




Private Sub txQRCode_KeyPress(KeyAscii As Integer)
Dim rc As Boolean
Dim STD_ID As String

On Error GoTo ERR_QR:



If KeyAscii = 13 Then


    DoEvents
    If txQRCode = "" Then Exit Sub
   
    If GetSTDQRCodeFromString(Trim(txQRCode), UserQrCode) Then
    
     
        If UserQrCode.STDn = "" Or UserQrCode.STDv = "" Then
            PopupMessage 2, "Please Rescan QRCode...", , , "STD QRCode Reader"
            txQRCode = ""
            Exit Sub
        End If
        With UserQrCode
            If UCase(Trim(.Code)) = UCase(Trim(Text1(1))) Then
                 
                  PopupMessage 2, "STD acquired :" & .STDn & vbCrLf & "STD Value = " & .STDv & vbCrLf & "STD ID = " & .STD_ID, , , "STD QRCode Reader"
           
                 Text1(5) = .STDn
                 Text1(9) = .STDv
                 STD_ID = .STD_ID
                 Text1(67) = STD_ID
                 
                 'Text1_KeyPress(5,13
                  Text1_LostFocus 9
                  
                
            Else
                MessageInfoTime = 3500
                PopupMessage 2, "ERROR : Please check STD Label" & vbCrLf & "STD Hanna Code = " & .Code & vbCrLf & "This Lot Code = " & Text1(1), , , "STD QRCode Reader"
            End If
        End With

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


Private Function SetClosedLotReadings(ByVal rc As Boolean, ByVal strDate As String)


    PicMain(0).Enabled = Not (rc)
    PicMain(1).Enabled = Not (rc)
    If rc Then
        blTable = "Closed Lot : " & strDate
    Else
        blTable = "Reading QC"
    
    End If
    Image2(0).Visible = Not (rc)
    Image2(1).Visible = Not (rc)
    Image2(2).Visible = Not (rc)
    Grd2.ReadOnly = rc



End Function

Private Sub CheckClosedLot()

If MyOperatore.IndexPrivilege >= 1 Then
    
        Call SetClosedLotReadings(False, strValidDate)
       
    Else
        Call SetClosedLotReadings(bClosedLot, strValidDate)
    End If
    

End Sub
