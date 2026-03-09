VERSION 5.00
Begin VB.Form F_SETTING 
   BackColor       =   &H004D3B37&
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
   ForeColor       =   &H00404040&
   Icon            =   "F_SETTING.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12000
   ScaleMode       =   0  'User
   ScaleWidth      =   19200
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   8655
      Index           =   3
      Left            =   120
      ScaleHeight     =   8655
      ScaleWidth      =   19215
      TabIndex        =   42
      Top             =   1440
      Visible         =   0   'False
      Width           =   19215
      Begin VB.ComboBox cmbLine 
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
         Left            =   2280
         TabIndex        =   122
         Text            =   "Combo1"
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame frCommandInside 
         BackColor       =   &H00008000&
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
         Index           =   4
         Left            =   10320
         TabIndex        =   99
         Top             =   7800
         Width           =   4455
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Save"
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
            TabIndex        =   100
            Top             =   120
            Width           =   4455
         End
      End
      Begin VB.TextBox TxSocieta 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   510
         Index           =   0
         Left            =   4440
         TabIndex        =   51
         Top             =   1200
         Width           =   10335
      End
      Begin VB.TextBox TxSocieta 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   510
         Index           =   1
         Left            =   4440
         TabIndex        =   50
         Top             =   2520
         Width           =   10335
      End
      Begin VB.TextBox TxSocieta 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   510
         Index           =   2
         Left            =   4440
         TabIndex        =   49
         Top             =   3840
         Width           =   10335
      End
      Begin VB.TextBox TxSocieta 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   510
         Index           =   3
         Left            =   4440
         TabIndex        =   48
         Top             =   5160
         Width           =   10335
      End
      Begin VB.TextBox TxSocieta 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   510
         Index           =   4
         Left            =   4440
         TabIndex        =   47
         Top             =   6480
         Width           =   10335
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   3
         Left            =   16440
         MouseIcon       =   "F_SETTING.frx":33E2
         MousePointer    =   99  'Custom
         ScaleHeight     =   1335
         ScaleWidth      =   9615
         TabIndex        =   43
         Top             =   3120
         Visible         =   0   'False
         Width           =   9615
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Login"
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
            Left            =   4560
            MouseIcon       =   "F_SETTING.frx":36EC
            MousePointer    =   99  'Custom
            TabIndex        =   44
            Top             =   840
            Width           =   465
         End
         Begin VB.Image Image4 
            Height          =   480
            Index           =   3
            Left            =   4560
            Picture         =   "F_SETTING.frx":39F6
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Label LbSocieta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line Leader"
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
         Index           =   4
         Left            =   4440
         TabIndex        =   56
         Top             =   4680
         Width           =   1320
      End
      Begin VB.Label LbSocieta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Workstation"
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
         Height          =   360
         Index           =   2
         Left            =   4440
         TabIndex        =   55
         Top             =   3360
         Width           =   1740
      End
      Begin VB.Label LbSocieta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   360
         Index           =   1
         Left            =   4440
         TabIndex        =   54
         Top             =   2040
         Width           =   1680
      End
      Begin VB.Label LbSocieta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department ( Line )"
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
         Index           =   0
         Left            =   4440
         TabIndex        =   53
         Top             =   720
         Width           =   2280
      End
      Begin VB.Label LbSocieta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "e-mail"
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
         Index           =   5
         Left            =   4440
         TabIndex        =   52
         Top             =   6000
         Width           =   705
      End
      Begin VB.Image Image3 
         Height          =   480
         Index           =   8
         Left            =   9360
         Picture         =   "F_SETTING.frx":6DD8
         Top             =   360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   0
         Left            =   13800
         Picture         =   "F_SETTING.frx":A1BA
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   4
         Left            =   120
         TabIndex        =   45
         Top             =   7440
         Width           =   2655
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   8655
      Index           =   4
      Left            =   0
      ScaleHeight     =   8655
      ScaleWidth      =   19215
      TabIndex        =   89
      Top             =   1440
      Width           =   19215
      Begin VB.OptionButton Option2 
         BackColor       =   &H00F0F0F0&
         Caption         =   "All numbers"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   405
         Index           =   0
         Left            =   8040
         TabIndex        =   133
         Top             =   3480
         Value           =   -1  'True
         Width           =   7095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00F0F0F0&
         Caption         =   "Only Odd numbers"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   8040
         TabIndex        =   132
         Top             =   3120
         Width           =   7095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00F0F0F0&
         Caption         =   "Only Even numbers"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   8040
         TabIndex        =   131
         Top             =   2760
         Width           =   7095
      End
      Begin VB.Frame frMethod 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Enabled         =   0   'False
         Height          =   975
         Left            =   960
         TabIndex        =   124
         Top             =   6840
         Width           =   6135
         Begin VB.OptionButton Option1 
            BackColor       =   &H00F0F0F0&
            Caption         =   "Chemical Production"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   0
            TabIndex        =   126
            Top             =   0
            Width           =   7095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00F0F0F0&
            Caption         =   "Chemical CLP Classification Software for Production"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   0
            TabIndex        =   125
            Top             =   360
            Width           =   7095
         End
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00F0F0F0&
         Caption         =   "Preparation : Open product classification after scan"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   405
         Left            =   960
         TabIndex        =   118
         Top             =   2280
         Width           =   6855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00F0F0F0&
         Caption         =   "Use . for Decimals "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   405
         Left            =   960
         TabIndex        =   115
         Top             =   3840
         Width           =   5295
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   4
         Left            =   10080
         MouseIcon       =   "F_SETTING.frx":D59C
         MousePointer    =   99  'Custom
         ScaleHeight     =   1335
         ScaleWidth      =   9615
         TabIndex        =   92
         Top             =   3000
         Visible         =   0   'False
         Width           =   9615
         Begin VB.Image Image4 
            Height          =   480
            Index           =   4
            Left            =   4560
            Picture         =   "F_SETTING.frx":D8A6
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Login"
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
            Index           =   14
            Left            =   4560
            MouseIcon       =   "F_SETTING.frx":10C88
            MousePointer    =   99  'Custom
            TabIndex        =   93
            Top             =   840
            Width           =   465
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00F0F0F0&
         Caption         =   "View : Full Screen Mode"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   405
         Left            =   960
         TabIndex        =   90
         Top             =   5400
         Width           =   9255
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   4800
         ScaleHeight     =   405
         ScaleWidth      =   375
         TabIndex        =   91
         Top             =   7920
         Width           =   375
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Lot Number"
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
         Height          =   315
         Left            =   8040
         TabIndex        =   130
         Top             =   2160
         Width           =   1800
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00C0C0C0&
         X1              =   8040
         X2              =   13800
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label Printer"
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
         Height          =   315
         Left            =   8040
         TabIndex        =   129
         Top             =   840
         Width           =   1425
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00C0C0C0&
         X1              =   8040
         X2              =   13800
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SetLabel Printer"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   255
         Index           =   7
         Left            =   8040
         MouseIcon       =   "F_SETTING.frx":10F92
         MousePointer    =   99  'Custom
         TabIndex        =   128
         Top             =   1440
         Width           =   2790
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   11400
         Picture         =   "F_SETTING.frx":1129C
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Index           =   3
         Left            =   9360
         Picture         =   "F_SETTING.frx":127B6
         Top             =   7680
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Administrator User Required : Enter UserID and Password"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   285
         Index           =   3
         Left            =   6720
         TabIndex        =   127
         Top             =   8280
         Width           =   5580
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00C0C0C0&
         X1              =   960
         X2              =   6720
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Software Method"
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
         Height          =   315
         Left            =   960
         TabIndex        =   123
         Top             =   6240
         Width           =   2055
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         X1              =   960
         X2              =   6720
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode Reader"
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
         Height          =   315
         Left            =   960
         TabIndex        =   119
         Top             =   840
         Width           =   1905
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monitor Settings"
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
         Height          =   315
         Left            =   960
         TabIndex        =   117
         Top             =   4920
         Width           =   1890
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         X1              =   960
         X2              =   6720
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   960
         X2              =   6720
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "International Settings"
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
         Height          =   315
         Left            =   960
         TabIndex        =   116
         Top             =   3360
         Width           =   2520
      End
      Begin VB.Label QrCodeLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Separator Character"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   300
         Index           =   1
         Left            =   960
         MouseIcon       =   "F_SETTING.frx":15B98
         MousePointer    =   99  'Custom
         TabIndex        =   114
         Top             =   1800
         Width           =   6495
      End
      Begin VB.Label QrCodeLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test Scans"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   300
         Index           =   0
         Left            =   960
         MouseIcon       =   "F_SETTING.frx":15EA2
         MousePointer    =   99  'Custom
         TabIndex        =   113
         Top             =   1440
         Width           =   1080
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   8655
      Index           =   1
      Left            =   240
      ScaleHeight     =   8655
      ScaleWidth      =   19215
      TabIndex        =   34
      Top             =   1800
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Frame frDatabaseType 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6165
         Left            =   9240
         TabIndex        =   101
         Top             =   600
         Visible         =   0   'False
         Width           =   10995
         Begin VB.Frame Frame3 
            BackColor       =   &H00644603&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
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
            Left            =   0
            TabIndex        =   102
            Top             =   0
            Width           =   11055
            Begin VB.Label lbInside 
               BackStyle       =   0  'Transparent
               Caption         =   "Database Table"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   1
               Left            =   240
               TabIndex        =   103
               Top             =   120
               Width           =   2895
               WordWrap        =   -1  'True
            End
            Begin VB.Image Image2 
               Height          =   240
               Left            =   10560
               Picture         =   "F_SETTING.frx":161AC
               Top             =   120
               Width           =   240
            End
         End
         Begin VB.Label lbData 
            BackStyle       =   0  'Transparent
            Caption         =   "Select form list to Edit Database Records"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Left            =   240
            TabIndex        =   110
            Top             =   5640
            Width           =   10335
         End
         Begin VB.Label lbDatabase 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chemical RM"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00964901&
            Height          =   450
            Index           =   5
            Left            =   0
            MouseIcon       =   "F_SETTING.frx":16BAE
            MousePointer    =   99  'Custom
            TabIndex        =   109
            Top             =   1680
            Width           =   10995
         End
         Begin VB.Label lbDatabase 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Formulations"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00964901&
            Height          =   450
            Index           =   4
            Left            =   0
            MouseIcon       =   "F_SETTING.frx":16EB8
            MousePointer    =   99  'Custom
            TabIndex        =   108
            Top             =   960
            Width           =   10935
         End
         Begin VB.Label lbDatabase 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Physical hazard Statement"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00964901&
            Height          =   450
            Index           =   3
            Left            =   0
            MouseIcon       =   "F_SETTING.frx":171C2
            MousePointer    =   99  'Custom
            TabIndex        =   107
            Top             =   4560
            Width           =   10965
         End
         Begin VB.Label lbDatabase 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Classification"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00964901&
            Height          =   450
            Index           =   2
            Left            =   3660
            MouseIcon       =   "F_SETTING.frx":174CC
            MousePointer    =   99  'Custom
            TabIndex        =   106
            Top             =   3840
            Width           =   3870
         End
         Begin VB.Label lbDatabase 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Machine"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00964901&
            Height          =   450
            Index           =   1
            Left            =   4800
            MouseIcon       =   "F_SETTING.frx":177D6
            MousePointer    =   99  'Custom
            TabIndex        =   105
            Top             =   3120
            Width           =   1635
         End
         Begin VB.Label lbDatabase 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hanna Code"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00964901&
            Height          =   450
            Index           =   0
            Left            =   0
            MouseIcon       =   "F_SETTING.frx":17AE0
            MousePointer    =   99  'Custom
            TabIndex        =   104
            Top             =   2400
            Width           =   11055
         End
         Begin VB.Shape shInside 
            BackColor       =   &H00808080&
            BorderColor     =   &H00808080&
            Height          =   6165
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   10995
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   1
         Left            =   17280
         MouseIcon       =   "F_SETTING.frx":17DEA
         MousePointer    =   99  'Custom
         ScaleHeight     =   1335
         ScaleWidth      =   9615
         TabIndex        =   35
         Top             =   6360
         Visible         =   0   'False
         Width           =   9615
         Begin VB.Image Image4 
            Height          =   480
            Index           =   1
            Left            =   4560
            Picture         =   "F_SETTING.frx":180F4
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Login"
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
            Left            =   4560
            MouseIcon       =   "F_SETTING.frx":1B4D6
            MousePointer    =   99  'Custom
            TabIndex        =   36
            Top             =   840
            Width           =   465
         End
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-------"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   225
         Index           =   8
         Left            =   11760
         TabIndex        =   136
         Top             =   1760
         Width           =   400
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   7
         Left            =   2760
         MouseIcon       =   "F_SETTING.frx":1B7E0
         MousePointer    =   99  'Custom
         TabIndex        =   134
         Top             =   3960
         Width           =   6135
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   11
         Left            =   7695
         Picture         =   "F_SETTING.frx":1BAEA
         Top             =   4320
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Export Hanna Code Table to dbCode Final"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   255
         Index           =   8
         Left            =   3330
         MouseIcon       =   "F_SETTING.frx":1EECC
         MousePointer    =   99  'Custom
         TabIndex        =   135
         Top             =   4440
         Width           =   4140
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   6
         Left            =   4200
         MouseIcon       =   "F_SETTING.frx":1F1D6
         MousePointer    =   99  'Custom
         TabIndex        =   120
         Top             =   5280
         Width           =   4695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recovery Database form File"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   255
         Index           =   6
         Left            =   4695
         MouseIcon       =   "F_SETTING.frx":1F4E0
         MousePointer    =   99  'Custom
         TabIndex        =   121
         Top             =   5760
         Width           =   2775
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   10
         Left            =   7695
         Picture         =   "F_SETTING.frx":1F7EA
         Top             =   5640
         Width           =   480
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--------"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   225
         Index           =   7
         Left            =   6960
         TabIndex        =   112
         Top             =   3600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--------"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   225
         Index           =   6
         Left            =   6960
         TabIndex        =   111
         Top             =   3300
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--------"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   225
         Index           =   5
         Left            =   6975
         TabIndex        =   82
         Top             =   3000
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   225
         Index           =   4
         Left            =   11760
         TabIndex        =   81
         Top             =   6120
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   225
         Index           =   3
         Left            =   11760
         TabIndex        =   80
         Top             =   1440
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   225
         Index           =   2
         Left            =   11760
         TabIndex        =   79
         Top             =   4560
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--------"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   225
         Index           =   1
         Left            =   6975
         TabIndex        =   78
         Top             =   1740
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--------"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   225
         Index           =   0
         Left            =   6975
         TabIndex        =   77
         Top             =   1440
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   3
         Left            =   10680
         MouseIcon       =   "F_SETTING.frx":22BCC
         MousePointer    =   99  'Custom
         TabIndex        =   69
         Top             =   5280
         Width           =   4695
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1455
         Index           =   2
         Left            =   10320
         MouseIcon       =   "F_SETTING.frx":22ED6
         MousePointer    =   99  'Custom
         TabIndex        =   68
         Top             =   3720
         Width           =   6015
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1095
         Index           =   1
         Left            =   10440
         MouseIcon       =   "F_SETTING.frx":231E0
         MousePointer    =   99  'Custom
         TabIndex        =   67
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1695
         Index           =   0
         Left            =   10440
         MouseIcon       =   "F_SETTING.frx":234EA
         MousePointer    =   99  'Custom
         TabIndex        =   63
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1695
         Index           =   5
         Left            =   3960
         MouseIcon       =   "F_SETTING.frx":237F4
         MousePointer    =   99  'Custom
         TabIndex        =   72
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1575
         Index           =   4
         Left            =   4320
         MouseIcon       =   "F_SETTING.frx":23AFE
         MousePointer    =   99  'Custom
         TabIndex        =   70
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratoy Manager or Administrator Login Required"
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
         Index           =   1
         Left            =   7320
         TabIndex        =   75
         Top             =   7440
         Width           =   4785
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   8
         Left            =   7680
         Picture         =   "F_SETTING.frx":23E08
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View / Edit Database"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   255
         Index           =   5
         Left            =   5385
         MouseIcon       =   "F_SETTING.frx":271EA
         MousePointer    =   99  'Custom
         TabIndex        =   73
         Top             =   2640
         Width           =   2070
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   7
         Left            =   7680
         Picture         =   "F_SETTING.frx":274F4
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Import  Excel"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   255
         Index           =   4
         Left            =   6120
         MouseIcon       =   "F_SETTING.frx":2A8D6
         MousePointer    =   99  'Custom
         TabIndex        =   71
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Database Path "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   255
         Index           =   3
         Left            =   11760
         MouseIcon       =   "F_SETTING.frx":2ABE0
         MousePointer    =   99  'Custom
         TabIndex        =   66
         Top             =   5760
         Width           =   2160
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   6
         Left            =   11160
         Picture         =   "F_SETTING.frx":2AEEA
         Top             =   5640
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reset Default Database ( Empty )"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   255
         Index           =   2
         Left            =   11760
         MouseIcon       =   "F_SETTING.frx":2E2CC
         MousePointer    =   99  'Custom
         TabIndex        =   65
         Top             =   4200
         Width           =   3210
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   5
         Left            =   11160
         Picture         =   "F_SETTING.frx":2E5D6
         Top             =   4080
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copy / Backup Database"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   255
         Index           =   1
         Left            =   11760
         MouseIcon       =   "F_SETTING.frx":319B8
         MousePointer    =   99  'Custom
         TabIndex        =   64
         Top             =   2640
         Width           =   2475
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   2
         Left            =   11160
         Picture         =   "F_SETTING.frx":31CC2
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Import Database"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   255
         Index           =   0
         Left            =   11760
         MouseIcon       =   "F_SETTING.frx":350A4
         MousePointer    =   99  'Custom
         TabIndex        =   62
         Top             =   1080
         Width           =   1665
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   1
         Left            =   11160
         Picture         =   "F_SETTING.frx":353AE
         Top             =   960
         Width           =   480
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   9
         Left            =   120
         TabIndex        =   37
         Top             =   7440
         Width           =   2655
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Index           =   1
         Left            =   9360
         MouseIcon       =   "F_SETTING.frx":38790
         MousePointer    =   99  'Custom
         Picture         =   "F_SETTING.frx":38A9A
         Top             =   6840
         Width           =   480
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   9135
      Index           =   0
      Left            =   0
      ScaleHeight     =   9135
      ScaleWidth      =   19215
      TabIndex        =   18
      Top             =   1320
      Visible         =   0   'False
      Width           =   19215
      Begin VB.TextBox Text1 
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
         ForeColor       =   &H00404040&
         Height          =   495
         Index           =   0
         Left            =   7560
         TabIndex        =   26
         Top             =   3480
         Width           =   4215
      End
      Begin VB.TextBox Text1 
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
         ForeColor       =   &H00404040&
         Height          =   495
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   7560
         PasswordChar    =   "*"
         TabIndex        =   20
         Top             =   4560
         Width           =   4215
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   0
         Left            =   4920
         MouseIcon       =   "F_SETTING.frx":3BE7C
         MousePointer    =   99  'Custom
         ScaleHeight     =   1335
         ScaleWidth      =   9615
         TabIndex        =   24
         Top             =   5760
         Width           =   9615
         Begin VB.Label Lab 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Login"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   4
            Left            =   0
            MouseIcon       =   "F_SETTING.frx":3C186
            MousePointer    =   99  'Custom
            TabIndex        =   25
            Top             =   840
            Width           =   9585
         End
         Begin VB.Image Image4 
            Height          =   480
            Index           =   0
            Left            =   4560
            Picture         =   "F_SETTING.frx":3C490
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   12120
         Picture         =   "F_SETTING.frx":3F872
         Top             =   3480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   735
         Index           =   6
         Left            =   4440
         MouseIcon       =   "F_SETTING.frx":42C54
         MousePointer    =   99  'Custom
         TabIndex        =   97
         Top             =   1680
         Width           =   5655
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   9
         Left            =   9360
         Picture         =   "F_SETTING.frx":42F5E
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratory Manager : Gibertini Riccardo"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   255
         Left            =   5160
         MouseIcon       =   "F_SETTING.frx":46340
         MousePointer    =   99  'Custom
         TabIndex        =   98
         Top             =   1920
         Width           =   3900
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   735
         Index           =   5
         Left            =   4200
         MouseIcon       =   "F_SETTING.frx":4664A
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   960
         Width           =   5775
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1095
         Index           =   7
         Left            =   4440
         MouseIcon       =   "F_SETTING.frx":46954
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   0
         Width           =   5655
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter UserID and Password"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   285
         Index           =   0
         Left            =   8280
         TabIndex        =   76
         Top             =   6480
         Width           =   2640
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   4
         Left            =   9360
         Picture         =   "F_SETTING.frx":46C5E
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   3
         Left            =   9360
         Picture         =   "F_SETTING.frx":4A040
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratory Manager : Gibertini Riccardo"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   255
         Left            =   5160
         MouseIcon       =   "F_SETTING.frx":4D422
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   480
         Width           =   3900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   360
         Index           =   1
         Left            =   7560
         TabIndex        =   27
         Tag             =   "User"
         Top             =   3000
         Width           =   1545
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Index           =   0
         Left            =   9360
         Picture         =   "F_SETTING.frx":4D72C
         Top             =   6000
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Administrator : Gibertini Riccardo"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   255
         Left            =   5895
         MouseIcon       =   "F_SETTING.frx":50B0E
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   1200
         Width           =   3210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   360
         Index           =   0
         Left            =   7560
         TabIndex        =   19
         Top             =   4080
         Width           =   1410
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   8
         Left            =   120
         TabIndex        =   30
         Top             =   7440
         Width           =   2655
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   8655
      Index           =   2
      Left            =   0
      ScaleHeight     =   8655
      ScaleWidth      =   19215
      TabIndex        =   38
      Top             =   1560
      Visible         =   0   'False
      Width           =   19215
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   2
         Left            =   4800
         MouseIcon       =   "F_SETTING.frx":50E18
         MousePointer    =   99  'Custom
         ScaleHeight     =   1335
         ScaleWidth      =   9615
         TabIndex        =   39
         Top             =   6480
         Visible         =   0   'False
         Width           =   9615
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Login"
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
            Left            =   4560
            MouseIcon       =   "F_SETTING.frx":51122
            MousePointer    =   99  'Custom
            TabIndex        =   40
            Top             =   840
            Width           =   465
         End
         Begin VB.Image Image4 
            Height          =   480
            Index           =   2
            Left            =   4560
            Picture         =   "F_SETTING.frx":5142C
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00644603&
         Caption         =   "Software License"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   8
         Left            =   0
         MouseIcon       =   "F_SETTING.frx":5480E
         MousePointer    =   99  'Custom
         TabIndex        =   96
         Top             =   0
         Width           =   3255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratoy Manager or Administrator Login Required"
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
         Index           =   2
         Left            =   7200
         TabIndex        =   74
         Top             =   6000
         Visible         =   0   'False
         Width           =   4785
      End
      Begin VB.Label lbStr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   0
         Left            =   3120
         TabIndex        =   60
         Top             =   3240
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   6
         Left            =   2040
         TabIndex        =   59
         Top             =   3240
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   570
         Left            =   10110
         TabIndex        =   58
         Top             =   2520
         Width           =   165
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   2055
         Index           =   3
         Left            =   1800
         Top             =   2520
         Width           =   15615
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   6
         Left            =   16560
         MouseIcon       =   "F_SETTING.frx":54B18
         MousePointer    =   99  'Custom
         Picture         =   "F_SETTING.frx":54E22
         Top             =   2760
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00644603&
         Caption         =   "Search Online Update"
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
         Height          =   495
         Index           =   5
         Left            =   2040
         MouseIcon       =   "F_SETTING.frx":58204
         MousePointer    =   99  'Custom
         TabIndex        =   57
         Top             =   3960
         Width           =   15135
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Index           =   2
         Left            =   9360
         MouseIcon       =   "F_SETTING.frx":5850E
         MousePointer    =   99  'Custom
         Picture         =   "F_SETTING.frx":58818
         Top             =   5400
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   7440
         Width           =   2655
      End
      Begin VB.Label lbMenuIntro 
         BackStyle       =   0  'Transparent
         Caption         =   "  SOFTWARE RELEASE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   2025
         Index           =   3
         Left            =   1800
         MouseIcon       =   "F_SETTING.frx":5BBFA
         MousePointer    =   99  'Custom
         TabIndex        =   61
         Top             =   2520
         Width           =   12615
      End
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   4
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   19215
      TabIndex        =   94
      Top             =   6720
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options : You can select or change options in this page..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   7
         Left            =   480
         TabIndex        =   95
         Top             =   120
         Width           =   5535
      End
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   3
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   19215
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratory Info / Hanna Department : fill Lab form and Save"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   3
         Left            =   480
         TabIndex        =   17
         Top             =   120
         Width           =   5805
      End
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H0070B0F0&
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   2
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   19215
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Update Program : Online update, check new releases..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   16
         Top             =   120
         Width           =   5340
      End
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H008FC9FA&
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   1
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   19215
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Settings : Copy, Import & Export Databases"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   15
         Top             =   120
         Width           =   5010
      End
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H00A0DFA0&
      BorderStyle     =   0  'None
      DrawWidth       =   7
      Height          =   615
      Index           =   0
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   19215
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Account Manager : Select Administrator, Laboratory Manager & QC Operator "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   14
         Top             =   120
         Width           =   7860
      End
   End
   Begin VB.PictureBox PicMainMenu 
      BackColor       =   &H00473733&
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
      Height          =   975
      Index           =   4
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   0
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   4
         Left            =   7680
         MouseIcon       =   "F_SETTING.frx":5BF04
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   83
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Options..."
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
            MouseIcon       =   "F_SETTING.frx":5C20E
            MousePointer    =   99  'Custom
            TabIndex        =   84
            Top             =   660
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   4
            Left            =   720
            MouseIcon       =   "F_SETTING.frx":5C518
            MousePointer    =   99  'Custom
            OLEDropMode     =   1  'Manual
            Picture         =   "F_SETTING.frx":5C822
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   3
         Left            =   5760
         MouseIcon       =   "F_SETTING.frx":5FC04
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   8
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   735
            MouseIcon       =   "F_SETTING.frx":5FF0E
            MousePointer    =   99  'Custom
            OLEDropMode     =   1  'Manual
            Picture         =   "F_SETTING.frx":60218
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Department Info"
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
            MouseIcon       =   "F_SETTING.frx":635FA
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   660
            Width           =   1890
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   3840
         MouseIcon       =   "F_SETTING.frx":63904
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
            MouseIcon       =   "F_SETTING.frx":63C0E
            MousePointer    =   99  'Custom
            Picture         =   "F_SETTING.frx":63F18
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Update"
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
            MouseIcon       =   "F_SETTING.frx":672FA
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   660
            Width           =   1890
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   1920
         MouseIcon       =   "F_SETTING.frx":67604
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
            MouseIcon       =   "F_SETTING.frx":6790E
            MousePointer    =   99  'Custom
            Picture         =   "F_SETTING.frx":67C18
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Database"
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
            Left            =   90
            MouseIcon       =   "F_SETTING.frx":6AFFA
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   660
            Width           =   1830
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00473733&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MouseIcon       =   "F_SETTING.frx":6B304
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
            Caption         =   "User Account"
            Enabled         =   0   'False
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
            MouseIcon       =   "F_SETTING.frx":6B60E
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   660
            Width           =   1080
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "F_SETTING.frx":6B918
            MousePointer    =   99  'Custom
            Picture         =   "F_SETTING.frx":6BC22
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Settings"
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
         Height          =   570
         Left            =   11760
         TabIndex        =   22
         Top             =   200
         Visible         =   0   'False
         Width           =   6930
      End
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   3
      Left            =   0
      MouseIcon       =   "F_SETTING.frx":6F004
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   10680
      Width           =   1815
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   1
      Left            =   8400
      MouseIcon       =   "F_SETTING.frx":6F30E
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   10560
      Width           =   2655
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   10
      Left            =   17640
      MouseIcon       =   "F_SETTING.frx":6F618
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Top             =   10560
      Width           =   1575
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   2
      Left            =   15240
      MouseIcon       =   "F_SETTING.frx":6F922
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   10680
      Width           =   2175
   End
   Begin VB.Label Label2 
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
      Left            =   17775
      MouseIcon       =   "F_SETTING.frx":6FC2C
      MousePointer    =   99  'Custom
      TabIndex        =   88
      Top             =   11715
      Width           =   1200
   End
   Begin VB.Label Label2 
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
      Left            =   15735
      MouseIcon       =   "F_SETTING.frx":6FF36
      MousePointer    =   99  'Custom
      TabIndex        =   87
      Top             =   11715
      Width           =   1230
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Settings"
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
      Index           =   10
      Left            =   9090
      MouseIcon       =   "F_SETTING.frx":70240
      MousePointer    =   99  'Custom
      TabIndex        =   86
      Top             =   11715
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select User"
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
      Left            =   360
      MouseIcon       =   "F_SETTING.frx":7054A
      MousePointer    =   99  'Custom
      TabIndex        =   85
      Top             =   11715
      Width           =   900
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_SETTING.frx":70854
      Height          =   480
      Index           =   0
      Left            =   18240
      MouseIcon       =   "F_SETTING.frx":73C36
      MousePointer    =   99  'Custom
      Picture         =   "F_SETTING.frx":73F40
      Top             =   11160
      Width           =   480
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
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   4800
      X2              =   4800
      Y1              =   120
      Y2              =   12000
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
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   480
      MouseIcon       =   "F_SETTING.frx":77322
      MousePointer    =   99  'Custom
      Picture         =   "F_SETTING.frx":7762C
      Top             =   11160
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_SETTING.frx":7AA0E
      Height          =   480
      Index           =   1
      Left            =   9360
      MouseIcon       =   "F_SETTING.frx":7DDF0
      MousePointer    =   99  'Custom
      Picture         =   "F_SETTING.frx":7E0FA
      Top             =   11160
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   15960
      MouseIcon       =   "F_SETTING.frx":814DC
      MousePointer    =   99  'Custom
      Picture         =   "F_SETTING.frx":817E6
      Top             =   11160
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   4
      Left            =   2040
      MouseIcon       =   "F_SETTING.frx":84BC8
      MousePointer    =   99  'Custom
      Picture         =   "F_SETTING.frx":84ED2
      Top             =   11160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   0
      X2              =   19200
      Y1              =   11040
      Y2              =   11040
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
End
Attribute VB_Name = "F_SETTING"
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

Private m_rc As Boolean
Private bFormSaved As Boolean
Private bSonoAutorizzato As Boolean
Private DatabaseIndex As Integer

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
        ElseIf TypeOf ctl Is Inet Then
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



For Each ctl In Controls
    With m_ControlPositions(i)
        If TypeOf ctl Is Line Then
            ctl.x1 = x_scale * .Left
            ctl.y1 = y_scale * .Top
            ctl.x2 = ctl.x1 + x_scale * .Width
            ctl.y2 = ctl.y1 + y_scale * .Height
        ElseIf TypeOf ctl Is Timer Then
        ElseIf TypeOf ctl Is Inet Then
        ElseIf TypeOf ctl Is Image Then
            ctl.Left = (x_scale * .Left) + IIf(x_scale = 1, 0, (x_scale - 1) * 200)
            ctl.Top = y_scale * .Top
        ElseIf TypeOf ctl Is Grid Then
           ctl.Left = x_scale * .Left
            ctl.Top = y_scale * .Top
            ctl.Width = x_scale * .Width
            ctl.Height = y_scale * .Height

             
           
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
Public Function DoShow(Optional ByRef Index As Integer, Optional ByRef sLot As String, Optional ByRef sCode As String, Optional MyImage As Image) As Boolean

    On Error GoTo ERR_SHOW
    
    GetOptions
    
    
    CheckRegistration
    Set DefaultMenu(4) = MyImage
    bSonoAutorizzato = False
    m_rc = False
    bFormSaved = False
    SelectProcedura Index
    
        Image5.Visible = bStampaOk
    
    
    mOk

    
    Me.Show
    
    If m_rc = True Then

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

Private Sub Check1_Click()
Dim rc As Boolean
rc = IIf(Check1.Value = 1, True, False)
bFullScreen = rc
SaveSetting App.Title, "Opzioni", "Full Screen Mode", bFullScreen
If Me.Visible Then PopupMessage 2, "Please Restart the Program to take affect changes..."
End Sub

Private Sub Check2_Click()
Dim rc As Boolean
rc = IIf(Check2.Value = 1, True, False)
bDotForDecimals = rc

Check2.ForeColor = IIf(rc, vbColorTextDarkBlue, vbBlack)
Check2.FontBold = rc


SaveSetting App.Title, "Notation", "bDotForDecimals", bDotForDecimals
End Sub

Private Sub Check3_Click()
Dim rc As Boolean
rc = IIf(Check3.Value = 1, True, False)
bOpenProductClassificationAfterScan = rc
SaveSetting App.Title, "BarcodeReader", "bOpenProductClassificationAfterScan", bOpenProductClassificationAfterScan
End Sub

Private Sub cmbLine_Click()
If Me.Visible Then
    If cmbLine <> "" Then
        TxSocieta(0) = cmbLine
        cmbLine.Visible = False
        Call SetUserLine(cmbLine, cmbLine.ListIndex)
    End If
End If
End Sub

Private Sub DefaultMenu_Click(Index As Integer)
DefaultMenuLabel_Click Index
End Sub

Private Sub DefaultMenuLabel_Click(Index As Integer)
Dim MyIndex As Integer
Dim MyName As String
Select Case Index

    Case 1
        Unload Me
    Case 2
        ' torna indietro
        If IndexProcedura = 0 Then
            MyIndex = PicMenu.Count - 1
        Else
            MyIndex = IndexProcedura - 1
        End If
        PicMenu_Click MyIndex
    Case 3
        If bExistAccount Then
            frmLogin.DoShow
            
        Else
            SelectProcedura (0)
            Picture4_Click (0)
        End If
    Case 5
        ' admin
        Call SplitName(Label7, MyName)
        Text1(0) = MyName
        Text1(1).SetFocus
     Case 6
        ' LabManager
        Call SplitName(Label10, MyName)
        Text1(0) = MyName
        Text1(1).SetFocus
        
    Case 7
        ' admin
        Call SplitName(Label8, MyName)
        Text1(0) = MyName
        Text1(1).SetFocus
        
    Case 10
        ' vai avanti
        If IndexProcedura = PicMenu.Count - 1 Then
            MyIndex = 0
        Else
            MyIndex = IndexProcedura + 1
        End If
        PicMenu_Click MyIndex
End Select

End Sub


Private Sub DisableImage_Click(Index As Integer)
Dim i As Integer
Dim sString As String


Select Case Index
    Case 0
        i = 1
        sString = "Only Line Leader ot TCO can Operate"

    Case 3
        i = 3
        sString = "Only Administrator can Operate"

End Select


PopupMessage 2, "Warning : " & sString & vbCrLf & "Please Login...", , True
CheckPrivilege i
Call CheckUser(Index, i)

End Sub

Private Sub Form_Initialize()

Call StartProcedure
SaveSizes
Me.WindowState = MainWindowState
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27
         DefaultMenuLabel_Click 1
    Case 37
        DefaultMenuLabel_Click 2
    Case 39
        DefaultMenuLabel_Click 0
End Select
End Sub
Private Sub Form_Load()
 Call SetPicForm
IndexProcedura = 99


If bFullScreen Then
    Me.WindowState = 2
Else
    Me.WindowState = 0
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)

Set F_SETTING = Nothing

End Sub

Private Sub frCommandInside_Click(Index As Integer)
    Select Case Index
        
        
        Case 4
            If SaveWorkStation Then
                PopupMessage 2, "Informations saved...", , , "CP Workstation"
            End If
    
    End Select
End Sub




Private Sub Image2_Click()
frDatabaseType.Visible = False
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
    Case 4
        If CheckPrivilege(3) Then
            If SaveWorkStation Then
                PopupMessage 2, "Department Info Saved!"
            End If
        End If
    Case 5
        ' ricerca aggiornamenti online
        Call SearchUpdate
    Case 8
        F_RegForm.DoShow True
        CheckRegistration
End Select
End Sub


Private Sub CheckRegistration()
If GetSetting(App.Title, "Autorizzazione", "Demo", False) Then
        Label1(8) = "Demo"
        Label1(8).BackColor = vbRed
        
Else
        Label1(8) = "Registered"
        Label1(8).BackColor = vbColorGreen
End If
        
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 3
    If i = IndexProcedura Then
    Else
        PicMenu(i).BackColor = &H473733
    End If
Next
End Sub

Private Sub Form_Resize()
On Error Resume Next

ResizeControls
End Sub




Private Sub Image1_Click()
CelarForm
Text1(0).SetFocus
End Sub

Private Sub Image3_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub Label2_Click(Index As Integer)
PicMenu_Click Index
End Sub



Private Sub Label5_Click(Index As Integer)
Select Case Index
    Case 7
        SelezionoStampanteSettings
        Image5.Visible = bStampaOk
End Select
End Sub

Private Sub Label9_Click(Index As Integer)
DisableImage_Click Index
End Sub

Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub

Private Sub lbDatabase_Click(Index As Integer)

Select Case DatabaseIndex
    Case 0
        
        F_IMPORT_EXCEL.DoShow (Index)
        
    Case 1
        Select Case Index
        Case 1
            F_MACHINE.DoShow
        Case Else
        
            FormDatabase.Left = Me.Left
            FormDatabase.Top = Me.Top
            FormDatabase.WindowState = Me.WindowState
            
            If FormDatabase.DoShow(Index) Then
            
                
                
                
            End If
            
            
            GetDBInfo
            
            
            DoEvents
            DoEvents
            DoEvents
        End Select
End Select

'frDatabaseType.Visible = False

End Sub

Private Sub lbMenuIntro_Click(Index As Integer)
Label1_Click 5
End Sub

Private Sub MenuDatabase_Click(Index As Integer)
'If Not (bSonoAutorizzato) Then
'    PopupMessage 2, "Warning : Line Leader or Administrator Login required...", , True
'    If bExistAccount Then
'        CheckPrivilege 1
'        Label2(9) = MyOperatore.Name
'        CheckUser (1)
'    Else''

'    End If
        '
'    Exit Sub
'End If
DatabaseIndex = 2


Select Case Index
    Case 0
        If CheckPrivilege(3) = False Then Exit Sub
        F_SEARCHARCH.DoShow
    Case 1
        If CheckPrivilege(3) = False Then Exit Sub
        F_DUPLICA.DoShow
    Case 2
        If CheckPrivilege(3) = False Then Exit Sub
        Call ResetUserDatabase
    Case 3
        If CheckPrivilege(3) = False Then Exit Sub
        F_PERCORSO_ARCHIVIO.DoShow
    Case 4
    
        ' import Excel : Hanna Code
        If CheckPrivilege(3) = False Then Exit Sub
        Frame3(1).BackColor = &H473733
        lbData = Replace(lbData, "and Edit", "to Import")
        frDatabaseType.Left = PicMain(1).Width / 2 - frDatabaseType.Width / 2
        frDatabaseType.Top = PicMain(1).Height / 2 - frDatabaseType.Height / 2
        lbInside(1) = "Database Table : Excel Import"
        DatabaseIndex = 0
        frDatabaseType.Visible = True
        
        
       
    
    Case 5
        lbData = Replace(lbData, "to Import", "and Edit")
        If CheckPrivilege(1) = False Then Exit Sub
        Frame3(1).BackColor = &H644603
        lbInside(1) = "Database Table : View / Edit"
        frDatabaseType.Left = PicMain(1).Width / 2 - frDatabaseType.Width / 2
        frDatabaseType.Top = PicMain(1).Height / 2 - frDatabaseType.Height / 2

        DatabaseIndex = 1
        frDatabaseType.Visible = True
    
    Case 6
        If CheckPrivilege(1) = False Then Exit Sub
        Call RecoveryDatabaseFromFile
        
    Case 7
        If CheckPrivilege(1) = False Then Exit Sub
        Call ExportHannaCodeTodbCodeFinal


End Select

GetDBInfo

End Sub



Private Sub Option1_Click(Index As Integer)

'If CheckPrivilege(3) = False Then Exit Sub



    If Index = 0 Then
        bCLPClassification = False
        Option1(0).ForeColor = vbColorTextDarkBlue
        Option1(0).FontBold = True
        Option1(0).Value = True
        
        Option1(1).ForeColor = vbBlack
        Option1(1).FontBold = False
        
    Else
        bCLPClassification = True
        Option1(1).ForeColor = vbColorTextDarkBlue
        Option1(1).FontBold = True
        Option1(1).Value = True
        
        Option1(0).ForeColor = vbBlack
        Option1(0).FontBold = False
    End If
    
    SaveSetting App.Title, "Settings", "bCLPClassification", bCLPClassification
    
    If Me.Visible Then PopupMessage 2, "Please Restart Program...", , , "Change Method"
    
    
End Sub

Private Sub Option2_Click(Index As Integer)
Dim i As Integer
Dim rc As Boolean

For i = 0 To Option2.UBound
If i = Index Then
Else
    Option2(i).ForeColor = vbBlack
    Option2(i).FontBold = False
    Option2(i).Value = False
End If
    
Next

Option2(Index).ForeColor = vbColorTextDarkBlue
Option2(Index).FontBold = True

iLotNumberType = Index '
SaveSetting App.Title, "LotNumber", "iLotNumberType", iLotNumberType



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
' If Index > 3 Then Exit Function
On Error GoTo ERR_SEL
For i = 0 To PicMenu.Count - 1
    If i = Index Then
        PicMenu(i).BackColor = &H6D5B57
        PicInfo(i).Visible = True
    Else
        PicInfo(i).Visible = False
        PicMenu(i).BackColor = &H473733
    End If
Next
Set Image4(Index) = Image3(Index)
Picture4(Index).BackColor = PicInfo(Index).BackColor
Lab(4) = Label2(Index)
IndexProcedura = Index

blTable.Visible = True
Select Case IndexProcedura
    Case 0
        GetAccountInfo
        FormSelected True
        
    Case 1
        If bExistAccount Then
            Call CheckUser(1, 1)
        Else
            bSonoAutorizzato = True
        End If
        GetDBInfo
        FormSelected False
    Case 2
        FormSelected False
    Case 3
        GetWorkStation
        FormSelected False
    
End Select
    PicMain(Index).Visible = True
    PicMain(Index).ZOrder
   ' blTable = Label2(IndexProcedura)
    Label2(9) = IIf(MyOperatore.Name <> "", MyOperatore.Name, Label2(9))
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_SEL:
    Resume Next
End Function


Private Sub FormSelected(ByVal bValue As Boolean)
        DefaultMenuLabel(5).Visible = bValue
        Label9(0).Visible = bValue
        Image4(3).Visible = bValue
End Sub
Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = PicMenu.LBound To PicMenu.UBound
    If i = IndexProcedura Then
        ' lascialo stare...
    ElseIf i = Index Then
        PicMenu(i).BackColor = &H4D3B37
    Else
        PicMenu(i).BackColor = &H473733
    End If
Next
End Sub


Private Sub SetPicForm()
On Error GoTo ERR_SET:
Dim i As Integer
For i = 0 To PicMain.Count - 1
PicMain(i).Left = 0
PicMain(i).Top = PicInfo(0).Top + PicInfo(0).Height
PicMain(i).Width = Me.Width
PicMain(i).Height = Line1.y1 - PicMain(i).Top
Next



PicInfo(0).BackColor = &H644603 ' vbColorTextLightBlue
Label1(0).ForeColor = vbWhite
'Picture1(0).BackColor = &HB76C00
'Line3(0).BorderColor = &HB76C00

PicInfo(1).BackColor = &H745613 ' vbColorTextBlue 'vbColorTextLightBlue
Label1(1).ForeColor = vbWhite
'Picture1(1).BackColor = vbTimBlue ' vbColorTextBlue
'Line3(1).BorderColor = vbTimBlue ' vbColorTextBlue



PicInfo(2).BackColor = &H846623    ' vbColorTextLightBlue
Label1(2).ForeColor = vbWhite
'Picture1(2).BackColor = &HB76C00
'Line3(2).BorderColor = &HB76C00

PicInfo(3).BackColor = &H947633    'vbColorTextLightBlue
Label1(3).ForeColor = vbWhite
'Picture1(3).BackColor = vbColorTextBlue
'Line3(3).BorderColor = vbColorTextBlue

PicInfo(4).BackColor = &HA48643 ' vbColorTextLightBlue
Label1(7).ForeColor = vbWhite
'Picture1(4).BackColor = vbColorTextDarkBlue
'Line3(4).BorderColor = vbColorTextDarkBlue 'vbColorTextBlue

'&
'&
Picture4(2).BackColor = vbColorTextBlue


For i = 0 To PicInfo.Count - 1
    
    PicInfo(i).Left = 0
    PicInfo(i).Top = PicMainMenu(4).Height
 
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
            
           ' PopupMessage 2, Lab(4), , , Label2(IndexProcedura), Image4(0)
 
 
            Select Case IndexProcedura
                Case 0
                    If CheckLogin Then
                        frmPreferenze.Left = Me.Left
                        frmPreferenze.Top = Me.Top
                        frmPreferenze.WindowState = Me.WindowState
                        frmPreferenze.DoShow , Text1(0), Text1(1)
                    Else
                        PopupMessage 2, "Wrong Password...", , True
                    End If
                    GetAccountInfo
                    Exit Sub
                   
                Case 1
                  
                Case 2
                 
                
            End Select
            
    

                   Frm.Top = Me.Top
                   Frm.Left = Me.Left
                   Frm.WindowState = Me.WindowState
                   MyNewIndex = IndexProcedura
                   rc = Frm.DoShow(MyNewIndex, MyLot, MyCode, Image3(IndexProcedura))
                  ' GoSub CheckNewProcedura
                      
            
        Case 1
            ' tabello lotti chiusi
            
        Case 2
            
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
    DisableImage(3).Visible = True
    DisableImage(2).Visible = False
    DisableImage(1).Visible = True
    DisableImage(0).Visible = True
    Label9(0).Visible = True
    blTable.Visible = False
    
End Sub

Private Sub QrCodeLabel_Click(Index As Integer)
Dim rc As Boolean
Select Case Index
    Case 0
        TestReader
    Case 1
         rc = SetQRSeparator
         Call SetFormSeparator(rc)

End Select
End Sub


Private Sub SetFormSeparator(ByVal rc As Boolean)
         If rc Then
            
            QrCodeLabel(1) = "Barcode Reader | Separator Character = " & sQRSeparator
            QrCodeLabel(1).ForeColor = &H964901
            
         Else
         
            QrCodeLabel(1) = "Barcode Reader : Please Set Separator Character"
            QrCodeLabel(1).ForeColor = vbRed
         
         
         End If
End Sub

Private Sub Text1_Change(Index As Integer)
Dim rc As Boolean

rc = IIf(Len(Text1(0)) > 0, True, False)
rc = IIf(Len(Text1(1)) > 0, rc, False)

DisableImage(0).Visible = Not (rc)
Label9(0).Visible = DisableImage(0).Visible
Picture4(0).Visible = rc

Image1.Visible = rc

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = Text1.Count - 1 Then
            If Text1(0) <> "" Then
                Picture4_Click 0
                Exit Sub
            Else
                Index = 0
            End If
        Else
            Index = Index + 1
        End If
        Text1(Index).SetFocus
    
    End If
End Sub


Private Sub StartProcedure()
Dim rc As Boolean

Call CelarForm


End Sub

Private Sub GetOptions()
Dim rc As Boolean
' opzioni....

lbStr(0) = App.Major & "." & App.Minor & "." & App.Revision

   rc = IIf(sQRSeparator = "", False, True)
   SetFormSeparator (rc)
  
    bFullScreen = GetSetting(App.Title, "Opzioni", "Full Screen Mode", False)
    Check1.Value = IIf(bFullScreen, 1, 0)
    Check2.Value = IIf(bDotForDecimals, 1, 0)
    Check3.Value = IIf(bOpenProductClassificationAfterScan, 1, 0)
    


     Option1_Click IIf(bCLPClassification, 1, 0)
     
     Option2_Click iLotNumberType

 

    Call GetUserLine
    Call SetLine(cmbLine, True)

    If UserLine <> "" And UserLineIndex > 0 Then
        
        cmbLine = UserLine
        TxSocieta(0) = UserLine
    
    End If
    
    

End Sub
Private Function SearchUpdate() As Boolean

    If FileExists(App.Path & "\smartupdate.exe") Then
        
        ApriEseguibile App.Path & "\smartupdate.exe"
        SaveSetting App.Title, "Opzioni", "Avvisa Update", True
    Else
        MessageInfoTime = 2000
        PopupMessage 2, ("Attenzione impossibile trovare SmartUpdate.exe, Si consiglia di Reinstallare il programma."), , True
        SaveSetting App.Title, "Opzioni", "Avvisa Update", False
    End If
    
    'Call UploadMyInfo(F_MAIN.Inet1)
End Function


Private Sub TxSocieta_Click(Index As Integer)
If Index = 0 Then
    cmbLine.Move TxSocieta(0).Left, TxSocieta(0).Top, TxSocieta(0).Width
    cmbLine.Visible = True
    cmbLine.ZOrder
Else
    cmbLine.Visible = False
    
End If
End Sub

Private Sub TxSocieta_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    
    If Index < TxSocieta.Count - 1 Then
        TxSocieta(Index + 1).SetFocus
    Else
        TxSocieta(0).SetFocus
    End If
    

End If
End Sub

Private Function SaveWorkStation() As Boolean
Dim i As Integer
Dim rc As Boolean
On Error GoTo ERR_SAVE
rc = True
With dbTabLaboratorio
    .filter = ""
    If .EOF Then .AddNew
    For i = 1 To 5
        .fields(i) = TxSocieta(i - 1)
    
    Next
    .Update
End With

    Call SetWorkStation
ERR_END:
    On Error GoTo 0
    SaveWorkStation = rc
    Exit Function
ERR_SAVE:
    rc = False
    MsgBox err.Description
    GoTo ERR_END:
End Function
Private Function GetWorkStation() As Boolean
Dim i As Integer
Dim rc As Boolean
On Error GoTo ERR_SAVE
rc = True
With dbTabLaboratorio
    .filter = ""
    If .EOF Then Exit Function
    For i = 1 To 5
        TxSocieta(i - 1) = IIf(IsNull(Trim(.fields(i))), "", Trim(.fields(i)))
    Next
    .Update
End With

    Call SetWorkStation
ERR_END:
    On Error GoTo 0
    GetWorkStation = rc
    Exit Function
ERR_SAVE:
    rc = False
    MsgBox err.Description
    GoTo ERR_END:
End Function


Private Function GetDBInfo()

'm_CreateArchivio(dbPath, MydbName)
Label11(3).Visible = True
Label11(3) = "Database Name : " & MydbName

'Label11(2).Visible = True
'Label11(2) = "Path : " & App.Path


'Label11(4).Visible = True
'Label11(4) = "Current Path : " & dbPath

With dbTabCode
    .filter = ""
    If .EOF Then
        Label11(5).Visible = False
    Else
        Label11(5).Visible = True
        Label11(5) = "Hanna Code Count : " & .RecordCount
    End If
End With

With dbTabRecipe
    .filter = ""
    If .EOF Then
        Label11(6).Visible = False
    Else
        Label11(6).Visible = True
        Label11(6) = "Recipes Count : " & .RecordCount
    End If
End With

With dbTabRawMaterial
    .filter = ""
    If .EOF Then
        Label11(7).Visible = False
    Else
        Label11(7).Visible = True
        Label11(7) = "Chemical RM Count : " & .RecordCount
    End If
End With

Label11(8) = "Actual Rel: " & dbCodeRelease & " ( " & dbCodeDate & " - " & dbCodeOperator & ")"

End Function



Private Function GetAccountInfo() As Boolean
Dim rc As Boolean
Dim i As Integer
'bExistAdministrator
'bExsistProductionManager

    Image6(4).Visible = False
    Image6(3).Visible = False
    Image6(9).Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label10.Visible = False
            
    With dbTabUserAccount
        .filter = ""
        If .EOF Then
            Picture4(0).ZOrder
            Picture4(0).Visible = True
            Lab(4) = "Enter User Account"
            
            GetAccountInfo = False

        Else
            ' visualizza Manager e Administrator
            .MoveFirst
            For i = 1 To .RecordCount
            
                If !IndexPrivilege = 1 Then
                    Image6(9).Visible = True
                    Label10.Visible = True
                    Label10.Caption = "Production Manager : " & Trim(!UserID)
                End If
                
                If !IndexPrivilege = 2 Then
                    Image6(4).Visible = True
                    Label7.Visible = True
                    Label7.Caption = "Line Leader : " & Trim(!UserID)
                End If
                If !IndexPrivilege = 3 Then
                    Image6(3).Visible = True
                    Label8.Visible = True
                    Label8.Caption = "Administrator : " & Trim(!UserID)
                End If
                .MoveNext
            Next
            GetAccountInfo = True
        End If
    
    
    End With

End Function

Private Function CheckLogin() As Boolean
Dim rc As Boolean
rc = False
    With dbTabUserAccount
        .filter = ""
        .filter = "USERID='" & Text1(0) & "'"
        If .EOF Then
            rc = True
            PopupMessage 2, "New User Account..."
        Else
            If !Password = Text1(1) Then
                rc = True
            Else
                rc = False
            End If
        End If
        
    
    End With
    
    CheckLogin = rc
    
End Function

Private Function CheckUser(ByVal Index As Integer, ByVal i As Integer)
Dim rc As Boolean
rc = True
If MyOperatore.IndexPrivilege < i Then rc = False
If MyOperatore.Name = "" Then rc = False


DisableImage(Index).Visible = Not (rc)
Label9(Index).Visible = DisableImage(Index).Visible
bSonoAutorizzato = rc

If Index = 3 Then
    frMethod.Enabled = bSonoAutorizzato
End If
End Function





