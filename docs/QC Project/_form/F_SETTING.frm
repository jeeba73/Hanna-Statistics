VERSION 5.00
Begin VB.Form F_SETTING 
   BackColor       =   &H00808080&
   Caption         =   "Settings QC"
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
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   8655
      Index           =   1
      Left            =   1680
      ScaleHeight     =   8655
      ScaleWidth      =   19215
      TabIndex        =   34
      Top             =   2160
      Visible         =   0   'False
      Width           =   19215
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   1
         Left            =   17280
         MouseIcon       =   "F_SETTING.frx":0A02
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
            Picture         =   "F_SETTING.frx":0D0C
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
            MouseIcon       =   "F_SETTING.frx":40EE
            MousePointer    =   99  'Custom
            TabIndex        =   36
            Top             =   840
            Width           =   465
         End
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   7
         Left            =   4320
         MouseIcon       =   "F_SETTING.frx":43F8
         MousePointer    =   99  'Custom
         TabIndex        =   109
         Top             =   3600
         Width           =   4695
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   5
         Left            =   4200
         MouseIcon       =   "F_SETTING.frx":4702
         MousePointer    =   99  'Custom
         TabIndex        =   73
         Top             =   2160
         Width           =   4695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FGCode Database"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   285
         Index           =   8
         Left            =   5715
         MouseIcon       =   "F_SETTING.frx":4A0C
         MousePointer    =   99  'Custom
         TabIndex        =   110
         Top             =   4080
         Width           =   1740
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   11
         Left            =   7680
         Picture         =   "F_SETTING.frx":4D16
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   480
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--------"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   270
         Index           =   0
         Left            =   6855
         TabIndex        =   108
         Top             =   4440
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   6
         Left            =   4200
         MouseIcon       =   "F_SETTING.frx":6230
         MousePointer    =   99  'Custom
         TabIndex        =   101
         Top             =   5280
         Width           =   4455
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   10
         Left            =   7695
         Picture         =   "F_SETTING.frx":653A
         Top             =   5640
         Width           =   480
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
         MouseIcon       =   "F_SETTING.frx":991C
         MousePointer    =   99  'Custom
         TabIndex        =   102
         Top             =   5760
         Width           =   2775
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--------"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   270
         Index           =   5
         Left            =   6855
         TabIndex        =   82
         Top             =   3000
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   270
         Index           =   4
         Left            =   11760
         TabIndex        =   81
         Top             =   6120
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   270
         Index           =   3
         Left            =   11760
         TabIndex        =   80
         Top             =   1440
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   270
         Index           =   2
         Left            =   11760
         TabIndex        =   79
         Top             =   4560
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--------"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   270
         Index           =   1
         Left            =   11760
         TabIndex        =   78
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   3
         Left            =   10440
         MouseIcon       =   "F_SETTING.frx":9C26
         MousePointer    =   99  'Custom
         TabIndex        =   70
         Top             =   5280
         Width           =   4695
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1455
         Index           =   2
         Left            =   10320
         MouseIcon       =   "F_SETTING.frx":9F30
         MousePointer    =   99  'Custom
         TabIndex        =   69
         Top             =   3720
         Width           =   6015
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1095
         Index           =   1
         Left            =   10440
         MouseIcon       =   "F_SETTING.frx":A23A
         MousePointer    =   99  'Custom
         TabIndex        =   68
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1695
         Index           =   0
         Left            =   10440
         MouseIcon       =   "F_SETTING.frx":A544
         MousePointer    =   99  'Custom
         TabIndex        =   64
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label MenuDatabase 
         BackStyle       =   0  'Transparent
         Height          =   1575
         Index           =   4
         Left            =   4200
         MouseIcon       =   "F_SETTING.frx":A84E
         MousePointer    =   99  'Custom
         TabIndex        =   71
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratoy Manager or Administrator Login Required"
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
         Left            =   6840
         TabIndex        =   76
         Top             =   7440
         Width           =   5040
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   8
         Left            =   7680
         Picture         =   "F_SETTING.frx":AB58
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Code Database"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   285
         Index           =   5
         Left            =   5280
         MouseIcon       =   "F_SETTING.frx":DF3A
         MousePointer    =   99  'Custom
         TabIndex        =   74
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   7
         Left            =   7680
         Picture         =   "F_SETTING.frx":E244
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Import Hanna Code Excel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   285
         Index           =   4
         Left            =   4920
         MouseIcon       =   "F_SETTING.frx":11626
         MousePointer    =   99  'Custom
         TabIndex        =   72
         Top             =   1080
         Width           =   2445
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Database Path "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   285
         Index           =   3
         Left            =   11760
         MouseIcon       =   "F_SETTING.frx":11930
         MousePointer    =   99  'Custom
         TabIndex        =   67
         Top             =   5760
         Width           =   2130
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   6
         Left            =   11160
         Picture         =   "F_SETTING.frx":11C3A
         Top             =   5640
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reset Default Database ( Empty )"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   285
         Index           =   2
         Left            =   11760
         MouseIcon       =   "F_SETTING.frx":1501C
         MousePointer    =   99  'Custom
         TabIndex        =   66
         Top             =   4200
         Width           =   3255
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   5
         Left            =   11160
         Picture         =   "F_SETTING.frx":15326
         Top             =   4080
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copy / Backup Database"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   285
         Index           =   1
         Left            =   11760
         MouseIcon       =   "F_SETTING.frx":18708
         MousePointer    =   99  'Custom
         TabIndex        =   65
         Top             =   2640
         Width           =   2385
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   2
         Left            =   11160
         Picture         =   "F_SETTING.frx":18A12
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Import Database"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   0
         Left            =   11760
         MouseIcon       =   "F_SETTING.frx":1BDF4
         MousePointer    =   99  'Custom
         TabIndex        =   63
         Top             =   1080
         Width           =   1710
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   1
         Left            =   11160
         Picture         =   "F_SETTING.frx":1C0FE
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
         MouseIcon       =   "F_SETTING.frx":1F4E0
         MousePointer    =   99  'Custom
         Picture         =   "F_SETTING.frx":1F7EA
         Top             =   6840
         Width           =   480
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   8655
      Index           =   4
      Left            =   0
      ScaleHeight     =   8655
      ScaleWidth      =   19215
      TabIndex        =   89
      Top             =   2040
      Width           =   19215
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   4
         Left            =   4800
         MouseIcon       =   "F_SETTING.frx":22BCC
         MousePointer    =   99  'Custom
         ScaleHeight     =   1335
         ScaleWidth      =   9615
         TabIndex        =   94
         Top             =   6240
         Visible         =   0   'False
         Width           =   9615
         Begin VB.Image Image4 
            Height          =   480
            Index           =   4
            Left            =   4560
            Picture         =   "F_SETTING.frx":22ED6
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
            MouseIcon       =   "F_SETTING.frx":262B8
            MousePointer    =   99  'Custom
            TabIndex        =   95
            Top             =   840
            Width           =   465
         End
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   320
         Left            =   11040
         TabIndex        =   92
         Text            =   "Text2"
         Top             =   7440
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
         Caption         =   "View : Full Screen Mode"
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   5160
         TabIndex        =   90
         Top             =   7920
         Width           =   9255
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   4800
         ScaleHeight     =   405
         ScaleWidth      =   375
         TabIndex        =   91
         Top             =   7920
         Width           =   375
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
         Left            =   7200
         MouseIcon       =   "F_SETTING.frx":265C2
         MousePointer    =   99  'Custom
         TabIndex        =   107
         Top             =   1080
         Width           =   1080
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
         Left            =   7200
         MouseIcon       =   "F_SETTING.frx":268CC
         MousePointer    =   99  'Custom
         TabIndex        =   106
         Top             =   1440
         Width           =   2670
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7200
         TabIndex        =   105
         Top             =   480
         Width           =   1905
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   7200
         X2              =   12960
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   3840
         Picture         =   "F_SETTING.frx":26BD6
         Top             =   960
         Width           =   480
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
         Left            =   480
         MouseIcon       =   "F_SETTING.frx":280F0
         MousePointer    =   99  'Custom
         TabIndex        =   104
         Top             =   1080
         Width           =   2790
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   480
         X2              =   6240
         Y1              =   840
         Y2              =   840
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   480
         TabIndex        =   103
         Top             =   480
         Width           =   1425
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   "      Evaluation QC :  minimum % of selected Readings"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4800
         TabIndex        =   93
         Top             =   7440
         Width           =   6255
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   8655
      Index           =   0
      Left            =   0
      ScaleHeight     =   8655
      ScaleWidth      =   19215
      TabIndex        =   18
      Top             =   2040
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
         Index           =   0
         Left            =   8160
         TabIndex        =   26
         Top             =   3000
         Width           =   4215
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
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   8160
         PasswordChar    =   "*"
         TabIndex        =   20
         Top             =   3600
         Width           =   4215
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   0
         Left            =   4800
         MouseIcon       =   "F_SETTING.frx":283FA
         MousePointer    =   99  'Custom
         ScaleHeight     =   1335
         ScaleWidth      =   9615
         TabIndex        =   24
         Top             =   4920
         Width           =   9615
         Begin VB.Label Lab 
            Alignment       =   2  'Center
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
            Index           =   4
            Left            =   0
            MouseIcon       =   "F_SETTING.frx":28704
            MousePointer    =   99  'Custom
            TabIndex        =   25
            Top             =   840
            Width           =   9585
         End
         Begin VB.Image Image4 
            Height          =   480
            Index           =   0
            Left            =   4560
            Picture         =   "F_SETTING.frx":28A0E
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   735
         Index           =   6
         Left            =   4440
         MouseIcon       =   "F_SETTING.frx":2BDF0
         MousePointer    =   99  'Custom
         TabIndex        =   99
         Top             =   1680
         Width           =   5655
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   9
         Left            =   9360
         Picture         =   "F_SETTING.frx":2C0FA
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratory Manager : Gibertini Riccardo"
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
         Left            =   2040
         MouseIcon       =   "F_SETTING.frx":2F4DC
         MousePointer    =   99  'Custom
         TabIndex        =   100
         Top             =   1920
         Width           =   7020
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   735
         Index           =   5
         Left            =   4200
         MouseIcon       =   "F_SETTING.frx":2F7E6
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
         MouseIcon       =   "F_SETTING.frx":2FAF0
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   0
         Width           =   5655
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter UserID and Password"
         ForeColor       =   &H000040C0&
         Height          =   285
         Index           =   0
         Left            =   8280
         TabIndex        =   77
         Top             =   6000
         Width           =   2760
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   5280
         Picture         =   "F_SETTING.frx":2FDFA
         Top             =   3285
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   4
         Left            =   9360
         Picture         =   "F_SETTING.frx":331DC
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   3
         Left            =   9360
         Picture         =   "F_SETTING.frx":365BE
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratory Manager : Gibertini Riccardo"
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
         Left            =   840
         MouseIcon       =   "F_SETTING.frx":399A0
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   480
         Width           =   8220
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
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
         Left            =   6120
         TabIndex        =   27
         Tag             =   "User"
         Top             =   3000
         Width           =   1740
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Index           =   0
         Left            =   9360
         Picture         =   "F_SETTING.frx":39CAA
         Top             =   5400
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   1815
         Index           =   1
         Left            =   4800
         Top             =   2640
         Width           =   9615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Administrator : Gibertini Riccardo"
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
         Left            =   3240
         MouseIcon       =   "F_SETTING.frx":3D08C
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   1200
         Width           =   5865
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   13200
         MouseIcon       =   "F_SETTING.frx":3D396
         MousePointer    =   99  'Custom
         Picture         =   "F_SETTING.frx":3D6A0
         Top             =   3240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   6120
         TabIndex        =   19
         Top             =   3600
         Width           =   1620
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
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   8655
      Index           =   3
      Left            =   0
      ScaleHeight     =   8655
      ScaleWidth      =   19215
      TabIndex        =   42
      Top             =   2040
      Visible         =   0   'False
      Width           =   19215
      Begin VB.TextBox TxSocieta 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00964901&
         Height          =   510
         Index           =   0
         Left            =   4440
         TabIndex        =   51
         Top             =   1200
         Width           =   10335
      End
      Begin VB.TextBox TxSocieta 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00964901&
         Height          =   510
         Index           =   1
         Left            =   4440
         TabIndex        =   50
         Top             =   2400
         Width           =   10335
      End
      Begin VB.TextBox TxSocieta 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00964901&
         Height          =   510
         Index           =   2
         Left            =   4440
         TabIndex        =   49
         Top             =   3720
         Width           =   10335
      End
      Begin VB.TextBox TxSocieta 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00964901&
         Height          =   510
         Index           =   3
         Left            =   4440
         TabIndex        =   48
         Top             =   5040
         Width           =   10335
      End
      Begin VB.TextBox TxSocieta 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00964901&
         Height          =   510
         Index           =   4
         Left            =   4440
         TabIndex        =   47
         Top             =   6240
         Width           =   10335
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   3
         Left            =   16440
         MouseIcon       =   "F_SETTING.frx":40A82
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
            MouseIcon       =   "F_SETTING.frx":40D8C
            MousePointer    =   99  'Custom
            TabIndex        =   44
            Top             =   840
            Width           =   465
         End
         Begin VB.Image Image4 
            Height          =   480
            Index           =   3
            Left            =   4560
            Picture         =   "F_SETTING.frx":41096
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Label LbSocieta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratory Manager"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Index           =   4
         Left            =   4440
         TabIndex        =   57
         Top             =   4560
         Width           =   3450
      End
      Begin VB.Label LbSocieta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Workstation"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Index           =   2
         Left            =   4440
         TabIndex        =   56
         Top             =   3240
         Width           =   2100
      End
      Begin VB.Label LbSocieta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Index           =   1
         Left            =   4440
         TabIndex        =   55
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label LbSocieta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Index           =   0
         Left            =   4440
         TabIndex        =   54
         Top             =   720
         Width           =   2040
      End
      Begin VB.Label LbSocieta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Index           =   5
         Left            =   4440
         TabIndex        =   53
         Top             =   5760
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   4
         Left            =   4440
         MouseIcon       =   "F_SETTING.frx":44478
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   7080
         Width           =   10335
      End
      Begin VB.Image Image3 
         Height          =   480
         Index           =   8
         Left            =   9360
         Picture         =   "F_SETTING.frx":44782
         Top             =   360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Index           =   0
         Left            =   13800
         Picture         =   "F_SETTING.frx":47B64
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
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   4
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   96
      Top             =   6720
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options : You can select or change options in this page..."
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   7
         Left            =   480
         TabIndex        =   97
         Top             =   360
         Width           =   5790
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
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratory Info / Hanna Department : fill Lab form and Save"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   3
         Left            =   480
         TabIndex        =   17
         Top             =   360
         Width           =   6195
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
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Update Program : Online update, check new releases..."
         ForeColor       =   &H00004080&
         Height          =   285
         Index           =   2
         Left            =   480
         TabIndex        =   16
         Top             =   360
         Width           =   5595
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
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Settings : Copy, Import & Export Databases"
         ForeColor       =   &H000040C0&
         Height          =   285
         Index           =   1
         Left            =   480
         TabIndex        =   15
         Top             =   360
         Width           =   5235
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
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Account Manager : Select Administrator, Laboratory Manager & QC Operator "
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   14
         Top             =   360
         Width           =   8340
      End
   End
   Begin VB.PictureBox PicMainMenu 
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
         Index           =   4
         Left            =   7680
         MouseIcon       =   "F_SETTING.frx":4AF46
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   4
            Left            =   0
            MouseIcon       =   "F_SETTING.frx":4B250
            MousePointer    =   99  'Custom
            TabIndex        =   84
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   4
            Left            =   720
            MouseIcon       =   "F_SETTING.frx":4B55A
            MousePointer    =   99  'Custom
            OLEDropMode     =   1  'Manual
            Picture         =   "F_SETTING.frx":4B864
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
         MouseIcon       =   "F_SETTING.frx":4EC46
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
            MouseIcon       =   "F_SETTING.frx":4EF50
            MousePointer    =   99  'Custom
            OLEDropMode     =   1  'Manual
            Picture         =   "F_SETTING.frx":4F25A
            Top             =   180
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   3
            Left            =   0
            MouseIcon       =   "F_SETTING.frx":5263C
            MousePointer    =   99  'Custom
            TabIndex        =   9
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
         MouseIcon       =   "F_SETTING.frx":52946
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
            MouseIcon       =   "F_SETTING.frx":52C50
            MousePointer    =   99  'Custom
            Picture         =   "F_SETTING.frx":52F5A
            Top             =   180
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   2
            Left            =   0
            MouseIcon       =   "F_SETTING.frx":5633C
            MousePointer    =   99  'Custom
            TabIndex        =   7
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
         MouseIcon       =   "F_SETTING.frx":56646
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   4
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   735
            MouseIcon       =   "F_SETTING.frx":56950
            MousePointer    =   99  'Custom
            Picture         =   "F_SETTING.frx":56C5A
            Top             =   180
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   1
            Left            =   90
            MouseIcon       =   "F_SETTING.frx":5A03C
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   720
            Width           =   1830
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MouseIcon       =   "F_SETTING.frx":5A346
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
            Left            =   105
            MouseIcon       =   "F_SETTING.frx":5A650
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   720
            Width           =   1800
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   735
            MouseIcon       =   "F_SETTING.frx":5A95A
            MousePointer    =   99  'Custom
            Picture         =   "F_SETTING.frx":5AC64
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Settings"
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
         Left            =   11760
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   6930
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   8655
      Index           =   2
      Left            =   0
      ScaleHeight     =   8655
      ScaleWidth      =   19215
      TabIndex        =   38
      Top             =   2040
      Visible         =   0   'False
      Width           =   19215
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   2
         Left            =   4800
         MouseIcon       =   "F_SETTING.frx":5E046
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
            MouseIcon       =   "F_SETTING.frx":5E350
            MousePointer    =   99  'Custom
            TabIndex        =   40
            Top             =   840
            Width           =   465
         End
         Begin VB.Image Image4 
            Height          =   480
            Index           =   2
            Left            =   4560
            Picture         =   "F_SETTING.frx":5E65A
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Software License"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   8
         Left            =   2040
         MouseIcon       =   "F_SETTING.frx":61A3C
         MousePointer    =   99  'Custom
         TabIndex        =   98
         Top             =   6840
         Width           =   15135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratoy Manager or Administrator Login Required"
         ForeColor       =   &H000040C0&
         Height          =   285
         Index           =   2
         Left            =   6960
         TabIndex        =   75
         Top             =   6000
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.Label lbStr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
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
         Index           =   0
         Left            =   3360
         TabIndex        =   61
         Top             =   3360
         Width           =   9015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   435
         Index           =   6
         Left            =   1920
         TabIndex        =   60
         Top             =   3360
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   24.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   570
         Left            =   10110
         TabIndex        =   59
         Top             =   2520
         Width           =   165
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00964901&
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
         MouseIcon       =   "F_SETTING.frx":61D46
         MousePointer    =   99  'Custom
         Picture         =   "F_SETTING.frx":62050
         Top             =   2760
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Search Online Update"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   5
         Left            =   2040
         MouseIcon       =   "F_SETTING.frx":65432
         MousePointer    =   99  'Custom
         TabIndex        =   58
         Top             =   3960
         Width           =   15135
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Index           =   2
         Left            =   9360
         MouseIcon       =   "F_SETTING.frx":6573C
         MousePointer    =   99  'Custom
         Picture         =   "F_SETTING.frx":65A46
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
         Caption         =   " SOFTWARE RELEASE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00964901&
         Height          =   2025
         Index           =   3
         Left            =   1800
         MouseIcon       =   "F_SETTING.frx":68E28
         MousePointer    =   99  'Custom
         TabIndex        =   62
         Top             =   2520
         Width           =   12615
      End
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   3
      Left            =   0
      MouseIcon       =   "F_SETTING.frx":69132
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   10560
      Width           =   1815
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   1
      Left            =   8280
      MouseIcon       =   "F_SETTING.frx":6943C
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
      MouseIcon       =   "F_SETTING.frx":69746
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
      MouseIcon       =   "F_SETTING.frx":69A50
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   12
      Left            =   17775
      MouseIcon       =   "F_SETTING.frx":69D5A
      MousePointer    =   99  'Custom
      TabIndex        =   88
      Top             =   11600
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   11
      Left            =   15735
      MouseIcon       =   "F_SETTING.frx":6A064
      MousePointer    =   99  'Custom
      TabIndex        =   87
      Top             =   11600
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   10
      Left            =   8955
      MouseIcon       =   "F_SETTING.frx":6A36E
      MousePointer    =   99  'Custom
      TabIndex        =   86
      Top             =   11600
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select User"
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
      Index           =   9
      Left            =   360
      MouseIcon       =   "F_SETTING.frx":6A678
      MousePointer    =   99  'Custom
      TabIndex        =   85
      Top             =   11600
      Width           =   1170
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_SETTING.frx":6A982
      Height          =   480
      Index           =   0
      Left            =   18240
      MouseIcon       =   "F_SETTING.frx":6DD64
      MousePointer    =   99  'Custom
      Picture         =   "F_SETTING.frx":6E06E
      Top             =   11040
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
      Y1              =   0
      Y2              =   11880
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
      MouseIcon       =   "F_SETTING.frx":71450
      MousePointer    =   99  'Custom
      Picture         =   "F_SETTING.frx":7175A
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_SETTING.frx":74B3C
      Height          =   480
      Index           =   1
      Left            =   9360
      MouseIcon       =   "F_SETTING.frx":77F1E
      MousePointer    =   99  'Custom
      Picture         =   "F_SETTING.frx":78228
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   15960
      MouseIcon       =   "F_SETTING.frx":7B60A
      MousePointer    =   99  'Custom
      Picture         =   "F_SETTING.frx":7B914
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   4
      Left            =   2040
      MouseIcon       =   "F_SETTING.frx":7ECF6
      MousePointer    =   99  'Custom
      Picture         =   "F_SETTING.frx":7F000
      Top             =   11040
      Visible         =   0   'False
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
            .Left = ctl.X1
            .Top = ctl.Y1
            .Width = ctl.X2 - ctl.X1
            .Height = ctl.Y2 - ctl.Y1
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
            ctl.X1 = x_scale * .Left
            ctl.Y1 = y_scale * .Top
            ctl.X2 = ctl.X1 + x_scale * .Width
            ctl.Y2 = ctl.Y1 + y_scale * .Height
        ElseIf TypeOf ctl Is Timer Then
        ElseIf TypeOf ctl Is Inet Then
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

    
    Me.Show vbModal
    
    If m_rc = True Then

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

Private Sub Check1_Click()
Dim rc As Boolean
rc = IIf(Check1.Value = 1, True, False)
bFullScreen = rc
SaveSetting App.Title, "Opzioni", "Full Screen Mode", bFullScreen
If Me.Visible Then PopupMessage 2, "Please Restart the Program to take affect changes..."
End Sub

Private Sub DefaultMenu_Click(Index As Integer)
Label1_Click 5
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


PopupMessage 2, "Warning : Only Laboratory Manager ot TCO can Operate" & vbCrLf & "Please Login...", , True
CheckPrivilege 1
CheckUser (Index)

End Sub

Private Sub Form_Initialize()

Call StartProcedure
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
 Call SetPicForm
IndexProcedura = 99
Dim i As Integer
'If Screen.Width - Me.Width > 1000 And bFullScreen Then
   ' Me.WindowState = 2
    For i = 0 To PicMain.Count - 1
        PicMain(i).Picture = LoadPicture(PictureMaxScreen)
      
    Next '
      'Me.Picture = LoadPicture(PictureMaxScreen)
'Else
   ' Me.WindowState = 0
'End If
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
        Label1(8).BackColor = vbRed
Else
        Label1(8).BackColor = vbColorGreen
End If
        
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
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

Private Sub lbMenuIntro_Click(Index As Integer)
Label1_Click 5
End Sub

Private Sub MenuDatabase_Click(Index As Integer)
'If Not (bSonoAutorizzato) Then
'    PopupMessage 2, "Warning : Laboratory Manager or Administrator Login required...", , True
'    If bExistAccount Then
'        CheckPrivilege 1
'        Label2(9) = MyOperatore.Name
'        CheckUser (1)
'    Else''

'    End If
        '
'    Exit Sub
'End If



Select Case Index
    Case 0
        If CheckPrivilege(3) = False Then Exit Sub
        
        If CheckCodeDB Then
        Else
            F_SEARCHARCH.DoShow
        End If
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
        F_IMPORT_EXCEL.DoShow
       
    
    Case 5
        FormCode.WindowState = Me.WindowState
        FormCode.Left = Me.Left
        FormCode.Top = Me.Top
        FormCode.DoShow
    Case 6
    
        If CheckPrivilege(1) = False Then Exit Sub
        Call RecoveryDatabaseFromFile
        
         PopupMessage 2, "Database recovery finished"
        
    Case 7
        FormFGCode.WindowState = Me.WindowState
        FormFGCode.Left = Me.Left
        FormFGCode.Top = Me.Top
        FormFGCode.DoShow
        
End Select

GetDBInfo

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
        PicMenu(i).BackColor = vbColorForeFixed
        PicInfo(i).Visible = True
    Else
        PicInfo(i).Visible = False
        PicMenu(i).BackColor = vbColorDarkFont
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
            CheckUser (1)
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
Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
For i = 0 To PicMenu.UBound
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
Dim i As Integer
For i = 0 To PicMain.Count - 1
PicMain(i).Left = 0
PicMain(i).Top = 2063
PicMain(i).Width = Me.Width
PicMain(i).Height = Line1.Y1 - PicMain(i).Top
Next



PicInfo(0).BackColor = &HB76C00 ' vbColorTextLightBlue
Label1(0).ForeColor = vbWhite
'Picture1(0).BackColor = &HB76C00
'Line3(0).BorderColor = &HB76C00

PicInfo(1).BackColor = vbTimBlue ' vbColorTextBlue 'vbColorTextLightBlue
Label1(1).ForeColor = vbWhite
'Picture1(1).BackColor = vbTimBlue ' vbColorTextBlue
'Line3(1).BorderColor = vbTimBlue ' vbColorTextBlue


PicInfo(2).BackColor = &HB76C00 ' vbColorTextLightBlue
Label1(2).ForeColor = vbWhite
'Picture1(2).BackColor = &HB76C00
'Line3(2).BorderColor = &HB76C00

PicInfo(3).BackColor = vbColorTextBlue 'vbColorTextLightBlue
Label1(3).ForeColor = vbWhite
'Picture1(3).BackColor = vbColorTextBlue
'Line3(3).BorderColor = vbColorTextBlue

PicInfo(4).BackColor = vbColorTextDarkBlue ' vbColorTextLightBlue
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
                        frmPreferenze.WindowState = Me.WindowState
                        frmPreferenze.Left = Me.Left
                        frmPreferenze.Top = Me.Top
                        frmPreferenze.DoShow , Text1(0), Text1(1)
                    Else
                        PopupMessage 2, "Wrong Password...", , True
                    End If
                    GetAccountInfo
                    Exit Sub
                   
                Case 1
                    Set Frm = F_READING
                Case 2
                    Set Frm = F_EVALUATION
                
            End Select
            
    
                Frm.WindowState = Me.WindowState
                   Frm.Top = Me.Top
                   Frm.Left = Me.Left
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

' opzioni....

lbStr(0) = App.Major & "." & App.Minor & "." & App.Revision


    MinNumberSelecterPerc = GetSetting(App.Title, "Options", "MinNumberSelecterPerc", 0.7)
    
    bFullScreen = GetSetting(App.Title, "Opzioni", "Full Screen Mode", False)
    Check1.Value = IIf(bFullScreen, 1, 0)
    Text2 = FormatNumber(MinNumberSelecterPerc * 100, 0)
    

End Sub
Private Function SearchUpdate() As Boolean

    If FileExists(App.PATH & "\smartupdate.exe") Then
        
        ApriEseguibile App.PATH & "\smartupdate.exe"
        SaveSetting App.Title, "Opzioni", "Avvisa Update", True
    Else
        MessageInfoTime = 2000
        PopupMessage 2, ("Attenzione impossibile trovare SmartUpdate.exe, Si consiglia di Reinstallare il programma."), , True
        SaveSetting App.Title, "Opzioni", "Avvisa Update", False
    End If
    
    'Call UploadMyInfo(F_MAIN.Inet1)
End Function

Private Sub Text2_Click()
Dim sString As String
sString = Text2
ripeti:
PopupMessage 2, "Please insert a number : 0 to 100 ..."
If F_InputBox.DoShow("Enter Tolerance : es . 70 ", , , , , sString) Then

    If IsNumeric(sString) Then
        If sString < 0 Or sString > 100 Then GoTo ripeti:
    
        Text2 = sString
        Text2_LostFocus
    End If
Else

    
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = TxtToNumber(KeyAscii)
End Sub

Private Sub Text2_LostFocus()
If IsNumeric(Text2) Then
If Text2 > 10 Then
    MinNumberSelecterPerc = FormatNumber(Text2 / 100, 2)
End If

SaveSetting App.Title, "Options", "MinNumberSelecterPerc", MinNumberSelecterPerc
    
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
        .Fields(i) = TxSocieta(i - 1)
    
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
    MsgBox Err.Description
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
        TxSocieta(i - 1) = IIf(IsNull(Trim(.Fields(i))), "", Trim(.Fields(i)))
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
    MsgBox Err.Description
    GoTo ERR_END:
End Function


Private Function GetDBInfo()
Dim sString As String
'm_CreateArchivio(dbPath, MydbName)
Label11(3).Visible = True


MydbName = "dbCode.mdb"


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

With dbTabFinishGood
    .filter = ""
    If .EOF Then
        Label11(0).Visible = False
    Else
        Label11(0).Visible = True
        Label11(0) = "FGCode Count : " & .RecordCount
    End If
End With




Call GetLastImport(, sString)
Label11(1).Caption = sString
End Function



Private Function GetAccountInfo() As Boolean
Dim rc As Boolean
Dim i As Integer
'bExistAdministrator
'bExistManager

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
                    Label10.Caption = "Laboratory Manager : " & Trim(!UserID)
                End If
                
                If !IndexPrivilege = 2 Then
                    Image6(4).Visible = True
                    Label7.Visible = True
                    Label7.Caption = "Tecnical Office ( TCO ) : " & Trim(!UserID)
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

Private Function CheckUser(ByVal Index As Integer)
Dim rc As Boolean
rc = True
If MyOperatore.IndexPrivilege < 1 Then rc = False
If MyOperatore.Name = "" Then rc = False


DisableImage(Index).Visible = Not (rc)
Label9(Index).Visible = DisableImage(Index).Visible
bSonoAutorizzato = rc
End Function
