VERSION 5.00
Begin VB.Form F_INFORMATION 
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
   Picture         =   "F_INFORMATION.frx":0000
   ScaleHeight     =   12000
   ScaleMode       =   0  'User
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   0
      Left            =   -120
      Picture         =   "F_INFORMATION.frx":1DED9
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Frame Frame2 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1815
         Left            =   2520
         TabIndex        =   119
         Top             =   3840
         Width           =   13935
         Begin VB.Image Image1 
            Height          =   480
            Left            =   4200
            Picture         =   "F_INFORMATION.frx":3BDB2
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
            TabIndex        =   122
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
            TabIndex        =   121
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
            TabIndex        =   120
            Top             =   1080
            Width           =   3645
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H000080DF&
            BorderColor     =   &H00006000&
            Height          =   1815
            Index           =   2
            Left            =   0
            Top             =   0
            Width           =   13935
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
         TabIndex        =   62
         Top             =   5145
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
         TabIndex        =   60
         Top             =   5160
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
         TabIndex        =   58
         Top             =   5160
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
         TabIndex        =   55
         Top             =   5145
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
         TabIndex        =   53
         Top             =   5145
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
         TabIndex        =   51
         Top             =   4200
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
         TabIndex        =   49
         Top             =   4185
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
         TabIndex        =   47
         Top             =   4185
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
         TabIndex        =   45
         Top             =   4200
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
         TabIndex        =   43
         Top             =   4185
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
         TabIndex        =   41
         Top             =   2145
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
         TabIndex        =   39
         Top             =   2145
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
         TabIndex        =   37
         Top             =   1545
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
         TabIndex        =   35
         Top             =   1545
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
         TabIndex        =   33
         Top             =   1545
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
         TabIndex        =   31
         Top             =   1545
         Width           =   2535
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   6
         Left            =   15000
         Picture         =   "F_INFORMATION.frx":3F194
         Top             =   6600
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
         Left            =   15600
         MouseIcon       =   "F_INFORMATION.frx":42576
         MousePointer    =   99  'Custom
         TabIndex        =   65
         Top             =   6720
         Width           =   3195
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00606060&
         Height          =   1575
         Index           =   1
         Left            =   11400
         Top             =   1320
         Width           =   6015
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00606060&
         Height          =   2655
         Index           =   0
         Left            =   1680
         Top             =   3240
         Visible         =   0   'False
         Width           =   15735
      End
      Begin VB.Label Label1 
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
         TabIndex        =   64
         Top             =   3120
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
         Index           =   21
         Left            =   13680
         TabIndex        =   63
         Top             =   4800
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
         Index           =   20
         Left            =   10920
         TabIndex        =   61
         Top             =   4800
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
         Index           =   19
         Left            =   8160
         TabIndex        =   59
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Label Label1 
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
         TabIndex        =   56
         Top             =   1200
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
         TabIndex        =   57
         Top             =   4800
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
         TabIndex        =   54
         Top             =   4800
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
         TabIndex        =   52
         Top             =   3840
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
         TabIndex        =   50
         Top             =   3840
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
         TabIndex        =   48
         Top             =   3840
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
         TabIndex        =   46
         Top             =   3840
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
         Height          =   345
         Index           =   11
         Left            =   2520
         TabIndex        =   44
         Top             =   3840
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
         TabIndex        =   42
         Top             =   1800
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
         TabIndex        =   40
         Top             =   1800
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
         TabIndex        =   38
         Top             =   1200
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
         TabIndex        =   36
         Top             =   1200
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
         TabIndex        =   34
         Top             =   1200
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
         TabIndex        =   32
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Left            =   9360
         Picture         =   "F_INFORMATION.frx":42880
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
         TabIndex        =   9
         Top             =   7440
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      FillColor       =   &H00004000&
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   2880
      ScaleHeight     =   855
      ScaleWidth      =   3975
      TabIndex        =   18
      Top             =   10920
      Width           =   3975
      Begin VB.Image ImageTAV 
         Height          =   480
         Index           =   5
         Left            =   1720
         MouseIcon       =   "F_INFORMATION.frx":45C62
         MousePointer    =   99  'Custom
         Picture         =   "F_INFORMATION.frx":45F6C
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
      MouseIcon       =   "F_INFORMATION.frx":4934E
      MousePointer    =   99  'Custom
      ScaleHeight     =   1815
      ScaleWidth      =   2775
      TabIndex        =   20
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
         MouseIcon       =   "F_INFORMATION.frx":49658
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   0
         Left            =   1200
         Picture         =   "F_INFORMATION.frx":49962
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
      TabIndex        =   19
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
         Left            =   12480
         TabIndex        =   29
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
         Left            =   9720
         TabIndex        =   27
         Top             =   825
         Width           =   2535
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
         Left            =   5760
         TabIndex        =   25
         Text            =   "Allumnium Perfectu"
         Top             =   825
         Width           =   3735
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
         TabIndex        =   23
         Text            =   "HI454323455"
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
         Left            =   1200
         TabIndex        =   0
         Text            =   "G56433"
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
         Left            =   12480
         TabIndex        =   30
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
         Left            =   9720
         TabIndex        =   28
         Top             =   480
         Width           =   2535
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
         Left            =   5760
         TabIndex        =   26
         Top             =   480
         Width           =   3735
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
         Left            =   3240
         TabIndex        =   24
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
         Left            =   1200
         TabIndex        =   22
         Top             =   480
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
         Index           =   3
         Left            =   5520
         MouseIcon       =   "F_INFORMATION.frx":4CD44
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   68
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
            MouseIcon       =   "F_INFORMATION.frx":4D04E
            MousePointer    =   99  'Custom
            TabIndex        =   69
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   735
            MouseIcon       =   "F_INFORMATION.frx":4D358
            MousePointer    =   99  'Custom
            Picture         =   "F_INFORMATION.frx":4D662
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
         MouseIcon       =   "F_INFORMATION.frx":50A44
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   15
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "F_INFORMATION.frx":50D4E
            MousePointer    =   99  'Custom
            Picture         =   "F_INFORMATION.frx":51058
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
            MouseIcon       =   "F_INFORMATION.frx":5443A
            MousePointer    =   99  'Custom
            TabIndex        =   16
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
         MouseIcon       =   "F_INFORMATION.frx":54744
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   13
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Preparation / Old Lot"
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
            MouseIcon       =   "F_INFORMATION.frx":54A4E
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   720
            Width           =   1875
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   735
            MousePointer    =   99  'Custom
            Picture         =   "F_INFORMATION.frx":54D58
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
         MouseIcon       =   "F_INFORMATION.frx":5813A
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   11
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Instrument"
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
            MouseIcon       =   "F_INFORMATION.frx":58444
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   735
            MouseIcon       =   "F_INFORMATION.frx":5874E
            MousePointer    =   99  'Custom
            Picture         =   "F_INFORMATION.frx":58A58
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
         TabIndex        =   17
         Top             =   360
         Width           =   2820
      End
   End
   Begin VB.PictureBox PicMain 
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   1
      Left            =   1440
      Picture         =   "F_INFORMATION.frx":5BE3A
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   66
      Top             =   2160
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
         Index           =   32
         Left            =   14160
         TabIndex        =   92
         Top             =   4185
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
         Left            =   11520
         TabIndex        =   90
         Top             =   4185
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
         Left            =   3600
         TabIndex        =   88
         Top             =   4185
         Width           =   7815
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
         Left            =   11520
         TabIndex        =   86
         Top             =   2985
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
         Left            =   8880
         TabIndex        =   84
         Top             =   2985
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
         Left            =   6240
         TabIndex        =   82
         Top             =   2985
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
         Index           =   26
         Left            =   3600
         TabIndex        =   80
         Top             =   2985
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
         Left            =   14160
         TabIndex        =   78
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
         Left            =   11520
         TabIndex        =   76
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
         Left            =   8880
         TabIndex        =   74
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
         Left            =   6240
         TabIndex        =   72
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
         Left            =   3600
         TabIndex        =   70
         Top             =   1665
         Width           =   2535
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
         Index           =   34
         Left            =   14160
         TabIndex        =   93
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Department"
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
         Left            =   11520
         TabIndex        =   91
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "QC Note"
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
         Left            =   3600
         TabIndex        =   89
         Top             =   3840
         Width           =   7815
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
         Index           =   31
         Left            =   11520
         TabIndex        =   87
         Top             =   2640
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
         Index           =   30
         Left            =   8880
         TabIndex        =   85
         Top             =   2640
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
         Index           =   29
         Left            =   6240
         TabIndex        =   83
         Top             =   2640
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
         Index           =   28
         Left            =   3600
         TabIndex        =   81
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "Prod. Machine Code"
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
         Left            =   14160
         TabIndex        =   79
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
         Index           =   26
         Left            =   11520
         TabIndex        =   77
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
         Index           =   25
         Left            =   8880
         TabIndex        =   75
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
         Index           =   24
         Left            =   6240
         TabIndex        =   73
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
         Index           =   23
         Left            =   3600
         TabIndex        =   71
         Top             =   1320
         Width           =   2535
      End
   End
   Begin ChemicalQC.ctlCalendar ctlCalendar1 
      Height          =   6960
      Left            =   11520
      TabIndex        =   8
      Top             =   3960
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
   Begin VB.PictureBox PicMain 
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   2
      Left            =   840
      Picture         =   "F_INFORMATION.frx":79D13
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   67
      Top             =   1560
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
         Index           =   42
         Left            =   12480
         TabIndex        =   116
         Top             =   3960
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
         Left            =   10080
         TabIndex        =   114
         Top             =   3960
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
         Index           =   40
         Left            =   7680
         TabIndex        =   111
         Top             =   3960
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
         Index           =   39
         Left            =   5280
         TabIndex        =   109
         Top             =   3960
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
         Index           =   38
         Left            =   14760
         TabIndex        =   106
         Top             =   2160
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
         Index           =   37
         Left            =   12360
         TabIndex        =   104
         Top             =   2160
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
         Index           =   36
         Left            =   9960
         TabIndex        =   101
         Top             =   2160
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
         Index           =   35
         Left            =   7560
         TabIndex        =   99
         Top             =   2160
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
         Left            =   5160
         TabIndex        =   96
         Text            =   "CTO-218"
         Top             =   2160
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
         Index           =   33
         Left            =   2760
         TabIndex        =   94
         Text            =   "HI83200"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label1 
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
         Index           =   49
         Left            =   10080
         TabIndex        =   118
         Top             =   3120
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
         Index           =   48
         Left            =   12480
         TabIndex        =   117
         Top             =   3600
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
         Index           =   47
         Left            =   10080
         TabIndex        =   115
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label1 
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
         Index           =   46
         Left            =   5280
         TabIndex        =   113
         Top             =   3120
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
         Index           =   45
         Left            =   7680
         TabIndex        =   112
         Top             =   3600
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
         Index           =   44
         Left            =   5280
         TabIndex        =   110
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label1 
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
         Index           =   43
         Left            =   12360
         TabIndex        =   108
         Top             =   1320
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
         Left            =   14760
         TabIndex        =   107
         Top             =   1800
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
         Left            =   12360
         TabIndex        =   105
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label1 
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
         Index           =   40
         Left            =   7560
         TabIndex        =   103
         Top             =   1320
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
         Index           =   39
         Left            =   9960
         TabIndex        =   102
         Top             =   1800
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
         Index           =   38
         Left            =   7560
         TabIndex        =   100
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label1 
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
         Index           =   37
         Left            =   2760
         TabIndex        =   98
         Top             =   1320
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
         Index           =   36
         Left            =   5160
         TabIndex        =   97
         Top             =   1800
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
         Index           =   35
         Left            =   2760
         TabIndex        =   95
         Top             =   1800
         Width           =   2295
      End
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   4
      Left            =   2040
      MouseIcon       =   "F_INFORMATION.frx":97BEC
      MousePointer    =   99  'Custom
      Picture         =   "F_INFORMATION.frx":97EF6
      Top             =   11040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   0
      Left            =   17640
      MouseIcon       =   "F_INFORMATION.frx":9B2D8
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   10680
      Width           =   1575
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   2
      Left            =   15240
      MouseIcon       =   "F_INFORMATION.frx":9B5E2
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
      MouseIcon       =   "F_INFORMATION.frx":9B8EC
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   10560
      Width           =   2655
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   15960
      MouseIcon       =   "F_INFORMATION.frx":9BBF6
      MousePointer    =   99  'Custom
      Picture         =   "F_INFORMATION.frx":9BF00
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_INFORMATION.frx":9F2E2
      Height          =   480
      Index           =   1
      Left            =   9360
      MouseIcon       =   "F_INFORMATION.frx":A26C4
      MousePointer    =   99  'Custom
      Picture         =   "F_INFORMATION.frx":A29CE
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   480
      MouseIcon       =   "F_INFORMATION.frx":A5DB0
      MousePointer    =   99  'Custom
      Picture         =   "F_INFORMATION.frx":A60BA
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
      DragIcon        =   "F_INFORMATION.frx":A949C
      Height          =   480
      Index           =   0
      Left            =   18240
      MouseIcon       =   "F_INFORMATION.frx":AC87E
      MousePointer    =   99  'Custom
      Picture         =   "F_INFORMATION.frx":ACB88
      Top             =   11040
      Width           =   480
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   3
      Left            =   0
      MouseIcon       =   "F_INFORMATION.frx":AFF6A
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   10560
      Width           =   1815
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
Private IndexDate As Integer
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



Private Sub ctlCalendar1_DateClicked(inputDate As Date)
Text1(IndexDate) = inputDate
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

End Sub

Private Sub Form_Load()
IndexFormProcedura = 99


End Sub
Private Sub Frame2_Click()
Frame2.Visible = False
Label1(22).Visible = Not (Frame2.Visible)
Shape1(0).Visible = Label1(22).Visible
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 3
    If i = IndexFormProcedura Then
    Else
        PicMenu(i).BackColor = &H404040
    End If
Next
Picture1.BackColor = &H4000&
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set F_INFORMATION = Nothing
End Sub

Private Sub Image3_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub Label2_Click(Index As Integer)
PicMenu_Click Index
End Sub


Private Sub PicMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Len(Text1(IndexText)) = 0 Then Text1(IndexText).BackColor = vbColorUnabled
End Sub

Private Sub PicMenu_Click(Index As Integer)
If IndexFormProcedura = Index Then
ElseIf Index = PicMenu.Count - 1 Then
    IndexMainProcedura = IndexMainProcedura + 1
    Unload Me
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
    Else
        PicMenu(i).BackColor = &H404040
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
        Picture4(0).BackColor = &H5000&
        rc = False
End Select
For i = 0 To 4
    Text1(i).Locked = Not (rc)
Next
Label2(4) = Label2(Index)
IndexFormProcedura = Index
PicMain(Index).Visible = True
PicMain(Index).ZOrder
blTable = "Information QC : " & Label2(IndexFormProcedura)
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
        PicMenu(i).BackColor = &H404040
    End If
Next
End Sub


Private Sub SetPicForm()
Dim i As Integer
ctlCalendar1.Left = Me.Width / 2 - ctlCalendar1.Width / 2
ctlCalendar1.Top = Me.Height / 2 - ctlCalendar1.Height / 2


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

Private Sub Text1_Change(Index As Integer)
Dim rc As Boolean
rc = IIf(Len(Text1(Index)) > 0, True, False)
Text1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbWhite, vbColorUnabled)
Select Case Index
    Case 3
        Frame2.Visible = Not (rc)
        Label1(22).Visible = Not (Frame2.Visible)
        Shape1(0).Visible = Label1(22).Visible

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
    ctlCalendar1.Visible = False
End If
End Sub
Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).BackColor = vbWhite
ctlCalendar1.ZOrder
IndexDate = Index
'ctlCalendar1.SetDate = FormatDateTime(Now, vbShortDate)

Select Case Index
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
    Case 23, 24, 29
        ' preparation
        ctlCalendar1.Left = Text1(Index).Left - ctlCalendar1.Width - 200
        ctlCalendar1.Top = PicMain(0).Top + Text1(Index).Top - ctlCalendar1.Height / 2
        ctlCalendar1.Visible = True
    Case 27
        ' preparation
        ctlCalendar1.Left = Text1(Index).Left + Text1(Index).Width + 200
        ctlCalendar1.Top = PicMain(0).Top + Text1(Index).Top - ctlCalendar1.Height / 2
        ctlCalendar1.Visible = True
    Case Else
        ctlCalendar1.Visible = False
        IndexDate = -1
End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = IIf(Len(Text1(Index)) > 0, vbWhite, vbColorUnabled)
End Sub

Private Function SaveProcedure()
Dim rc As Boolean

    rc = SaveForm
    PicMenu(3).Visible = rc
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
