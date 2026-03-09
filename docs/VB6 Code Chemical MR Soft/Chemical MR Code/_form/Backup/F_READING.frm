VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form F_READING 
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
   Picture         =   "F_READING.frx":0000
   ScaleHeight     =   12000
   ScaleMode       =   0  'User
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000080DF&
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   0
      Left            =   0
      MouseIcon       =   "F_READING.frx":1DED9
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
         Caption         =   "Set Standard"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   720
         MouseIcon       =   "F_READING.frx":1E1E3
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   0
         Left            =   1200
         Picture         =   "F_READING.frx":1E4ED
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00606060&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1455
         Left            =   6480
         TabIndex        =   102
         Top             =   240
         Visible         =   0   'False
         Width           =   8775
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   106
            Text            =   "3.23"
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   105
            Text            =   "0.4"
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   104
            Text            =   "2.4"
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   103
            Text            =   "3.09"
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000080DF&
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
            Height          =   465
            Index           =   39
            Left            =   6000
            TabIndex        =   112
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000080DF&
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
            Height          =   465
            Index           =   38
            Left            =   4080
            TabIndex        =   111
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000080DF&
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
            Height          =   465
            Index           =   37
            Left            =   1920
            TabIndex        =   110
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000080DF&
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
            Height          =   465
            Index           =   36
            Left            =   0
            TabIndex        =   109
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000080DF&
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
            Index           =   35
            Left            =   0
            TabIndex        =   108
            Top             =   0
            Width           =   3735
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H000080DF&
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
            Index           =   34
            Left            =   4080
            TabIndex        =   107
            Top             =   0
            Width           =   3735
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Left            =   10440
         TabIndex        =   30
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Left            =   8520
         TabIndex        =   28
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Left            =   6600
         TabIndex        =   26
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         TabIndex        =   22
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Top             =   840
         Width           =   1815
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
         Height          =   345
         Index           =   4
         Left            =   10440
         TabIndex        =   31
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
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
         Height          =   345
         Index           =   3
         Left            =   8520
         TabIndex        =   29
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "# Test"
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
         Left            =   6600
         TabIndex        =   27
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000080DF&
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
         TabIndex        =   23
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000080DF&
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
         TabIndex        =   21
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
         Index           =   0
         Left            =   0
         MouseIcon       =   "F_READING.frx":218CF
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
            MouseIcon       =   "F_READING.frx":21BD9
            MousePointer    =   99  'Custom
            Picture         =   "F_READING.frx":21EE3
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Set Standard"
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
            MouseIcon       =   "F_READING.frx":252C5
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
         MouseIcon       =   "F_READING.frx":255CF
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   13
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Test Specifications"
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
            MouseIcon       =   "F_READING.frx":258D9
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
            Picture         =   "F_READING.frx":25BE3
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
         MouseIcon       =   "F_READING.frx":28FC5
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   11
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Test Table "
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
            MouseIcon       =   "F_READING.frx":292CF
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   735
            MouseIcon       =   "F_READING.frx":295D9
            MousePointer    =   99  'Custom
            Picture         =   "F_READING.frx":298E3
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
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   0
      Left            =   0
      Picture         =   "F_READING.frx":2CCC5
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   6
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
         Index           =   30
         Left            =   11880
         TabIndex        =   98
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H000080DF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1575
         Left            =   3960
         TabIndex        =   89
         Top             =   4800
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
            Caption         =   "3 - Enter Test Info and Save"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   2640
            TabIndex        =   92
            Top             =   960
            Width           =   2805
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2 - Select PH : Click pH1 or pH2"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   2640
            TabIndex        =   91
            Top             =   600
            Width           =   3135
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1 - Enter STD number or select form List : Click SFG Standard"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   2640
            TabIndex        =   90
            Top             =   240
            Width           =   6165
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   1440
            Picture         =   "F_READING.frx":4AB9E
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
         Index           =   13
         Left            =   15720
         TabIndex        =   48
         Top             =   2760
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
         Index           =   12
         Left            =   13800
         TabIndex        =   46
         Top             =   2760
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
         Index           =   11
         Left            =   6960
         TabIndex        =   44
         Top             =   2760
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
         Left            =   5040
         TabIndex        =   42
         Top             =   2760
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
         Left            =   3120
         TabIndex        =   40
         Top             =   2760
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
         Left            =   8880
         TabIndex        =   38
         Top             =   1200
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
         Left            =   6960
         TabIndex        =   36
         Top             =   1200
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
         Left            =   5040
         TabIndex        =   34
         Top             =   1200
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
         Left            =   3120
         TabIndex        =   32
         Top             =   1200
         Width           =   1815
      End
      Begin FlexCell.Grid Grd1 
         Height          =   3480
         Left            =   2520
         TabIndex        =   96
         Top             =   3600
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
         BackColor       =   &H00A65911&
         Caption         =   "pH Value"
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
         Index           =   32
         Left            =   11880
         TabIndex        =   99
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Target Value "
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
         Left            =   3120
         TabIndex        =   95
         Top             =   1920
         Width           =   5655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   855
         Left            =   15480
         MouseIcon       =   "F_READING.frx":4DF80
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   240
         Width           =   2775
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
         MouseIcon       =   "F_READING.frx":4E28A
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   480
         Width           =   1815
      End
      Begin VB.Image Image4 
         Height          =   480
         Index           =   3
         Left            =   15840
         Picture         =   "F_READING.frx":4E594
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "pH2"
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
         Index           =   15
         Left            =   14760
         MouseIcon       =   "F_READING.frx":51976
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000080DF&
         Caption         =   "pH1"
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
         Left            =   11880
         MouseIcon       =   "F_READING.frx":51C80
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "ph max"
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
         Left            =   15720
         TabIndex        =   49
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "pH min"
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
         Left            =   13800
         TabIndex        =   47
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
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
         Left            =   6960
         TabIndex        =   45
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
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
         Left            =   5040
         TabIndex        =   43
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "VALUE ( ppm )"
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
         Left            =   3120
         TabIndex        =   41
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "+/- % "
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
         Left            =   8880
         TabIndex        =   39
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Accuracy"
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
         Left            =   6960
         TabIndex        =   37
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "MR STD"
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
         Index           =   6
         Left            =   5040
         TabIndex        =   35
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
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
         Height          =   345
         Index           =   5
         Left            =   3120
         TabIndex        =   33
         Top             =   840
         Width           =   1815
      End
      Begin VB.Image DisableImage 
         Height          =   480
         Left            =   9360
         Picture         =   "F_READING.frx":51F8A
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
   Begin VB.PictureBox PicMain 
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   1
      Left            =   0
      Picture         =   "F_READING.frx":5536C
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   24
      Top             =   2880
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
         Index           =   29
         Left            =   12480
         TabIndex        =   87
         Top             =   4320
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
         Left            =   6720
         TabIndex        =   85
         Top             =   4320
         Width           =   5655
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
         Left            =   3360
         TabIndex        =   83
         Top             =   4320
         Width           =   3255
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
         Left            =   15840
         TabIndex        =   81
         Top             =   2760
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
         Index           =   25
         Left            =   13920
         TabIndex        =   79
         Top             =   2760
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
         Index           =   24
         Left            =   12000
         TabIndex        =   77
         Top             =   2760
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
         Index           =   23
         Left            =   10080
         TabIndex        =   75
         Top             =   2760
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
         Index           =   22
         Left            =   6840
         TabIndex        =   73
         Top             =   2760
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
         Index           =   21
         Left            =   4440
         TabIndex        =   71
         Top             =   2760
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
         Index           =   20
         Left            =   2040
         TabIndex        =   69
         Top             =   2760
         Width           =   2295
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   13320
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   1080
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
         Index           =   18
         Left            =   10800
         TabIndex        =   63
         Top             =   1080
         Width           =   2415
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
         Index           =   17
         Left            =   8880
         TabIndex        =   61
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
         Index           =   16
         Left            =   6960
         TabIndex        =   59
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
         Index           =   15
         Left            =   5040
         TabIndex        =   57
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
         Index           =   14
         Left            =   3120
         TabIndex        =   55
         Top             =   1080
         Width           =   1815
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
         TabIndex        =   54
         Top             =   6600
         Width           =   3975
         Begin VB.Image ImageTAV 
            Height          =   480
            Index           =   5
            Left            =   1720
            MouseIcon       =   "F_READING.frx":73245
            MousePointer    =   99  'Custom
            Picture         =   "F_READING.frx":7354F
            Top             =   160
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
         Index           =   19
         Left            =   13320
         TabIndex        =   67
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
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
         Index           =   31
         Left            =   12480
         TabIndex        =   88
         Top             =   3960
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "Note"
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
         Left            =   6720
         TabIndex        =   86
         Top             =   3960
         Width           =   5655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "QC OPERATOR"
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
         Left            =   3360
         TabIndex        =   84
         Top             =   3960
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "WEIGHT (mg)"
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
         Left            =   15840
         TabIndex        =   82
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "TURBID."
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
         Left            =   13920
         TabIndex        =   80
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
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
         Height          =   465
         Index           =   26
         Left            =   12000
         TabIndex        =   78
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00A65911&
         Caption         =   "SPECTR. (ABS)"
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
         Left            =   10080
         TabIndex        =   76
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "METER 3 (ppm)"
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
         Left            =   6840
         TabIndex        =   74
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "METER 2 (ppm)"
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
         Left            =   4440
         TabIndex        =   72
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "METER 1 (ppm)"
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
         Left            =   2040
         TabIndex        =   70
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "HEAD"
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
         Left            =   13320
         TabIndex        =   65
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "MACHINE OP."
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
         Left            =   10800
         TabIndex        =   64
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "PROD. DATE"
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
         Left            =   8880
         TabIndex        =   62
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "TIME"
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
         Left            =   6960
         TabIndex        =   60
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "QC DATE"
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
         Index           =   17
         Left            =   5040
         TabIndex        =   58
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00964901&
         Caption         =   "TEST TYPE"
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
         Index           =   16
         Left            =   3120
         TabIndex        =   56
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.PictureBox PicMain 
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   2
      Left            =   0
      Picture         =   "F_READING.frx":76931
      ScaleHeight     =   7815
      ScaleWidth      =   19215
      TabIndex        =   25
      Top             =   2880
      Width           =   19215
      Begin FlexCell.Grid Grd2 
         Height          =   6600
         Left            =   480
         TabIndex        =   97
         Top             =   360
         Width           =   18240
         _ExtentX        =   32173
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
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   15720
         MouseIcon       =   "F_READING.frx":9480A
         MousePointer    =   99  'Custom
         TabIndex        =   101
         Top             =   7080
         Width           =   3015
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
         MouseIcon       =   "F_READING.frx":94B14
         MousePointer    =   99  'Custom
         TabIndex        =   100
         Top             =   7200
         Width           =   1995
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   360
         MouseIcon       =   "F_READING.frx":94E1E
         MousePointer    =   99  'Custom
         TabIndex        =   94
         Top             =   7080
         Width           =   3015
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete Selected Test"
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
         Left            =   1200
         MouseIcon       =   "F_READING.frx":95128
         MousePointer    =   99  'Custom
         TabIndex        =   93
         Top             =   7200
         Width           =   2070
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   480
         Picture         =   "F_READING.frx":95432
         Top             =   7100
         Width           =   480
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
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   4
      Left            =   2040
      MouseIcon       =   "F_READING.frx":98814
      MousePointer    =   99  'Custom
      Picture         =   "F_READING.frx":98B1E
      Top             =   11040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Index           =   0
      Left            =   17640
      MouseIcon       =   "F_READING.frx":9BF00
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
      MouseIcon       =   "F_READING.frx":9C20A
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
      MouseIcon       =   "F_READING.frx":9C514
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   10560
      Width           =   2655
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   15960
      MouseIcon       =   "F_READING.frx":9C81E
      MousePointer    =   99  'Custom
      Picture         =   "F_READING.frx":9CB28
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "F_READING.frx":9FF0A
      Height          =   480
      Index           =   1
      Left            =   9360
      MouseIcon       =   "F_READING.frx":A32EC
      MousePointer    =   99  'Custom
      Picture         =   "F_READING.frx":A35F6
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   480
      MouseIcon       =   "F_READING.frx":A69D8
      MousePointer    =   99  'Custom
      Picture         =   "F_READING.frx":A6CE2
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
      DragIcon        =   "F_READING.frx":AA0C4
      Height          =   480
      Index           =   0
      Left            =   18240
      MouseIcon       =   "F_READING.frx":AD4A6
      MousePointer    =   99  'Custom
      Picture         =   "F_READING.frx":AD7B0
      Top             =   11040
      Width           =   480
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   3
      Left            =   0
      MouseIcon       =   "F_READING.frx":B0B92
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   10560
      Width           =   1815
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
    SetCombo
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



Private Sub Combo1_Click()
If Combo1 = "Enter" Then
    Text1(19).SetFocus
Else
    Text1(19) = Combo1
End If

Combo1.Visible = False
End Sub


Private Sub Combo2_Click()
Combo2.Visible = False
If Combo2 = "Enter" Then
    Text1(14).SetFocus
    'Combo2.ListIndex = Combo2.ListCount - 1
Else
    Text1(14) = Combo2
End If

End Sub

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
Call SetGrid
End Sub

Private Sub Form_Load()
IndexFormProcedura = 99
End Sub

Private Sub Frame2_Click()
Frame2.Visible = False
End Sub

Private Sub Label10_Click()
Dim rc As Boolean
If Label7.Caption = "Show More" Then
    Label7.Caption = "Show Less"
    rc = True
Else
    Label7.Caption = "Show More"
    rc = False
End If
Call SetViewGrid2(rc)
End Sub

Private Sub PicMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Len(Text1(IndexText)) = 0 Then Text1(IndexText).BackColor = vbColorUnabled
IndexText = 0
Picture1.BackColor = &H4000&
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set F_READING = Nothing
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


Private Sub Label3_Click()
Grd1.ZOrder
Grd1.Visible = Not (Grd1.Visible)
End Sub

Private Sub Label6_Click()
PopupMessage 2, "Delete Selected Test", , , , DefaultMenu(4)
End Sub



Private Sub PicMenu_Click(Index As Integer)
If IndexFormProcedura = Index Then
'ElseIf Index = PicMenu.Count - 1 Then
   ' IndexMainProcedura = IndexMainProcedura + 1
   ' Unload Me
Else
    Grd1.Cell(0, 0).SetFocus
    Grd2.Cell(0, 0).SetFocus
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
    Frame3.Visible = False
    Select Case Index
        Case 0
            Picture4(0).BackColor = &H80DF&
            rc = False
        Case 1
            rc = True
            Picture4(0).BackColor = &H70DF&
            
            Text1(15) = FormatDataLAT(Date)
            Text1(16) = FormatTimeLAT(FormatDateTime(Now, vbShortTime))
        Case 2
            rc = False
            Picture4(0).BackColor = &H60DF&
            Frame3.Visible = True
    End Select
    For i = 2 To 4
        Label1(i).Visible = rc
        Text1(i).Visible = rc
    Next
    Label2(4) = Label2(Index)
    IndexFormProcedura = Index
    PicMain(Index).Visible = True
    PicMain(Index).ZOrder
    blTable = "Reading QC : " & Label2(IndexFormProcedura)
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
ctlCalendar1.ZOrder
IndexDate = Index
'ctlCalendar1.SetDate = FormatDateTime(Now, vbShortDate)

Select Case Index
    Case 14
        If Combo2 = "Enter" Then
            Exit Sub
        End If
        If Text1(Index) <> "" Then
            'If IsNumeric(Text1(Index)) And Text1(Index) > 0 And Text1(Index) < 4 Then
                'Combo1 = Text1(Index)
           ' End If
        End If
        Combo2.ZOrder
        Combo2.Visible = True
    Case 19
        If Combo1 = "Enter" Then Exit Sub
        If Text1(Index) <> "" Then
            If IsNumeric(Text1(Index)) And Text1(Index) > 0 And Text1(Index) < 4 Then
                Combo1 = Text1(Index)
            End If
        End If
        Combo1.ZOrder
        Combo1.Visible = True
    Case 3
    
        ctlCalendar1.Left = Text1(Index).Left
        ctlCalendar1.Top = Frame1.Top + Text1(Index).Top + Text1(Index).Height + 120
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

Private Sub SetCombo()
Dim i As Integer

    Combo1.Clear
    For i = 1 To 4
        Combo1.AddItem i
    Next
    Combo1.AddItem "Enter"
    Combo2.Clear
    'For i = 1 To 4
        Combo2.AddItem "OLD A"
        Combo2.AddItem "EPP"
        Combo2.AddItem "P"
        Combo2.AddItem "Enter"
    'Next
    
End Sub



Private Function SetGrid()

       '------------------------------------------------
        '       SET TABELLA Test
        '------------------------------------------------
    With Grd2
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        .DefaultFont.Size = 14
        .DefaultRowHeight = 40
        .Cols = 19
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Standard"
        .Column(1).Width = 150
        .Cell(0, 2).Text = "STD Value"
        .Column(2).Width = 150
        .Cell(0, 3).Text = "#"
        .Column(3).Width = 50
        .Cell(0, 4).Text = "TEST"
        .Column(4).Width = 150
        .Cell(0, 5).Text = "QC DATE"
        .Column(5).Width = 120
        .Cell(0, 6).Text = "TIME"
        .Column(6).Width = 100
        .Cell(0, 7).Text = "PROD. DATE"
        .Column(7).Width = 120
        .Cell(0, 8).Text = "MACHINE OPERATOR"
        .Column(8).Width = 200
        
        
        .Cell(0, 9).Text = "HEAD"
        .Column(9).Width = 80
        .Cell(0, 10).Text = "METER 1 [ppm]"
        .Column(10).Width = 170
        .Cell(0, 11).Text = "METER 2 [ppm]"
        .Column(11).Width = 170
        .Cell(0, 12).Text = "METER 3 [ppm]"
        .Column(12).Width = 170
        .Cell(0, 13).Text = "SPECTR. [ABS]"
        .Column(13).Width = 150
        
        .Cell(0, 14).Text = "pH"
        .Column(14).Width = 80
        .Cell(0, 15).Text = "TIRB."
        .Column(15).Width = 120
        .Cell(0, 16).Text = "WEIGHT [mg]"
        .Column(16).Width = 150
        .Cell(0, 17).Text = "QC OPERATOR"
        .Column(17).Width = 200
        .Cell(0, 18).Text = "NOTE"
        .Column(18).Width = 300
        
        
        
               
        
        
                      
        
               
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
        Next
        .ReadOnly = True
        .AutoRedraw = True
        .Refresh
        
    End With
    
    With Grd1
      .Rows = 1

        .AutoRedraw = False
        
        .DisplayFocusRect = False 'Show ComboBox arrow with a single click
        .DrawMode = cellOwnerDraw
        .SelectionMode = cellSelectionByRow
        .DefaultFont.Size = 14
        .DefaultRowHeight = 40
        .Cols = 16
        .Cell(0, 0).Text = "n."
        .Column(0).Width = 0
        .Cell(0, 1).Text = "Standard"
        .Column(1).Width = 150
        .Cell(0, 2).Text = "STD Value"
        .Column(2).Width = 150
        .Cell(0, 3).Text = "MR STD"
        .Column(3).Width = 100
        .Cell(0, 4).Text = "VALUE [ppm]"
        .Column(4).Width = 150
        .Cell(0, 5).Text = "MIN [ppm]"
        .Column(5).Width = 120
        .Cell(0, 6).Text = "MAX [ppm]"
        .Column(6).Width = 120
        
        
        .Cell(0, 7).Text = "ACCURACY"
        .Column(7).Width = 120
        .Cell(0, 8).Text = "+/- %"
        .Column(8).Width = 120
                
        .Cell(0, 9).Text = "pH1 Value"
        .Column(9).Width = 100
        .Cell(0, 10).Text = "pH1 MIN"
        .Column(10).Width = 100
        .Cell(0, 11).Text = "pH1 MAX"
        .Column(11).Width = 100
        .Cell(0, 12).Text = "pH2 Value"
        .Column(12).Width = 100
        .Cell(0, 13).Text = "pH2 MIN"
        .Column(13).Width = 100
        .Cell(0, 14).Text = "pH2 MAX"
        .Column(14).Width = 100

        .Cell(0, 15).Text = "ID"
        .Column(15).Width = 0

        
        For i = 1 To .Cols - 1
            .Column(i).Alignment = cellCenterCenter
            
        Next
       ' .ReadOnly = True
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
    

End Function
Private Sub SetViewGrid2(ByVal bValue As Boolean)
    With Grd2
        .Column(1).Width = IIf(bValue, 150, 0)
        .Column(4).Width = IIf(bValue, 150, 0)
        .Column(8).Width = IIf(bValue, 200, 0)
        .Column(9).Width = IIf(bValue, 80, 0)
        .Column(13).Width = IIf(bValue, 150, 0)
        .Column(15).Width = IIf(bValue, 120, 0)
        .Column(16).Width = IIf(bValue, 150, 0)
        .Column(17).Width = IIf(bValue, 200, 0)
        .Column(18).Width = IIf(bValue, 300, 0)
    End With
End Sub
