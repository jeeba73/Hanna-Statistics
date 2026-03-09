VERSION 5.00
Begin VB.Form frmGaussiana 
   BorderStyle     =   0  'None
   ClientHeight    =   11970
   ClientLeft      =   15
   ClientTop       =   -105
   ClientWidth     =   19170
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGaussiana.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmGaussiana.frx":6852
   ScaleHeight     =   12000
   ScaleMode       =   0  'User
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   35
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MouseIcon       =   "frmGaussiana.frx":2472B
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   42
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   735
            MouseIcon       =   "frmGaussiana.frx":24A35
            MousePointer    =   99  'Custom
            Picture         =   "frmGaussiana.frx":24D3F
            Top             =   180
            Width           =   480
         End
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
            Index           =   4
            Left            =   75
            MouseIcon       =   "frmGaussiana.frx":28121
            MousePointer    =   99  'Custom
            TabIndex        =   43
            Top             =   720
            Width           =   1860
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   1920
         MouseIcon       =   "frmGaussiana.frx":2842B
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   40
         Top             =   0
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
            Index           =   1
            Left            =   0
            MouseIcon       =   "frmGaussiana.frx":28735
            MousePointer    =   99  'Custom
            TabIndex        =   41
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   735
            MouseIcon       =   "frmGaussiana.frx":28A3F
            MousePointer    =   99  'Custom
            Picture         =   "frmGaussiana.frx":28D49
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
         MouseIcon       =   "frmGaussiana.frx":2C12B
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   38
         Top             =   0
         Width           =   1935
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
            MouseIcon       =   "frmGaussiana.frx":2C435
            MousePointer    =   99  'Custom
            TabIndex        =   39
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   735
            MouseIcon       =   "frmGaussiana.frx":2C73F
            MousePointer    =   99  'Custom
            Picture         =   "frmGaussiana.frx":2CA49
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
         MouseIcon       =   "frmGaussiana.frx":2FE2B
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   36
         Top             =   0
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
            MouseIcon       =   "frmGaussiana.frx":30135
            MousePointer    =   99  'Custom
            TabIndex        =   37
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   735
            MouseIcon       =   "frmGaussiana.frx":3043F
            MousePointer    =   99  'Custom
            Picture         =   "frmGaussiana.frx":30749
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Graph QC"
         BeginProperty Font 
            Name            =   "Whitney-Light"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   450
         Left            =   16875
         TabIndex        =   44
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   8280
      ScaleHeight     =   1335
      ScaleWidth      =   1215
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frame_DATI 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   19215
      Begin VB.CommandButton Command1 
         Caption         =   "Head4"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   8640
         TabIndex        =   33
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Head3"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   8640
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Head2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   7440
         TabIndex        =   31
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Head1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   7440
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   4
         Left            =   2640
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   29
         Top             =   645
         Width           =   135
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   3
         Left            =   5520
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   28
         Top             =   645
         Width           =   135
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   2
         Left            =   5520
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   27
         Top             =   1005
         Width           =   135
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   2640
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   26
         Top             =   360
         Width           =   135
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   0
         Left            =   5520
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   25
         Top             =   285
         Width           =   135
      End
      Begin VB.PictureBox PicZOOM 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   1
         Left            =   12960
         MouseIcon       =   "frmGaussiana.frx":33B2B
         MousePointer    =   99  'Custom
         Picture         =   "frmGaussiana.frx":33E35
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   6
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox PicZOOM 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   0
         Left            =   12960
         MouseIcon       =   "frmGaussiana.frx":37217
         MousePointer    =   99  'Custom
         Picture         =   "frmGaussiana.frx":37521
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   5
         Top             =   240
         Width           =   480
      End
      Begin VB.OptionButton Opt_GRAPH 
         Caption         =   "Gaussian"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Opt_GRAPH 
         Caption         =   "Gaussiana"
         Height          =   255
         Index           =   2
         Left            =   10440
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton Opt_GRAPH 
         Caption         =   "Average"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Out of Range n. :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   24
         Top             =   600
         Width           =   1725
      End
      Begin VB.Label lbl_ris 
         AutoSize        =   -1  'True
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   23
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2s :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   5760
         TabIndex        =   20
         Top             =   600
         Width           =   330
      End
      Begin VB.Label lbl_ris 
         AutoSize        =   -1  'True
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   6240
         TabIndex        =   19
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "3s :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   5760
         TabIndex        =   18
         Top             =   255
         Width           =   330
      End
      Begin VB.Label lbl_ris 
         AutoSize        =   -1  'True
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   6240
         TabIndex        =   17
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lbl_ris 
         AutoSize        =   -1  'True
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   2160
         TabIndex        =   16
         Top             =   960
         Width           =   180
      End
      Begin VB.Label lbl_ris 
         AutoSize        =   -1  'True
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   4440
         TabIndex        =   15
         Top             =   600
         Width           =   240
      End
      Begin VB.Label lbl_ris 
         AutoSize        =   -1  'True
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   6240
         TabIndex        =   14
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lbl_ris 
         AutoSize        =   -1  'True
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4440
         TabIndex        =   13
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lbl_ris 
         AutoSize        =   -1  'True
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "s :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   5760
         TabIndex        =   11
         Top             =   960
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Reference weight :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Out of Range % : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Average  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   8
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Checked num :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      MouseIcon       =   "frmGaussiana.frx":3A903
      MousePointer    =   99  'Custom
      ScaleHeight     =   3735
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   2520
      Width           =   4335
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "H1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   120
         Width           =   465
      End
      Begin VB.Label lbl_VALUE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   840
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   9615.023
      X2              =   9615.023
      Y1              =   120.301
      Y2              =   12030.08
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   4807.512
      X2              =   4807.512
      Y1              =   0
      Y2              =   11909.77
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   14422.54
      X2              =   14422.54
      Y1              =   0
      Y2              =   11909.77
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
      TabIndex        =   50
      Top             =   10200
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   480.751
      X2              =   18749.29
      Y1              =   10706.77
      Y2              =   10706.77
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   4
      Left            =   480
      MouseIcon       =   "frmGaussiana.frx":3AC0D
      MousePointer    =   99  'Custom
      Picture         =   "frmGaussiana.frx":3AF17
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   3
      Left            =   4560
      MouseIcon       =   "frmGaussiana.frx":3E2F9
      MousePointer    =   99  'Custom
      Picture         =   "frmGaussiana.frx":3E603
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "frmGaussiana.frx":419E5
      Height          =   480
      Index           =   1
      Left            =   9360
      MouseIcon       =   "frmGaussiana.frx":44DC7
      MousePointer    =   99  'Custom
      Picture         =   "frmGaussiana.frx":450D1
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      Height          =   480
      Index           =   2
      Left            =   14160
      MouseIcon       =   "frmGaussiana.frx":484B3
      MousePointer    =   99  'Custom
      Picture         =   "frmGaussiana.frx":487BD
      Top             =   11040
      Width           =   480
   End
   Begin VB.Image DefaultMenu 
      DragIcon        =   "frmGaussiana.frx":4BB9F
      Height          =   480
      Index           =   0
      Left            =   18240
      MouseIcon       =   "frmGaussiana.frx":4EF81
      MousePointer    =   99  'Custom
      Picture         =   "frmGaussiana.frx":4F28B
      Top             =   11040
      Width           =   480
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   4
      Left            =   0
      TabIndex        =   49
      Top             =   10680
      Width           =   1695
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1815
      Index           =   2
      Left            =   13080
      TabIndex        =   48
      Top             =   10200
      Width           =   2775
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Index           =   1
      Left            =   8280
      TabIndex        =   47
      Top             =   10560
      Width           =   2655
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   975
      Index           =   0
      Left            =   17640
      TabIndex        =   46
      Top             =   10800
      Width           =   1455
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Index           =   3
      Left            =   3840
      TabIndex        =   45
      Top             =   10560
      Width           =   1815
   End
End
Attribute VB_Name = "frmGaussiana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const GRID_SPACING = 120
Private Const HANDLE_WIDTH = 80
Private Const HANDLE_HALF_WIDTH = HANDLE_WIDTH / 2
Public Opt_GRAPHico As String
Public Sel_X
Public Sel_Y
Public NumWeights As String
Public sFileName As String
Public INDEX_G As Integer
Private INDEX_H As Integer
Private TipoGrafico As Integer
Private bActivate As Boolean
Private m_rc As Boolean
Private cMedia  As Double
Private cDeviazione As Double
Private cTotali As Integer
Private cUm As String
Private cNominale As Double
Private cQMR As Double
Private cT1 As Double
Private cT2 As Double
Private MaxDev As Double
Private x_height As Double
Private y_height As Double
Private xoffset As Double
Private percent_data_x As Double
Private yoffset As Double
Private percent_data_y As Double
Private xData As Double
Private yData As Double
Private Zero As Integer
Private pltColor As OLE_COLOR
Private x_data_screen As Double
Private y_data_screen As Double
Private StartTime As String
Private StopTime As String
Private cStep

Public Function DoShow(ByVal NumPesate As Integer) As Boolean
    Dim m_FlgLoading As Boolean
    Dim mIndex As Integer
    
    On Error GoTo ERR_SHOW
    
    m_rc = False
    
    If NumPesate = 0 Then GoTo ERR_END
    bActivate = True
    Picture1.Cls
    INDEX_G = 0
    INDEX_H = 1
    Opt_GRAPHico = ("Proiezione delle Medie")
    
    '---------------------------------------------------------
    ' numero di pesate
      NumWeights = NumPesate
    '---------------------------------------------------------
    'Call GetDatiLotto
    Call CheckHeads
    Call Dati_Lotto(INDEX_H)
    
    cStep = 5
    If (cNominale / 1000) > 1 Then
        cStep = cStep * 50
    Else
    End If


    Opt_GRAPH_Click 0
    
    Me.Show vbModal
    
    If m_rc = True Then
            'in caso di modifiche
    End If
    
    DoShow = m_rc
    
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_SHOW:
    MsgBox err.Description
    m_rc = False
    Resume ERR_END
End Function
Private Function CheckHeads() As Boolean
Dim i As Integer
Dim MyHeads As Integer
    
    MyHeads = MyWeightCheck.numHeads
    
    If MyPostazione.Department = "Powder" Then
    Else
        MyHeads = 1
        Command1(0).Caption = "N Final"
        Label2 = "N"
    End If
    
    
    For i = 0 To 3
        Command1(i).Visible = IIf(i <= MyHeads - 1, True, False)
    Next


End Function
Private Sub Command1_Click(Index As Integer)

INDEX_H = Index + 1
    Call Dati_Lotto(INDEX_H)
    Opt_GRAPH_Click INDEX_G
    Label2 = "H" & INDEX_H
    Command1(0).Enabled = True
    Command1(1).Enabled = True
    Command1(2).Enabled = True
    Command1(3).Enabled = True
    Command1(Index).Enabled = False
    Select Case Index
        
    End Select
End Sub



Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
    DropShadow Me.hWnd
End Sub
Private Sub TranslMenu()
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Dim Ctl As Control
   For Each Ctl In Me.Controls
      If TypeOf Ctl Is Menu Then
            Ctl.Caption = (Ctl.Caption)
      End If
   Next Ctl
End Sub
Private Sub Form_Initialize()
' inizio col visualizzare la gaussiana.....
On Local Error Resume Next

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Local Error GoTo err:
Select Case KeyCode
        Case vbKeyEscape
                Unload Me
        Case 38
                PicZOOM_Click 0
        Case 40
                PicZOOM_Click 1
        Case 35
                Dati_Lotto (INDEX_G)
                Form_Resize
        Case 37
                Opt_GRAPH(INDEX_G - 1).Value = True
        Case 39
                Opt_GRAPH(INDEX_G + 1).Value = True
End Select
err:
End Sub




Public Sub Dati_Lotto(Optional ByVal Index As Integer)
On Error GoTo ERR_DATI:

Dim cx As Double
Dim a As Integer
Dim t As Integer
Dim s As Integer
Dim nOutOfRange As Integer
'Index = 1
cMedia = 0
cDeviazione = 0
t = 0
s = 0
nOutOfRange = 0
Dim Tolerance As Double
Dim ToleranceNeg As Double

   Tolerance = MyWeightCheck.DataControllo.RefWeight + MyWeightCheck.DataControllo.s3
   ToleranceNeg = MyWeightCheck.DataControllo.RefWeight - MyWeightCheck.DataControllo.s3
   
    With MyWeightCheck


    '-----------------------------------------------------------------------------
    '
    '            calcolo media
    '
    '------------------------------------------------------------------------------
    
    
         For a = 1 To NumWeights
         
            If (.LotCheck(a).h(Index)) = "" Then
                 Debug.Print "dato nullo"
            Else
                If (.LotCheck(a).h(Index)) <= 0 Then
                    Debug.Print "dato negativo"
                Else
                    If (.LotCheck(a).bRecCCData) Then
                        cx = .LotCheck(a).h(Index)
                        t = t + 1
                        cMedia = cMedia + cx
                        
                        MyGraphicCheck.LotCheck(t) = .LotCheck(a)
                        
                        If cx > Tolerance Or cx < ToleranceNeg Then
                            nOutOfRange = nOutOfRange + 1
                        End If
                        If t = 1 Then StartTime = .LotCheck(a).Time
                    Else
                        Debug.Print "dato non registrato"
                    End If
                End If
            End If
         Next
         
         If t = 0 Then Exit Sub
         
         
         
        StopTime = .LotCheck(a - 1).Time
         cMedia = cMedia / t
        MyGraphicCheck.NumWeights = t
        
        .DataControllo.media(Index) = cMedia
        
    '-----------------------------------------------------------------------------
    '
    '            calcolo deviazione standard
    '
    '------------------------------------------------------------------------------
        s = 0
        MaxDev = 0
         For a = 1 To NumWeights
            If (.LotCheck(a).h(Index)) = "" Then
            Else
                 If (.LotCheck(a).h(Index)) <= 0 Then
                Else
                    If (.LotCheck(a).bRecCCData) Then
                        s = s + 1
                        cx = .LotCheck(a).h(Index)
                        cDeviazione = cDeviazione + (cMedia - cx) ^ 2
                        cDevSt(s, Index) = Sqr(cDeviazione / s)
                        If cDevSt(s, Index) > MaxDev Then MaxDev = cDevSt(s, Index)
                    End If
                 End If
            End If
        Next
        If nOutOfRange = 0 Then
         .DataControllo.OutOfRangeData(Index) = 0
         .DataControllo.OutOfRangeDataPerc(Index) = 0
        Else
         .DataControllo.OutOfRangeData(Index) = nOutOfRange
         .DataControllo.OutOfRangeDataPerc(Index) = FormatNumber(100 * nOutOfRange / t, 1)
        End If
        

        
        cDeviazione = Sqr(cDeviazione / t)
        
        
        .DataControllo.devst(Index) = cDeviazione
        
        
        cTotali = MyGraphicCheck.NumWeights
           
        cUm = "mg"
        cNominale = .DataControllo.RefWeight
        cQMR = .DataControllo.Tolerance
        cT1 = .DataControllo.s
        cT2 = .DataControllo.s2
        
    '-----------------------------------------------------------------------------
    '
    '            stampa a schermo i dati
    '
    '------------------------------------------------------------------------------
        lbl_ris(0) = t
        lbl_ris(1) = cNominale & cUm
        lbl_ris(2) = .DataControllo.OutOfRangeData(Index)
        lbl_ris(3) = Chr(177) & cQMR & cUm
        lbl_ris(4) = FormatNumber(cMedia, 2) & cUm
        lbl_ris(5) = .DataControllo.OutOfRangeDataPerc(Index) & "%"
        lbl_ris(7) = Chr(177) & cT1 & cUm
        lbl_ris(8) = Chr(177) & cT2 & cUm
           
           
    '-----------------------------------------------------------------------------
    '
    '            assegnazione variabili grafico > MyGraphicCheck
    '
    '------------------------------------------------------------------------------
           
        MyGraphicCheck.DataControllo = .DataControllo
        
        
        
End With

coef_graph = 0.9
coef_graph_2 = 1.1

If (cNominale / 1000) > 1 Then
    coef_graph = 0.99
    coef_graph_2 = 1.01
Else

End If

Xmin = Int((cNominale - cQMR) * coef_graph)
Xmax = Int((cNominale + cQMR) * coef_graph_2)
Ydev = IIf(MaxDev = 0, 1, MaxDev)
cYMax = (1 / (cDeviazione * Sqr(2 * 3.14)) * Exp(-(cMedia - cMedia) ^ 2 / (2 * cDeviazione ^ 2))) * cTotali
cYMax = IIf(cYMax = 0, cTotali, cYMax)

'InitGraph Picture1, 0, cTotali, Xmin, Xmax, 10, 10, TipoGrafico


ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_DATI:
    MsgBox err.Description
    Resume Next

End Sub



Private Sub Form_Resize()
On Local Error Resume Next

Picture1.Move 60, 60, Me.ScaleWidth - 120, Me.ScaleHeight - 120 - frame_DATI.Height
Picture2.Move Picture1.Left, Picture1.Top, Picture1.Width, Picture1.Height
frame_DATI.Move 60, Picture1.Top + Picture1.Height, Me.ScaleWidth - 120
PicZOOM(0).Left = frame_DATI.Width - PicZOOM(0).Width - 340
PicZOOM(1).Left = frame_DATI.Width - PicZOOM(1).Width - 340
If Opt_GRAPH(1) Then
            Label2.Visible = False
            Picture1.Cls
            InitGraph Picture1, Xmin, Xmax, 0, cYMax, 10, 10, TipoGrafico, StartTime, StopTime
            'Take X() and Y() and plot the data
            PlotData Picture1, INDEX_H
                
                DISEGNA_RETTE_Y Picture1, cNominale + MyWeightCheck.DataControllo.s3, vbBlack, 0, "+3s"
                DISEGNA_RETTE_Y Picture1, cNominale - MyWeightCheck.DataControllo.s3, vbBlack, 0, "-3s"
                DISEGNA_RETTE_Y Picture1, cNominale - cT1, vbGreen, 0, "-s"
                DISEGNA_RETTE_Y Picture1, cNominale - cT2, vbRed, 0, "-2s"
                DISEGNA_RETTE_Y Picture1, cNominale + cT1, vbGreen, 0, "+s"
                DISEGNA_RETTE_Y Picture1, cNominale + cT2, vbRed, 0, "+2s"
                DISEGNA_RETTE_Y Picture1, cNominale, vbBlue, 0, "STD"
                
                
ElseIf Opt_GRAPH(0) Then
        
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '                                MEDIE
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Picture1.Cls
            Label2.Visible = True
            InitGraph Picture1, 0, cTotali, Xmin, Xmax, 10, 10, TipoGrafico, StartTime, StopTime
    
                DisegnaMedie Picture1, INDEX_H
                DISEGNA_RETTE Picture1, cNominale + MyWeightCheck.DataControllo.s3, vbBlack, 0, cNominale + MyWeightCheck.DataControllo.s3, "+3s"
                DISEGNA_RETTE Picture1, cNominale - MyWeightCheck.DataControllo.s3, vbBlack, 0, cNominale - MyWeightCheck.DataControllo.s3, "-3s"
                DISEGNA_RETTE Picture1, cNominale - cT1, vbGreen, 4, cNominale - cT1, " -s"
                DISEGNA_RETTE Picture1, cNominale - cT2, vbRed, 4, cNominale - cT2, "-2s"
                DISEGNA_RETTE Picture1, cNominale + cT1, vbGreen, 4, cNominale + cT1, " +s"
                DISEGNA_RETTE Picture1, cNominale + cT2, vbRed, 4, cNominale + cT2, "+2s"
                DISEGNA_RETTE Picture1, cNominale, vbBlue, 0, cNominale, "ref"

                DISEGNA_RETTE Picture1, cMedia, &H800000, 0    ', FormatNumber(cMedia, 2)
                
Else

            Picture1.Cls
            InitGraph Picture1, 0, cTotali, 0, Ydev, 10, 10, TipoGrafico, StartTime, StopTime
            DisegnaDEVST Picture1, INDEX_H

End If

End Sub



Private Sub Form_Unload(Cancel As Integer)
bActivate = False
Set frmGaussiana = Nothing
End Sub

Private Sub lbl_VALUE_Click()
lbl_VALUE.Visible = False
End Sub

Private Sub mnuesci_Click()
Unload Me
End Sub

Private Sub mnustampa_Click()
Dim str_grafico As String
If Me.WindowState = 2 Then Me.WindowState = 0
str_grafico = Opt_GRAPHico & " Graph"
Call DefaultValue(str_grafico, INDEX_H)
Call PrintGrafico(Picture1, str_grafico, INDEX_H)
DoEvents

End Sub

Public Sub Opt_GRAPH_Click(Index As Integer)
If Opt_GRAPH(Index) Then Opt_GRAPHico = Opt_GRAPH(Index).Caption
On Local Error Resume Next
INDEX_G = Index
TipoGrafico = Index
lbl_VALUE.Visible = False
If Opt_GRAPH(Index) = True Then
    Select Case Index
        Case 1 ' gaussiana
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Label2.Visible = False
            Picture1.Cls
            InitGraph Picture1, Xmin, Xmax, 0, cYMax, 10, 10, TipoGrafico, StartTime, StopTime
            'Take X() and Y() and plot the data
            PlotData Picture1, INDEX_H
                
                DISEGNA_RETTE_Y Picture1, cNominale + MyWeightCheck.DataControllo.s3, vbBlack, 0, "+3s"
                DISEGNA_RETTE_Y Picture1, cNominale - MyWeightCheck.DataControllo.s3, vbBlack, 0, "-3s"
                DISEGNA_RETTE_Y Picture1, cNominale - cT1, vbGreen, 0, "-s"
                DISEGNA_RETTE_Y Picture1, cNominale - cT2, vbRed, 0, "-2s"
                DISEGNA_RETTE_Y Picture1, cNominale + cT1, vbGreen, 0, "+s"
                DISEGNA_RETTE_Y Picture1, cNominale + cT2, vbRed, 0, "+2s"
                DISEGNA_RETTE_Y Picture1, cNominale, vbBlue, 0, "STD"
                
            
            
        Case 0 ' medie
       
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Label2.Visible = True
            Picture1.Cls
            InitGraph Picture1, 0, cTotali, Xmin, Xmax, 10, 10, TipoGrafico, StartTime, StopTime
    
                DisegnaMedie Picture1, INDEX_H
                
                DISEGNA_RETTE Picture1, cNominale + MyWeightCheck.DataControllo.s3, vbBlack, 0, cNominale + MyWeightCheck.DataControllo.s3, "LCS"
                DISEGNA_RETTE Picture1, cNominale - MyWeightCheck.DataControllo.s3, vbBlack, 0, cNominale - MyWeightCheck.DataControllo.s3, "LCI"
                DISEGNA_RETTE Picture1, cNominale - cT1, vbGreen, 4, cNominale - cT1, " -s"
                DISEGNA_RETTE Picture1, cNominale - cT2, vbRed, 4, cNominale - cT2, "-2s"
                DISEGNA_RETTE Picture1, cNominale + cT1, vbGreen, 4, cNominale + cT1, " +s"
                DISEGNA_RETTE Picture1, cNominale + cT2, vbRed, 4, cNominale + cT2, "+2s"
                DISEGNA_RETTE Picture1, cNominale, vbBlue, 0, cNominale, "ref"
                DISEGNA_RETTE Picture1, cMedia, &H800000, 0   ', FormatNumber(cMedia, 2)

    
        Case 2 ' Deviazione STANDARD
            Label2.Visible = False
            Picture1.Cls
            InitGraph Picture1, 0, cTotali, 0, Ydev, 10, 10, TipoGrafico, StartTime, StopTime
            DisegnaDEVST Picture1, INDEX_H
                
    End Select
End If

End Sub

Private Sub Picture1_DblClick()
'If Me.WindowState = 2 Then
'    Me.WindowState = 0
'    Else
'    Me.WindowState = 2
'End If
'
End Sub

Private Sub Picture1_LostFocus()
lbl_VALUE.Visible = False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Local Error Resume Next
'Dim pesate As Integer
'Dim Valore As Double
'pesate = Int(X_Convrs(x))
'Valore = check_value(pesate)

'If Opt_GRAPH(0) Then
   ' lbl_VALUE.Caption = (" Pesata n.") & pesate & (" - Netto : ") & Valore & cUm & " "
'    Else
'    Exit Sub
'End If

'lbl_VALUE.Visible = True
'If x + lbl_VALUE.Width > Picture1.Width Then x = x - lbl_VALUE.Width
'lbl_VALUE.Move x, y + lbl_VALUE.Height + 60
'lbl_VALUE.ZOrder
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Sel_X = x
Sel_Y = y


End Sub
Private Function X_Convrs(x_data_screen)

            x_height = X_Axis_Max - X_Axis_Min
            X_Convrs = (x_data_screen - X0) / (X1 - X0) * x_height + X_Axis_Min
End Function
Private Function Y_Convrs(y_data_screen)

            y_height = Y_Axis_Max - Y_Axis_Min
            Y_Convrs = (y_data_screen - Y0) / (Y1 - Y0) * y_height + Y_Axis_Min
End Function

Public Sub DISEGNA_RETTE_Y(Frm As PictureBox, Value, Color, stile, stringa) ', XData As Double, YData As Double, pltColor As Long)
On Local Error Resume Next
                    Frm.DrawStyle = stile
                    Frm.CurrentX = ConvertiX(Value)
                    Frm.CurrentY = ConvertiY(0)
                    Frm.Line -(ConvertiX(Value), ConvertiY(cYMax)), Color
                    Frm.CurrentX = Frm.CurrentX - 50
                    Frm.CurrentY = Y0 + 500
                    Frm.Print stringa
    
                   
End Sub

Public Sub DISEGNA_RETTE(Frm As PictureBox, Value, Color, stile, Optional stringa, Optional stringa2) ', XData As Double, YData As Double, pltColor As Long)
                
                    Frm.DrawStyle = stile
                    Frm.CurrentX = ConvertiX(0)
                    Frm.CurrentY = ConvertiY(Value)
                    Frm.Line -(ConvertiX(cTotali), ConvertiY(Value)), Color
                    Frm.DrawStyle = stile
                    Frm.CurrentY = Frm.CurrentY - 100
                    Frm.CurrentX = X0 - 500
                    If Len(stringa) > 0 Then Frm.Print stringa
End Sub
Public Sub DISEGNA_RETTE_NOMINALE(Frm As PictureBox, Value, Color, stile, Optional stringa, Optional stringa2) ', XData As Double, YData As Double, pltColor As Long)
                
                    Frm.DrawStyle = stile
                    Frm.CurrentX = ConvertiX(0)
                    Frm.CurrentY = ConvertiY(Value)
                    Frm.DrawWidth = 1.5
                    Frm.Line -(ConvertiX(cTotali), ConvertiY(Value)), Color
                    Frm.DrawWidth = 1
                    Frm.CurrentY = Frm.CurrentY - 100
                   'If Len(stringa2) > 0 Then Frm.Print stringa2
                    Frm.CurrentX = X0 - 500
                    If Len(stringa) > 0 Then Frm.Print stringa
End Sub
Public Function ConvertiX(Valore)
On Local Error Resume Next
 x_height = X_Axis_Max - X_Axis_Min
 xoffset = Valore - X_Axis_Min
 percent_data_x = xoffset / x_height
 ConvertiX = X0 + ((X1 - X0) * percent_data_x)
End Function
Public Function ConvertiY(Valore)
On Local Error Resume Next
 y_height = Y_Axis_Max - Y_Axis_Min
 yoffset = Valore - Y_Axis_Min
 percent_data_y = yoffset / y_height
 ConvertiY = Y0 + ((Y1 - Y0) * percent_data_y)
End Function
Public Sub DisegnaMedie(Grafico As PictureBox, Optional ByVal Index As Integer) ', XData As Double, YData As Double, pltColor As Long)
Dim j As Integer
Dim xData As Double
Dim yData As Double
Dim x_data_screen
Dim y_data_screen
Dim pltColor

With MyGraphicCheck

'Index = 1
pltColor = vbBlue
        For j = 1 To .NumWeights  '+ 1
           xData = j
           'If .LotCheck(j).H(Index) = "" Then .LotCheck(j).H(Index) = Xmin
           'If .LotCheck(j).H(Index) < 0 Then .LotCheck(j).H(Index) = Xmin
           yData = CDbl(.LotCheck(j).h(Index))
        
                
                
                x_height = X_Axis_Max - X_Axis_Min
                y_height = Y_Axis_Max - Y_Axis_Min
                
                xoffset = xData - X_Axis_Min
                yoffset = yData - Y_Axis_Min
                
               ' If (xoffset >= 0) And (yoffset >= 0) Then
                    percent_data_x = xoffset / x_height
                    percent_data_y = yoffset / y_height
                    
                    x_data_screen = X0 + ((X1 - X0) * percent_data_x)
                    y_data_screen = Y0 - ((Y0 - Y1) * percent_data_y)
                        
                        
                    If j = 1 Then
                    
                            Grafico.CurrentX = x_data_screen
                            Grafico.CurrentY = y_data_screen
                    End If
                    
                    
                        If Pr_X = -100 Or Pr_Y = -100 Then
                            Pr_X = x_data_screen
                            Pr_Y = y_data_screen
                            'Exit Sub
                        End If
                       'Grafico.DrawWidth = 1.5
                        Grafico.Line -(x_data_screen, y_data_screen), pltColor
                       ' Grafico.DrawWidth = 1
                    Pr_X = x_data_screen
                    Pr_Y = y_data_screen
               ' Else

               ' End If

        Next j

        For j = 1 To .NumWeights '- 1
           xData = j
           If .LotCheck(j).h(Index) = "" Then .LotCheck(j).h(Index) = Xmin
           If .LotCheck(j).h(Index) < 0 Then .LotCheck(j).h(Index) = Xmin
           
           yData = CDbl(.LotCheck(j).h(Index))
        
       
            pltColor = IIf((Abs(yData) > Abs(cNominale + MyWeightCheck.DataControllo.s3)) Or (Abs(yData) < Abs(cNominale - MyWeightCheck.DataControllo.s3)), vbRed, vbBlue)
                
                
                x_height = X_Axis_Max - X_Axis_Min
                y_height = Y_Axis_Max - Y_Axis_Min
                
                xoffset = xData - X_Axis_Min
                yoffset = yData - Y_Axis_Min
                
             '   If (xoffset >= 0) And (yoffset >= 0) Then
                    percent_data_x = xoffset / x_height
                    percent_data_y = yoffset / y_height
                    
                    x_data_screen = X0 + ((X1 - X0) * percent_data_x)
                    y_data_screen = Y0 - ((Y0 - Y1) * percent_data_y)
                        
                        
                    If j = 1 Then
                    
                            Grafico.CurrentX = x_data_screen
                            Grafico.CurrentY = y_data_screen
                    
                    End If
                    
                        If Pr_X = -100 Or Pr_Y = -100 Then
                            Pr_X = x_data_screen
                            Pr_Y = y_data_screen
                            'Exit Sub
                        End If
  
            Grafico.Line (x_data_screen - HANDLE_HALF_WIDTH, y_data_screen - HANDLE_HALF_WIDTH)-Step(HANDLE_WIDTH, HANDLE_WIDTH), _
                pltColor, BF
            Grafico.Line (x_data_screen - HANDLE_HALF_WIDTH, y_data_screen - HANDLE_HALF_WIDTH)-Step(HANDLE_WIDTH, HANDLE_WIDTH), _
                , B

              '  End If
        Next j
        
    End With
End Sub


Public Sub PicZOOM_Click(Index As Integer)
Dim PreMin
Dim PreMax
PreMin = Xmin
PreMax = Xmax



            
            Select Case Index
                Case 0 ' INGRANDISCI
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                If Opt_GRAPH(2) Then
                    Ydev = Ydev - Ydev / 50
                Else
                    Xmin = Xmin + cStep
                    Xmax = Xmax - cStep
                End If
                 
                 
                Case 1 ' RIMPICCIOLISCI
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
               ' If cStep * 100 < 1 Then cStep = cStep * 100
                
                If Opt_GRAPH(2) Then
                    Ydev = Ydev + Ydev / 50
                Else
                    Xmin = Xmin - cStep
                    Xmax = Xmax + cStep
                End If
                  
            End Select
    
    ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
            If Xmin >= cMedia Then
                Xmin = PreMin
                Xmax = PreMax
                
              '  cStep = cStep / 50
                Exit Sub
            End If
            
            If Xmax <= cMedia Then
                Xmin = PreMin
                Xmax = PreMax
                  '  cStep = cStep / 50
                    Exit Sub
            End If
    ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Form_Resize
End Sub

Public Sub PlotData(Frm As PictureBox, Optional Index As Integer) ', XData As Double, YData As Double, pltColor As Long)
Dim j As Integer

        For j = Xmin To Xmax 'Step 0.04 * cStep
           xData = j
           yData = (1 / (MyWeightCheck.DataControllo.devst(Index) * Sqr(2 * 3.14)) * Exp(-(xData - MyWeightCheck.DataControllo.media(Index)) ^ 2 / (2 * MyWeightCheck.DataControllo.devst(Index) ^ 2))) * cTotali
        
            Zero = Frm.Top + Frm.Height - 760
                
                pltColor = vbBlue
                
                x_height = X_Axis_Max - X_Axis_Min
                y_height = Y_Axis_Max - Y_Axis_Min
                
                xoffset = xData - X_Axis_Min
                yoffset = yData - Y_Axis_Min
                
                If (xoffset >= 0) And (yoffset >= 0) Then
                    percent_data_x = xoffset / x_height
                    percent_data_y = yoffset / y_height
                    
                    x_data_screen = X0 + ((X1 - X0) * percent_data_x)
                    y_data_screen = Y0 - ((Y0 - Y1) * percent_data_y)
                    
                    If Pr_X = -100 Or Pr_Y = -100 Then
                        Pr_X = x_data_screen
                        Pr_Y = y_data_screen
                        'Exit Sub
                    End If
                    
                    Frm.Line (Pr_X, Pr_Y)-(x_data_screen, y_data_screen), pltColor
                    
                    Pr_X = x_data_screen
                    Pr_Y = y_data_screen
                End If
        Next j
    Exit Sub
err:
If err.NUMBER = 11 Then 'divisione per 0
DISEGNA_RETTE_Y Picture1, cNominale, vbRed, 0, ""

End If
End Sub







Public Sub DisegnaDEVST(Frm As PictureBox, Optional Index As Integer) ', XData As Double, YData As Double, pltColor As Long)
Dim j As Integer
        
        For j = 1 To MyGraphicCheck.NumWeights
           xData = j
           yData = cDevSt(j, Index)
        
                pltColor = vbBlue
                
                x_height = X_Axis_Max - X_Axis_Min
                y_height = Y_Axis_Max - Y_Axis_Min
                
                xoffset = xData - X_Axis_Min
                yoffset = yData - Y_Axis_Min
                
                If (xoffset >= 0) And (yoffset >= 0) Then
                    percent_data_x = xoffset / x_height
                    percent_data_y = yoffset / y_height
                    
                    x_data_screen = X0 + ((X1 - X0) * percent_data_x)
                    y_data_screen = Y0 - ((Y0 - Y1) * percent_data_y)
                        
                        
                    If j = 1 Then
                    
                            Frm.CurrentX = x_data_screen
                            Frm.CurrentY = y_data_screen
                    
                    End If
                    
                    
                        If Pr_X = -100 Or Pr_Y = -100 Then
                            Pr_X = x_data_screen
                            Pr_Y = y_data_screen
                            'Exit Sub
                        End If
        
                        Frm.Line -(x_data_screen, y_data_screen), pltColor
                   
                    Pr_X = x_data_screen
                    Pr_Y = y_data_screen
                End If

        Next j
        
        
        
        For j = 1 To MyGraphicCheck.NumWeights
           xData = j
           yData = cDevSt(j, Index)
        
                pltColor = vbBlue
                
                x_height = X_Axis_Max - X_Axis_Min
                y_height = Y_Axis_Max - Y_Axis_Min
                
                xoffset = xData - X_Axis_Min
                yoffset = yData - Y_Axis_Min
                
                If (xoffset >= 0) And (yoffset >= 0) Then
                    percent_data_x = xoffset / x_height
                    percent_data_y = yoffset / y_height
                    
                    x_data_screen = X0 + ((X1 - X0) * percent_data_x)
                    y_data_screen = Y0 - ((Y0 - Y1) * percent_data_y)
                        
                        
                    If j = 1 Then
                    
                            Frm.CurrentX = x_data_screen
                            Frm.CurrentY = y_data_screen
                    
                    End If
                    
                        If Pr_X = -100 Or Pr_Y = -100 Then
                            Pr_X = x_data_screen
                            Pr_Y = y_data_screen
                            'Exit Sub
                        End If
  
            Frm.Line (x_data_screen - HANDLE_HALF_WIDTH, y_data_screen - HANDLE_HALF_WIDTH)-Step(HANDLE_WIDTH, HANDLE_WIDTH), _
                vbWhite, BF
            Frm.Line (x_data_screen - HANDLE_HALF_WIDTH, y_data_screen - HANDLE_HALF_WIDTH)-Step(HANDLE_WIDTH, HANDLE_WIDTH), _
                , B
            
 
                End If
        Next j

End Sub

Private Function check_value(ByVal numero As Integer, Optional ByVal Index As Integer)
Index = 1
If numero < 1 Then Exit Function
If MyWeightCheck.LotCheck(numero).h(Index) = "" Then MyWeightCheck.LotCheck(numero).h(Index) = 0
    check_value = CDbl(MyWeightCheck.LotCheck(numero).h(Index))
End Function



