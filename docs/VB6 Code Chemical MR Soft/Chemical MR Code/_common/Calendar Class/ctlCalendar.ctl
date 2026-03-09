VERSION 5.00
Begin VB.UserControl ctlCalendar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   PropertyPages   =   "ctlCalendar.ctx":0000
   ScaleHeight     =   7410
   ScaleMode       =   0  'User
   ScaleWidth      =   8142.406
   ToolboxBitmap   =   "ctlCalendar.ctx":0031
   Begin VB.Frame fraDays 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   6090
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   8055
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   34
         Left            =   6840
         TabIndex        =   63
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   33
         Left            =   5760
         TabIndex        =   62
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   32
         Left            =   4680
         TabIndex        =   61
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   31
         Left            =   3600
         TabIndex        =   60
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   30
         Left            =   2520
         TabIndex        =   59
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   29
         Left            =   1440
         TabIndex        =   58
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   28
         Left            =   360
         TabIndex        =   57
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   27
         Left            =   6840
         TabIndex        =   56
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   26
         Left            =   5760
         TabIndex        =   55
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   25
         Left            =   4680
         TabIndex        =   54
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   24
         Left            =   3600
         TabIndex        =   53
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   23
         Left            =   2535
         TabIndex        =   52
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   22
         Left            =   1440
         TabIndex        =   51
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   21
         Left            =   360
         TabIndex        =   50
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   20
         Left            =   6840
         TabIndex        =   49
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   19
         Left            =   5760
         TabIndex        =   48
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   18
         Left            =   4680
         TabIndex        =   47
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   17
         Left            =   3600
         TabIndex        =   46
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   16
         Left            =   2520
         TabIndex        =   45
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   15
         Left            =   1440
         TabIndex        =   44
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   14
         Left            =   360
         TabIndex        =   43
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   13
         Left            =   6840
         TabIndex        =   42
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   12
         Left            =   5760
         TabIndex        =   41
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   11
         Left            =   4680
         TabIndex        =   40
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   10
         Left            =   3600
         TabIndex        =   39
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   9
         Left            =   2520
         TabIndex        =   38
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   8
         Left            =   1440
         TabIndex        =   37
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   7
         Left            =   360
         TabIndex        =   36
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   6
         Left            =   6840
         TabIndex        =   35
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   5
         Left            =   5760
         TabIndex        =   34
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   4
         Left            =   4680
         TabIndex        =   33
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   3
         Left            =   3600
         TabIndex        =   32
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   2
         Left            =   2520
         TabIndex        =   31
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   1
         Left            =   1440
         TabIndex        =   30
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   0
         Left            =   360
         TabIndex        =   29
         Top             =   840
         Width           =   615
      End
      Begin VB.Shape shpToday 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         FillColor       =   &H00FF0000&
         Height          =   705
         Index           =   1
         Left            =   6360
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sat"
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
         Index           =   6
         Left            =   6840
         TabIndex        =   20
         Top             =   75
         Width           =   615
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fri"
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
         Index           =   5
         Left            =   5760
         TabIndex        =   19
         Top             =   75
         Width           =   495
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thu"
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
         Index           =   4
         Left            =   4560
         TabIndex        =   18
         Top             =   75
         Width           =   705
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wed"
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
         Index           =   3
         Left            =   3480
         TabIndex        =   17
         Top             =   75
         Width           =   885
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tue"
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
         Index           =   2
         Left            =   2400
         TabIndex        =   16
         Top             =   75
         Width           =   705
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mon"
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
         Index           =   1
         Left            =   1320
         TabIndex        =   15
         Top             =   75
         Width           =   855
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sun"
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
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   75
         Width           =   705
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   35
         Left            =   360
         TabIndex        =   13
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   36
         Left            =   1440
         TabIndex        =   12
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   37
         Left            =   2520
         TabIndex        =   11
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   38
         Left            =   3600
         TabIndex        =   10
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   39
         Left            =   4680
         TabIndex        =   9
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   40
         Left            =   5760
         TabIndex        =   8
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label lblDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "30"
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
         Height          =   525
         Index           =   41
         Left            =   6840
         TabIndex        =   7
         Top             =   5400
         Width           =   615
      End
      Begin VB.Shape shpSelected 
         BorderColor     =   &H00404040&
         FillColor       =   &H00644603&
         FillStyle       =   0  'Solid
         Height          =   705
         Left            =   6240
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdBack 
      Height          =   600
      Left            =   1800
      Picture         =   "ctlCalendar.ctx":0343
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7920
      Width           =   660
   End
   Begin VB.CommandButton cmdFwd 
      Height          =   600
      Left            =   4200
      Picture         =   "ctlCalendar.ctx":068A
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7920
      Width           =   675
   End
   Begin VB.Frame fraWeek 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1280
      Left            =   7560
      TabIndex        =   21
      Top             =   3840
      Width           =   358
      Begin VB.Label lblWeeks 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "30"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   27
         Top             =   1100
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblWeeks 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "30"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   26
         Top             =   885
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblWeeks 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "30"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   25
         Top             =   660
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblWeeks 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "30"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblWeeks 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "30"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   23
         Top             =   225
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblWeeks 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "30"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   345
      End
   End
   Begin ChemicalMR.ctlDateScroll ctlDateScroll 
      Height          =   645
      Left            =   840
      TabIndex        =   66
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1138
      Alignment       =   2
      Text            =   "01/01/2004"
      BackColor       =   6571523
      ForeColor       =   16777215
      DateMode        =   0
      Month           =   "January"
      Year            =   "01/01/2004"
   End
   Begin VB.Line lineMain 
      X1              =   482.513
      X2              =   3206.7
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Image1Label 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Index           =   0
      Left            =   -120
      MouseIcon       =   "ctlCalendar.ctx":09D0
      MousePointer    =   99  'Custom
      TabIndex        =   65
      Top             =   -240
      Width           =   1455
   End
   Begin VB.Label Image1Label 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Index           =   1
      Left            =   6840
      MouseIcon       =   "ctlCalendar.ctx":0CDA
      MousePointer    =   99  'Custom
      TabIndex        =   64
      Top             =   -240
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   7440
      MouseIcon       =   "ctlCalendar.ctx":0FE4
      MousePointer    =   99  'Custom
      Picture         =   "ctlCalendar.ctx":12EE
      Top             =   180
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   240
      MouseIcon       =   "ctlCalendar.ctx":46D0
      MousePointer    =   99  'Custom
      Picture         =   "ctlCalendar.ctx":49DA
      Top             =   180
      Width           =   480
   End
   Begin VB.Line lineWeekNums 
      Visible         =   0   'False
      X1              =   7599.579
      X2              =   7599.579
      Y1              =   3840
      Y2              =   5140
   End
   Begin VB.Shape shpHighlight 
      BorderStyle     =   3  'Dot
      Height          =   255
      Left            =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblYear 
      BackStyle       =   0  'Transparent
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
      Height          =   615
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblTodayShape 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape shpToday 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H00FF0000&
      Height          =   215
      Index           =   0
      Left            =   60
      Shape           =   2  'Oval
      Top             =   2060
      Width           =   320
   End
   Begin VB.Label lblToday 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oggi : 02/05/2004"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   435
      TabIndex        =   3
      Top             =   2055
      Width           =   1605
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "February "
      BeginProperty Font 
         Name            =   "Whitney-Light"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   1080
      TabIndex        =   2
      Top             =   60
      Width           =   6135
   End
   Begin VB.Label lblBackground 
      BackColor       =   &H00644603&
      Height          =   840
      Left            =   -120
      TabIndex        =   28
      Top             =   0
      Width           =   9930
   End
   Begin VB.Menu mnuAlarmsMain 
      Caption         =   "Alarms"
      Begin VB.Menu mnuToday 
         Caption         =   "Go to Today"
      End
      Begin VB.Menu mnuAddAlarm 
         Caption         =   "Add New Alarm"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAlarm 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "ctlCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Option Explicit
Dim m_CurrentDate As Date
Dim m_LastSelected As Integer
Dim m_SelectedColor As OLE_COLOR
Dim m_AlarmColor As OLE_COLOR
Dim m_LastNextDaysColor As OLE_COLOR
Dim m_WeekNumberColor As OLE_COLOR
Dim m_Alarms As New cAlarmGroup
Dim m_UseAlarms As Boolean
Dim m_ForegroundColor As OLE_COLOR
Dim m_Today As Boolean
Dim m_WeekStartsWith As VbDayOfWeek
Dim m_ShowLastMonthDays As Boolean
Dim m_ShowNextMonthDays As Boolean
Dim m_DateOffset As Integer
Dim m_ToolTipText As Boolean
Dim m_AllowRightClick As Boolean
Dim m_LastClicked As Integer
Dim m_ShowWeekNumber As Boolean
Dim m_ShowWeekNumberLeft As Boolean
Dim m_ShowLastMonthButton As Boolean
Dim m_ShowNextMonthButton As Boolean
Dim m_ShowSelected As Boolean
Dim m_ShowShortDays As Boolean

Dim WithEvents m_HeaderFont As StdFont
Attribute m_HeaderFont.VB_VarHelpID = -1
Dim WithEvents m_DayFont As StdFont
Attribute m_DayFont.VB_VarHelpID = -1
Dim WithEvents m_TodayFont As StdFont
Attribute m_TodayFont.VB_VarHelpID = -1
Dim WithEvents m_ColumnFont As StdFont
Attribute m_ColumnFont.VB_VarHelpID = -1

Public Event DateClicked(inputDate As Date)
Public Event DateDblClicked(inputDate As Date)
Public Event MonthChanged(inputDate As Date)
Public Event WeekHeadingClicked(weekday As VbDayOfWeek)
Public Event WeekHeadingDblClicked(weekday As VbDayOfWeek)
Public Event MonthHeadingClicked(inputDate As Date)
Public Event MonthHeadingDblClicked(inputDate As Date)
Public Event LastButtonClicked(inputDate As Date)
Public Event NextButtonClicked(inputDate As Date)
Public Event TodayClicked(inputDate As Date)
Public Event AlarmSelected(UID As Double)
Public Event AddNewAlarm(inputDate As Date)
Public Event WeekNumberClicked(weekNumber As Integer)
Public Event WeekNumberDblClicked(weekNumber As Integer)

Public Sub ShowDate(inputDate As Date, Optional alarmDays As cAlarmGroup)
    m_CurrentDate = inputDate
    If Not alarmDays Is Nothing Then
        Set m_Alarms = alarmDays
    End If
    Call SetDate
End Sub
Public Sub SetAlarms(cAlarms As cAlarmGroup)
    Set m_Alarms = cAlarms
    Call SetDate
End Sub
Public Sub SetDate()
    Dim nCount As Integer
    Dim nAlarm As Integer
    Dim nStartingDate As Date
    Dim nDayOfWeek As Integer
    Dim alarmDays As cAlarmGroup
    Dim dHoldDate As Date
    Dim nAlarmCount As Integer
    
    UserControl.AutoRedraw = False
        
    m_LastSelected = -1
    lblMonth.Caption = Format(m_CurrentDate, "mmmm yyyy")
        
    If m_WeekStartsWith = vbUseSystemDayOfWeek Then
        m_WeekStartsWith = vbSunday
    End If
    For nCount = lblHeader.LBound To lblHeader.UBound
        If m_ShowShortDays = True Then
            lblHeader(nCount).Caption = Mid(WeekdayName(((nCount + m_WeekStartsWith - 1) Mod 7) + 1, True), 1, 1)
        Else
            lblHeader(nCount).Caption = WeekdayName(((nCount + m_WeekStartsWith - 1) Mod 7) + 1, True)
        End If
    Next nCount
    
    If BackwardsDate = False Then
        nDayOfWeek = Format(Format(m_CurrentDate, "mm/01/yyyy"), "w")
    Else
        nDayOfWeek = Format(Format(m_CurrentDate, "1/mm/yyyy"), "w") - 1
    End If
    nDayOfWeek = (nDayOfWeek + 7 - m_WeekStartsWith) Mod 7
    If nDayOfWeek = 0 Then
        nDayOfWeek = 7
    End If
    If BackwardsDate = False Then
        nStartingDate = DateAdd("d", -nDayOfWeek, Format(m_CurrentDate, "mm/01/yyyy"))
    Else
        nStartingDate = DateAdd("d", -nDayOfWeek, Format(m_CurrentDate, "1/mm/yyyy"))
    End If
    
    nAlarm = 1
    m_DateOffset = 0
    Set alarmDays = m_Alarms.GetAlarmDays(m_CurrentDate)
    For nCount = 0 To 41
        lblDays(nCount).ToolTipText = ""
        lblDays(nCount).Caption = Format(nStartingDate, "d")
        If Format(nStartingDate, "mm/yy") = Format(m_CurrentDate, "mm/yy") Then
            If m_DateOffset = 0 Then
                m_DateOffset = nCount
            End If
            nAlarmCount = 0
            If m_UseAlarms = False Or (m_UseAlarms = True And nAlarm > alarmDays.Count) Then
                lblDays(nCount).ForeColor = m_ForegroundColor
                lblDays(nCount).Font.Bold = False
                lblDays(nCount).Tag = "0"
            ElseIf nAlarm <= alarmDays.Count Then
                Do While CInt(Format(alarmDays(nAlarm).dateTime, "dd")) = CInt(Format(nStartingDate, "dd"))
                    nAlarmCount = nAlarmCount + 1
                    nAlarm = nAlarm + 1
                    If nAlarm > alarmDays.Count Then
                        Exit Do
                    End If
                Loop
                If nAlarmCount > 0 Then
                    lblDays(nCount).ForeColor = m_AlarmColor
                    lblDays(nCount).Font.Bold = True
                    lblDays(nCount).Tag = "1"
                    If m_ToolTipText = True Then
                        lblDays(nCount).ToolTipText = nAlarmCount & " Alarm" & IIf(nAlarmCount > 1, "s", "") & IIf(m_AllowRightClick = True, " - Right Click for Details", "")
                    End If
                Else
                    lblDays(nCount).ForeColor = m_ForegroundColor
                    lblDays(nCount).Font.Bold = False
                    lblDays(nCount).Tag = "0"
                End If
            End If
            
            If DatePart("d", Date) = DatePart("d", nStartingDate) Then
                shpToday(1).Left = ((lblDays(nCount).Left + lblDays(nCount).Width / 2) - shpToday(1).Width / 2) + 7
                shpToday(1).Top = ((lblDays(nCount).Top + lblDays(nCount).Height / 2) - shpToday(1).Height / 2) + 5
            End If
            lblDays(nCount).Visible = True
        Else
            If (nCount < 10 And m_ShowLastMonthDays = True) Or (nCount > 28 And m_ShowNextMonthDays = True) Then
                lblDays(nCount).Visible = True
                lblDays(nCount).Font.Bold = False
                lblDays(nCount).ForeColor = m_LastNextDaysColor
                lblDays(nCount).Tag = "-1"
            Else
                lblDays(nCount).Visible = False
                lblDays(nCount).Font.Bold = False
                lblDays(nCount).Tag = "-1"
            End If
        End If
        
        nStartingDate = DateAdd("d", 1, nStartingDate)
    
    Next nCount
    
    If m_ShowWeekNumber = True Then
        If BackwardsDate = False Then
            nStartingDate = DateAdd("d", -nDayOfWeek, Format(m_CurrentDate, "mm/01/yyyy"))
        Else
            nStartingDate = DateAdd("d", -nDayOfWeek, Format(m_CurrentDate, "1/mm/yyyy"))
        End If
        For nCount = lblWeeks.LBound To lblWeeks.UBound
            lblWeeks(nCount).Caption = Format(nStartingDate, "ww", m_WeekStartsWith, vbFirstFourDays)
'            If CInt(lblWeeks(nCount).Caption) > 52 Then
'                lblWeeks(nCount).Caption = "1"
'            End If
            lblWeeks(nCount).ForeColor = m_WeekNumberColor
            nStartingDate = DateAdd("d", 7, nStartingDate)
        Next nCount
    End If
    
    If ColorBackground = ColorToday Or m_Today = False Then
        shpToday(0).Visible = False
        shpToday(1).Visible = False
        lblToday.Left = 80
    Else
        shpToday(0).Visible = True
        shpToday(1).Visible = IIf(Format(Date, "mm/yy") = Format(m_CurrentDate, "mm/yy"), True, False)
        lblToday.Left = 440
    End If
    lblTodayShape.Visible = m_Today
    lblToday.Visible = m_Today

    lblToday.Caption = "Oggi : " & FormatDateTime(Date, vbShortDate)
    SetShape

    UserControl.AutoRedraw = True
    Set alarmDays = Nothing
End Sub
Private Function BackwardsDate() As Boolean
    If DateAdd("m", 1, CDate("1/1/2000")) = CDate("2/1/2000") Then
        BackwardsDate = False
    Else
        BackwardsDate = True
    End If
End Function

Private Sub cmdBack_Click()
    ctlDateScroll.Visible = False
    m_CurrentDate = DateAdd("m", -1, m_CurrentDate)
    Call SetDate
    RaiseEvent LastButtonClicked(m_CurrentDate)
End Sub

Private Sub cmdFwd_Click()
    ctlDateScroll.Visible = False
    m_CurrentDate = DateAdd("m", 1, m_CurrentDate)
    Call SetDate
    RaiseEvent NextButtonClicked(m_CurrentDate)
End Sub

Private Sub ctlDateScroll_Change()
    If IsDate(ctlDateScroll.Month & " " & Format(m_CurrentDate, "dd") & " ," & ctlDateScroll.Year) = True Then
        ShowDate ctlDateScroll.Month & " " & Format(m_CurrentDate, "dd") & " ," & ctlDateScroll.Year
        RaiseEvent MonthChanged(m_CurrentDate)
    End If
End Sub

Private Sub ctlDateScroll_LostFocus()
    ctlDateScroll.Visible = False
End Sub

Private Sub Image1_Click(Index As Integer)
If Index = 1 Then cmdFwd_Click
If Index = 0 Then cmdBack_Click
End Sub

Private Sub Image1Label_Click(Index As Integer)
Image1_Click Index
End Sub

Private Sub lblBackground_Click()
    ctlDateScroll.Visible = False
    RaiseEvent MonthHeadingClicked(m_CurrentDate)
End Sub

Private Sub lblBackground_DblClick()
    RaiseEvent MonthHeadingDblClicked(m_CurrentDate)
End Sub

Private Sub SetShape(Optional newSelection As Integer = -1)
    If shpSelected.Visible = True Then
        shpSelected.Visible = False
        If m_LastSelected <> -1 Then
            If lblDays(m_LastSelected).Tag = "1" Then
                lblDays(m_LastSelected).ForeColor = m_AlarmColor
            Else
                lblDays(m_LastSelected).ForeColor = m_ForegroundColor
            End If
        End If
        
        If newSelection = -1 Then
            m_LastSelected = m_DateOffset + DatePart("d", m_CurrentDate) - 1
        Else
            m_LastSelected = newSelection
        End If
        
        shpSelected.Left = ((lblDays(m_LastSelected).Left + lblDays(m_LastSelected).Width / 2) - shpSelected.Width / 2) + 7
        shpSelected.Top = ((lblDays(m_LastSelected).Top + lblDays(m_LastSelected).Height / 2) - shpSelected.Height / 2) + 5
        lblDays(m_LastSelected).ForeColor = m_SelectedColor
        m_ShowSelected = True
        shpSelected.Visible = True
    End If
End Sub

Private Sub lblDays_DblClick(Index As Integer)
    Dim nRaiseType As Integer
        
    ctlDateScroll.Visible = False
    m_ShowSelected = True
    shpSelected.Visible = True
    If lblDays(Index).Tag <> "-1" Then
        If BackwardsDate = False Then
            m_CurrentDate = CDate(Format(m_CurrentDate, "m") & "/" & lblDays(Index).Caption & "/" & Format(m_CurrentDate, "yy"))
        Else
            m_CurrentDate = CDate(lblDays(Index).Caption & "/" & Format(m_CurrentDate, "m") & "/" & Format(m_CurrentDate, "yy"))
        End If
        nRaiseType = 1
    Else
        If BackwardsDate = False Then
            If Index <= 7 Then
                m_CurrentDate = CDate(Format(DateAdd("m", -1, m_CurrentDate), "m") & "/" & lblDays(Index).Caption & "/" & Format(DateAdd("m", -1, m_CurrentDate), "yy"))
            Else
                m_CurrentDate = CDate(Format(DateAdd("m", 1, m_CurrentDate), "m") & "/" & lblDays(Index).Caption & "/" & Format(DateAdd("m", 1, m_CurrentDate), "yy"))
            End If
        Else
            If Index <= 7 Then
                m_CurrentDate = CDate(lblDays(Index).Caption & "/" & Format(DateAdd("m", -1, m_CurrentDate), "m") & "/" & Format(DateAdd("m", -1, m_CurrentDate), "yy"))
            Else
                m_CurrentDate = CDate(lblDays(Index).Caption & "/" & Format(DateAdd("m", 1, m_CurrentDate), "m") & "/" & Format(DateAdd("m", 1, m_CurrentDate), "yy"))
            End If
        End If
        SetDate
        nRaiseType = 2
    End If
    SetShape
        
    If nRaiseType = 1 And m_LastClicked = CInt(Format(m_CurrentDate, "dd")) Then
        RaiseEvent DateDblClicked(m_CurrentDate)
    ElseIf nRaiseType = 2 Then
        RaiseEvent MonthChanged(m_CurrentDate)
    End If
End Sub

Private Sub lblDays_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cEvents As cAlarmGroup
    Dim nCount As Integer
    Dim nRaiseType As Integer
    Dim bIsActive As Boolean
    
    ctlDateScroll.Visible = False
    m_ShowSelected = True
    shpSelected.Visible = True
    bIsActive = IIf(lblDays(Index).Tag = "-1", False, True)
    If lblDays(Index).Tag <> "-1" Then
        If BackwardsDate = False Then
            m_CurrentDate = CDate(Format(m_CurrentDate, "m") & "/" & lblDays(Index).Caption & "/" & Format(m_CurrentDate, "yy"))
        Else
            m_CurrentDate = CDate(lblDays(Index).Caption & "/" & Format(m_CurrentDate, "m") & "/" & Format(m_CurrentDate, "yy"))
        End If
        nRaiseType = 1
    Else
        If BackwardsDate = False Then
            If Index <= 7 Then
                m_CurrentDate = CDate(Format(DateAdd("m", -1, m_CurrentDate), "m") & "/" & lblDays(Index).Caption & "/" & Format(DateAdd("m", -1, m_CurrentDate), "yy"))
            Else
                m_CurrentDate = CDate(Format(DateAdd("m", 1, m_CurrentDate), "m") & "/" & lblDays(Index).Caption & "/" & Format(DateAdd("m", 1, m_CurrentDate), "yy"))
            End If
        Else
            If Index <= 7 Then
                m_CurrentDate = CDate(lblDays(Index).Caption & "/" & Format(DateAdd("m", -1, m_CurrentDate), "m") & "/" & Format(DateAdd("m", -1, m_CurrentDate), "yy"))
            Else
                m_CurrentDate = CDate(lblDays(Index).Caption & "/" & Format(DateAdd("m", 1, m_CurrentDate), "m") & "/" & Format(DateAdd("m", 1, m_CurrentDate), "yy"))
            End If
        End If
        SetDate
        nRaiseType = 2
        
    End If
    m_LastClicked = CInt(Format(m_CurrentDate, "dd"))
    SetShape
    DoEvents
    
    If Button = 1 Then
        If nRaiseType = 1 Then
            RaiseEvent DateClicked(m_CurrentDate)
        Else
            RaiseEvent MonthChanged(m_CurrentDate)
        End If

    ElseIf Button = 2 And UseAlarms = True And m_AllowRightClick = True And bIsActive = True Then
        mnuSep.Visible = False
        Set cEvents = m_Alarms.GetEvents(m_CurrentDate, ccDaily)
        For nCount = 1 To cEvents.Count
            mnuSep.Visible = True
            If mnuAlarm.Count < nCount Then
                Load mnuAlarm(nCount - 1)
            End If
            mnuAlarm(nCount - 1).Visible = True
            mnuAlarm(nCount - 1).Caption = Format(cEvents(nCount).dateTime, "hh:mm AMPM") & " - " & ShowFormat(cEvents(nCount).message)
            mnuAlarm(nCount - 1).Tag = cEvents(nCount).UID
        Next nCount
        
        If mnuAlarm.Count > cEvents.Count Then
            For nCount = cEvents.Count To mnuAlarm.Count - 1
                If nCount > 0 Then
                    Unload mnuAlarm(nCount)
                End If
            Next nCount
        End If
        
        If nRaiseType = 2 Then
            RaiseEvent MonthChanged(m_CurrentDate)
        End If
        DoEvents
        mnuAddAlarm.Visible = True
        If cEvents.Count = 0 Then
            mnuSep.Visible = False
            mnuAlarm(0).Visible = False
        Else
        End If
        PopupMenu mnuAlarmsMain
        Set cEvents = Nothing
    ElseIf Button = 2 And UseAlarms = False And m_AllowRightClick = True And bIsActive = True Then
        If mnuAlarm.Count > 1 Then
            For nCount = 1 To mnuAlarm.Count - 1
                Unload mnuAlarm(nCount)
            Next nCount
        End If
        mnuAddAlarm.Visible = False
        mnuSep.Visible = False
        mnuAlarm(0).Visible = False
        PopupMenu mnuAlarmsMain
    Else
        If nRaiseType = 1 Then
            RaiseEvent DateClicked(m_CurrentDate)
        Else
            RaiseEvent MonthChanged(m_CurrentDate)
        End If
    End If
End Sub

Private Sub lblHeader_Click(Index As Integer)
    Dim nWeekday As VbDayOfWeek
    ctlDateScroll.Visible = False
    nWeekday = (Index + m_WeekStartsWith) Mod 7
    If nWeekday = 0 Then
        nWeekday = vbSaturday
    End If
    RaiseEvent WeekHeadingClicked(nWeekday)
End Sub

Private Sub lblHeader_DblClick(Index As Integer)
    Dim nWeekday As VbDayOfWeek
    nWeekday = (Index + m_WeekStartsWith) Mod 7
    If nWeekday = 0 Then
        nWeekday = vbSaturday
    End If
    RaiseEvent WeekHeadingDblClicked(nWeekday)
End Sub

Private Sub lblMonth_DblClick()
    RaiseEvent MonthHeadingDblClicked(m_CurrentDate)
End Sub

Private Sub lblMonth_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        m_ShowSelected = True
        shpSelected.Visible = True
        SetShape
        With ctlDateScroll
           .Locked = True
            .BackColor = vbColorBlueProgram        'ColorForegroundHeader
            .ForeColor = vbWhite ' ColorBackgroundHeader
            .Month = Format(m_CurrentDate, "mmmm")
            .Year = Format(m_CurrentDate, "yyyy")
            .Visible = Not .Visible
            DoEvents
            .MonthSetFocus
           .Locked = False
        End With
    End If
    RaiseEvent MonthHeadingClicked(m_CurrentDate)
End Sub

Private Sub lblToday_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ctlDateScroll.Visible = False
        m_CurrentDate = Date
        ShowSelected = True
        Call SetDate
        RaiseEvent TodayClicked(m_CurrentDate)
    End If
End Sub

Private Sub lblTodayShape_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ctlDateScroll.Visible = False
        m_CurrentDate = Date
        ShowSelected = True
        Call SetDate
        RaiseEvent TodayClicked(m_CurrentDate)
    End If
End Sub

Private Sub lblWeeks_Click(Index As Integer)
    ctlDateScroll.Visible = False
End Sub

Private Sub lblWeeks_DblClick(Index As Integer)
    RaiseEvent WeekNumberDblClicked(CInt(lblWeeks(Index).Caption))
End Sub

Private Sub lblWeeks_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nCount As Integer
    
    If Button = 1 Then
        RaiseEvent WeekNumberClicked(CInt(lblWeeks(Index).Caption))
    ElseIf Button = 2 And m_AllowRightClick = True Then
        If mnuAlarm.Count > 1 Then
            For nCount = 1 To mnuAlarm.Count - 1
                Unload mnuAlarm(nCount)
            Next nCount
        End If
        mnuAddAlarm.Visible = False
        mnuSep.Visible = False
        mnuAlarm(0).Visible = False
        PopupMenu mnuAlarmsMain
    End If
End Sub

Private Sub lblYear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        m_ShowSelected = True
        shpSelected.Visible = True
        SetShape
        With ctlDateScroll
           ' .Locked = True
            .BackColor = vbColorBlueProgram        'ColorForegroundHeader
            .ForeColor = vbWhite
            .Month = Format(m_CurrentDate, "mmmm")
            .Year = Format(m_CurrentDate, "yyyy")
            .Visible = Not .Visible
            DoEvents
            .YearSetFocus
            .Locked = False
        End With
    End If
    RaiseEvent MonthHeadingClicked(m_CurrentDate)
End Sub

Private Sub mnuAddAlarm_Click()
    RaiseEvent AddNewAlarm(m_CurrentDate)
End Sub

Private Sub mnuAlarm_Click(Index As Integer)
    RaiseEvent AlarmSelected(CDbl(mnuAlarm(Index).Tag))
End Sub

Private Sub mnuToday_Click()
    If m_Today = True Then
        m_CurrentDate = Date
        ShowSelected = True
        Call SetDate
    End If
    RaiseEvent TodayClicked(m_CurrentDate)
End Sub

Private Sub UserControl_Initialize()
    m_CurrentDate = Date
'    Call SetDate
End Sub

Public Sub Refresh()
    SetDate
    UserControl.Refresh
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal vNewValue As Boolean)
    UserControl.Enabled = vNewValue
    PropertyChanged
End Property

Public Property Get ShowTodayLabel() As Boolean
    ShowTodayLabel = m_Today
End Property

Public Property Let ShowTodayLabel(ByVal vNewValue As Boolean)
    m_Today = vNewValue
    PropertyChanged
    UserControl_Resize
    SetDate
End Property

Public Property Get ColorBackgroundHeader() As OLE_COLOR
    ColorBackgroundHeader = lblBackground.BackColor
End Property
Public Property Let ColorBackgroundHeader(ByVal vNewValue As OLE_COLOR)
    lblBackground.BackColor = vNewValue
    PropertyChanged
End Property

Public Property Get ColorForegroundHeader() As OLE_COLOR
    ColorForegroundHeader = lblMonth.ForeColor
End Property
Public Property Let ColorForegroundHeader(ByVal vNewValue As OLE_COLOR)
    lblMonth.ForeColor = vNewValue
    PropertyChanged
End Property

Public Property Get ColorSelectedBack() As OLE_COLOR
    ColorSelectedBack = shpSelected.FillColor
End Property
Public Property Let ColorSelectedBack(ByVal vNewValue As OLE_COLOR)
    shpSelected.FillColor = vNewValue
    PropertyChanged
End Property

Public Property Get ColorSelectedFore() As OLE_COLOR
    ColorSelectedFore = m_SelectedColor
End Property
Public Property Let ColorSelectedFore(ByVal vNewValue As OLE_COLOR)
    m_SelectedColor = vNewValue
    PropertyChanged
    SetShape
End Property

Public Property Get ColorToday() As OLE_COLOR
    ColorToday = shpToday(0).BorderColor
End Property
Public Property Let ColorToday(ByVal vNewValue As OLE_COLOR)
    shpToday(0).BorderColor = vNewValue
    shpToday(1).BorderColor = vNewValue
    PropertyChanged
End Property

Public Property Get ColorDayColumn() As OLE_COLOR
    ColorDayColumn = lblHeader(0).ForeColor
End Property
Public Property Let ColorDayColumn(ByVal vNewValue As OLE_COLOR)
    Dim nCount As Integer
    For nCount = 0 To lblHeader.UBound
        lblHeader(nCount).ForeColor = vNewValue
    Next nCount
    PropertyChanged
End Property

Public Property Get ColorBackground() As OLE_COLOR
    ColorBackground = UserControl.BackColor
End Property
Public Property Let ColorBackground(ByVal vNewValue As OLE_COLOR)
    UserControl.BackColor = vNewValue
    fraDays.BackColor = vNewValue
    fraWeek.BackColor = vNewValue
    PropertyChanged
End Property

Public Property Get ColorForeground() As OLE_COLOR
    ColorForeground = m_ForegroundColor
End Property
Public Property Let ColorForeground(ByVal vNewValue As OLE_COLOR)
    m_ForegroundColor = vNewValue
    lblToday.ForeColor = vNewValue
    PropertyChanged
    SetDate
End Property

Public Property Get ColorLastNextMonthDayColor() As OLE_COLOR
    ColorLastNextMonthDayColor = m_LastNextDaysColor
End Property
Public Property Let ColorLastNextMonthDayColor(ByVal vNewValue As OLE_COLOR)
    m_LastNextDaysColor = vNewValue
    PropertyChanged
    SetDate
End Property

Public Property Get ColorWeekNumber() As OLE_COLOR
    ColorWeekNumber = m_WeekNumberColor
End Property
Public Property Let ColorWeekNumber(ByVal vNewValue As OLE_COLOR)
    m_WeekNumberColor = vNewValue
    PropertyChanged
    SetDate
End Property

Public Property Get ColorButtons() As OLE_COLOR
    ColorButtons = cmdBack.BackColor
End Property
Public Property Let ColorButtons(ByVal vNewValue As OLE_COLOR)
    cmdBack.BackColor = vNewValue
    cmdFwd.BackColor = vNewValue
    PropertyChanged
End Property

Public Property Get ColorAlarms() As OLE_COLOR
    ColorAlarms = m_AlarmColor
End Property
Public Property Let ColorAlarms(ByVal vNewValue As OLE_COLOR)
    Dim nCount As Integer
    For nCount = 0 To lblDays.UBound
        If lblDays(nCount).Tag = "1" Then
            lblDays(nCount).ForeColor = vNewValue
        End If
    Next nCount
    m_AlarmColor = vNewValue
    PropertyChanged
End Property

Public Property Get ColorLine() As OLE_COLOR
    ColorLine = lineMain.BorderColor
End Property
Public Property Let ColorLine(ByVal vNewValue As OLE_COLOR)
    lineMain.BorderColor = vNewValue
    lineWeekNums.BorderColor = vNewValue
    PropertyChanged
End Property

Public Property Get ShowLastMonthButton() As Boolean
    ShowLastMonthButton = m_ShowLastMonthButton
End Property
Public Property Let ShowLastMonthButton(ByVal vNewValue As Boolean)
    m_ShowLastMonthButton = vNewValue
    cmdBack.Visible = vNewValue
    PropertyChanged
End Property
Public Property Get ShowNextMonthButton() As Boolean
    ShowNextMonthButton = m_ShowNextMonthButton
End Property
Public Property Let ShowNextMonthButton(ByVal vNewValue As Boolean)
    m_ShowNextMonthButton = vNewValue
    cmdFwd.Visible = vNewValue
    PropertyChanged
End Property

Public Property Get ShowSelected() As Boolean
    ShowSelected = m_ShowSelected
End Property
Public Property Let ShowSelected(ByVal vNewValue As Boolean)
    m_ShowSelected = vNewValue
    shpSelected.Visible = vNewValue
    If vNewValue = False And m_LastSelected <> -1 Then
        lblDays(m_LastSelected).ForeColor = m_ForegroundColor
    End If
    PropertyChanged
    SetDate
End Property

Public Property Get UseAlarms() As Boolean
    UseAlarms = m_UseAlarms
End Property
Public Property Let UseAlarms(ByVal vNewValue As Boolean)
    m_UseAlarms = vNewValue
    PropertyChanged
    SetDate
End Property

Public Property Get ShowWeekNumbers() As Boolean
    ShowWeekNumbers = m_ShowWeekNumber
End Property
Public Property Let ShowWeekNumbers(ByVal vNewValue As Boolean)
    Dim nCount As Integer
    m_ShowWeekNumber = vNewValue
    lineWeekNums.Visible = vNewValue
    For nCount = lblWeeks.LBound To lblWeeks.UBound
        lblWeeks(nCount).Visible = vNewValue
    Next nCount
    UserControl_Resize
    SetDate
End Property

Public Property Get ShowWeekNumberLeft() As Boolean
    ShowWeekNumberLeft = m_ShowWeekNumberLeft
End Property
Public Property Let ShowWeekNumberLeft(ByVal vNewValue As Boolean)
    Dim nCount As Integer
    m_ShowWeekNumberLeft = vNewValue
    UserControl_Resize
    SetDate
End Property

Public Property Get weekStartsWith() As VbDayOfWeek
    weekStartsWith = m_WeekStartsWith
End Property
Public Property Let weekStartsWith(ByVal vNewValue As VbDayOfWeek)
    m_WeekStartsWith = vNewValue
    PropertyChanged
    SetDate
End Property

Public Property Get ShowLastMonthDays() As Boolean
    ShowLastMonthDays = m_ShowLastMonthDays
End Property
Public Property Let ShowLastMonthDays(ByVal vNewValue As Boolean)
    m_ShowLastMonthDays = vNewValue
    PropertyChanged
    SetDate
End Property

Public Property Get ShowNextMonthDays() As Boolean
    ShowNextMonthDays = m_ShowNextMonthDays
End Property
Public Property Let ShowNextMonthDays(ByVal vNewValue As Boolean)
    m_ShowNextMonthDays = vNewValue
    PropertyChanged
    SetDate
End Property

Public Property Get ShowShortDays() As Boolean
    ShowShortDays = m_ShowShortDays
End Property
Public Property Let ShowShortDays(ByVal vNewValue As Boolean)
    m_ShowShortDays = vNewValue
    PropertyChanged
    SetDate
End Property

Public Property Get ShowToolTipText() As Boolean
    ShowToolTipText = m_ToolTipText
End Property
Public Property Let ShowToolTipText(ByVal vNewValue As Boolean)
    m_ToolTipText = vNewValue
    PropertyChanged
    SetDate
End Property

Public Property Get AllowRightClick() As Boolean
    AllowRightClick = m_AllowRightClick
End Property
Public Property Let AllowRightClick(ByVal vNewValue As Boolean)
    m_AllowRightClick = vNewValue
    PropertyChanged
End Property

Public Property Get FontHeader() As StdFont
Attribute FontHeader.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
Attribute FontHeader.VB_UserMemId = -512
    Set FontHeader = m_HeaderFont
End Property
Public Property Set FontHeader(ByVal vNewValue As StdFont)
    Set m_HeaderFont = vNewValue
    SetFont
End Property
Private Sub m_HeaderFont_FontChanged(ByVal PropertyName As String)
    SetFont
End Sub

Public Property Get FontDay() As StdFont
    Set FontDay = m_DayFont
End Property
Public Property Set FontDay(ByVal vNewValue As StdFont)
    Set m_DayFont = vNewValue
    SetFont
End Property
Private Sub m_DayFont_FontChanged(ByVal PropertyName As String)
    'm_DayFont.Bold = False
    SetFont
End Sub

Public Property Get FontToday() As StdFont
    Set FontToday = m_TodayFont
End Property
Public Property Set FontToday(ByVal vNewValue As StdFont)
    Set m_TodayFont = vNewValue
    SetFont
End Property
Private Sub m_TodayFont_FontChanged(ByVal PropertyName As String)
    SetFont
End Sub

Public Property Get FontColumn() As StdFont
    Set FontColumn = m_ColumnFont
End Property
Public Property Set FontColumn(ByVal vNewValue As StdFont)
    Set m_ColumnFont = vNewValue
    SetFont
End Property
Private Sub m_ColumnFont_FontChanged(ByVal PropertyName As String)
    SetFont
End Sub


Public Sub About()
Attribute About.VB_UserMemId = -552
End Sub
Private Sub SetFont()
    Dim frmObj
    
    UserControl.AutoRedraw = False
    For Each frmObj In UserControl
        If TypeOf frmObj Is Label Then
            If frmObj.Name = "lblMonth" Then
                ConfigFont m_HeaderFont, frmObj.Font
            ElseIf frmObj.Name = "lblToday" Then
                ConfigFont m_TodayFont, frmObj.Font
            ElseIf frmObj.Name = "lblDays" Or frmObj.Name = "lblWeeks" Then
                ConfigFont m_DayFont, frmObj.Font
            ElseIf frmObj.Name = "lblHeader" Then
                ConfigFont m_ColumnFont, frmObj.Font
            End If
        End If
    Next
    UserControl.AutoRedraw = True
    DoEvents
    
    Call SetDate
End Sub
Private Sub ConfigFont(sourceFont As StdFont, destFont As StdFont)
    destFont.Bold = sourceFont.Bold
    destFont.Charset = sourceFont.Charset
    destFont.Italic = sourceFont.Italic
    destFont.Name = sourceFont.Name
    destFont.Size = sourceFont.Size
    destFont.Strikethrough = sourceFont.Strikethrough
    destFont.Underline = sourceFont.Underline
    destFont.Weight = sourceFont.Weight
End Sub
Public Property Get CurrentDate() As Date
    CurrentDate = Format(m_CurrentDate, "mm/dd/yyyy")
End Property
Public Sub NextMonth()
    m_CurrentDate = DateAdd("m", 1, m_CurrentDate)
    Call SetDate
End Sub
Public Sub LastMonth()
    m_CurrentDate = DateAdd("m", -1, m_CurrentDate)
    Call SetDate
End Sub
Private Sub UserControl_InitProperties()
    ShowLastMonthButton = True
    ShowNextMonthButton = True
    ShowLastMonthDays = True
    ShowNextMonthDays = True
    ShowTodayLabel = True
    ShowToolTipText = True
    ShowWeekNumbers = False
    ShowWeekNumberLeft = True
    AllowRightClick = False
    ShowShortDays = False
    ColorBackgroundHeader = lblBackground.BackColor
    ColorForegroundHeader = lblMonth.ForeColor
    ColorSelectedBack = shpSelected.FillColor
    ColorSelectedFore = vbWhite
    ColorToday = shpToday(0).BorderColor
    ColorDayColumn = lblHeader(0).ForeColor
    ColorAlarms = lblDays(0).ForeColor
    ColorBackground = vbWhite
    ColorForeground = vbBlack
    ColorButtons = cmdBack.BackColor
    ColorLine = &H80000008
    ColorLastNextMonthDayColor = &H808080
    ColorWeekNumber = &H808080
    weekStartsWith = vbSunday
    ShowSelected = True
    UserControl.Width = 8115
    UserControl.Height = 2320 * 3
    SetDate
    PropertyChanged
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ctlDateScroll.Visible = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ShowLastMonthButton = PropBag.ReadProperty("ShowLastMonthButton", ShowLastMonthButton)
    ShowNextMonthButton = PropBag.ReadProperty("ShowNextMonthButton", ShowNextMonthButton)
    ShowLastMonthDays = PropBag.ReadProperty("ShowLastMonthDays", ShowLastMonthDays)
    ShowNextMonthDays = PropBag.ReadProperty("ShowNextMonthDays", ShowNextMonthDays)
    ShowTodayLabel = PropBag.ReadProperty("ShowTodayLabel", ShowTodayLabel)
    ColorBackgroundHeader = PropBag.ReadProperty("ColorBackgroundHeader", ColorBackgroundHeader)
    ColorForegroundHeader = PropBag.ReadProperty("ColorForegroundHeader", ColorForegroundHeader)
    ColorSelectedBack = PropBag.ReadProperty("ColorSelectedBack", ColorSelectedBack)
    ColorSelectedFore = PropBag.ReadProperty("ColorSelectedFore", ColorSelectedFore)
    ColorToday = PropBag.ReadProperty("ColorToday", ColorToday)
    ColorDayColumn = PropBag.ReadProperty("ColorDayColumn", ColorDayColumn)
    ColorAlarms = PropBag.ReadProperty("ColorAlarms", ColorAlarms)
    ColorBackground = PropBag.ReadProperty("ColorBackground", ColorBackground)
    ColorForeground = PropBag.ReadProperty("ColorForeground", ColorForeground)
    ColorButtons = PropBag.ReadProperty("ColorButtons", ColorButtons)
    ColorLine = PropBag.ReadProperty("ColorLine", ColorLine)
    ColorWeekNumber = PropBag.ReadProperty("ColorWeekNumber", ColorWeekNumber)
    ColorLastNextMonthDayColor = PropBag.ReadProperty("ColorLastNextMonthDayColor", ColorLastNextMonthDayColor)
    weekStartsWith = PropBag.ReadProperty("WeekStartsWith", weekStartsWith)
    ShowSelected = PropBag.ReadProperty("ShowSelected", ShowSelected)
    ShowToolTipText = PropBag.ReadProperty("ShowToolTipText", ShowToolTipText)
    ShowWeekNumbers = PropBag.ReadProperty("ShowWeekNumbers", ShowWeekNumbers)
    ShowWeekNumberLeft = PropBag.ReadProperty("ShowWeekNumberLeft", ShowWeekNumberLeft)
    AllowRightClick = PropBag.ReadProperty("AllowRightClick", AllowRightClick)
    UseAlarms = PropBag.ReadProperty("UseAlarms", UseAlarms)
    ShowShortDays = PropBag.ReadProperty("ShowShortDays", ShowShortDays)
    Set m_HeaderFont = PropBag.ReadProperty("FontHeader", lblMonth.Font)
    Set m_DayFont = PropBag.ReadProperty("FontDay", lblDays(1).Font)
    Set m_TodayFont = PropBag.ReadProperty("FontToday", lblToday.Font)
    Set m_ColumnFont = PropBag.ReadProperty("FontColumn", lblHeader(1).Font)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    SetDate
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ShowLastMonthButton", ShowLastMonthButton)
    Call PropBag.WriteProperty("ShowNextMonthButton", ShowNextMonthButton)
    Call PropBag.WriteProperty("ShowLastMonthDays", ShowLastMonthDays)
    Call PropBag.WriteProperty("ShowNextMonthDays", ShowNextMonthDays)
    Call PropBag.WriteProperty("ShowTodayLabel", ShowTodayLabel)
    Call PropBag.WriteProperty("ColorBackgroundHeader", ColorBackgroundHeader)
    Call PropBag.WriteProperty("ColorForegroundHeader", ColorForegroundHeader)
    Call PropBag.WriteProperty("ColorSelectedBack", ColorSelectedBack)
    Call PropBag.WriteProperty("ColorSelectedFore", ColorSelectedFore)
    Call PropBag.WriteProperty("ColorToday", ColorToday)
    Call PropBag.WriteProperty("ColorDayColumn", ColorDayColumn)
    Call PropBag.WriteProperty("ColorAlarms", ColorAlarms)
    Call PropBag.WriteProperty("ColorBackground", ColorBackground)
    Call PropBag.WriteProperty("ColorForeground", ColorForeground)
    Call PropBag.WriteProperty("ColorButtons", ColorButtons)
    Call PropBag.WriteProperty("ColorLastNextMonthDayColor", ColorLastNextMonthDayColor)
    Call PropBag.WriteProperty("ColorLine", ColorLine)
    Call PropBag.WriteProperty("ColorWeekNumber", ColorWeekNumber)
    Call PropBag.WriteProperty("WeekStartsWith", weekStartsWith)
    Call PropBag.WriteProperty("ShowSelected", ShowSelected)
    Call PropBag.WriteProperty("ShowToolTipText", ShowToolTipText)
    Call PropBag.WriteProperty("ShowWeekNumbers", ShowWeekNumbers)
    Call PropBag.WriteProperty("ShowWeekNumberLeft", ShowWeekNumberLeft)
    Call PropBag.WriteProperty("AllowRightClick", AllowRightClick)
    Call PropBag.WriteProperty("UseAlarms", UseAlarms)
    Call PropBag.WriteProperty("ShowShortDays", ShowShortDays)
    Call PropBag.WriteProperty("FontHeader", m_HeaderFont)
    Call PropBag.WriteProperty("FontDay", m_DayFont)
    Call PropBag.WriteProperty("FontToday", m_TodayFont)
    Call PropBag.WriteProperty("FontColumn", m_ColumnFont)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    SetDate
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 8115
        fraDays.Left = 40
    lblBackground.Width = UserControl.Width + 200
    cmdFwd.Left = UserControl.Width - cmdFwd.Width - 40
    ctlDateScroll.Left = UserControl.Width / 2 - ctlDateScroll.Width / 2
    lblMonth.Width = UserControl.Width
    lblMonth.Left = 0

    UserControl.Height = 2320 * 3
 
End Sub

