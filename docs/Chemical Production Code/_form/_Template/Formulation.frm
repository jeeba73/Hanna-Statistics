VERSION 5.00
Begin VB.Form Formulation 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12000
   ScaleWidth      =   19200
   Begin VB.PictureBox PicHover 
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   675
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   675
      Begin VB.Label imOver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "é"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   175
         TabIndex        =   64
         Top             =   80
         Width           =   330
      End
      Begin VB.Label lblHoverClick 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00808080&
         Height          =   570
         Left            =   60
         TabIndex        =   63
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.PictureBox PBContainerViewport 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9975
      Index           =   0
      Left            =   -120
      ScaleHeight     =   9975
      ScaleWidth      =   19245
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1080
      Width           =   19245
      Begin VB.PictureBox PBContainer 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   34215
         Left            =   360
         ScaleHeight     =   34215
         ScaleWidth      =   18555
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   360
         Width           =   18555
         Begin VB.Frame frCommandInside 
            BackColor       =   &H00644603&
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
            Height          =   495
            Index           =   1
            Left            =   9600
            TabIndex        =   68
            Top             =   1440
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Select Recipe"
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
               TabIndex        =   69
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.Frame frCommandInside 
            BackColor       =   &H00644603&
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
            Height          =   495
            Index           =   0
            Left            =   5880
            TabIndex        =   66
            Top             =   1440
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Select Hanna Code"
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
               TabIndex        =   67
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.PictureBox PBBottom 
            BackColor       =   &H00322D2B&
            BorderStyle     =   0  'None
            Height          =   4485
            Left            =   -315
            ScaleHeight     =   4485
            ScaleWidth      =   65250
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   27990
            Width           =   65255
            Begin VB.Line Line1 
               BorderColor     =   &H00505050&
               Index           =   6
               X1              =   8880
               X2              =   8880
               Y1              =   3760
               Y2              =   540
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Donate"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   345
               Index           =   19
               Left            =   9195
               TabIndex        =   60
               Top             =   2250
               Width           =   750
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Send news"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   345
               Index           =   18
               Left            =   9195
               TabIndex        =   59
               Top             =   1755
               Width           =   1170
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Report an error"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   345
               Index           =   17
               Left            =   9195
               TabIndex        =   58
               Top             =   1260
               Width           =   1575
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Feedback"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00E0E0E0&
               Height          =   345
               Index           =   16
               Left            =   9195
               TabIndex        =   57
               Top             =   645
               Width           =   1125
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00B57CC3&
               BorderWidth     =   2
               Index           =   2
               X1              =   9210
               X2              =   9700
               Y1              =   450
               Y2              =   450
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Politics"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   345
               Index           =   15
               Left            =   5160
               TabIndex        =   56
               Top             =   3270
               Width           =   705
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Video"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   345
               Index           =   14
               Left            =   5160
               TabIndex        =   55
               Top             =   2760
               Width           =   615
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Entertainment"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   345
               Index           =   13
               Left            =   5160
               TabIndex        =   54
               Top             =   2250
               Width           =   1410
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gaming"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   345
               Index           =   12
               Left            =   5160
               TabIndex        =   53
               Top             =   1755
               Width           =   825
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reviews"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   345
               Index           =   11
               Left            =   5160
               TabIndex        =   52
               Top             =   1260
               Width           =   870
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sections"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00E0E0E0&
               Height          =   345
               Index           =   10
               Left            =   5160
               TabIndex        =   51
               Top             =   645
               Width           =   990
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00B57CC3&
               BorderWidth     =   2
               Index           =   1
               X1              =   5175
               X2              =   5665
               Y1              =   450
               Y2              =   450
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rss Feed"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   345
               Index           =   9
               Left            =   1245
               TabIndex        =   50
               Top             =   2250
               Width           =   1020
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "About our ads"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   345
               Index           =   8
               Left            =   1245
               TabIndex        =   49
               Top             =   1755
               Width           =   1455
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "About this page"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   345
               Index           =   7
               Left            =   1245
               TabIndex        =   48
               Top             =   1260
               Width           =   1605
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "About"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00E0E0E0&
               Height          =   345
               Index           =   6
               Left            =   1245
               TabIndex        =   47
               Top             =   645
               Width           =   690
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00505050&
               Index           =   4
               X1              =   4800
               X2              =   4800
               Y1              =   3700
               Y2              =   480
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00B57CC3&
               BorderWidth     =   2
               Index           =   0
               X1              =   1260
               X2              =   1750
               Y1              =   450
               Y2              =   450
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4620
            Index           =   5
            Left            =   555
            ScaleHeight     =   4620
            ScaleWidth      =   15165
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   18585
            Width           =   15165
            Begin VB.Line Line1 
               BorderColor     =   &H00B7B7B7&
               Index           =   5
               X1              =   300
               X2              =   14250
               Y1              =   4455
               Y2              =   4455
            End
            Begin VB.Label Label3 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Why not go for a multi-port charger instead of a single-port one?"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00676767&
               Height          =   795
               Index           =   5
               Left            =   6795
               TabIndex        =   45
               Top             =   1785
               Width           =   7860
            End
            Begin VB.Label lblTitleNews 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "The best USB phone charger"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Index           =   5
               Left            =   6795
               TabIndex        =   44
               Top             =   975
               Width           =   7830
            End
            Begin VB.Image Image1 
               Height          =   3000
               Index           =   5
               Left            =   1545
               Picture         =   "Formulation.frx":0000
               Stretch         =   -1  'True
               Top             =   420
               Width           =   4500
            End
            Begin VB.Label lblCategory 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gadgetry"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B57CC3&
               Height          =   315
               Index           =   5
               Left            =   6795
               TabIndex        =   43
               Top             =   435
               Width           =   1020
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "5Hs"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   315
               Index           =   5
               Left            =   555
               TabIndex        =   42
               Top             =   435
               Width           =   420
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "À"
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   12
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   255
               Index           =   10
               Left            =   255
               TabIndex        =   41
               Top             =   495
               Width           =   210
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4620
            Index           =   4
            Left            =   555
            ScaleHeight     =   4620
            ScaleWidth      =   15165
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   23220
            Width           =   15165
            Begin VB.Label Label3 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "You may have to settle, if just a little bit."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00676767&
               Height          =   795
               Index           =   4
               Left            =   6795
               TabIndex        =   39
               Top             =   2235
               Width           =   7860
            End
            Begin VB.Label lblTitleNews 
               BackStyle       =   0  'Transparent
               Caption         =   "OnePlus may return to low-cost phones with the 8 Lite"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Index           =   4
               Left            =   6795
               TabIndex        =   38
               Top             =   975
               Width           =   7845
            End
            Begin VB.Image Image1 
               Height          =   3000
               Index           =   4
               Left            =   1545
               Picture         =   "Formulation.frx":2216
               Stretch         =   -1  'True
               Top             =   420
               Width           =   4500
            End
            Begin VB.Label lblCategory 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B57CC3&
               Height          =   315
               Index           =   4
               Left            =   6795
               TabIndex        =   37
               Top             =   435
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "7Hs"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   315
               Index           =   4
               Left            =   555
               TabIndex        =   36
               Top             =   435
               Width           =   420
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "À"
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   12
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   255
               Index           =   9
               Left            =   255
               TabIndex        =   35
               Top             =   495
               Width           =   210
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4620
            Index           =   3
            Left            =   555
            ScaleHeight     =   4620
            ScaleWidth      =   15165
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   13950
            Width           =   15165
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "À"
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   12
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   255
               Index           =   8
               Left            =   255
               TabIndex        =   33
               Top             =   495
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "4Hs"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   315
               Index           =   3
               Left            =   555
               TabIndex        =   32
               Top             =   435
               Width           =   420
            End
            Begin VB.Label lblCategory 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Space"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B57CC3&
               Height          =   315
               Index           =   3
               Left            =   6795
               TabIndex        =   31
               Top             =   435
               Width           =   660
            End
            Begin VB.Image Image1 
               Height          =   3000
               Index           =   3
               Left            =   1545
               Picture         =   "Formulation.frx":3F92
               Stretch         =   -1  'True
               Top             =   420
               Width           =   4500
            End
            Begin VB.Label lblTitleNews 
               BackStyle       =   0  'Transparent
               Caption         =   "NASA hopes OSIRIS-REx data will explain an asteroid's mini-eruptions"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Index           =   3
               Left            =   6795
               TabIndex        =   30
               Top             =   975
               Width           =   7815
            End
            Begin VB.Label Label3 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Samples sent back to Earth might help solve the mystery."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00676767&
               Height          =   795
               Index           =   3
               Left            =   6795
               TabIndex        =   29
               Top             =   2235
               Width           =   7860
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00B7B7B7&
               Index           =   3
               X1              =   300
               X2              =   14250
               Y1              =   4455
               Y2              =   4455
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4620
            Index           =   2
            Left            =   555
            ScaleHeight     =   4620
            ScaleWidth      =   15165
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   9315
            Width           =   15165
            Begin VB.Line Line1 
               BorderColor     =   &H00B7B7B7&
               Index           =   2
               X1              =   300
               X2              =   14250
               Y1              =   4455
               Y2              =   4455
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Elon Musk is also considering an electric dirt bike."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00676767&
               Height          =   795
               Index           =   2
               Left            =   6795
               TabIndex        =   27
               Top             =   2235
               Width           =   7875
            End
            Begin VB.Label lblTitleNews 
               BackStyle       =   0  'Transparent
               Caption         =   "Tesla's electric ATV should launch at the same time as the Cybertruck"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   975
               Index           =   2
               Left            =   6795
               TabIndex        =   26
               Top             =   975
               Width           =   7815
            End
            Begin VB.Image Image1 
               Height          =   3000
               Index           =   2
               Left            =   1545
               Picture         =   "Formulation.frx":4983
               Stretch         =   -1  'True
               Top             =   420
               Width           =   4500
            End
            Begin VB.Label lblCategory 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Transportation"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B57CC3&
               Height          =   315
               Index           =   2
               Left            =   6795
               TabIndex        =   25
               Top             =   435
               Width           =   1620
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "3Hs"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   315
               Index           =   2
               Left            =   555
               TabIndex        =   24
               Top             =   435
               Width           =   420
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "À"
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   12
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00B7B7B7&
               Height          =   255
               Index           =   7
               Left            =   255
               TabIndex        =   23
               Top             =   495
               Width           =   210
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(C)2019 No copyright on this example."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00B7B7B7&
            Height          =   315
            Index           =   6
            Left            =   945
            TabIndex        =   61
            Top             =   32745
            Width           =   4200
         End
      End
   End
   Begin VB.PictureBox PBFooter 
      BackColor       =   &H00886010&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   12
      Top             =   11040
      Width           =   19215
      Begin VB.Timer Timer2 
         Interval        =   10
         Left            =   8400
         Top             =   120
      End
      Begin VB.Label Lab 
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
         Left            =   17745
         MouseIcon       =   "Formulation.frx":6891
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   630
         Width           =   1200
      End
      Begin VB.Label Lab 
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
         Left            =   15345
         MouseIcon       =   "Formulation.frx":6B9B
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Formulation"
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
         Index           =   7
         Left            =   8880
         MouseIcon       =   "Formulation.frx":6EA5
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   630
         Width           =   1380
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   0
         Left            =   9360
         MouseIcon       =   "Formulation.frx":71AF
         MousePointer    =   99  'Custom
         Picture         =   "Formulation.frx":74B9
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   3
         Left            =   15600
         MouseIcon       =   "Formulation.frx":A89B
         MousePointer    =   99  'Custom
         Picture         =   "Formulation.frx":ABA5
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   4
         Left            =   18240
         MouseIcon       =   "Formulation.frx":DF87
         MousePointer    =   99  'Custom
         Picture         =   "Formulation.frx":E291
         Top             =   120
         Width           =   480
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   0
         Left            =   8760
         TabIndex        =   16
         Top             =   -120
         Width           =   1695
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   2
         Left            =   0
         TabIndex        =   15
         Top             =   -120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   3
         Left            =   14760
         TabIndex        =   14
         Top             =   -120
         Width           =   2175
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   4
         Left            =   17280
         TabIndex        =   13
         Top             =   -120
         Width           =   1935
      End
   End
   Begin VB.PictureBox PBTitle 
      BackColor       =   &H00644603&
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
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   19215
      TabIndex        =   0
      Top             =   0
      Width           =   19215
      Begin ChemicalProduction.ucScrollAdd ucScrollAdd2 
         Left            =   10800
         Top             =   120
         _ExtentX        =   1138
         _ExtentY        =   423
      End
      Begin ChemicalProduction.ucScrollAdd ucScrollAdd1 
         Left            =   9840
         Top             =   120
         _ExtentX        =   1138
         _ExtentY        =   423
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   4
         Left            =   7680
         MouseIcon       =   "Formulation.frx":11673
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            MouseIcon       =   "Formulation.frx":1197D
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   4
            Left            =   720
            MouseIcon       =   "Formulation.frx":11C87
            MousePointer    =   99  'Custom
            Picture         =   "Formulation.frx":11F91
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00A48643&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MouseIcon       =   "Formulation.frx":15373
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   7
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "Formulation.frx":1567D
            MousePointer    =   99  'Custom
            Picture         =   "Formulation.frx":15987
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Formulation"
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
            Left            =   90
            MouseIcon       =   "Formulation.frx":18D69
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   720
            Width           =   1830
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   1920
         MouseIcon       =   "Formulation.frx":19073
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   5
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   720
            MouseIcon       =   "Formulation.frx":1937D
            MousePointer    =   99  'Custom
            Picture         =   "Formulation.frx":19687
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Material Requisition"
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
            MouseIcon       =   "Formulation.frx":1CA69
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   720
            Width           =   1830
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   3840
         MouseIcon       =   "Formulation.frx":1CD73
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   720
            MouseIcon       =   "Formulation.frx":1D07D
            MousePointer    =   99  'Custom
            Picture         =   "Formulation.frx":1D387
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            MouseIcon       =   "Formulation.frx":20769
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   720
            Width           =   1890
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   3
         Left            =   5760
         MouseIcon       =   "Formulation.frx":20A73
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   720
            MouseIcon       =   "Formulation.frx":20D7D
            MousePointer    =   99  'Custom
            Picture         =   "Formulation.frx":21087
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            MouseIcon       =   "Formulation.frx":24469
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   720
            Width           =   1890
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formulation"
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
         Height          =   495
         Left            =   16605
         TabIndex        =   9
         Top             =   240
         Width           =   2325
      End
   End
   Begin VB.PictureBox PBContainerViewport 
      BorderStyle     =   0  'None
      Height          =   6975
      Index           =   1
      Left            =   12120
      ScaleHeight     =   6975
      ScaleWidth      =   7215
      TabIndex        =   65
      Top             =   1680
      Width           =   7215
   End
End
Attribute VB_Name = "Formulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_rc As Boolean



Private Type ControlPositionType
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    FontSize As Single
End Type


Private m_ControlPositions() As ControlPositionType
Private m_FormWid As Single
Private m_FormHgt As Single

Private IndexProcedura As Integer
Private IndexDashCommInside As Integer



Private Sub Form_Load()
    'Add scrolling capabilities to the container
    ucScrollAdd1.AddScroll PBContainerViewport(0)
    'Note that PBContainerViewport is our target now, wich define the width and height of view of the scrolling area.
    'inside PBContainerViewport we put our previous PBContainer wich has the real size of the full area to scroll
    'and the elements inside (To understand it better play changing the backcolor of the containers)
    
    'Make it scroll with MouseWheel
    ucScrollAdd1.TrackMouseWheel Vertical
    'Limit the resize of the window
   ' ucScrollAdd1.ResizeWindowLimit 450, 148 ', 1097, 2000
    
    'Notice: We dont need to call 'ucScrollAdd1.RemoveFromContainer PBTitle' because its no more inside the scrolling area
    '(its not inside the TARGET container, before it was the 'Form', now its only 'PBContainerViewport')
    
    'Automatize the process of resizing the target used as a viewport when the form resize
    ucScrollAdd1.ResizeTargetOnFormResize 0, 0
    
    ucScrollAdd1.RemoveFromContainer PBFooter
    
    'Remove the buttons on the scrollbars (just to give a different look)
    ucScrollAdd1.UCScrollV.ShowButtons = False
    ucScrollAdd1.UCScrollH.ShowButtons = False
    
    
    Dim i As Integer
    If Screen.Width - Me.Width > 1000 And bFullScreen Then
        Me.WindowState = 2
    
    End If


    For i = PBContainerViewport.LBound To PBContainerViewport.UBound
        PBContainerViewport(i).Move 0, PBTitle.Height, Me.ScaleWidth, Me.ScaleHeight - PBTitle.Height
    Next
  
    RSBottom PicHover, Me, -1350
    RSRight PicHover, Me, -450
    
    PBContainer.Top = 0
    PBContainer.Left = 0
    PBContainer.Width = Me.Width
    

    PBContainerViewport(0).ZOrder
    PBFooter.ZOrder
    
End Sub

Private Sub Form_Resize()

    PBTitle.Width = Me.Width
    PBFooter.Top = Me.ScaleHeight - PBFooter.Height
    PBFooter.Width = Me.Width
 
    
    'Resize the container (needed to show the full bottom box on maximized state)
    'First resize our container
    ucScrollAdd1.ContainerW = Me.ScaleWidth
    'But also need to resize PBContainer wich hide the width of the bottom box

    
      ResizeControls

    
End Sub


Private Sub ucScrollAdd1_ScrollH(Value As Long)
    Form_Resize
End Sub
Private Sub PicHover_Click()
ucScrollAdd1.UCScrollV.ScrollToValue 0
End Sub
Private Sub lblHoverClick_Click()
    ucScrollAdd1.UCScrollV.ScrollToValue 0
End Sub
Private Sub imOver_Click()
ucScrollAdd1.UCScrollV.ScrollToValue 0
End Sub

'========================================
'Vertical scroll event
'========================================
Private Sub ucScrollAdd1_ScrollV(Value As Long)
    
    'Just log the value for no reason
   
        
    PicHover.ZOrder
    'Show a button to scroll to top
    If Not (ucScrollAdd1.UCScrollV Is Nothing) Then
        If (ucScrollAdd1.UCScrollV.Value > 100) Then
            PicHover.Visible = True
        Else
            PicHover.Visible = False
        End If
    Else
        PicHover.Visible = False
    End If
    
   
    
End Sub

'Poorly made resizing functions just for the example
Private Sub RSRight(c As Control, Source As Object, adjust As Long, Optional LimitLeft& = -1, Optional LimitRight& = -1)
On Error Resume Next
Dim aux&
    aux& = (Source.ScaleWidth - c.Width) + adjust
    If (Err.NUMBER > 0) Then aux& = (Source.Width - c.Width) + adjust
    If (aux < LimitLeft) And (LimitLeft <> -1) Then aux = LimitLeft
    If (aux > LimitRight&) And (LimitRight& <> -1) Then aux = LimitRight&
    c.Left = aux
End Sub

Private Sub RSWidth(c As Control, Source As Object, adjust As Long, Optional LimitLeft& = 0, Optional LimitRight& = -1)
Dim aux&
    aux& = Source.Width + adjust
    If (aux < LimitLeft) Then aux = LimitLeft
    If (aux > LimitRight&) And (LimitRight& <> -1) Then aux = LimitRight&
    c.Width = aux
End Sub

Private Sub RSCenter(c As Control, Source As Object, Optional adjust As Long = 0, Optional LimitLeft& = -1, Optional LimitRight& = -1)
Dim aux&
    aux& = ((Source.Width / 2) - (c.Width / 2)) + adjust
    If (aux < LimitLeft) And (LimitLeft <> -1) Then aux = LimitLeft
    If (aux > LimitRight&) And (LimitRight& <> -1) Then aux = LimitRight&
    c.Left = aux
End Sub

Private Sub RSBottom(c As Control, Source As Object, adjust As Long, Optional LimitBot& = -1)
On Error Resume Next
Dim aux&
    aux& = (Source.ScaleHeight - c.Height) + adjust
    If (Err.NUMBER > 0) Then aux& = (Source.Height - c.Height) + adjust
    If (aux < LimitBot) And (LimitBot <> -1) Then aux = LimitBot
    c.Top = aux
End Sub

Private Sub RSLeft(c As Control, Source As Object, adjust As Long, Optional LimitLeft& = -1, Optional LimitRight& = -1)
Dim aux&
    aux& = Source.Left + adjust
    If (aux < LimitLeft) And (LimitLeft <> -1) Then
        aux = LimitLeft
    ElseIf (aux > LimitRight&) And (LimitRight& <> -1) Then
        aux = LimitRight&
    End If
    c.Left = aux
End Sub



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
        ElseIf TypeOf ctl Is ucScrollAdd Then

        Else
            .Left = ctl.Left
           ' MsgBox (TypeName(ctl))
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

'If Not (bStazioneEsterna) Then
'm_ControlGridFontSize = 1 ' y_scale * 0.8
'm_ControlGridColWidth = 1 ' x_scale

'End If
'm_ControlGridRowHeight = 1 '1.4


For Each ctl In Controls
    With m_ControlPositions(i)
        If TypeOf ctl Is Line Then
            ctl.x1 = x_scale * .Left
            ctl.y1 = y_scale * .Top
            ctl.x2 = ctl.x1 + x_scale * .Width
            ctl.y2 = ctl.y1 + y_scale * .Height
        ElseIf TypeOf ctl Is Timer Then
        ElseIf TypeOf ctl Is Inet Then
        ElseIf TypeOf ctl Is ucScrollAdd Then
        ElseIf TypeOf ctl Is Grid Then
           ctl.Left = x_scale * .Left
            ctl.Top = y_scale * .Top
            ctl.Width = x_scale * .Width
            ctl.Height = y_scale * .Height

                ctl.DefaultFont.Size = 12 * m_ControlGridFontSize
                ctl.DefaultRowHeight = 30 * m_ControlGridRowHeight
           
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

Private Sub Form_Initialize()

SaveSizes
End Sub

Public Function DoShow(Optional ByVal ID As Long) As Boolean

    On Error GoTo ERR_SHOW
    
    m_rc = False
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
    Resume ERR_END
End Function




Private Sub DefaultMenu_Click(Index As Integer)
Select Case Index
    Case 0
        Unload Me
End Select
End Sub



Private Sub DefaultMenuLabel_Click(Index As Integer)
DefaultMenu_Click Index
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27
        Unload Me
    Case 37
        DefaultMenuLabel_Click 2
    Case 39
        DefaultMenuLabel_Click 0
End Select
End Sub

Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To PicMenu.Count - 1
 
    If i = IndexProcedura Then
        ' lascialo stare...
    ElseIf i = Index Then
        PicMenu(i).BackColor = &H886010
    Else
        PicMenu(i).BackColor = &H644603
    End If
Next
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

If Index > PicMenu.UBound Then Exit Function


For i = 0 To PicMenu.UBound
    If i = Index Then
        PicMenu(i).BackColor = &HA48643
    Else
        PicMenu(i).BackColor = &H644603
    End If
Next
blTable = Label2(Index)
IndexProcedura = Index

PBContainerViewport(Index).ZOrder
PBContainerViewport(Index).Visible = True

Select Case IndexProcedura
    Case 0

    Case 1
    
End Select

PBFooter.ZOrder


End Function



Private Sub Label2_Click(Index As Integer)
PicMenu_Click Index
End Sub


Private Sub PicMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
PicMenu_Click Index
End Sub
Private Sub Image3_Click(Index As Integer)
PicMenu_Click Index
End Sub

Private Sub frCommandInside_Click(Index As Integer)
    Select Case Index
        Case 0

        Case 1
        
        Case 2

    End Select
End Sub

Private Sub frCommandInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
IndexDashCommInside = Index
Dim i As Integer
    For i = 0 To frCommandInside.UBound
        If i = Index Then
            frCommandInside(i).BackColor = &H846623
            lbCommandInside(i).ForeColor = vbWhite
            If i = 5 Then
                frCommandInside(i).BackColor = &H20A020
            End If
            
        Else
            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 5 Then
                frCommandInside(i).BackColor = &H8000&
            End If
        End If
    
    Next
 
 
End Sub
Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub
Private Sub lbCommandInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
frCommandInside_MouseMove Index, Button, Shift, X, Y
End Sub
Private Sub PBTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub PBTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub PBTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
FrmMove = False
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FrmMove = True
    DragX = X
    DragY = Y
    If Me.WindowState = 2 Then
        FrmMove = False
       
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nx, ny
    If Me.WindowState = 2 Then
        FrmMove = False
        Exit Sub
    End If
    nx = Me.Left + X - DragX
    ny = Me.Top + Y - DragY
    Me.Left = nx
    Me.Top = ny
    FrmMove = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer



If Me.WindowState = 2 Then
    FrmMove = False
End If
Dim nx, ny
    If FrmMove Then
        nx = Me.Left + X - DragX
        ny = Me.Top + Y - DragY
        Me.Left = nx
        Me.Top = ny
    End If
    
For i = 0 To 3
    If i = IndexProcedura Then
    Else
        PicMenu(i).BackColor = &H644603
    End If
Next
End Sub
