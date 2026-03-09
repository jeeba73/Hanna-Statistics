VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form FormDatabase 
   BackColor       =   &H00F0F0F0&
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
   Icon            =   "FormCode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   12000
   ScaleWidth      =   19200
   Begin VB.Frame frInside 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9375
      Index           =   1
      Left            =   6840
      TabIndex        =   14
      Top             =   1440
      Width           =   12255
      Begin VB.PictureBox PicType 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         Height          =   2295
         Left            =   840
         ScaleHeight     =   2295
         ScaleWidth      =   6975
         TabIndex        =   39
         Top             =   4560
         Visible         =   0   'False
         Width           =   6975
         Begin VB.ComboBox cmbType 
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Preparation Type"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2235
            TabIndex        =   41
            Top             =   600
            Width           =   2250
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   6600
            Picture         =   "FormCode.frx":0A02
            Top             =   120
            Width           =   240
         End
      End
      Begin VB.PictureBox PicUm 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         Height          =   2295
         Left            =   6000
         ScaleHeight     =   2295
         ScaleWidth      =   6975
         TabIndex        =   35
         Top             =   2160
         Visible         =   0   'False
         Width           =   6975
         Begin VB.ComboBox cmbUM 
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   6600
            Picture         =   "FormCode.frx":1404
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lbUM 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "label"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3225
            TabIndex        =   38
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select from list"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   2820
            TabIndex        =   37
            Top             =   720
            Width           =   1380
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00F0F0F0&
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
         Height          =   735
         Index           =   5
         Left            =   0
         TabIndex        =   15
         Top             =   120
         Width           =   12255
         Begin VB.Frame frCommandInside 
            BackColor       =   &H00886010&
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
            Index           =   8
            Left            =   5520
            TabIndex        =   68
            Top             =   120
            Visible         =   0   'False
            Width           =   3255
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Open Revision history"
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
               Index           =   8
               Left            =   0
               TabIndex        =   69
               Top             =   120
               Width           =   3255
            End
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
            Index           =   5
            Left            =   8880
            TabIndex        =   33
            Top             =   120
            Visible         =   0   'False
            Width           =   3255
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Open Formulation"
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
               Index           =   5
               Left            =   0
               TabIndex        =   34
               Top             =   120
               Width           =   3255
            End
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            X1              =   120
            X2              =   12000
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label lbInside 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Specifications"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00644603&
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1590
         End
      End
      Begin FlexCell.Grid Grd2 
         Height          =   8295
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   14631
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColorActiveCellSel=   15790320
         BackColorBkg    =   15790320
         BackColorFixed  =   15790320
         BackColorFixedSel=   15790320
         BackColorScrollBar=   15592423
         BorderColor     =   15790320
         CellBorderColor =   15790320
         CellBorderColorFixed=   15790320
         Cols            =   5
         DefaultFontName =   "Century Gothic"
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DefaultRowHeight=   20
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   15790320
         Rows            =   1
         ScrollBarStyle  =   0
         MultiSelect     =   0   'False
      End
      Begin VB.Frame frCommandInside 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Caption         =   "Image14"
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
         Left            =   8880
         TabIndex        =   45
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
         Begin VB.Label lbCommandInside 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Product Classification"
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
            TabIndex        =   46
            Top             =   120
            Width           =   3255
         End
      End
   End
   Begin VB.Frame frInside 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9495
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   6495
      Begin VB.ComboBox cmbLineHannaCode 
         BackColor       =   &H00886010&
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
         Height          =   435
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   80
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   420
         Index           =   0
         Left            =   3240
         TabIndex        =   20
         Top             =   1200
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00F0F0F0&
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
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   6255
         Begin VB.Line Line5 
            BorderColor     =   &H00E0E0E0&
            X1              =   240
            X2              =   6120
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Label lbInside 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Record List / Search"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00644603&
            Height          =   285
            Index           =   2
            Left            =   240
            TabIndex        =   28
            Top             =   0
            Width           =   2235
         End
      End
      Begin FlexCell.Grid Grd1 
         Height          =   7695
         Left            =   240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1800
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   13573
         AllowUserSort   =   -1  'True
         Appearance      =   0
         BackColor1      =   15790320
         BackColor2      =   15790320
         BackColorActiveCellSel=   15790320
         BackColorBkg    =   15790320
         BackColorFixed  =   15790320
         BackColorFixedSel=   15790320
         BackColorScrollBar=   15592423
         BorderColor     =   15790320
         CellBorderColor =   15790320
         CellBorderColorFixed=   15790320
         Cols            =   5
         DefaultFontName =   "Century Gothic"
         DefaultFontSize =   8.25
         BoldFixedCell   =   0   'False
         DisplayRowIndex =   -1  'True
         DrawMode        =   1
         DefaultRowHeight=   20
         FixedRowColStyle=   0
         ForeColorFixed  =   4210752
         GridColor       =   15790320
         ReadOnly        =   -1  'True
         Rows            =   1
         ScrollBarStyle  =   0
         SelectionMode   =   1
         MultiSelect     =   0   'False
      End
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
      Index           =   7
      Left            =   11280
      TabIndex        =   50
      Top             =   6120
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label lbCommandInside 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "View Classification"
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
         Index           =   7
         Left            =   0
         TabIndex        =   51
         Top             =   120
         Width           =   6615
      End
   End
   Begin VB.PictureBox PicHover 
      BackColor       =   &H00886010&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   675
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   675
      Begin VB.Label lblHoverClick 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00808080&
         Height          =   570
         Left            =   60
         TabIndex        =   44
         Top             =   0
         Width           =   585
      End
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
         TabIndex        =   43
         Top             =   80
         Width           =   330
      End
   End
   Begin VB.PictureBox PBFooter 
      BackColor       =   &H00886010&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   22
      Top             =   11040
      Width           =   19215
      Begin VB.Timer Timer2 
         Interval        =   10
         Left            =   8400
         Top             =   120
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   2
         Left            =   0
         TabIndex        =   25
         Top             =   -120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   4
         Left            =   17280
         TabIndex        =   23
         Top             =   -120
         Width           =   1935
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   0
         Left            =   8760
         TabIndex        =   27
         Top             =   -120
         Width           =   1695
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   3
         Left            =   14760
         TabIndex        =   24
         Top             =   -120
         Width           =   2175
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Database Table"
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
         Left            =   8760
         MousePointer    =   99  'Custom
         TabIndex        =   49
         Top             =   660
         Width           =   1620
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
         Left            =   15360
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   660
         Width           =   1230
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
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   660
         Width           =   1200
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
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   6525
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
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
         Width           =   6045
      End
      Begin VB.Label lbRecords 
         BackStyle       =   0  'Transparent
         Caption         =   "Records"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   120
         Width           =   4335
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   0
         Left            =   9360
         MousePointer    =   99  'Custom
         Picture         =   "FormCode.frx":1E06
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   3
         Left            =   15600
         MousePointer    =   99  'Custom
         Picture         =   "FormCode.frx":51E8
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   4
         Left            =   18240
         Picture         =   "FormCode.frx":85CA
         Top             =   120
         Width           =   480
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   1
         Left            =   3960
         TabIndex        =   26
         Top             =   -240
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.PictureBox PBTitle 
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
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   1
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   4
         Left            =   7680
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   11
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Export DB"
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
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   640
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   4
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "FormCode.frx":B9AC
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   0
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   8
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "FormCode.frx":ED8E
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New"
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
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   640
            Width           =   1830
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   1920
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   6
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "FormCode.frx":12170
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Save"
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
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   640
            Width           =   1830
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   3840
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   4
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "FormCode.frx":15552
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Refresh Table"
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
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   640
            Width           =   1890
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   3
         Left            =   5760
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   2
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   720
            MousePointer    =   99  'Custom
            Picture         =   "FormCode.frx":18934
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Delete"
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
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   640
            Width           =   1890
         End
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hanna Code"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   630
         Left            =   10560
         TabIndex        =   10
         Top             =   120
         Width           =   8370
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   19200
      _ExtentX        =   33867
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox ctlCalendar1 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   0
      Width           =   1000
   End
   Begin VB.Frame frRevisionHistory 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9375
      Left            =   240
      TabIndex        =   52
      Top             =   3120
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Frame frInside 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "16800"
         Height          =   9015
         Index           =   6
         Left            =   1080
         TabIndex        =   53
         Top             =   360
         Width           =   17055
         Begin VB.Frame frExcel 
            BackColor       =   &H00206020&
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
            Left            =   960
            TabIndex        =   81
            Top             =   6960
            Width           =   3015
            Begin VB.Label lbExcel 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Export Excel"
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
               Left            =   0
               TabIndex        =   82
               Top             =   120
               Width           =   3015
            End
            Begin VB.Image Image 
               Height          =   480
               Left            =   120
               MousePointer    =   99  'Custom
               OLEDropMode     =   1  'Manual
               Picture         =   "FormCode.frx":1BD16
               Top             =   0
               Width           =   480
            End
         End
         Begin VB.ComboBox cmbRevType 
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9720
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   5760
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox txRevision 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   4
            Left            =   2160
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   78
            Top             =   6240
            Width           =   13815
         End
         Begin VB.TextBox txRevision 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   13560
            TabIndex        =   76
            Top             =   5760
            Width           =   2415
         End
         Begin VB.TextBox txRevision 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   6240
            TabIndex        =   74
            Top             =   5760
            Width           =   2415
         End
         Begin VB.TextBox txRevision 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   9720
            TabIndex        =   72
            Top             =   5760
            Width           =   2415
         End
         Begin VB.TextBox txRevision 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   2160
            TabIndex        =   70
            Top             =   5760
            Width           =   2415
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "l"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   4
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   17055
            Begin VB.Line Line10 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   16920
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label lbInside 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Revision History"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00606060&
               Height          =   375
               Index           =   5
               Left            =   0
               TabIndex        =   61
               Top             =   75
               Width           =   13215
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Recipe Formulation"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00644603&
               Height          =   255
               Left            =   15165
               TabIndex        =   60
               Top             =   120
               Width           =   1725
            End
         End
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
            Index           =   9
            Left            =   12960
            TabIndex        =   57
            Top             =   6960
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Clear form"
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
               Index           =   9
               Left            =   0
               TabIndex        =   58
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00886010&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   5880
            TabIndex        =   54
            Top             =   2400
            Width           =   5055
            Begin VB.Label lbChem 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Enter all fileds and Save"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   6
               Left            =   1380
               TabIndex        =   56
               Top             =   720
               Width           =   2340
            End
            Begin VB.Label lbChem 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Empty List..."
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   55
               Top             =   360
               Width           =   4995
            End
         End
         Begin FlexCell.Grid Grid2 
            Height          =   4695
            Left            =   0
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   600
            Width           =   16935
            _ExtentX        =   29871
            _ExtentY        =   8281
            AllowUserSort   =   -1  'True
            Appearance      =   0
            BackColor1      =   15790320
            BackColor2      =   15790320
            BackColorBkg    =   15790320
            BackColorFixed  =   15790320
            BackColorFixedSel=   15790320
            BackColorScrollBar=   15592423
            BorderColor     =   15790320
            CellBorderColor =   15790320
            CellBorderColorFixed=   15790320
            Cols            =   5
            DefaultFontName =   "Segoe UI"
            DefaultFontSize =   8.25
            BoldFixedCell   =   0   'False
            DisplayRowIndex =   -1  'True
            DrawMode        =   1
            DefaultRowHeight=   20
            FixedRowColStyle=   0
            ForeColorFixed  =   6571523
            GridColor       =   15790320
            Rows            =   1
            ScrollBarStyle  =   0
            SelectionMode   =   3
            MultiSelect     =   0   'False
            DateFormat      =   0
         End
         Begin VB.Label lbFunction 
            BackStyle       =   0  'Transparent
            Height          =   855
            Index           =   5
            Left            =   8400
            TabIndex        =   65
            Top             =   7440
            Width           =   1815
         End
         Begin VB.Label lbFunction 
            BackStyle       =   0  'Transparent
            Height          =   855
            Index           =   4
            Left            =   6840
            TabIndex        =   64
            Top             =   7440
            Width           =   1575
         End
         Begin VB.Label lbAcquisition 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   240
            TabIndex        =   79
            Top             =   6240
            Width           =   1695
         End
         Begin VB.Label lbAcquisition 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Operator"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   12480
            TabIndex        =   77
            Top             =   5760
            Width           =   855
         End
         Begin VB.Label lbAcquisition 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Rev Number"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   4920
            TabIndex        =   75
            Top             =   5760
            Width           =   1215
         End
         Begin VB.Label lbAcquisition 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   8880
            TabIndex        =   73
            Top             =   5760
            Width           =   735
         End
         Begin VB.Label lbAcquisition 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Rev Date"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   240
            TabIndex        =   71
            Top             =   5760
            Width           =   1695
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delete Specifics"
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
            Height          =   255
            Left            =   8640
            TabIndex        =   67
            Top             =   7875
            Width           =   1500
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Save Specifics"
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
            Height          =   255
            Left            =   6960
            TabIndex        =   66
            Top             =   7875
            Width           =   1335
         End
         Begin VB.Image ImCode 
            Height          =   240
            Index           =   5
            Left            =   7440
            Picture         =   "FormCode.frx":1F0F8
            ToolTipText     =   "4000"
            Top             =   7485
            Width           =   240
         End
         Begin VB.Image ImCode 
            Height          =   240
            Index           =   4
            Left            =   9120
            Picture         =   "FormCode.frx":1FAFA
            ToolTipText     =   "4000"
            Top             =   7485
            Width           =   240
         End
         Begin VB.Line Line11 
            BorderColor     =   &H00D0D0D0&
            X1              =   0
            X2              =   16920
            Y1              =   5400
            Y2              =   5400
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter/Edit/ Delete  Revision Specifics : Enter all fields and Save or Export to Excel"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Index           =   2
            Left            =   5265
            TabIndex        =   63
            Top             =   8640
            Width           =   6435
         End
      End
   End
End
Attribute VB_Name = "FormDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


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

Private MyLbHelpCount As Integer

Private IndexTabella As Integer
Private MaxIndex As Integer



Private bHilight As Boolean

Private m_rc As Boolean

Private MyID As Long
Private MyIndexRecord As Integer
Private lRow As Long
Private lCol As Long
Private DatabaseIndex As Integer

Private DatabaseString As String
Private dbTab As ADODB.Recordset
Private ID_CHEMICAL As Long


Private uMR As MRType

Private IndexVisibleFrame As Integer

Private SelectedCode As String
Private UmComponent As String

Private HannaCode As String
Private bSetPercentageLastComponent As Boolean
Private bFlagOpenRecipe As Boolean

Private RevisionID As Boolean
Private bCloneRecipe As Boolean
Private MRLocation As String
Private MRDescription As String



Private Sub cmbType_Click()
If lRow > 0 Then
    Grd2.Cell(lRow, 2).Text = cmbType
    Grd2.Cell(lRow, 2).Alignment = cellCenterCenter
    
End If
End Sub



Private Sub cmbTypeHannaCode_Click()
GlobalSearch
End Sub





Private Sub cmbRevType_Click()
txRevision(1) = cmbRevType
cmbRevType.Visible = False
End Sub

Private Sub cmbUM_Click()
If lRow > 0 Then
    Grd2.Cell(lRow, 2).Text = cmbUM
    Grd2.Cell(lRow, 2).Alignment = cellCenterCenter
    If SetbUmMassa(cmbUM) Then
        If Grd2.Cell(8, 2).Text = "" Then
            Grd2.Cell(8, 2).Text = 1
        End If
    End If
End If
End Sub

Private Sub Form_Activate()
Me.WindowState = MainWindowState
End Sub

Private Sub Form_Initialize()
SaveSizes
End Sub

Private Sub Form_Unload(Cancel As Integer)

Timer2.Enabled = False

Set FormDatabase = Nothing

End Sub

Private Sub frExcel_Click()

    Grid2.ExportToExcel USER_DESKTOP & "\" & FormatNomeFile(HannaCode) & "_RevHistory.xls", True, True
    MessageInfoTime = 2500
    PopupMessage 2, "File correcly created on Desktop", , , FormatNomeFile(HannaCode) & "_RevHistory.xls"
End Sub


Private Sub Grd2_ButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim Value As String
Value = Grd2.Cell(Row, Col).Text
Select Case Row
    Case 12
        lRow = Row
        If Value <> "" Then cmbType = Value
        
        PicType.Visible = True
    Case 10, 13, 15
     
    Case 14
    
End Select
End Sub

Private Sub Grd2_OwnerDrawCell(ByVal Row As Long, ByVal Col As Long, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Handled As Boolean)
Dim rc As Boolean
If DatabaseIndex = 4 Then
    
    rc = IIf(Len(Grd2.Cell(Row, Col).Text) > 0, True, False)
    
    
    If Row = 1 And Col = 2 Then
            
            
            PicMenu(6).Visible = rc
        
    End If
    If Row = 12 And Col = 2 Then
        Debug.Print Grd2.Cell(Row, Col).Text
        frCommandInside(5).Visible = rc
        frCommandInside(8).Visible = rc
        
    End If
End If

End Sub



Private Sub Grid2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)


RevisionID = 0

With Grid2

    If FirstRow > 0 Then
    
        RevisionID = .Cell(FirstRow, 6).Text
        txRevision(0) = .Cell(FirstRow, 1).Text
        txRevision(1) = .Cell(FirstRow, 3).Text
        txRevision(2) = .Cell(FirstRow, 2).Text
        txRevision(3) = .Cell(FirstRow, 5).Text
        txRevision(4) = .Cell(FirstRow, 4).Text
        
    End If

End With
End Sub

Private Sub Image_Click()
frExcel_Click
End Sub



Private Sub InExport_Click()

    If CheckPrivilege(3) Then
        If F_MsgBox.DoShow("Export Critical " & DatabaseString & " to Excel?") Then
        
            Dim ExcelName As String
            ExcelName = FormatNomeFile("Critical " & DatabaseString & " List_" & FormatDataLAT(Now())) & ".xls"
            
            If F_InputBox.DoShow("Please Set Excel Name", "Critical List", , , , ExcelName) Then
    
                ExcelName = USER_DESKTOP & "\" & FormatNomeFile(ExcelName) & ".xls"
                
                
                Grd1.ExportToExcel ExcelName, True, True
                MessageInfoTime = 2500
                PopupMessage 2, "Excel Done..." & vbCrLf & ExcelName & vbCrLf & "File on Desktop."
                
            End If
            
        End If
    End If
            
End Sub

Private Sub lbExcel_Click()
frExcel_Click
End Sub

Private Sub lbInside_Click(Index As Integer)
    Select Case Index
        
        Case 5
            ' rev history table
            If HannaCode <> "" Then Call GetRecipeRevision(Grid2, HannaCode)
    
    End Select
End Sub



Private Sub txRevision_Change(Index As Integer)
Dim rc As Boolean
rc = IIf(Len(txRevision(Index)) > 0, True, False)
txRevision(Index).BackColor = IIf(rc, vbWhite, &HE0E0E0)
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
        ElseIf TypeOf ctl Is Menu Then
        ElseIf TypeOf ctl Is Inet Then
        ElseIf TypeOf ctl Is Image Then
            ctl.Left = (x_scale * .Left) + IIf(x_scale = 1, 0, (x_scale - 1) * 200)
            ctl.Top = y_scale * .Top
        ElseIf TypeOf ctl Is Timer Then
        ElseIf TypeOf ctl Is ucScrollAdd Then
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



Private Sub Form_Load()
Dim i As Integer


If Screen.Width - Me.Width > 1000 And bFullScreen Then
    Me.WindowState = 2
    'Me.Picture = LoadPicture()
   
End If


  MyID = 0
  MyIndexRecord = 3



   

End Sub


Public Function DoShow(ByVal Index As Integer) As Boolean
Dim i As Integer
    On Error GoTo ERR_SHOW
    
    m_rc = False
    mOk
    
    
    Call SetDatabaseType(Index)
    
    Call SetPretarationType(cmbType)

Dim sStr1 As String
Dim sStr2 As String


Call GetLastImport(sStr1, sStr2)

Label11(0) = sStr1
Label11(1) = sStr2




    Call CopyDatabaseGrd1
    
    

    Me.Show vbModal
    
    

    
    If m_rc = True Then
    End If
ERR_END:
    On Error GoTo 0
    DoShow = m_rc
    Exit Function
ERR_SHOW:
    m_rc = False
    MsgBox "DOSHOW ERROR  " & Err.Description
    Resume ERR_END
End Function




Private Sub frChemicals_Click()
ImCode_Click 0
End Sub

Private Sub frCommandInside_Click(Index As Integer)
Select Case Index

    Case 0
        ' save
        
    Case 1
     
        Lab(7) = "Exit Database Table"
        blTable = DatabaseString
    Case 2
    
        

        
    Case 3
       
    Case 4
        If MyID > 0 Then Call F_PICTOGRAM.DoShow(MyID, 1)
        
        
    Case 5

        
    Case 6
        frInside(5).Visible = False
   
       
    Case 7
        Call F_PICTOGRAM.DoShow(0, 0, Grd2.Cell(1, 2).Text, False)
        
    Case 8
        Call OpenRevisionHistory
    Case 9
        Call ClearRevisionForm
        
End Select
End Sub

Private Sub frHannaCode_Click()
ImCode_Click 2
End Sub


Private Sub Grd1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)



If FirstRow > 0 Then
    
    
    Select Case DatabaseIndex
        Case 0
            ' hanna code
            
                
                MyID = Grd1.Cell(FirstRow, 3).Text
                Call SetGridEditCode(Grd2)
                Call CopyCodeGrd2(Grd2, MyID)
                
        Case 1
            ' pipette
            
                MyID = Grd1.Cell(FirstRow, 4).Text
                Call SetGridEditPipetta(Grd2)
                Call CopyPipettaGrd2(Grd2, MyID)

        Case 2
            ' Code Classification
                
                MyID = Grd1.Cell(FirstRow, 3).Text
                Call SetGridEditClassification(Grd2)
                Call CopyClassificationGrd2(Grd2, MyID)
                
                frCommandInside(7).Visible = IIf(Len(Grd2.Cell(Grd2.Rows - 1, 2).Text) > 0, True, False)
                
     
        Case 3
            ' Frasi H
                MyID = Grd1.Cell(FirstRow, 3).Text
                Call SetGridEditFrasiH(Grd2)
                Call CopyFrasiHGrd2(Grd2, MyID)
                      
                      
        Case 4

                
        Case 5
            ' Chemical MR
                MyID = Grd1.Cell(FirstRow, 3).Text
                HannaCode = Trim(Grd1.Cell(FirstRow, 1).Text)
                Call SetGridEditChemicalRM(Grd2)
                Call CopyChemicalRMGrd2(Grd2, MyID, MRLocation, MRDescription)
    
                
              
                
    End Select
End If
End Sub




Private Sub grd1_Click()
ctlCalendar1.Visible = False

bCloneRecipe = False
 
End Sub

Private Sub Grd2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

Dim Value As String
lRow = 0

PicType.Left = Grd2.Width / 2 - PicType.Width / 2
PicType.Top = Grd2.Height / 2 - PicType.Height / 2

PicUm.Left = Grd2.Width / 2 - PicUm.Width / 2
PicUm.Top = Grd2.Height / 2 - PicUm.Height / 2

PicUm.Visible = False
PicType.Visible = False

MessageInfoTime = 2000
    Select Case DatabaseIndex
        Case 1
            ' Production Way
        Case 2
            ' Code Classification
          
                Select Case FirstCol
                    Case 1
                       ' Dim MyID As Long
                        
                        If FirstRow = 7 Then
                           ' MyID = GetIDMR(Grd2.Cell(1, 2).Text)
                           ' Call F_PICTOGRAM.DoShow(0, 0, Grd2.Cell(1, 2).Text, False)
                        
                        End If
                
                End Select
                
                
        Case 3
            ' Frasi H

        Case 4
          

        Case 5
            ' Chemical MR
        
                
                Select Case FirstCol
                    Case 1
                        Dim MyID As Long
                        
                       ' If FirstRow = 5 And IsNumeric(Grd2.Cell(1, 2).Text) Then
                          ' MyID = GetMRbyID(Grd2.Cell(1, 2).Text)
                          '  Call F_PICTOGRAM.DoShow(MyID, 1, , False)
                      ' End If
                
                End Select

            
        
    
    End Select


End Sub


Private Sub Image1_Click()
PicUm_Click
End Sub

Private Sub Image2_Click()
PicType_Click
End Sub

Private Sub Image3_Click(Index As Integer)

    If CheckPrivilege(3) = False Then Exit Sub

    Select Case Index
        Case 0
            ' pulisci maschera
           
            Call CleanCode
            Grd2.Cell(1, 2).SetFocus
        Case 1
           
                If CheckPrivilege(3) Then
                    If SaveRecord Then
                        Select Case DatabaseIndex
                            Case 4 ' recipe
                               
                
                            Case 5 ' rawmaterials
                                frCommandInside(4).Visible = True
                        
                        End Select
                        CopyDatabaseGrd1
                    End If
                End If
         
        Case 2
            ' refresh table
         
            Text1(0) = ""
            CopyDatabaseGrd1
        Case 3
          
            If CheckPrivilege(3) Then CancellaTab
        Case 4
         
            If CheckPrivilege(3) Then
                If F_MsgBox.DoShow("Export " & DatabaseString & " to Excel?") Then
                
                    Call DBCodeToExcel(ProgressBar1, DatabaseIndex, DatabaseString)
                End If
            End If
        Case 5
        
        Case 6
          
    
    End Select
End Sub



Private Sub Form_Resize()
On Error Resume Next


ResizeControls
frRevisionHistory.Move frInside(0).Left, frInside(0).Top, Me.Width - frInside(0).Left * 2, frInside(0).Height

End Sub




Private Sub ctlCalendar1_DateClicked(inputDate As Date)
ctlCalendar1.Visible = False
End Sub



Private Sub DefaultMenu_Click(Index As Integer)
Dim MyIndex As Integer
Select Case Index
    Case 0
      
        If frRevisionHistory.Visible Then
            
            frRevisionHistory.Visible = False
       
        Else
        
         
            If F_MsgBox.DoShow("Exit " & DatabaseString & " database?", "Database") Then
                Unload Me
            End If
        End If
        
        
    Case 2
        ' Open Report folder
        OpenWithDefault (USER_DOCUMENTI & PathRequisition)
      
    Case 1
        ' filtro
        
       
        
    Case 3
        

            ' avanti di 10
            Call ScorriTabella(False)
       
    Case 4
        
       
            ' indietro di 10
            Call ScorriTabella(True)
      
    Case 5
   
    Case 6
      
    Case 7
     
    Case 8
        m_rc = True
        Unload Me
    Case 9
       
    Case 10
      
    Case 11
      
    Case 12

        
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
         DefaultMenu_Click 3
    Case 39
        DefaultMenu_Click 4
    Case 38
        DefaultMenu_Click 3
    Case 40
        DefaultMenu_Click 4
        
End Select
End Sub


Private Function RiempiGrid(ByRef Grd As Grid, Optional ByVal Code As String)
Dim i As Integer
Dim t As Integer
Dim MaxCount As Integer

    On Error GoTo ERR_GRID
    ' --------------------------------------
    '
    ' --------------------------------------
    frCommandInside(5).Visible = False
    frCommandInside(8).Visible = False
    For i = 1 To Grd2.Rows - 1
        Grd2.Cell(i, 2).Text = ""
    Next
    
    
    Grd2.Cell(1, 2).BackColor = &HFFFFFF
    Grd2.Cell(1, 2).ForeColor = vbBlack
    
    CopyDatabaseGrd1
    
    
    
    
   
ERR_END:
   
    IndexTabella = 1
    MaxIndex = IIf(Int((Grd2.Rows - 1) / 10) < (Grd2.Rows - 1) / 10, (Int((Grd2.Rows - 1) / 10)) + 1, Int((Grd2.Rows - 1) / 10))
    If MaxIndex = 0 Then MaxIndex = 1

    Exit Function
ERR_GRID:
    MessageInfoTime = 2000
    Text1(0) = ""
    PopupMessage 2, Err.Description
    GoTo ERR_END:
End Function



Private Sub ImageTAV_Click(Index As Integer)
Select Case Index
        Case 0
            Unload Me
        
        Case 2
        

End Select
End Sub

Private Sub ImCode_Click(Index As Integer)
Dim UserSTDNumber As String
Dim UserHannaCode As String
Dim MyCodeID As Long


    frCommandInside(2).Visible = False
    

    Select Case Index
        Case 0

        Case 1

        Case 2

            
        Case 3

                        
        Case 4
            ' aggiungi revision specifics
            If AddRevision(HannaCode, txRevision(2)) Then
                 Call GetRecipeRevision(Grid2, HannaCode)
            End If
            
            frExcel.Visible = IIf(Grid2.Rows > 1, True, False)
            Frame6.Visible = IIf(Grid2.Rows > 1, False, True)
        Case 5
            ' delete revision specifics
            If DeleteRevision(HannaCode, txRevision(2)) Then
                 Call GetRecipeRevision(Grid2, HannaCode)
            End If
            
            frExcel.Visible = IIf(Grid2.Rows > 1, True, False)
            Frame6.Visible = IIf(Grid2.Rows > 1, False, True)
            
            
            
    End Select
End Sub


Private Sub Label1_Click(Index As Integer)
Select Case Index
    Case 0
        
    Case 1
        frChemicals_Click
End Select

End Sub

Private Sub Label2_Click(Index As Integer)
Image3_Click Index
End Sub



Private Sub lbChem_Click(Index As Integer)
Select Case Index
    Case 0, 1
        frChemicals_Click
    Case 2, 3
        frHannaCode_Click
End Select


End Sub

Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub

Private Sub lbFunction_Click(Index As Integer)
ImCode_Click Index
End Sub

Private Sub PicType_Click()
PicType.Visible = False

End Sub


Private Sub PicUm_Click()
PicUm.Visible = False


End Sub


Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
For i = 0 To PicMenu.Count - 1
    If i = Index Then
        PicMenu(i).BackColor = &H886010
    Else
        PicMenu(i).BackColor = &H644603
    End If
Next
End Sub

Private Sub PicMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Image3_Click Index
End Sub



Private Sub Text1_Change(Index As Integer)
If InStr(UCase(Text1(0)), UCase("code")) Then

Else
    
   GlobalSearch

End If



End Sub


Private Sub Timer2_Timer()

Dim i As Integer
    '
    ' start form
    '
      bHilight = True
 
    Call CleanCode
 
    
    Timer2.Enabled = False
    
    

    
End Sub



Private Sub ScorriTabella(ByVal bValue As Boolean)

Dim MyRow As Integer
If Grd1.Rows > 1 Then
    MyRow = IIf(bValue, (IndexTabella * 10) + 20, (IndexTabella * 10) - 19)
    IndexTabella = IIf(bValue, IndexTabella + 1, IndexTabella - 1)
    If IndexTabella < 1 Then
        IndexTabella = 1
        Grd1.Cell(1, 1).EnsureVisible
    ElseIf MyRow >= Grd1.Rows Then
        Grd1.Cell(Grd1.Rows - 1, 1).EnsureVisible
        IndexTabella = MaxIndex
    'ElseIf IndexTabella >= MaxIndex - 1 And Not (bValue) Then
        'grd1.Cell((IndexTabella) * 10, 1).EnsureVisible
    Else
         
        Grd1.Cell(MyRow, 1).EnsureVisible
    
    End If
End If

End Sub


Private Function SetNumCol(ByVal sString As String) As Integer

Select Case Trim(UCase(sString))
    Case UCase(("Code"))
        SetNumCol = 2
    Case UCase(("Description"))
        SetNumCol = 3
    Case UCase(("Lot"))
        SetNumCol = 1
    'Case UCase(("Cliente"))
       ' SetNumCol = 5
    
    End Select

End Function
Private Function GlobalSearch()


    RiempiGrid Grd1, Text1(0)
End Function



Private Function CancellaTab() As Boolean
        If F_MsgBox.DoShow(("Delete Selected Record ?"), "Database", , ("Delete"), ("No")) Then
            
            If DeleteRecord(MyID) Then
               ' Text1(0) = ""
                GlobalSearch
            Else
            End If
        End If
End Function

Private Sub CleanCode()
Dim i As Integer
On Error GoTo ERR_CLEAN:
With Grd2
    For i = 1 To .Rows - 1
        .Cell(i, 2).Text = ""
    Next
    
    Select Case DatabaseIndex
        Case 4
        
            .Cell(15, 2).Text = "pcs"
            .Cell(15, 2).Alignment = cellCenterCenter
            
    End Select
End With

MyID = 0
lRow = 0
ID_CHEMICAL = 0
UmComponent = ""
SelectedCode = ""

ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_CLEAN:
    GoTo ERR_END:
End Sub





'
'----------- edit insert code ---------------'
'


Private Sub Grd2_EditRow(ByVal Row As Long)
Debug.Print "Edit Row ", Row
Debug.Print
lRow = Row


End Sub

Private Sub Grd2_LeaveCell(ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)


    Select Case DatabaseIndex
        Case 0
            ' hanna code
            
            Call Grd2_Code_LeaveCell(Grd2, Row, Col, NewRow, NewCol, Cancel, lRow)

        Case 1
            ' Pipetta
            Call Grd2_Pipetta_LeaveCell(Grd2, Row, Col, NewRow, NewCol, Cancel, lRow)
       
        Case 2
            ' Code Classification
            Call Grd2_Classification_LeaveCell(Grd2, Row, Col, NewRow, NewCol, Cancel, lRow)
        Case 3
            ' Frasi H
            Call Grd2_FrasiH_LeaveCell(Grd2, Row, Col, NewRow, NewCol, Cancel, lRow)
        Case 4
            ' Recipe
         
        Case 5
            ' Chemical RM
           Call Grd2_ChemicalRM_LeaveCell(Grd2, Row, Col, NewRow, NewCol, Cancel, lRow)
            
        
    
    End Select




End Sub

Private Function SaveRecord() As Boolean
Dim rc As Boolean

If Grd2.Cell(1, 2).Text <> "" Then
    Select Case DatabaseIndex
        Case 0
            ' hanna code
            
            rc = SaveDatabaseCode(Grd2)

        Case 1
            ' Pipetta
           rc = SaveDatabasePipette(Grd2, MyID)
        Case 2
            ' Code Classification
            rc = SaveDatabaseClassification(Grd2)
        Case 3
            ' Frasi H
            rc = SaveDatabaseFrasiH(Grd2)
        Case 4
            ' Recipe
          
            
        Case 5
            ' Chemical RM
           rc = SaveDatabaseChemicalRM(Grd2, MRLocation, MRDescription)
            
        
    
    End Select
Else
    PopupMessage 2, "Please fill all required fields..", , True
End If



   
    
    
    
    
    
    SaveRecord = rc
End Function
Private Function DeleteRecord(ByVal MyID As Long) As Boolean
Dim rc As Boolean


    If MyID = 0 Then Exit Function
    
    
    rc = CancellaCodeByID(dbTab, MyID)
    DeleteRecord = rc

End Function


Private Sub SetDatabaseType(ByVal Index As Integer)

DatabaseIndex = Index

frCommandInside(7).Visible = False

    Select Case Index
        Case 0
            ' hanna code
            Set dbTab = dbTabCode
            DatabaseString = "Hanna Code"
            Call SetGridCode(Grd1)
            Call SetGridEditCode(Grd2)
            Call AddComboCode(Combo1)
          
  
        Case 1
            Set dbTab = dbTabPipette
            DatabaseString = "Equipment"
            Call SetGridPipette(Grd1)
            Call SetGridEditPipetta(Grd2)
            Call AddComboPipetta(Combo1)

        Case 2
            ' Code Classification
            Set dbTab = dbTabCodeClassification
            DatabaseString = "Product Classification"
            
            Call SetGridClassification(Grd1)
            Call SetGridEditClassification(Grd2)
            Call AddComboClassification(Combo1)
            frCommandInside(7).ZOrder
            'frCommandInside(7).Visible = True
            
        Case 3
            ' Frasi H
            Set dbTab = dbTabFrasiH
            DatabaseString = "Physical hazard statement"
            
            Call SetGridFrasiH(Grd1)
            Call SetGridEditFrasiH(Grd2)
            Call AddComboFrasiH(Combo1)
            
        Case 4

           
        Case 5
            ' Chemical RM
           
            
            Set dbTab = dbTabMR
            DatabaseString = "Chemical MR"
            
            Call SetGridChemicalRM(Grd1)
            Call SetGridEditChemicalRM(Grd2)
            Call AddComboChemicalRM(Combo1)
            
          '  Call SetChemicalsXRecipe
            
          
            frCommandInside(4).Visible = True
            
             Call SetUM(cmbUM)
    
    End Select



    
    blTable = DatabaseString
    
    
    
    


End Sub



Private Sub CopyDatabaseGrd1()

Grd1.Rows = 1

    Select Case DatabaseIndex
        Case 0

        Case 1
            Call FillGridPipette(Grd1, Text1(0), , Combo1)
            
        Case 2
            ' Code Classification
            Call CopyClassificationGrd1(Grd1, Text1(0), , Combo1)
        Case 3
            ' Frasi H
            Call CopyFrasiHGrd1(Grd1, Text1(0), , Combo1)
        Case 4
            ' Recipe
         
        Case 5
            ' Chemical RM
            Call CopyChemicalRMGrd1(Grd1, Text1(0), , Combo1)
        
    
    End Select

    lbRecords = "Database Records # " & Grd1.Rows - 1
  
End Sub


Public Function SetUM(ByVal cmb As ComboBox) As Boolean

    With cmb
        .Clear
       ' .AddItem "L"
       ' .AddItem "kg"
       ' .AddItem "pcs"
        .AddItem "mL"
        .AddItem "g"
       ' .AddItem "mg"
        .ListIndex = 1
    End With
End Function

Public Function SetUMPeso(ByVal cmb As ComboBox) As Boolean

    With cmb
        .Clear
        .AddItem "kg"
        .AddItem "g"
        .AddItem "mg"
        .AddItem "µg"
        .ListIndex = 1
    End With
End Function



Private Sub OpenRevisionHistory()

lbInside(5) = HannaCode & " : Revision History"

frExcel.Visible = IIf(Grid2.Rows > 1, True, False)
Frame6.Visible = IIf(Grid2.Rows > 1, False, True)

                
frRevisionHistory.BackColor = &HF0F0F0
frRevisionHistory.Move frInside(0).Left, frInside(0).Top, Me.Width - frInside(0).Left * 2, frInside(0).Height
frRevisionHistory.ZOrder
frRevisionHistory.Visible = True


End Sub


Private Sub ClearRevisionForm()



Dim i As Integer
For i = 0 To txRevision.UBound
    txRevision(i) = ""
Next
txRevision(3) = MyOperatore.Name

End Sub


Private Sub AddcmbRevType()


    With cmbRevType
        .AddItem "Revision"
        .AddItem "Improvement"
        .AddItem "Issue"
        .ListIndex = 0
    End With

End Sub
Private Sub txRevision_Click(Index As Integer)
Dim userCode As String
Dim Answer As String
Dim Selected As String
Dim bNumber As Boolean
Dim sString As String
Dim rc As Boolean

    Selected = lbAcquisition(Index) ' "Preparation"
    Answer = txRevision(Index)
    sString = "Please Enter " & lbAcquisition(Index)
  
    bNumber = False
    
    If txRevision(3) = "" Then txRevision(3) = MyOperatore.Name

    Select Case Index
        Case 0
        
            If Answer = "" Then Answer = FormatDataLAT(Now())
        Case 1
            ' type
            cmbRevType.ZOrder
            cmbRevType.Visible = True
            Exit Sub
        Case 2
            'rev number
             If uMR.Rev <> "" Then
                If IsNumeric(uMR.Rev) Then
                    Answer = uMR.Rev
                End If
             End If
    End Select
    
    
    If txRevision(Index).Locked Then Exit Sub
    
    
  
    If F_InputBox.DoShow(sString, Selected, , , , Answer, , bNumber, Me.Top) Then
    
        txRevision(Index) = Answer
        
        Select Case Index
            Case 0
                ' isdate?
                If IsDate(Answer) Then
                     txRevision(Index) = FormatDataLAT(Answer)
                Else
                    PopupMessage 2, "Please enter a valid Date...", , True
                End If
        End Select
    End If
    
    
    
    
End Sub

Private Function DeleteRevision(ByVal Code As String, ByVal RevNumber As String) As Boolean
Dim rc As Boolean
Dim i As Integer

rc = True

For i = 1 To txRevision.UBound
    If txRevision(i) = "" Then
        rc = False
        PopupMessage 2, "Please Select a Revision form the table...", , True, "Delete Revision"
        DeleteRevision = rc
        Exit Function
    End If
Next

With dbTabMRrevisionHistory
    .filter = ""
    .filter = "Recipe='" & Replace(Code, "'", "''") & "' and RevNumber='" & RevNumber & "'"
    If .EOF Then
        
    Else
        If F_MsgBox.DoShow("Delete Rev : " & RevNumber & "?", Code, , "Delete", "Exit") Then
            .Delete
            .Update
        Else
            rc = False
        End If
    End If
     
End With

DeleteRevision = rc
End Function
Private Function AddRevision(ByVal Code As String, ByVal RevNumber As String) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim OldDate As Date

rc = True

For i = 1 To txRevision.UBound
    If txRevision(i) = "" Then
        rc = False
        PopupMessage 2, "Please enter all fields...", , True, "Revision History"
        AddRevision = rc
        Exit Function
    End If
Next

If Code = "" Or RevNumber = "" Then
    
    rc = False
    PopupMessage 2, "Please enter all fields...", , True, "Revision History"
    AddRevision = rc
    Exit Function
        
End If




With dbTabMRrevisionHistory
    .filter = ""
    .filter = "Recipe='" & Replace(Code, "'", "''") & "' and RevNumber='" & RevNumber & "'"
    If .EOF Then
        .AddNew
    Else
        .MoveFirst
        OldDate = FormatDataLAT(Trim(!RevDate))
        If F_MsgBox.DoShow("Rev : " & RevNumber & " already exsists." & vbCrLf & "Rev Date : " & OldDate, Code, , "Modify", "Exit") Then
        
        Else
            AddRevision = False
            Exit Function
        End If
    End If
        
        !RevDate = txRevision(0)
        !Recipe = Code
        !RevNumber = txRevision(2)
        !Type = txRevision(1)
        !Description = IIf(Len(txRevision(4)) > 255, Left(txRevision(4), 255), txRevision(4))
        !Operator = txRevision(3)
        
        .Update
End With

AddRevision = rc
End Function



