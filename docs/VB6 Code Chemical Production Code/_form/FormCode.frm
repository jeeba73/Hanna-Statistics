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
   Begin VB.Frame frRevisionHistory 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9375
      Left            =   1200
      TabIndex        =   120
      Top             =   3480
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Frame frInside 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "16800"
         Height          =   9015
         Index           =   6
         Left            =   1080
         TabIndex        =   121
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
            TabIndex        =   149
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
               TabIndex        =   150
               Top             =   120
               Width           =   3015
            End
            Begin VB.Image Image 
               Height          =   480
               Left            =   120
               MousePointer    =   99  'Custom
               OLEDropMode     =   1  'Manual
               Picture         =   "FormCode.frx":0A02
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
            TabIndex        =   148
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
            TabIndex        =   146
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
            TabIndex        =   144
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
            TabIndex        =   142
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
            TabIndex        =   140
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
            TabIndex        =   138
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
            TabIndex        =   127
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
               TabIndex        =   129
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
               TabIndex        =   128
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
            TabIndex        =   125
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
               TabIndex        =   126
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00886010&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   5880
            TabIndex        =   122
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
               TabIndex        =   124
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
               TabIndex        =   123
               Top             =   360
               Width           =   4995
            End
         End
         Begin FlexCell.Grid Grid2 
            Height          =   4695
            Left            =   0
            TabIndex        =   130
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
            TabIndex        =   133
            Top             =   7440
            Width           =   1815
         End
         Begin VB.Label lbFunction 
            BackStyle       =   0  'Transparent
            Height          =   855
            Index           =   4
            Left            =   6840
            TabIndex        =   132
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
            TabIndex        =   147
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
            TabIndex        =   145
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
            TabIndex        =   143
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
            TabIndex        =   141
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
            TabIndex        =   139
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
            TabIndex        =   135
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
            TabIndex        =   134
            Top             =   7875
            Width           =   1335
         End
         Begin VB.Image ImCode 
            Height          =   240
            Index           =   5
            Left            =   7440
            Picture         =   "FormCode.frx":3DE4
            ToolTipText     =   "4000"
            Top             =   7485
            Width           =   240
         End
         Begin VB.Image ImCode 
            Height          =   240
            Index           =   4
            Left            =   9120
            Picture         =   "FormCode.frx":47E6
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
            TabIndex        =   131
            Top             =   8640
            Width           =   6435
         End
      End
   End
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
      Begin VB.PictureBox PicLine 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         Height          =   2295
         Left            =   480
         ScaleHeight     =   2295
         ScaleWidth      =   6975
         TabIndex        =   39
         Top             =   1080
         Visible         =   0   'False
         Width           =   6975
         Begin VB.ComboBox cmbLine 
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
            Caption         =   "Select Line from list"
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
            Left            =   2445
            TabIndex        =   41
            Top             =   600
            Width           =   1830
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   6600
            Picture         =   "FormCode.frx":51E8
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
            Picture         =   "FormCode.frx":5BEA
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
            TabIndex        =   136
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
               TabIndex        =   137
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
         Top             =   960
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
         TabIndex        =   89
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
            TabIndex        =   90
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
         TabIndex        =   152
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
         Begin VB.CheckBox chMix 
            BackColor       =   &H00F0F0F0&
            Caption         =   "Only Recipes/Mixes"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00644603&
            Height          =   285
            Left            =   3840
            TabIndex        =   42
            Top             =   0
            Visible         =   0   'False
            Width           =   2295
         End
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
   Begin VB.PictureBox PBContainerViewport 
      BorderStyle     =   0  'None
      Height          =   9735
      Left            =   480
      ScaleHeight     =   9735
      ScaleWidth      =   19215
      TabIndex        =   43
      Top             =   1320
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Frame PBContainer 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   29615
         Left            =   0
         TabIndex        =   44
         Top             =   -960
         Width           =   19215
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "&H00F0F0F0&"
            Height          =   7335
            Index           =   5
            Left            =   1200
            TabIndex        =   91
            Top             =   20040
            Visible         =   0   'False
            Width           =   15255
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
               Index           =   6
               Left            =   6720
               TabIndex        =   98
               Top             =   4680
               Width           =   3255
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Back To Recipe"
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
                  Index           =   6
                  Left            =   0
                  TabIndex        =   99
                  Top             =   120
                  Width           =   3255
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H000040C0&
               BorderStyle     =   0  'None
               Height          =   1335
               Left            =   5880
               TabIndex        =   95
               Top             =   1800
               Width           =   5055
               Begin VB.Label lbChem 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Open Recipe/Chemical to add Components..."
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
                  Index           =   4
                  Left            =   315
                  TabIndex        =   100
                  Top             =   720
                  Width           =   4380
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Empty List..."
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   2
                  Left            =   1920
                  TabIndex        =   96
                  Top             =   360
                  Width           =   1155
               End
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
               Index           =   3
               Left            =   0
               TabIndex        =   92
               Top             =   0
               Width           =   15255
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Formulation"
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
                  Left            =   14100
                  TabIndex        =   94
                  Top             =   180
                  Width           =   1050
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe Components : D002/1"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000040C0&
                  Height          =   285
                  Index           =   4
                  Left            =   0
                  TabIndex        =   93
                  Top             =   120
                  Width           =   3345
               End
               Begin VB.Line Line8 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   0
                  X2              =   15240
                  Y1              =   480
                  Y2              =   480
               End
            End
            Begin FlexCell.Grid Grid3 
               Height          =   3735
               Left            =   0
               TabIndex        =   97
               TabStop         =   0   'False
               Top             =   600
               Width           =   15255
               _ExtentX        =   26908
               _ExtentY        =   6588
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
               ForeColorFixed  =   8937488
               GridColor       =   15790320
               Rows            =   1
               ScrollBarStyle  =   0
               SelectionMode   =   3
               MultiSelect     =   0   'False
               DateFormat      =   0
            End
            Begin VB.Line Line9 
               BorderColor     =   &H00D0D0D0&
               X1              =   120
               X2              =   15240
               Y1              =   4440
               Y2              =   4440
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "&H00F0F0F0&"
            Height          =   6135
            Index           =   3
            Left            =   1200
            TabIndex        =   64
            Top             =   8680
            Width           =   17055
            Begin FlexCell.Grid Grid1 
               Height          =   3615
               Left            =   120
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   -240
               Width           =   16935
               _ExtentX        =   29871
               _ExtentY        =   6376
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
            Begin VB.Frame frHannaCode 
               Appearance      =   0  'Flat
               BackColor       =   &H000040C0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1335
               Left            =   5880
               TabIndex        =   68
               Top             =   1080
               Visible         =   0   'False
               Width           =   5055
               Begin VB.Label lbChem 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Click + to add chemicals"
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
                  Index           =   3
                  Left            =   45
                  TabIndex        =   76
                  Top             =   720
                  Width           =   5010
               End
               Begin VB.Label lbChem 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Empty List..."
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   2
                  Left            =   0
                  TabIndex        =   75
                  Top             =   360
                  Width           =   4995
               End
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
               Index           =   2
               Left            =   0
               TabIndex        =   65
               Top             =   0
               Width           =   17055
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe Link with Hanna Codes"
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
                  Index           =   2
                  Left            =   14115
                  TabIndex        =   67
                  Top             =   180
                  Width           =   2655
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hanna Codes"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00606060&
                  Height          =   345
                  Index           =   3
                  Left            =   0
                  TabIndex        =   66
                  Top             =   75
                  Width           =   1890
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   0
                  X2              =   16800
                  Y1              =   480
                  Y2              =   480
               End
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Delete Code"
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
               Left            =   8760
               TabIndex        =   74
               Top             =   5040
               Width           =   1245
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Add  Code"
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
               Left            =   7080
               TabIndex        =   73
               Top             =   5040
               Width           =   1080
            End
            Begin VB.Label lbFunction 
               BackStyle       =   0  'Transparent
               Height          =   855
               Index           =   3
               Left            =   8400
               TabIndex        =   72
               Top             =   4275
               Width           =   1815
            End
            Begin VB.Label lbFunction 
               BackStyle       =   0  'Transparent
               Height          =   855
               Index           =   2
               Left            =   6840
               TabIndex        =   71
               Top             =   4275
               Width           =   1575
            End
            Begin VB.Line Line6 
               BorderColor     =   &H00D0D0D0&
               X1              =   0
               X2              =   16800
               Y1              =   4320
               Y2              =   4320
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Select Hanna Code and link to Recipe"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   70
               Top             =   5400
               Width           =   17055
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   2
               Left            =   7440
               Picture         =   "FormCode.frx":65EC
               Top             =   4680
               Width           =   240
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   3
               Left            =   9120
               Picture         =   "FormCode.frx":6FEE
               Top             =   4680
               Width           =   240
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1455
            Index           =   4
            Left            =   1200
            TabIndex        =   59
            Top             =   18000
            Width           =   17055
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
               Left            =   8520
               TabIndex        =   62
               Top             =   600
               Width           =   3255
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Close"
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
                  TabIndex        =   63
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
               Index           =   0
               Left            =   5040
               TabIndex        =   60
               Top             =   600
               Width           =   3255
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Save Formulation"
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
                  TabIndex        =   61
                  Top             =   120
                  Width           =   3255
               End
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "16800"
            Height          =   7335
            Index           =   2
            Left            =   1200
            TabIndex        =   45
            Top             =   960
            Width           =   17055
            Begin VB.Frame SetPercentageLastComponent 
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
               Left            =   10800
               TabIndex        =   114
               Top             =   4560
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbSetPercentageLastComponent 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Set   %   last component"
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
                  TabIndex        =   115
                  Top             =   120
                  Width           =   3015
               End
            End
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
               Index           =   2
               Left            =   0
               TabIndex        =   112
               Top             =   4560
               Visible         =   0   'False
               Width           =   4815
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Open Component Table"
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
                  Index           =   2
                  Left            =   0
                  TabIndex        =   113
                  Top             =   120
                  Width           =   4815
               End
            End
            Begin VB.PictureBox PicUmComponent 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               CausesValidation=   0   'False
               Height          =   2295
               Left            =   480
               ScaleHeight     =   2295
               ScaleWidth      =   6975
               TabIndex        =   101
               Top             =   1080
               Visible         =   0   'False
               Width           =   6975
               Begin VB.ComboBox cmbUM2 
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
                  TabIndex        =   102
                  Top             =   1200
                  Width           =   2655
               End
               Begin VB.Image Image6 
                  Height          =   240
                  Left            =   6600
                  Picture         =   "FormCode.frx":79F0
                  Top             =   120
                  Width           =   240
               End
               Begin VB.Label lbUM2 
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
                  TabIndex        =   104
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
                  Index           =   0
                  Left            =   2820
                  TabIndex        =   103
                  Top             =   720
                  Width           =   1380
               End
            End
            Begin VB.Frame frQuantityCheck 
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   0  'None
               Caption         =   "Frame7"
               Height          =   615
               Left            =   1800
               TabIndex        =   80
               Top             =   6000
               Width           =   12855
               Begin VB.PictureBox PicPerc 
                  BackColor       =   &H000040C0&
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   2280
                  ScaleHeight     =   255
                  ScaleWidth      =   255
                  TabIndex        =   81
                  Top             =   240
                  Width           =   255
               End
               Begin VB.Image imPerc 
                  Height          =   240
                  Left            =   11880
                  Picture         =   "FormCode.frx":83F2
                  Top             =   240
                  Width           =   240
               End
               Begin VB.Label lbTotalWL 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   9840
                  TabIndex        =   87
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.Label lbTotalWKg 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   5880
                  TabIndex        =   86
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total Weight ( L )"
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
                  Height          =   255
                  Left            =   7920
                  TabIndex        =   85
                  Top             =   240
                  Width           =   1620
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Total Weight ( Kg )"
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
                  Height          =   255
                  Left            =   3720
                  TabIndex        =   84
                  Top             =   240
                  Width           =   1785
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Percentage Check"
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
                  Height          =   255
                  Left            =   240
                  TabIndex        =   83
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.Label lbPerc 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "."
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
                  Left            =   2760
                  TabIndex        =   82
                  Top             =   240
                  Width           =   60
               End
               Begin VB.Shape Shape1 
                  BorderColor     =   &H00E0E0E0&
                  Height          =   735
                  Left            =   -1320
                  Top             =   -2040
                  Width           =   11295
               End
            End
            Begin VB.Frame frChemicals 
               BackColor       =   &H00886010&
               BorderStyle     =   0  'None
               Height          =   1335
               Left            =   5880
               TabIndex        =   51
               Top             =   1800
               Visible         =   0   'False
               Width           =   5055
               Begin VB.Label lbChem 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Empty List..."
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   53
                  Top             =   360
                  Width           =   4995
               End
               Begin VB.Label lbChem 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Click + to add chemicals"
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
                  Index           =   1
                  Left            =   45
                  TabIndex        =   52
                  Top             =   720
                  Width           =   5010
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
               Index           =   3
               Left            =   13920
               TabIndex        =   49
               Top             =   4560
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Delete Table"
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
                  Index           =   3
                  Left            =   0
                  TabIndex        =   50
                  Top             =   120
                  Width           =   3015
               End
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
               Index           =   0
               Left            =   0
               TabIndex        =   46
               Top             =   0
               Width           =   17055
               Begin VB.Label Label3 
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
                  TabIndex        =   48
                  Top             =   120
                  Width           =   1725
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe : "
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00606060&
                  Height          =   345
                  Index           =   1
                  Left            =   0
                  TabIndex        =   47
                  Top             =   75
                  Width           =   1200
               End
               Begin VB.Line Line4 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   0
                  X2              =   16920
                  Y1              =   480
                  Y2              =   480
               End
            End
            Begin FlexCell.Grid GridChemicals 
               Height          =   3615
               Left            =   0
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   600
               Width           =   16935
               _ExtentX        =   29871
               _ExtentY        =   6376
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
            Begin VB.Shape Shape2 
               BorderColor     =   &H00D0D0D0&
               Height          =   930
               Left            =   0
               Top             =   5880
               Width           =   16935
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Set Each Component Quantity in Recipe: click darker cells and set quantity"
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
               Index           =   0
               Left            =   5430
               TabIndex        =   88
               Top             =   7080
               Width           =   6105
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00D0D0D0&
               X1              =   0
               X2              =   16920
               Y1              =   4320
               Y2              =   4320
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   1
               Left            =   9120
               Picture         =   "FormCode.frx":8DF4
               ToolTipText     =   "4000"
               Top             =   4605
               Width           =   240
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   0
               Left            =   7440
               Picture         =   "FormCode.frx":97F6
               ToolTipText     =   "4000"
               Top             =   4605
               Width           =   240
            End
            Begin VB.Label lbFunction 
               BackStyle       =   0  'Transparent
               Height          =   855
               Index           =   0
               Left            =   6840
               TabIndex        =   58
               Top             =   4440
               Width           =   1575
            End
            Begin VB.Label lbFunction 
               BackStyle       =   0  'Transparent
               Height          =   855
               Index           =   1
               Left            =   8400
               TabIndex        =   57
               Top             =   4440
               Width           =   1815
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Add Comp."
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
               TabIndex        =   56
               Top             =   4875
               Width           =   1155
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Delete Comp."
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
               TabIndex        =   55
               Top             =   4875
               Width           =   1380
            End
         End
         Begin VB.Line Line2 
            Visible         =   0   'False
            X1              =   9600
            X2              =   9600
            Y1              =   120
            Y2              =   39600
         End
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
      TabIndex        =   118
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
         TabIndex        =   119
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
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   675
      Begin VB.Label lblHoverClick 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00808080&
         Height          =   570
         Left            =   60
         TabIndex        =   79
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
         TabIndex        =   78
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
      Begin VB.Frame frClassification 
         BackColor       =   &H00000080&
         BorderStyle     =   0  'None
         Caption         =   "Image14"
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
         Left            =   11160
         TabIndex        =   153
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
         Begin VB.Label lbClassification 
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
            Left            =   0
            TabIndex        =   154
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.Frame frCritical 
         BackColor       =   &H00886010&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         TabIndex        =   116
         Top             =   120
         Visible         =   0   'False
         Width           =   8415
         Begin VB.ComboBox cmbLineRecipe 
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
            ForeColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   151
            Top             =   120
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00886010&
            Caption         =   "Only Recipes with Critical RM"
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
            Height          =   285
            Left            =   120
            TabIndex        =   117
            Top             =   240
            Width           =   3735
         End
         Begin VB.Image InExport 
            Height          =   480
            Left            =   3960
            MouseIcon       =   "FormCode.frx":A1F8
            MousePointer    =   99  'Custom
            Picture         =   "FormCode.frx":A502
            Top             =   120
            Visible         =   0   'False
            Width           =   480
         End
      End
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
         TabIndex        =   111
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
         TabIndex        =   110
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
         TabIndex        =   109
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
         Picture         =   "FormCode.frx":D8E4
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   3
         Left            =   15600
         MousePointer    =   99  'Custom
         Picture         =   "FormCode.frx":10CC6
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   4
         Left            =   18240
         Picture         =   "FormCode.frx":140A8
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
         Index           =   6
         Left            =   11760
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   107
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Export Recipe to RM"
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
            Index           =   6
            Left            =   0
            MousePointer    =   99  'Custom
            TabIndex        =   108
            Top             =   640
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   6
            Left            =   720
            MousePointer    =   99  'Custom
            OLEDropMode     =   1  'Manual
            Picture         =   "FormCode.frx":1748A
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   5
         Left            =   9600
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   105
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   5
            Left            =   720
            MousePointer    =   99  'Custom
            OLEDropMode     =   1  'Manual
            Picture         =   "FormCode.frx":19E7C
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Import Mixes from RM"
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
            Index           =   5
            Left            =   0
            MousePointer    =   99  'Custom
            TabIndex        =   106
            Top             =   640
            Width           =   1890
         End
      End
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
            Picture         =   "FormCode.frx":1C86E
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
            Picture         =   "FormCode.frx":1FC50
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
            Picture         =   "FormCode.frx":23032
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
            Picture         =   "FormCode.frx":26414
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
            Picture         =   "FormCode.frx":297F6
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
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   13740
         TabIndex        =   10
         Top             =   240
         Width           =   5190
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


Private uRecipe As RecipeType

Private IndexVisibleFrame As Integer

Private SelectedCode As String
Private UmComponent As String

Private RecipeCode As String
Private bSetPercentageLastComponent As Boolean
Private bFlagOpenRecipe As Boolean

Private RevisionID As Boolean
Private bCloneRecipe As Boolean



Private Sub Check1_Click()
Dim rc As Boolean
rc = IIf(Check1.Value = 1, True, False)
Check1.ForeColor = IIf(rc, &H70B070, vbWhite)
Check1.FontBold = rc
InExport.Visible = rc
If DatabaseIndex = 4 Then cmbLineRecipe.Visible = Not (rc)
DoEvents
CopyDatabaseGrd1

End Sub

Private Sub cmbLine_Click()
If lRow > 0 Then
    Grd2.Cell(lRow, 2).Text = cmbLine
    Grd2.Cell(lRow, 2).Alignment = cellCenterCenter
    
End If
End Sub



Private Sub cmbLineHannaCode_Click()
GlobalSearch
End Sub

Private Sub cmbLineRecipe_Click()

    Call SerarchRecipePerLine(Grd1, cmbLineRecipe)
    

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

Private Sub cmbUM2_Click()
Dim i As Integer
'Dim Perc As String
    uRecipe.UmMultiple = cmbUM2
    
    With GridChemicals
        For i = 1 To .Rows - 1
            .Cell(i, 5).Text = cmbUM2
            
           ' Perc = .Cell(i, 6).Text
            
            'If Perc <> "" Then
                'Call SetComponentWeightByPerc(Perc, i)
          '  End If
            
             
        Next
    End With
    
    Call CheckTotalsAndPercentage
    
End Sub





Private Sub Combo1_Click()
    Select Case DatabaseIndex
        
        Case 4
            ' Recipe/Formulation...
           SearchHannaCodeFormRecipe
    
    End Select
End Sub


Private Function SearchHannaCodeFormRecipe()


If Combo1 = "Hanna Code" Then
    
    Dim strRecipe As String
    Dim strHannaCode As String
    If FormCodes.DoShow(strHannaCode, , , , strRecipe) Then
        Text1(0) = strRecipe
        
        Combo1.ListIndex = 0
        
        RiempiGrid Grd1, Text1(0)
        PopupMessage 3, "Selected Hanna Code : " & strHannaCode, , , strRecipe
        
    End If
    
End If


End Function


Private Sub Text1_Click(Index As Integer)
 Select Case DatabaseIndex
        
        Case 4
            ' Recipe/Formulation...
           SearchHannaCodeFormRecipe
    
    End Select
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

    Grid2.ExportToExcel USER_DESKTOP & "\" & FormatNomeFile(RecipeCode) & "_RevHistory.xls", True, True
    MessageInfoTime = 2500
    PopupMessage 2, "File correcly created on Desktop", , , FormatNomeFile(RecipeCode) & "_RevHistory.xls"
End Sub


Private Sub Grd2_ButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim Value As String
Value = Grd2.Cell(Row, Col).Text
Select Case Row
    Case 3
        lRow = Row
        If Value <> "" Then cmbLine = Value
        
        PicLine.Visible = True
    Case 10, 13, 15
        lRow = Row
        lbUM = Grd2.Cell(Row, 1).Text
         If Value <> "" Then cmbUM = Value
        PicUm.Visible = True
    Case 14
    
        Call SetMinQtyMultiple(Grd2)
                       

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

Private Sub Image6_Click()
UmComponent = cmbUM2
PicUmComponent.Visible = False
End Sub

Private Sub InExport_Click()
If PBContainerViewport.Visible Then Exit Sub
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

Private Sub lbClassification_Click()
frClassification_Click
End Sub

Private Sub lbExcel_Click()
frExcel_Click
End Sub

Private Sub lbInside_Click(Index As Integer)
    Select Case Index
        
        Case 5
            ' rev history table
            If RecipeCode <> "" Then Call GetRecipeRevision(Grid2, RecipeCode)
    
    End Select
End Sub

Private Sub lbSetPercentageLastComponent_Click()
SetPercentageLastComponent_Click
End Sub

Private Sub lbUM2_Click()
Image6_Click
End Sub

Private Sub PicUmComponent_Click()
Image6_Click
End Sub



Private Sub SetPercentageLastComponent_Click()


bSetPercentageLastComponent = True

Call SetComponentWeightByPerc("", 0)

End Sub





Private Sub txRevision_Change(Index As Integer)
Dim rc As Boolean
rc = IIf(Len(txRevision(Index)) > 0, True, False)
txRevision(Index).BackColor = IIf(rc, vbWhite, &HE0E0E0)
End Sub


Private Sub ucScrollAdd1_ScrollH(Value As Long)
    Form_Resize
End Sub
Private Sub PicHover_Click()
PBContainer.Top = 0
End Sub
Private Sub lblHoverClick_Click()
    PBContainer.Top = 0
    
End Sub
Private Sub imOver_Click()
PBContainer.Top = 0
End Sub

'Poorly made resizing functions just for the example
Private Sub RSRight(c As Control, Source As Object, adjust As Long, Optional LimitLeft& = -1, Optional LimitRight& = -1)
On Error Resume Next
Dim aux&
    aux& = (Source.ScaleWidth - c.Width) + adjust
    If (err.NUMBER > 0) Then aux& = (Source.Width - c.Width) + adjust
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
    If (err.NUMBER > 0) Then aux& = (Source.Height - c.Height) + adjust
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

Private Sub chMix_Click()
Dim rc As Boolean
rc = IIf(chMix.Value = 1, True, False)
chMix.FontBold = rc
CopyDatabaseGrd1
End Sub

Private Sub Form_Load_Scroll()

    PBContainer.Top = 0
    PBContainer.Left = 0
    PBContainer.Width = Me.Width
    IndexVisibleFrame = 2
   
    
    Dim i As Integer
    If Screen.Width - Me.Width > 1000 And bFullScreen Then
        Me.WindowState = 2
    
    End If



        PBContainerViewport.Move 0, PBTitle.Height, Me.ScaleWidth, Me.ScaleHeight - PBTitle.Height

  
    RSBottom PicHover, Me, -1350
    RSRight PicHover, Me, -450
   

    PBContainerViewport.ZOrder
    PBFooter.ZOrder
    
    
    
End Sub
Private Sub Form_Load()
Dim i As Integer


        PBContainerViewport.Left = 0
        PBContainerViewport.Top = PBTitle.Height
        PBContainerViewport.Width = Me.ScaleWidth
        PBContainerViewport.Height = PBFooter.Top
        
        
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
    
    
    Label11(0) = "File Name : " & dbCodeName
    Label11(1) = "Actual Rel. " & dbCodeRelease & " ( " & dbCodeDate & " - " & dbCodeOperator & ")"
    Label11(0).Visible = True
    Label11(1).Visible = True
    
    DoEvents
    
    


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
    MsgBox "DOSHOW ERROR  " & err.Description
    Resume Next
End Function




Private Sub frChemicals_Click()
ImCode_Click 0
End Sub

Private Sub frCommandInside_Click(Index As Integer)
Select Case Index

    Case 0
        ' save
        Call SaveRecipe
    Case 1
        PBContainerViewport.Visible = False
        Lab(7) = "Exit Database Table"
        blTable = DatabaseString
    Case 2
    
        Call SetComponentComponentRecipe

        
    Case 3
        If F_MsgBox.DoShow("Delete all Components?", uRecipe.Code) Then
            If DeleteRecipeComponentByCode(uRecipe.Code) Then
                CheckChemicalsPerRecipe
                PopupMessage 2, "Records deleted...", , , uRecipe.Code
                
            End If
            
        End If
    Case 4
        If MyID > 0 Then Call F_PICTOGRAM.DoShow(MyID, 1)
        
        
    Case 5
    
        'If bFlagOpenRecipe = False Then
      '
      '       Call Form_Load_Scroll
      '       DoEvents
      '
      '       bFlagOpenRecipe = True
      '
      '  End If
       'lbInside(1) = Grd2.Cell(1, 2).Text & "  :  " & Grd2.Cell(2, 2).Text & "  |  Recipe Components"

        
      '  PBContainerViewport.Visible = True
      '  Lab(7) = "Exit Formulation"
      '  blTable = "Recipe Formulation"
      '  PBContainer.Top = 0
      
      If uRecipe.Code = "" Then
        
        uRecipe.Code = Grd2.Cell(1, 2).Text
      
        Call SetMyRecipeByCode(uRecipe.Code, uRecipe)
      End If
      
       zRecipe = uRecipe
       FormFormulation.Top = Me.Top
       FormFormulation.Left = Me.Left
       FormFormulation.Width = Me.Width
       FormFormulation.Height = Me.Height
       FormFormulation.DoShow
        
    Case 6
        frInside(5).Visible = False
         PBContainer.Top = 0
       
    Case 7
        Call F_PICTOGRAM.DoShow(MyID, 0, Grd2.Cell(1, 2).Text, True)
        
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


uRecipe = RecipeClean
If FirstRow > 0 Then
    
    
    Select Case DatabaseIndex
        Case 0
            ' hanna code
            
                
                MyID = Grd1.Cell(FirstRow, 3).Text
                Call SetGridEditCode(Grd2)
                Call CopyCodeGrd2(Grd2, MyID)
                
        Case 1
            ' Production Way
        
                
                MyID = Grd1.Cell(FirstRow, 2).Text
                Call SetGridEditProduction(Grd2)
                Call CopyProductionGrd2(Grd2, MyID)

            
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
            ' Recipe
                
                uRecipe = RecipeClean
                RecipeCode = Trim(Grd1.Cell(FirstRow, 1).Text)
                MyID = Trim(Grd1.Cell(FirstRow, 3).Text)
                Call SetGridEditRecipe(Grd2)
                Call CopyRecipeGrid2(Grd2, MyID)
                Call GetRecipeRevision(Grid2, RecipeCode)
               
                
                Call AddcmbRevType
                ' chemicals x recipe
                Call SetMyRecipeByCode(RecipeCode, uRecipe)
                Call CheckChemicalsPerRecipe
                
                PicMenu(6).Visible = True
                
        Case 5
            ' Chemical RM
                MyID = Grd1.Cell(FirstRow, 3).Text
                RecipeCode = Trim(Grd1.Cell(FirstRow, 1).Text)
                Call SetGridEditChemicalRM(Grd2)
                Call CopyChemicalRMGrd2(Grd2, MyID)
                frCommandInside(5).Visible = Grd2.Cell(12, 2).Text
                
              
                
    End Select
End If
End Sub

Private Sub CheckChemicalsPerRecipe()
                
    Call GetChemicalsPerRecipe(GridChemicals, uRecipe)
    Call CheckTotalsAndPercentage
    
   ' frCommandInside(5).Visible = True
    
    frChemicals.Visible = IIf(GridChemicals.Rows > 1, False, True)
    
    Call GetHannaCodePerRecipe(Grid1, uRecipe)
    Grid1.SelectionMode = cellSelectionByRow
                
End Sub


Private Sub grd1_Click()
ctlCalendar1.Visible = False

bCloneRecipe = False
 
End Sub





Private Sub Grd2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

Dim Value As String
lRow = 0

PicLine.Left = Grd2.Width / 2 - PicLine.Width / 2
PicLine.Top = Grd2.Height / 2 - PicLine.Height / 2

PicUm.Left = Grd2.Width / 2 - PicUm.Width / 2
PicUm.Top = Grd2.Height / 2 - PicUm.Height / 2

PicUm.Visible = False
PicLine.Visible = False

MessageInfoTime = 2000
    Select Case DatabaseIndex
        Case 0
            ' hanna code
            


        Case 1
            ' Production Way
        

            
        Case 2
            ' Code Classification
          
                Select Case FirstCol
                    Case 1
                       ' Dim MyID As Long
                        
                        If FirstRow = 7 Then
                           ' MyID = GetIDRowMaterial(Grd2.Cell(1, 2).Text)
                           ' Call F_PICTOGRAM.DoShow(0, 0, Grd2.Cell(1, 2).Text, False)
                        
                        End If
                
                End Select
                
                
        Case 3
            ' Frasi H

        Case 4
            ' Recipe
            Value = Grd2.Cell(FirstRow, 2).Text

                Select Case FirstRow
                    Case 3
                        lRow = FirstRow
                        If Value <> "" Then cmbLine = Value
                        
                        PicLine.Visible = True
                    Case 10, 13, 15
                        lRow = FirstRow
                        lbUM = Grd2.Cell(FirstRow, 1).Text
                         If Value <> "" Then cmbUM = Value
                        PicUm.Visible = True
                        
                        Call SetMinQtyMultiple(Grd2)
                    Case 9, 11
                        Call SetMinQtyMultiple(Grd2)
                    Case 14
                    
                        Call SetMinQtyMultiple(Grd2)
                       
                    
                        PopupMessage 2, "Minimum Quantity : Enter Min Q.ty and Multiple..."
                    Case 16
                        PopupMessage 2, "Select Recipe form list to Add or Delete Component/Mix in Recipe..."
                    Case 18
                       ' PopupMessage 2, "Please enter a valid Date...", , , "Revision Date"
                End Select

        Case 5
            ' Chemical RM
            Value = Grd2.Cell(FirstRow, 2).Text
            
            
                Select Case FirstRow
                    Case 3
                        'lRow = FirstRow
                        'If Value <> "" Then cmbLine = Value
                        
                       ' PicLine.Visible = True
                    Case 7
                        lRow = FirstRow
                        lbUM = Grd2.Cell(FirstRow, 1).Text
                         If Value <> "" Then cmbUM = Value
                        PicUm.Visible = True
                    Case 14
                       ' PopupMessage 2, "Minimum Quantity : Enter Min Q.ty and Multiple..."
                    Case 16
                        'PopupMessage 2, "Open RecipeForProduction to Add or Delete Component/Mix in Recipe..."

                End Select
                
                Select Case FirstCol
                    Case 1
                        Dim MyID As Long
                        
                        If FirstRow = 5 Then
                            MyID = GetIDRowMaterial(Grd2.Cell(1, 2).Text)
                            Call F_PICTOGRAM.DoShow(MyID, 1, , False)
                        
                        End If
                
                End Select

            
        
    
    End Select


End Sub



Private Sub GridChemicals_Click()

Dim sString As String
Dim Perc As String

If lRow > 0 And lCol = 6 Then
    
    
    sString = GridChemicals.Cell(0, lCol).Text
    Perc = GridChemicals.Cell(lRow, lCol).Text
    
    
    Call SetPercentageNumber(sString, SelectedCode, Perc, lRow)

End If


End Sub


Private Function SetPercentageNumber(ByVal sString As String, ByVal SelectedCode As String, ByVal Perc As String, ByVal FirstRow As Long)

If F_InputBox.DoShow(sString, SelectedCode, , , , Perc, , True, Me.Top) Then
    Call SetComponentWeightByPerc(Perc, FirstRow)
    
End If

End Function
Private Sub GridChemicals_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
Dim Qty As String
Dim Perc As String
Dim sString As String
Dim sNote As String
Dim Value As String
SelectedCode = GridChemicals.Cell(FirstRow, 1).Text
Qty = GridChemicals.Cell(FirstRow, FirstCol).Text
sString = GridChemicals.Cell(0, FirstCol).Text
sNote = GridChemicals.Cell(FirstRow, FirstCol).Text
Perc = GridChemicals.Cell(FirstRow, 6).Text

lRow = 0
frCommandInside(2).Visible = False

If FirstRow > 0 Then
frCommandInside(2).Visible = IIf(GridChemicals.Cell(FirstRow, 9).Text, True, False)
lbCommandInside(2) = SelectedCode & " - Component Table"
lCol = FirstCol
lRow = FirstRow

    Select Case FirstCol
        Case 4
            ' Q.ty to produce
            
            If F_InputBox.DoShow(sString, SelectedCode, , , , Qty, , True, Me.Top) Then
            
                
                If IsNumeric(Qty) Or Qty = "" Then
                    GridChemicals.Cell(FirstRow, 4).Text = Qty
                    GridChemicals.Cell(FirstRow, 4).Alignment = cellCenterCenter
                    Call CheckTotalsAndPercentage
                End If
                
            End If
        Case 5
          ' Value = GridChemicals.Cell(FirstRow, 5).Text
           ' lRow = FirstRow
            
          '  lbUM2 = lbInside(1) ' GridChemicals.Cell(FirstRow, 1).Text
             
           '  If Value <> "" Then cmbUM2 = Value
             
           '  PicUmComponent.Left = GridChemicals.Width / 2 - PicUmComponent.Width / 2
           '  PicUmComponent.Top = GridChemicals.Height / 2 - PicUmComponent.Height / 2
           '  PicUmComponent.Visible = True

        Case 6  ' perc
            
           Call SetPercentageNumber(sString, SelectedCode, Perc, lRow)
        Case 7
              ' note
            Perc = GridChemicals.Cell(FirstRow, 7).Text
            If F_InputBox.DoShow(sString, SelectedCode, , , , Perc, , True, Me.Top) Then
                GridChemicals.Cell(FirstRow, 7).Text = Perc
                GridChemicals.Cell(FirstRow, 7).Alignment = cellCenterCenter
            End If
            
    
         Case 8
            ' note
            If F_InputBox.DoShow(sString, SelectedCode, , , , sNote, , , Me.Top) Then
                GridChemicals.Cell(FirstRow, 8).Text = sNote
                GridChemicals.Cell(FirstRow, 8).Alignment = cellLeftCenter
            End If
    
    End Select

    

End If
End Sub


Private Sub Image1_Click()
PicUm_Click
End Sub

Private Sub Image2_Click()
PicLine_Click
End Sub

Private Sub Image3_Click(Index As Integer)

    If CheckPrivilege(3) = False Then Exit Sub

    Select Case Index
        Case 0
            ' pulisci maschera
            If PBContainerViewport.Visible Then Exit Sub
            Call CleanCode
            Grd2.Cell(1, 2).SetFocus
        Case 1
            If PBContainerViewport.Visible Then
                frCommandInside_Click 0
            Else
            
                If CheckPrivilege(3) Then
                    If SaveRecord Then
                        Select Case DatabaseIndex
                            Case 4 ' recipe
                                If uRecipe.Code = "" Then
                                    uRecipe.Code = Trim(Grd2.Cell(1, 2).Text)
                                End If
                                Call SetMyRecipeByCode(uRecipe.Code, uRecipe)
                                Call CheckChemicalsPerRecipe
                                Call SaveRecipeDataPerRevision(uRecipe)
                
                            Case 5 ' rawmaterials
                                frCommandInside(4).Visible = True
                        
                        End Select
                        CopyDatabaseGrd1
                    End If
                End If
            End If
        Case 2
            ' refresh table
            If PBContainerViewport.Visible Then Exit Sub
            Text1(0) = ""
            CopyDatabaseGrd1
        Case 3
            If PBContainerViewport.Visible Then Exit Sub
            If CheckPrivilege(3) Then CancellaTab
        Case 4
            If PBContainerViewport.Visible Then Exit Sub
            If CheckPrivilege(3) Then
                If F_MsgBox.DoShow("Export " & DatabaseString & " to Excel?") Then
                
                    Call DBCodeToExcel(ProgressBar1, DatabaseIndex, DatabaseString)
                End If
            End If
        Case 5
            If PBContainerViewport.Visible Then Exit Sub
             If ImportMixFromChemicalRM Then CopyDatabaseGrd1
             
        Case 6
            If PBContainerViewport.Visible Then Exit Sub
            If F_MsgBox.DoShow("Copy Recipe " & Grd2.Cell(1, 2).Text & " to Chemicals RM Mixes ? ") Then
                ExportRecipeToChemicalsRM
            End If
    
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
        If PBContainerViewport.Visible Then
            PBContainerViewport.Visible = False
            Lab(7) = "Exit Database Table"
            blTable = DatabaseString

        ElseIf frRevisionHistory.Visible Then
            
            frRevisionHistory.Visible = False
            frCritical.Visible = True
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
        
        If PBContainerViewport.Visible Then
        
            ' Previous
            If IndexVisibleFrame >= 2 Then
                MyIndex = IndexVisibleFrame - 1
                If frInside(MyIndex).Visible = False Then
                    MyIndex = IndexVisibleFrame - 3
                End If
                
                PBContainer.Top = -(frInside(MyIndex).Top - 480)
            Else
                PBContainer.Top = 0
            End If
    
    
    
        Else
            ' avanti di 10
            Call ScorriTabella(False)
        End If
    Case 4
        
        If PBContainerViewport.Visible Then
            
             ' forward
            If IndexVisibleFrame < frInside.UBound Then
                MyIndex = IndexVisibleFrame + 1
                If frInside(MyIndex).Visible = False Then
                    MyIndex = IndexVisibleFrame + 3
                End If
                PBContainer.Top = -(frInside(MyIndex).Top - 480)
            Else
                PBContainer.Top = 0
            End If
        
        Else
            ' indietro di 10
            Call ScorriTabella(True)
        End If
    
    
    
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
        DefaultMenu_Click 0
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
    PopupMessage 2, err.Description
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
Dim UserChCode As String
Dim UserHannaCode As String
Dim MyCodeID As Long


    frCommandInside(2).Visible = False
    

    Select Case Index
        Case 0
            'add
            If FormChemicalRM.DoShow(UserChCode) Then
                If UserChCode <> "" Then
                
                    frChemicals.Visible = False
                    
                    Call CopyUserChCodeInGrid(GridChemicals, UserChCode, uRecipe)
                    Call CheckTotalsAndPercentage
                
                End If
                
            End If
        Case 1
            ' delete
            If lRow > 0 Then
                If F_MsgBox.DoShow("Delete Component " & GridChemicals.Cell(lRow, 1).Text & " ? ", "Recipe : " & uRecipe.Code) Then
                    GridChemicals.ReadOnly = False
                    GridChemicals.Selection.DeleteByRow
                    GridChemicals.ReadOnly = True
                    Call CheckTotalsAndPercentage
                End If
            End If
            
        Case 2
            ' add hanna code per recipe
               If FormCodes.DoShow(UserHannaCode) Then
                If UserHannaCode <> "" Then
                
                    frHannaCode.Visible = False
                    
                    Call CopyUserHannaCodeInGrid(Grid1, UserHannaCode, uRecipe)
                
                End If
                
            End If
            
        Case 3
            ' delete hanna code per recipe
            If lRow > 0 Then
                If F_MsgBox.DoShow("Delete Recipe in Hanna Code " & Grid1.Cell(lRow, 1).Text & " ? ", "Recipe : " & uRecipe.Code) Then
                    Call DeleteRecipePerCode(Grid1.Cell(lRow, 1).Text, uRecipe.Code)
                    Call GetHannaCodePerRecipe(Grid1, uRecipe)
                    'Grid1.ReadOnly = False
                    'Grid1.Selection.DeleteByRow
                    'Grid1.ReadOnly = True
                    
                End If
            End If
                        
                        
        Case 4
            ' aggiungi revision specifics
            If AddRevision(RecipeCode, txRevision(2)) Then
                 Call GetRecipeRevision(Grid2, RecipeCode)
            End If
            
            frExcel.Visible = IIf(Grid2.Rows > 1, True, False)
            Frame6.Visible = IIf(Grid2.Rows > 1, False, True)
        Case 5
            ' delete revision specifics
            If DeleteRevision(RecipeCode, txRevision(2)) Then
                 Call GetRecipeRevision(Grid2, RecipeCode)
            End If
            
            frExcel.Visible = IIf(Grid2.Rows > 1, True, False)
            Frame6.Visible = IIf(Grid2.Rows > 1, False, True)
            
            
            
    End Select
End Sub

Private Sub imPerc_Click()
CheckTotalsAndPercentage
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
    Case 0
        Image6_Click
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

Private Sub PicLine_Click()
PicLine.Visible = False

End Sub


Private Sub PicUm_Click()
PicUm.Visible = False


End Sub


Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To PicMenu.Count - 1
    If i = Index Then
        PicMenu(i).BackColor = &H886010
    Else
        PicMenu(i).BackColor = &H644603
    End If
Next
End Sub

Private Sub PicMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
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
            ' Production Way
        
            Call Grd2_Production_LeaveCell(Grd2, Row, Col, NewRow, NewCol, Cancel, lRow)
            
        Case 2
            ' Code Classification
            Call Grd2_Classification_LeaveCell(Grd2, Row, Col, NewRow, NewCol, Cancel, lRow)
        Case 3
            ' Frasi H
            Call Grd2_FrasiH_LeaveCell(Grd2, Row, Col, NewRow, NewCol, Cancel, lRow)
        Case 4
            ' Recipe
            Call Grd2_Recipe_LeaveCell(Grd2, Row, Col, NewRow, NewCol, Cancel, lRow)
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
            ' Production Way
        
            rc = SaveDatabaseProduction(Grd2)
            
        Case 2
            ' Code Classification
            rc = SaveDatabaseClassification(Grd2)
        Case 3
            ' Frasi H
            rc = SaveDatabaseFrasiH(Grd2)
        Case 4
            ' Recipe
            Dim NewCode As String
            NewCode = Trim(UCase(Grd2.Cell(1, 2).Text))
            bCloneRecipe = IIf(NewCode <> Trim(UCase(RecipeCode)), True, False)
            
            rc = SaveDatabaseRecipe(Grd2, bCloneRecipe, RecipeCode)
            
            
        Case 5
            ' Chemical RM
           rc = SaveDatabaseChemicalRM(Grd2)
            
        
    
    End Select
    
   If rc Then
        bAddNewDatabaseRelease = True
   End If
    
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
Private Function CheckLottoAperto(ByVal sCode As String) As Boolean

Dim rc As Boolean
Dim Path As String
Dim FSO As New Scripting.FileSystemObject
Dim Cartella As Folder
Dim FileGenerico As file

    rc = False
    
     
    Path = USER_TEMP_PATH
    Set Cartella = FSO.GetFolder(Path)
    
        For Each FileGenerico In Cartella.Files
            If InStr(FileGenerico.Name, USER_ESTENSIONE) Then
                If InStr(FileGenerico.Name, FormatNomeFile(sCode)) Then
                   
                    rc = True
                End If
                
            End If
        Next

    
    CheckLottoAperto = rc

End Function




Private Sub SetDatabaseType(ByVal Index As Integer)
frCritical.Visible = False
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
            Call SetLine(cmbLineHannaCode, True)
            cmbLineHannaCode.Visible = True
            frClassification.Visible = True
            

        Case 1
            ' Production Way
            Set dbTab = dbTabProductionWay
            DatabaseString = "Production Way"
            
            Call SetGridProduction(Grd1)
            Call SetGridEditProduction(Grd2)
            Call AddComboProduction(Combo1)
            
       
            
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
            ' Recipe
            Check1.Caption = "Only Recipes with Critical RM"
            frCritical.Visible = True
            Set dbTab = dbTabRecipe
            DatabaseString = "Recipes"
            Call SetGridRecipe(Grd1)
            Call SetGridEditRecipe(Grd2)
            Call SetGridRecipeRevision(Grid2)
            
            Call AddComboRecipe(Combo1)
            
            Call SetChemicalsXRecipe
            
            Call SetLine(cmbLine)
            Call SetLine(cmbLineRecipe, True)
            
            cmbLineRecipe.Visible = True
            
            Call SetUM(cmbUM)
            Call SetUM(cmbUM2)
         
            frClassification.Visible = True


            PicMenu(5).Visible = True
           
        Case 5
            ' Chemical RM
            Check1.Caption = "Only Chemicals with Critical RM"
            frCritical.Visible = True
            
            Set dbTab = dbTabRawMaterial
            DatabaseString = "Chemical RM"
            
            Call SetGridChemicalRM(Grd1)
            Call SetGridEditChemicalRM(Grd2)
            Call AddComboChemicalRM(Combo1)
            
          '  Call SetChemicalsXRecipe
            
            chMix.Visible = True
            frCommandInside(4).Visible = True
            
             Call SetUM(cmbUM)
             
             frClassification.Visible = True
    
    End Select



    
    blTable = DatabaseString
    
    
    
    


End Sub

Private Sub SetChemicalsXRecipe()
    
    Call SetDatabaseComponentGrid(GridChemicals)
    Call SetCodeGrid(Grid1)
    
    Grid1.Column(6).Width = 0
    Grid1.ReadOnly = True

End Sub

Private Sub CopyDatabaseGrd1()

Dim rc As Boolean
rc = IIf(Check1.Value = 1, True, False)

Grd1.Rows = 1

    Select Case DatabaseIndex
        Case 0
            ' hanna code
            Call FillGridCode(Grd1, Text1(0), , Combo1, cmbLineHannaCode)

        Case 1
            ' Production Way

            Call CopyProductionGrd1(Grd1, Text1(0), , Combo1)
            
        Case 2
            ' Code Classification
            Call CopyClassificationGrd1(Grd1, Text1(0), , Combo1)
        Case 3
            ' Frasi H
            Call CopyFrasiHGrd1(Grd1, Text1(0), , Combo1)
        Case 4
            ' Recipe
            
            Call FillGridRecipe(Grd1, Text1(0), , Combo1, rc)
            
            cmbLineRecipe_Click
            
        Case 5
            ' Chemical RM
            Call CopyChemicalRMGrd1(Grd1, Text1(0), , Combo1, IIf(chMix.Value = 1, True, False), rc)
        
    
    End Select

    lbRecords = "Database Records # " & Grd1.Rows - 1
    
    If Grid1.Rows > 1 Then Grd1.Column(2).AutoFit

End Sub


Public Function SetUM(ByVal Cmb As ComboBox) As Boolean

    With Cmb
        .Clear
        .AddItem "L"
        .AddItem "kg"
        .AddItem "pcs"
        .AddItem "mL"
        .AddItem "g"
        .AddItem "mg"
        .ListIndex = 5
    End With
End Function

Public Function SetUMPeso(ByVal Cmb As ComboBox) As Boolean

    With Cmb
        .Clear
        .AddItem "kg"
        .AddItem "g"
        .AddItem "mg"
        .ListIndex = 1
    End With
End Function



Private Sub CheckTotalsAndPercentage()
Dim rc As Boolean
Dim strPerc As String
Dim Totali As Double
Dim bUmMassa As Boolean


If Me.Visible = False Then Exit Sub
If uRecipe.Density = 0 Then uRecipe.Density = 1


Select Case uRecipe.bUmMassa
    Case True
    
        If SetbUmMassa(uRecipe.UmMultiple) Then
            Totali = uRecipe.Multiple
        Else
            Totali = (uRecipe.Multiple / uRecipe.Density)
        End If
        
        uRecipe.TotalWeightKg = Totali
        uRecipe.UmTotalWeightKg = uRecipe.UmMultiple
        uRecipe.UmTotalWeightL = SetUmVolume(uRecipe.UmMultiple)
        
        
        
        
        Label6 = "Total Weight ( " & (uRecipe.UmMultiple) & " )"
        Label7 = "Total Weight ( " & SetUmVolume(uRecipe.UmMultiple) & " )"
        

        

        
    Case False
        Totali = uRecipe.Multiple * uRecipe.Density
        uRecipe.TotalWeightKg = Totali
        
        uRecipe.UmTotalWeightKg = SetUmVolume(uRecipe.UmMultiple)
        uRecipe.UmTotalWeightL = uRecipe.UmMultiple
        
        Label6 = "Total Weight ( " & uRecipe.UmTotalWeightKg & " )"
        Label7 = "Total Weight ( " & uRecipe.UmMultiple & " )"
        
        
        
        
        
End Select

'Dim Perc As Double
Dim i As Integer
For i = 1 To GridChemicals.Rows - 1
    strPerc = GridChemicals.Cell(i, 6).Text
    If strPerc = "" Then Exit For
    Call SetComponentWeightByPerc(CDbl(strPerc), i)
Next

'rc = CheckPercentageByWeight(GridChemicals, strPerc, Totali * Um(uRecipe.UmTotalWeightKg))


'PicPerc.BackColor = IIf(rc, &H8000&, &H40C0&)
'CheckPercentage


  

Totali = FormatNumber(Totali, iVirgola(Totali))

    
lbTotalWKg = Totali
lbTotalWL = FormatNumber((Totali / uRecipe.Density), iVirgola(Totali))


End Sub

Private Function CalculateLastPercentage(ByRef Row As Long) As String
Dim rc As Boolean
Dim i As Integer
Dim TotalPerc As Double

    If lbPerc <> "" Then

        TotalPerc = CDbl(Replace(lbPerc, "%", ""))
    
    
        CalculateLastPercentage = ""
        Row = 0
        
        With GridChemicals
            If .Rows < 2 Then Exit Function
            For i = 1 To .Rows - 1
                If .Cell(i, 6).Text = "" Then

                    CalculateLastPercentage = FormatNumber(100 - TotalPerc, 4)
                    Row = i
                    Exit Function
                End If
            Next
            TotalPerc = 0
            For i = 1 To .Rows - 2
               
                    TotalPerc = TotalPerc + Replace(.Cell(i, 6).Text, "%", "")
                    CalculateLastPercentage = FormatNumber(100 - TotalPerc, 4)
                   
                
                
               
            Next
             Row = .Rows - 1
            
        End With
    End If
    
End Function

Private Function SetComponentWeightByPerc(ByVal Perc As String, ByVal Row As Long)

Dim TotalW As Double
Dim Qty As Double
Dim UmComponent As String

   
   If bSetPercentageLastComponent Then
        bSetPercentageLastComponent = False
        Perc = CalculateLastPercentage(Row)
        
     
    End If
    
    If Perc = "" Then
            
            GridChemicals.Cell(Row, 4).Text = ""
            
            GridChemicals.Cell(Row, 6).Text = ""
            Call CheckPercentage
        Exit Function
    End If
    
    

    With uRecipe
        
        UmComponent = GridChemicals.Cell(Row, 5).Text
        
        If UmComponent = "" Then
        
            UmComponent = .UmMultiple
            GridChemicals.Cell(Row, 5).Text = .UmMultiple
            
        End If
    
        TotalW = .TotalWeightKg * Um(.UmTotalWeightKg)
        
        If TotalW = 0 Then
            PopupMessage 2, "Warning : Please check Recipes Specification...", , True, .Code
            Exit Function
        End If
        
    Select Case uRecipe.bUmMassa
        Case True
            If SetbUmMassa(uRecipe.UmMultiple) Then
            Else
                TotalW = TotalW * uRecipe.Density
            End If
        
            
        Case False
        
            'If SetbUmMassa(uRecipe.UmMultiple) Then
            
           ' Else
                'TotalW = TotalW / uRecipe.Density
           ' End If
            
          
            
    End Select
        
        
        If IsNumeric(Perc) Or Perc <> "" Then
            Qty = FormatNumber((Perc * (TotalW / 100) / Um(UmComponent)), 3)
            GridChemicals.Cell(Row, 4).Text = Qty
            GridChemicals.Cell(Row, 4).Alignment = cellRightCenter
            
            GridChemicals.Cell(Row, 6).Text = FormatNumber(Perc, 4)
            GridChemicals.Cell(Row, 6).Alignment = cellCenterCenter
            
            Call CheckPercentage
        End If
        
    End With
    
    
    
End Function



Private Sub CheckPercentage()
Dim i As Integer
Dim rc As Boolean
Dim TotPerc As Double
Dim strPerc As String
strPerc = ""
rc = False
SetPercentageLastComponent.Visible = False
With GridChemicals
    If .Rows > 1 Then
        SetPercentageLastComponent.Visible = True
        For i = 1 To .Rows - 1
            If .Cell(i, 6).Text = "" Then
                rc = False
                Exit For
            End If
            TotPerc = TotPerc + CDbl(.Cell(i, 6).Text)
        Next
        rc = IIf((TotPerc) = 100, True, False)
        strPerc = TotPerc
    End If

End With

PicPerc.BackColor = IIf(rc, &H8000&, &H40C0&)

lbPerc = strPerc & " %"


End Sub



Private Sub SaveRecipe()
Dim rc As Boolean
If CheckChemicalsInRecipe(GridChemicals) Then

   
    rc = SaveFullRecipe(GridChemicals, uRecipe.Code)
    
    If rc Then
         PopupMessage 2, "Recipe correctly saved...", , , uRecipe.Code
         
         Call CopyRecipeGrid2(Grd2, MyID)
    Else
        GoTo err
    End If
    
Else
err:
    PopupMessage 2, "Check all Chemicals and Weights before saving...", , True, uRecipe.Code
End If
End Sub

Private Sub SetComponentComponentRecipe()

    Dim strRecipe As String
    
    
        strRecipe = SelectedCode
        
        lbInside(4) = "Components : " & strRecipe

        Call SetComponentGrid(Grid3)
        Grid3.Column(7).Width = 0
        Grid3.Column(8).Width = 0
        Call AddComponentGrid(Grid3, strRecipe)
        
        Frame4.Visible = IIf(Grid3.Rows > 1, False, True)
        
        frInside(5).Top = frInside(3).Top
        frInside(5).Left = frInside(3).Left
        frInside(5).Width = frInside(3).Width
        frInside(5).ZOrder
        frInside(5).Visible = True
        PBContainer.Top = -(frInside(5).Top - 480)
         
End Sub


Private Function ImportMixFromChemicalRM() As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim strCode As String
    rc = False
    With dbTabRawMaterial
        .filter = ""
        .filter = "bMix=true"
        If .EOF Then
            PopupMessage 2, "No Mixes in Chemicals RM Database...", , , "Recipe Database"
        Else
            .MoveFirst
            For i = 1 To .RecordCount
                
                strCode = IIf(IsNull(!Code), "", Trim(!Code))
                
                If CheckExistinCodeInTabRecipe(strCode) Then
                    t = t + 1
                    With dbTabRecipe
                        .AddNew
                        !Code = strCode
                        
                        !Description = IIf(IsNull(dbTabRawMaterial!Description), "", Trim(dbTabRawMaterial!Description))
                        
                        .Update
                    
                    End With
                
                End If
            
                .MoveNext
            Next
            If t > 0 Then
                PopupMessage 2, "N." & t & " Chemicals Mix imported..." & vbCrLf & "Goto Recipe Database to fill specifics", , , "Recipe Database"
            Else
                PopupMessage 2, "All Chemical RM Mixes are already visible in Recipes Database", , , "Recipe Database"
            End If
            rc = True
        End If
    End With
    
    ImportMixFromChemicalRM = rc
    
End Function
Private Function CheckExistinCodeInTabRecipe(ByVal strCode As String) As Boolean
Dim rc As Boolean
    With dbTabRecipe
        .filter = ""
        .filter = "Code='" & strCode & "'"
        rc = .EOF
    End With
    CheckExistinCodeInTabRecipe = rc

End Function


Private Function ExportRecipeToChemicalsRM() As Boolean

    Dim RecipeCode As String
    
    RecipeCode = Trim(Grd2.Cell(1, 2).Text)
    

    With dbTabRawMaterial
        .filter = ""
        .filter = "Code='" & RecipeCode & "'"
        If .EOF Then
            .AddNew
            !Code = RecipeCode
        End If
        
        !Description = Trim(Grd2.Cell(2, 2).Text)
        !bMix = True
        .Update
            
        PopupMessage 2, "New Chemical RM Mix :  " & RecipeCode & " copied correctly.." & vbCrLf & "Goto Chemical RM Database to fill specifics"

    
        
    End With
    
    
    
End Function


Private Sub OpenRevisionHistory()

lbInside(5) = RecipeCode & " : Revision History"

frExcel.Visible = IIf(Grid2.Rows > 1, True, False)
Frame6.Visible = IIf(Grid2.Rows > 1, False, True)

                
frRevisionHistory.BackColor = &HF0F0F0
frRevisionHistory.Move frInside(0).Left, frInside(0).Top, Me.Width - frInside(0).Left * 2, frInside(0).Height
frRevisionHistory.ZOrder
frRevisionHistory.Visible = True
frCritical.Visible = False

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

Private Sub cmbRevType_Click()
txRevision(1) = cmbRevType
cmbRevType.Visible = False
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
             If uRecipe.Rev <> "" Then
                If IsNumeric(uRecipe.Rev) Then
                    Answer = uRecipe.Rev
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

With dbTabRecipeRevisionHistory
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




With dbTabRecipeRevisionHistory
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


Private Sub frClassification_Click()
Dim Index As Integer
Select Case DatabaseIndex

    Case 0 'hanna code
        Index = 0
    Case 4 ' recipe
        Index = 2
    Case 5 'raw material
        Index = 1
         
End Select

Call OpenProductCalssification(Grd2.Cell(1, 2).Text, Index)

End Sub


