VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form FormDatabase 
   BackColor       =   &H00F0F0F0&
   Caption         =   "Chemical MR"
   ClientHeight    =   12000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19200
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   12000
   ScaleWidth      =   19200
   Begin VB.PictureBox PBFooter 
      BackColor       =   &H00886010&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   53
      Top             =   11040
      Width           =   19215
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   1
         Left            =   3960
         TabIndex        =   64
         Top             =   -240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   4
         Left            =   18240
         Picture         =   "FormDatabase.frx":0000
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   3
         Left            =   15600
         MousePointer    =   99  'Custom
         Picture         =   "FormDatabase.frx":33E2
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   0
         Left            =   9360
         MousePointer    =   99  'Custom
         Picture         =   "FormDatabase.frx":67C4
         Top             =   120
         Width           =   480
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
         TabIndex        =   63
         Top             =   120
         Width           =   4335
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
         TabIndex        =   62
         Top             =   360
         Visible         =   0   'False
         Width           =   6045
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
         TabIndex        =   61
         Top             =   600
         Visible         =   0   'False
         Width           =   6525
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
         TabIndex        =   60
         Top             =   660
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
         Left            =   15360
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   660
         Width           =   1230
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
         TabIndex        =   58
         Top             =   660
         Width           =   1620
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   3
         Left            =   14760
         TabIndex        =   57
         Top             =   -120
         Width           =   2175
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   0
         Left            =   8760
         TabIndex        =   56
         Top             =   -120
         Width           =   1695
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   4
         Left            =   17280
         TabIndex        =   55
         Top             =   -120
         Width           =   1935
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1215
         Index           =   2
         Left            =   0
         TabIndex        =   54
         Top             =   -120
         Visible         =   0   'False
         Width           =   1935
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
      Index           =   8
      Left            =   15480
      TabIndex        =   49
      Top             =   1440
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
         TabIndex        =   50
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Frame frRevisionHistory 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9495
      Left            =   17640
      TabIndex        =   20
      Top             =   840
      Visible         =   0   'False
      Width           =   19215
      Begin VB.Frame frInside 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "16800"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9015
         Index           =   6
         Left            =   1080
         TabIndex        =   21
         Top             =   240
         Width           =   17055
         Begin VB.Frame Frame6 
            BackColor       =   &H00886010&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   5880
            TabIndex        =   35
            Top             =   2400
            Width           =   5055
            Begin VB.Label lbChem 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Empty List..."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   37
               Top             =   360
               Width           =   4995
            End
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
               TabIndex        =   36
               Top             =   720
               Width           =   2340
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
            TabIndex        =   33
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
               TabIndex        =   34
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
            Index           =   4
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   17055
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hanna Code Revision Table"
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
               Left            =   14475
               TabIndex        =   32
               Top             =   120
               Width           =   2415
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
               TabIndex        =   31
               Top             =   75
               Width           =   13215
            End
            Begin VB.Line Line10 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   16920
               Y1              =   480
               Y2              =   480
            End
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
            TabIndex        =   29
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            Index           =   3
            Left            =   13560
            TabIndex        =   26
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
            Height          =   585
            Index           =   4
            Left            =   2160
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Top             =   6240
            Width           =   13815
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
            TabIndex        =   24
            Top             =   5760
            Visible         =   0   'False
            Width           =   2415
         End
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
            TabIndex        =   22
            Top             =   6960
            Width           =   3015
            Begin VB.Image Image 
               Height          =   480
               Left            =   120
               MousePointer    =   99  'Custom
               OLEDropMode     =   1  'Manual
               Picture         =   "FormDatabase.frx":9BA6
               Top             =   0
               Width           =   480
            End
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
               TabIndex        =   23
               Top             =   120
               Width           =   3015
            End
         End
         Begin FlexCell.Grid Grid2 
            Height          =   4695
            Left            =   0
            TabIndex        =   38
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
            Index           =   4
            Left            =   8400
            TabIndex        =   39
            Top             =   7440
            Width           =   1815
         End
         Begin VB.Label lbFunction 
            BackStyle       =   0  'Transparent
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
            Index           =   5
            Left            =   6840
            TabIndex        =   40
            Top             =   7440
            Width           =   1575
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
            TabIndex        =   48
            Top             =   8640
            Width           =   6435
         End
         Begin VB.Line Line11 
            BorderColor     =   &H00D0D0D0&
            X1              =   0
            X2              =   16920
            Y1              =   5400
            Y2              =   5400
         End
         Begin VB.Image ImCode 
            Height          =   240
            Index           =   4
            Left            =   9120
            Picture         =   "FormDatabase.frx":CF88
            ToolTipText     =   "4000"
            Top             =   7485
            Width           =   240
         End
         Begin VB.Image ImCode 
            Height          =   240
            Index           =   5
            Left            =   7440
            Picture         =   "FormDatabase.frx":D98A
            ToolTipText     =   "4000"
            Top             =   7485
            Width           =   240
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Save Specifics"
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   6960
            TabIndex        =   47
            Top             =   7875
            Width           =   1335
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delete Specifics"
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   8640
            TabIndex        =   46
            Top             =   7875
            Width           =   1500
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
            TabIndex        =   45
            Top             =   5760
            Width           =   1695
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
            TabIndex        =   44
            Top             =   5760
            Width           =   735
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
            TabIndex        =   43
            Top             =   5760
            Width           =   1215
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
            TabIndex        =   42
            Top             =   5760
            Width           =   855
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
            TabIndex        =   41
            Top             =   6240
            Width           =   1695
         End
      End
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00644603&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00644603&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   420
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   1560
      Width           =   2895
   End
   Begin VB.PictureBox PicMainMenu 
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
      Index           =   4
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   19215
      TabIndex        =   5
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   4
         Left            =   7680
         MouseIcon       =   "FormDatabase.frx":E38C
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   16
         Top             =   0
         Width           =   1935
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Export DB Code"
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
            MouseIcon       =   "FormDatabase.frx":E696
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   720
            Width           =   1890
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   4
            Left            =   720
            MouseIcon       =   "FormDatabase.frx":E9A0
            MousePointer    =   99  'Custom
            Picture         =   "FormDatabase.frx":ECAA
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   0
         Left            =   0
         MouseIcon       =   "FormDatabase.frx":1208C
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   12
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   720
            MouseIcon       =   "FormDatabase.frx":12396
            MousePointer    =   99  'Custom
            Picture         =   "FormDatabase.frx":126A0
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New Code"
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
            MouseIcon       =   "FormDatabase.frx":15A82
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   720
            Width           =   1830
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   1
         Left            =   1920
         MouseIcon       =   "FormDatabase.frx":15D8C
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   10
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   720
            MouseIcon       =   "FormDatabase.frx":16096
            MousePointer    =   99  'Custom
            Picture         =   "FormDatabase.frx":163A0
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Save Code"
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
            MouseIcon       =   "FormDatabase.frx":19782
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   720
            Width           =   1830
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   2
         Left            =   3840
         MouseIcon       =   "FormDatabase.frx":19A8C
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   8
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   2
            Left            =   720
            MouseIcon       =   "FormDatabase.frx":19D96
            MousePointer    =   99  'Custom
            Picture         =   "FormDatabase.frx":1A0A0
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Refresh Code Table"
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
            MouseIcon       =   "FormDatabase.frx":1D482
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   720
            Width           =   1890
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00644603&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   3
         Left            =   5760
         MouseIcon       =   "FormDatabase.frx":1D78C
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   1935
         TabIndex        =   6
         Top             =   0
         Width           =   1935
         Begin VB.Image Image3 
            Height          =   480
            Index           =   3
            Left            =   720
            MouseIcon       =   "FormDatabase.frx":1DA96
            MousePointer    =   99  'Custom
            Picture         =   "FormDatabase.frx":1DDA0
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Delete Code"
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
            MouseIcon       =   "FormDatabase.frx":21182
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   720
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
         Left            =   16200
         TabIndex        =   14
         Top             =   240
         Width           =   2550
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   6600
      Top             =   8520
   End
   Begin VB.Timer Timer3 
      Interval        =   250
      Left            =   6480
      Top             =   9600
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H008080FF&
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   0
      Left            =   5760
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6600
      Top             =   9120
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   0
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   19200
      _ExtentX        =   33867
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin FlexCell.Grid Grd1 
      Height          =   8655
      Left            =   240
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   2160
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   15266
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
   Begin FlexCell.Grid Grd2 
      Height          =   8775
      Left            =   6960
      TabIndex        =   52
      Top             =   2160
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   15478
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Specification"
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
      Index           =   1
      Left            =   6840
      TabIndex        =   15
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   7
      Left            =   14880
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   17280
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label DefaultMenuLabel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   9
      Left            =   3960
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "FormDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private bModMenu As Boolean
Private TimerCount As Integer
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
Private dIndexProcedura As Integer
Private m_Procedura As Boolean
Private CampioneSelezionato As String

Private bHilight As Boolean
Private DataIndex As Integer
Private m_rc As Boolean

Private MyID As Long
Private MyIndexRecord As Integer
Private lRow As Long
Private MRCode As String
Private RevisionID As Long
Private RevDate As String



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
            .Left = Ctl.x1
            .Top = Ctl.y1
            .Width = Ctl.x2 - Ctl.x1
            .Height = Ctl.y2 - Ctl.y1
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
            Ctl.x1 = x_scale * .Left
            Ctl.y1 = y_scale * .Top
            Ctl.x2 = Ctl.x1 + x_scale * .Width
            Ctl.y2 = Ctl.y1 + y_scale * .Height
        ElseIf TypeOf Ctl Is Timer Then
        ElseIf TypeOf Ctl Is Inet Then
        ElseIf TypeOf Ctl Is Grid Then
           Ctl.Left = x_scale * .Left
            Ctl.Top = y_scale * .Top
            Ctl.Width = x_scale * .Width
            Ctl.Height = y_scale * .Height

             
           
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

Private Sub Form_Initialize()
SaveSizes
End Sub

Private Sub Form_Load()
Dim i As Integer
If Screen.Width - Me.Width > 1000 And bFullScreen Then
    Me.WindowState = 2
    'Me.Picture = LoadPicture(PictureMaxScreen)
   
End If
Call SetGridCode(Grd1)
Call SetGridEditCode(Grd2)

Call GetLastImport(Me, 0)
Call AddCombo
    MyID = 0
    MyIndexRecord = 3
  

End Sub


Public Function DoShow() As Boolean
Dim i As Integer
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



Private Sub frCommandInside_Click(Index As Integer)
Select Case Index
    Case 8
        Call OpenRevisionHistory
    Case 9
        Call ClearRevisionForm
End Select
End Sub

Private Sub frExcel_Click()

Grid2.ExportToExcel USER_DESKTOP & "\" & FormatNomeFile(MRCode) & "_RevHistory.xls", True, True
DoEvents
MessageInfoTime = 2500
PopupMessage 2, "File correcly created on Desktop", , , FormatNomeFile(MRCode) & "_RevHistory.xls"

End Sub

Private Sub Grd1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

If FirstRow > 0 Then
    
    MyID = Grd1.Cell(FirstRow, 7).Text
    
    MRCode = Trim(Grd1.Cell(FirstRow, 1).Text)
    
    Call SetGridEditCode(Grd2)
    Call CopyCodeGrd2(MyID)
    
    If CheckLottoAperto(Grd1.Cell(FirstRow, 1).Text) Then
    
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

Private Sub Image3_Click(Index As Integer)

    
    
    Select Case Index
        Case 0
            ' pulisci maschera
            frCommandInside(8).Visible = False
            Call CleanCode
            Grd2.Cell(1, 2).SetFocus
        Case 1
            If CheckPrivilege(3) Then
                If SaveCode Then
                
                    GlobalSearch
                End If
            End If
        Case 2
            ' refresh table
            
            Text1(0) = ""
        Case 3
            If CheckPrivilege(3) Then CancellaTab
        Case 4
            If CheckPrivilege(3) Then
                If F_MsgBox.DoShow("Export  DB Hanna Code to Excel?") Then
                
                    Call dbChemicalMRToExcel(ProgressBar1)
                End If
            End If

    
    End Select
End Sub



Private Sub Form_Resize()
ResizeControls
frRevisionHistory.Move 0, ProgressBar1.Top, Me.Width, Me.Height - ProgressBar1.Top - 20
End Sub





Private Sub DefaultMenu_Click(Index As Integer)
Select Case Index
    Case 0
    
        If frRevisionHistory.Visible Then
            
            frRevisionHistory.Visible = False
            'frCritical.Visible = True
        Else
        
            'If bFlagOpenRecipe Then ucScrollAdd1.Terminate
            If F_MsgBox.DoShow("Exit Hanna Code database?", "Database") Then
                Unload Me
            End If
        End If
    
      
    Case 2
        ' Open Report folder
        OpenWithDefault (USER_DOCUMENTI & PathReport)
      
    Case 1
        ' filtro
        
       
        
    Case 4
        ' avanti di 10
        Call ScorriTabella(True)
    Case 3
        ' indietro di 10
        Call ScorriTabella(False)
    
    
    
    Case 5
   
    Case 6
      
    Case 7
     
    Case 8
        m_Procedura = True
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
        DefaultMenuLabel_Click 2
    Case 39
        DefaultMenuLabel_Click 0
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
    FillGridCode Grd, Code, , Combo1

ERR_END:
   
    IndexTabella = 1
    MaxIndex = IIf(Int((Grd2.Rows - 1) / 10) < (Grd2.Rows - 1) / 10, (Int((Grd2.Rows - 1) / 10)) + 1, Int((Grd2.Rows - 1) / 10))
    If MaxIndex = 0 Then MaxIndex = 1
    
        lbRecords = "Database Records # " & Grd1.Rows - 1

    Exit Function
ERR_GRID:
    MessageInfoTime = 2000
    Text1(0) = ""
    PopupMessage 2, err.Description
    GoTo ERR_END:
End Function

Private Sub CopyCodeGrd2(ByVal lId As Long)
If lId = 0 Then Exit Sub
Dim strUserDecimal As String

 With dbTabCode
        
    .Filter = ""
    .Filter = "ID='" & lId & "'"
    If .EOF Then Exit Sub
    .MoveFirst
    
    
End With
strUserDecimal = IIf(IsNull(Trim(dbTabCode!Decimal)), 0, Trim((dbTabCode!Decimal)))
If Not (IsNumeric(strUserDecimal)) Then strUserDecimal = 0
strUserDecimal = FormatDecimal(strUserDecimal)


With Grd2
   
    .Cell(1, 2).Text = IIf(IsNull(Trim(dbTabCode!Code)), "", Trim(dbTabCode!Code))
    .Cell(2, 2).Text = IIf(IsNull(Trim(dbTabCode!Description)), "", Trim(dbTabCode!Description))      '& "SFG Description"
    .Cell(3, 2).Text = IIf(IsNull(Trim(dbTabCode!Line)), "", Trim(dbTabCode!Line))      '& "Line"
    .Cell(4, 2).Text = IIf(IsNull(Trim(dbTabCode!Recipe)), "", Trim(dbTabCode!Recipe))      '& "Recipe"
    
    .Cell(5, 2).Text = IIf(IsNull(Trim(dbTabCode!QCMethod)), "", Trim(dbTabCode!QCMethod))      '& "QC Method"
    .Cell(6, 2).Text = IIf(IsNull(Trim(dbTabCode!MeterFamily1)), "", Trim(dbTabCode!MeterFamily1))      '& "Meter Family 1"
    .Cell(7, 2).Text = IIf(IsNull(Trim(dbTabCode!MeterFamily2)), "", Trim(dbTabCode!MeterFamily2))      '& "Meter Family 2"
    
    .Cell(8, 2).Text = IIf(IsNull(Trim(dbTabCode!ParameterMethod)), "", Trim(dbTabCode!ParameterMethod))      '& "Parameter Method"
    .Cell(9, 2).Text = IIf(IsNull(Trim(dbTabCode!ParameterFormula)), "", Trim(dbTabCode!ParameterFormula))      '& "Parameter Formula"
    .Cell(10, 2).Text = IIf(IsNull(Trim(dbTabCode!MeasurementUnit)), "", Trim(dbTabCode!MeasurementUnit))      '& "Measurement Unit"
    
    
    
    .Cell(12, 2).Text = IIf(IsNull(Trim(dbTabCode!RangeMin)), "", Trim(dbTabCode!RangeMin))      '& "Range Min"
    .Cell(13, 2).Text = IIf(IsNull(Trim(dbTabCode!RangeMax)), "", Trim(dbTabCode!RangeMax))      '& "Range Max"
    .Cell(14, 2).Text = IIf(IsNull(Trim(dbTabCode!Decimal)), "0", Trim((dbTabCode!Decimal)))      '& "Decimal"
    
    
    
    
    
    .Cell(16, 2).Text = IIf(IsNull(Trim(dbTabCode!Fixed)), "", Trim(dbTabCode!Fixed))      '& "Fixed"
    .Cell(17, 2).Text = IIf(IsNull(Trim(dbTabCode!AndOr)), "", Trim(dbTabCode!AndOr))      '& "And / Or"
    .Cell(18, 2).Text = IIf(IsNull(Trim(dbTabCode!Percentage)), "", Trim(dbTabCode!Percentage))      '& "Percentage %"
    .Cell(19, 2).Text = IIf(IsNull(Trim(dbTabCode!QCRestriction)), "", Trim(dbTabCode!QCRestriction))      '& "QC Restriction"
  '  .Cell(20, 2).Text = IIf(IsNull(Trim(dbTabCode!STDMR)), "", Trim(dbTabCode!STDMR))      '& "STDMR "
    
    .Cell(22, 2).Text = IIf(IsNull(Trim(dbTabCode!STD1Value)), "", Trim(dbTabCode!STD1Value))      '& "Value"
    .Cell(23, 2).Text = IIf(IsNull(Trim(dbTabCode!STD1Min)), "", Trim(dbTabCode!STD1Min))      '& "Min"
    .Cell(24, 2).Text = IIf(IsNull(Trim(dbTabCode!STD1Max)), "", Trim(dbTabCode!STD1Max))      '& "Max"
    
    
    .Cell(26, 2).Text = IIf(IsNull(Trim(dbTabCode!STD2Value)), "", Trim(dbTabCode!STD2Value))      '& "Value"
    .Cell(27, 2).Text = IIf(IsNull(Trim(dbTabCode!STD2Min)), "", Trim(dbTabCode!STD2Min))      '& "Min"
    .Cell(28, 2).Text = IIf(IsNull(Trim(dbTabCode!STD2Max)), "", Trim(dbTabCode!STD2Max))      '& "Max"
    
    
    .Cell(30, 2).Text = IIf(IsNull(Trim(dbTabCode!STD3Value)), "", Trim(dbTabCode!STD3Value))      '& "Value"
    .Cell(31, 2).Text = IIf(IsNull(Trim(dbTabCode!STD3Min)), "", Trim(dbTabCode!STD3Min))      '& "Min"
    .Cell(32, 2).Text = IIf(IsNull(Trim(dbTabCode!STD3Max)), "", Trim(dbTabCode!STD3Max))      '& "Max"
    
    
    .Cell(34, 2).Text = IIf(IsNull(Trim(dbTabCode!STD4Value)), "", Trim(dbTabCode!STD4Value))      '& "Value"
    .Cell(35, 2).Text = IIf(IsNull(Trim(dbTabCode!STD4Min)), "", Trim(dbTabCode!STD4Min))      '& "Min"
    .Cell(36, 2).Text = IIf(IsNull(Trim(dbTabCode!STD4Max)), "", Trim(dbTabCode!STD4Max))      '& "Max"
    
    .Cell(38, 2).Text = IIf(IsNull(Trim(dbTabCode!STD5Value)), "", Trim(dbTabCode!STD5Value))      '& "Value"
    .Cell(39, 2).Text = IIf(IsNull(Trim(dbTabCode!STD5Min)), "", Trim(dbTabCode!STD5Min))      '& "Min"
    .Cell(40, 2).Text = IIf(IsNull(Trim(dbTabCode!STD5Max)), "", Trim(dbTabCode!STD5Max))      '& "Max"
    
    
    .Cell(42, 2).Text = IIf(IsNull(Trim(dbTabCode!STD6Value)), "", Trim(dbTabCode!STD6Value))      '& "Value"
    .Cell(43, 2).Text = IIf(IsNull(Trim(dbTabCode!STD6Min)), "", Trim(dbTabCode!STD6Min))      '& "Min"
    .Cell(44, 2).Text = IIf(IsNull(Trim(dbTabCode!STD6Max)), "", Trim(dbTabCode!STD6Max))      '& "Max"
            
    
    .Cell(46, 2).Text = IIf(IsNull(Trim(dbTabCode!pH1Value)), "", Trim(dbTabCode!pH1Value))      '& "Value"
    .Cell(47, 2).Text = IIf(IsNull(Trim(dbTabCode!pH1Min)), "", Trim(dbTabCode!pH1Min))      '& "Min"
    .Cell(48, 2).Text = IIf(IsNull(Trim(dbTabCode!pH1Max)), "", Trim(dbTabCode!pH1Max))      '& "Max"
            
    
    .Cell(50, 2).Text = IIf(IsNull(Trim(dbTabCode!pH2Value)), "", Trim(dbTabCode!pH2Value))      '& "Value"
    .Cell(51, 2).Text = IIf(IsNull(Trim(dbTabCode!pH2Min)), "", Trim(dbTabCode!pH2Min))      '& "Min"
    .Cell(52, 2).Text = IIf(IsNull(Trim(dbTabCode!pH2Max)), "", Trim(dbTabCode!pH2Max))      '& "Max"
    
    .Cell(54, 2).Text = IIf(IsNull(Trim(dbTabCode!pH3Value)), "", Trim(dbTabCode!pH3Value))      '& "Value"
    .Cell(55, 2).Text = IIf(IsNull(Trim(dbTabCode!pH3Min)), "", Trim(dbTabCode!pH3Min))      '& "Min"
    .Cell(56, 2).Text = IIf(IsNull(Trim(dbTabCode!pH3Max)), "", Trim(dbTabCode!pH3Max))      '& "Max"
    
    
    .Cell(58, 2).Text = IIf(IsNull(Trim(dbTabCode!WeightValue)), "", Trim(dbTabCode!WeightValue))      '& "Value"
    .Cell(59, 2).Text = IIf(IsNull(Trim(dbTabCode!WeightMin)), "", Trim(dbTabCode!WeightMin))      '& "Min"
    .Cell(60, 2).Text = IIf(IsNull(Trim(dbTabCode!WeightMax)), "", Trim(dbTabCode!WeightMax))      '& "Max"
    Dim sString As String
    sString = IIf(IsNull(Trim(dbTabCode!Certified)), "", Trim(dbTabCode!Certified))
    .Cell(61, 2).Text = IIf(sString = True, "Yes", "No")    '& "Certified"

    .Cell(62, 2).Text = IIf(IsNull(Trim(dbTabCode!RevDate)), "", Trim(dbTabCode!RevDate))      '& "Max"
    
    
    .Cell(63, 2).Text = IIf(IsNull(Trim(dbTabCode!STDMR)), "", Trim(dbTabCode!STDMR))      '& "STDMR "
    .Cell(64, 2).Text = IIf(IsNull(Trim(dbTabCode!STDMR2)), "", Trim(dbTabCode!STDMR2))      '& "STDMR "
    
    .Cell(65, 2).Text = IIf(IsNull(Trim(dbTabCode!MS)), "", Trim(dbTabCode!MS))      '& "STDMR "
    .Cell(66, 2).Text = IIf(IsNull(Trim(dbTabCode!MSEXP)), "", Trim(dbTabCode!MSEXP))
    .Cell(67, 2).Text = IIf(IsNull(Trim(dbTabCode!STDMatrix)), "", Trim(dbTabCode!STDMatrix))      '& "STDMR "
    .Cell(68, 2).Text = IIf(IsNull(Trim(dbTabCode!STDVolume)), "", Trim(dbTabCode!STDVolume))
    .Cell(69, 2).Text = IIf(IsNull(Trim(dbTabCode!STDExp)), "", Trim(dbTabCode!STDExp))      '& "STDMR "
    .Cell(70, 2).Text = IIf(IsNull(Trim(dbTabCode!STDNote)), "", Trim(dbTabCode!STDNote))
    
    .Cell(71, 2).Text = IIf(IsNull(Trim(dbTabCode!FWParameterFormula)), "", Trim(dbTabCode!FWParameterFormula))
    
        'dbCode.Execute ("ALTER TABLE TabCode ADD FWParameterFormula varchar(60) WITH COMPRESSION")
        'dbCode.Execute ("ALTER TABLE TabCode ADD STDMR2 varchar(60) WITH COMPRESSION")
        'dbCode.Execute ("ALTER TABLE TabCode ADD MS varchar(60) WITH COMPRESSION")
        'dbCode.Execute ("ALTER TABLE TabCode ADD MSEXP varchar(60) WITH COMPRESSION")
        'dbCode.Execute ("ALTER TABLE TabCode ADD STDMatrix varchar(60) WITH COMPRESSION")
        'dbCode.Execute ("ALTER TABLE TabCode ADD STDVolume varchar(60) WITH COMPRESSION")
        'dbCode.Execute ("ALTER TABLE TabCode ADD STDExp varchar(60) WITH COMPRESSION")
        'dbCode.Execute ("ALTER TABLE TabCode ADD STDNote varchar(60) WITH COMPRESSION")
        
    
    .Cell(12, 2).Text = Format$(.Cell(12, 2).Text, strUserDecimal)
    .Cell(13, 2).Text = Format$(.Cell(13, 2).Text, strUserDecimal)
    .Cell(16, 2).Text = Format$(.Cell(16, 2).Text, strUserDecimal)
    .Cell(22, 2).Text = Format$(.Cell(22, 2).Text, strUserDecimal)
    .Cell(23, 2).Text = Format$(.Cell(23, 2).Text, strUserDecimal)
    .Cell(24, 2).Text = Format$(.Cell(24, 2).Text, strUserDecimal)
    
    .RowHeight(20) = 0
    
    Dim i As Integer
    For i = 2 To .Cols - 1
        .Column(i).Alignment = cellCenterCenter ' cellLeftCenter
        
    Next
    
    
    'For i = 1 To .Rows - 1
    '    .Cell(i, 2).BackColor = vbColorUnabled
    '    .Cell(i, 2).ForeColor = vbColorForeFixed  ' vbColorTextDarkBlue
    '    .Cell(i, 2).FontBold = True
    '    .Cell(i, 2).Locked = True
    '    .Cell(i, 2).ForeColor = vbColorForeFixed
    '    If i = 11 Or i = 15 Or i = 20 Or i = 24 Or i = 28 Or i = 32 Or i = 36 Or i = 40 Or i = 44 Or i = 48 Or i = 52 Then
    '        .Cell(i, 2).Alignment = cellCenterCenter
    '        .Cell(i, 2).BackColor = vbColorUnabled ' vbColorTextBlue ' &HF0F0F0
    '        .Cell(i, 2).ForeColor = vbWhite  ' &HF0F0F0
    '    End If
        
    'Next


    frCommandInside(8).Visible = True
End With

End Sub


Private Sub ImageTAV_Click(Index As Integer)
Select Case Index
        Case 0
            Unload Me
        
        Case 2
        

End Select
End Sub



Private Sub Label2_Click(Index As Integer)
Image3_Click Index
End Sub

Private Sub lbCommandInside_Click(Index As Integer)
frCommandInside_Click Index
End Sub

Private Sub lbExcel_Click()
frExcel_Click
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

CleanCode

End Sub

Private Sub Timer2_Timer()

Dim i As Integer
    '
    ' start form
    '
    
         AddcmbRevType
     Call SetGridRecipeRevision(Grid2)
      bHilight = True
     RiempiGrid Grd1
     
     

    
    Timer2.Enabled = False
    
    
    
    
End Sub



Private Sub ScorriTabella(ByVal bValue As Boolean)

Dim MyRow As Integer
If Grd1.Rows > 1 Then
    MyRow = IIf(bValue, (IndexTabella * 10) + 10, (IndexTabella * 10) - 19)
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
    
        If F_MsgBox.DoShow(("Delete Selected Code ?"), "Database", , ("Delete"), ("No")) Then
            frCommandInside(8).Visible = False
            If CancellaRecord(MyID) Then
               ' Text1(0) = ""
                GlobalSearch
            Else
            End If
        End If
End Function

Private Sub CleanCode()
Dim i As Integer
With Grd2
    For i = 1 To .Rows - 1
        .Cell(i, 2).Text = ""
    Next
End With
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
Dim sValue As String
Dim iDecimal As Integer
Dim sString As String
Debug.Print "Leave ", Row, Col
With Grd2
    sValue = .Cell(Row, Col).Text
    If Col = 2 Then
        If lRow = Row Then
        
            Select Case Row
                Case 1
                    ' CODE
                    If Len(sValue) = 0 Then
                        PopupMessage 2, "Warning : Code must be a valid value...."
                       
                    End If
                    
                Case 14
                    ' Decimal
                    If sValue = "" Then Exit Sub
                    sString = "Warning : Decimal ERROR"
                    
                    If IsNumeric(sValue) Then
                        
                        If InStr(sValue, ".") Or InStr(sValue, ",") Then
                            PopupMessage 2, "Warning : Decimal must be an Integer ( es.0,1,2,3 )"
                            GoSub err
                        End If
                        If Val(sValue) > 10 Then
                            GoSub err
                        End If
                    Else
                        sString = sString & vbCrLf & "Filed must be a Number..."
                        GoSub err
                    End If
                Case 16
                    ' Decimal
                    If sValue = "" Then Exit Sub
                    sString = "Warning : Fixed ERROR"
                    If IsNumeric(sValue) Then

                    Else
                        sString = sString & vbCrLf & "Filed must be a Number..."
                        GoSub err
                    End If
                Case 22, 26, 30, 34, 38, 42
                    'STDi
                    If sValue = "" Then Exit Sub
                    
                    sString = "Warning : STD Value ERROR" & vbCrLf & "Filed must be a Number..."
                    If Not (IsNumeric(sValue)) Then
                         GoSub err
                        Exit Sub
                    End If
                    
                    CheckToleranceSTDValue sValue, Row
                Case 46 To 56
                    ' pH
                   If sValue = "" Then Exit Sub
                    
                    sString = "Warning : pH Value ERROR"
                    If Not (IsNumeric(sValue)) Then
                         sString = sString & vbCrLf & "pH must be a Number..."
                         GoSub err
                        Exit Sub
                    Else
                        If Val(sValue) > 14 Then
                            sString = sString & vbCrLf & "pH must be < 14 ..."
                            GoSub err
                        End If
                    End If
                Case 62
                    If sValue = "" Then Exit Sub
                    If Not (IsDate(sValue)) Then
                         sString = sString & vbCrLf & "Revision Date must be a Date..."
                         GoSub err
                        Exit Sub
                    Else
                       .Cell(Row, Col).Text = FormatDataLAT(sValue)
                       
                    End If
            
            End Select
        
        
        
        End If
    End If
End With

Exit Sub

err:
PopupMessage 2, sString
Grd2.Cell(Row, Col).Text = ""
Return
End Sub



Private Function CheckToleranceSTDValue(ByVal strValue As String, ByVal Index As Integer)
Dim MinValue As String
Dim MaxValue As String
Dim Perc As Double
Dim Restr As Double
Dim MyDecimalGrid As String
Dim RangeMin As String
Dim RangeMax As String
Dim Fixed As String
Dim AndOr As String
Dim MeasurementUnit As String
Dim UserDecimal As String



On Error GoTo ERR_CHECK

    MyDecimalGrid = Grd2.Cell(14, 2).Text
    MeasurementUnit = Grd2.Cell(10, 2).Text
    Fixed = Grd2.Cell(16, 2).Text
    AndOr = Grd2.Cell(17, 2).Text
    
    If Not (IsNumeric(MyDecimalGrid)) Then
        
        PopupMessage 2, "Please Enter Decimal...."
        Grd2.Cell(14, 2).SetFocus
        Exit Function
    End If
    
    RangeMin = Grd2.Cell(12, 2).Text
    RangeMax = Grd2.Cell(13, 2).Text
    
    UserDecimal = FormatDecimal(Grd2.Cell(14, 2).Text)
    
    
    
    
    
    If strValue <> "" Then

        If Trim(RangeMax) = "" Or Trim(RangeMin) = "" Then
            GoTo okcheck
        End If
        If (Trim(RangeMax) <> "" Or Trim(RangeMax) <> "0") And (Trim(RangeMin) <> "" Or Trim(RangeMin) <> "0") Then
            If CDbl(strValue) <= CDbl(RangeMax) And CDbl(strValue) >= CDbl(RangeMin) Then
                ' ok!!!
okcheck:
                If IsNumeric(Val(Grd2.Cell(18, 2).Text)) Then
                    
                    Perc = Val(Grd2.Cell(18, 2).Text)
                    If Perc > 0 Then Perc = Perc / 100
                Else
                    Perc = 0
                End If
                If IsNumeric(Val(Grd2.Cell(19, 2).Text)) Then
                    Restr = Val(Grd2.Cell(19, 2).Text) / 100
                Else
                    Restr = 0
                End If
                
                If StandardCal(strValue, Fixed, AndOr, Perc, Restr, UserDecimal, MinValue, MaxValue) Then
                    Grd2.Cell(Index + 1, 2).Text = MinValue
                    Grd2.Cell(Index + 2, 2).Text = MaxValue
                  
                End If
            Else
                MessageInfoTime = 2000
                PopupMessage 2, "Warning : Wrong value..." & vbCrLf & "Please Check Tolerance ( Min " & RangeMin & MeasurementUnit & " // Max " & RangeMax & MeasurementUnit & " )", , , "Reagent Range"
                strValue = ""
                Grd2.Cell(Index + 1, 2).Text = ""
                Grd2.Cell(Index + 2, 2).Text = ""
            End If
        End If
    End If
    
ERR_END:
    On Error GoTo 0
    Exit Function
ERR_CHECK:
    MsgBox err.Description
    Resume Next
End Function

Private Function SaveCode() As Boolean
Dim rc As Boolean
Dim MyNewCode As String
Dim RangeMin As String
Dim RangeMax As String

On Error GoTo ERR_SAVE
rc = True
    MyNewCode = Trim(Grd2.Cell(1, 2).Text)
    RangeMin = Trim(Grd2.Cell(12, 2).Text)
    RangeMax = Trim(Grd2.Cell(13, 2).Text)
    
    If MyNewCode = "" Then
        PopupMessage 2, "Please Enter a valid Code!"
        Exit Function
    End If
    
    With dbTabCode
        .Filter = ""
        .Filter = "Code='" & MyNewCode & "' and RangeMin='" & RangeMin & "' and RangeMax='" & RangeMax & "'"
        If .EOF Then
        
            .AddNew
        Else
            If F_MsgBox.DoShow("Code already exsist. Replace Info?") Then
            Else
                Exit Function
            End If
            
        End If
        

      
        
        !Code = Trim(Grd2.Cell(1, 2).Text)
        !Description = Trim(Grd2.Cell(2, 2).Text)
        !Line = Trim(Grd2.Cell(3, 2).Text)
        !Recipe = Trim(Grd2.Cell(4, 2).Text)
        !QCMethod = Trim(Grd2.Cell(5, 2).Text)
        !MeterFamily1 = Trim(Grd2.Cell(6, 2).Text)
        !MeterFamily2 = Trim(Grd2.Cell(7, 2).Text)
        !ParameterMethod = Trim(Grd2.Cell(8, 2).Text)
        !ParameterFormula = Trim(Grd2.Cell(9, 2).Text)
        !MeasurementUnit = Trim(Grd2.Cell(10, 2).Text)
        
        !RangeMin = Trim(Grd2.Cell(12, 2).Text)
        !RangeMax = Trim(Grd2.Cell(13, 2).Text)
        
        If Trim(Grd2.Cell(14, 2).Text) <> "" And Trim(Grd2.Cell(14, 2).Text) <> "/" Then
            !Decimal = Trim(CInt(Grd2.Cell(14, 2).Text))
        End If
        
        !Fixed = Trim(Grd2.Cell(16, 2).Text)
        !AndOr = Trim(Grd2.Cell(17, 2).Text)
        !Percentage = Trim(Grd2.Cell(18, 2).Text)
        !QCRestriction = Trim(Grd2.Cell(19, 2).Text)
        
        
        !STD1Value = Trim(Grd2.Cell(22, 2).Text)
        !STD1Min = Trim(Grd2.Cell(23, 2).Text)
        !STD1Max = Trim(Grd2.Cell(24, 2).Text)
        
        !STD2Value = Trim(Grd2.Cell(26, 2).Text)
        !STD2Min = Trim(Grd2.Cell(27, 2).Text)
        !STD2Max = Trim(Grd2.Cell(28, 2).Text)
        
        !STD3Value = Trim(Grd2.Cell(30, 2).Text)
        !STD3Min = Trim(Grd2.Cell(31, 2).Text)
        !STD3Max = Trim(Grd2.Cell(32, 2).Text)
        
        !STD4Value = Trim(Grd2.Cell(34, 2).Text)
        !STD4Min = Trim(Grd2.Cell(35, 2).Text)
        !STD4Max = Trim(Grd2.Cell(36, 2).Text)
        
        !STD5Value = Trim(Grd2.Cell(38, 2).Text)
        !STD5Min = Trim(Grd2.Cell(39, 2).Text)
        !STD5Max = Trim(Grd2.Cell(40, 2).Text)
        
        !STD6Value = Trim(Grd2.Cell(42, 2).Text)
        !STD6Min = Trim(Grd2.Cell(43, 2).Text)
        !STD6Max = Trim(Grd2.Cell(44, 2).Text)
        
        !pH1Value = Trim(Grd2.Cell(46, 2).Text)
        !pH1Min = Trim(Grd2.Cell(47, 2).Text)
        !pH1Max = Trim(Grd2.Cell(48, 2).Text)
        !pH2Value = Trim(Grd2.Cell(50, 2).Text)
        !pH2Min = Trim(Grd2.Cell(51, 2).Text)
        !pH2Max = Trim(Grd2.Cell(52, 2).Text)
        
        !pH3Value = Trim(Grd2.Cell(54, 2).Text)
        !pH3Min = Trim(Grd2.Cell(55, 2).Text)
        !pH3Max = Trim(Grd2.Cell(56, 2).Text)
        !WeightValue = Trim(Grd2.Cell(58, 2).Text)
        !WeightMin = Trim(Grd2.Cell(59, 2).Text)
        !WeightMax = Trim(Grd2.Cell(60, 2).Text)

        !Certified = IIf(UCase(Trim(Grd2.Cell(61, 2).Text)) = UCase("Yes"), True, False)
        !RevDate = Trim(Grd2.Cell(62, 2).Text)
        !DateModified = Now()
        
        
        !STDMR = Trim(Grd2.Cell(63, 2).Text)
        !STDMR2 = Trim(Grd2.Cell(64, 2).Text)
        !MS = Trim(Grd2.Cell(65, 2).Text)
        !MSEXP = Trim(Grd2.Cell(66, 2).Text)
        !STDMatrix = Trim(Grd2.Cell(67, 2).Text)
        !STDVolume = Trim(Grd2.Cell(68, 2).Text)
        !STDExp = Trim(Grd2.Cell(69, 2).Text)
        !STDNote = Trim(Grd2.Cell(70, 2).Text)
        !FWParameterFormula = Trim(Grd2.Cell(71, 2).Text)

        .Update
        
        .Close
        .Open "SELECT *  FROM TabCode ORDER BY Code", dbCode, adOpenKeyset, adLockOptimistic, adCmdText
    End With
    
    
    
        
       ' .Cell(63, 1).Text = "  " & "MR1"
       ' .Cell(64, 1).Text = "  " & "MR2"
       ' .Cell(65, 1).Text = "  " & "MS"
       ' .Cell(66, 1).Text = "  " & "MS EXP (days)"
       ' .Cell(67, 1).Text = "  " & "STD Matrix"
       ' .Cell(68, 1).Text = "  " & "STD Volume (ml)"
       ' .Cell(69, 1).Text = "  " & "STD EXP (days)"
       ' .Cell(70, 1).Text = "  " & "STD Note"
       ' .Cell(71, 1).Text = "  " & "FW Parameter Formula"


ERR_END:
    On Error GoTo 0
    If rc Then
        PopupMessage 2, "Code : " & MyNewCode & " saved!"
    Else
        PopupMessage 2, "Warning : a problem occurred, please check all entries before Save"
    End If
    
    SaveCode = rc
    Exit Function
    
ERR_SAVE:
    rc = False
    MsgBox err.Description
    Resume Next

End Function
Private Function CancellaRecord(ByVal MyID As Long) As Boolean
Dim rc As Boolean
On Error GoTo ERR_CAN
Dim MyNewCode As String


    If MyID = 0 Then Exit Function
    
    
    rc = True
    With dbTabCode
        .Filter = ""
        .Filter = "ID='" & MyID & "'"
        If .EOF Then
            rc = False
        Else
        
        MyNewCode = Trim(!Code)
        .Delete
        .Update
        End If
    End With
ERR_END:
    On Error GoTo 0
    If rc Then
        PopupMessage 2, "Code : " & MyNewCode & " Deleted!"
    Else
        PopupMessage 2, "Warning : a problem occurred...."
    End If
    
    CancellaRecord = rc
    Exit Function
    
ERR_CAN:
    rc = False
    MsgBox err.Description
    Resume ERR_END:
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

Private Sub AddCombo()

    Combo1.Clear
    Combo1.AddItem "Hanna Code"
    Combo1.AddItem "Chemical MR"
    Combo1.ListIndex = 0
End Sub


'-------------------------------------------------------------------
'
'
'     REVIsiON HISTORY
'
''
'-------------------------------------------------------------------

Private Sub OpenRevisionHistory()

Call ClearRevisionForm

Call GetRecipeRevision(Grid2, MRCode)

lbInside(5) = MRCode & " : Revision History"

frExcel.Visible = IIf(Grid2.Rows > 1, True, False)
Frame6.Visible = IIf(Grid2.Rows > 1, False, True)

                
frRevisionHistory.BackColor = &HF0F0F0
'frRevisionHistory.Move frInside(0).Left, frInside(0).Top, Me.Width - frInside(0).Left * 2, frInside(0).Height
frRevisionHistory.ZOrder
frRevisionHistory.Visible = True
'frCritical.Visible = False

End Sub
Private Sub lbInside_Click(Index As Integer)
    Select Case Index
        
        Case 5
            ' rev history table
            If MRCode <> "" Then Call GetRecipeRevision(Grid2, MRCode)
    
    End Select
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
        Case 0
            'rev number
             'If RevDate <> "" Then
             '   If IsDate(RevDate) Then
             '       Answer = RevDate
             '   End If
             'End If
             
    End Select
    
    
    If txRevision(Index).Locked Then Exit Sub
    
    
  
    If F_InputBox.DoShow(sString, Selected, , , , Answer, , bNumber) Then
    
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
    .Filter = ""
    .Filter = "Recipe='" & Replace(Code, "'", "''") & "' and RevNumber='" & RevNumber & "'"
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
    .Filter = ""
    .Filter = "Recipe='" & Replace(Code, "'", "''") & "' and RevNumber='" & RevNumber & "'"
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

If RevDate <> "" And IsDate(RevDate) Then

    If IsDate(txRevision(0)) Then
        If CDate(RevDate) < CDate(txRevision(0)) Then
        
            Grd2.Cell(62, 2).Text = txRevision(0)
            PopupMessage 2, "Revision Date is changed. Please Save Hanna Code...", , , MRCode
        End If
    Else
    
    
        
    
    End If
Else
    If IsDate(txRevision(0)) Then
        Grd2.Cell(62, 2).Text = txRevision(0)
        PopupMessage 2, "Revision Date is changed. Please Save Hanna Code...", , , MRCode
    End If
        

End If

AddRevision = rc
End Function

Private Sub lbFunction_Click(Index As Integer)
ImCode_Click Index
End Sub



Private Sub ImCode_Click(Index As Integer)
Dim UserChCode As String
Dim UserHannaCode As String
Dim MyCodeID As Long


    Select Case Index
      
                        
        Case 5
            ' aggiungi revision specifics
            If AddRevision(MRCode, txRevision(2)) Then
                 Call GetRecipeRevision(Grid2, MRCode)
            End If
            
            frExcel.Visible = IIf(Grid2.Rows > 1, True, False)
            Frame6.Visible = IIf(Grid2.Rows > 1, False, True)
        Case 4
            ' delete revision specifics
            If DeleteRevision(MRCode, txRevision(2)) Then
                 Call GetRecipeRevision(Grid2, MRCode)
            End If
            
            frExcel.Visible = IIf(Grid2.Rows > 1, True, False)
            Frame6.Visible = IIf(Grid2.Rows > 1, False, True)
            
    End Select
End Sub




Private Sub cmbRevType_Click()
txRevision(1) = cmbRevType
cmbRevType.Visible = False
End Sub
