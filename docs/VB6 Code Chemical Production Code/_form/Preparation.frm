VERSION 5.00
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.7#0"; "FlexCell.ocx"
Begin VB.Form frmPreparation 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Chemical Production"
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
   Icon            =   "Preparation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   12000
   ScaleWidth      =   19200
   Begin VB.PictureBox PBContainerViewport 
      BackColor       =   &H00FFFFFF&
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
      Height          =   9975
      Index           =   0
      Left            =   120
      ScaleHeight     =   9975
      ScaleWidth      =   19245
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   960
      Width           =   19245
      Begin VB.PictureBox PBContainer 
         BackColor       =   &H00E0E0E0&
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
         Height          =   50000
         Left            =   240
         ScaleHeight     =   50002.14
         ScaleMode       =   0  'User
         ScaleWidth      =   19155
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   19155
         Begin VB.Frame frInside 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "&H00F0F0F0&"
            Height          =   9015
            Index           =   5
            Left            =   840
            TabIndex        =   175
            Top             =   29160
            Visible         =   0   'False
            Width           =   17055
            Begin VB.Frame frExcel3 
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
               TabIndex        =   191
               Top             =   6960
               Width           =   3015
               Begin VB.Label Label7 
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
                  TabIndex        =   192
                  Top             =   120
                  Width           =   3015
               End
               Begin VB.Image Image2 
                  Height          =   480
                  Left            =   120
                  MousePointer    =   99  'Custom
                  OLEDropMode     =   1  'Manual
                  Picture         =   "Preparation.frx":29F2
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
               TabIndex        =   190
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
               TabIndex        =   189
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
               TabIndex        =   188
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
               TabIndex        =   187
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
               TabIndex        =   186
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
               TabIndex        =   185
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
               Index           =   3
               Left            =   0
               TabIndex        =   183
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
                  Alignment       =   2  'Center
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
                  Index           =   4
                  Left            =   0
                  TabIndex        =   184
                  Top             =   75
                  Width           =   16935
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
               Index           =   19
               Left            =   12960
               TabIndex        =   181
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
                  Index           =   19
                  Left            =   0
                  TabIndex        =   182
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00886010&
               BorderStyle     =   0  'None
               Height          =   1335
               Left            =   5880
               TabIndex        =   178
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
                  Index           =   0
                  Left            =   1380
                  TabIndex        =   180
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
                  Index           =   1
                  Left            =   0
                  TabIndex        =   179
                  Top             =   360
                  Width           =   4995
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
               Index           =   21
               Left            =   12960
               TabIndex        =   176
               Top             =   7560
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Exit Revision"
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
                  Index           =   21
                  Left            =   0
                  TabIndex        =   177
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin FlexCell.Grid Grid5 
               Height          =   4695
               Left            =   0
               TabIndex        =   193
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
               Index           =   1
               Left            =   8400
               TabIndex        =   203
               Top             =   7440
               Width           =   1815
            End
            Begin VB.Label lbFunction 
               BackStyle       =   0  'Transparent
               Height          =   855
               Index           =   0
               Left            =   6840
               TabIndex        =   202
               Top             =   7440
               Width           =   1575
            End
            Begin VB.Label lbRevHist 
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
               TabIndex        =   201
               Top             =   6240
               Width           =   1695
            End
            Begin VB.Label lbRevHist 
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
               TabIndex        =   200
               Top             =   5760
               Width           =   855
            End
            Begin VB.Label lbRevHist 
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
               TabIndex        =   199
               Top             =   5760
               Width           =   1215
            End
            Begin VB.Label lbRevHist 
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
               TabIndex        =   198
               Top             =   5760
               Width           =   735
            End
            Begin VB.Label lbRevHist 
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
               TabIndex        =   197
               Top             =   5760
               Width           =   1695
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Delete Specifics"
               ForeColor       =   &H00808080&
               Height          =   255
               Left            =   8640
               TabIndex        =   196
               Top             =   7875
               Width           =   1500
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Save Specifics"
               ForeColor       =   &H00808080&
               Height          =   255
               Left            =   6960
               TabIndex        =   195
               Top             =   7875
               Width           =   1335
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   0
               Left            =   7440
               Picture         =   "Preparation.frx":5DD4
               ToolTipText     =   "4000"
               Top             =   7485
               Width           =   240
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   1
               Left            =   9120
               Picture         =   "Preparation.frx":67D6
               ToolTipText     =   "4000"
               Top             =   7485
               Width           =   240
            End
            Begin VB.Line Line11 
               BorderColor     =   &H00D0D0D0&
               X1              =   240
               X2              =   16800
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
               Index           =   0
               Left            =   5265
               TabIndex        =   194
               Top             =   8640
               Width           =   6435
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Caption         =   "&H00E0E0E0&"
            Height          =   9015
            Index           =   4
            Left            =   960
            TabIndex        =   143
            Top             =   40798
            Visible         =   0   'False
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
               Index           =   22
               Left            =   12960
               TabIndex        =   173
               Top             =   7560
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Exit Notes"
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
                  Index           =   22
                  Left            =   0
                  TabIndex        =   174
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.TextBox txPreparation 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
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
               Left            =   6000
               TabIndex        =   169
               Top             =   5760
               Width           =   2415
            End
            Begin VB.ComboBox cmbNotes 
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
               Left            =   6000
               Style           =   2  'Dropdown List
               TabIndex        =   168
               Top             =   5760
               Visible         =   0   'False
               Width           =   2415
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00886010&
               BorderStyle     =   0  'None
               Height          =   1335
               Left            =   5880
               TabIndex        =   154
               Top             =   2400
               Width           =   5055
               Begin VB.Label lbChem 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Empty List..."
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Index           =   5
                  Left            =   0
                  TabIndex        =   156
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
                  TabIndex        =   155
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
               Index           =   17
               Left            =   12960
               TabIndex        =   152
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
                  Index           =   17
                  Left            =   0
                  TabIndex        =   153
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00886010&
               BorderStyle     =   0  'None
               Caption         =   "&H00E0E0E0&"
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
               Left            =   0
               TabIndex        =   149
               Top             =   0
               Width           =   17055
               Begin VB.Label Label23 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe Preparation"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Left            =   15165
                  TabIndex        =   151
                  Top             =   120
                  Width           =   1725
               End
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Preparation Notes"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   3
                  Left            =   7560
                  TabIndex        =   150
                  Top             =   75
                  Width           =   2055
               End
            End
            Begin VB.TextBox txPreparation 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
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
               TabIndex        =   148
               Top             =   5760
               Width           =   2415
            End
            Begin VB.TextBox txPreparation 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
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
               Left            =   10320
               TabIndex        =   147
               Top             =   5760
               Width           =   5655
            End
            Begin VB.TextBox txPreparation 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
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
               Index           =   3
               Left            =   2160
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   146
               Top             =   6240
               Width           =   13815
            End
            Begin VB.Frame frExcel2 
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
               TabIndex        =   144
               Top             =   6960
               Width           =   3015
               Begin VB.Image Image1 
                  Height          =   480
                  Left            =   120
                  MousePointer    =   99  'Custom
                  OLEDropMode     =   1  'Manual
                  Picture         =   "Preparation.frx":71D8
                  Top             =   0
                  Width           =   480
               End
               Begin VB.Label Label4 
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
                  TabIndex        =   145
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin FlexCell.Grid Grid4 
               Height          =   4695
               Left            =   0
               TabIndex        =   157
               TabStop         =   0   'False
               Top             =   600
               Width           =   17055
               _ExtentX        =   30083
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
               Height          =   975
               Index           =   5
               Left            =   6720
               TabIndex        =   159
               Top             =   7320
               Width           =   1695
            End
            Begin VB.Label lbFunction 
               BackStyle       =   0  'Transparent
               Height          =   975
               Index           =   4
               Left            =   8400
               TabIndex        =   158
               Top             =   7320
               Width           =   1815
            End
            Begin VB.Label lbRevision 
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
               Left            =   5160
               TabIndex        =   170
               Top             =   5760
               Width           =   735
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
               TabIndex        =   165
               Top             =   8640
               Width           =   6435
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   4
               Left            =   9120
               Picture         =   "Preparation.frx":A5BA
               ToolTipText     =   "4000"
               Top             =   7485
               Width           =   240
            End
            Begin VB.Image ImCode 
               Height          =   240
               Index           =   5
               Left            =   7440
               Picture         =   "Preparation.frx":AFBC
               ToolTipText     =   "4000"
               Top             =   7485
               Width           =   240
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Save Note"
               ForeColor       =   &H00808080&
               Height          =   255
               Left            =   7080
               TabIndex        =   164
               Top             =   7875
               Width           =   1005
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Delete Note"
               ForeColor       =   &H00808080&
               Height          =   255
               Left            =   8640
               TabIndex        =   163
               Top             =   7875
               Width           =   1170
            End
            Begin VB.Label lbRevision 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
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
               TabIndex        =   162
               Top             =   5760
               Width           =   1695
            End
            Begin VB.Label lbRevision 
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
               Index           =   2
               Left            =   9240
               TabIndex        =   161
               Top             =   5760
               Width           =   855
            End
            Begin VB.Label lbRevision 
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
               Index           =   3
               Left            =   240
               TabIndex        =   160
               Top             =   6240
               Width           =   1695
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Caption         =   "&H00F0F0F0&"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8655
            Index           =   2
            Left            =   1080
            TabIndex        =   26
            Top             =   19080
            Visible         =   0   'False
            Width           =   17295
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Index           =   18
               Left            =   10560
               TabIndex        =   140
               Top             =   1440
               Width           =   1080
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   13
               Left            =   120
               TabIndex        =   120
               Top             =   5880
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Scan QRCode"
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
                  Index           =   13
                  Left            =   0
                  TabIndex        =   121
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Index           =   17
               Left            =   6720
               TabIndex        =   112
               Top             =   4800
               Width           =   6255
            End
            Begin VB.Frame CommandRecalculate 
               BackColor       =   &H000000C0&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Left            =   8040
               TabIndex        =   110
               Top             =   5880
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandRecalculate 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recalculate"
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
                  TabIndex        =   111
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.PictureBox PicTolerance 
               BorderStyle     =   0  'None
               Height          =   135
               Left            =   6720
               ScaleHeight     =   135
               ScaleWidth      =   4095
               TabIndex        =   108
               Top             =   4200
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0F0F0&
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
               Height          =   400
               Index           =   16
               Left            =   12960
               Locked          =   -1  'True
               TabIndex        =   106
               Text            =   "-21 %"
               Top             =   3840
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0F0F0&
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
               Height          =   400
               Index           =   15
               Left            =   12960
               Locked          =   -1  'True
               TabIndex        =   104
               Text            =   "-21 %"
               Top             =   3360
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   14
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   102
               Text            =   "-200,221"
               Top             =   3840
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0F0F0&
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
               Height          =   400
               Index           =   12
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   100
               Text            =   "-21 %"
               Top             =   4800
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0F0F0&
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
               Height          =   400
               Index           =   11
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   98
               Text            =   "-200,221"
               Top             =   4320
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Height          =   600
               Index           =   13
               Left            =   6720
               TabIndex        =   96
               Text            =   "1229,998"
               Top             =   3600
               Width           =   4095
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   10
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   94
               Text            =   "1300,400"
               Top             =   3360
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Index           =   9
               Left            =   13680
               TabIndex        =   92
               Top             =   1800
               Width           =   2415
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   9
               Left            =   6960
               TabIndex        =   81
               Top             =   8040
               Visible         =   0   'False
               Width           =   3015
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
                  Index           =   9
                  Left            =   0
                  TabIndex        =   82
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00307030&
               BorderStyle     =   0  'None
               Height          =   495
               Index           =   8
               Left            =   11160
               TabIndex        =   79
               Top             =   5880
               Visible         =   0   'False
               Width           =   3015
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
                  Index           =   8
                  Left            =   0
                  TabIndex        =   80
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   7
               Left            =   14280
               TabIndex        =   77
               Top             =   5880
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Exit Acquisition"
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
                  TabIndex        =   78
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   76
               Top             =   1080
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Left            =   7200
               Locked          =   -1  'True
               TabIndex        =   75
               Top             =   1080
               Width           =   5295
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Left            =   13680
               Locked          =   -1  'True
               TabIndex        =   74
               Top             =   1080
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Left            =   2280
               TabIndex        =   73
               Top             =   1440
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Index           =   4
               Left            =   7200
               TabIndex        =   72
               Top             =   1440
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Index           =   5
               Left            =   13680
               TabIndex        =   71
               Top             =   1440
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Index           =   6
               Left            =   2280
               TabIndex        =   70
               Top             =   1800
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Index           =   7
               Left            =   7200
               TabIndex        =   69
               Top             =   1800
               Width           =   2415
            End
            Begin VB.TextBox txAcquisition 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Index           =   8
               Left            =   11640
               TabIndex        =   68
               Top             =   1800
               Width           =   840
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00886010&
               BorderStyle     =   0  'None
               Caption         =   "l"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   2
               Left            =   120
               TabIndex        =   27
               Top             =   240
               Width           =   17295
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Acquisition Details"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   2
                  Left            =   -120
                  TabIndex        =   29
                  Top             =   120
                  Width           =   17235
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Preparation"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   15960
                  TabIndex        =   28
                  Top             =   120
                  Width           =   1050
               End
            End
            Begin VB.Label lbNote 
               Alignment       =   2  'Center
               BackColor       =   &H00886010&
               Caption         =   "Note"
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
               Height          =   360
               Left            =   120
               TabIndex        =   142
               Top             =   5360
               Visible         =   0   'False
               Width           =   17130
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Exp"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   18
               Left            =   10125
               TabIndex        =   141
               Top             =   1440
               Width           =   300
            End
            Begin VB.Label lbCritical 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Critical RM"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H001050D0&
               Height          =   360
               Left            =   960
               TabIndex        =   137
               Top             =   2280
               Visible         =   0   'False
               Width           =   15210
            End
            Begin VB.Label lbChemical 
               Alignment       =   2  'Center
               BackColor       =   &H00307030&
               BackStyle       =   0  'Transparent
               Caption         =   "CM4434"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00307030&
               Height          =   660
               Left            =   0
               TabIndex        =   134
               Top             =   7200
               Width           =   17235
            End
            Begin VB.Label lbAcquisition 
               BackStyle       =   0  'Transparent
               Caption         =   "Note"
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
               Index           =   17
               Left            =   6720
               TabIndex        =   113
               Top             =   4560
               Width           =   1095
            End
            Begin VB.Label lbRecalculate 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Acquisition exceed Tolerance : recalculate Recipe Total Weight or Save?"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   5400
               TabIndex        =   109
               Top             =   6720
               Visible         =   0   'False
               Width           =   6555
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tolerance %"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   16
               Left            =   11700
               TabIndex        =   107
               Top             =   3840
               Width           =   1065
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tolerance ( g )"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   15
               Left            =   11490
               TabIndex        =   105
               Top             =   3360
               Width           =   1275
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Real Weight ( g )"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   1740
               TabIndex        =   103
               Top             =   3840
               Width           =   1485
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Variance %"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   12
               Left            =   1740
               TabIndex        =   101
               Top             =   4800
               Width           =   1485
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Variance ( g )"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   11
               Left            =   1740
               TabIndex        =   99
               Top             =   4320
               Width           =   1485
            End
            Begin VB.Label lbAcquisition 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Acquisition Weight ( g )"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   13
               Left            =   6720
               TabIndex        =   97
               Top             =   3240
               Width           =   2085
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Theoretical Weight ( g )"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   10
               Left            =   1740
               TabIndex        =   95
               Top             =   3360
               Width           =   1485
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Package"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   9
               Left            =   12720
               TabIndex        =   93
               Top             =   1800
               Width           =   795
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "RM Code"
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
               Left            =   1080
               TabIndex        =   91
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00B0B0B0&
               X1              =   960
               X2              =   16200
               Y1              =   2760
               Y2              =   2760
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
               Index           =   1
               Left            =   5520
               TabIndex        =   90
               Top             =   1080
               Width           =   1575
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cas"
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
               Left            =   12840
               TabIndex        =   89
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Manufacturer"
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
               Left            =   840
               TabIndex        =   88
               Top             =   1440
               Width           =   1335
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Manufacturer Code"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   5340
               TabIndex        =   87
               Top             =   1440
               Width           =   1755
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Manufacturer Lot "
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   11985
               TabIndex        =   86
               Top             =   1440
               Width           =   1590
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Delivery Date"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   6
               Left            =   945
               TabIndex        =   85
               Top             =   1800
               Width           =   1230
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qty delivered"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   7
               Left            =   5865
               TabIndex        =   84
               Top             =   1800
               Width           =   1230
            End
            Begin VB.Label lbAcquisition 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Week Delivery"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   8
               Left            =   10200
               TabIndex        =   83
               Top             =   1800
               Width           =   1305
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Caption         =   "&H00F0F0F0&"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6480
            Index           =   3
            Left            =   1560
            TabIndex        =   123
            Top             =   12600
            Visible         =   0   'False
            Width           =   17175
            Begin VB.Frame Frame3 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Caption         =   "l"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   0
               Left            =   0
               TabIndex        =   128
               Top             =   0
               Width           =   17175
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Set Hanna Code Lot Number"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00105010&
                  Height          =   255
                  Left            =   14550
                  TabIndex        =   130
                  Top             =   120
                  Width           =   2535
               End
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hanna Code"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   12.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00105010&
                  Height          =   285
                  Index           =   0
                  Left            =   90
                  TabIndex        =   129
                  Top             =   105
                  Width           =   17085
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
               Index           =   16
               Left            =   14160
               TabIndex        =   126
               Top             =   5880
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Exit Lot"
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
                  Index           =   16
                  Left            =   0
                  TabIndex        =   127
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   15
               Left            =   240
               TabIndex        =   124
               Top             =   5880
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Add Acquisition"
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
                  Index           =   15
                  Left            =   0
                  TabIndex        =   125
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin FlexCell.Grid Grid3 
               Height          =   4935
               Left            =   0
               TabIndex        =   133
               TabStop         =   0   'False
               Top             =   600
               Width           =   17175
               _ExtentX        =   30295
               _ExtentY        =   8705
               AllowUserSort   =   -1  'True
               Appearance      =   0
               BackColor1      =   14737632
               BackColor2      =   14737632
               BackColorActiveCellSel=   14737632
               BackColorBkg    =   14737632
               BackColorFixed  =   14737632
               BackColorFixedSel=   14737632
               BackColorScrollBar=   15592423
               BorderColor     =   14737632
               CellBorderColor =   14737632
               CellBorderColorFixed=   14737632
               Cols            =   5
               DefaultFontName =   "Segoe UI"
               DefaultFontSize =   9.75
               BoldFixedCell   =   0   'False
               DisplayRowIndex =   -1  'True
               DrawMode        =   1
               DefaultRowHeight=   20
               FixedRowColStyle=   0
               ForeColorFixed  =   8937488
               GridColor       =   14737632
               Rows            =   1
               ScrollBarStyle  =   0
               SelectionMode   =   3
               MultiSelect     =   0   'False
               DateFormat      =   0
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   17280
               Y1              =   5640
               Y2              =   5640
            End
         End
         Begin VB.TextBox txQRCode 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   840
            TabIndex        =   122
            Text            =   "Text1"
            Top             =   17760
            Visible         =   0   'False
            Width           =   14535
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
            Index           =   12
            Left            =   14880
            TabIndex        =   118
            Top             =   840
            Visible         =   0   'False
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Back To Acquisition"
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
               Index           =   12
               Left            =   0
               TabIndex        =   119
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Caption         =   "&H00F0F0F0&"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6480
            Index           =   1
            Left            =   1080
            TabIndex        =   30
            Top             =   11040
            Width           =   17175
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   11
               Left            =   3120
               TabIndex        =   116
               Top             =   5880
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Delete Acquisition"
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
                  Index           =   11
                  Left            =   0
                  TabIndex        =   117
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00644603&
               BorderStyle     =   0  'None
               Caption         =   "Image14"
               Height          =   495
               Index           =   10
               Left            =   0
               TabIndex        =   114
               Top             =   5880
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Add Acquisition"
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
                  Index           =   10
                  Left            =   0
                  TabIndex        =   115
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
               Index           =   2
               Left            =   14160
               TabIndex        =   40
               Top             =   5880
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe Table"
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
                  TabIndex        =   41
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00D0D0D0&
               BorderStyle     =   0  'None
               Caption         =   "l"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   8
               Left            =   0
               TabIndex        =   31
               Top             =   0
               Width           =   17175
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Acquisitions"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   12.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00105010&
                  Height          =   285
                  Index           =   8
                  Left            =   7830
                  TabIndex        =   33
                  Top             =   100
                  Width           =   1485
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select Component "
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00105010&
                  Height          =   255
                  Left            =   15345
                  TabIndex        =   32
                  Top             =   120
                  Width           =   1740
               End
            End
            Begin FlexCell.Grid Grid2 
               Height          =   4935
               Left            =   0
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   600
               Width           =   17175
               _ExtentX        =   30295
               _ExtentY        =   8705
               AllowUserSort   =   -1  'True
               Appearance      =   0
               BackColor1      =   14737632
               BackColor2      =   14737632
               BackColorActiveCellSel=   14737632
               BackColorBkg    =   14737632
               BackColorFixed  =   14737632
               BackColorFixedSel=   14737632
               BackColorScrollBar=   15592423
               BorderColor     =   14737632
               CellBorderColor =   14737632
               CellBorderColorFixed=   14737632
               Cols            =   5
               DefaultFontName =   "Segoe UI"
               DefaultFontSize =   9.75
               BoldFixedCell   =   0   'False
               DisplayRowIndex =   -1  'True
               DrawMode        =   1
               DefaultRowHeight=   20
               FixedRowColStyle=   0
               ForeColorFixed  =   8937488
               GridColor       =   14737632
               Rows            =   1
               ScrollBarStyle  =   0
               SelectionMode   =   3
               MultiSelect     =   0   'False
               DateFormat      =   0
            End
            Begin VB.Line Line12 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   17280
               Y1              =   5640
               Y2              =   5640
            End
         End
         Begin VB.Frame frInside 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Caption         =   "&H00F0F0F0&"
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
            Index           =   0
            Left            =   960
            TabIndex        =   17
            Top             =   1080
            Width           =   17175
            Begin VB.CheckBox Check1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Test Lot"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   11040
               TabIndex        =   246
               Top             =   1800
               Width           =   1575
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
               Index           =   25
               Left            =   0
               TabIndex        =   240
               Top             =   8280
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Material Requisition"
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
                  Height          =   240
                  Index           =   25
                  Left            =   0
                  TabIndex        =   241
                  Top             =   120
                  Width           =   2925
               End
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   9
               Left            =   11040
               TabIndex        =   204
               Top             =   1440
               Visible         =   0   'False
               Width           =   1695
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
               Index           =   20
               Left            =   11040
               TabIndex        =   171
               Top             =   7680
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Revision History"
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
                  Index           =   20
                  Left            =   0
                  TabIndex        =   172
                  Top             =   120
                  Width           =   3015
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
               Index           =   18
               Left            =   14160
               TabIndex        =   166
               Top             =   7680
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Preparation Notes"
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
                  Index           =   18
                  Left            =   0
                  TabIndex        =   167
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Index           =   8
               Left            =   11040
               TabIndex        =   138
               Top             =   720
               Width           =   1695
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
               Left            =   6240
               TabIndex        =   135
               Top             =   7680
               Width           =   3015
               Begin VB.Image Image 
                  Height          =   480
                  Left            =   120
                  MousePointer    =   99  'Custom
                  OLEDropMode     =   1  'Manual
                  Picture         =   "Preparation.frx":B9BE
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
                  TabIndex        =   136
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
               Index           =   14
               Left            =   11040
               TabIndex        =   131
               Top             =   7080
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Preparation Lot Number"
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
                  Index           =   14
                  Left            =   0
                  TabIndex        =   132
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame frCommandInside 
               BackColor       =   &H00008000&
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
               Index           =   4
               Left            =   0
               TabIndex        =   66
               Top             =   7080
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Save Preparation"
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
                  TabIndex        =   67
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
               Index           =   5
               Left            =   3120
               TabIndex        =   64
               Top             =   7080
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Add Acquisition"
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
                  TabIndex        =   65
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
               Index           =   6
               Left            =   6240
               TabIndex        =   60
               Top             =   7080
               Visible         =   0   'False
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Product Calssification"
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
                  TabIndex        =   61
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Index           =   7
               Left            =   14040
               TabIndex        =   58
               Top             =   1440
               Width           =   2535
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Index           =   6
               Left            =   8400
               TabIndex        =   56
               Top             =   1080
               Width           =   1695
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Left            =   2400
               TabIndex        =   49
               Top             =   720
               Width           =   2655
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Left            =   7200
               TabIndex        =   48
               Top             =   720
               Width           =   1935
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Left            =   14880
               TabIndex        =   47
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Left            =   3960
               TabIndex        =   46
               Top             =   1080
               Width           =   1695
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Index           =   4
               Left            =   14040
               TabIndex        =   45
               Top             =   1080
               Width           =   2535
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Caption         =   "l"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   5
               Left            =   960
               TabIndex        =   43
               Top             =   0
               Width           =   15255
               Begin VB.Line Line8 
                  BorderColor     =   &H00B0B0B0&
                  X1              =   240
                  X2              =   15480
                  Y1              =   480
                  Y2              =   480
               End
               Begin VB.Label lbInside 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe for production | Preparation"
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00105010&
                  Height          =   270
                  Index           =   5
                  Left            =   0
                  TabIndex        =   44
                  Top             =   120
                  Width           =   3945
               End
            End
            Begin VB.TextBox txFormulation 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
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
               Index           =   5
               Left            =   2400
               TabIndex        =   42
               Top             =   1440
               Width           =   7695
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
               Index           =   1
               Left            =   3120
               TabIndex        =   38
               Top             =   7680
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Delete Component from Recipe"
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
                  TabIndex        =   39
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
               Left            =   0
               TabIndex        =   36
               Top             =   7680
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Add Component to  Recipe"
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
                  TabIndex        =   37
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame Frame2 
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
               Left            =   6000
               TabIndex        =   23
               Top             =   3840
               Visible         =   0   'False
               Width           =   5055
               Begin VB.Label Label1 
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
                  Index           =   0
                  Left            =   1920
                  TabIndex        =   24
                  Top             =   555
                  Width           =   1155
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
               Index           =   3
               Left            =   14160
               TabIndex        =   21
               Top             =   7080
               Width           =   3015
               Begin VB.Label lbCommandInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Acquisitions Table"
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
                  TabIndex        =   22
                  Top             =   120
                  Width           =   3015
               End
            End
            Begin VB.Frame Frame3 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Caption         =   "l"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   540
               Index           =   1
               Left            =   0
               TabIndex        =   18
               Top             =   2280
               Width           =   17175
               Begin VB.Label lbRecipeDensity 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Select Component "
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00644603&
                  Height          =   240
                  Left            =   120
                  TabIndex        =   63
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1665
               End
               Begin VB.Label lbExpand 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "+ "
                  BeginProperty Font 
                     Name            =   "Century Gothic"
                     Size            =   20.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00105010&
                  Height          =   480
                  Left            =   16440
                  TabIndex        =   20
                  Top             =   0
                  Width           =   840
               End
               Begin VB.Label lbInside 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Recipe"
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
                  Height          =   345
                  Index           =   1
                  Left            =   0
                  TabIndex        =   19
                  Top             =   105
                  Width           =   17085
               End
            End
            Begin FlexCell.Grid Grid1 
               Height          =   3855
               Left            =   0
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   2880
               Width           =   17175
               _ExtentX        =   30295
               _ExtentY        =   6800
               AllowUserSort   =   -1  'True
               Appearance      =   0
               BackColor1      =   14737632
               BackColor2      =   14737632
               BackColorBkg    =   14737632
               BackColorFixed  =   14737632
               BackColorFixedSel=   14737632
               BackColorScrollBar=   15592423
               BorderColor     =   14737632
               CellBorderColor =   14737632
               CellBorderColorFixed=   14737632
               Cols            =   5
               DefaultFontName =   "Segoe UI"
               DefaultFontSize =   9.75
               BoldFixedCell   =   0   'False
               DisplayRowIndex =   -1  'True
               DrawMode        =   1
               DefaultRowHeight=   20
               FixedRowColStyle=   0
               ForeColorFixed  =   6571523
               GridColor       =   14737632
               Rows            =   1
               ScrollBarStyle  =   0
               SelectionMode   =   3
               MultiSelect     =   0   'False
               DateFormat      =   0
            End
            Begin VB.Label lbFormulation 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Prep .Lot"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   9
               Left            =   10185
               TabIndex        =   205
               Top             =   1440
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Exp Date"
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   8
               Left            =   10080
               TabIndex        =   139
               Top             =   720
               Width           =   795
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
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
               Index           =   7
               Left            =   12960
               TabIndex        =   59
               Top             =   1440
               Width           =   825
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Preparation Week"
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
               Index           =   6
               Left            =   6360
               TabIndex        =   57
               Top             =   1080
               Width           =   1620
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "Recipe by"
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
               Left            =   1200
               TabIndex        =   55
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "Preparation Date"
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
               Left            =   5400
               TabIndex        =   54
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label lbFormulation 
               BackStyle       =   0  'Transparent
               Caption         =   "# Prep Week"
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
               Left            =   13440
               TabIndex        =   53
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Planned Preparation Week"
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
               Left            =   1200
               TabIndex        =   52
               Top             =   1080
               Width           =   2385
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Planning Reference"
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
               Left            =   12000
               TabIndex        =   51
               Top             =   1080
               Width           =   1725
            End
            Begin VB.Label lbFormulation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Note"
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
               Index           =   5
               Left            =   1200
               TabIndex        =   50
               Top             =   1440
               Width           =   450
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00B0B0B0&
               X1              =   960
               X2              =   16800
               Y1              =   2160
               Y2              =   2160
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00D0D0D0&
               X1              =   -120
               X2              =   17040
               Y1              =   6840
               Y2              =   6840
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Procedure : Recipe Preparation"
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   18975
         End
      End
   End
   Begin VB.PictureBox PBTitle 
      BackColor       =   &H00105010&
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
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   0
      Top             =   0
      Width           =   19215
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00105010&
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
         Left            =   2160
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   2175
         TabIndex        =   238
         Top             =   0
         Width           =   2175
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
            MousePointer    =   99  'Custom
            TabIndex        =   239
            Top             =   640
            Width           =   2070
         End
         Begin VB.Image Image3 
            Height          =   480
            Index           =   1
            Left            =   840
            MousePointer    =   99  'Custom
            Picture         =   "Preparation.frx":EDA0
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox PicMenu 
         BackColor       =   &H00307030&
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
         MousePointer    =   99  'Custom
         ScaleHeight     =   1095
         ScaleWidth      =   2175
         TabIndex        =   1
         Top             =   0
         Width           =   2175
         Begin VB.Image Image3 
            Height          =   480
            Index           =   0
            Left            =   840
            MousePointer    =   99  'Custom
            Picture         =   "Preparation.frx":12182
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Preparation"
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
            TabIndex        =   2
            Top             =   640
            Width           =   2070
         End
      End
      Begin VB.Label lbLine 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preparation"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   62
         Top             =   200
         Visible         =   0   'False
         Width           =   19215
      End
      Begin VB.Label lbWait 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "Wait : Loading Data..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5760
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   7575
      End
      Begin VB.Label blTable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preparation"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   16620
         TabIndex        =   3
         Top             =   200
         Width           =   2310
      End
   End
   Begin VB.PictureBox PBContainerViewport 
      BackColor       =   &H00F0F0F0&
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
      Height          =   9975
      Index           =   1
      Left            =   360
      ScaleHeight     =   9975
      ScaleWidth      =   19095
      TabIndex        =   206
      Top             =   1080
      Width           =   19095
      Begin VB.Frame frIRequisition 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "&H00F0F0F0&"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Index           =   0
         Left            =   1800
         TabIndex        =   226
         Top             =   4680
         Width           =   15255
         Begin VB.Frame frCommandInside 
            BackColor       =   &H00644603&
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
            Index           =   26
            Left            =   6000
            TabIndex        =   244
            Top             =   4200
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Delete Component"
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
               Index           =   26
               Left            =   0
               TabIndex        =   245
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   6
            Left            =   0
            TabIndex        =   233
            Top             =   0
            Width           =   15255
            Begin VB.Line Line9 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   15240
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label lbInside 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Preparation : Material Requisition Table"
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
               Index           =   6
               Left            =   0
               TabIndex        =   235
               Top             =   120
               Width           =   4470
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Material Requisition"
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
               Left            =   13395
               TabIndex        =   234
               Top             =   180
               Visible         =   0   'False
               Width           =   1755
            End
         End
         Begin VB.Frame frCommandInside 
            BackColor       =   &H00644603&
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
            Index           =   24
            Left            =   12240
            TabIndex        =   231
            Top             =   4200
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Update Table"
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
               Index           =   24
               Left            =   0
               TabIndex        =   232
               Top             =   120
               Width           =   3015
            End
         End
         Begin VB.Frame Frame1 
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
            Left            =   5400
            TabIndex        =   229
            Top             =   1200
            Visible         =   0   'False
            Width           =   5055
            Begin VB.Label Label1 
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
               Index           =   4
               Left            =   1920
               TabIndex        =   230
               Top             =   555
               Width           =   1155
            End
         End
         Begin VB.Frame frCommandInside 
            BackColor       =   &H00644603&
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
            Index           =   23
            Left            =   9120
            TabIndex        =   227
            Top             =   4200
            Width           =   3015
            Begin VB.Label lbCommandInside 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Open pdf Folder"
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
               Index           =   23
               Left            =   0
               TabIndex        =   228
               Top             =   120
               Width           =   3015
            End
         End
         Begin FlexCell.Grid Grid6 
            Height          =   3135
            Left            =   0
            TabIndex        =   236
            TabStop         =   0   'False
            Top             =   720
            Width           =   15255
            _ExtentX        =   26908
            _ExtentY        =   5530
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
            SelectionMode   =   1
            MultiSelect     =   0   'False
            DateFormat      =   0
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Click on Component to change Quantity"
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
            Left            =   0
            TabIndex        =   237
            Top             =   4200
            Width           =   3645
         End
      End
      Begin VB.Frame frIRequisition 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   3975
         Index           =   1
         Left            =   1800
         TabIndex        =   207
         Top             =   720
         Width           =   15615
         Begin VB.Frame Frame3 
            BackColor       =   &H00F0F0F0&
            BorderStyle     =   0  'None
            Caption         =   "l"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   7
            Left            =   0
            TabIndex        =   214
            Top             =   0
            Width           =   15255
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Material Requisition"
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
               Left            =   13395
               TabIndex        =   216
               Top             =   180
               Visible         =   0   'False
               Width           =   1755
            End
            Begin VB.Label lbInside 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Material Requisition Document"
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
               Index           =   7
               Left            =   0
               TabIndex        =   215
               Top             =   120
               Width           =   3510
            End
            Begin VB.Line Line6 
               BorderColor     =   &H00B0B0B0&
               X1              =   0
               X2              =   15240
               Y1              =   480
               Y2              =   480
            End
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   213
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   1
            Left            =   6480
            TabIndex        =   212
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   2
            Left            =   12120
            TabIndex        =   211
            Top             =   960
            Width           =   3135
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   210
            Top             =   1440
            Width           =   2175
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   4
            Left            =   6480
            TabIndex        =   209
            Top             =   1440
            Width           =   8775
         End
         Begin VB.TextBox txDocument 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   5
            Left            =   1920
            TabIndex        =   208
            Top             =   1920
            Width           =   4575
         End
         Begin VB.Image impdf 
            Height          =   480
            Left            =   6000
            Picture         =   "Preparation.frx":14B74
            Top             =   3240
            Width           =   480
         End
         Begin VB.Label lbpdf 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Save pdf for Material Requisition"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00644603&
            Height          =   255
            Left            =   6600
            TabIndex        =   224
            Top             =   3360
            Width           =   3105
         End
         Begin VB.Label lbDocument 
            BackStyle       =   0  'Transparent
            Caption         =   "Document No: MR-"
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   223
            Top             =   1005
            Width           =   1935
         End
         Begin VB.Label lbDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operator"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   1
            Left            =   5520
            TabIndex        =   222
            Top             =   1005
            Width           =   885
         End
         Begin VB.Label lbDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Production line no./dep."
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   2
            Left            =   9600
            TabIndex        =   221
            Top             =   1005
            Width           =   2370
         End
         Begin VB.Label lbDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "today "
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   220
            Top             =   1485
            Width           =   630
         End
         Begin VB.Label lbDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Planning Reference"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   4
            Left            =   4560
            TabIndex        =   219
            Top             =   1440
            Width           =   1860
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fill Document form and save pdf for material requisition"
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
            Left            =   10200
            TabIndex        =   218
            Top             =   2040
            Width           =   4920
         End
         Begin VB.Label lbDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dep. Manager"
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   217
            Top             =   1965
            Width           =   1395
         End
         Begin VB.Label lbCommand 
            BackColor       =   &H00C0FFC0&
            Height          =   735
            Left            =   5640
            TabIndex        =   225
            Top             =   3120
            Width           =   4095
         End
      End
   End
   Begin VB.PictureBox PicHover 
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
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   675
      TabIndex        =   13
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
         TabIndex        =   15
         Top             =   80
         Width           =   330
      End
      Begin VB.Label lblHoverClick 
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
         ForeColor       =   &H00808080&
         Height          =   570
         Left            =   60
         TabIndex        =   14
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.PictureBox PBFooter 
      BackColor       =   &H00307030&
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
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   19215
      TabIndex        =   4
      Top             =   11040
      Width           =   19215
      Begin VB.Timer TimerBeginForm 
         Interval        =   1
         Left            =   8400
         Top             =   120
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   2
         Left            =   480
         MousePointer    =   99  'Custom
         Picture         =   "Preparation.frx":17566
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mat.Req. folder"
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
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   243
         Top             =   675
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label DefaultMenuLabel 
         BackStyle       =   0  'Transparent
         Height          =   1335
         Index           =   2
         Left            =   0
         MousePointer    =   99  'Custom
         TabIndex        =   242
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
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
         TabIndex        =   10
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
         Left            =   15345
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   660
         Width           =   1230
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Preparation / Stand by"
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
         Left            =   8460
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   660
         Width           =   2220
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   0
         Left            =   9360
         MousePointer    =   99  'Custom
         Picture         =   "Preparation.frx":1A948
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   3
         Left            =   15600
         MousePointer    =   99  'Custom
         Picture         =   "Preparation.frx":1DD2A
         Top             =   120
         Width           =   480
      End
      Begin VB.Image DefaultMenu 
         Height          =   480
         Index           =   4
         Left            =   18240
         MousePointer    =   99  'Custom
         Picture         =   "Preparation.frx":2110C
         Top             =   120
         Width           =   480
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
         Index           =   0
         Left            =   8760
         TabIndex        =   7
         Top             =   -120
         Width           =   1695
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
         Index           =   3
         Left            =   14760
         TabIndex        =   6
         Top             =   -120
         Width           =   2175
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
         Index           =   4
         Left            =   17280
         TabIndex        =   5
         Top             =   -120
         Width           =   1935
      End
   End
   Begin VB.Line Line1 
      X1              =   9000
      X2              =   10200
      Y1              =   5760
      Y2              =   6240
   End
End
Attribute VB_Name = "frmPreparation"
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
Private IndexVisibleFrame As Integer

Private SelectedCode As String

Private uRecipe() As RecipeType
Private SelectedMixCode As String
Private SelectedRecipeCode As String


Private lRowHanna As Long
Private lColHanna As Long

Private lRowRecipe As Long
Private lColRecipe As Long

Private lRowMixes As Long
Private lColMixes As Long

Private lRowMaterialReq As Long
Private lColMaterialReq As Long

Private lRowCombo As Long

Private IndexRecipe As Integer
Private indexMix As Integer
Private IndexComponent   As Integer


Private ProductionWay() As ProdWay
Private uPreparation As RecipeForProduction
Private uMaterialRequisition As MaterialRequisition

Private SettingName As String
Private bImportata As Boolean
Private bIfDataPath As Boolean
Private bfrInsideMoveTop As Boolean

Private bCancelUpdate As Boolean

Private RecipeCode As String
Private CHCode As String

Private Frame3Top As Long
Private Grid1Height As Long
Private PreparationID As Long
Private userAcquisition As PrepAcquisition
Private userAcquisitionClean As PrepAcquisition

Private AcquisitionID As Long
Private AcquisitionChCode As String
Private AcquisitionWeight As Double
Private lAcquisitionRow As Long

Private ComponentID As Long
Private ComponentChCode As String
Private ComponentWeight As Double
Private lComponentRow As Long
Private bPreparationClosed As Boolean

Private NotesID As Long

Private RevisionID As Long

Private lGid6Row As Long
Private lGid6Col As Long


Private strLine As String
Private Nr As String
Private nrWeek As String



Private Sub SetColumnWidth()

Dim ctl As Control
Dim i As Integer
For Each ctl In Controls
    If TypeOf ctl Is Grid Then
            For i = 1 To ctl.Cols - 1
                ctl.Column(i).Width = (m_ControlGridColWidth / m_ControlGridColWidthOld) * ctl.Column(i).Width
          Next
    End If
Next



m_ControlGridFontSizeOld = m_ControlGridFontSize
m_ControlGridColWidthOld = m_ControlGridColWidth
m_ControlGridRowHeightOld = m_ControlGridRowHeight


 Frame3Top = Frame3(1).Top
 Grid1Height = Grid1.Height

    
End Sub


Public Function DoShow(ByVal RecCode As String, ByVal FileName As String, ByVal Preparation_ID As Long, Optional bClosed As Boolean) As Boolean

    On Error GoTo ERR_SHOW
    
    
    bCODLine = IIf(InStr(UserLine, "59"), True, False)
    
    bPreparationClosed = bClosed
    m_rc = False
    mOk
    PreparationID = Preparation_ID
    bIfDataPath = IIf(USER_PATH = USER_DATA_PATH, True, False)

    SettingName = FileName
    bImportata = IIf(FileName <> "", True, False)
    RecipeCode = RecCode
    
    
    Call SetGridRecipeRevision(Grid5)
    

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






Private Sub Check1_Click()
Dim rc As Boolean
rc = IIf(Check1.Value = 1, True, False)
Check1.ForeColor = IIf(rc, vbColorGreen, vbBlack)
uRecipe(1).bTestLot = rc
End Sub

Private Sub Form_Activate()
Me.WindowState = MainWindowState
End Sub








Private Sub Grid1_DblClick()
If CHCode = "" Or bPreparationClosed Then Exit Sub
PBContainer.Top = -(frInside(2).Top - 680)
Call FillUserRMCode(CHCode, False)

End Sub

Private Sub Grid1_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)




 lComponentRow = 0

    frCommandInside(6).Visible = False
    
    If FirstRow > 0 Then
        If Grid1.Cell(FirstRow, 7).Text <> "" Then
            lComponentRow = FirstRow - 1
            CHCode = Trim(Grid1.Cell(FirstRow, 2).Text)
        
            ComponentWeight = Trim(Grid1.Cell(FirstRow, 7).Text)
        End If
        frCommandInside(6).Visible = IIf(CHCode <> "", True, False)
        
        If Len(Trim(Grid1.Cell(FirstRow, 14).Text)) > 0 Then
            MessageInfoTime = 2000
            
            PopupMessage 2, Grid1.Cell(FirstRow, 14).Text, , , CHCode & " | Critical RM"
       
        End If
        
        
        
        
    End If
    
End Sub



Private Sub Grid2_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
AcquisitionID = 0
AcquisitionChCode = ""
AcquisitionWeight = 0
lAcquisitionRow = 0
If FirstRow > 0 Then
    lAcquisitionRow = FirstRow
    AcquisitionID = Grid2.Cell(FirstRow, 15).Text
    AcquisitionChCode = Grid2.Cell(FirstRow, 1).Text
    AcquisitionWeight = CDbl(Grid2.Cell(FirstRow, 4).Text)
    frCommandInside(11).Visible = Not (bPreparationClosed)
    
End If

End Sub

Private Sub Grid3_Click()

If lRowHanna > 0 And lColHanna = 7 Then
    Call SetLotHannaCode
End If


End Sub
Private Sub SetLotHannaCode()
  ' Q.ty to produce
Dim Recipe As String
Dim Lot As String
Dim strHannaCode As String
Dim sString As String

    With Grid3

        strHannaCode = Trim(.Cell(lRowHanna, 1).Text)
        Lot = Trim(Grid3.Cell(lRowHanna, 7).Text)
            
        Call InputLotNumber(strHannaCode, Lot)
    End With
           
End Sub


Private Function InputLotNumber(ByVal strHannaCode As String, ByRef Lot As String) As Boolean
Dim rc As Boolean
rc = True

    On Error GoTo ERR_LOT:
    
    
    If F_InputBox.DoShow("Please enter Preparation Lot number", strHannaCode & " - Preparation Lot", , , , Lot, , True, Me.Top) Then
    
    If Len(Lot) <> 4 And (Not (bCODLine) And uPreparation.Recipes(1).bIsMix) Then
        rc = False
        PopupMessage 2, "Lot must be 4 digits : es. 0001", , True, "LOT NUMBERT ERROR"
        GoTo ERR_END
    
    End If

    If (bCODLine And Not (uPreparation.Recipes(1).bIsMix)) And lRowHanna > 0 Then
        uPreparation.HannaCodes(lRowHanna).LotNumber = Lot
        Grid3.Cell(lRowHanna, 7).Text = Lot
        Exit Function
    End If
   

    If CheckPreparationLot(Lot, uPreparation.Recipes(1).Line, False, uPreparation) Then
           
           If bImportata Then bImportata = False
        
            If uPreparation.Recipes(1).bIsMix Then
            
                uRecipe(1).PreparationLotMix = Lot
            
            Else

            
                uRecipe(1).PreparationLotMix = Lot
                
                Call SetAllHannaCodeLot(Lot)

           End If
    Else
        rc = False
    End If
        

End If

ERR_END:
    On Error GoTo 0
    InputLotNumber = rc
    Exit Function
ERR_LOT:
    rc = False
    MsgBox err.Description
    Resume ERR_END


End Function

Private Function SetAllHannaCodeLot(ByVal Lot As String)
Dim i As Integer
If (bCODLine And Not (uPreparation.Recipes(1).bIsMix)) Then Exit Function
    With Grid3
        .AutoRedraw = False
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                .Cell(i, 7).Text = Lot
                .Cell(i, 7).Alignment = cellCenterCenter
                uPreparation.HannaCodes(i).LotNumber = Lot
            Next
            .Refresh
            .AutoRedraw = True
        End If
    End With
                
End Function



Private Sub Grid3_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)

lRowHanna = FirstRow
lColHanna = FirstCol

End Sub



Private Sub Grid6_Click()

Call ChangeMaterialPreparationReqQty(Grid6, lGid6Row, lGid6Col)

End Sub

Private Sub Grid6_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)
lGid6Row = FirstRow
lGid6Col = FirstCol
End Sub

Private Sub Image_Click()
frExcel_Click
End Sub







Private Sub lbChemical_Change()
lbInside(2) = lbChemical
End Sub

Private Sub lbCommandRecalculate_Click()
CommandRecalculate_Click
End Sub

Private Sub lbCritical_Change()
lbCritical.Visible = IIf(Len(lbCritical) > 0, True, False)
End Sub

Private Sub lbExcel_Click()
frExcel_Click
End Sub

Private Sub lbExpand_Click()
Grid1.ZOrder
Frame3(1).ZOrder
    If Trim(lbExpand) = "+" Then
        lbExpand = "-"
        Frame3(1).Top = 0
        Grid1.Top = Frame3(1).Top + Frame3(1).Height + 60
        Grid1.Height = Line5.y1 - Grid1.Top - 120
        
    Else
        lbExpand = "+"
        
        Frame3(1).Top = Frame3Top
        Grid1.Top = Frame3(1).Top + Frame3(1).Height + 60
        Grid1.Height = Grid1Height

    
    End If
End Sub

Private Sub TimerBeginForm_Timer()
    
Call StartUpForm

TimerBeginForm.Enabled = False

End Sub

Private Sub StartUpForm()
    
    Call InitForm
    
    Dim i As Integer
    
    
    If bfrInsideMoveTop = False Then
        For i = 3 To frInside.UBound
            frInside(i).Top = frInside(i).Top - (frInside(2).Height) * m_ControlGridRowHeight
        Next
        bfrInsideMoveTop = True
    End If
    

    
    For i = txFormulation.LBound To txFormulation.UBound
        txFormulation(i) = ""
        
    Next
    
   
    '--------------------------------------
    '
    '   Recipe importata
    '
    '--------------------------------------
    
    txFormulation(7) = MyOperatore.Name
    
    If bImportata Then
        GetFileInfo
    Else
       
    End If
    
    
    '--------------------------------------


End Sub
Private Sub InitForm()



  
    uPreparation = uPreparationClean
    
    ReDim uRecipe(0)

    PicMenu(1).Visible = bImportata
   

    
    SelectedCode = ""
    SelectedMixCode = ""
    SelectedRecipeCode = ""
    lRowHanna = 0
    lColHanna = 0
    lRowRecipe = 0
    lColRecipe = 0
    lRowMixes = 0
    lColMixes = 0
    lRowCombo = 0
    IndexRecipe = 0
    indexMix = 0
    
    lRowMaterialReq = 0
    lColMaterialReq = 0
    
    
    Dim Grid(10) As Grid
    
    Set Grid(0) = Grid1
    Set Grid(1) = Grid2
    
    Set Grid(2) = Grid3
    
    Set Grid(3) = Grid6
    'Set Grid(4) = Grid5
  
   ' Set Grid(6) = Grid7
   
   Call SetGridNotes(Grid4)
    
    Call SetAllPreparationGrid(Grid())
    Call SetColumnWidth
    
    Grid1.FrozenCols = 2
    Grid2.FrozenCols = 2
   ' Grid3.FrozenCols = 2
   ' Grid4.FrozenCols = 2
   ' Grid5.FrozenCols = 2
  
   ' Grid7.FrozenCols = 2
   '
    
    

End Sub
Private Sub Form_Load()


    PBContainer.Top = 0
    PBContainer.Left = 0
    PBContainer.Width = Me.Width
  
    
    
    Dim i As Integer
    If Screen.Width - Me.Width > 1000 And bFullScreen Then
        Me.WindowState = 2
    
    End If


    For i = PBContainerViewport.LBound To PBContainerViewport.UBound
        PBContainerViewport(i).Move 0, PBTitle.Height, Me.ScaleWidth, Me.ScaleHeight - PBTitle.Height
    Next
  
    RSBottom PicHover, Me, -1350
    RSRight PicHover, Me, -450
 

    PBContainerViewport(0).ZOrder
    PBFooter.ZOrder
    
    Select Case iLotNumberType
    
    Case 0
        
    Case 1
        lbFormulation(9) = "Even Lot"
    Case 2
        lbFormulation(9) = "Odd Lot"
    End Select
    
    
    
End Sub

Private Sub Form_Resize()




    
    lbWait.Left = Me.Width / 2 - lbWait.Width / 2
    PBTitle.Width = Me.Width
    PBFooter.Top = Me.ScaleHeight - PBFooter.Height
    PBFooter.Width = Me.Width
 
    
    'Resize the container (needed to show the full bottom box on maximized state)
    'First resize our container
   
    
    
    ResizeControls

    SetColumnWidth
 
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmPreparation = Nothing
End Sub







Private Sub txFormulation_Change(Index As Integer)
Dim rc As Boolean
rc = IIf(Len(txFormulation(Index)) > 0, True, False)

    Select Case Index
        Case 1, 2, 6, 8
            txFormulation(Index).BackColor = IIf(rc, vbWhite, vbRed)
        Case 9
            If uRecipe(1).bTestLot = False Then
                If Len(txFormulation(Index)) <> 4 Then
                    
                    txFormulation(Index) = Right(txFormulation(Index), 4)
                
                End If
            End If
    
    End Select

End Sub

Private Sub txFormulation_Click(Index As Integer)
Dim Answer As String
Dim Selected As String
Dim sString As String
Dim bNumber As Boolean

Selected = "Preparation"
Answer = txFormulation(Index)
sString = lbFormulation(Index)

bNumber = IIf(Index = 2, True, False)


If bPreparationClosed Then Exit Sub

' lot!

If Index = 9 Then

        SetFormulationLot

    Exit Sub
End If

If Index = 1 Then If Answer = "" Then Answer = FormatDataLAT(Now())
If Index = 6 Then If Answer = "" Then Answer = PreparationWeek(Now())
If Index = 8 Then
    ' exp date
    If txFormulation(1) = "" Then
        PopupMessage 2, "Please enter Preparation Date first...", , True, "Exp Date"
        Exit Sub
    Else
        SetuPreparationExpDate
        Answer = uPreparation.ExpDate
    End If
End If
If Index = 7 Then
    
    If frmLogin.DoShow Then
        txFormulation(7) = MyOperatore.Name
        Exit Sub
    Else
    
        Exit Sub
    End If
    
End If

        
        If F_InputBox.DoShow(sString, Selected, , , , Answer, , bNumber, Me.Top) Then
        
            txFormulation(Index) = Answer
            
            Select Case Index
                Case 1
                    ' isdate?
                    If IsDate(Answer) Then
                         txFormulation(Index) = FormatDataLAT(Answer)
                         uPreparation.PreparationDate = Answer
                         SetuPreparationExpDate
                         
                    Else
                        PopupMessage 2, "Please enter a valid Date...", , True
                    End If
                Case 6
                    uPreparation.PrepWeek = Answer
                Case 2
                    uPreparation.numPrepWeek = Answer
                Case 8
                    ' exp date
                    If IsDate(Answer) Then
                         txFormulation(Index) = FormatDataExp(Answer)
                         uPreparation.ExpDate = txFormulation(Index)
                    Else
                        PopupMessage 2, "Please enter a valid Exp Date ( MM/YYYY ) ...", , True
                    End If
                Case 9
                    ' lot

                       
                
            End Select
        End If
        
        




End Sub





Private Sub txFormulation_LostFocus(Index As Integer)
    
    Select Case Index
        
        Case 1, 2, 3
           
    
    End Select
    

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

m_ControlGridFontSizeOld = 1
m_ControlGridColWidthOld = 1
m_ControlGridRowHeightOld = 1

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
        ElseIf TypeOf ctl Is Image Then
            ctl.Left = (x_scale * .Left) + IIf(x_scale = 1, 0, (x_scale - 1) * 200)
            ctl.Top = y_scale * .Top
        ElseIf TypeOf ctl Is ucScrollAdd Then
        ElseIf TypeOf ctl Is Grid Then
           ctl.Left = x_scale * .Left
            ctl.Top = y_scale * .Top
            ctl.Width = x_scale * .Width
            ctl.Height = y_scale * .Height

               ' ctl.DefaultFont.Size = 12 * m_ControlGridFontSize
               ' ctl.DefaultRowHeight = 30 * m_ControlGridRowHeight
           
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






Private Sub DefaultMenuLabel_Click(Index As Integer)
DefaultMenu_Click Index
End Sub



Private Sub DefaultMenu_Click(Index As Integer)
Dim MyIndex As Integer
Select Case Index
    Case 0
        If F_MsgBox.DoShow("Quit Preparation?", "Exit Preparation / Stand By") Then
            
            If Grid2.Rows > 1 And Not (bPreparationClosed) Then
                
                
                If F_MsgBox.DoShow("Save Preparation?") Then
                    frCommandInside_Click 4
                Else
                End If
            
            End If
            
            Unload Me
        End If
    Case 2
   ' Debug.Print PathPrepRequisition
            ApriIlReportFolder (USER_DOCUMENTI & PathPrepRequisition)
    Case 3
        ' Previous
         If IndexVisibleFrame > 1 Then
            MyIndex = IndexVisibleFrame - 1
            
            
            PBContainer.Top = -(frInside(MyIndex).Top - 680)
        Else
            PBContainer.Top = 0
         End If
    
    
    
    Case 4
        ' forward
        If IndexVisibleFrame < frInside.UBound Then
            MyIndex = IndexVisibleFrame + 1
            
            PBContainer.Top = -(frInside(MyIndex).Top - 680)
        Else
            PBContainer.Top = 0
        End If
          
End Select
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

Private Sub PicMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To PicMenu.Count - 1
 
    If i = IndexProcedura Then
        ' lascialo stare...
    ElseIf i = Index Then
        PicMenu(i).BackColor = &H206020
    Else
        PicMenu(i).BackColor = &H105010
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
        PicMenu(i).BackColor = &H307030
    Else
        PicMenu(i).BackColor = &H105010
    End If
Next
blTable = Label2(Index)
IndexProcedura = Index

PBContainerViewport(Index).ZOrder
PBContainerViewport(Index).Visible = True

DefaultMenu(2).Visible = IIf(IndexProcedura = 1, True, False)
DefaultMenuLabel(2).Visible = DefaultMenu(2).Visible
Lab(5).Visible = DefaultMenu(2).Visible

Select Case IndexProcedura
    Case 0
        
    Case 1
      
          txDocument(1) = MyOperatore.Name
          txDocument(3) = txFormulation(1)
          txDocument(4) = txFormulation(4)
          txDocument(0) = GetMaterialRequisitionNumber
          txDocument(2) = GetLineNumber(uPreparation.Recipes(1).Line)
          txDocument(5) = GetSetting(App.Title, "PREP_" & "MaterialRequisition", "Dep.Manager", "Kis Laszlo")
    
        
End Select

PBFooter.ZOrder


End Function
Private Function GetMaterialRequisitionNumber() As String

On Error GoTo ERR_GET:

With uPreparation

    If txFormulation(1) = "" Then Exit Function
    'nrWeek = Week(txFormulation(6))
    If IsDate(txFormulation(1)) Then
        nrWeek = Week(txFormulation(1))
    Else
        nrWeek = txFormulation(6)
    End If
    nrWeek = Format(nrWeek, "00")
    strLine = GetLineNumber(.Recipes(1).Line)
    
    Nr = GetSetting(App.Title, "PREP_" & strLine, nrWeek, "00") + 1
    Nr = Format(Nr, "00")
ERR_END:
    On Error GoTo 0
    GetMaterialRequisitionNumber = strLine & nrWeek & Nr
    
End With

    Exit Function
ERR_GET:
    MsgBox err.Description
    Resume ERR_END:

End Function

Private Function GetLineNumber(ByVal Line As String) As String
GetLineNumber = Mid(Line, 2, 2)
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
Dim rc As Boolean
Dim UserChCode As String

    txQRCode.Visible = False


    bCancelUpdate = False
        
    Select Case Index
        Case 0
            'add chemical
            frCommandInside_Click 5
            OpenComponentDatabase False
        Case 1
            ' delete component
            Call DeleteComponent
        Case 2
            ' scroll top
            
            PBContainer.Top = 0
        Case 3
            ' scroll to acquisitions
            frInside(1).ZOrder
            PBContainer.Top = -(frInside(1).Top - 680)
        Case 4
            Call SavePreparation
        Case 5
            ' add acquisition
            Call AddAcquisition
        Case 6
            ' open product classification
            Call OpenProductCalssification(CHCode, 1)
        Case 7
            ' exit acquisition
            Call SetAcquisition(False)
            
        Case 8
            ' save acquisition
             Call SetAcquisition(True)
        Case 9
            ' Acquisition : open product classification
            Call OpenProductCalssification(txAcquisition(0), 1)
        Case 10
            ' add acquisition
            frCommandInside_Click 5
        Case 11
            ' delete acquisition
            Call DeleteAcquisition
        Case 12
            ' back to acquisition
            PBContainer.Top = -(frInside(2).Top - 680)
            frCommandInside(12).Visible = False
        Case 13
            ' SCANN BARCODE
            Call ScanQRCode
        Case 14
            ' goto hanna codes table
            
            Call SetFormulationLot
  
        Case 15
            Call AddAcquisition
        Case 16
            ' goto preparation...
             PBContainer.Top = 0
            frInside(3).Visible = False
        Case 17
            Call ClearPrepNotesForm
        Case 18
            If SettingName <> "" Then
                AddcmbNotes
                lbInside(3).ForeColor = vbWhite
                Call GetPreparationNotes(Grid4, SettingName)
                 frExcel2.Visible = IIf(Grid4.Rows > 1, True, False)
                Frame6.Visible = IIf(Grid4.Rows > 1, False, True)
            
                Call ClearPrepNotesForm
                frInside(4).Visible = True
                PBContainer.Top = -(frInside(4).Top - 680)
            Else
                PopupMessage 2, "Please save Preparation first..."
            End If
        Case 19
          
            Call ClearRevisionForm
        Case 20
              Call OpenRevisionHistory
        Case 21
            frInside(5).Visible = False
            PBContainer.Top = 0
        Case 22
            frInside(4).Visible = False
            PBContainer.Top = 0
        Case 23
             ' Debug.Print PathPrepRequisition
            ApriIlReportFolder (USER_DOCUMENTI & PathPrepRequisition)
            
        Case 24
            frCommandInside_Click 25
        Case 25
             ' material requisition ALL RECIPE
            Call SavePreparation
            Call SetMaterialRequisitionComponents
        Case 26
           ' MATERIAL REQUISITION : delete record
            Dim SelectedComponent As String
            SelectedComponent = Grid6.Cell(lRowMaterialReq, 1).Text
            If SelectedComponent <> "" Then
                If F_MsgBox.DoShow("Warning : Delete Component from Table?", SelectedComponent) Then
                    Call MaterialRequisitionDeleteRecord(Grid6)
                End If
            
            End If
            
    End Select
End Sub

Private Function SetFormulationLot()
Dim Lot As String
Dim rc As Boolean

If uRecipe(1).bTestLot Then
    Lot = "Test-" & FormatDataLAT(Now())
    If F_InputBox.DoShow("Please enter Test Lot number", "TEST LOT", , , , Lot, , , Me.Top) Then
        Lot = Format$(Lot, "0000")
        uRecipe(1).PreparationLotMix = Lot
    End If

Else

        If uPreparation.Recipes(1).bIsMix Then
          
          Lot = txFormulation(9)
          rc = InputLotNumber(uPreparation.Recipes(1).Code, Lot)
          If rc = False Then Exit Function
         
          
        Else


          Lot = txFormulation(9)
          rc = InputLotNumber(uPreparation.Recipes(1).Code, Lot)
          If rc = False Then Exit Function
          'txFormulation(9) = Lot
          
         
         If (bCODLine And Not (uPreparation.Recipes(1).bIsMix)) Then
            frInside(3).Left = frInside(1).Left
            frInside(3).Top = frInside(1).Top
            frInside(3).ZOrder
            frInside(3).Visible = True
            PBContainer.Top = -(frInside(3).Top - 680)
        End If
            
        End If
        
        
        If IsNumeric(Lot) Then Lot = LotNumberControl(Lot)
        Lot = Format$(Lot, "0000")
End If
        uRecipe(1).PreparationLotMix = Lot
        txFormulation(9) = Lot
        txFormulation(9).BackColor = vbWhite
        txFormulation(9).ForeColor = vbBlack
        
            
            
End Function
Private Sub frInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


Dim i As Integer
    For i = 0 To frCommandInside.UBound

            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 4 Or i = 13 Or i = 8 Then
                frCommandInside(i).BackColor = &H8000&
            End If

    
    Next
 
 
End Sub

Private Sub frCommandInside_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
IndexDashCommInside = Index
Dim i As Integer
    For i = 0 To frCommandInside.UBound
        If i = Index Then
            ' quando ci passo sopra....
            frCommandInside(i).BackColor = &H846623
            lbCommandInside(i).ForeColor = vbWhite
            If i = 4 Or i = 13 Or i = 8 Then
                frCommandInside(i).BackColor = &H20A020
            End If
            
        Else
            frCommandInside(i).BackColor = &H644603
            lbCommandInside(i).ForeColor = &HE0E0E0
             If i = 4 Or i = 13 Or i = 8 Then
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
    
For i = 0 To PicMenu.UBound
    If i = IndexProcedura Then
    Else
        PicMenu(i).BackColor = &H105010
    End If
Next

End Sub


'-----------------------------------------------------------------------------------------------
'   SettingName
'-----------------------------------------------------------------------------------------------

Private Sub SetSettingName()
' LINE+DATERECIPE+PREPARATIONWEEK+PLANNEDPREPARATION
SettingName = FormatNomeFile(Trim(uRecipe(1).Code) & "." & Trim(uRecipe(1).Line) & "." & txFormulation(1) & "." & txFormulation(2) & "." & txFormulation(3)) & "." & USER_ESTENSIONE_RFP

End Sub


Private Function ChecktxFormulation() As Boolean
Dim rc As Boolean
Dim i As Integer
    rc = True
    For i = txFormulation.LBound To txFormulation.UBound - 1
        If Len(txFormulation(i)) = 0 Then
            rc = False
            MessageInfoTime = 1500
            PopupMessage 2, "Please Enter field : " & lbFormulation(i), , True, "Preparation"
            txFormulation(i).SetFocus
            rc = False
            Exit For
        End If
    Next
    ChecktxFormulation = rc
End Function


'-----------------------------------------------------------------------------------------------
'
'
'                                   GetReceiptFromFile

'
'-----------------------------------------------------------------------------------------------


Private Sub GetFileInfo()
Dim rc As Boolean
 rc = GetReceiptFromFile
 
 bImportata = rc
 
End Sub

Private Function GetReceiptFromFile() As Boolean
Dim i As Integer
Dim rc As Boolean

On Error GoTo ERR_GET:
   
rc = True

    lbWait.Visible = True
    lbLine.Visible = False
    
    ReDim uRecipe(0)
    
    
    uPreparation = uPreparationClean
    
    
    If SettingName = "" Then
            MessageInfoTime = 2000
            PopupMessage 2, "Warning : File non found! ", , True, "Preparation"
            rc = False
            GoTo ERR_END
    
    End If
    
    If FileExists(USER_PREPARATION_PATH & SettingName) Then
        USER_PATH = USER_PREPARATION_PATH
    ElseIf FileExists(USER_PREPARATION_PATH & "Data\" & SettingName) Then
        USER_PATH = USER_PREPARATION_PATH & "Data\"
    Else
        rc = False
        PopupMessage 2, "No file Preparation found...", , True, SettingName
        GoTo ERR_END
        
    End If
    
    
    If GetSettingData(SettingName, "iPreparation", "bOpen", True) Then

        Else
       
            blTable.Visible = False
            blTable = "Preparation : Closed"
            'PicMenu(1).Visible = False
            bPreparationClosed = True

    End If
    
    
    
    Call SetClosedPreparation(bPreparationClosed)
    Call PreparationGetSetting(uPreparation, SettingName, RecipeCode)
     
    
    With uPreparation

        uRecipe = .Recipes

        txFormulation(1) = IIf(IsNull(.PreparationDate), FormatDataLAT(Now()), .PreparationDate)
        txFormulation(5) = .Note
        txFormulation(3) = .PlannedPrepWeek
        txFormulation(4) = .PlanningReference
        txFormulation(2) = .numPrepWeek
        txFormulation(0) = .RecipeBy
        txFormulation(6) = .PrepWeek
        txFormulation(8) = .ExpDate
        Check1.Value = IIf(uRecipe(1).bTestLot, 1, 0)
        
        
        
        txFormulation(9).BackColor = vbWhite
        txFormulation(9).ForeColor = vbBlack
        If .Recipes(1).PreparationLotMix = "" Or .Recipes(1).PreparationLotMix = "0000" Then
            If uRecipe(1).bTestLot = False Then
                Call CheckPreparationLot(.Recipes(1).PreparationLotMix, .Recipes(1).Line, True, uPreparation)
                
                
                If .Recipes(1).bIsMix Then
                
                Else
                   
                   For i = 1 To UBound(.HannaCodes)
                    If (bCODLine And Not (uPreparation.Recipes(1).bIsMix)) Then
                        ' li abbiamo già impostati
                    Else
                      .HannaCodes(i).LotNumber = .Recipes(1).PreparationLotMix
                    End If
                    
                   Next
                   
                   End If
                
                txFormulation(9).BackColor = vbColorRosaTabella
             End If
             
             
        End If
        
        
        
        If uRecipe(1).bTestLot Then
            ' se cambio ba lotto normale salvato a TEST allora devo risalvare tutto
             If .Recipes(1).bIsMix Then
             
             Else
                If (bCODLine And Not (uPreparation.Recipes(1).bIsMix)) Then
                Else
                    For i = 1 To UBound(.HannaCodes)
                       .HannaCodes(i).LotNumber = .Recipes(1).PreparationLotMix
                    Next
                End If
                
            End If
             
            txFormulation(9).BackColor = vbColorGreen
            txFormulation(9).ForeColor = vbWhite
             
        End If
        
       ' If (bCODLine And Not (uPreparation.Recipes(1).bIsMix)) Then
            ' per COD si è deciso di mettere il Lotto del primo HannaCode
           ' If IsNumeric(.HannaCodes(1).LotNumber) Then
              '  txFormulation(9) = .HannaCodes(1).LotNumber
           ' End If
       ' Else
        
            txFormulation(9) = .Recipes(1).PreparationLotMix
        'End If
        
        If .RecipeCount = 0 Then
            MessageInfoTime = 2000
            PopupMessage 2, "Warning : No recipe found! ", , True, "Preparation"
            rc = False
            GoTo ERR_END
        End If
        
        
        lbInside(1) = uRecipe(1).Code & " | " & uRecipe(1).Description
        
    End With
    
    
    Call FillGridPreparationFromFile(Grid1, uPreparation, 1, PreparationID)
    Call FillGridPreparationFromFile(Grid2, uPreparation, 2, PreparationID)
    Call FillGridPreparationFromFile(Grid3, uPreparation, 3, PreparationID)
    

    uRecipe(1).ActualWeight = uPreparation.Recipes(1).ActualWeight

    lbLine.Visible = True
    lbLine = lbInside(1) ' uRecipe(1).Line
    
    blTable.Visible = True

    If uRecipe(1).bUmMassa Then
    Else
        lbRecipeDensity = "Density : " & uRecipe(1).Density
        lbRecipeDensity.Visible = True
    End If
    
    If uPreparation.Recipes(1).TotalWeightKg = 0 Then
        PopupMessage 2, "Error : Recipe without Total Weight to Produce", , True
        Unload Me
    End If
    
    lbFormulation(9).Visible = True ' IIf(uPreparation.Recipes(1).bIsMix, True, False)
    txFormulation(9).Visible = True ' IIf(uPreparation.Recipes(1).bIsMix, True, False)
    
    frCommandInside(14).Visible = True ' IIf(uPreparation.Recipes(1).bIsMix, False, True)
    lbInside(0) = IIf(uPreparation.Recipes(1).bIsMix, "Mix Lot", "Hanna Code Lot")
ERR_END:

    On Error GoTo 0
     uRecipe = uPreparation.Recipes
    lbWait.Visible = False
    lbLine.Visible = True
    GetReceiptFromFile = rc
    Exit Function
ERR_GET:
    rc = False
    MsgBox err.Description
    Resume Next
    
End Function


Private Function AddAcquisition()
  
    ClearAcquisition
    frInside(2).ZOrder
    frInside(2).Visible = True
    txAcquisition(0).SetFocus
    PBContainer.Top = -(frInside(2).Top - 680)
End Function

Private Sub txAcquisition_Change(Index As Integer)
Dim rc As Boolean
    rc = IIf(Len(txAcquisition(Index)) > 0, True, False)
    
    txAcquisition(Index).BackColor = IIf(rc, vbWhite, &HF0F0F0)
    
    Select Case Index
        Case 0, 1
            frCommandInside(9).Visible = rc
            lbChemical = txAcquisition(0) & " | " & txAcquisition(1)
            
            lbCritical = GetCriticalRM(txAcquisition(0))
            lbNote = GetNoteRM(uRecipe(1).Code, txAcquisition(0))
            lbNote.Visible = IIf(lbNote <> "", True, False)
            
            If rc Then
                txAcquisition(3).BackColor = IIf(txAcquisition(3) <> "", vbWhite, &HF0F0F0)
                txAcquisition(4).BackColor = IIf(txAcquisition(4) <> "", vbWhite, &HF0F0F0)
                txAcquisition(5).BackColor = IIf(txAcquisition(5) <> "", vbWhite, &HF0F0F0)
            End If
          
        Case 13
            ' peso acquisito
            
            CommandRecalculate.Visible = False
            lbRecalculate.Visible = False
            PicTolerance.Visible = False
            
            If rc And IsNumeric(txAcquisition(Index)) Then
                Call SetVarianceAcquisition(IndexComponent)
               ' txAcquisition(17).SetFocus
            End If

            If Len(txAcquisition(0)) > 0 Then
                frCommandInside(8).Visible = rc
            Else
                 frCommandInside(8).Visible = False
            End If
            
            
    End Select
End Sub



Private Sub OpenComponentDatabase(ByVal rc As Boolean)
Dim userCode As String

    userCode = txFormulation(0)
    rc = FormChemicalRM.DoShow(userCode, , IIf(rc, RecipeCode, ""))
    If rc Then
        
        If userCode <> "" Then
            Call FillUserRMCode(userCode, False)
        End If
    
    End If
            
End Sub



Private Sub txAcquisition_Click(Index As Integer)
Dim userCode As String
Dim Answer As String
Dim Selected As String
Dim bNumber As Boolean
Dim sString As String
Dim rc As Boolean


    Selected = lbAcquisition(Index) ' "Preparation"
    Answer = txAcquisition(Index)
    sString = "Please Enter " & lbAcquisition(Index)
    
    bNumber = IIf(Index = 11 Or Index = 13, True, False)
    

    Select Case Index
        Case 0
            ' importa RMCode
            
            OpenComponentDatabase True
            Exit Sub

    End Select
    
    
    If txAcquisition(Index).Locked Then Exit Sub
    If txAcquisition(0) = "" Then Exit Sub
    
    If F_InputBox.DoShow(sString, Selected, , , , Answer, , bNumber, Me.Top) Then
    
        txAcquisition(Index) = Answer
        
        Select Case Index
            Case 6
                ' isdate?
                If IsDate(Answer) Then
                     txAcquisition(Index) = FormatDataLAT(Answer)
                Else
                    PopupMessage 2, "Please enter a valid Date...", , True
                End If
        End Select
    End If
    
    
    
    
End Sub

Private Function CheckPreparationDetail()
Dim rc As Boolean

    rc = True
    
    If txFormulation(6) = "" Then
        rc = False
        txFormulation(6).BackColor = vbRed
    End If
    If txFormulation(2) = "" Then
        rc = False
        txFormulation(2).BackColor = vbRed
    End If
    If txFormulation(1) = "" Then
        rc = False
        txFormulation(1).BackColor = vbRed
    End If

    CheckPreparationDetail = rc


End Function
Private Function SetAcquisition(ByVal rc As Boolean)
Dim mrc As Boolean
Dim i As Integer

   

    If CheckPreparationDetail = False Then
        frCommandInside(12).Visible = True
        PBContainer.Top = 0
        PopupMessage 2, "Please fill all Preparation Details first..."
    
        Exit Function
    End If
    

    
    lbCritical = ""
    
    If rc Then
    
    
        For i = 3 To 5
            If txAcquisition(i) = "" Then
                PopupMessage 2, "Please enter " & lbAcquisition(i) & "..", , True, lbAcquisition(i)
                txAcquisition_Click i
                Exit Function
            End If
        Next
        
        
        
        mrc = SaveAcquisition
        If mrc Then
            PopupMessage 2, "Acquisition done..."
        Else
            PopupMessage 2, "Warning : Acquisition not Saved...", , True, txAcquisition(0)
        End If
    Else
       '
    End If
    PBContainer.Top = 0
    frInside(2).Visible = False
    
End Function

Private Function FillUserRMCode(ByVal userCode As String, ByVal bFromBarcode As Boolean) As Boolean
Dim rc As Boolean
Dim Manufacurer As String
Dim ManufacturerCode As String

    On Error GoTo ERR_FILL:
    
    If bFromBarcode = False Then Call ClearAcquisition
    
    frInside(2).Visible = True
    
    rc = True
    
    
    With dbTabRMxRecipe
        .filter = ""
        .filter = "CHCode='" & userCode & "' and RecipeCode='" & RecipeCode & "'"
        
        If .EOF Then
        
            ' attenzione non fa parte della ricetta!!!!!
            If IfChemicalInRecipeType(userCode, uRecipe(1), IndexComponent) Then
            
                
               If bFromBarcode = False Then txAcquisition(0) = uRecipe(1).RmxRecipe(IndexComponent).CHCode
                If bFromBarcode = False Then txAcquisition(1) = uRecipe(1).RmxRecipe(IndexComponent).Description
                txAcquisition(2) = uRecipe(1).RmxRecipe(IndexComponent).Cas
                uRecipe(1).RmxRecipe(IndexComponent).TolerancePerc = 1
           
                GoTo cont
            Else
            
                If Warning.DoShow("Warning : " & userCode & " is not a Recipe Component." & vbCrLf & "Add component Anyway?", userCode, , "Proceed", "Exit") Then
                    
                    Call GetComponentSpecifics(userCode)
                    IndexComponent = 999
                Else
                
                    FillUserRMCode = False
                    Exit Function
                
                End If
            End If
        Else
        
            Call IfChemicalInRecipeType(userCode, uRecipe(1), IndexComponent)

            If bFromBarcode = False Then txAcquisition(0) = IIf(IsNull(Trim(!CHCode)), "", Trim(!CHCode))
            If bFromBarcode = False Then txAcquisition(1) = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
            txAcquisition(2) = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
           
            uRecipe(1).RmxRecipe(IndexComponent).TolerancePerc = CheckDot(IIf(IsNull(Trim(!TolerancePerc)), 1, Trim(!TolerancePerc)))
            If uRecipe(1).RmxRecipe(IndexComponent).TolerancePerc = 0 Then uRecipe(1).RmxRecipe(IndexComponent).TolerancePerc = 1
            
cont:
            Call SetTxAcquisition(userCode)
            
           
            
            
        End If
    
    
    End With
ERR_END:
    On Error GoTo 0
    FillUserRMCode = rc
    Exit Function
ERR_FILL:
    rc = False
    PopupMessage 2, err.Description
    Resume ERR_END
End Function

Private Function SetVarianceAcquisition(ByVal i As Integer)

Dim Variance        As Double
Dim VariancePerc    As Double
Dim RealWeight      As Double
Dim ActualWeight    As Double
Dim bRecalculate    As Boolean
Dim bCorrection     As Boolean
    
    If i = 999 Then
    
            If txAcquisition(13) <> "" And IsNumeric(txAcquisition(13)) Then
                ActualWeight = CDbl(txAcquisition(13))
            Else
                ActualWeight = 0
            End If
            
            RealWeight = ActualWeight
            
            txAcquisition(11) = ""
            txAcquisition(12) = ""
            txAcquisition(13) = Trim(PadString(txAcquisition(13)))
            txAcquisition(14) = FormatNumber(RealWeight) & " g  "
            txAcquisition(16) = ""
            txAcquisition(15) = ""
            
            PicTolerance.Visible = False
            
            
    Else
    
        
        With uRecipe(1).RmxRecipe(i)
            
            '.RmxRecipe(i).ActualTheoreticalWeight = .RmxRecipe(i).TheoreticalWeight - .RmxRecipe(i).RealWeight
            
            If txAcquisition(13) <> "" Then
                ActualWeight = CDbl(txAcquisition(13))
            Else
                ActualWeight = 0
            End If
            
            RealWeight = .RealWeight + ActualWeight
            txAcquisition(14) = PadString(RealWeight) & " g  "
            
            
            Variance = RealWeight - .TheoreticalWeight
            VariancePerc = (Variance / .TheoreticalWeight) * 100
                
            txAcquisition(11) = PadString(Variance) & " g  "
            txAcquisition(12) = FormatNumber(VariancePerc) & " %  "
            txAcquisition(13) = Trim(PadString(txAcquisition(13)))
            txAcquisition(14) = FormatNumber(RealWeight) & " g  "
            txAcquisition(16) = FormatNumber(.TolerancePerc, 2) & " %  "
            txAcquisition(15) = FormatNumber(.TheoreticalWeight * (.TolerancePerc / 100), 2) & " g  "
            
            
            
            If ActualWeight > 0 And .bAddedInPreparation = False Then
            
                Dim MyColor As OLE_COLOR
                
                MyColor = ColorTolerance(Variance, .TheoreticalWeight * (.TolerancePerc / 100), bRecalculate, bCorrection)
                
                    PicTolerance.BackColor = MyColor
                    PicTolerance.Visible = True
                    
                    If bCorrection Then
                        uPreparation.bCorrection = True
                    End If
                
                
                
                If RealWeight > .TheoreticalWeight Then
                    lbRecalculate.Visible = bRecalculate
                    CommandRecalculate.Visible = bRecalculate
                Else
                    lbRecalculate.Visible = False
                    CommandRecalculate.Visible = False
                    bRecalculate = False
                End If
            End If
        
        End With
    End If


End Function



Private Sub SetTxAcquisition(ByVal userCode As String)

Dim i As Integer
Dim Variance As Double
Dim VariancePerc As Double
Dim Manufacurer As String
Dim ManufacturerCode As String
With uRecipe(1)
    
    For i = 0 To .RmxRecipeCount
        If .RmxRecipe(i).CHCode = userCode Then
            IndexComponent = i
            ' .RmxRecipe(i).ActualTheoreticalWeight = .RmxRecipe(i).TheoreticalWeight - .RmxRecipe(i).RealWeight

            txAcquisition(10) = PadString(.RmxRecipe(i).TheoreticalWeight) & " g  "
         
            
             Call SetVarianceAcquisition(IndexComponent)
             
                 ' se ho una mix acquisisco Manufacturer.... da TabRawMaterial
            'If .RmxRecipe(IndexComponent).bMix Then
            
                Call GetRawMaterialManufacturer(.RmxRecipe(IndexComponent).CHCode, Manufacurer, ManufacturerCode)
                txAcquisition(3) = Manufacurer
                txAcquisition(4) = ManufacturerCode
            
            'End If
    
            Exit For
        Else
            IndexComponent = 999
        End If
    Next
End With

userAcquisition.bFromBarcode = False
userAcquisition.bRecipeComponent = True
Frame3(2).BackColor = &H886010
lbInside(2) = "Acquisition Details"
lbInside(2).ForeColor = vbWhite
            
End Sub

Private Sub GetComponentSpecifics(ByVal userCode As String)


    With dbTabRawMaterial
        .filter = ""
        .filter = "Code='" & userCode & "'"
        
        If .EOF Then

                Exit Sub
          
        Else
            
            txAcquisition(0) = IIf(IsNull(Trim(!Code)), "", Trim(!Code))
            txAcquisition(1) = IIf(IsNull(Trim(!Description)), "", Trim(!Description))
            txAcquisition(2) = IIf(IsNull(Trim(!Cas)), "", Trim(!Cas))
            
            userAcquisition.bFromBarcode = False
            userAcquisition.bRecipeComponent = False
            Frame3(2).BackColor = &HFFFF&
            lbInside(2) = "Acquisition Details : Correction"
            lbInside(2).ForeColor = &H105010
        End If
    
    
    End With
End Sub

Private Function ClearAcquisition()
Dim i As Integer
    
    For i = txAcquisition.LBound To txAcquisition.UBound
        txAcquisition(i) = ""
        txAcquisition(i).BackColor = &HF0F0F0
    Next
    lbNote = ""
    lbNote.Visible = False
    lbChemical = ""
    lbInside(2) = "Acquisition Details"
    PicTolerance.Visible = False
    lbRecalculate.Visible = False
    CommandRecalculate.Visible = False
                           
    Frame3(2).BackColor = &H886010
    lbInside(2) = "Acquisition Details"
    lbInside(2).ForeColor = vbWhite
    
    userAcquisition = userAcquisitionClean
    AcquisitionID = 0
    AcquisitionChCode = ""
    AcquisitionWeight = 0
    frCommandInside(11).Visible = False
End Function


Private Function SaveAcquisition() As Boolean
Dim rc As Boolean
    rc = True
    
    '-------------------------------------------
    ' salvo l'acquisizione in userAcquisition
    ' se non avevo il componente lo aggiungo
    ' in urecipe(1).RmxRecipe(IndexComponent)
    '-------------------------------------------
    
    Call SetNewUserAcquisition
    
    '-------------------------------------------
    ' salvo l'acquisizione in TabAcquisition
    '-------------------------------------------
    
    Call SaveAcquisitionInTabAcquisition
    '-------------------------------------------
    ' Aggiungo Row in Grid2 ' acquisitions
    '-------------------------------------------
    With Grid2
        .AutoRedraw = False
        
        Call AddNewRowInAcquisition(Grid2, userAcquisition)
        
        .Refresh
        .AutoRedraw = True
    End With
    '-------------------------------------------
    ' Salvo su file :
    ' cancello e risalvo? o aggiungo e basta?
    '
    '-------------------------------------------
    Dim MaxCount As Integer
    MaxCount = uRecipe(1).AcquisitionCount
    ReDim Preserve uRecipe(1).Acquisitions(MaxCount)
    
    uRecipe(1).Acquisitions(MaxCount) = userAcquisition
    
    '-------------------------------------------
    ' ricarico la Gird1
    '-------------------------------------------
    Call FillGridPreparationFromFile(Grid1, uPreparation, 1, PreparationID)
    
    uRecipe(1).ActualWeight = uPreparation.Recipes(1).ActualWeight
    
    SaveAcquisition = rc
End Function

Private Sub SetNewUserAcquisition()
Dim AcquisitionsCount As Integer
    With userAcquisition
        .ActualWeight = txAcquisition(13)
        .PrepBarcode.Code = txAcquisition(0)
        .PrepBarcode.ChemicalName = txAcquisition(1)
        .PrepBarcode.Cas = txAcquisition(2)
        .PrepBarcode.Manufacturer = txAcquisition(3)
        .PrepBarcode.ManufacturerCode = txAcquisition(4)
        .PrepBarcode.ManufacturerLot = txAcquisition(5)
        .PrepBarcode.DeliveryDate = txAcquisition(6)
        .PrepBarcode.QtyDelivered = txAcquisition(7)
        .PrepBarcode.WeekDelPackageNumber = txAcquisition(8)
        .PrepBarcode.Package = txAcquisition(9)
        .ExpDate = txAcquisition(18)
        .Note = txAcquisition(17)
        .AcquisitionTime = Now()
        .Operator = MyOperatore.Name
      
         If .bRecipeComponent Then
             uRecipe(1).RmxRecipe(IndexComponent).RealWeight = uRecipe(1).RmxRecipe(IndexComponent).RealWeight + .ActualWeight
        Else
            ' aggiungo il componente alla ricetta
            Call AddNewComponentToRecipe(uRecipe(1).RmxRecipe(), uRecipe(1).RmxRecipeCount, txAcquisition(0), RecipeCode, txAcquisition(13))
            IndexComponent = uRecipe(1).RmxRecipeCount
        End If
    End With
    
    
    AcquisitionsCount = uRecipe(1).AcquisitionCount
    AcquisitionsCount = AcquisitionsCount + 1
    uRecipe(1).AcquisitionCount = AcquisitionsCount
    
    'ReDim Preserve uRecipe(1).Acquisitions(AcquisitionsCount)
    
    'uRecipe(1).Acquisitions(AcquisitionsCount) = userAcquisition
    'uRecipe(1).RmxRecipe(IndexComponent).RealWeight = uRecipe(1).RmxRecipe(IndexComponent).RealWeight + userAcquisition.ActualWeight
    userAcquisition.Index = AcquisitionsCount
    
       
    If uRecipe(1).RmxRecipe(IndexComponent).bAddedInPreparation Then
        uRecipe(1).RmxRecipe(IndexComponent).TheoreticalWeight = uRecipe(1).RmxRecipe(IndexComponent).RealWeight
    End If
    
    
    
    uPreparation.Recipes(1) = uRecipe(1)
    
End Sub



' recalculation

Private Sub CommandRecalculate_Click()
Dim NewTotal As Double
Dim Perc     As Double
Dim Variance As Double
Dim VariancePerc As Double
Dim i As Integer

    With uRecipe(1)
    
        '--------------------------------------------
        ' ricalcolo total weight + theoretical weight
        '--------------------------------------------
        
       ' Perc = .RmxRecipe(IndexComponent).Perc
       ' Variance = CDbl(txAcquisition(13)) - .RmxRecipe(IndexComponent).TheoreticalWeight
       ' VariancePerc = (Variance / .RmxRecipe(IndexComponent).TheoreticalWeight)
       
       If .RmxRecipe(IndexComponent).TheoreticalWeight = 0 And .RmxRecipe(IndexComponent).bAddedInPreparation Then
            .RmxRecipe(IndexComponent).TheoreticalWeight = .RmxRecipe(IndexComponent).RealWeight
       End If
        
        Perc = CDbl(txAcquisition(13) + .RmxRecipe(IndexComponent).RealWeight) / (.RmxRecipe(IndexComponent).TheoreticalWeight)
               
        NewTotal = .TotalWeightKg * Perc
            
        If F_MsgBox.DoShow("Recalculate Recipe Total Weight?" & vbCrLf & "New Total Weight : " & Trim(PadString(NewTotal)) & " kg", RecipeCode) Then
    
            
            userAcquisition.bRecalculation = True
            uRecipe(1).bRecalculation = True
            
            
            .bRecalculation = True
            
            .TotalWeightKg = NewTotal
            
            For i = 0 To .RmxRecipeCount
                If .RmxRecipe(i).Perc = 0 Then
                    .RmxRecipe(i).RealPerc = (.RmxRecipe(i).RealWeight / (.ActualWeight * 1000)) * 100
                End If
                
                Perc = IIf(.RmxRecipe(i).Perc = 0, .RmxRecipe(i).RealPerc, .RmxRecipe(i).Perc)
                
               .RmxRecipe(i).TheoreticalWeight = NewTotal * 1000 * Perc / 100
                
            Next
                
                
            frCommandInside_Click 8
            
            
        End If
        
        
    End With
End Sub




Private Function SaveAcquisitionInTabAcquisition()

With dbTabAcquisition
    .AddNew
    !AcquisitionTime = userAcquisition.AcquisitionTime
    !Code = userAcquisition.PrepBarcode.Code
    !ChemicalName = userAcquisition.PrepBarcode.ChemicalName
    !Cas = userAcquisition.PrepBarcode.Cas
    !Manufacturer = userAcquisition.PrepBarcode.Manufacturer
    !ManufacturerCode = userAcquisition.PrepBarcode.ManufacturerCode
    !ManufacturerLot = userAcquisition.PrepBarcode.ManufacturerLot
    !DeliveryDate = userAcquisition.PrepBarcode.DeliveryDate
    !QtyDelivered = userAcquisition.PrepBarcode.QtyDelivered
    !Package = userAcquisition.PrepBarcode.Package
    !WeekDelPackageNumber = userAcquisition.PrepBarcode.WeekDelPackageNumber
    !Index = userAcquisition.Index
    !ActualWeight = userAcquisition.ActualWeight
    !bRecalculation = userAcquisition.bRecalculation
    !bRecipeComponent = userAcquisition.bRecipeComponent
    !bFromBarcode = userAcquisition.bFromBarcode
    !Note = userAcquisition.Note
    !Operator = userAcquisition.Operator
    !RecipeCode = uPreparation.Recipes(1).Code
    !PrepWeek = uPreparation.PrepWeek
    !NumberPrepWeek = uPreparation.numPrepWeek
    !FileName = SettingName
    !PreparationID = PreparationID
    !ExpDate = userAcquisition.ExpDate
    .Update
    
    userAcquisition.ID = !ID

End With


End Function
Private Function DeleteComponent()
Dim IndexComp As Integer

Dim i As Integer

If CHCode <> "" Then

     
    If F_MsgBox.DoShow("Delete Component : " & CHCode & vbCrLf & "Weight : " & Trim(PadString(ComponentWeight)) & "g", RecipeCode, True) Then
    
    Else
        Exit Function
    End If

    With uRecipe(1)
    
        Debug.Print .RmxRecipe(lComponentRow).CHCode
        .RmxRecipe(lComponentRow).bDeleted = True
    End With
    

    
cont:

    
    '-----------------------------------
    ' cancello dalla tabella
    '-----------------------------------
    
    Grid1.ReadOnly = False
    Grid1.Selection.DeleteByRow
    Grid1.ReadOnly = True
    
    '-------------------------------------------
    ' ricarico la Gird1
    '-------------------------------------------
    
    uPreparation.Recipes(1) = uRecipe(1)
    
    Call FillGridPreparationFromFile(Grid1, uPreparation, 1, PreparationID)
    
    uRecipe(1).ActualWeight = uPreparation.Recipes(1).ActualWeight
    
    PopupMessage 2, "Component Deleted....", , , ComponentChCode
    frCommandInside(1).Visible = False
End If


End Function


Private Function DeleteAcquisition()
Dim IndexComp As Integer

Dim i As Integer

If AcquisitionID > 0 And AcquisitionChCode <> "" Then

     
    If F_MsgBox.DoShow("Delete Acquisition Code : " & AcquisitionChCode & vbCrLf & "Weight : " & Trim(PadString(AcquisitionWeight)) & "g", RecipeCode, True) Then
    
    Else
        Exit Function
    End If
    '-----------------------------------
    ' sottraggo il peso inserito...
    '-----------------------------------
    With uRecipe(1)
    
      '  Debug.Print .Acquisitions(lAcquisitionRow).ActualWeight
        .Acquisitions(lAcquisitionRow).bDeleted = True
        
        
        For i = 0 To .RmxRecipeCount
        
            If .RmxRecipe(i).CHCode = AcquisitionChCode Then
                
                ' sottraggo...
                
                If .RmxRecipe(i).RealWeight < AcquisitionWeight Then
           
                    .RmxRecipe(i).RealWeight = 0
                    
                Else
                    .RmxRecipe(i).RealWeight = .RmxRecipe(i).RealWeight - AcquisitionWeight
                    
                End If
                GoTo cont:
            End If
        
        Next
    
    End With
    
    If F_MsgBox.DoShow("Code not found in Recipe. Delete acquisition anyway?", AcquisitionChCode) Then
    Else
        Exit Function
    End If
    
    
    
    
cont:
    '-----------------------------------
    ' cancello la riga
    '-----------------------------------
        
    Call DeleteRowInTabAcquisition(AcquisitionID)
    
    '-----------------------------------
    ' cancello dalla tabella
    '-----------------------------------
    
    Grid2.ReadOnly = False
    Grid2.Selection.DeleteByRow
    Grid2.ReadOnly = True
    
    '-------------------------------------------
    ' ricarico la Gird1
    '-------------------------------------------
    
    uPreparation.Recipes(1) = uRecipe(1)
    
    Call FillGridPreparationFromFile(Grid1, uPreparation, 1, PreparationID)
    
    uRecipe(1).ActualWeight = uPreparation.Recipes(1).ActualWeight
    
    PopupMessage 2, "Acquisition Deleted....", , , AcquisitionChCode
    frCommandInside(11).Visible = False
End If

End Function


Private Function SavePreparation()
Dim i As Integer

lbWait = "Wait : Saving Data | Preparation..."
lbLine.Visible = False
lbWait.Visible = True

    '---------------------------------------------
    ' aggiorna RealPerc / Variance / VariancePerc
    '---------------------------------------------
    Call SetPreparationPercentageInRmxRecipe
    
    
    
    With uPreparation
        .Recipes(1) = uRecipe(1)
        .bSaved = True
        .Note = txFormulation(5)
        .ExpDate = txFormulation(8)
        .numPrepWeek = txFormulation(2)
        .PlannedPrepWeek = txFormulation(3)
        .PlanningReference = txFormulation(4)
        .PreparationDate = txFormulation(1)
        .PrepWeek = txFormulation(6)
        .RecipeBy = txFormulation(0)
        .PreparationLot = txFormulation(9)
        .bPesatoTuttiComponenti = SetbPesatoTuttiComponenti
        .OperatorPrep = txFormulation(7)
        If .HannaCodesCount > 0 Then
            For i = 1 To .HannaCodesCount ' UBound(.HannaCodes)
             Debug.Print .HannaCodes(i).Code
                If .HannaCodes(i).bHide = False Then
                    .HannaCodes(i).ExpDate = .ExpDate
                End If
            Next
        End If
    End With
    txFormulation(9).BackColor = vbWhite
    txFormulation(9).ForeColor = vbBlack
    '-------------------------------------------
    ' Salva e aggiorna TabPreparation
    '-------------------------------------------
    Call AggiornaTabPreparation(PreparationID, uPreparation)
   
    '-------------------------------------------
    ' Salva e aggiorna File
    '-------------------------------------------
    Call PeparationSaveSetting(uPreparation, SettingName)
    
    lbWait.Visible = False
    lbLine.Visible = True
    PopupMessage 2, "Preparation correctly Saved", , , uRecipe(1).Code
    
    
    PicMenu(1).Visible = True

    
End Function


Private Function SetPreparationPercentageInRmxRecipe()

Dim NewTotal As Double
Dim Perc     As Double
Dim Variance As Double
Dim VariancePerc As Double
Dim i As Integer

    With uRecipe(1)
            If .ActualWeight <> 0 Then
            For i = 0 To .RmxRecipeCount
              
              If .RmxRecipe(i).RealWeight <> 0 Then
                   .RmxRecipe(i).RealPerc = (.RmxRecipe(i).RealWeight / (.ActualWeight * 1000)) * 100
            
                    If (.RmxRecipe(i).RealWeight) > 0 And (.RmxRecipe(i).TheoreticalWeight) > 0 Then
                        .RmxRecipe(i).Variance = .RmxRecipe(i).RealWeight - .RmxRecipe(i).TheoreticalWeight
                    End If
                    
                    If (.RmxRecipe(i).Variance) <> 0 And (.RmxRecipe(i).RealWeight) > 0 Then
                         .RmxRecipe(i).VariancePerc = (.RmxRecipe(i).Variance / .RmxRecipe(i).RealWeight) * 100
                    End If
               End If
               
            Next
            
            End If
    End With
    
End Function
Private Function SetbPesatoTuttiComponenti() As Boolean
Dim rc As Boolean
Dim i As Integer

    rc = True
    
    With Grid1
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                'Debug.Print .Cell(i, 2).Text
                If .Cell(i, 2).Text <> "" Then
                    If IsNumeric(.Cell(i, 7).Text) Then
                        Debug.Print .Cell(i, 7).Text
                        If CDbl(.Cell(i, 7).Text) = 0 Then
                            rc = False
                            SetbPesatoTuttiComponenti = rc
                            Exit Function
                        End If
                    Else
                        rc = False
                    End If
                End If
            Next
        Else
            rc = False
        End If
    End With

    SetbPesatoTuttiComponenti = rc

End Function
Private Function ScanQRCode() As Boolean

MessageInfoTime = 200
Call ClearAcquisition
txQRCode = ""
txQRCode.Top = 0
txQRCode.Visible = True
txQRCode.SetFocus
PBContainer.Top = -(frInside(2).Top - 680)
PopupMessage 2, "Scan QRCode....", , , "Preparation"


End Function





Private Sub txQRCode_KeyPress(KeyAscii As Integer)
Dim rc As Boolean

On Error GoTo ERR_QR:

Dim UserQrCode As Barcode

If KeyAscii = 13 Then

    
    DoEvents
    If txQRCode = "" Then Exit Sub
   
    If CheckQRCode_HANNA(Trim(txQRCode), UserQrCode) Then
    

        
        If FillUserRMCode(UserQrCode.Code, True) Then
        
        
            ' se è un componente o se ho comunque deciso di acquisirlo allora
            ' copio i dati in form
            ' altrimentri niente...
            
        
            txAcquisition(0) = UserQrCode.Code
            txAcquisition(1) = UserQrCode.ChemicalName
           ' txAcquisition(2) = ""
            
            txAcquisition(3) = UserQrCode.Manufacturer
            txAcquisition(4) = UserQrCode.ManufacturerCode
            txAcquisition(5) = UserQrCode.ManufacturerLot
            txAcquisition(6) = UserQrCode.DeliveryDate
            txAcquisition(7) = UserQrCode.QtyDelivered
            txAcquisition(8) = UserQrCode.WeekDelPackageNumber
            txAcquisition(9) = UserQrCode.Package
           
            txQRCode.Visible = True
            
            PopupMessage 2, "Read Code : " & UserQrCode.Code
        
            ' inserisci automaticamente la pesata...
            'txAcquisition_Click 13
            
        
            If bOpenProductClassificationAfterScan Then
                
                frCommandInside_Click 9
            
            End If
        
        
        
        End If
    Else
        PopupMessage 2, "Please Check QRCode or Scan Again...", , True, "QRCode"
    End If
    

End If

ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_QR:
    MessageInfoTime = 2000
    PopupMessage 2, err.Description & vbCrLf & "Please repeat reading...", , , "QR Code Reader"
    Resume ERR_END:


End Sub


Private Sub SetClosedPreparation(ByVal rc As Boolean)

    
    frCommandInside(4).Visible = Not (rc)
    frCommandInside(0).Visible = Not (rc)
    frCommandInside(1).Visible = Not (rc)
    frCommandInside(5).Visible = Not (rc)
    frCommandInside(10).Visible = Not (rc)
    frExcel.Move frCommandInside(4).Left, frCommandInside(4).Top, frCommandInside(4).Width, frCommandInside(4).Height
    frExcel.Visible = rc
End Sub

Private Sub frExcel_Click()
            
    Dim ExcelFilename As String
    ' export LOT Excel
    If SettingName = "" Then
    Else

        ExcelFilename = "PREP_" & FormatNomeFile(Trim(uRecipe(1).Code) & "." & Trim(uRecipe(1).Line) & "." & txFormulation(2) & "." & txFormulation(6) & ".LOT_" & txFormulation(9))
        
        
        
        PopupMessage 2, "Exporting data to Excel : please wait...." & vbCrLf & ExcelFilename
        If Len(ExcelFilename) > 40 Then ExcelFilename = Left$(ExcelFilename, 40)
        ExcelFilename = ExcelFilename & ".xls"
        Call EsportaPreparationExcel(SettingName, ExcelFilename, uPreparation)
    End If


End Sub


Private Sub SetuPreparationExpDate()
If uRecipe(1).Exp = "" Then uRecipe(1).Exp = GetRecipeExp(uRecipe(1).Code)
uPreparation.ExpDate = SetExpDate(uPreparation.PreparationDate, uRecipe(1).Exp)
txFormulation(8) = uPreparation.ExpDate
End Sub
















Private Sub Grid4_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)


NotesID = 0

    With Grid4
    
        If FirstRow > 0 Then
        
            NotesID = .Cell(FirstRow, 5).Text
            txPreparation(0) = .Cell(FirstRow, 1).Text
            txPreparation(1) = .Cell(FirstRow, 2).Text
            txPreparation(2) = .Cell(FirstRow, 4).Text
            txPreparation(3) = .Cell(FirstRow, 3).Text
            
        End If
    
    End With


End Sub

Private Sub lbFunction_Click(Index As Integer)
ImCode_Click Index
End Sub

Private Sub ImCode_Click(Index As Integer)
Select Case Index
                        
                        
        Case 5
            ' aggiungi Notes specifics
            If AddNotes Then
                 Call GetPreparationNotes(Grid4, SettingName)
            End If
            
            frExcel2.Visible = IIf(Grid4.Rows > 1, True, False)
            Frame6.Visible = IIf(Grid4.Rows > 1, False, True)
        Case 4
            ' delete Notes specifics
            If DeleteNotes(NotesID) Then
                 Call GetPreparationNotes(Grid4, SettingName)
            End If
            
            frExcel2.Visible = IIf(Grid4.Rows > 1, True, False)
            Frame6.Visible = IIf(Grid4.Rows > 1, False, True)
            
                                   
        Case 0
            ' aggiungi revision specifics
            If CheckPrivilege(2) Then
                If AddRevision(RecipeCode, txRevision(2)) Then
                     Call GetRecipeRevision(Grid5, RecipeCode)
                End If
           
            
            
                frExcel3.Visible = IIf(Grid5.Rows > 1, True, False)
                Frame4.Visible = IIf(Grid5.Rows > 1, False, True)
             End If
        Case 1
            ' delete revision specifics
            If DeleteRevision(RecipeCode, txRevision(2)) Then
                 Call GetRecipeRevision(Grid5, RecipeCode)
            End If
            
            frExcel3.Visible = IIf(Grid5.Rows > 1, True, False)
            Frame4.Visible = IIf(Grid5.Rows > 1, False, True)
            
             
            
            
            
            
            

End Select

End Sub




Private Sub ClearPrepNotesForm()



Dim i As Integer
For i = 0 To txPreparation.UBound
    txPreparation(i) = ""
Next
txPreparation(2) = MyOperatore.Name

End Sub


Private Sub txPreparation_Click(Index As Integer)
Dim userCode As String
Dim Answer As String
Dim Selected As String
Dim bNumber As Boolean
Dim sString As String
Dim rc As Boolean

    Selected = lbRevision(Index) ' "Preparation"
    Answer = txPreparation(Index)
    sString = "Please Enter " & lbRevision(Index)
  
    bNumber = False
    
    If txPreparation(2) = "" Then txPreparation(2) = MyOperatore.Name
    
    cmbNotes.Visible = False
    
    Select Case Index
        Case 0
            If Answer = "" Then Answer = FormatDataLAT(Now())
        Case 1
            ' type
            cmbNotes.ZOrder
            cmbNotes.Visible = True
            Exit Sub
        Case 2
        
    End Select
    
    
    If txPreparation(Index).Locked Then Exit Sub
    
    
  
    If F_InputBox.DoShow(sString, Selected, , , , Answer, , bNumber, Me.Top) Then
    
        txPreparation(Index) = Answer
        
        Select Case Index
            Case 0
                ' isdate?
                If IsDate(Answer) Then
                     txPreparation(Index) = FormatDataLAT(Answer)
                Else
                    PopupMessage 2, "Please enter a valid Date...", , True
                End If
        End Select
    End If
    
    
    
    
End Sub

Private Function DeleteNotes(ByVal ID As Long) As Boolean
Dim rc As Boolean
Dim i As Integer

rc = True

For i = 1 To txPreparation.UBound
    If txPreparation(i) = "" Then
        rc = False
        PopupMessage 2, "Please Select a Note form the table...", , True, "Delete Notes"
        DeleteNotes = rc
        Exit Function
    End If
Next

With dbTabPreparationNotes
    .filter = ""
    .filter = "ID='" & ID & "'"
    If .EOF Then
        
    Else
        If F_MsgBox.DoShow("Delete Note?", "Preparation Notes", , "Delete", "Exit") Then
            .Delete
            .Update
        Else
            rc = False
        End If
    End If
     
        
        
        
    



End With

DeleteNotes = rc
End Function
Private Function AddNotes() As Boolean
Dim rc As Boolean
Dim i As Integer
Dim OldDate As Date

rc = True

For i = 1 To txPreparation.UBound
    If txPreparation(i) = "" Then
        rc = False
        PopupMessage 2, "Please enter all fields...", , True, "Notes History"
        AddNotes = rc
        Exit Function
    End If
Next





With dbTabPreparationNotes
    .filter = ""
    .filter = "filename='" & SettingName & "' and NoteDate='" & txPreparation(0) & "'"
    If .EOF Then
        .AddNew
    Else
        .MoveFirst
        OldDate = FormatDataLAT(Trim(!NoteDate))
        If F_MsgBox.DoShow("Note Date : " & OldDate & " already exsists.", "Add Preparation Note", , "Add", "Exit") Then
            .AddNew
        Else
            AddNotes = False
            Exit Function
        End If
    End If
        
        !NoteDate = txPreparation(0)
        !Type = txPreparation(1)
        !Description = IIf(Len(txPreparation(3)) > 255, Left(txPreparation(3), 255), txPreparation(3))
        !Operator = txPreparation(2)
        !FileName = SettingName
        .Update


End With

AddNotes = rc
End Function
Private Sub Image1_Click()
frExcel2_Click
End Sub
Private Sub Label4_Click()
frExcel2_Click
End Sub
Private Sub frExcel2_Click()
    If SettingName <> "" Then
        If Grid4.ExportToExcel(USER_DESKTOP & "\" & SettingName & "_PreparationNote_History.xls", True, True) Then
            MessageInfoTime = 2500
            PopupMessage 2, "File correcly created on Desktop", , , RecipeCode & "_Note_History.xls"
        End If
    End If
End Sub


Private Sub cmbNotes_Click()
txPreparation(1) = cmbNotes
cmbNotes.Visible = False
End Sub

Private Sub AddcmbNotes()


    With cmbNotes
        .AddItem "Revision"
        .AddItem "Improvement"
        .AddItem "Issue"
        .ListIndex = 0
    End With

End Sub











Private Sub OpenRevisionHistory()

            
AddcmbRevType

lbInside(4) = RecipeCode & " : Revision History"

Call GetRecipeRevision(Grid5, RecipeCode)

frExcel3.Visible = IIf(Grid5.Rows > 1, True, False)
Frame4.Visible = IIf(Grid5.Rows > 1, False, True)



frInside(5).Visible = True
frInside(5).ZOrder
PBContainer.Top = -(frInside(5).Top - 680)


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

On Error GoTo ERR_CHECK:

    Selected = lbRevHist(Index) ' "Preparation"
    Answer = txRevision(Index)
    sString = "Please Enter " & lbRevHist(Index)
  
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
             If uRecipe(1).Rev <> "" Then
                If IsNumeric(uRecipe(1).Rev) Then
                    Answer = uRecipe(1).Rev
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
    
ERR_END:
    On Error GoTo 0
    Exit Sub
ERR_CHECK:
    rc = False
    GoTo ERR_END:
    
    
    
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




Private Sub Grid5_SelChange(ByVal FirstRow As Long, ByVal FirstCol As Long, ByVal LastRow As Long, ByVal LastCol As Long)


RevisionID = 0

With Grid5

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


Private Sub frExcel3_Click()
    Grid5.ExportToExcel USER_DESKTOP & "\" & FormatNomeFile(RecipeCode) & "_RevHistory.xls", True, True
    MessageInfoTime = 2500
    PopupMessage 2, "File correcly created on Desktop", , , FormatNomeFile(RecipeCode) & "_RevHistory.xls"
End Sub













'-----------------------------------------------------------------------------------------------
'
'
'                                   MATERIAL REQUISITION

'
'-----------------------------------------------------------------------------------------------

Private Sub lbCommand_Click()
SaveSetting App.Title, "PREP_" & strLine, nrWeek, Nr
SaveSetting App.Title, "MaterialRequisition", "Dep.Manager", txDocument(5)
Call StampaMaterialRequisition

End Sub



Private Sub lbpdf_Click()
lbCommand_Click
End Sub


Private Sub txDocument_LostFocus(Index As Integer)
    Select Case Index
        Case 2
            SaveSetting App.Title, "Workstation", "no.department", txDocument(Index)
        
    End Select

End Sub


Private Function GetRecipeCodeFromMixes(ByRef Index As Integer) As String
Dim bMix As Boolean
Dim MixCode As String
Dim RecipeCode As String

Dim i As Integer
With Grid4
  If .Rows > 1 Then
        For i = 1 To .Rows - 1
        
            bMix = .Cell(i, 15).Text
            
            If bMix Then
            
                MixCode = .Cell(i, 1).Text
                Index = i
                GetRecipeCodeFromMixes = MixCode
                Exit Function
        
            End If
        Next
  End If

End With

End Function








Private Sub SetMaterialRequisitionComponents()
Dim xDocument() As String
Dim rc As Boolean

    If CheckTxTxFormulation(xDocument()) Then
        Call GotoMaterialRequisition
    End If

   ' If bImportata Then
   '     If F_MsgBox.DoShow("Material Requisition already done. Create new Material Requisition?", SelectedRecipeCode) = False Then
   '         rc = AddMaterialRequisitionFromFile(Grid6, IndexRecipe)
   '         If rc = False Then GoTo MrComponent:
   '         PicMenu_Click 1
   '     Else
   '         GoTo MrComponent
   '     End If
   ' Else
MrComponent:
       ' If SelectedRecipeCode = "" Then
          '  PopupMessage 2, "Please select a Recipe first..", , True
       ' Else
            
       ' End If
  '  End If
End Sub

Private Sub GotoMaterialRequisition()
Dim rc As Boolean
Dim Index As Integer
Dim i As Integer
Dim MaterialReqRecipe As RecipeType
Dim MaterialReqMixes() As RecipeType
Dim ArrayMaterialReqRecipe() As RecipeType


    If IndexRecipe = 0 Then IndexRecipe = 1

        ReDim ArrayMaterialReqRecipe(1)
        Index = IndexRecipe
        

        rc = SetMaterialRequisitionPreparation(uRecipe, MaterialReqRecipe, Grid1, True, SettingName)
        ArrayMaterialReqRecipe(1) = MaterialReqRecipe
        ArrayMaterialReqRecipe(1).Line = uPreparation.Recipes(1).Line
        
        
        
        'txDocument(2) = uPreparation.Recipes(1).Line


        If rc Then
            lRowMaterialReq = 0
            lColMaterialReq = 0
        
            Call AddRecipeToMaterialReqGrid(Grid6, ArrayMaterialReqRecipe(), True)
            PicMenu_Click 1
    
        End If
      
        
      
   

End Sub
Private Sub StampaMaterialRequisition()
Dim rc As Boolean
Dim FileName As String
Dim xDocument() As String
Dim strHannaCode As String
Dim strRecipe As String

    rc = True
     
    rc = CheckTxDocument(xDocument())
    If rc = False Then Exit Sub
    
    CloseSettingDataFile
    
    lbWait.Caption = "Material Requisition PDF file : Wait while Saving Data..."
    lbWait.Visible = True
    lbLine.Visible = False



    rc = MaterialRequisitionSaveSettingsFile(Grid6, xDocument(), SettingName, IndexRecipe)
    
    Call SaveMaterialRequisitionForRecipeForProductionInDatabase
    
    If uPreparation.HannaCodesCount = 0 Then
    Else
        strHannaCode = SetNeHannaCodeQtyString(uPreparation.HannaCodes)
    End If
    strRecipe = "Recipe : " & uPreparation.Recipes(1).Code & " | " & uPreparation.Recipes(1).Description
    
    
    If rc Then rc = MaterialRequisitionSaveSettingsTempFile(Grid6, xDocument(), FileName, strHannaCode, strRecipe)
    If rc Then rc = ReportStampato(FileName)
    If rc Then
        PopupMessage 2, "Document Succesfully Generated...", , , "Material Requisition : MR-" & xDocument(0)
        
        
        ' user temp / data Material Requisition  TEMP file! da cancellare! lo uso solo per stampare...
        
        If FileExists(USER_PATH & FileName) Then
        

            Kill USER_PATH & FileName

        End If
        
        
    End If
    CloseSettingDataFile
    lbWait.Visible = False
    lbLine.Visible = True


End Sub


Private Function SaveMaterialRequisitionForRecipeForProductionInDatabase() As Boolean
Dim rc As Boolean
Dim strMaterialRequisition As String
On Error GoTo SaveReceipt
    rc = True
    With dbTabReceiptForProduction
        .filter = ""
        .filter = "FileName ='" & SettingName & IIf(bIfDataPath, "' and bClosed=true", "' and bClosed=false")
        If .EOF Then
                
            .AddNew
                
            !Recipe = GetStrRecipe(uRecipe)
            !Description = GetStrDescriptionRecipe(uRecipe)
            !Line = GetStrLineRecipe(uRecipe)
            !PlanningReference = txFormulation(4)
            
            !DataRecipe = txFormulation(1)
            !RecipeWeek = txFormulation(2)
            !PlannedPreparation = txFormulation(3)
            !Operator = txFormulation(0)
            !bClosed = bIfDataPath
            !Note = txFormulation(5)
            !FileName = SettingName
        Else
        
        End If
        
        If IsNull(Trim(!MaterialRequisitionNumber)) Or Trim(!MaterialRequisitionNumber) = "" Then
            strMaterialRequisition = txDocument(0)
        Else
            strMaterialRequisition = CheckStrMaterialRequisition(Trim(!MaterialRequisitionNumber), txDocument(0))
        End If
        
        !MaterialRequisitionNumber = strMaterialRequisition
        !bMaterialRequisitionPrinted = True
        .Update
    
    End With

ERR_END:
    On Error GoTo 0
    SaveMaterialRequisitionForRecipeForProductionInDatabase = rc
    Exit Function
SaveReceipt:
    rc = False
    MsgBox err.Description
    Resume Next
    
End Function



Private Function ReportStampato(ByVal FileName As String) As Boolean
    Dim rc As Boolean
    Dim NumReport As String
    On Error GoTo ERR_SAVE
    rc = True

    NumReport = FormatNomeFile(txDocument(0) & "." & txDocument(1))

    rc = OkStampa(NumReport, bSeStampa, FileName, True)
     
ERR_END:
    On Error GoTo 0
    ReportStampato = rc
    Exit Function
ERR_SAVE:
    rc = False
    Resume ERR_END
End Function

Private Function CheckTxDocument(ByRef xDocument() As String) As Boolean
Dim rc As Boolean
Dim i As Integer
    rc = True
    ReDim xDocument(txDocument.UBound)
    
    For i = txDocument.LBound To txDocument.UBound
        xDocument(i) = txDocument(i)
        If Len(txDocument(i)) = 0 Then
            rc = False
            PopupMessage 2, "Please Enter field : " & lbDocument(i), , True, "Preparation Data"
            txDocument(i).SetFocus
            Exit For
        End If
    Next
    CheckTxDocument = rc
End Function

Private Function CheckTxTxFormulation(ByRef xDocument() As String) As Boolean
Dim rc As Boolean
Dim i As Integer
    rc = True
    ReDim xDocument(txFormulation.UBound)
    
    For i = txFormulation.LBound To txFormulation.UBound
        xDocument(i) = txFormulation(i)
        
        If i = 5 Then GoTo cont:
        
        If Len(txFormulation(i)) = 0 Then
            rc = False
            PopupMessage 2, "Please Enter field : " & lbFormulation(i), , True, "Preparation Data"
            txFormulation(i).SetFocus
            Exit For
        End If
cont:
    Next
    CheckTxTxFormulation = rc
End Function


Public Function AddMaterialRequisitionFromFile(ByVal Grd As Grid, ByVal Index As Integer) As Boolean
Dim rc As Boolean
Dim i As Integer
Dim t As Integer
Dim X As Integer
Dim RowsCount As Integer

    On Error GoTo ERR_ADD:

    If Index = 0 Then Index = 1
    
    rc = True
    

    If SettingName = "" Then
        rc = False
        GoTo ERR_END
     
    End If

    CloseSettingDataFile


    For i = 0 To txDocument.UBound
        txDocument(i) = GetSettingData(SettingName, "Material Requisition" & Index, "txDocument(" & i & ")", "")
    Next
    RowsCount = GetSettingData(SettingName, "Material Requisition" & Index, "Rows", 0)
    
    
    If RowsCount = 0 Then
        ' ho riaperto la Recipe ma non hpo fatto Material requisition
        rc = False
        GoTo ERR_END
    End If
    
    
    With Grd
        .AutoRedraw = False
        .Rows = 1
        For i = 1 To RowsCount
            .AddItem "", False
            For t = 1 To .Cols - 1
                .Cell(i, t).Text = GetSettingData(SettingName, "Material Requisition" & Index, "Grd(" & i & "," & t & ")", "")
                .Column(t).Alignment = cellLeftCenter
                .Column(t).Width = 150
                .Cell(0, t).FontBold = True
            
            Next
        Next
        
        .Column(2).Width = 250
        .Column(3).Width = 100
        .Column(5).Width = 100
        .Column(4).Alignment = cellRightCenter
        .Column(5).Alignment = cellCenterCenter
        .Column(6).Alignment = cellCenterCenter
        .Refresh
        .AutoRedraw = True
    End With

ERR_END:
    On Error GoTo 0
    AddMaterialRequisitionFromFile = rc
    Exit Function
ERR_ADD:
    rc = False
    MsgBox err.Description
    GoTo ERR_END
End Function




